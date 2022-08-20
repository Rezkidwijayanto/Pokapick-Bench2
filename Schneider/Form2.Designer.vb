<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form2
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form2))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtIP = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtPort = New System.Windows.Forms.TextBox()
        Me.btnConnect = New System.Windows.Forms.Button()
        Me.btnDisconnect = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtAddress = New System.Windows.Forms.TextBox()
        Me.btnRead = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtValue = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtNewValue = New System.Windows.Forms.TextBox()
        Me.btnWrite = New System.Windows.Forms.Button()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.cboRegType = New System.Windows.Forms.ComboBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.txtFrom = New System.Windows.Forms.TextBox()
        Me.txtTo = New System.Windows.Forms.TextBox()
        Me.btnConvert = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(50, 61)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(54, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Server IP:"
        '
        'txtIP
        '
        Me.txtIP.Location = New System.Drawing.Point(110, 58)
        Me.txtIP.Name = "txtIP"
        Me.txtIP.Size = New System.Drawing.Size(124, 20)
        Me.txtIP.TabIndex = 1
        Me.txtIP.Text = "127.0.0.1"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(254, 61)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(29, 13)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Port:"
        '
        'txtPort
        '
        Me.txtPort.Location = New System.Drawing.Point(289, 58)
        Me.txtPort.Name = "txtPort"
        Me.txtPort.Size = New System.Drawing.Size(65, 20)
        Me.txtPort.TabIndex = 1
        Me.txtPort.Text = "502"
        '
        'btnConnect
        '
        Me.btnConnect.Location = New System.Drawing.Point(53, 100)
        Me.btnConnect.Name = "btnConnect"
        Me.btnConnect.Size = New System.Drawing.Size(147, 35)
        Me.btnConnect.TabIndex = 2
        Me.btnConnect.Text = "CONNECT"
        Me.btnConnect.UseVisualStyleBackColor = True
        '
        'btnDisconnect
        '
        Me.btnDisconnect.Enabled = False
        Me.btnDisconnect.Location = New System.Drawing.Point(207, 100)
        Me.btnDisconnect.Name = "btnDisconnect"
        Me.btnDisconnect.Size = New System.Drawing.Size(147, 35)
        Me.btnDisconnect.TabIndex = 2
        Me.btnDisconnect.Text = "DISCONNECT"
        Me.btnDisconnect.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(67, 179)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(51, 13)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "Address :"
        '
        'txtAddress
        '
        Me.txtAddress.Location = New System.Drawing.Point(124, 176)
        Me.txtAddress.Name = "txtAddress"
        Me.txtAddress.Size = New System.Drawing.Size(110, 20)
        Me.txtAddress.TabIndex = 1
        '
        'btnRead
        '
        Me.btnRead.Location = New System.Drawing.Point(240, 176)
        Me.btnRead.Name = "btnRead"
        Me.btnRead.Size = New System.Drawing.Size(114, 46)
        Me.btnRead.TabIndex = 3
        Me.btnRead.Text = "Get Data"
        Me.btnRead.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(78, 205)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(40, 13)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "Value :"
        '
        'txtValue
        '
        Me.txtValue.Location = New System.Drawing.Point(124, 202)
        Me.txtValue.Name = "txtValue"
        Me.txtValue.Size = New System.Drawing.Size(110, 20)
        Me.txtValue.TabIndex = 1
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(50, 239)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(65, 13)
        Me.Label5.TabIndex = 0
        Me.Label5.Text = "New Value :"
        '
        'txtNewValue
        '
        Me.txtNewValue.Location = New System.Drawing.Point(124, 236)
        Me.txtNewValue.Name = "txtNewValue"
        Me.txtNewValue.Size = New System.Drawing.Size(110, 20)
        Me.txtNewValue.TabIndex = 1
        '
        'btnWrite
        '
        Me.btnWrite.Location = New System.Drawing.Point(240, 228)
        Me.btnWrite.Name = "btnWrite"
        Me.btnWrite.Size = New System.Drawing.Size(114, 34)
        Me.btnWrite.TabIndex = 3
        Me.btnWrite.Text = "Sent Data"
        Me.btnWrite.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(58, 150)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(60, 13)
        Me.Label6.TabIndex = 0
        Me.Label6.Text = "Reg Type :"
        '
        'cboRegType
        '
        Me.cboRegType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboRegType.Font = New System.Drawing.Font("Consolas", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboRegType.Items.AddRange(New Object() {"Coil Outputs      (00000)", "Discrete Inputs    (10000)", "Register Inputs   (30000)", "Holding Registers (40000)"})
        Me.cboRegType.Location = New System.Drawing.Point(124, 144)
        Me.cboRegType.Name = "cboRegType"
        Me.cboRegType.Size = New System.Drawing.Size(230, 26)
        Me.cboRegType.TabIndex = 4
        '
        'Button1
        '
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(86, 310)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(114, 34)
        Me.Button1.TabIndex = 5
        Me.Button1.Text = "Save"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Image = CType(resources.GetObject("Button2.Image"), System.Drawing.Image)
        Me.Button2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button2.Location = New System.Drawing.Point(207, 310)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(114, 34)
        Me.Button2.TabIndex = 6
        Me.Button2.Text = "Cancel"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(240, 268)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(114, 34)
        Me.Button3.TabIndex = 7
        Me.Button3.Text = "Write Single"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'txtFrom
        '
        Me.txtFrom.Location = New System.Drawing.Point(421, 100)
        Me.txtFrom.Name = "txtFrom"
        Me.txtFrom.Size = New System.Drawing.Size(170, 20)
        Me.txtFrom.TabIndex = 8
        '
        'txtTo
        '
        Me.txtTo.Location = New System.Drawing.Point(421, 126)
        Me.txtTo.Name = "txtTo"
        Me.txtTo.Size = New System.Drawing.Size(170, 20)
        Me.txtTo.TabIndex = 9
        '
        'btnConvert
        '
        Me.btnConvert.Location = New System.Drawing.Point(477, 158)
        Me.btnConvert.Name = "btnConvert"
        Me.btnConvert.Size = New System.Drawing.Size(114, 34)
        Me.btnConvert.TabIndex = 10
        Me.btnConvert.Text = "Convert"
        Me.btnConvert.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(477, 202)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(114, 34)
        Me.Button4.TabIndex = 11
        Me.Button4.Text = "Convert 1"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(477, 242)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(114, 34)
        Me.Button5.TabIndex = 12
        Me.Button5.Text = "Cek Array"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Form2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(627, 372)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.btnConvert)
        Me.Controls.Add(Me.txtTo)
        Me.Controls.Add(Me.txtFrom)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.cboRegType)
        Me.Controls.Add(Me.btnWrite)
        Me.Controls.Add(Me.btnRead)
        Me.Controls.Add(Me.btnDisconnect)
        Me.Controls.Add(Me.btnConnect)
        Me.Controls.Add(Me.txtPort)
        Me.Controls.Add(Me.txtNewValue)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtValue)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtAddress)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtIP)
        Me.Controls.Add(Me.Label1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form2"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Device Configuration"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtIP As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtPort As System.Windows.Forms.TextBox
    Friend WithEvents btnConnect As System.Windows.Forms.Button
    Friend WithEvents btnDisconnect As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtAddress As System.Windows.Forms.TextBox
    Friend WithEvents btnRead As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtValue As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtNewValue As System.Windows.Forms.TextBox
    Friend WithEvents btnWrite As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cboRegType As System.Windows.Forms.ComboBox
    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents txtFrom As TextBox
    Friend WithEvents txtTo As TextBox
    Friend WithEvents btnConvert As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents Button5 As Button
End Class
