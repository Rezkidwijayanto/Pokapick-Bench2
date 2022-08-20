Imports System.Data
Imports System.Data.SqlClient
Imports EasyModbus
Imports System.Text
Imports System.Threading


Public Class Form2
    Dim modbusClient As ModbusClient

    Private Sub btnConnect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConnect.Click
        Try
            modbusClient = New ModbusClient(txtIP.Text, Val(txtPort.Text))
            modbusClient.Connect()

            btnConnect.Enabled = False
            btnDisconnect.Enabled = True
        Catch ex As Exception
            MsgBox("Error Connect! " & ex.Message)
        End Try
    End Sub
    Sub LihatIP()
        Dim Conn As SqlConnection
        Dim SQl1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()
            SQl1 = Conn.CreateCommand
            SQl1.CommandText = "SELECT *FROM s_config WHERE (NoID=1)"
            Reader = SQl1.ExecuteReader
            If Reader.Read Then
                txtIP.Text = Reader("IPAddress").ToString
                txtPort.Text = Reader("Port").ToString
            Else
                txtIP.Text = ""
                txtPort.Text = ""

            End If
            Conn.Close()
            Reader.Close()

        Catch ex As Exception

        End Try
    End Sub
    Private Sub btnDisconnect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDisconnect.Click
        Try
            modbusClient.Disconnect()

            btnConnect.Enabled = True
            btnDisconnect.Enabled = False
        Catch ex As Exception
            MsgBox("Error Disconnect! " & ex.Message)
        End Try
    End Sub

    Private Sub btnRead_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRead.Click
        Try
            Dim StartAddress As Integer = Val(txtAddress.Text)

            Select Case cboRegType.SelectedIndex
                Case 0
                    If StartAddress > 0 Then StartAddress = StartAddress - 1
                    Dim ReadValues() As Boolean = modbusClient.ReadCoils(StartAddress, 1)
                    txtValue.Text = ReadValues(0)
                Case 1
                    If StartAddress > 10000 Then StartAddress = StartAddress - 10001
                    Dim ReadValues() As Boolean = modbusClient.ReadDiscreteInputs(StartAddress, 1)
                    txtValue.Text = ReadValues(0)
                Case 2
                    If StartAddress > 30000 Then StartAddress = StartAddress - 30001
                    Dim ReadValues() As Integer = modbusClient.ReadInputRegisters(StartAddress, 1)
                    txtValue.Text = ReadValues(0)
                Case 3
                    If StartAddress > 40000 Then StartAddress = StartAddress - 40001
                    Dim ReadValues() As Integer = modbusClient.ReadHoldingRegisters(StartAddress, 1)
                    txtValue.Text = ReadValues(0)
            End Select

        Catch ex As Exception
            MsgBox("Error Read! " & ex.Message)
        End Try
    End Sub

    Private Sub btnWrite_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnWrite.Click
        Try
            Dim StartAddress As Integer = Val(txtAddress.Text)

            Select Case cboRegType.SelectedIndex
                Case 0
                    If StartAddress > 0 Then StartAddress = StartAddress - 1
                    Dim WriteVals(0) As Boolean
                    WriteVals(0) = Val(txtNewValue.Text)
                    modbusClient.WriteMultipleCoils(StartAddress, WriteVals)
                    'modbusClient.WriteSingleCoil(StartAddress, WriteVals)
                Case 1
                    'If StartAddress > 10000 Then StartAddress = StartAddress - 10001
                    'Dim WriteVals(0) As Boolean
                    'WriteVals(0) = Val(txtNewValue.Text)
                    'modbusClient.WriteSingleRegister
                    'modbusClient.WriteMultipleCoils(StartAddress, WriteVals)
                    'modbusClient.WriteSingleCoil(StartAddress, WriteVals)
                    '
                Case 2
                    If StartAddress > 30000 Then StartAddress = StartAddress - 30001
                    'Dim WriteVals As String
                    'WriteVals = txtNewValue.Text
                    'modbusClient.WriteMultipleRegisters(StartAddress, txtNewValue.Text)
                    'modbusClient.WriteSingleRegister()

                    'If Dim StartAddress As Integer = 515
                    Dim Values As String
                    Values = txtNewValue.Text
                    'modBusClient.WriteMultipleRegisters("515", txtRef.Text)
                    modbusClient.WriteSingleRegister(StartAddress, txtNewValue.Text)
                    'modbusClient.WriteSingleRegister()
                    '
                Case 3

                    If StartAddress > 40000 Then StartAddress = StartAddress - 40001
                    Dim WriteVals(0) As Integer
                    WriteVals(0) = Val(txtNewValue.Text)
                    modbusClient.WriteMultipleRegisters(StartAddress, WriteVals)
            End Select
            
        Catch ex As Exception
            MsgBox("Error Write! " & ex.Message)
        End Try
    End Sub

    Private Sub cboRegType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboRegType.SelectedIndexChanged
        Dim iSel As Integer = cboRegType.SelectedIndex
        If iSel > -1 Then
            If iSel = 0 Then
                txtAddress.Text = "00001"
                btnWrite.Enabled = True
            End If
            If iSel = 1 Then
                txtAddress.Text = "10001"
                btnWrite.Enabled = False
            End If

            If iSel = 2 Then
                txtAddress.Text = "30001"
                btnWrite.Enabled = True
            End If
            If iSel = 3 Then
                txtAddress.Text = "40001"
                btnWrite.Enabled = True
            End If

        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()

    End Sub
    Sub Simpan()
        Dim Conn As SqlConnection
        Dim SQl1 As New SqlCommand
        Try
            Conn = GetConnect()
            Conn.Open()
            SQl1 = Conn.CreateCommand
            SQl1.CommandText = "UPDATE s_config SET NoID=1,IPAddress='" & txtIP.Text & "',Port='" & txtPort.Text & "'" _
                             & "WHERE (NoID=1)"
            SQl1.ExecuteReader()
            Conn.Close()

            MessageBox.Show("Configuration Saved!", "Success", MessageBoxButtons.OK)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles Me.Load
        Call LihatIP()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call Simpan()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try

            Dim StartAddress As Integer = Val(txtAddress.Text)
            'Dim StartAddress As Integer = 300
            'Dim Values As String
            'Values = txtNewValue.Text
            'modBusClient.WriteMultipleRegisters("515", txtRef.Text)
            modbusClient.WriteSingleRegister(StartAddress, txtNewValue.Text)

        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString)
        End Try
    End Sub
    Sub ConvertASCII()

    End Sub

    Private Sub BtnConvert_Click(sender As Object, e As EventArgs) Handles btnConvert.Click
        'Dim Sendbytes() As Byte = hConverter.HexToBinary(txtFrom.Text)
        Try
            txtTo.Text = ""
            Dim cb() As Byte
            'txtFrom.Clear()
            If txtFrom.Text <> "" Then
                cb = ASCIIEncoding.ASCII.GetBytes(txtFrom.Text)
                Dim MWval_reff(19) As Integer

                For x As Integer = 0 To cb.Length - 1
                    'txtFrom.Text = ""
                    txtTo.AppendText(cb(x).ToString & " ")
                    'Dim arrayRef() as Char=cb(x).to
                    'MessageBox.Show(cb(x).ToString)
                    MWval_reff(x) = Convert.ToInt64(cb(x))
                    Try
                        Dim StartAddress As Integer = Val(txtAddress.Text)
                        modbusClient.WriteMultipleRegisters(StartAddress, MWval_reff)
                    Catch ex As Exception

                    End Try
                Next
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString)
        End Try
        'publisher.Send(Sendbytes, Sendbytes.Length)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            Dim sepCH() As Char = New Char() {" "c, ","c}
            Dim s() As String
            txtFrom.Clear()
            If txtTo.Text <> "" Then
                s = txtTo.Text.Split(sepCH, StringSplitOptions.RemoveEmptyEntries)
                For x As Integer = 0 To s.Length - 1
                    txtFrom.AppendText(Convert.ToChar(CInt(s(x))))
                Next
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        'Dim Jumlahref As Integer

    End Sub
End Class