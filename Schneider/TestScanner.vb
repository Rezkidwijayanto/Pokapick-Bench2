Imports System.IO.Ports
Imports System.Data.SqlClient

Public Class TestScanner
    Dim WithEvents Com1 As New SerialPort

    Private Sub TestScanner_Load(sender As Object, e As EventArgs) Handles MyBase.Load


    End Sub

    Private Sub Com1_DataReceived(sender As Object, e As System.IO.Ports.SerialDataReceivedEventArgs) Handles Com1.DataReceived
        'Dim returnStr As String
        'returnStr = Com1.ReadExisting 'Com1.ReceivedBytesThreshold
        'Rec 'eiveSerialData(returnStr)

    End Sub
    Function ReceiveSerialData() As String
        ' Receive strings from a serial port.
        'Dim returnStr As String = ""

        'Dim com1 As IO.Ports.SerialPort = Nothing
        'Try
        'Com1 = My.Computer.Ports.OpenSerialPort("COM1")
        'Com1.ReadTimeout = 10000
        'Do
        'Dim Incoming As String = com1.ReadLine()
        'If Incoming Is Nothing Then
        'Exit Do
        'Else
        'returnStr &= Incoming & vbCrLf
        'txtData.Text = returnStr
        'End If
        'Loop
        'Catch ex As TimeoutException
        'returnStr = "Error: Serial Port read timed out."
        'Finally
        'If com1 IsNot Nothing Then com1.Close()
        'End Try

        'Return returnStr
    End Function

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        Try
            'Timer1.Start()
            'If txtScan.Text = txtData.Text Then
            'lblLabel.Text = txtScan.Text
            'lblDetik.Text = "Cocok"
            'txtScan.Text = ""
            'Else
            'lblDetik.Text = "scan"
            'txtScan.Text = ""
            'End If
            'Timer1.Stop()
        Catch ex As Exception

        End Try
    End Sub


    Sub Bandingkan()
        Try
            If txtScan.Text = txtData.Text Then
                lblLabel.Text = txtScan.Text
                lblDetik.Text = "Cocok"
            Else
                lblDetik.Text = "scan"
                txtScan.Text = ""
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Dim xx As String = txtScan.Text
        'Dim xstr As String = xx.Split(",").First

        'Dim xxstr() As String = xx.Split(",")
        'MsgBox(xstr)
        'MsgBox(xxstr(xxstr.Count + 2))
        'txtData.Text = Right(txtScan.Text, 2)

        'Dim xx As String = txtScan.Text
        'xx = xx.Substring(0, 4)

        'txtData.Text = xx
        Dim Scan As String = txtScan.Text
        Dim Data As String = Mid(Scan, 4, 8)

        txtData.Text = Data

        If txtData.Text = txtData1.Text Then
            lblLabel.Text = "Benar"
        Else
            lblLabel.Text = "Salah"
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim SQL2 As New SqlCommand

        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "DELETE FROM m_refRunning"
            SQL1.ExecuteReader()
            Conn.Close()

            Conn.Open()
            SQL2 = Conn.CreateCommand
            SQL2.CommandText = "DELETE FROM m_refRunningItem"
            SQL2.ExecuteReader()
            Conn.Close()

        Catch ex As Exception

        End Try
    End Sub
End Class