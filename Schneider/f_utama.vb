Imports System.IO
Imports System.IO.Ports
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports EasyModbus
Imports System.Text
Imports System.Threading
Imports RawPrint
'Imports IniParser
'Imports IniParser.Model
Imports System.ComponentModel

Public Class f_utama
    Dim modbusClient1 As ModbusClient
    Dim modbusClient2 As ModbusClient
    Dim modbusScrewing As ModbusClient
    'Dim modbusScrew As ModbusClient

    Private m_NumOfSeconds As Integer = 0

    Dim WithEvents Com1 As New SerialPort
    Dim Item1X As Integer = 0
    Dim Item1Y As Integer = 0
    Dim zplFile As String
    Dim zplString(9) As String
    Dim Item1Data As String = 0
    Dim Item1Type As String
    Dim Darkness As Integer
    Dim zebraPrinter As IPrinter = New Printer()

    Delegate Sub SetTextCallback(ByVal [text] As String)
    Dim LastQR As String

    Dim MW_RESET_BENCH_1 As Integer = 11500
    Dim MW_RESET_BENCH_2 As Integer = 11600
    Dim MW_RESET_BENCH_3 As Integer = 11700
    Dim MW_RESET_BENCH_4 As Integer = 11800

    Dim MW_COMPLETE_BENCH_2 As Integer = 12000
    Dim MW_COMPLETE_BENCH_3 As Integer = 12001
    Dim MW_COMPLETE_BENCH_4 As Integer = 12002
    'Dim rs As New Resizer
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Try
            lblTime.Text = Format(Now, "HH:mm:ss").ToString

        Catch ex As Exception

        End Try

        Try

            If modbusClient1.Connected And modbusClient2.Connected Then
                If modbusClient1.ReadHoldingRegisters(MW_RESET_BENCH_2, 1)(0) = 1 Then
                    RESET_BECNH_2()
                    CLEAR_CHECKLIST_ST2()
                    modbusClient1.WriteSingleRegister(MW_RESET_BENCH_2, 0)
                End If

                If modbusClient2.ReadHoldingRegisters(MW_RESET_BENCH_3, 1)(0) = 1 Then
                    RESET_BECNH_3()
                    CLEAR_CHECKLIST_ST3()
                    modbusClient2.WriteSingleRegister(MW_RESET_BENCH_3, 0)
                End If

                If modbusClient2.ReadHoldingRegisters(MW_RESET_BENCH_4, 1)(0) = 1 Then
                    RESET_BECNH_4()
                    CLEAR_CHECKLIST_ST4()
                    lblHMI4.Text = ""
                    FINISH_SCAN_QR = False
                    modbusClient2.WriteSingleRegister(MW_RESET_BENCH_4, 0)
                End If

            End If

        Catch ex As Exception

        End Try


    End Sub


    Private Sub CLEAR_CHECKLIST_ST2()
        modbusClient1.WriteSingleRegister(7010, 0)
        modbusClient1.WriteSingleRegister(7011, 0)
        modbusClient1.WriteSingleRegister(7012, 0)
        modbusClient1.WriteSingleRegister(7013, 0)
    End Sub

    Private Sub CLEAR_CHECKLIST_ST3()
        modbusClient2.WriteSingleRegister(7014, 0)
        modbusClient2.WriteSingleRegister(7015, 0)
        modbusClient2.WriteSingleRegister(7016, 0)
        modbusClient2.WriteSingleRegister(7017, 0)
        modbusClient2.WriteSingleRegister(7018, 0)
        modbusClient2.WriteSingleRegister(7019, 0)
        modbusClient2.WriteSingleRegister(7020, 0)

        modbusClient2.WriteSingleRegister(7032, 0)
        modbusClient2.WriteSingleRegister(7033, 0)

    End Sub

    Private Sub CLEAR_CHECKLIST_ST4()
        modbusClient2.WriteSingleRegister(5000, 0)
        modbusClient2.WriteSingleRegister(5001, 0)
        modbusClient2.WriteSingleRegister(5002, 0)

        modbusClient2.WriteSingleRegister(7021, 0)
        modbusClient2.WriteSingleRegister(7022, 0)
        modbusClient2.WriteSingleRegister(7023, 0)
        modbusClient2.WriteSingleRegister(7024, 0)



    End Sub




    Sub BacaLoadingUtama()
        Try

            tmrScan.Enabled = True

            'Baca Nomor ID dari PLC
            Dim StartAddress As Integer = 10001
            Dim ReadValue() As Integer = modbusClient1.ReadHoldingRegisters(StartAddress, 1)

            lblNoID.Text = ReadValue(0)

            Call PanggilRef()

            If lblRef.Text <> "" Then
                lblSequence.Text = "1"
                lblLabel.Text = "WAIT TO PROCESS"
                tmrSequence.Enabled = True
            End If

        Catch ex As Exception
            'MessageBox.Show(ex.Message.ToString)
        End Try
    End Sub
    Sub MasukPart1Bench1()

        Dim cb() As Byte
        'txtConvert.Text = ""
        If txtScan.Text <> "" Then
            cb = ASCIIEncoding.ASCII.GetBytes(txtScan.Text)
            Dim MWval_reff(39) As Integer 'jumlah address hingga 30
            For x As Integer = 0 To cb.Length - 1
                txtConvPart.AppendText(cb(x).ToString & "")
                'txtFrom.Text = ""
                'txtTo.AppendText(cb(x).ToString & "")
                'Dim arrayRef() as Char=cb(x).to
                'MessageBox.Show(cb(x).ToString)
                MWval_reff(x) = Convert.ToInt64(cb(x))
                Try
                    Dim StartAddress As Integer = 600 'mulai dari MW600
                    modbusClient1.WriteMultipleRegisters(StartAddress, MWval_reff)

                    'Call MasukPart1PPBench1()

                Catch ex As Exception

                End Try
            Next

        End If

    End Sub
    Sub MasukPart1PPBench1()

        Dim cb() As Byte
        'txtConvert.Text = ""
        If lblHMI1.Text <> "" Then
            cb = ASCIIEncoding.ASCII.GetBytes(lblHMI1.Text)
            Dim MWval_reff(39) As Integer 'jumlah address hingga 30
            For x As Integer = 0 To cb.Length - 1
                txtConvPart.AppendText(cb(x).ToString & "")
                'txtFrom.Text = ""
                'txtTo.AppendText(cb(x).ToString & "")
                'Dim arrayRef() as Char=cb(x).to
                'MessageBox.Show(cb(x).ToString)
                MWval_reff(x) = Convert.ToInt64(cb(x))
                Try
                    Dim StartAddress As Integer = Val(txtMW.Text)
                    modbusClient1.WriteMultipleRegisters(StartAddress, MWval_reff)
                Catch ex As Exception

                End Try
            Next

        End If

    End Sub
    Sub MasukPart2Bench1()

        Dim cb() As Byte
        'txtConvert.Text = ""
        If txtScan.Text <> "" Then
            cb = ASCIIEncoding.ASCII.GetBytes(txtScan.Text)
            Dim MWval_reff(39) As Integer 'jumlah address hingga 20
            For x As Integer = 0 To cb.Length - 1
                'txtConvert.AppendText(cb(x).ToString & "")
                'txtFrom.Text = ""
                'txtTo.AppendText(cb(x).ToString & "")
                'Dim arrayRef() as Char=cb(x).to
                'MessageBox.Show(cb(x).ToString)
                MWval_reff(x) = Convert.ToInt64(cb(x))
                Try
                    Dim StartAddress As Integer = 700 'mulai dari MW400
                    modbusClient1.WriteMultipleRegisters(StartAddress, MWval_reff)

                    'Call MasukPart2PPBench1()

                Catch ex As Exception

                End Try
            Next

        End If

    End Sub
    Sub MasukPart2PPBench1()

        Dim cb() As Byte
        'txtConvert.Text = ""
        If txtScan.Text <> "" Then
            cb = ASCIIEncoding.ASCII.GetBytes(txtScan.Text)
            Dim MWval_reff(39) As Integer 'jumlah address hingga 20
            For x As Integer = 0 To cb.Length - 1
                'txtConvert.AppendText(cb(x).ToString & "")
                'txtFrom.Text = ""
                'txtTo.AppendText(cb(x).ToString & "")
                'Dim arrayRef() as Char=cb(x).to
                'MessageBox.Show(cb(x).ToString)
                MWval_reff(x) = Convert.ToInt64(cb(x))
                Try
                    Dim StartAddress As Integer = Val(txtMW.Text) 'mulai dari MW400
                    modbusClient1.WriteMultipleRegisters(StartAddress, MWval_reff)

                Catch ex As Exception

                End Try
            Next

        End If

    End Sub

    Sub MasukPart3Bench1()

        Dim cb() As Byte
        'txtConvert.Text = ""
        If txtScan.Text <> "" Then
            cb = ASCIIEncoding.ASCII.GetBytes(txtScan.Text)
            Dim MWval_reff(39) As Integer 'jumlah address hingga 20
            For x As Integer = 0 To cb.Length - 1
                'txtConvert.AppendText(cb(x).ToString & "")
                'txtFrom.Text = ""
                'txtTo.AppendText(cb(x).ToString & "")
                'Dim arrayRef() as Char=cb(x).to
                'MessageBox.Show(cb(x).ToString)
                MWval_reff(x) = Convert.ToInt64(cb(x))
                Try
                    Dim StartAddress As Integer = 800 'mulai dari MW400
                    modbusClient1.WriteMultipleRegisters(StartAddress, MWval_reff)

                    'Call MasukPart3PPBench1()

                Catch ex As Exception

                End Try
            Next

        End If

    End Sub
    Sub MasukPart3PPBench1()

        Dim cb() As Byte
        'txtConvert.Text = ""
        If txtScan.Text <> "" Then
            cb = ASCIIEncoding.ASCII.GetBytes(txtScan.Text)
            Dim MWval_reff(39) As Integer 'jumlah address hingga 20
            For x As Integer = 0 To cb.Length - 1
                'txtConvert.AppendText(cb(x).ToString & "")
                'txtFrom.Text = ""
                'txtTo.AppendText(cb(x).ToString & "")
                'Dim arrayRef() as Char=cb(x).to
                'MessageBox.Show(cb(x).ToString)
                MWval_reff(x) = Convert.ToInt64(cb(x))
                Try
                    Dim StartAddress As Integer = Val(txtMW.Text) 'mulai dari MW400
                    modbusClient1.WriteMultipleRegisters(StartAddress, MWval_reff)
                Catch ex As Exception

                End Try
            Next

        End If

    End Sub
    Sub MasukPart4Bench1()

        Dim cb() As Byte
        'txtConvert.Text = ""
        If txtScan.Text <> "" Then
            cb = ASCIIEncoding.ASCII.GetBytes(txtScan.Text)
            Dim MWval_reff(39) As Integer 'jumlah address hingga 20
            For x As Integer = 0 To cb.Length - 1
                'txtConvert.AppendText(cb(x).ToString & "")
                'txtFrom.Text = ""
                'txtTo.AppendText(cb(x).ToString & "")
                'Dim arrayRef() as Char=cb(x).to
                'MessageBox.Show(cb(x).ToString)
                MWval_reff(x) = Convert.ToInt64(cb(x))
                Try
                    Dim StartAddress As Integer = 900 'mulai dari MW400
                    modbusClient1.WriteMultipleRegisters(StartAddress, MWval_reff)

                    'Call MasukPart4PPBench1()

                Catch ex As Exception

                End Try
            Next

        End If

    End Sub
    Sub MasukPart4PPBench1()

        Dim cb() As Byte
        'txtConvert.Text = ""
        If txtScan.Text <> "" Then
            cb = ASCIIEncoding.ASCII.GetBytes(txtScan.Text)
            Dim MWval_reff(39) As Integer 'jumlah address hingga 20
            For x As Integer = 0 To cb.Length - 1
                'txtConvert.AppendText(cb(x).ToString & "")
                'txtFrom.Text = ""
                'txtTo.AppendText(cb(x).ToString & "")
                'Dim arrayRef() as Char=cb(x).to
                'MessageBox.Show(cb(x).ToString)
                MWval_reff(x) = Convert.ToInt64(cb(x))
                Try
                    Dim StartAddress As Integer = Val(txtMW.Text) 'mulai dari MW400
                    modbusClient1.WriteMultipleRegisters(StartAddress, MWval_reff)
                Catch ex As Exception

                End Try
            Next

        End If

    End Sub
    Sub MasukPart5Bench1()

        Dim cb() As Byte
        'txtConvert.Text = ""
        If txtScan.Text <> "" Then
            cb = ASCIIEncoding.ASCII.GetBytes(txtScan.Text)
            Dim MWval_reff(39) As Integer 'jumlah address hingga 20
            For x As Integer = 0 To cb.Length - 1
                'txtConvert.AppendText(cb(x).ToString & "")
                'txtFrom.Text = ""
                'txtTo.AppendText(cb(x).ToString & "")
                'Dim arrayRef() as Char=cb(x).to
                'MessageBox.Show(cb(x).ToString)
                MWval_reff(x) = Convert.ToInt64(cb(x))
                Try
                    Dim StartAddress As Integer = 1000 'mulai dari MW400
                    modbusClient1.WriteMultipleRegisters(StartAddress, MWval_reff)

                    'Call MasukPart5PPBench1()

                Catch ex As Exception

                End Try
            Next

        End If

    End Sub
    Sub MasukPart5PPBench1()

        Dim cb() As Byte
        'txtConvert.Text = ""
        If txtScan.Text <> "" Then
            cb = ASCIIEncoding.ASCII.GetBytes(txtScan.Text)
            Dim MWval_reff(39) As Integer 'jumlah address hingga 20
            For x As Integer = 0 To cb.Length - 1
                'txtConvert.AppendText(cb(x).ToString & "")
                'txtFrom.Text = ""
                'txtTo.AppendText(cb(x).ToString & "")
                'Dim arrayRef() as Char=cb(x).to
                'MessageBox.Show(cb(x).ToString)
                MWval_reff(x) = Convert.ToInt64(cb(x))
                Try
                    Dim StartAddress As Integer = Val(txtMW.Text) 'mulai dari MW400
                    modbusClient1.WriteMultipleRegisters(StartAddress, MWval_reff)
                Catch ex As Exception

                End Try
            Next

        End If

    End Sub
    Sub MasukPart6Bench1()

        Dim cb() As Byte
        'txtConvert.Text = ""
        If txtScan.Text <> "" Then
            cb = ASCIIEncoding.ASCII.GetBytes(txtScan.Text)
            Dim MWval_reff(39) As Integer 'jumlah address hingga 20
            For x As Integer = 0 To cb.Length - 1
                'txtConvert.AppendText(cb(x).ToString & "")
                'txtFrom.Text = ""
                'txtTo.AppendText(cb(x).ToString & "")
                'Dim arrayRef() as Char=cb(x).to
                'MessageBox.Show(cb(x).ToString)
                MWval_reff(x) = Convert.ToInt64(cb(x))
                Try
                    Dim StartAddress As Integer = 1100 'mulai dari MW400
                    modbusClient1.WriteMultipleRegisters(StartAddress, MWval_reff)
                Catch ex As Exception

                End Try
            Next

        End If

    End Sub
    Sub MasukPart7Bench1()

        Dim cb() As Byte
        'txtConvert.Text = ""
        If txtScan.Text <> "" Then
            cb = ASCIIEncoding.ASCII.GetBytes(txtScan.Text)
            Dim MWval_reff(39) As Integer 'jumlah address hingga 20
            For x As Integer = 0 To cb.Length - 1
                'txtConvert.AppendText(cb(x).ToString & "")
                'txtFrom.Text = ""
                'txtTo.AppendText(cb(x).ToString & "")
                'Dim arrayRef() as Char=cb(x).to
                'MessageBox.Show(cb(x).ToString)
                MWval_reff(x) = Convert.ToInt64(cb(x))
                Try
                    Dim StartAddress As Integer = 1200 'mulai dari MW400
                    modbusClient1.WriteMultipleRegisters(StartAddress, MWval_reff)
                Catch ex As Exception

                End Try
            Next

        End If

    End Sub
    Sub MasukPart8Bench1()

        Dim cb() As Byte
        'txtConvert.Text = ""
        If txtScan.Text <> "" Then
            cb = ASCIIEncoding.ASCII.GetBytes(txtScan.Text)
            Dim MWval_reff(39) As Integer 'jumlah address hingga 20
            For x As Integer = 0 To cb.Length - 1
                'txtConvert.AppendText(cb(x).ToString & "")
                'txtFrom.Text = ""
                'txtTo.AppendText(cb(x).ToString & "")
                'Dim arrayRef() as Char=cb(x).to
                'MessageBox.Show(cb(x).ToString)
                MWval_reff(x) = Convert.ToInt64(cb(x))
                Try
                    Dim StartAddress As Integer = 1300 'mulai dari MW400
                    modbusClient1.WriteMultipleRegisters(StartAddress, MWval_reff)
                Catch ex As Exception

                End Try
            Next

        End If

    End Sub
    Sub MasukPart1Bench2()

        Dim cb() As Byte
        'txtConvert.Text = ""
        If lblHMI2.Text <> "" Then
            cb = ASCIIEncoding.ASCII.GetBytes(lblHMI2.Text)
            Dim MWval_reff(39) As Integer 'jumlah address hingga 30
            For x As Integer = 0 To cb.Length - 1
                txtConvPart.AppendText(cb(x).ToString & "")
                'txtFrom.Text = ""
                'txtTo.AppendText(cb(x).ToString & "")
                'Dim arrayRef() as Char=cb(x).to
                'MessageBox.Show(cb(x).ToString)
                MWval_reff(x) = Convert.ToInt64(cb(x))
                Try
                    Dim StartAddress As Integer = 1800 'mulai dari MW400
                    modbusClient2.WriteMultipleRegisters(StartAddress, MWval_reff)

                    'Call MasukPart1PPBench1()

                Catch ex As Exception

                End Try
            Next

        End If

    End Sub
    Sub MasukPart2Bench2()

        Dim cb() As Byte
        'txtConvert.Text = ""
        If lblHMI2.Text <> "" Then
            cb = ASCIIEncoding.ASCII.GetBytes(lblHMI2.Text)
            Dim MWval_reff(39) As Integer 'jumlah address hingga 30
            For x As Integer = 0 To cb.Length - 1
                'txtConvPart.AppendText(cb(x).ToString & "")
                'txtFrom.Text = ""
                'txtTo.AppendText(cb(x).ToString & "")
                'Dim arrayRef() as Char=cb(x).to
                'MessageBox.Show(cb(x).ToString)
                MWval_reff(x) = Convert.ToInt64(cb(x))
                Try
                    Dim StartAddress As Integer = 1900 'mulai dari MW400
                    modbusClient2.WriteMultipleRegisters(StartAddress, MWval_reff)

                    'Call MasukPart1PPBench1()

                Catch ex As Exception

                End Try
            Next

        End If

    End Sub
    Sub MasukPart3Bench2()

        Dim cb() As Byte
        'txtConvert.Text = ""
        If lblHMI2.Text <> "" Then
            cb = ASCIIEncoding.ASCII.GetBytes(lblHMI2.Text)
            Dim MWval_reff(39) As Integer 'jumlah address hingga 30
            For x As Integer = 0 To cb.Length - 1
                'txtConvPart.AppendText(cb(x).ToString & "")
                'txtFrom.Text = ""
                'txtTo.AppendText(cb(x).ToString & "")
                'Dim arrayRef() as Char=cb(x).to
                'MessageBox.Show(cb(x).ToString)
                MWval_reff(x) = Convert.ToInt64(cb(x))
                Try
                    Dim StartAddress As Integer = 2000 'mulai dari MW400
                    modbusClient2.WriteMultipleRegisters(StartAddress, MWval_reff)

                    'Call MasukPart1PPBench1()

                Catch ex As Exception

                End Try
            Next

        End If

    End Sub
    Sub MasukPart4Bench2()

        Dim cb() As Byte
        'txtConvert.Text = ""
        If lblHMI2.Text <> "" Then
            cb = ASCIIEncoding.ASCII.GetBytes(lblHMI2.Text)
            Dim MWval_reff(39) As Integer 'jumlah address hingga 30
            For x As Integer = 0 To cb.Length - 1
                'txtConvPart.AppendText(cb(x).ToString & "")
                'txtFrom.Text = ""
                'txtTo.AppendText(cb(x).ToString & "")
                'Dim arrayRef() as Char=cb(x).to
                'MessageBox.Show(cb(x).ToString)
                MWval_reff(x) = Convert.ToInt64(cb(x))
                Try
                    Dim StartAddress As Integer = 2100 'mulai dari MW400
                    modbusClient2.WriteMultipleRegisters(StartAddress, MWval_reff)

                    'Call MasukPart1PPBench1()

                Catch ex As Exception

                End Try
            Next

        End If

    End Sub
    Sub MasukPart5Bench2()

        Dim cb() As Byte
        'txtConvert.Text = ""
        If lblHMI2.Text <> "" Then
            cb = ASCIIEncoding.ASCII.GetBytes(lblHMI2.Text)
            Dim MWval_reff(39) As Integer 'jumlah address hingga 30
            For x As Integer = 0 To cb.Length - 1
                'txtConvPart.AppendText(cb(x).ToString & "")
                'txtFrom.Text = ""
                'txtTo.AppendText(cb(x).ToString & "")
                'Dim arrayRef() as Char=cb(x).to
                'MessageBox.Show(cb(x).ToString)
                MWval_reff(x) = Convert.ToInt64(cb(x))
                Try
                    Dim StartAddress As Integer = 2200 'mulai dari MW400
                    modbusClient2.WriteMultipleRegisters(StartAddress, MWval_reff)

                    'Call MasukPart1PPBench1()

                Catch ex As Exception

                End Try
            Next

        End If

    End Sub
    Sub MasukPart6Bench2()

        Dim cb() As Byte
        'txtConvert.Text = ""
        If lblHMI2.Text <> "" Then
            cb = ASCIIEncoding.ASCII.GetBytes(lblHMI2.Text)
            Dim MWval_reff(39) As Integer 'jumlah address hingga 30
            For x As Integer = 0 To cb.Length - 1
                'txtConvPart.AppendText(cb(x).ToString & "")
                'txtFrom.Text = ""
                'txtTo.AppendText(cb(x).ToString & "")
                'Dim arrayRef() as Char=cb(x).to
                'MessageBox.Show(cb(x).ToString)
                MWval_reff(x) = Convert.ToInt64(cb(x))
                Try
                    Dim StartAddress As Integer = 2300 'mulai dari MW400
                    modbusClient2.WriteMultipleRegisters(StartAddress, MWval_reff)

                    'Call MasukPart1PPBench1()

                Catch ex As Exception

                End Try
            Next

        End If

    End Sub
    Sub MasukPart7Bench2()

        Dim cb() As Byte
        'txtConvert.Text = ""
        If lblHMI2.Text <> "" Then
            cb = ASCIIEncoding.ASCII.GetBytes(lblHMI2.Text)
            Dim MWval_reff(39) As Integer 'jumlah address hingga 30
            For x As Integer = 0 To cb.Length - 1
                'txtConvPart.AppendText(cb(x).ToString & "")
                'txtFrom.Text = ""
                'txtTo.AppendText(cb(x).ToString & "")
                'Dim arrayRef() as Char=cb(x).to
                'MessageBox.Show(cb(x).ToString)
                MWval_reff(x) = Convert.ToInt64(cb(x))
                Try
                    Dim StartAddress As Integer = 2400 'mulai dari MW400
                    modbusClient2.WriteMultipleRegisters(StartAddress, MWval_reff)

                    'Call MasukPart1PPBench1()

                Catch ex As Exception

                End Try
            Next

        End If

    End Sub
    Sub MasukPart1Bench3()

        Dim cb() As Byte
        'txtConvert.Text = ""
        If lblHMI3.Text.Length > 5 Then
            cb = ASCIIEncoding.ASCII.GetBytes(lblHMI3.Text)
            Dim MWval_reff(39) As Integer 'jumlah address hingga 30
            For x As Integer = 0 To cb.Length - 1
                'txtConvPart.AppendText(cb(x).ToString & "")
                'txtFrom.Text = ""
                'txtTo.AppendText(cb(x).ToString & "")
                'Dim arrayRef() as Char=cb(x).to
                'MessageBox.Show(cb(x).ToString)
                MWval_reff(x) = Convert.ToInt64(cb(x))
                Try
                    Dim StartAddress As Integer = 2800 'mulai dari MW400
                    modbusClient2.WriteMultipleRegisters(StartAddress, MWval_reff)

                    'Call MasukPart1PPBench1()

                Catch ex As Exception

                End Try
            Next

        End If

    End Sub
    Sub MasukPart2Bench3()

        Dim cb() As Byte
        'txtConvert.Text = ""
        If lblHMI3.Text.Length > 5 Then
            cb = ASCIIEncoding.ASCII.GetBytes(lblHMI3.Text)
            Dim MWval_reff(39) As Integer 'jumlah address hingga 30
            For x As Integer = 0 To cb.Length - 1
                'txtConvPart.AppendText(cb(x).ToString & "")
                'txtFrom.Text = ""
                'txtTo.AppendText(cb(x).ToString & "")
                'Dim arrayRef() as Char=cb(x).to
                'MessageBox.Show(cb(x).ToString)
                MWval_reff(x) = Convert.ToInt64(cb(x))
                Try
                    Dim StartAddress As Integer = 2900 'mulai dari MW400
                    modbusClient2.WriteMultipleRegisters(StartAddress, MWval_reff)

                    'Call MasukPart1PPBench1()

                Catch ex As Exception

                End Try
            Next

        End If

    End Sub
    Sub MasukPart3Bench3()

        Dim cb() As Byte
        'txtConvert.Text = ""
        If lblHMI3.Text.Length > 5 Then
            cb = ASCIIEncoding.ASCII.GetBytes(lblHMI3.Text)
            Dim MWval_reff(39) As Integer 'jumlah address hingga 30
            For x As Integer = 0 To cb.Length - 1
                'txtConvPart.AppendText(cb(x).ToString & "")
                'txtFrom.Text = ""
                'txtTo.AppendText(cb(x).ToString & "")
                'Dim arrayRef() as Char=cb(x).to
                'MessageBox.Show(cb(x).ToString)
                MWval_reff(x) = Convert.ToInt64(cb(x))
                Try
                    Dim StartAddress As Integer = 3000 'mulai dari MW400
                    modbusClient2.WriteMultipleRegisters(StartAddress, MWval_reff)

                    'Call MasukPart1PPBench1()

                Catch ex As Exception

                End Try
            Next

        End If

    End Sub
    Sub MasukPart4Bench3()

        Dim cb() As Byte
        'txtConvert.Text = ""
        If lblHMI3.Text.Length > 5 Then
            cb = ASCIIEncoding.ASCII.GetBytes(lblHMI3.Text)
            Dim MWval_reff(39) As Integer 'jumlah address hingga 30
            For x As Integer = 0 To cb.Length - 1
                'txtConvPart.AppendText(cb(x).ToString & "")
                'txtFrom.Text = ""
                'txtTo.AppendText(cb(x).ToString & "")
                'Dim arrayRef() as Char=cb(x).to
                'MessageBox.Show(cb(x).ToString)
                MWval_reff(x) = Convert.ToInt64(cb(x))
                Try
                    Dim StartAddress As Integer = 3100 'mulai dari MW400
                    modbusClient2.WriteMultipleRegisters(StartAddress, MWval_reff)

                    'Call MasukPart1PPBench1()

                Catch ex As Exception

                End Try
            Next

        End If

    End Sub
    Sub MasukPart5Bench3()

        Dim cb() As Byte
        'txtConvert.Text = ""
        If lblHMI3.Text.Length > 6 Then
            cb = ASCIIEncoding.ASCII.GetBytes(lblHMI3.Text)
            Dim MWval_reff(39) As Integer 'jumlah address hingga 30
            For x As Integer = 0 To cb.Length - 1
                'txtConvPart.AppendText(cb(x).ToString & "")
                'txtFrom.Text = ""
                'txtTo.AppendText(cb(x).ToString & "")
                'Dim arrayRef() as Char=cb(x).to
                'MessageBox.Show(cb(x).ToString)
                MWval_reff(x) = Convert.ToInt64(cb(x))
                Try
                    Dim StartAddress As Integer = 3200 'mulai dari MW400
                    modbusClient2.WriteMultipleRegisters(StartAddress, MWval_reff)

                    'Call MasukPart1PPBench1()

                Catch ex As Exception

                End Try
            Next

        End If

    End Sub
    Sub MasukPart6Bench3()

        Dim cb() As Byte
        'txtConvert.Text = ""
        If lblHMI3.Text.Length > 5 Then
            cb = ASCIIEncoding.ASCII.GetBytes(lblHMI3.Text)
            Dim MWval_reff(39) As Integer 'jumlah address hingga 30
            For x As Integer = 0 To cb.Length - 1
                'txtConvPart.AppendText(cb(x).ToString & "")
                'txtFrom.Text = ""
                'txtTo.AppendText(cb(x).ToString & "")
                'Dim arrayRef() as Char=cb(x).to
                'MessageBox.Show(cb(x).ToString)
                MWval_reff(x) = Convert.ToInt64(cb(x))
                Try
                    Dim StartAddress As Integer = 3300 'mulai dari MW400
                    modbusClient2.WriteMultipleRegisters(StartAddress, MWval_reff)

                    'Call MasukPart1PPBench1()

                Catch ex As Exception

                End Try
            Next

        End If

    End Sub
    Sub MasukPart7Bench3()

        Dim cb() As Byte
        'txtConvert.Text = ""
        If lblHMI3.Text.Length > 5 Then
            cb = ASCIIEncoding.ASCII.GetBytes(lblHMI3.Text)
            Dim MWval_reff(39) As Integer 'jumlah address hingga 30
            For x As Integer = 0 To cb.Length - 1
                'txtConvPart.AppendText(cb(x).ToString & "")
                'txtFrom.Text = ""
                'txtTo.AppendText(cb(x).ToString & "")
                'Dim arrayRef() as Char=cb(x).to
                'MessageBox.Show(cb(x).ToString)
                MWval_reff(x) = Convert.ToInt64(cb(x))
                Try
                    Dim StartAddress As Integer = 3400 'mulai dari MW400
                    modbusClient2.WriteMultipleRegisters(StartAddress, MWval_reff)

                    'Call MasukPart1PPBench1()

                Catch ex As Exception

                End Try
            Next

        End If

    End Sub
    Sub MasukPart1Bench4()

        Dim cb() As Byte
        'txtConvert.Text = ""
        If lblHMI4.Text <> "" Then
            cb = ASCIIEncoding.ASCII.GetBytes(lblHMI4.Text)
            Dim MWval_reff(39) As Integer 'jumlah address hingga 30
            For x As Integer = 0 To cb.Length - 1
                'txtConvPart.AppendText(cb(x).ToString & "")
                'txtFrom.Text = ""
                'txtTo.AppendText(cb(x).ToString & "")
                'Dim arrayRef() as Char=cb(x).to
                'MessageBox.Show(cb(x).ToString)
                MWval_reff(x) = Convert.ToInt64(cb(x))
                Try
                    Dim StartAddress As Integer = 3800 'mulai dari MW400
                    modbusClient2.WriteMultipleRegisters(StartAddress, MWval_reff)

                    'Call MasukPart1PPBench1()

                Catch ex As Exception

                End Try
            Next

        End If

    End Sub
    Sub MasukPart2Bench4()

        Dim cb() As Byte
        'txtConvert.Text = ""
        If lblHMI4.Text <> "" Then
            cb = ASCIIEncoding.ASCII.GetBytes(lblHMI4.Text)
            Dim MWval_reff(39) As Integer 'jumlah address hingga 30
            For x As Integer = 0 To cb.Length - 1
                'txtConvPart.AppendText(cb(x).ToString & "")
                'txtFrom.Text = ""
                'txtTo.AppendText(cb(x).ToString & "")
                'Dim arrayRef() as Char=cb(x).to
                'MessageBox.Show(cb(x).ToString)
                MWval_reff(x) = Convert.ToInt64(cb(x))
                Try
                    Dim StartAddress As Integer = 3900 'mulai dari MW400
                    modbusClient2.WriteMultipleRegisters(StartAddress, MWval_reff)

                    'Call MasukPart1PPBench1()

                Catch ex As Exception

                End Try
            Next

        End If

    End Sub
    Sub MasukPart3Bench4()

        Dim cb() As Byte
        'txtConvert.Text = ""
        If lblHMI4.Text <> "" Then
            cb = ASCIIEncoding.ASCII.GetBytes(lblHMI4.Text)
            Dim MWval_reff(39) As Integer 'jumlah address hingga 30
            For x As Integer = 0 To cb.Length - 1
                'txtConvPart.AppendText(cb(x).ToString & "")
                'txtFrom.Text = ""
                'txtTo.AppendText(cb(x).ToString & "")
                'Dim arrayRef() as Char=cb(x).to
                'MessageBox.Show(cb(x).ToString)
                MWval_reff(x) = Convert.ToInt64(cb(x))
                Try
                    Dim StartAddress As Integer = 4000 'mulai dari MW400
                    modbusClient2.WriteMultipleRegisters(StartAddress, MWval_reff)

                    'Call MasukPart1PPBench1()

                Catch ex As Exception

                End Try
            Next

        End If

    End Sub
    Sub MasukPart4Bench4()

        Dim cb() As Byte
        'txtConvert.Text = ""
        If lblHMI4.Text <> "" Then
            cb = ASCIIEncoding.ASCII.GetBytes(lblHMI4.Text)
            Dim MWval_reff(39) As Integer 'jumlah address hingga 30
            For x As Integer = 0 To cb.Length - 1
                'txtConvPart.AppendText(cb(x).ToString & "")
                'txtFrom.Text = ""
                'txtTo.AppendText(cb(x).ToString & "")
                'Dim arrayRef() as Char=cb(x).to
                'MessageBox.Show(cb(x).ToString)
                MWval_reff(x) = Convert.ToInt64(cb(x))
                Try
                    Dim StartAddress As Integer = 4100 'mulai dari MW400
                    modbusClient2.WriteMultipleRegisters(StartAddress, MWval_reff)

                    'Call MasukPart1PPBench1()

                Catch ex As Exception

                End Try
            Next

        End If

    End Sub
    Sub MasukPart5Bench4()

        Dim cb() As Byte
        'txtConvert.Text = ""
        If lblHMI4.Text <> "" Then
            cb = ASCIIEncoding.ASCII.GetBytes(lblHMI4.Text)
            Dim MWval_reff(39) As Integer 'jumlah address hingga 30
            For x As Integer = 0 To cb.Length - 1
                'txtConvPart.AppendText(cb(x).ToString & "")
                'txtFrom.Text = ""
                'txtTo.AppendText(cb(x).ToString & "")
                'Dim arrayRef() as Char=cb(x).to
                'MessageBox.Show(cb(x).ToString)
                MWval_reff(x) = Convert.ToInt64(cb(x))
                Try
                    Dim StartAddress As Integer = 4200 'mulai dari MW400
                    modbusClient2.WriteMultipleRegisters(StartAddress, MWval_reff)

                    'Call MasukPart1PPBench1()

                Catch ex As Exception

                End Try
            Next

        End If

    End Sub
    Sub MasukPart6Bench4()

        Dim cb() As Byte
        'txtConvert.Text = ""
        If lblHMI4.Text <> "" Then
            cb = ASCIIEncoding.ASCII.GetBytes(lblHMI4.Text)
            Dim MWval_reff(39) As Integer 'jumlah address hingga 30
            For x As Integer = 0 To cb.Length - 1
                'txtConvPart.AppendText(cb(x).ToString & "")
                'txtFrom.Text = ""
                'txtTo.AppendText(cb(x).ToString & "")
                'Dim arrayRef() as Char=cb(x).to
                'MessageBox.Show(cb(x).ToString)
                MWval_reff(x) = Convert.ToInt64(cb(x))
                Try
                    Dim StartAddress As Integer = 4300 'mulai dari MW400
                    modbusClient2.WriteMultipleRegisters(StartAddress, MWval_reff)

                    'Call MasukPart1PPBench1()

                Catch ex As Exception

                End Try
            Next

        End If

    End Sub
    Sub ConnectPLC1()
        Try
            Dim konek As Integer

            'SEKA

            modbusClient1 = New ModbusClient(lblPLCip1.Text, 502)

            If Not modbusClient1.Connected Then
                modbusClient1.Connect()
            End If


            konek = 1
            'lblConStatus.Text = "Connected"

            If konek = 1 Then
                'lblPLCStatus.Text = "ON"
                btmPLC1.BackColor = Color.GreenYellow
                btmPLC1.Text = "ON"
                btmPLC1.ForeColor = Color.Red
                lblLabel.Text = "INPUT NEW REFERENCE TO RUN"
            Else
                btmPLC1.BackColor = Color.Red
                btmPLC1.Text = "OFF"
                btmPLC1.ForeColor = Color.Black
                lblLabel.Text = "CHECK PLC CONNECTION!"
                'Call BacaHoldingRegister()
                'Else
                Exit Sub
            End If

        Catch ex As Exception
            MsgBox("Error Connected to PLC! " & ex.Message)
            btmPLC1.BackColor = Color.Red
            btmPLC1.Text = "OFF"
            btmPLC1.ForeColor = Color.Black

        End Try
    End Sub
    Sub ConnectScrewing1()
        Try
            modbusScrewing = New ModbusClient("192.168.1.102", 502)

            If Not modbusScrewing.Connected Then
                modbusScrewing.Connect()
            End If
        Catch ex As Exception
            MessageBox.Show("PLC Screw 102 Not Connected", "PLC Screwing", MessageBoxButtons.OK)
            End
        End Try

    End Sub
    Sub ConnectPLC2()
        Try
            Dim konek As Integer
            'SEKA
            modbusClient2 = New ModbusClient(lblPLCip2.Text, 502)

            If Not modbusClient2.Connected Then
                modbusClient2.Connect()
            End If



            konek = 1
            'lblConStatus.Text = "Connected"

            If konek = 1 Then
                'lblPLCStatus.Text = "ON"
                btmPLC2.BackColor = Color.GreenYellow
                btmPLC2.Text = "ON"
                btmPLC2.ForeColor = Color.Red
            Else
                btmPLC2.BackColor = Color.Red
                btmPLC2.Text = "OFF"
                btmPLC2.ForeColor = Color.Black
                'Call BacaHoldingRegister()
                'Else
                Exit Sub
            End If

        Catch ex As Exception
            MsgBox("Error Connected to PLC! " & ex.Message)
            btmPLC2.BackColor = Color.Red
            btmPLC2.Text = "OFF"
            btmPLC2.ForeColor = Color.Black

        End Try
    End Sub
    Sub BacaLine()
        Dim TW As New StreamReader("Line.ini")
        Dim koneksi As String
        koneksi = TW.ReadLine
        Dim NomorLine As String
        NomorLine = koneksi

        lblStation.Text = NomorLine

        'conn = New SqlConnection(koneksi)
        'conn = New SqlConnection("server = QONOD_1;database = Timbangan;Trusted_Connection = yes")

    End Sub
    Private Sub f_utama_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Call Config()
            lblDetik.Text = "0"
            lblPrint.Text = "0"
            lblLabel2.Text = ""

            '1280, 728
            ResizeFormClass.SubResize(Me, (Me.Width / 1280) * 100, (Me.Height / 720) * 100)
            'rs.FindAllControls(Me)
            lblDate.Text = Format(Now, "dd/mm/yyy").ToString
            'lblLabel.Text = "Initializing..."

            Call ConnectPLC1()
            Call ConnectPLC2()
            Call ConnectScrewing1()


            'SEKA
            RESET_BECNH_4()
            RESET_BECNH_2()
            RESET_BECNH_3()

            Call BacaLine()

            lblPCB.Text = ""
            'Call BersihkanMWRunning()

            Call BacaLoadingUtama()
            Call BACA_SIGNAL()

            'Call BACA_MATERIAL_COMPLETE()

            'lblPCB.Text = ""
            'tmrPCBA.Enabled = True

            tmStation2.Enabled = True


            txtScan.Focus()

            SerialBench2.PortName = "COM4" 'txtCom.Text
            SerialBench2.BaudRate = "9600" 'txtBaudRate.Text
            SerialBench2.Parity = IO.Ports.Parity.None
            SerialBench2.StopBits = IO.Ports.StopBits.One
            SerialBench2.DataBits = 8            'Open our serial port
            SerialBench2.Open()

            SerialBench3.PortName = "COM3" 'txtCom.Text
            SerialBench3.BaudRate = "9600" 'txtBaudRate.Text
            SerialBench3.Parity = IO.Ports.Parity.None
            SerialBench3.StopBits = IO.Ports.StopBits.One
            SerialBench3.DataBits = 8            'Open our serial port
            SerialBench3.Open()

            SerialBench4.PortName = "COM5" 'txtCom.Text
            SerialBench4.BaudRate = "9600" 'txtBaudRate.Text
            SerialBench4.Parity = IO.Ports.Parity.None
            SerialBench4.StopBits = IO.Ports.StopBits.One
            SerialBench4.DataBits = 8            'Open our serial port
            SerialBench4.Open()


        Catch ex As Exception

        End Try
    End Sub
    Private Sub BACA_SIGNAL()
        'Baca Reference yg sedang terbuka
        Dim StartAddressref As Integer = 10001
        Dim ReadValueref() As Integer = modbusClient1.ReadHoldingRegisters(StartAddressref, 1)
        lblNoID.Text = ReadValueref(0)

        Dim StartAddress As Integer = 10002
        Dim ReadValue() As Integer = modbusClient1.ReadHoldingRegisters(StartAddress, 1)
        lblMWStation1.Text = ReadValue(0)

        Dim StartAddress3 As Integer = 10003
        Dim ReadValue3() As Integer = modbusClient1.ReadHoldingRegisters(StartAddress3, 1)
        lblMWStation2.Text = ReadValue3(0)

        If lblMWStation2.Text = "1" Then
            'Call PanggilRef()
            Call BACA_MATERIAL_COMPLETE()

            'Tutup Signal dari MW Station 1
            Dim StartAddress1 As Integer = 10003
            Dim values(0) As Integer
            values(0) = 0
            modbusClient1.WriteMultipleRegisters(StartAddress1, values)
            'End If
        End If
    End Sub

    Sub BersihkanMWRunning()
        '===============MW8000 (PLC 106)
        Try
            Dim StartAddress As Integer = 8000
            Dim Values(0) As Integer
            Values(0) = 0
            modbusClient1.WriteMultipleRegisters(StartAddress, Values)

            '============MW8001,8002,8003 (PLC 107)
            Dim StartAddress8001 As Integer = 8001
            Dim Values8001(3) As Integer
            Values8001(3) = 0
            modbusClient2.WriteMultipleRegisters(StartAddress8001, Values8001)


            '============= 9000 (PLC)
            Dim StartAddress2 As Integer = 9000
            Dim Values2(0) As Integer
            Values2(0) = 0
            modbusClient1.WriteMultipleRegisters(StartAddress2, Values2)

            '============= 9001,9002,9003 (PLC 107)
            Dim StartAddres9001 As Integer = 9001
            Dim Values9001(3) As Integer
            Values9001(3) = 0
            modbusClient1.WriteMultipleRegisters(StartAddres9001, Values9001)


            '============= 7000 - 7024
            Dim StartAddres3 As Integer = 7000
            Dim Values3(25) As Integer
            Values3(25) = 0
            modbusClient1.WriteMultipleRegisters(StartAddres3, Values3)
            '=============

            Dim StartAddress4 As Integer = 106
            Dim MW106(0) As Integer
            MW106(0) = 0
            modbusClient1.WriteMultipleRegisters(StartAddress4, MW106)

            Dim StartAddress5 As Integer = 107
            Dim MW107(0) As Integer
            MW107(0) = 0
            modbusClient2.WriteMultipleRegisters(StartAddress5, MW107)

        Catch ex As Exception

        End Try
    End Sub
    Sub MatikanStatusPrint()
        Try
            'Matikan Status Print 9000
            Dim StartAddress As Integer = 9000
            Dim Values(0) As Integer
            Values(0) = 0
            modbusClient1.WriteMultipleRegisters(StartAddress, Values)
        Catch ex As Exception

        End Try

    End Sub
    Sub PanggilRef()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "SELECT *FROM m_refrunning WHERE (NoID=" & Val(lblNoID.Text) & ")"
            Reader = SQL1.ExecuteReader
            If Reader.Read Then
                'lblNoID.Text = Val(Reader("NoID"))
                lblSO.Text = Reader("SONo").ToString
                lblRef.Text = Reader("Model").ToString
                lblItem.Text = Reader("ItemNo").ToString
                lblQty.Text = Val(Reader("Qty")) '.ToString
                lblTarget.Text = Val(Reader("Qty"))
                lblRunQty.Text = Val(Reader("RunQty"))
                lblVolt.Text = Reader("Voltage").ToString
                lblLayer.Text = Reader("Layer").ToString
                lblBuzzer.Text = Reader("Buzzer").ToString
                lblFoot.Text = Reader("Foot").ToString
                lblPCB.Text = Reader("PCBA").ToString
                lblType.Text = Reader("Type").ToString


                'Call TotalPartSt1()
                Call dpStation1()
                Call dpStation2()
                Call dpStation3()
                Call dpStation4()

                Call HitungRunQty()

            Else
                'MessageBox.Show("Data tidak ditemukan!", "Error", MessageBoxButtons.OK)
            End If
            Reader.Close()
            Conn.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Panggil", MessageBoxButtons.OK)
        End Try
    End Sub
    Sub TotalSt1Masterial()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "SELECT COUNT(PartNo) AS tPart FROM m_refPartList WHERE (Model='" & lblRef.Text & "')AND(Station=1)"
            Reader = SQL1.ExecuteReader
            If Reader.Read Then
                lblSt1TotalMaterial.Text = Val(Reader("tPart"))
            Else
                lblSt1TotalMaterial.Text = "00"
            End If
            Reader.Close()
            Conn.Close()

        Catch ex As Exception
            lblSt1TotalMaterial.Text = "00"
        End Try
    End Sub
    Sub TotalSt2Masterial()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "SELECT COUNT(PartNo) AS tPart FROM m_refPartList WHERE (Model='" & lblRef.Text & "')AND(Station=2)"
            Reader = SQL1.ExecuteReader
            If Reader.Read Then
                lblSt2TotalMaterial.Text = Val(Reader("tPart"))
            Else
                lblSt2TotalMaterial.Text = "00"
            End If
            Reader.Close()
            Conn.Close()

        Catch ex As Exception
            lblSt2TotalMaterial.Text = "00"
        End Try
    End Sub
    Sub TotalSt3Masterial()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "SELECT COUNT(PartNo) AS tPart FROM m_refPartList WHERE (Model='" & lblRef.Text & "')AND(Station=3)"
            Reader = SQL1.ExecuteReader
            If Reader.Read Then
                lblSt3TotalMaterial.Text = Val(Reader("tPart"))
            Else
                lblSt3TotalMaterial.Text = "00"
            End If
            Reader.Close()
            Conn.Close()

        Catch ex As Exception
            lblSt2TotalMaterial.Text = "00"
        End Try
    End Sub
    Sub TotalSt4Masterial()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "SELECT COUNT(PartNo) AS tPart FROM m_refPartList WHERE (Model='" & lblRef.Text & "')AND(Station=4)"
            Reader = SQL1.ExecuteReader
            If Reader.Read Then
                lblSt4TotalMaterial.Text = Val(Reader("tPart"))
            Else
                lblSt4TotalMaterial.Text = "00"
            End If
            Reader.Close()
            Conn.Close()

        Catch ex As Exception
            lblSt2TotalMaterial.Text = "00"
        End Try
    End Sub
    Sub St1MasterialScan()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "SELECT COUNT(PartNo) AS tPart FROM m_refPartList WHERE (Model='" & lblRef.Text & "')AND(Station=1)AND(Scan=1)"
            Reader = SQL1.ExecuteReader

            If Reader.Read Then
                lblSt1ScanTotal.Text = Val(Reader("tPart"))
            Else
                lblSt1ScanTotal.Text = "00"
            End If

            Reader.Close()
            Conn.Close()

        Catch ex As Exception
            lblSt1ScanTotal.Text = "00"
        End Try
    End Sub
    Sub St2MasterialScan()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "SELECT COUNT(PartNo) AS tPart FROM m_refPartList WHERE (Model='" & lblRef.Text & "')AND(Station=2)AND(Scan=1)"
            Reader = SQL1.ExecuteReader
            If Reader.Read Then
                lblSt2ScanTotal.Text = Val(Reader("tPart"))
            Else
                lblSt2ScanTotal.Text = "00"
            End If
            Reader.Close()
            Conn.Close()

        Catch ex As Exception
            lblSt2ScanTotal.Text = "00"
        End Try
    End Sub
    Sub St3MasterialScan()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "SELECT COUNT(PartNo) AS tPart FROM m_refPartList WHERE (Model='" & lblRef.Text & "')AND(Station=3)AND(Scan=1)"
            Reader = SQL1.ExecuteReader
            If Reader.Read Then
                lblSt3ScanTotal.Text = Val(Reader("tPart"))
            Else
                lblSt3ScanTotal.Text = "00"
            End If
            Reader.Close()
            Conn.Close()

        Catch ex As Exception
            lblSt2ScanTotal.Text = "00"
        End Try
    End Sub
    Sub St4MasterialScan()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "SELECT COUNT(PartNo) AS tPart FROM m_refPartList WHERE (Model='" & lblRef.Text & "')AND(Station=4)AND(Scan=1)"
            Reader = SQL1.ExecuteReader
            If Reader.Read Then
                lblSt4ScanTotal.Text = Val(Reader("tPart"))
            Else
                lblSt4ScanTotal.Text = "00"
            End If
            Reader.Close()
            Conn.Close()

        Catch ex As Exception
            lblSt2ScanTotal.Text = "00"
        End Try
    End Sub
    Sub dpStation1()
        Try
            Dim conn As SqlConnection
            Dim SQL1 As New SqlCommand
            Dim Pilih As New SqlDataAdapter
            Dim datWarga As New DataSet
            Dim datTable As New DataTable

            conn = GetConnect()
            conn.Open()

            SQL1 = conn.CreateCommand
            SQL1.CommandText = "SELECT *FROM m_refPartList WHERE (Model='" & lblRef.Text & "')AND(Station=1) ORDER BY SEQ ASC"
            Pilih.SelectCommand = SQL1
            Pilih.Fill(datWarga, "m_refPartList")
            datTable = datWarga.Tables("m_refPartList")

            Dim i As Integer
            Dim x As Integer

            'Displaydata()
            LvSt1.Items.Clear()
            For i = 0 To (datTable.Rows.Count - 1)
                With datTable.Rows(i)
                    Dim lSingleItem As ListViewItem
                    'Dim Tot = FormatCurrency("Amount")
                    lSingleItem = LvSt1.Items.Add(.Item("Model").ToString)
                    lSingleItem.SubItems.Add(.Item("Item").ToString)
                    lSingleItem.SubItems.Add(.Item("PartNo").ToString)
                    lSingleItem.SubItems.Add(.Item("Scan").ToString)
                    lSingleItem.SubItems.Add(.Item("PokNo").ToString)

                    'lSingleItem.SubItems.Add(.Item("RefNo").ToString)
                    'lSingleItem.SubItems.Add(.Item("ItemNo").ToString)
                    'lSingleItem.SubItems(3).Text = CDbl(lSingleItem.SubItems(3).Text).ToString("N2")
                    'lSingleItem.SubItems.Add(.Item("Qty").ToString)
                    'lSingleItem.SubItems(4).Text = CDbl(lSingleItem.SubItems(4).Text).ToString("N2")
                    'lSingleItem.SubItems.Add(.Item("Type").ToString)
                End With

                For x = 0 To LvSt1.Items.Count - 1
                    If LvSt1.Items(x).SubItems(3).Text = "1" Then
                        LvSt1.Items(x).BackColor = Color.Yellow
                    Else
                        'LvSt1.Items(x).BackColor = Color.White
                    End If
                Next x

            Next i

            Call TotalSt1Masterial()
            Call St1MasterialScan()

            Call CekScanStatusSt1()

            conn.Close()
        Catch ex As Exception

        End Try
    End Sub
    Sub dpStation2()
        Try
            Dim conn As SqlConnection
            Dim SQL1 As New SqlCommand
            Dim Pilih As New SqlDataAdapter
            Dim datWarga As New DataSet
            Dim datTable As New DataTable

            conn = GetConnect()
            conn.Open()

            SQL1 = conn.CreateCommand
            SQL1.CommandText = "SELECT *FROM m_refPartList WHERE (Model='" & lblRef.Text & "')AND(Station=2) ORDER BY SEQ ASC"
            Pilih.SelectCommand = SQL1
            Pilih.Fill(datWarga, "m_refPartList")
            datTable = datWarga.Tables("m_refPartList")

            Dim i As Integer
            Dim x As Integer

            'Displaydata()
            LvSt2.Items.Clear()
            For i = 0 To (datTable.Rows.Count - 1)
                With datTable.Rows(i)
                    Dim lSingleItem As ListViewItem
                    'Dim Tot = FormatCurrency("Amount")
                    lSingleItem = LvSt2.Items.Add(.Item("Model").ToString)
                    lSingleItem.SubItems.Add(.Item("Item").ToString)
                    lSingleItem.SubItems.Add(.Item("PartNo").ToString)
                    lSingleItem.SubItems.Add(.Item("Scan").ToString)
                    lSingleItem.SubItems.Add(.Item("PokNo").ToString)

                    'lSingleItem.SubItems.Add(.Item("RefNo").ToString)
                    'lSingleItem.SubItems.Add(.Item("ItemNo").ToString)
                    'lSingleItem.SubItems(3).Text = CDbl(lSingleItem.SubItems(3).Text).ToString("N2")
                    'lSingleItem.SubItems.Add(.Item("Qty").ToString)
                    'lSingleItem.SubItems(4).Text = CDbl(lSingleItem.SubItems(4).Text).ToString("N2")
                    'lSingleItem.SubItems.Add(.Item("Type").ToString)
                End With

                For x = 0 To LvSt2.Items.Count - 1
                    If LvSt2.Items(x).SubItems(3).Text = "1" Then
                        LvSt2.Items(x).BackColor = Color.Yellow
                    Else
                        'LvSt1.Items(x).BackColor = Color.White
                    End If
                Next x

            Next i

            Call HapusStatusPartBench2()

            Call TotalSt2Masterial()
            Call St2MasterialScan()

            Call CekScanStatusSt2()

            'Call TotalOrder()

            'Call TotalItem()

            conn.Close()
        Catch ex As Exception

        End Try
    End Sub
    Sub dpStation3()
        Try
            Dim conn As SqlConnection
            Dim SQL1 As New SqlCommand
            Dim Pilih As New SqlDataAdapter
            Dim datWarga As New DataSet
            Dim datTable As New DataTable

            conn = GetConnect()
            conn.Open()

            SQL1 = conn.CreateCommand
            SQL1.CommandText = "SELECT *FROM m_refPartList WHERE (Model='" & lblRef.Text & "')AND(Station=3) ORDER BY SEQ ASC"
            Pilih.SelectCommand = SQL1
            Pilih.Fill(datWarga, "m_refPartList")
            datTable = datWarga.Tables("m_refPartList")

            Dim i As Integer
            Dim x As Integer

            'Displaydata()
            LvSt3.Items.Clear()
            For i = 0 To (datTable.Rows.Count - 1)
                With datTable.Rows(i)
                    Dim lSingleItem As ListViewItem
                    'Dim Tot = FormatCurrency("Amount")
                    lSingleItem = LvSt3.Items.Add(.Item("Model").ToString)
                    lSingleItem.SubItems.Add(.Item("Item").ToString)
                    lSingleItem.SubItems.Add(.Item("PartNo").ToString)
                    lSingleItem.SubItems.Add(.Item("Scan").ToString)
                    lSingleItem.SubItems.Add(.Item("PokNo").ToString)

                    'lSingleItem.SubItems.Add(.Item("RefNo").ToString)
                    'lSingleItem.SubItems.Add(.Item("ItemNo").ToString)
                    'lSingleItem.SubItems(3).Text = CDbl(lSingleItem.SubItems(3).Text).ToString("N2")
                    'lSingleItem.SubItems.Add(.Item("Qty").ToString)
                    'lSingleItem.SubItems(4).Text = CDbl(lSingleItem.SubItems(4).Text).ToString("N2")
                    'lSingleItem.SubItems.Add(.Item("Type").ToString)
                End With

                For x = 0 To LvSt3.Items.Count - 1
                    If LvSt3.Items(x).SubItems(3).Text = "1" Then
                        LvSt3.Items(x).BackColor = Color.Yellow
                    Else
                        'LvSt1.Items(x).BackColor = Color.White
                    End If
                Next x

            Next i
            Call HapusStatusPartBench3()

            Call TotalSt3Masterial()
            Call St3MasterialScan()

            Call CekScanStatusSt3()

            'Call TotalOrder()

            'Call TotalItem()

            conn.Close()
        Catch ex As Exception

        End Try
    End Sub
    Sub dpStation4()
        Try
            Dim conn As SqlConnection
            Dim SQL1 As New SqlCommand
            Dim Pilih As New SqlDataAdapter
            Dim datWarga As New DataSet
            Dim datTable As New DataTable

            conn = GetConnect()
            conn.Open()

            SQL1 = conn.CreateCommand
            SQL1.CommandText = "SELECT *FROM m_refPartList WHERE (Model='" & lblRef.Text & "')AND(Station=4) ORDER BY Item ASC"
            Pilih.SelectCommand = SQL1
            Pilih.Fill(datWarga, "m_refPartList")
            datTable = datWarga.Tables("m_refPartList")

            Dim i As Integer
            Dim x As Integer

            'Displaydata()
            LvSt4.Items.Clear()
            For i = 0 To (datTable.Rows.Count - 1)
                With datTable.Rows(i)
                    Dim lSingleItem As ListViewItem
                    'Dim Tot = FormatCurrency("Amount")
                    lSingleItem = LvSt4.Items.Add(.Item("Model").ToString)
                    lSingleItem.SubItems.Add(.Item("Item").ToString)
                    lSingleItem.SubItems.Add(.Item("PartNo").ToString)
                    lSingleItem.SubItems.Add(.Item("Scan").ToString)
                    lSingleItem.SubItems.Add(.Item("PokNo").ToString)

                    'lSingleItem.SubItems.Add(.Item("RefNo").ToString)
                    'lSingleItem.SubItems.Add(.Item("ItemNo").ToString)
                    'lSingleItem.SubItems(3).Text = CDbl(lSingleItem.SubItems(3).Text).ToString("N2")
                    'lSingleItem.SubItems.Add(.Item("Qty").ToString)
                    'lSingleItem.SubItems(4).Text = CDbl(lSingleItem.SubItems(4).Text).ToString("N2")
                    'lSingleItem.SubItems.Add(.Item("Type").ToString)
                End With

                For x = 0 To LvSt4.Items.Count - 1
                    If LvSt4.Items(x).SubItems(3).Text = "1" Then
                        LvSt4.Items(x).BackColor = Color.Yellow
                    Else
                        'LvSt1.Items(x).BackColor = Color.White
                    End If
                Next x

            Next i
            Call HapusStatusPartBench4()

            Call TotalSt4Masterial()
            Call St4MasterialScan()

            Call CekScanStatusSt4()

            'Call TotalOrder()

            'Call TotalItem()

            conn.Close()
        Catch ex As Exception

        End Try
    End Sub

    Sub MasukQtyRunningMWBench2()

        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader
        Dim Jumlah As Integer

        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "SELECT SUM(QtyRunSt2) AS RunST2 FROM m_refRunningItem WHERE ( SONO = '" & lblSO.Text & "')"
            Reader = SQL1.ExecuteReader
            If Reader.Read Then
                Jumlah = Val(Reader("RunST2"))

                'Kirim ke HMI Bench 2
                Dim Start311 As Integer = 311
                Dim Val311(0) As Integer
                Val311(0) = Jumlah
                modbusClient1.WriteMultipleRegisters(Start311, Val311)

            Else
                Jumlah = 0

            End If
            Conn.Close()
            Reader.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message,"Error Qty",MessageBoxButtons.OK)
        End Try
    End Sub
    Sub MasukQtyRunningMWBench3()

        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader
        Dim Jumlah As Integer

        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "SELECT SUM(QtyRunSt3) AS RunST3 FROM m_refRunningItem WHERE ( SONO ='" & lblSO.Text & "')"
            Reader = SQL1.ExecuteReader
            If Reader.Read Then
                Jumlah = Val(Reader("RunST3"))

                'Kirim ke HMI Bench 4
                Dim StartAddress As Integer = 312
                Dim Values(0) As Integer
                Values(0) = Jumlah
                modbusClient2.WriteMultipleRegisters(StartAddress, Values)

            Else
                Jumlah = 0

            End If
            Conn.Close()
            Reader.Close()

        Catch ex As Exception

        End Try

    End Sub
    Sub MasukQtyRunningMWBench4()

        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader
        Dim Jumlah As Integer

        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "SELECT SUM(QtyRunSt4) AS RunST4 FROM m_refRunningItem WHERE (SONO = '" & lblSO.Text & "')"
            Reader = SQL1.ExecuteReader
            If Reader.Read Then
                Jumlah = Val(Reader("RunST4"))
                lblRunQty.Text = Jumlah

                'Kirim ke HMI Bench 4
                Dim StartAddress As Integer = 313
                Dim Values(0) As Integer
                Values(0) = Jumlah
                modbusClient2.WriteMultipleRegisters(StartAddress, Values)
            Else
                Jumlah = 0

            End If
            Conn.Close()
            Reader.Close()

        Catch ex As Exception

        End Try

    End Sub
    Sub MasukPartListHMI()
        Try
            If txtStation.Text = "1" And txtStationItem.Text = "1" Then
                Call MasukPart1Bench1()
            ElseIf txtStation.Text = "1" And txtStationItem.Text = "2" Then
                Call MasukPart2Bench1()
            ElseIf txtStation.Text = "1" And txtStationItem.Text = "3" Then
                Call MasukPart3Bench1()
            ElseIf txtStation.Text = "1" And txtStationItem.Text = "4" Then
                Call MasukPart4Bench1()
            ElseIf txtStation.Text = "1" And txtStationItem.Text = "5" Then
                Call MasukPart5Bench1()
            ElseIf txtStation.Text = "1" And txtStationItem.Text = "6" Then
                Call MasukPart6Bench1()
            ElseIf txtStation.Text = "1" And txtStationItem.Text = "7" Then
                Call MasukPart7Bench1()
            ElseIf txtStation.Text = "1" And txtStationItem.Text = "8" Then
                Call MasukPart8Bench1()
            ElseIf txtStation.Text = "2" And txtStationItem.Text = "1" Then
                Call MasukPart1Bench2()
            ElseIf txtStation.Text = "2" And txtStationItem.Text = "2" Then
                Call MasukPart2Bench2()
            ElseIf txtStation.Text = "2" And txtStationItem.Text = "3" Then
                Call MasukPart3Bench2()
            ElseIf txtStation.Text = "2" And txtStationItem.Text = "4" Then
                Call MasukPart4Bench2()
            ElseIf txtStation.Text = "2" And txtStationItem.Text = "5" Then
                Call MasukPart5Bench2()
            ElseIf txtStation.Text = "2" And txtStationItem.Text = "6" Then
                Call MasukPart6Bench2()
            ElseIf txtStation.Text = "2" And txtStationItem.Text = "7" Then
                Call MasukPart7Bench2()
            ElseIf txtStation.Text = "3" And txtStationItem.Text = "1" Then
                Call MasukPart1Bench3()
            ElseIf txtStation.Text = "3" And txtStationItem.Text = "2" Then
                Call MasukPart2Bench3()
            ElseIf txtStation.Text = "3" And txtStationItem.Text = "3" Then
                Call MasukPart3Bench3()
            ElseIf txtStation.Text = "3" And txtStationItem.Text = "4" Then
                Call MasukPart4Bench3()
            ElseIf txtStation.Text = "3" And txtStationItem.Text = "5" Then
                Call MasukPart5Bench3()
            ElseIf txtStation.Text = "3" And txtStationItem.Text = "6" Then
                Call MasukPart6Bench3()
            ElseIf txtStation.Text = "3" And txtStationItem.Text = "7" Then
                Call MasukPart7Bench3()
            ElseIf txtStation.Text = "4" And txtStationItem.Text = "1" Then
                Call MasukPart1Bench4()
            ElseIf txtStation.Text = "4" And txtStationItem.Text = "2" Then
                Call MasukPart2Bench4()
            ElseIf txtStation.Text = "4" And txtStationItem.Text = "3" Then
                Call MasukPart3Bench4()
            ElseIf txtStation.Text = "4" And txtStationItem.Text = "4" Then
                Call MasukPart4Bench4()
            ElseIf txtStation.Text = "4" And txtStationItem.Text = "5" Then
                Call MasukPart5Bench4()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Transfer Material TypeOf HMI", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    Sub CekScanStatusSt1()
        Dim Target As Integer = Val(lblSt1TotalMaterial.Text)
        Dim Running As Integer = Val(lblSt1ScanTotal.Text)
        If Target = Running Then
            lblSt1Complete.Text = "COMPLETE"
            lblSt1Complete.Visible = True
            Call StatusCompleteBench1()
            'Call TypeModel1Bench1()
        Else
            lblSt1Complete.Text = "SCAN"
            lblSt1Complete.Visible = False
            Call StatusProgressBench1()
        End If
    End Sub
    Sub CekScanStatusSt2()
        Dim Target As Integer = Val(lblSt2TotalMaterial.Text)
        Dim Running As Integer = Val(lblSt2ScanTotal.Text)
        If Target = Running Then
            lblSt2Complete.Text = "COMPLETE"
            lblSt2Complete.Visible = True
            Call StatusCompleteBench2()
            'Call HidupPokaPickModel1Bench2()
        Else
            lblSt2Complete.Text = "SCAN"
            lblSt2Complete.Visible = False
            Call StatusProgressBench2()
        End If
    End Sub
    Sub CekScanStatusSt3()
        Dim Target As Integer = Val(lblSt3TotalMaterial.Text)
        Dim Running As Integer = Val(lblSt3ScanTotal.Text)
        If Target = Running Then
            lblSt3Complete.Text = "COMPLETE"
            lblSt3Complete.Visible = True
            Call StatusCompleteBench3()
            'Call HidupPokaPickModel1Bench3()

        Else
            lblSt3Complete.Text = "SCAN"
            lblSt3Complete.Visible = False
            Call StatusProgressBench3()
        End If
    End Sub
    Sub CekScanStatusSt4()
        Dim Target As Integer = Val(lblSt4TotalMaterial.Text)
        Dim Running As Integer = Val(lblSt4ScanTotal.Text)
        If Target = Running Then
            lblSt4Complete.Text = "COMPLETE"
            lblSt4Complete.Visible = True
            Call StatusCompleteBench4()
            'Call HidupPokaPickModel1Bench4()
        Else
            lblSt4Complete.Text = "SCAN"
            lblSt4Complete.Visible = False
            Call StatusProgressBench4()
            'Call HidupPokaPickModel1Bench4()

        End If
    End Sub
    Sub LihatAlamatMWPokapick()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader

        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "SELECT *FROM m_refMW WHERE (PP=" & Val(txtMW.Text) & ")"
            Reader = SQL1.ExecuteReader

            If Reader.Read Then
                txtPPmw.Text = Reader("MW").ToString
            Else
                txtPPmw.Text = ""
            End If
            Conn.Close()
            Reader.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString)
        End Try
    End Sub
    Sub ScanPartBench2()
        Dim Conn As SqlConnection
        Dim SQl1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()
            SQl1 = Conn.CreateCommand
            SQl1.CommandText = "SELECT *FROM m_refPartList WHERE (Model='" & lblRef.Text & "' AND PartNo='" & lblHMI2.Text & "' AND Station=2 AND Scan=0)"
            Reader = SQl1.ExecuteReader

            If Reader.Read Then
                ' Jika ada maka masukan status Scan
                txtStation.Text = Val(Reader("Station"))
                txtStationItem.Text = Val(Reader("ItemStation"))
                'txtMW.Text = Val(Reader("MW"))
                Dim NoMW As Integer = Val(Reader("MW"))
                Dim Checklist As Integer = Val(Reader("Checklist"))

                txtMW.Text = NoMW

                'Call LihatAlamatMWPokapick()

                Dim Conn1 As SqlConnection
                Dim SQL2 As New SqlCommand
                Conn1 = GetConnect()
                Conn1.Open()

                SQL2 = Conn1.CreateCommand
                SQL2.CommandText = "UPDATE m_refPartList SET Scan=1 WHERE (Model='" & lblRef.Text & "' AND PartNo='" & lblHMI2.Text & "' AND Station=2)" 'AND(Item=" & Val(txtStationItem.Text) & ")"
                SQL2.ExecuteReader()
                Conn1.Close()

                Call BACA_MATERIAL_COMPLETE()

                '===== Update Checklist HMI Bench 2 (PLC 106)
                Dim AddressCheck As Integer = Checklist
                Dim ValCheck(0) As Integer
                ValCheck(0) = 1
                modbusClient1.WriteMultipleRegisters(AddressCheck, ValCheck)

                '=====Tandai Nomor Pokapick yg akan Aktif (PLC 107)====
                'Dim StartAddress107 As Integer = NoMW 'Val(txtMW.Text)
                'Dim Values107(0) As Integer
                'Values107(0) = 1
                'modbusClient2.WriteMultipleRegisters(StartAddress107, Values107)
                '============================================

                Call StatusMaterialHMIProgressB2()

                Call dpStation1()
                Call dpStation2()

                '=====Beri Signal pada Station 1====
                Dim StartAddress1 As Integer = 10002
                Dim Values1(0) As Integer
                Values1(0) = 1
                modbusClient1.WriteMultipleRegisters(StartAddress1, Values1)
                '============================================

                txtScan.Text = ""
                lblLabel.Text = "Station 2: Material Correct!"
                lblLabel2.Text = ""

                Call CompleteScankeMWBench2()

                'Call ScanPartBench1()
                'MessageBox.Show("Scan Berhasil")
            Else
                'Call CekQrCodeBench2()
                'MessageBox.Show("Station 2: Wrong Material","Material")
                lblLabel.Text = "Station 2: Wrong Material"
                lblLabel2.Text = lblHMI2.Text

                Dim Status As String = "MATERIAL FAIL"
                Dim cb() As Byte
                'txtConvert.Text = ""

                cb = ASCIIEncoding.ASCII.GetBytes(Status)
                Dim MWval_reff(25) As Integer 'jumlah address hingga 20
                For x As Integer = 0 To cb.Length - 1
                    txtConvPart.AppendText(cb(x).ToString & "")
                    MWval_reff(x) = Convert.ToInt64(cb(x))
                    Try
                        Dim StartAddress As Integer = 2710 'Alamat MW Status Material Scan (PLC 106)
                        modbusClient1.WriteMultipleRegisters(StartAddress, MWval_reff)

                    Catch ex As Exception

                    End Try
                Next

            End If

            Conn.Close()
            Reader.Close()

        Catch ex As Exception
            'MessageBox.Show(ex.Message, "Scan Part Bench 2", MessageBoxButtons.OK)
        End Try
    End Sub
    Sub CekScanStation2()
        'tmStation2.Enabled = True
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "SELECT *FROM m_refPartList WHERE (Model='" & lblRef.Text & "' AND Station=2 AND Scan=0)" 'AND(PartNo='" & lblHMI2.Text & "')AND(Station=2)AND(Scan=0)"
            Reader = SQL1.ExecuteReader

            If Reader.Read Then
                Call ScanPartBench2()
            Else
                Call CekQrCodeBench2()
            End If

            Conn.Close()
            Reader.Close()

            'tmStation2.Enabled = False

        Catch ex As Exception

        End Try
    End Sub
    Sub CekScanStation3()
        'tmStation2.Enabled = True

        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "SELECT *FROM m_refPartList WHERE (Model='" & lblRef.Text & "' AND Station=3 AND Scan=0)"
            Reader = SQL1.ExecuteReader
            If Reader.Read Then
                Call ScanPartBench3()
            Else
                Call CekQrCodeBench3()
            End If
            Conn.Close()
            Reader.Close()

            'tmStation2.Enabled = False

        Catch ex As Exception

        End Try
    End Sub
    Sub ScanPartBench3()
        Dim Conn As SqlConnection
        Dim SQl1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()
            SQl1 = Conn.CreateCommand
            SQl1.CommandText = "SELECT *FROM m_refPartList WHERE (Model='" & lblRef.Text & "')AND(Station=3)AND(PartNo='" & lblHMI3.Text & "')AND(Scan=0)"
            Reader = SQl1.ExecuteReader

            If Reader.Read Then
                ' Jika ada maka masukan status Scan
                txtStation.Text = Val(Reader("Station"))
                txtStationItem.Text = Val(Reader("ItemStation"))
                'txtMW.Text = Val(Reader("MW"))
                Dim NoMW As Integer = Val(Reader("MW"))
                Dim Checklist As Integer = Val(Reader("Checklist"))

                '===== Update Checklist HMI Bench 3 (PLC 107)
                Dim AddressCheck As Integer = Checklist
                Dim ValCheck(0) As Integer
                ValCheck(0) = 1
                modbusClient2.WriteMultipleRegisters(AddressCheck, ValCheck)



                Dim Conn1 As SqlConnection
                Dim SQL2 As New SqlCommand
                Conn1 = GetConnect()
                Conn1.Open()

                SQL2 = Conn1.CreateCommand
                SQL2.CommandText = "UPDATE m_refPartList SET Scan=1 WHERE (Model='" & lblRef.Text & "' AND Station=3)AND(PartNo='" & lblHMI3.Text & "')" 'AND(Item=" & Val(txtStationItem.Text) & ")"
                SQL2.ExecuteReader()
                Conn1.Close()

                Call BACA_MATERIAL_COMPLETE()

                lblLabel.Text = "Station 3: Marial Correct!"
                lblLabel2.Text = ""


                Call StatusMaterialHMIProgressB3()

                Call dpStation1()
                Call dpStation3()

                '=====Beri Signal pada Station 1====
                Dim StartAddress1 As Integer = 10002
                Dim Values1(0) As Integer
                Values1(0) = 1
                modbusClient1.WriteMultipleRegisters(StartAddress1, Values1)
                '============================================

                txtScan.Text = ""
                'lblLabel.Text = "SCAN MATERIAL..."

                Call CompleteScankeMWBench3()

                'Call ScanPartBench1()
                'MessageBox.Show("Scan Berhasil")
            Else
                'Call CekQrCodeBench3()
                lblLabel2.Text = lblHMI3.Text
                lblLabel.Text = "Station 3: Wrong Material"

                Dim Status As String = "WRONG MATERIAL"
                Dim cb() As Byte
                'txtConvert.Text = ""

                cb = ASCIIEncoding.ASCII.GetBytes(Status)
                Dim MWval_reff(25) As Integer 'jumlah address hingga 20
                For x As Integer = 0 To cb.Length - 1
                    txtConvPart.AppendText(cb(x).ToString & "")
                    MWval_reff(x) = Convert.ToInt64(cb(x))
                    Try
                        Dim StartAddress As Integer = 3710 'Alamat MW Status Material Scan (PLC 106)
                        modbusClient2.WriteMultipleRegisters(StartAddress, MWval_reff)

                    Catch ex As Exception

                    End Try
                Next
            End If

            Conn.Close()
            Reader.Close()

        Catch ex As Exception
            'MessageBox.Show(ex.Message, "Scan Part Bench 3", MessageBoxButtons.OK)
        End Try
    End Sub


    Private Sub CEKPCB_DOUBLE()
        Dim CONN As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim READER As SqlDataReader

        Try
            CONN = GetConnect()
            CONN.Open()
            SQL1 = CONN.CreateCommand
            SQL1.CommandText = "SELECT *FROM m_refRunningitem WHERE ( MODEL = '" & lblRef.Text & "' ) AND (PCBNo='" & lblHMI4.Text & "')"
            READER = SQL1.ExecuteReader
            If READER.Read Then
                'MessageBox.Show("PCB N")
                lblLabel.Text = "Station 4: PCB NO. ALREADY USED"
                lblLabel2.Text = lblHMI4.Text

                Dim Salah As String = "PCB ALREADY USED"
                Dim cb() As Byte
                'txtConvert.Text = ""
                'If lblHMI4.Text.Length > 6 Then
                cb = ASCIIEncoding.ASCII.GetBytes(Salah)
                Dim MWval_reff(49) As Integer 'jumlah address hingga 30
                For x As Integer = 0 To cb.Length - 1
                    txtConvPart.AppendText(cb(x).ToString & "")
                    'txtFrom.Text = ""
                    'txtTo.AppendText(cb(x).ToString & "")
                    'Dim arrayRef() as Char=cb(x).to
                    'MessageBox.Show(cb(x).ToString)
                    MWval_reff(x) = Convert.ToInt64(cb(x))
                    Try
                        Dim StartAddress As Integer = 4800 'Alamat PCB di HMI sampai 4850
                        modbusClient2.WriteMultipleRegisters(StartAddress, MWval_reff)

                        'Call MasukPart1PPBench1()

                    Catch ex As Exception

                    End Try
                Next
            Else
                lblMasterPCB.Text = ""

                Dim Scan As String = lblHMI4.Text
                Dim Data As String = Mid(Scan, 4, 8)
                lblMasterPCB.Text = Data

                If Data = lblPCB.Text Then

                    'Try
                    Dim Conn1 As SqlConnection
                    Dim SQl2 As New SqlCommand
                    Conn1 = GetConnect()
                    Conn1.Open()
                    SQl2 = Conn1.CreateCommand
                    SQl2.CommandText = "UPDATE m_refRunningItem SET ScanPCB=1 , PCBNo='" & lblHMI4.Text & "',Status='C' WHERE QRCode = '" & lblLastQR.Text & "' AND St = 4"
                    SQl2.ExecuteReader()
                    Conn1.Close()

                    Call KirimPCBtoHMI()

                    '=========== Matikan Perintah Scan PCB (PLC 107)
                    Dim StartAddress As Integer = 5001
                    Dim Values(0) As Integer
                    Values(0) = 1
                    modbusClient2.WriteMultipleRegisters(StartAddress, Values)

                    Dim Start5000 As Integer = 5000
                    Dim Val5000(0) As Integer
                    Val5000(0) = 0
                    modbusClient2.WriteMultipleRegisters(Start5000, Val5000)
                    '===============================================

                    lblLabel.Text = "Station 4: PCB Correct"
                    lblLabel2.Text = lblHMI4.Text

                    'SEKA
                    SEQUENCE_PICKUPBENCH4()

                    'tmrPick4.Enabled = True

                    'Call SEQUENCE_PICKUPBENCH4()
                    tmrPick4.Enabled = True
                    tmrPCBA.Enabled = False
                Else
                    Call WrongPCBA()
                    'lblHMI4.Text = "0"
                    lblLabel.Text = "Station 4: Wrong PCB"
                    lblLabel2.Text = lblHMI4.Text
                End If
            End If
            CONN.Close()
            READER.Close()

        Catch ex As Exception

        End Try

    End Sub

    Dim FINISH_SCAN_QR As Boolean = False

    Sub CekScanStation4()
        'tmStation2.Enabled = True

        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "SELECT *FROM m_refPartList WHERE (Model='" & lblRef.Text & "' AND Station=4 AND Scan=0)"
            Reader = SQL1.ExecuteReader
            If Reader.Read Then
                Call ScanPartBench4()
            Else
                'If lblScanPCBA.Text = "1" Then
                'Call CEKPCB_DOUBLE()
                'Else
                If Not FINISH_SCAN_QR Then
                    Call CekQrCodeBench4()
                End If
                'End If

            End If
            Conn.Close()
            Reader.Close()

        Catch ex As Exception

        End Try
    End Sub
    Private Sub WrongPCBA()
        'KirimPCBtoHMI
        Dim Salah As String = "WRONG QR PCBA"
        Dim cb() As Byte
        'txtConvert.Text = ""
        'If lblHMI4.Text.Length > 6 Then
        cb = ASCIIEncoding.ASCII.GetBytes(Salah)
        Dim MWval_reff(49) As Integer 'jumlah address hingga 30
        For x As Integer = 0 To cb.Length - 1
            txtConvPart.AppendText(cb(x).ToString & "")
            'txtFrom.Text = ""
            'txtTo.AppendText(cb(x).ToString & "")
            'Dim arrayRef() as Char=cb(x).to
            'MessageBox.Show(cb(x).ToString)
            MWval_reff(x) = Convert.ToInt64(cb(x))
            Try
                Dim StartAddress As Integer = 4800 'Alamat PCB di HMI sampai 4850
                modbusClient2.WriteMultipleRegisters(StartAddress, MWval_reff)

                'Call MasukPart1PPBench1()

            Catch ex As Exception

            End Try
        Next

        'End If

    End Sub

    'End Sub

    Private Sub KirimPCBtoHMI()
        Dim cb() As Byte
        'txtConvert.Text = ""
        If lblHMI4.Text.Length > 6 Then
            cb = ASCIIEncoding.ASCII.GetBytes(lblHMI4.Text)
            Dim MWval_reff(49) As Integer 'jumlah address hingga 30
            For x As Integer = 0 To cb.Length - 1
                txtConvPart.AppendText(cb(x).ToString & "")
                'txtFrom.Text = ""
                'txtTo.AppendText(cb(x).ToString & "")
                'Dim arrayRef() as Char=cb(x).to
                'MessageBox.Show(cb(x).ToString)
                MWval_reff(x) = Convert.ToInt64(cb(x))
                Try
                    Dim StartAddress As Integer = 4800 'Alamat PCB di HMI sampai 4850
                    modbusClient2.WriteMultipleRegisters(StartAddress, MWval_reff)

                    'Call MasukPart1PPBench1()

                Catch ex As Exception

                End Try
            Next

        End If

    End Sub
    Private Sub HapusPCBdiHMI()
        Try

            ' Hapus PCB number di HMI Bench 4, mulai dari MW4800 - MW4850
            Dim StartAddress As Integer = 4800
            Dim Values(49) As Integer
            Values(49) = 0
            modbusClient2.WriteMultipleRegisters(StartAddress, Values)
        Catch ex As Exception

        End Try

    End Sub
    Sub ScanPartBench4()
        Dim Conn As SqlConnection
        Dim SQl1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()
            SQl1 = Conn.CreateCommand
            SQl1.CommandText = "SELECT *FROM m_refPartList WHERE (Model='" & lblRef.Text & "' AND Station=4) AND (PartNo='" & lblHMI4.Text & "')AND(Scan=0)"
            Reader = SQl1.ExecuteReader

            If Reader.Read Then
                ' Jika ada maka masukan status Scan
                txtStation.Text = Val(Reader("Station"))
                txtStationItem.Text = Val(Reader("ItemStation"))
                'txtMW.Text = Val(Reader("MW"))
                Dim NoMW As Integer = Val(Reader("MW"))
                Dim Checklist As Integer = Val(Reader("Checklist"))


                Dim Conn1 As SqlConnection
                Dim SQL2 As New SqlCommand
                Conn1 = GetConnect()
                Conn1.Open()

                SQL2 = Conn1.CreateCommand
                SQL2.CommandText = "UPDATE m_refPartList SET Scan=1 WHERE (Model='" & lblRef.Text & "' AND Station=4)AND(PartNo='" & lblHMI4.Text & "')" 'AND(Item=" & Val(txtStationItem.Text) & ")"
                SQL2.ExecuteReader()
                Conn1.Close()

                'Call MasukPartListHMI()

                Call BACA_MATERIAL_COMPLETE()

                '===== Update Checklist HMI Bench 1 (PLC 106)
                Dim AddressCheck As Integer = Checklist
                Dim ValCheck(0) As Integer
                ValCheck(0) = 1
                modbusClient2.WriteMultipleRegisters(AddressCheck, ValCheck)


                '=====Tandai Nomor Pokapick yg akan Aktif (PLC 106)====
                'Dim StartAddress As Integer = NoMW
                'Dim Values(0) As Integer
                'Values(0) = 1
                'modbusClient1.WriteMultipleRegisters(StartAddress, Values)
                '============================================

                '=====Tandai Nomor Pokapick yg akan Aktif (PLC 107)====
                'Dim StartAddress107 As Integer = NoMW 'Val(txtMW.Text)
                'Dim Values107(0) As Integer
                'Values107(0) = 1
                'modbusClient2.WriteMultipleRegisters(StartAddress107, Values107)
                '============================================

                Call StatusMaterialHMIProgressB4()

                Call dpStation1()
                Call dpStation4()

                '=====Beri Signal pada Station 1====
                Dim StartAddress1 As Integer = 10002
                Dim Values1(0) As Integer
                Values1(0) = 1
                modbusClient1.WriteMultipleRegisters(StartAddress1, Values1)
                '============================================

                txtScan.Text = ""
                lblLabel.Text = "Station 4: Material Correct!"
                lblLabel2.Text = ""

                Call CompleteScankeMWBench4()

                'MessageBox.Show("Scan Berhasil")
            Else
                lblLabel2.Text = lblHMI4.Text
                lblLabel.Text = "Station 4: Wrong Material"

                Dim Status As String = "WRONG MATERIAL"
                Dim cb() As Byte
                'txtConvert.Text = ""

                cb = ASCIIEncoding.ASCII.GetBytes(Status)
                Dim MWval_reff(25) As Integer 'jumlah address hingga 20
                For x As Integer = 0 To cb.Length - 1
                    txtConvPart.AppendText(cb(x).ToString & "")
                    MWval_reff(x) = Convert.ToInt64(cb(x))
                    Try
                        Dim StartAddress As Integer = 4710 'Alamat MW Status Material Scan (PLC 106)
                        modbusClient2.WriteMultipleRegisters(StartAddress, MWval_reff)

                    Catch ex As Exception

                    End Try
                Next
            End If

            Conn.Close()
            Reader.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Scan Part Bench 4", MessageBoxButtons.OK)
        End Try
    End Sub

    Sub StatusMaterialHMIProgressB2()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "SELECT *FROM m_refPartList WHERE (Model='" & lblRef.Text & "' AND Station=2 AND Scan=0)"
            Reader = SQL1.ExecuteReader
            If Reader.Read Then
                Dim Status As String = "PROGRESS"
                Dim cb() As Byte
                'txtConvert.Text = ""

                cb = ASCIIEncoding.ASCII.GetBytes(Status)
                Dim MWval_reff(19) As Integer 'jumlah address hingga 20
                For x As Integer = 0 To cb.Length - 1
                    txtConvPart.AppendText(cb(x).ToString & "")
                    MWval_reff(x) = Convert.ToInt64(cb(x))
                    Try
                        Dim StartAddress As Integer = 2710 'mulai dari MW600
                        modbusClient1.WriteMultipleRegisters(StartAddress, MWval_reff)

                    Catch ex As Exception

                    End Try
                Next
            Else

                Dim Status As String = "COMPLETE"
                Dim cb() As Byte
                'txtConvert.Text = ""

                cb = ASCIIEncoding.ASCII.GetBytes(Status)
                Dim MWval_reff(19) As Integer 'jumlah address hingga 20
                For x As Integer = 0 To cb.Length - 1
                    txtConvPart.AppendText(cb(x).ToString & "")
                    MWval_reff(x) = Convert.ToInt64(cb(x))
                    Try
                        Dim StartAddress As Integer = 2710 'Alamat MW Status Material Scan (PLC 106)
                        modbusClient1.WriteMultipleRegisters(StartAddress, MWval_reff)

                    Catch ex As Exception

                    End Try
                Next

            End If
            Conn.Close()
            Reader.Close()

        Catch ex As Exception

        End Try

    End Sub
    Sub StatusMaterialHMIProgressB3()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "SELECT *FROM m_refPartList WHERE (Model='" & lblRef.Text & "' AND Station=3 AND Scan=0)"
            Reader = SQL1.ExecuteReader
            If Reader.Read Then
                Dim Status As String = "PROGRESS"
                Dim cb() As Byte
                'txtConvert.Text = ""

                cb = ASCIIEncoding.ASCII.GetBytes(Status)
                Dim MWval_reff(19) As Integer 'jumlah address hingga 20
                For x As Integer = 0 To cb.Length - 1
                    txtConvPart.AppendText(cb(x).ToString & "")
                    MWval_reff(x) = Convert.ToInt64(cb(x))
                    Try
                        Dim StartAddress As Integer = 3710 'mulai dari MW600
                        modbusClient2.WriteMultipleRegisters(StartAddress, MWval_reff)

                    Catch ex As Exception

                    End Try
                Next
            Else

                Dim Status As String = "COMPLETE"
                Dim cb() As Byte
                'txtConvert.Text = ""

                cb = ASCIIEncoding.ASCII.GetBytes(Status)
                Dim MWval_reff(19) As Integer 'jumlah address hingga 20
                For x As Integer = 0 To cb.Length - 1
                    txtConvPart.AppendText(cb(x).ToString & "")
                    MWval_reff(x) = Convert.ToInt64(cb(x))
                    Try
                        Dim StartAddress As Integer = 3710 'Alamat MW Status Material Scan (PLC 106)
                        modbusClient2.WriteMultipleRegisters(StartAddress, MWval_reff)

                    Catch ex As Exception

                    End Try
                Next

            End If
            Conn.Close()
            Reader.Close()

        Catch ex As Exception

        End Try

    End Sub
    Sub StatusMaterialHMIProgressB4()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "SELECT *FROM m_refPartList WHERE (Model='" & lblRef.Text & "' AND Station=4 AND Scan=0)"
            Reader = SQL1.ExecuteReader
            If Reader.Read Then
                Dim Status As String = "PROGRESS"
                Dim cb() As Byte
                'txtConvert.Text = ""

                cb = ASCIIEncoding.ASCII.GetBytes(Status)
                Dim MWval_reff(19) As Integer 'jumlah address hingga 20
                For x As Integer = 0 To cb.Length - 1
                    txtConvPart.AppendText(cb(x).ToString & "")
                    MWval_reff(x) = Convert.ToInt64(cb(x))
                    Try
                        Dim StartAddress As Integer = 4710 'mulai dari MW600
                        modbusClient2.WriteMultipleRegisters(StartAddress, MWval_reff)

                    Catch ex As Exception

                    End Try
                Next
            Else

                Dim Status As String = "COMPLETE"
                Dim cb() As Byte
                'txtConvert.Text = ""

                cb = ASCIIEncoding.ASCII.GetBytes(Status)
                Dim MWval_reff(19) As Integer 'jumlah address hingga 20
                For x As Integer = 0 To cb.Length - 1
                    txtConvPart.AppendText(cb(x).ToString & "")
                    MWval_reff(x) = Convert.ToInt64(cb(x))
                    Try
                        Dim StartAddress As Integer = 4710 'Alamat MW Status Material Scan (PLC 106)
                        modbusClient2.WriteMultipleRegisters(StartAddress, MWval_reff)

                    Catch ex As Exception

                    End Try
                Next

            End If
            Conn.Close()
            Reader.Close()

        Catch ex As Exception

        End Try

    End Sub

    Private Sub SEQUENCE_PICKUPBENCH2()
        Dim CONN As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim READER As SqlDataReader
        Dim Sequence As Integer

        Try
            CONN = GetConnect()
            CONN.Open()
            SQL1 = CONN.CreateCommand
            SQL1.CommandText = "SELECT TOP 1 Model,Station,ONOFF,SEQ,ItemStation,MW,OFFADDRESS,CheckList FROM m_RefPartList WHERE (Model='" & lblRef.Text & "' AND Station=2 AND ONOFF=0) ORDER BY SEQ ASC"
            READER = SQL1.ExecuteReader
            If READER.Read Then
                txtStation.Text = Val(READER("Station"))
                txtStationItem.Text = Val(READER("ItemStation"))
                txtMW2.Text = Val(READER("MW"))
                txtOFF2.Text = Val(READER("OFFADDRESS"))
                Dim Checklist As Integer = Val(READER("Checklist"))
                Dim ItemStation As Integer = Val(READER("ItemStation"))
                Sequence = Val(READER("SEQ"))
                'Screwing = Val(READER("Screw"))

                '=====HIDUP POKAPICK BENCH2 (106)====
                Dim StartAddress1 As Integer = Val(txtMW2.Text)
                Dim Values1(0) As Integer
                Values1(0) = 1
                modbusClient1.WriteMultipleRegisters(StartAddress1, Values1)
                '============================================

                Dim CONN1 As SqlConnection
                Dim SQL2 As New SqlCommand
                CONN1 = GetConnect()
                CONN1.Open()
                SQL2 = CONN1.CreateCommand
                SQL2.CommandText = "UPDATE m_refPartList SET ONOFF=1 WHERE (Model='" & lblRef.Text & "' AND Station=2 AND Seq=" & Sequence & " AND ItemStation=" & ItemStation & ")"
                SQL2.ExecuteReader()
                CONN1.Close()

            Else
                lblLabel.Text = "Station 2: PICKUP COMPLETE"

                COMPLETE_BECNH_2()

                Call MasukQtyRunningMWBench2()

                Dim StartAddress1 As Integer = 9001
                Dim Values1(0) As Integer
                Values1(0) = 1
                modbusClient1.WriteMultipleRegisters(StartAddress1, Values1)

                'tmrQtyB2.Enabled = True
                Call KOSONGKAN_SEQUENCEBENCH2()
                'Call MasukQtyRunningMWBench2()

                tmrPick2.Enabled = False

            End If
            CONN.Close()
            READER.Close()

        Catch ex As Exception

        End Try
    End Sub



    'SEKA SUB COMPLETE PRODUCT
    Private Sub COMPLETE_BECNH_2()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand 

        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "UPDATE M_REFRUNNINGITEM SET NEXTSCAN = 3 , ST = 3 , QtyRunSt2 = 1 WHERE  QRCODE ='" & lblHMI2.Text & "'  "
            SQL1.ExecuteNonQuery()

            modbusClient1.WriteSingleRegister(MW_COMPLETE_BENCH_2, 1)
        Catch ex As Exception

        End Try
       
    End Sub

    Private Sub COMPLETE_BECNH_3()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand

        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "UPDATE M_REFRUNNINGITEM SET NEXTSCAN = 4 , ST = 4  , QtyRunSt3 = 1  WHERE QRCODE ='" & lblHMI3.Text & "'  "
            SQL1.ExecuteNonQuery()
            modbusClient2.WriteSingleRegister(MW_COMPLETE_BENCH_3, 1)
        Catch ex As Exception

        End Try

    End Sub

    Private Sub COMPLETE_BECNH_4()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand


        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "UPDATE M_REFRUNNINGITEM SET NEXTSCAN = 5 , ST = 5  , QtyRunSt4 = 1  WHERE QRCODE ='" & lblLastQR.Text & "'  "
            SQL1.ExecuteNonQuery()

            modbusClient2.WriteSingleRegister(MW_COMPLETE_BENCH_4, 1)

            FINISH_SCAN_QR = False

        Catch ex As Exception

        End Try

    End Sub



    Private Sub SEQUENCE_PICKUPBENCH3()
        Dim CONN As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim READER As SqlDataReader
        Dim Sequence As Integer
        Dim Screwing As Integer

        Try
            CONN = GetConnect()
            CONN.Open()
            SQL1 = CONN.CreateCommand
            SQL1.CommandText = "SELECT TOP 1 *FROM m_RefPartList WHERE (Model='" & lblRef.Text & "' AND Station=3 AND ONOFF=0) ORDER BY Seq ASC"
            READER = SQL1.ExecuteReader
            If READER.Read Then
                txtStation.Text = Val(READER("Station"))
                txtStationItem.Text = Val(READER("ItemStation"))
                txtMW3.Text = Val(READER("MW"))
                txtOFF3.Text = Val(READER("OFFADDRESS"))
                Dim Checklist As Integer = Val(READER("Checklist"))
                Dim ItemStation As Integer = Val(READER("ItemStation"))
                Sequence = Val(READER("Seq"))
                Screwing = Val(READER("Screw"))

                If Screwing = 1 Then
                    '=====HIDUP SCREWING (102)====
                    Dim StartAddress As Integer = Val(txtMW3.Text)
                    Dim Values(0) As Integer
                    Values(0) = 1
                    modbusScrewing.WriteMultipleRegisters(StartAddress, Values)

                    lblScrew3.Text = "SCREWING"
                    'modbusClient1.WriteMultipleRegisters(StartAddress, Values)
                    '============================================
                ElseIf Screwing = 0 Then
                    '=====HIDUP POKAPICK (106)====
                    Dim StartAddress1 As Integer = Val(txtMW3.Text)
                    Dim Values1(0) As Integer
                    Values1(0) = 1
                    modbusClient2.WriteMultipleRegisters(StartAddress1, Values1)

                    lblScrew3.Text = ""

                    '============================================
                End If

                Dim CONN1 As SqlConnection
                Dim SQL2 As New SqlCommand
                CONN1 = GetConnect()
                CONN1.Open()
                SQL2 = CONN1.CreateCommand
                SQL2.CommandText = "UPDATE m_refPartList SET ONOFF=1 WHERE (Model='" & lblRef.Text & "' AND Station=3 AND Seq=" & Sequence & " AND ItemStation=" & ItemStation & ")"
                SQL2.ExecuteReader()
                CONN1.Close()

                tmrScrew3.Enabled = True
                tmrPick3.Enabled = True

            Else
                'tmrPickup.Enabled = False
                'tmrScrewing.Enabled = False
                'Call SEQUENCE_SCREWINGBENCH3()

                lblLabel.Text = "Station 3: PICKUP COMPLETE"

                COMPLETE_BECNH_3()


                Call MasukQtyRunningMWBench3()

                Dim StartAddress1 As Integer = 9002
                Dim Values1(0) As Integer
                Values1(0) = 1
                modbusClient2.WriteMultipleRegisters(StartAddress1, Values1)

                Call KOSONGKAN_SEQUENCEBENCH3()
                tmrScrew3.Enabled = False
                tmrPick3.Enabled = False

                lblScrew3.Text = ""

            End If
            CONN.Close()
            READER.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error 192.168.1.102", MessageBoxButtons.OK)
        End Try
    End Sub
    Private Sub SEQUENCE_PICKUPBENCH4()
        Dim CONN As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim READER As SqlDataReader
        Dim Sequence As Integer
        Dim Screwing As Integer

        Try
            CONN = GetConnect()
            CONN.Open()
            SQL1 = CONN.CreateCommand
            SQL1.CommandText = "SELECT TOP 1 *FROM m_RefPartList WHERE (Model='" & lblRef.Text & "' AND Station=4 AND ONOFF=0) ORDER BY SEQ ASC"
            READER = SQL1.ExecuteReader
            If READER.Read Then
                txtStation.Text = Val(READER("Station"))
                txtStationItem.Text = Val(READER("ItemStation"))
                txtMW4.Text = Val(READER("MW"))
                txtOFF4.Text = Val(READER("OFFADDRESS"))
                Dim Checklist As Integer = Val(READER("Checklist"))
                Dim ItemStation As Integer = Val(READER("ItemStation"))
                Sequence = Val(READER("Seq"))
                Screwing = Val(READER("Screw"))


                '=====HIDUP POKAPICK (106)====
                Dim StartAddress1 As Integer = Val(txtMW4.Text)
                Dim Values1(0) As Integer
                Values1(0) = 1
                modbusClient2.WriteMultipleRegisters(StartAddress1, Values1)




                'lblScrew4.Text = ""WRITE
                '============================================

                'tmrPick4.Enabled = True

                Dim CONN1 As SqlConnection
                Dim SQL2 As New SqlCommand
                CONN1 = GetConnect()
                CONN1.Open()
                SQL2 = CONN1.CreateCommand
                SQL2.CommandText = "UPDATE m_refPartList SET ONOFF=1 WHERE (Model='" & lblRef.Text & "' AND Station=4 AND Seq=" & Sequence & " AND ItemStation=" & ItemStation & ")"
                SQL2.ExecuteReader()
                CONN1.Close()
            Else
                'tmrPickup.Enabled = False
                'tmrScrewing.Enabled = False
                'Call SEQUENCE_SCREWINGBENCH4()

                '==============================================
                tmrQtyB4.Enabled = True 'Untuk masukan Qty running ke HMI
                Call HitungRunQty()

                '
                Dim StartAddress As Integer = 5001
                Dim Values(0) As Integer
                Values(0) = 0
                modbusClient2.WriteMultipleRegisters(StartAddress, Values)

                Dim Start5000 As Integer = 5000
                Dim Val5000(0) As Integer
                Val5000(0) = 0
                modbusClient2.WriteMultipleRegisters(Start5000, Val5000)
                '===============================================

                lblLabel.Text = "Station 4: PICKUP COMPLETE"

                COMPLETE_BECNH_4()

                Dim StartAddress1 As Integer = 9003
                Dim Values1(0) As Integer
                Values1(0) = 1
                modbusClient2.WriteMultipleRegisters(StartAddress1, Values1)

                Call KOSONGKAN_SEQUENCEBENCH4()
                tmrPick4.Enabled = False
                'tmrScrew4.Enabled = False

            End If
            CONN.Close()
            READER.Close()

        Catch ex As Exception

        End Try
    End Sub
    Private Sub SEQUENCE_SCREWINGBENCH3()
        If lblScrew3.Text = "SCREWING" Then
            tmrScrew3.Enabled = True
            tmrPick3.Enabled = False
        Else
            tmrScrew3.Enabled = False
            tmrPick3.Enabled = True
        End If
    End Sub
    Private Sub SEQUENCE_SCREWINGBENCH4()
        If lblScrew4.Text = "SCREWING" Then
            tmrScrew4.Enabled = True
            tmrPick4.Enabled = False
        Else
            tmrScrew4.Enabled = False
            tmrPick4.Enabled = True
        End If
    End Sub
    Private Sub KOSONGKAN_SEQUENCEBENCH2()
        Dim CONN As SqlConnection
        Dim SQL1 As New SqlCommand
        Try
            CONN = GetConnect()
            CONN.Open()
            SQL1 = CONN.CreateCommand
            SQL1.CommandText = "UPDATE m_refPartList SET ONOFF=0 WHERE (Model='" & lblRef.Text & "' AND Station=2)"
            SQL1.ExecuteReader()
            CONN.Close()

        Catch ex As Exception

        End Try
    End Sub
    Private Sub KOSONGKAN_SEQUENCEBENCH3()
        Dim CONN As SqlConnection
        Dim SQL1 As New SqlCommand
        Try
            CONN = GetConnect()
            CONN.Open()
            SQL1 = CONN.CreateCommand
            SQL1.CommandText = "UPDATE m_refPartList SET ONOFF=0 WHERE (Model='" & lblRef.Text & "' AND Station=3)"
            SQL1.ExecuteReader()
            CONN.Close()

        Catch ex As Exception

        End Try
    End Sub
    Private Sub KOSONGKAN_SEQUENCEBENCH4()
        Dim CONN As SqlConnection
        Dim SQL1 As New SqlCommand
        Try
            CONN = GetConnect()
            CONN.Open()
            SQL1 = CONN.CreateCommand
            SQL1.CommandText = "UPDATE m_refPartList SET ONOFF=0 WHERE (Model='" & lblRef.Text & "' AND Station=4)"
            SQL1.ExecuteReader()
            CONN.Close()

        Catch ex As Exception

        End Try
    End Sub
    Private Sub START_SEQUENCEBENCH2()
        Dim CONN As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim READER As SqlDataReader
        Dim Sequence As Integer
        'Dim MW As Integer
        Try
            'Call KOSONGKAN_SEQUENCE()

            CONN = GetConnect()
            CONN.Open()
            SQL1 = CONN.CreateCommand
            SQL1.CommandText = "SELECT TOP 1 *FROM m_RefPartList WHERE (Model='" & lblRef.Text & "' AND Station=2 AND ONOFF=0) ORDER BY Seq ASC"
            READER = SQL1.ExecuteReader
            If READER.Read Then
                txtStation.Text = Val(READER("Station"))
                txtStationItem.Text = Val(READER("ItemStation"))
                txtMW2.Text = Val(READER("MW"))
                txtOFF2.Text = Val(READER("OFFADDRESS"))
                Dim Checklist As Integer = Val(READER("Checklist"))
                Sequence = Val(READER("Seq"))

                '=====HIDUP POKAPICK (106)====
                Dim StartAddress As Integer = Val(txtMW2.Text)
                Dim Values(0) As Integer
                Values(0) = 1
                modbusClient1.WriteMultipleRegisters(StartAddress, Values)
                '============================================

                Dim CONN1 As SqlConnection
                Dim SQL2 As New SqlCommand
                CONN1 = GetConnect()
                CONN1.Open()
                SQL2 = CONN1.CreateCommand
                SQL2.CommandText = "UPDATE m_refPartList SET ONOFF=1 WHERE (Model='" & lblRef.Text & "' AND Station=2 AND Seq=" & Sequence & ")"
                SQL2.ExecuteReader()
                CONN1.Close()

                tmrPick2.Enabled = True
                'tmrScrewing.Enabled = True

            Else
                lblLabel.Text = "MATERIAL COMPLETE"
            End If
            CONN.Close()
            READER.Close()

        Catch ex As Exception

        End Try
    End Sub
    Private Sub START_SEQUENCEBENCH3()
        Dim CONN As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim READER As SqlDataReader
        Dim Sequence As Integer
        'Dim MW As Integer
        Try
            'Call KOSONGKAN_SEQUENCE()

            CONN = GetConnect()
            CONN.Open()
            SQL1 = CONN.CreateCommand
            SQL1.CommandText = "SELECT TOP 1 *FROM m_RefPartList WHERE (Model='" & lblRef.Text & "' AND Station=3 AND ONOFF=0) ORDER BY Seq ASC"
            READER = SQL1.ExecuteReader
            If READER.Read Then
                txtStation.Text = Val(READER("Station"))
                txtStationItem.Text = Val(READER("ItemStation"))
                txtMW3.Text = Val(READER("MW"))
                txtOFF3.Text = Val(READER("OFFADDRESS"))
                Dim Checklist As Integer = Val(READER("Checklist"))
                Sequence = Val(READER("Seq"))

                '=====HIDUP POKAPICK BENCH 3 (107)====
                Dim StartAddress As Integer = Val(txtMW3.Text)
                Dim Values(0) As Integer
                Values(0) = 1
                modbusClient2.WriteMultipleRegisters(StartAddress, Values)
                '============================================

                Dim CONN1 As SqlConnection
                Dim SQL2 As New SqlCommand
                CONN1 = GetConnect()
                CONN1.Open()
                SQL2 = CONN1.CreateCommand
                SQL2.CommandText = "UPDATE m_refPartList SET ONOFF=1 WHERE (Model='" & lblRef.Text & "' AND Station=3 AND Seq=" & Sequence & ")"
                SQL2.ExecuteReader()
                CONN1.Close()

                tmrPick3.Enabled = True
                'tmrScrewing.Enabled = True

            Else
                lblLabel.Text = "Station 3: MATERIAL COMPLETE"

                tmrPick3.Enabled = False
            End If
            CONN.Close()
            READER.Close()

        Catch ex As Exception

        End Try
    End Sub
    Private Sub START_SEQUENCEBENCH4()
        Dim CONN As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim READER As SqlDataReader
        Dim Sequence As Integer
        'Dim MW As Integer
        Try
            'Call KOSONGKAN_SEQUENCE()

            CONN = GetConnect()
            CONN.Open()
            SQL1 = CONN.CreateCommand
            SQL1.CommandText = "SELECT TOP 1 *FROM m_RefPartList WHERE (Model='" & lblRef.Text & "' AND Station=4 AND ONOFF=0) ORDER BY Seq ASC"
            READER = SQL1.ExecuteReader
            If READER.Read Then
                txtStation.Text = Val(READER("Station"))
                txtStationItem.Text = Val(READER("ItemStation"))
                txtMW4.Text = Val(READER("MW"))
                txtOFF4.Text = Val(READER("OFFADDRESS"))
                Dim Checklist As Integer = Val(READER("Checklist"))
                Sequence = Val(READER("Seq"))

                '=====HIDUP POKAPICK BENCH 4 (107)====
                Dim StartAddress As Integer = Val(txtMW4.Text)
                Dim Values(0) As Integer
                Values(0) = 1
                modbusClient2.WriteMultipleRegisters(StartAddress, Values)
                '============================================

                tmrPCBA.Enabled = True

                Dim CONN1 As SqlConnection
                Dim SQL2 As New SqlCommand
                CONN1 = GetConnect()
                CONN1.Open()
                SQL2 = CONN1.CreateCommand
                SQL2.CommandText = "UPDATE m_refPartList SET ONOFF=1 WHERE (Model='" & lblRef.Text & "' AND Station=4 AND Seq=" & Sequence & ")"
                SQL2.ExecuteReader()
                CONN1.Close()

                'tmrPick4.Enabled = True
                'tmrScrewing.Enabled = True

            Else
                lblLabel.Text = "Station 4: ASSEMBLY COMPLETE"
                tmrPick4.Enabled = False

            End If
            CONN.Close()
            READER.Close()

        Catch ex As Exception

        End Try
    End Sub
    Private Sub CekQrCodeBench2()

        'tmStation2.Enabled = True
        Dim Conn As SqlConnection
        Dim SQl1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()
            SQl1 = Conn.CreateCommand
            SQl1.CommandText = "Select *FROM m_refRunningItem WHERE (QRCode='" & lblHMI2.Text & "') AND ( NEXTSCAN = 2 OR ST = 2 ) "
            'AND(NoID=" & Val(lblNoID.Text) & ")AND(St=1)"
            Reader = SQl1.ExecuteReader
            If Reader.Read Then

                Call KOSONGKAN_SEQUENCEBENCH2()

                lblQr.Text = lblHMI2.Text
                'Conn.Close()
                Dim SQL2 As New SqlCommand
                Dim Conn1 As SqlConnection
                'Try
                Conn1 = GetConnect()
                Conn1.Open()

                SQL2 = Conn1.CreateCommand
                'SEKA CHANGE NEXT SCAN
                '                SQL2.CommandText = "UPDATE m_refRunningItem SET ScanQRSt2=1,QtyRunSt2=1,St=2,QtyRunSt3=0,QtyRunSt4=0,ScanQrSt3=0,ScanQrSt4=0,Status='R',NextScan=3  WHERE (QRCode='" & lblHMI2.Text & "')AND(NoID=" & Val(lblNoID.Text) & ")AND(St=1)"
                SQL2.CommandText = "UPDATE m_refRunningItem SET ScanQRSt2=1,QtyRunSt2=1,St=2,QtyRunSt3=0,QtyRunSt4=0,ScanQrSt3=0,ScanQrSt4=0,Status='R'   WHERE (QRCode='" & lblHMI2.Text & "')AND(NoID=" & Val(lblNoID.Text) & ")AND(St=1)"
                SQL2.ExecuteReader()
                Conn1.Close()

                tmrQtyB2.Enabled = True

                lblQr.Text = lblHMI2.Text
                'Kirim ke %MW8001 Nilai 5 (PLC 107)
                Dim StartAddress As Integer = 8001
                Dim Values(0) As Integer
                Values(0) = 5
                modbusClient1.WriteMultipleRegisters(StartAddress, Values)

                '=========== MATIKAN DI HMI 2
                Dim Add2703 As Integer = 2703
                Dim Val2703(0) As Integer
                Val2703(0) = 0
                modbusClient1.WriteMultipleRegisters(Add2703, Val2703)

                'Call MasukQtyRunningMWBench2()
                'Call KOSONGKAN_SEQUENCEBENCH2()

                Call START_SEQUENCEBENCH2()

                Call MasukSerialNoBench2()

                lblLabel.Text = "Station 2: PSN Correct!"
                lblLabel2.Text = lblHMI2.Text

                modbusClient1.WriteSingleRegister(MW_COMPLETE_BENCH_2, 0)
                'txtScan.Text = ""
                'Exit Sub
            Else
                Call SALAHQR_BENCH2()

            End If

            '            Call MasukSerialNoBench2()


            Conn.Close()
            Reader.Close()

            'tmStation2.Enabled = False

        Catch ex As Exception
            'MessageBox.Show(ex.Message.ToString)
        End Try
        'ElseIf lblStation.Text = "2" Then
        'Call LihatScanQRSt2()
        'End If
    End Sub
    Private Sub SAVE_FAILBENCH2()
        Dim CONN As SqlConnection
        Dim SQL1 As New SqlCommand
        Try
            CONN = GetConnect()
            CONN.Open()
            SQL1 = CONN.CreateCommand
            SQL1.CommandText = "UPDATE m_Fail SET St2='" & lblHMI2.Text & "' WHERE (NoID=1)"
            SQL1.ExecuteReader()
            CONN.Close()

        Catch ex As Exception

        End Try
    End Sub

    Private Sub SALAHQR_BENCH2()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader

        Dim Add2703 As Integer = 2703
        Dim Val2703(0) As Integer

        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "Select *FROM m_refRunningItem WHERE ( QRCode='" & lblHMI2.Text & "' AND NEXTSCAN = 2 ) " 'AND(NoID=" & Val(lblNoID.Text) & ")"
            Reader = SQL1.ExecuteReader
            If Reader.Read Then
                Dim Posisi As Integer = Val(Reader("NextScan"))
                lblLabel.Text = "Station 2: WRONG PSN"
                lblLabel2.Text = "PSN: " & lblHMI2.Text & " BACK TO STATION " & Posisi

                '=========== TAMPILKAN DI HMI 2
              
                Val2703(0) = 1
                modbusClient1.WriteMultipleRegisters(Add2703, Val2703)

            Else
                'lblHMI2.Text = "0"
                lblLabel2.Text = lblHMI2.Text
                lblLabel.Text = "Station 2: WRONG PSN"
                txtScan.Text = ""

                Val2703(0) = 1
                modbusClient1.WriteMultipleRegisters(Add2703, Val2703)
                'Exit Sub
            End If
            Conn.Close()
            Reader.Close()

        Catch ex As Exception

        End Try
    End Sub

    Private Sub SALAHQR_BENCH3()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader

        '=========== TAMPILKAN DI HMI 3
        Dim Add2704 As Integer = 2704
        Dim Val2704(0) As Integer

        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "Select *FROM m_refRunningItem WHERE (QRCode='" & lblHMI3.Text & "') AND NEXTSCAN = 3  " 'AND(NoID=" & Val(lblNoID.Text) & ")"
            Reader = SQL1.ExecuteReader
            If Reader.Read Then
                Dim Posisi As Integer = Val(Reader("NextScan"))
                lblLabel.Text = "Station 3: WRONG PSN"
                lblLabel2.Text = "PSN: " & lblHMI3.Text & " BACK TO STATION " & Posisi

             
                Val2704(0) = 1
                modbusClient2.WriteMultipleRegisters(Add2704, Val2704)

            Else
                'lblHMI2.Text = "0"
                lblLabel2.Text = lblHMI3.Text
                lblLabel.Text = "Station 3: WRONG PSN"
                txtScan.Text = ""

                Val2704(0) = 1
                modbusClient2.WriteMultipleRegisters(Add2704, Val2704)
                'Exit Sub
            End If
            Conn.Close()
            Reader.Close()

        Catch ex As Exception

        End Try
    End Sub

    Private Sub SALAHQR_BENCH4()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader

        '=========== TAMPILKAN DI HMI 4
        Dim Add2705 As Integer = 2705
        Dim Val2705(0) As Integer

        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "Select *FROM m_refRunningItem WHERE (QRCode='" & lblHMI4.Text & "') AND NEXTSCAN = 4  " 'AND(NoID=" & Val(lblNoID.Text) & ")"
            Reader = SQL1.ExecuteReader
            If Reader.Read Then
                Dim Posisi As Integer = Val(Reader("NextScan"))
                lblLabel.Text = "Station 4: WRONG PSN"
                lblLabel2.Text = "PSN: " & lblHMI4.Text & " BACK TO STATION " & Posisi

              
                Val2705(0) = 1
                modbusClient2.WriteMultipleRegisters(Add2705, Val2705)

            Else
                'lblHMI2.Text = "0"
                lblLabel2.Text = lblHMI4.Text
                lblLabel.Text = "Station 4: WRONG PSN"
                txtScan.Text = ""

                Val2705(0) = 1
                modbusClient2.WriteMultipleRegisters(Add2705, Val2705)
                'Exit Sub
            End If
            Conn.Close()
            Reader.Close()

        Catch ex As Exception

        End Try
    End Sub

    Private Sub CekQrCodeBench3()
        'tmStation2.Enabled = True

        Dim Conn As SqlConnection
        Dim SQl1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Call KOSONGKAN_SEQUENCEBENCH3()

            Conn = GetConnect()
            Conn.Open()
            SQl1 = Conn.CreateCommand
            SQl1.CommandText = "Select *FROM m_refRunningItem WHERE (QRCode='" & lblHMI3.Text & "')AND(NoID=" & Val(lblNoID.Text) & ") AND ( NEXTSCAN=3 OR ST = 3)"
            Reader = SQl1.ExecuteReader
            If Reader.Read Then
                lblQr.Text = lblHMI3.Text
                'Conn.Close()
                Dim SQL2 As New SqlCommand
                Dim Conn1 As SqlConnection
                'Try
                Conn1 = GetConnect()
                Conn1.Open()

                SQL2 = Conn1.CreateCommand
                '                SQL2.CommandText = "UPDATE m_refRunningItem SET ScanQRSt3=1,QtyRunSt3=1,ScanQrSt4=0,QtyRunSt4=0,St=3,Status='R',NextScan=4  WHERE (QRCode='" & lblHMI3.Text & "') AND (NoID=" & Val(lblNoID.Text) & ") AND (St=2)"
                SQL2.CommandText = "UPDATE m_refRunningItem SET ScanQRSt3=1,QtyRunSt3=1,ScanQrSt4=0,QtyRunSt4=0,St=3,Status='R'   WHERE (QRCode='" & lblHMI3.Text & "') AND (NoID=" & Val(lblNoID.Text) & ") AND (St=2)"
                SQL2.ExecuteReader()
                Conn1.Close()

                '======== Hidup Pokapick Station 3 (107)
                Dim StartAddress As Integer = 8002
                Dim Values(0) As Integer
                Values(0) = 5
                modbusClient2.WriteMultipleRegisters(StartAddress, Values)

                lblQr.Text = lblHMI3.Text
                'Conn.Close()
                'Call UpdateStatusRunning()
                'Call HitungRunQty()
                tmrQtyB3.Enabled = True
                'PLC 107

                '=========== MATIKAN DI HMI 3
                Dim Add2704 As Integer = 2704
                Dim Val2704(0) As Integer
                Val2704(0) = 0
                modbusClient2.WriteMultipleRegisters(Add2704, Val2704)

                'Call MasukQtyRunningMWBench3()
                'Call KOSONGKAN_SEQUENCEBENCH3()

                Call START_SEQUENCEBENCH3()
                Call MasukSerialNoBench3()


                lblLabel.Text = "Station 3: PSN Correct!"
                lblLabel2.Text = lblHMI3.Text

                modbusClient2.WriteSingleRegister(MW_COMPLETE_BENCH_3, 0)

                'Exit Sub

            Else

                Call SALAHQR_BENCH3()

            End If

            Conn.Close()
            Reader.Close()

            'tmStation2.Enabled = False
        Catch ex As Exception
            'MessageBox.Show(ex.Message.ToString)
        End Try
        'ElseIf lblStation.Text = "2" Then
        'Call LihatScanQRSt2()
        'End If
    End Sub
    Private Sub UpdateLastQRToHeader()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Try
            Conn = GetConnect()
            Conn.Open()

        Catch ex As Exception

        End Try
    End Sub

    Sub CekQrCodeBench4()
        'tmStation2.Enabled = True

        Dim Conn As SqlConnection
        Dim SQl1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()
            SQl1 = Conn.CreateCommand
            SQl1.CommandText = "Select *FROM m_refRunningItem WHERE (QRCode='" & lblHMI4.Text & "') AND(St=4 OR NEXTSCAN = 4)" '(Model='" & Val(lblNoID.Text) & " AND QRCode='" & lblHMI4.Text & "' AND St=3)"
            '            SQl1.CommandText = "Select *FROM m_refRunningItem WHERE ( QRCode='" & TextBox3.Text & "') AND (NoID=" & Val(lblNoID.Text) & ")AND( ST = 4 OR NEXTSCAN = 4 )" '(Model='" & Val(lblNoID.Text) & " AND QRCode='" & lblHMI4.Text & "' AND St=3)"
            Reader = SQl1.ExecuteReader

            If Reader.Read Then

                Call KOSONGKAN_SEQUENCEBENCH4()

                'MessageBox.Show("Test")
                LastQR = Reader("QRCode").ToString
                lblLastQR.Text = lblHMI4.Text

                'Conn.Close()
                Dim SQL2 As New SqlCommand
                Dim Conn1 As SqlConnection
                'Try
                Conn1 = GetConnect()
                Conn1.Open()

                SQL2 = Conn1.CreateCommand
                '                SQL2.CommandText = "UPDATE m_refRunningItem SET ScanQRSt4=1,QtyRunSt4=1,St=4,Status='R',NextScan=5  WHERE (QRCode='" & lblHMI4.Text & "') AND (NoID=" & Val(lblNoID.Text) & ") AND (St=3)"
                SQL2.CommandText = "UPDATE m_refRunningItem SET ScanQRSt4=1,QtyRunSt4=1,St=4,Status='R'  WHERE (QRCode='" & lblHMI4.Text & "') AND (NoID=" & Val(lblNoID.Text) & ") AND (St=3)"
                SQL2.ExecuteReader()
                Conn1.Close()

                'PLC 107
                Dim StartAddress As Integer = 8003
                Dim Values(0) As Integer
                Values(0) = 5
                modbusClient2.WriteMultipleRegisters(StartAddress, Values)

                '=========== MATIKAN DI HMI 4
                Dim Add2705 As Integer = 2705
                Dim Val2705(0) As Integer
                Val2705(0) = 0
                modbusClient2.WriteMultipleRegisters(Add2705, Val2705)

                'Call MasukQtyRunningMWBench4()
                'Call KOSONGKAN_SEQUENCEBENCH4()


                Call START_SEQUENCEBENCH4()


                Call HapusPCBdiHMI()
                Call MasukSerialNoBench4()
                'tmrQtyB4.Enabled = True

                txtScan.Text = ""

                lblLabel.Text = "Station 4: PSN Correct!"
                lblLabel2.Text = lblHMI4.Text

                modbusClient2.WriteSingleRegister(MW_COMPLETE_BENCH_4, 0)

                'Call HapusPCBdiHMI()

                FINISH_SCAN_QR = True

            Else

                FINISH_SCAN_QR = False
                tmrPick4.Enabled = False
                Call SALAHQR_BENCH4()

            End If

            Conn.Close()
            Reader.Close()

            'tmStation2.Enabled = False

        Catch ex As Exception
            'MessageBox.Show(ex.Message.ToString)
        End Try

    End Sub

    Sub UpdateStatusRunning()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "UPDATE m_refRunning SET Status='R' WHERE (NoID=" & Val(lblNoID.Text) & ")"
            SQL1.ExecuteReader()
            Conn.Close()

        Catch ex As Exception

        End Try
    End Sub
    Sub BuatQRCodeBaru()
        Dim Conn As SqlConnection
        Dim SQl1 As New SqlCommand
        Dim Reader As SqlDataReader
        Dim Awal As String = "8B"
        Dim Model As String = lblRef.Text
        'Dim ItemNo As String = txtItem.Text
        Dim Tahun = Format(Now, "yy").ToString
        Dim Week = DatePart("ww", Now, vbSunday)


        Dim NomorBarcode As String

        Try
            Conn = GetConnect()
            Conn.Open()
            SQl1 = Conn.CreateCommand
            SQl1.CommandText = "SELECT COUNT(ScanQRSt1) AS ScanQRSt1 FROM m_refrunningitem WHERE (NoID=" & Val(lblNoID.Text) & ")"
            'SQl1.CommandText = "SELECT RIGHT(QRCode,5) AS QRCode FROM m_refRunningItem WHERE (NoID=" & Val(lblNoID.Text) & ") ORDER BY RIGHT(QRCode,5) DESC"
            Reader = SQl1.ExecuteReader
            If Reader.Read Then
                lblQrItem.Text = Format(Reader("ScanQRSt1") + 1, "00000")
                NomorBarcode = Awal & "," & Model & "," & Tahun & Week & "," & lblQrItem.Text
                lblQr.Text = NomorBarcode

            Else
                lblQrItem.Text = "00001"
                NomorBarcode = Awal & "," & Model & "," & Tahun & Week & "," & lblQrItem.Text
                lblQr.Text = NomorBarcode

            End If
            Conn.Close()
            Reader.Close()

            'NomorBarcode = Awal & "," & Model & "," & Tahun & Week & "," & ItemNo

            'txtMWSerial.Text = NomorBarcode
            'f_utama.lblQr.Text = NomorBarcode
        Catch ex As Exception

        End Try
    End Sub
    Sub NoItem()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()

            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "SELECT TOP 1 *FROM m_refRunningItem WHERE (NoID=" & Val(lblNoID.Text) & ") ORDER BY Item DESC"
            Reader = SQL1.ExecuteReader

            If Reader.Read Then
                lblNoItem.Text = Val(Reader("Item")) + 1
            Else
                lblNoItem.Text = "1"
            End If
            Conn.Close()
            Reader.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error Item", MessageBoxButtons.OK)
        End Try
    End Sub
    Sub MasukItemBaru1()
        Dim Conn As SqlConnection
        Dim SQl1 As New SqlCommand
        Dim Tanggal = Format(Now, "yyyy/MM/dd HH:mm:ss").ToString

        Try
            Call NoItem()
            Call BuatQRCodeBaru()

            Conn = GetConnect()
            Conn.Open()
            SQl1 = Conn.CreateCommand
            SQl1.CommandText = "INSERT INTO m_refRunningItem" _
                            & "(NoID,Item,SONo,QRCode,Model,Qty,QtyRun,ScanQRSt1,iDate,Status,St)" _
                            & "VALUES (" & Val(lblNoID.Text) & "," & Val(lblNoItem.Text) & ",'" & lblSO.Text & "','" & lblQr.Text & "','" & lblRef.Text & "'," & Val(lblQty.Text) & ",0,0,'" & Tanggal & "','R',1)"
            SQl1.ExecuteReader()
            Conn.Close()

            'Call HitungRunQty()
            'Call MasukQRCodeMW()
            'Call MasukQtyRunningMWBench1()

            txtScan.Text = ""
            txtScan.Focus()

            tmrSequence.Enabled = True

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error Item Baru", MessageBoxButtons.OK)
        End Try
    End Sub
    Sub MasukQRCodeMW1()
        Dim cb() As Byte
        'txtConvert.Text = ""
        If txtScan.Text <> "" Then
            cb = ASCIIEncoding.ASCII.GetBytes(txtScan.Text)
            Dim MWval_reff(29) As Integer 'jumlah address hingga 20
            For x As Integer = 0 To cb.Length - 1
                'txtConvert.AppendText(cb(x).ToString & "")
                'txtFrom.Text = ""
                'txtTo.AppendText(cb(x).ToString & "")
                'Dim arrayRef() as Char=cb(x).to
                'MessageBox.Show(cb(x).ToString)
                MWval_reff(x) = Convert.ToInt64(cb(x))
                Try
                    Dim StartAddress As Integer = 500 'mulai dari MW400
                    modbusClient1.WriteMultipleRegisters(StartAddress, MWval_reff)
                Catch ex As Exception

                End Try
            Next

        End If
    End Sub
    Sub HitungRunQty()
        Dim Conn As SqlConnection
        Dim SQl1 As New SqlCommand
        Dim Reader As SqlDataReader
        Dim Jumlah As Integer

        Try
            Conn = GetConnect()
            Conn.Open()
            SQl1 = Conn.CreateCommand
            SQl1.CommandText = "SELECT SUM(QtyRunSt4) AS tQtyRun FROM m_refRunningItem WHERE (NoID=" & Val(lblNoID.Text) & ")"
            Reader = SQl1.ExecuteReader
            If Reader.Read Then
                Jumlah = Val(Reader("tQtyRun"))
                lblRunQty.Text = Jumlah
                Call UpdateRunQty()
                Call TargetComplete()
            Else
                lblRunQty.Text = "00"
            End If
            Conn.Close()
            Reader.Close()

        Catch ex As Exception
            'lblRunQty.Text = "00"
        End Try
    End Sub
    Private Sub UpdateRunQty()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Jumlah As Integer = Val(lblRunQty.Text)

        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "UPDATE m_refRunning SET RunQty=" & Jumlah & " WHERE (NoID=" & Val(lblNoID.Text) & ")"
            SQL1.ExecuteReader()
            Conn.Close()

            'Call TargetComplete()

        Catch ex As Exception

        End Try
    End Sub
    Private Sub TargetComplete()
        Dim Target As Integer = Val(lblTarget.Text)
        Dim jumRunning As Integer = Val(lblRunQty.Text)

        If Target = jumRunning Then
            Dim Conn As SqlConnection
            Dim SQL1 As New SqlCommand
            Try
                Conn = GetConnect()
                Conn.Open()
                SQL1 = Conn.CreateCommand
                SQL1.CommandText = "UPDATE m_refRunning SET Status='R',RunQty=" & jumRunning & " WHERE (NoID=" & Val(lblNoID.Text) & ")"
                SQL1.ExecuteReader()
                Conn.Close()

                'SerialBench2.Close()
                'SerialBench3.Close()
                'SerialBench4.Close()

                lblLabel.Text = "TARGET COMPLETE"

            Catch ex As Exception

            End Try
            'Else

        End If
    End Sub
    Sub MasukItemRunning()
        Dim Conn As SqlConnection
        Dim SQl1 As New SqlCommand
        Dim TargetQty As Integer
        TargetQty = Val(lblQty.Text)
        Dim Tanggal = Format(Now, "yyyy/MM/dd HH:mm:ss").ToString

        Try
            Conn = GetConnect()
            Conn.Open()
            SQl1 = Conn.CreateCommand
            SQl1.CommandText = "INSERT INTO m_refRunningItem" _
                            & "(NoID,Item,QrCode,Model,Qty,iDate,Status,ScanQRSt1,PCBNo)" _
                            & "VALUES (" & Val(lblNoID.Text) & "," & Val(lblNoItem.Text) & ",'" & lblQr.Text & "','" & lblRef.Text & "'," & TargetQty & ",'" & Tanggal & "','R',1,'" & lblPCB.Text & "')"
            SQl1.ExecuteReader()
            Conn.Close()


        Catch ex As Exception

        End Try
    End Sub
    Sub LihatStatusScanQRSt1()
        Dim Conn As SqlConnection
        Dim SQl1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()
            SQl1 = Conn.CreateCommand
            SQl1.CommandText = "SELECT *FROM m_refRunning WHERE (Model='" & lblRef.Text & "')AND(ScanQRSt1=1)AND(NoID=" & Val(lblNoID.Text) & ")"
            Reader = SQl1.ExecuteReader
            If Reader.Read Then
                'Call LihatScanPartSt2()
            Else
                lblLabel.Text = "SCAN QR CODE Station 1"

            End If
            Conn.Close()
            Reader.Close()

        Catch ex As Exception

        End Try
    End Sub

    Sub StatusProgressBench1()

        Dim StartAddress As Integer = 1700
        Dim Values(0) As Integer
        Values(0) = 1
        modbusClient1.WriteMultipleRegisters(StartAddress, Values)

    End Sub
    Sub StatusProgressBench2()

        Dim StartAddress As Integer = 2700
        Dim Values(0) As Integer
        Values(0) = 1
        modbusClient2.WriteMultipleRegisters(StartAddress, Values)

    End Sub
    Sub StatusProgressBench3()

        Dim StartAddress As Integer = 3700
        Dim Values(0) As Integer
        Values(0) = 1
        modbusClient2.WriteMultipleRegisters(StartAddress, Values)

    End Sub
    Sub StatusProgressBench4()

        Dim StartAddress As Integer = 3700
        Dim Values(0) As Integer
        Values(0) = 1
        modbusClient2.WriteMultipleRegisters(StartAddress, Values)

    End Sub
    Sub StatusfailBench1()

        Dim StartAddress As Integer = 1701
        Dim Values(0) As Integer
        Values(0) = 1
        modbusClient1.WriteMultipleRegisters(StartAddress, Values)

    End Sub
    Sub StatusfailBench2()

        Dim StartAddress As Integer = 2701
        Dim Values(0) As Integer
        Values(0) = 1
        modbusClient2.WriteMultipleRegisters(StartAddress, Values)

    End Sub
    Sub StatusfailBench3()

        Dim StartAddress As Integer = 3701
        Dim Values(0) As Integer
        Values(0) = 1
        modbusClient2.WriteMultipleRegisters(StartAddress, Values)

    End Sub
    Sub StatusfailBench4()

        Try
            Dim StartAddress As Integer = 4701
            Dim Values(0) As Integer
            Values(0) = 1
            modbusClient2.WriteMultipleRegisters(StartAddress, Values)
        Catch ex As Exception
            MessageBox.Show("PLC Poor Connection", "PLC Error", MessageBoxButtons.OK)
        End Try

    End Sub
    Sub StatusCompleteBench1()

        Try
            Dim StartAddress As Integer = 1702
            Dim Values(0) As Integer
            Values(0) = 1
            modbusClient1.WriteMultipleRegisters(StartAddress, Values)

        Catch ex As Exception
            MessageBox.Show("PLC Poor Connection", "PLC Error", MessageBoxButtons.OK)
        End Try

        'Call TypeModel1Bench1() 'Hidupkan Pokapick Model1

    End Sub

    Sub StatusCompleteBench2()

        Try
            Dim StartAddress As Integer = 2702
            Dim Values(0) As Integer
            Values(0) = 1
            modbusClient1.WriteMultipleRegisters(StartAddress, Values)
        Catch ex As Exception

        End Try

        'Call TypeModel1Bench1() 'Hidupkan Pokapick Model1

    End Sub
    Sub StatusCompleteBench3()

        Try
            Dim StartAddress As Integer = 3702
            Dim Values(0) As Integer
            Values(0) = 1
            modbusClient2.WriteMultipleRegisters(StartAddress, Values)
        Catch ex As Exception

        End Try

        'Call TypeModel1Bench1() 'Hidupkan Pokapick Model1

    End Sub
    Sub StatusCompleteBench4()

        Try
            Dim StartAddress As Integer = 4702
            Dim Values(0) As Integer
            Values(0) = 1
            modbusClient2.WriteMultipleRegisters(StartAddress, Values)

        Catch ex As Exception

        End Try

        'Call TypeModel1Bench1() 'Hidupkan Pokapick Model1

    End Sub

    Sub HapusStatusPartBench2()
        'Status Progress
        Dim StartAddress As Integer = 2700
        Dim Values(0) As Integer
        Values(0) = 0
        modbusClient1.WriteMultipleRegisters(StartAddress, Values)

        'Status Fail
        Dim StartAddress1 As Integer = 2701
        Dim Values1(0) As Integer
        Values1(0) = 0
        modbusClient1.WriteMultipleRegisters(StartAddress1, Values1)

        'Status Complete
        Dim StartAddress2 As Integer = 2702
        Dim Values2(0) As Integer
        Values2(0) = 0
        modbusClient1.WriteMultipleRegisters(StartAddress2, Values2)

    End Sub
    Sub HapusStatusPartBench3()
        'Status Progress
        Dim StartAddress As Integer = 3700
        Dim Values(0) As Integer
        Values(0) = 0
        modbusClient2.WriteMultipleRegisters(StartAddress, Values)

        'Status Fail
        Dim StartAddress1 As Integer = 3701
        Dim Values1(0) As Integer
        Values1(0) = 0
        modbusClient2.WriteMultipleRegisters(StartAddress1, Values1)

        'Status Complete
        Dim StartAddress2 As Integer = 3702
        Dim Values2(0) As Integer
        Values2(0) = 0
        modbusClient2.WriteMultipleRegisters(StartAddress2, Values2)

    End Sub
    Sub HapusStatusPartBench4()
        'Status Progress
        Dim StartAddress As Integer = 4700
        Dim Values(0) As Integer
        Values(0) = 0
        modbusClient2.WriteMultipleRegisters(StartAddress, Values)

        'Status Fail
        Dim StartAddress1 As Integer = 4701
        Dim Values1(0) As Integer
        Values1(0) = 0
        modbusClient2.WriteMultipleRegisters(StartAddress1, Values1)

        'Status Complete
        Dim StartAddress2 As Integer = 4702
        Dim Values2(0) As Integer
        Values2(0) = 0
        modbusClient2.WriteMultipleRegisters(StartAddress2, Values2)

    End Sub
    Sub PostingScan()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "UPDATE s_masterdetail SET Sc=1 WHERE (SONo='" & lblSO.Text & "')AND(ListPartNo='" & txtScan.Text & "')"
            SQL1.ExecuteReader()
            Conn.Close()

        Catch ex As Exception

        End Try
    End Sub
    Sub PassScanSt1()

    End Sub
    Sub JumlahScanS1()
        Dim Conn As SqlConnection
        Dim SQl1 As New SqlCommand
        Dim Reader As SqlDataReader
        Dim Jumlah As Double

        Try
            Conn = GetConnect()
            Conn.Open()
            SQl1 = Conn.CreateCommand
            SQl1.CommandText = "SELECT COUNT(Sc) AS tSc FROM s_masterdetail WHERE (SONo='" & lblSO.Text & "')AND(Sc=1)"
            Reader = SQl1.ExecuteReader
            If Reader.Read Then
                Jumlah = Format(Reader("tSc"), "#,##0").ToString
                'lblSt1Scan.Text = Jumlah
            Else
                'lblSt1Scan.Text = "00"
            End If
            Conn.Close()
            Reader.Close()

        Catch ex As Exception

        End Try
    End Sub
    Sub SimpanRunning()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim SQL2 As New SqlCommand

        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "UPDATE m_refrunning SET Status='S',BarCode='" & lblQr.Text & "',RunQty=" & Val(lblRunQty.Text) & " WHERE (NoID=" & Val(lblNoID.Text) & ")"
            SQL1.ExecuteReader()
            Conn.Close()

            Conn.Open()
            SQL2 = Conn.CreateCommand
            SQL2.CommandText = "UPDATE m_refPartList SET Scan=0 WHERE (Model='" & lblRef.Text & "')"
            SQL2.ExecuteReader()
            Conn.Close()

            Call FormBlank()

            '======= Kosongkan WM350
            Dim StartAddress As Integer = 350
            Dim Values(0) As Integer
            Values(0) = 0
            modbusClient1.WriteMultipleRegisters(StartAddress, Values)

            Dim StartAddress1 As Integer = 106
            Dim Values1(0) As Integer
            Values1(0) = 0
            modbusClient1.WriteMultipleRegisters(StartAddress1, Values1)

            Dim StartAddress2 As Integer = 107
            Dim Values2(0) As Integer
            Values2(0) = 0
            modbusClient2.WriteMultipleRegisters(StartAddress2, Values2)


        Catch ex As Exception

        End Try
    End Sub
    Private Sub f_utama_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Q AndAlso e.Modifiers = Keys.Control) Then
            'If lblSO.Text <> "" Then
            'MessageBox.Show("Save Material Running First!")
            'Else
            End
            'End If
        End If
        If e.KeyCode = Keys.Home Then
            Dim Scanner As New TestScanner
            Scanner.ShowDialog()

        End If
    End Sub

    Sub FormBlank()
        lblNoID.Text = ""
        lblSO.Text = ""
        lblRef.Text = ""
        lblBuzzer.Text = ""
        lblFoot.Text = ""
        lblItem.Text = ""
        lblLayer.Text = ""
        lblType.Text = ""
        lblVolt.Text = ""
        lblQty.Text = ""
        lblTarget.Text = "00"
        lblRunQty.Text = "00"
        tmrSequence.Enabled = False

        txtConvPart.Text = ""
        lblLabel.Text = "WAIT TO PROCESS..."

        lblSt1TotalMaterial.Text = "0"
        lblSt2TotalMaterial.Text = "0"
        lblSt3TotalMaterial.Text = "0"
        lblSt4TotalMaterial.Text = "0"

        lblSt4ScanTotal.Text = "0"
        lblSt3ScanTotal.Text = "0"
        lblSt2ScanTotal.Text = "0"
        lblSt1ScanTotal.Text = "0"

        lblSt1ScanTotal.Text = "00"
        lblSt2ScanTotal.Text = "00"
        lblSt3ScanTotal.Text = "00"
        lblSt4ScanTotal.Text = "00"
        lblSequence.Text = "0"
        lblQr.Text = ""

        lblSt1ScanTotal.Visible = False
        lblSt2ScanTotal.Visible = False
        lblSt3ScanTotal.Visible = False
        lblSt4ScanTotal.Visible = False

        LvSt1.Items.Clear()
        LvSt2.Items.Clear()
        LvSt3.Items.Clear()
        LvSt4.Items.Clear()

        'Call HapusStatusPartBench1()
        Call HapusStatusPartBench2()
        Call HapusStatusPartBench3()
        Call HapusStatusPartBench4()


        'lblList1.Visible = False
        'lblList2.Visible = False
        'lblList3.Visible = False
        'lblList4.Visible = False
        'lblBarcodeNo.Visible = False

        'lblSt1Pass.Visible = False


    End Sub

    Private Sub f_utama_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        'rs.ResizeAllControls(Me)
    End Sub
    Sub TandaiScan()
        Dim Conn As SqlConnection
        Dim SQl1 As New SqlCommand
        Try
            Conn = GetConnect()
            Conn.Open()
            SQl1 = Conn.CreateCommand
            SQl1.CommandText = "UPDATE m_refPartList SET Scan=1 WHERE (Model='" & lblRef.Text & "')AND(Item=" & Val(lblPartItem.Text) & ")"
            SQl1.ExecuteReader()
            Conn.Close()

            Call dpStation1()
            Call dpStation2()
            Call dpStation3()
            Call dpStation4()

        Catch ex As Exception

        End Try
    End Sub
    Private Sub txtScan_KeyDown(sender As Object, e As KeyEventArgs) Handles txtScan.KeyDown
        If e.KeyCode = Keys.Enter Then

            If lblSequence.Text = "1" And txtScan.Text.Length > 6 Then
                'Call ScanPart()
                'tmrPrint.Enabled = False
            ElseIf lblSequence.Text = "2" And txtScan.Text.Length > 15 Then


            Else
                lblLabel.Text = "SELECT REFERENCE NUMBER!"
                txtScan.Text = ""
            End If

        End If
    End Sub

    Private Sub LvSt1_KeyDown(sender As Object, e As KeyEventArgs) Handles LvSt1.KeyDown
        If e.KeyCode = Keys.Escape Then
            txtScan.Focus()
        End If
    End Sub

    Sub MasukSerialNoBench2()
        If lblStation.Text = "1" Or lblStation.Text = "2" Then
            Dim cb() As Byte
            'txtConvert.Text = ""
            If lblHMI2.Text.Length > 6 Then
                cb = ASCIIEncoding.ASCII.GetBytes(lblHMI2.Text)
                Dim MWval_reff(29) As Integer 'jumlah address hingga 20
                For x As Integer = 0 To cb.Length - 1
                    'txtConvert.AppendText(cb(x).ToString & "")
                    'txtFrom.Text = ""
                    MWval_reff(x) = Convert.ToInt64(cb(x))
                    Try
                        Dim StartAddress As Integer = 6100 'mulai dari MW400
                        modbusClient1.WriteMultipleRegisters(StartAddress, MWval_reff)
                    Catch ex As Exception

                    End Try
                Next

            End If
        Else

        End If
    End Sub
    Sub MasukSerialNoBench3()
        If lblStation.Text = "1" Or lblStation.Text = "2" Then
            Dim cb() As Byte
            'txtConvert.Text = ""
            If lblHMI3.Text.Length > 5 Then
                cb = ASCIIEncoding.ASCII.GetBytes(lblHMI3.Text)
                Dim MWval_reff(29) As Integer 'jumlah address hingga 20
                For x As Integer = 0 To cb.Length - 1
                    'txtConvert.AppendText(cb(x).ToString & "")
                    'txtFrom.Text = ""
                    'txtTo.AppendText(cb(x).ToString & "")
                    'Dim arrayRef() as Char=cb(x).to
                    'MessageBox.Show(cb(x).ToString)
                    MWval_reff(x) = Convert.ToInt64(cb(x))
                    Try
                        Dim StartAddress As Integer = 6200 'mulai dari MW400
                        modbusClient2.WriteMultipleRegisters(StartAddress, MWval_reff)
                    Catch ex As Exception

                    End Try
                Next

            End If
        Else

        End If
    End Sub
    Sub MasukSerialNoBench4()
        If lblStation.Text = "1" Or lblStation.Text = "2" Then
            Dim cb() As Byte
            'txtConvert.Text = ""
            If lblHMI4.Text.Length > 6 Then
                cb = ASCIIEncoding.ASCII.GetBytes(lblHMI4.Text)
                Dim MWval_reff(29) As Integer 'jumlah address hingga 20
                For x As Integer = 0 To cb.Length - 1
                    'txtConvert.AppendText(cb(x).ToString & "")
                    MWval_reff(x) = Convert.ToInt64(cb(x))
                    Try
                        Dim StartAddress As Integer = 6300 'mulai dari MW400
                        modbusClient2.WriteMultipleRegisters(StartAddress, MWval_reff)
                    Catch ex As Exception

                    End Try
                Next

            End If
        Else

        End If
    End Sub

    Sub CekStatusPrint()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader
        Dim Status As Integer
        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "SELECT *FROM m_refrunningitem WHERE (QRCode='" & lblQr.Text & "')"
            Reader = SQL1.ExecuteReader
            If Reader.Read Then
                Status = Val(Reader("Prt"))
                If Status = 0 Then
                    Call PrintQRCode()
                ElseIf Status = 2 Then
                    Exit Sub
                End If
                'Call PrintQRCode()
            Else
                Exit Sub
            End If
            Conn.Close()
            Reader.Close()

        Catch ex As Exception

        End Try
    End Sub
    Sub UpdateStatusPrint()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "UPDATE m_refrunningitem SET Prt=1"
        Catch ex As Exception

        End Try

    End Sub
    Sub PrintQRCode() 'As Boolean
        Try
            Dim Item1Size = Val(txtItemSize.Text)
            'Item1Data = "8B,XVGM2H,1945,00001" 'lblQr.Text
            Item1Data = lblQr.Text

            Darkness = Val(txtDarknessValue.Text)
            Item1X = Val(txtXValue.Text)
            Item1Y = Val(txtYValue.Text)
            Item1Type = txtType.Text

            Dim zplStringBuffer As String = ""
            zplFile = ""
            '^BXR,5,200
            zplString(0) = String.Format("^FO {0},{1}{2},{3},200^FD{4}^FS", 70, 35, "^BXR", 4, Item1Data)

            For a As Integer = 0 To zplString.Length - 1
                zplStringBuffer = zplStringBuffer & zplString(a)
            Next

            zplFile = String.Format("^XA^MD{0}{1}^XZ", Darkness, zplStringBuffer)
            System.IO.File.WriteAllText(System.IO.Path.Combine("E:\LabelFile\LastLabel - LA01.txt"), zplFile)
            Dim path As String = System.IO.Path.Combine("E:\LabelFile\LastLabel - LA01.txt")

            zebraPrinter.PrintRawFile("ZDesigner GX430t", path)

            lblPrint.Text = "0"
            'RawPrinterHelper.SendStringToPrinter("ZDesigner GX430t", Zplfile)

            'Matikan Status Print 9000
            Dim StartAddress As Integer = 9000
            Dim Values(0) As Integer
            Values(0) = 0
            modbusClient1.WriteMultipleRegisters(StartAddress, Values)

            'tmrSequence.Enabled = False

            'Call MasukItemBaru()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error Print", MessageBoxButtons.OK)
        End Try
    End Sub
    Sub Config()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "SELECT *FROM s_Config WHERE (NoID=1)"
            Reader = SQL1.ExecuteReader
            If Reader.Read Then
                txtYValue.Text = Val(Reader("Barcode_Ypos_value")) '.ToString
                txtXValue.Text = Val(Reader("Barcode_Xpos_value")) '.ToString
                txtDarknessValue.Text = Val(Reader("Darkness_Dark_value")) '.ToString
                txtPrinterName.Text = Reader("PrinterSt1Name") '.ToString
                txtItemSize.Text = Val(Reader("ItemSize")) '.ToString
                txtType.Text = Reader("Type").ToString

            Else
                txtXValue.Text = ""
                txtYValue.Text = ""
                txtDarknessValue.Text = ""
                txtPrinterName.Text = ""
                txtItemSize.Text = ""
            End If
            Conn.Close()
            Reader.Close()

        Catch ex As Exception

        End Try
    End Sub
    Private Sub TmrPrint_Tick(sender As Object, e As EventArgs)
        Try
            'Dim StartAddress As Integer = 9000
            'Dim ReadValue() As Integer = modbusClient1.ReadHoldingRegisters(StartAddress, 1)
            'lblPrint.Text = ReadValue(0)

            'If lblPrint.Text = "1" Then
            'Call MasukItemBaru()
            'Else

            'Me.Enabled = False
            'End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub BtnTestPrint_Click(sender As Object, e As EventArgs) Handles btnTestPrint.Click
        Try
            'Call PrintQRCode()
            Dim Item1Size = Val(txtItemSize.Text)
            'Item1Data = "8B,XVGM2H,1945,00001" 'lblQr.Text
            Item1Data = lblQr.Text

            Darkness = Val(txtDarknessValue.Text)
            Item1X = Val(txtXValue.Text)
            Item1Y = Val(txtYValue.Text)
            Item1Type = txtType.Text

            Dim zplStringBuffer As String = ""
            zplFile = ""
            '^BXR,5,200
            zplString(0) = String.Format("^FO {0},{1}{2},{3},200^FD{4}^FS", 70, 35, "^BXR", 4, Item1Data)

            For a As Integer = 0 To zplString.Length - 1
                zplStringBuffer = zplStringBuffer & zplString(a)
            Next

            zplFile = String.Format("^XA^MD{0}{1}^XZ", Darkness, zplStringBuffer)
            System.IO.File.WriteAllText(System.IO.Path.Combine("E:\LabelFile\LastLabel - LA01.txt"), zplFile)
            Dim path As String = System.IO.Path.Combine("E:\LabelFile\LastLabel - LA01.txt")

            zebraPrinter.PrintRawFile("ZDesigner GX430t", path)

            'RawPrinterHelper.SendStringToPrinter("ZDesigner GX430t", Zplfile)

            txtScan.Focus()

        Catch ex As Exception

        End Try
    End Sub

    Private Sub TxtScan_TextChanged(sender As Object, e As EventArgs) Handles txtScan.TextChanged
        'Try
        'If lblSO.Text <> "" Then
        'tmrClearScan.Start()
        'Else
        'txtScan.Text = ""
        'Exit Sub
        'End If
        'Catch ex As Exception

        'End Try
    End Sub

    Private Sub TmrScan_Tick(sender As Object, e As EventArgs) Handles tmrScan.Tick
        Try

            'Baca Reference yg sedang terbuka
            Dim StartAddressref As Integer = 10001
            Dim ReadValueref() As Integer = modbusClient1.ReadHoldingRegisters(StartAddressref, 1)
            lblNoID.Text = ReadValueref(0)

            Dim StartAddress As Integer = 10002
            Dim ReadValue() As Integer = modbusClient1.ReadHoldingRegisters(StartAddress, 1)
            lblMWStation1.Text = ReadValue(0)

            Dim StartAddress3 As Integer = 10003
            Dim ReadValue3() As Integer = modbusClient1.ReadHoldingRegisters(StartAddress3, 1)
            lblMWStation2.Text = ReadValue3(0)

            If lblMWStation2.Text = "1" Then
                Call PanggilRef()
                Call BACA_MATERIAL_COMPLETE()

                'Tutup Signal dari MW Station 1
                Dim StartAddress1 As Integer = 10003
                Dim values(0) As Integer
                values(0) = 0
                modbusClient1.WriteMultipleRegisters(StartAddress1, values)
                'End If
            End If

        Catch ex As Exception

        End Try
    End Sub
    Sub Station2PP()
        Try
            Dim MW9001 As Integer
            Dim Address9001 As Integer = 9001
            Dim Read9001() As Integer = modbusClient1.ReadHoldingRegisters(Address9001, 1)
            MW9001 = Read9001(0)
            lblPP2.Text = MW9001

            If lblPP2.Text = "1" Then

                Dim Address8001 As Integer = 8001
                Dim Val8001(0) As Integer
                Val8001(0) = 0
                modbusClient1.WriteMultipleRegisters(Address8001, Val8001)
                'lblLabel.Text = "Terima"
            Else
                Exit Sub
            End If

        Catch ex As Exception

        End Try
    End Sub
    Sub Station3PP()
        Try
            Dim MW9002 As Integer
            Dim Address9002 As Integer = 9002
            Dim Read9002() As Integer = modbusClient2.ReadHoldingRegisters(Address9002, 1)
            MW9002 = Read9002(0)
            lblPP3.Text = MW9002

            If lblPP3.Text = "1" Then

                Dim Address8002 As Integer = 8002
                Dim Val8002(0) As Integer
                Val8002(0) = 0
                modbusClient2.WriteMultipleRegisters(Address8002, Val8002)
            Else
                Exit Sub
            End If
        Catch ex As Exception

        End Try
    End Sub
    Sub Station4PP()
        Try
            Dim MW9003 As Integer
            Dim Address9003 As Integer = 9003
            Dim Read9003() As Integer = modbusClient2.ReadHoldingRegisters(Address9003, 1)
            MW9003 = Read9003(0)
            lblPP4.Text = MW9003

            If lblPP4.Text = "1" Then
                Dim Address8003 As Integer = 8003
                Dim Val8003(0) As Integer
                Val8003(0) = 0
                modbusClient2.WriteMultipleRegisters(Address8003, Val8003)
            Else
                Exit Sub
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub BACA_MATERIAL_COMPLETE()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "SELECT *FROM m_refPartList WHERE (Model='" & lblRef.Text & "')AND(Scan=0)"
            Reader = SQL1.ExecuteReader
            If Reader.Read Then
                lblSequence.Text = "1"
            Else
                'Call CompleteScankeMW()
                lblSequence.Text = "2"

            End If
            Conn.Close()
            Reader.Close()

        Catch ex As Exception

        End Try
    End Sub
    Private Sub TmrSequence_Tick(sender As Object, e As EventArgs) Handles tmrSequence.Tick
        'Dim Conn As SqlConnection
        'Dim SQL1 As New SqlCommand
        'Dim Reader As SqlDataReader
        'Try
        'conn = GetConnect()
        'conn.Open()
        'SQL1 = Conn.CreateCommand
        'SQL1.CommandText = "SELECT *FROM m_refPartList WHERE (Model='" & lblRef.Text & "')AND(Scan=0)"
        'Reader = SQL1.ExecuteReader
        'If Reader.Read Then
        'lblSequence.Text = "1"
        'Else
        'Call CompleteScankeMW()
        'lblSequence.Text = "2"

        'End If
        'conn.Close()
        'Reader.Close()

        'Catch ex As Exception

        'End Try
    End Sub
    Sub CompleteScankeMW()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "SELECT *FROM m_refPartList WHERE (Model='" & lblRef.Text & "')AND(Scan=0)"
            Reader = SQL1.ExecuteReader
            If Reader.Read Then
                Exit Sub
            Else
                '================= Kirim angka 1 ke MW 106 dan 107
                Dim StartAddress As Integer = 106
                Dim Values(0) As Integer
                Values(0) = 1
                modbusClient1.WriteMultipleRegisters(StartAddress, Values)

                Dim StartAddress1 As Integer = 107
                Dim Values1(0) As Integer
                modbusClient2.WriteMultipleRegisters(StartAddress1, Values1)

            End If
            Conn.Close()
            Reader.Close()

        Catch ex As Exception

        End Try
    End Sub
    Sub CompleteScankeMWBench2()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "SELECT *FROM m_refPartList WHERE (Model='" & lblRef.Text & "' AND Station=2)AND(Scan=0)"
            Reader = SQL1.ExecuteReader
            If Reader.Read Then
                Exit Sub
            Else
                '================= Kirim angka 1 ke MW 2702
                Dim StartAddress2702 As Integer = 2702
                Dim Values2702(0) As Integer
                Values2702(0) = 1
                modbusClient1.WriteMultipleRegisters(StartAddress2702, Values2702)

                '================= Kirim angka 1 ke MW 106 dan 107
                Dim StartAddress As Integer = 106
                Dim Values(0) As Integer
                Values(0) = 1
                modbusClient1.WriteMultipleRegisters(StartAddress, Values)

                Dim StartAddress1 As Integer = 107
                Dim Values1(0) As Integer
                modbusClient2.WriteMultipleRegisters(StartAddress1, Values1)

            End If
            Conn.Close()
            Reader.Close()

        Catch ex As Exception

        End Try
    End Sub
    Sub CompleteScankeMWBench3()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "SELECT *FROM m_refPartList WHERE (Model='" & lblRef.Text & "' AND Station=3)AND(Scan=0)"
            Reader = SQL1.ExecuteReader
            If Reader.Read Then
                Exit Sub
            Else
                '================= Kirim angka 1 ke MW 3702 (PLC 107)
                Dim StartAddress3702 As Integer = 3702
                Dim Values3702(0) As Integer
                Values3702(0) = 1
                modbusClient2.WriteMultipleRegisters(StartAddress3702, Values3702)

                '================= Kirim angka 1 ke MW 106 dan 107
                'Dim StartAddress As Integer = 106
                'Dim Values(0) As Integer
                'Values(0) = 1
                'modbusClient1.WriteMultipleRegisters(StartAddress, Values)

                'Dim StartAddress1 As Integer = 107
                'Dim Values1(0) As Integer
                'modbusClient2.WriteMultipleRegisters(StartAddress1, Values1)

            End If
            Conn.Close()
            Reader.Close()

        Catch ex As Exception

        End Try
    End Sub
    Sub CompleteScankeMWBench4()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "SELECT *FROM m_refPartList WHERE (Model='" & lblRef.Text & "' AND Station=4)AND(Scan=0)"
            Reader = SQL1.ExecuteReader
            If Reader.Read Then
                Exit Sub
            Else
                '================= Kirim angka 1 ke MW 4702 (PLC 107)
                Dim StartAddress4702 As Integer = 4702
                Dim Values4702(0) As Integer
                Values4702(0) = 1
                modbusClient2.WriteMultipleRegisters(StartAddress4702, Values4702)

                '================= Kirim angka 1 ke MW 106 dan 107
                'Dim StartAddress As Integer = 106
                'Dim Values(0) As Integer
                'Values(0) = 1
                'modbusClient1.WriteMultipleRegisters(StartAddress, Values)

                'Dim StartAddress1 As Integer = 107
                'Dim Values1(0) As Integer
                'modbusClient2.WriteMultipleRegisters(StartAddress1, Values1)

            End If
            Conn.Close()
            Reader.Close()

        Catch ex As Exception

        End Try
    End Sub
    Private Sub TmrClearScan_Tick(sender As Object, e As EventArgs) Handles tmrClearScan.Tick
        'txtScan.Text = ""
        Try
            'tmrClearScan.Start()
            'If lblSequence.Text = "1" And txtScan.Text.Length > 6 Then
            'Call ScanPart()
            'tmrPrint.Enabled = False
            'ElseIf lblSequence.Text = "2" And txtScan.Text.Length > 15 Then

            'Call CekQRCode()
            'tmrPrint.Enabled = True

            'Else
            'lblLabel.Text = "SELECT REFERENCE NUMBER!"
            'txtScan.Text = ""
            'End If
            'tmrClearScan.Stop()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub f_utama_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        modbusClient1.Disconnect()
        modbusClient2.Disconnect()

    End Sub

    Private Sub SerialBench2_DataReceived(sender As Object, e As SerialDataReceivedEventArgs) Handles SerialBench2.DataReceived
        Try
            ReceivedText2(SerialBench2.ReadExisting())

        Catch ex As Exception
            'txtData.Text = ""
        End Try
    End Sub
    Private Sub ReceivedText2(ByVal [text] As String)

        Try
            If Me.lblHMI2.InvokeRequired Then
                Dim x As New SetTextCallback(AddressOf ReceivedText2)
                Me.Invoke(x, New Object() {(text)})
            Else

                lblHMI2.Text = ""
                Me.lblHMI2.Text = [text]
                'SerialBench4.DiscardInBuffer()
                'SerialBench4.DiscardOutBuffer()
            End If
        Catch ex As Exception

        End Try

    End Sub
    Private Sub SerialBench3_DataReceived(sender As Object, e As SerialDataReceivedEventArgs) Handles SerialBench3.DataReceived
        Try
            ReceivedText3(SerialBench3.ReadExisting())

        Catch ex As Exception
            'txtData.Text = ""
        End Try
    End Sub
    Private Sub ReceivedText3(ByVal [text] As String)

        If Me.lblHMI3.InvokeRequired Then
            Dim x As New SetTextCallback(AddressOf ReceivedText3)
            Me.Invoke(x, New Object() {(text)})
        Else
            lblHMI3.Text = ""
            Me.lblHMI3.Text = [text]

            'If lblSequence.Text = "1" And lblHMI3.Text.Length > 6 Then
            'Call ScanPartBench3()
            'Else
            'Exit Sub
            'End If
        End If

    End Sub
    Private Sub SerialBench4_DataReceived(sender As Object, e As SerialDataReceivedEventArgs) Handles SerialBench4.DataReceived
        Try
            ReceivedText4(SerialBench4.ReadExisting())

        Catch ex As Exception
            'txtData.Text = ""
        End Try
    End Sub
    Private Sub ReceivedText4(ByVal [text] As String)

        If Me.lblHMI4.InvokeRequired Then
            Dim x As New SetTextCallback(AddressOf ReceivedText4)
            Me.Invoke(x, New Object() {(text)})
            'MessageBox.Show("Invoke")
        Else
            lblHMI4.Text = ""
            Me.lblHMI4.Text = [text]
            'SerialBench4.DiscardInBuffer()
            'SerialBench4.DiscardOutBuffer()
            'SerialBench4.Dispose()
            'SerialBench4.Close()

        End If

    End Sub

    Private Sub lblHMI2_TextChanged(sender As Object, e As EventArgs) Handles lblHMI2.TextChanged
        Try

            If lblHMI2.Text.Length > 5 Then
                Call CekScanStation2()
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub lblHMI3_TextChanged(sender As Object, e As EventArgs) Handles lblHMI3.TextChanged
        Try
            If lblHMI3.Text.Length > 6 Then
                Call CekScanStation3()
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub lblHMI4_TextChanged(sender As Object, e As EventArgs) Handles lblHMI4.TextChanged
        Try
            'If lblHMI4.Text.Length > 6 Then
            lblBacaScan.Text = lblHMI4.Text
            Call CekScanStation4()
            'End If

        Catch ex As Exception
            'MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub TmStation2_Tick(sender As Object, e As EventArgs) Handles tmStation2.Tick

        Call Station2PP()
        Call Station3PP()
        Call Station4PP()

    End Sub
    Sub SignalPCBA()
        Try
            Dim MW5000 As Integer
            Dim Address5000 As Integer = 5000
            Dim Read5000() As Integer = modbusClient2.ReadHoldingRegisters(Address5000, 1)
            MW5000 = Read5000(0)

            lblScanPCBA.Text = MW5000

            If lblScanPCBA.Text = "1" Then
                Call CEKPCB_DOUBLE()
            Else
                Exit Sub
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub TmrPCBA_Tick(sender As Object, e As EventArgs) Handles tmrPCBA.Tick
        Try
            Call SignalPCBA()

        Catch ex As Exception

        End Try
    End Sub

    Private Sub lblHMI4_Validated(sender As Object, e As EventArgs) Handles lblHMI4.Validated
        'Call CekScanStation4()
    End Sub

    Private Sub TmrQtyB2_Tick(sender As Object, e As EventArgs) Handles tmrQtyB2.Tick
        Try
            Dim MW9001 As Integer
            Dim Address9001 As Integer = 9001
            Dim Read9001() As Integer = modbusClient1.ReadHoldingRegisters(Address9001, 1)
            MW9001 = Read9001(0)

            If MW9001 = 1 Then
                Call MasukQtyRunningMWBench2()
                tmrQtyB2.Enabled = False
            End If

            'Call MasukQtyRunningMWBench2()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TmrQtyB3_Tick(sender As Object, e As EventArgs) Handles tmrQtyB3.Tick
        Try
            Dim MW9002 As Integer
            Dim Address9002 As Integer = 9002
            Dim Read9002() As Integer = modbusClient2.ReadHoldingRegisters(Address9002, 1)
            MW9002 = Read9002(0)

            If MW9002 = 1 Then
                Call MasukQtyRunningMWBench3()
                tmrQtyB3.Enabled = False
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TmrQtyB4_Tick(sender As Object, e As EventArgs) Handles tmrQtyB4.Tick
        Dim MW9003 As Integer
        Dim Address9003 As Integer = 9003
        Dim Read9003() As Integer = modbusClient2.ReadHoldingRegisters(Address9003, 1)
        MW9003 = Read9003(0)

        If MW9003 = 1 Then
            Call MasukQtyRunningMWBench4()
            tmrQtyB4.Enabled = False
        End If
    End Sub

    Private Sub TmrPick2_Tick(sender As Object, e As EventArgs) Handles tmrPick2.Tick
        Try
            '==================================
            Dim Nilai As Integer
            Dim StartAddress As Integer = Val(txtOFF2.Text)
            Dim ReadValue() As Integer = modbusClient1.ReadHoldingRegisters(StartAddress, 1)
            Nilai = ReadValue(0)

            If Nilai = 1 Then

                Dim OFFMW As Integer = Val(txtOFF2.Text)
                Dim VALOFF(0) As Integer
                VALOFF(0) = 0
                modbusClient1.WriteMultipleRegisters(OFFMW, VALOFF)

                Dim AddMW As Integer = Val(txtMW2.Text)
                Dim ValMW(0) As Integer
                ValMW(0) = 0
                modbusClient1.WriteMultipleRegisters(AddMW, ValMW)

                Call SEQUENCE_PICKUPBENCH2()

            Else
                Exit Sub
            End If
            'lblSrew802.Text = ReadValue(0)
            '==================================

        Catch ex As Exception

        End Try
    End Sub

    Private Sub TmrPick3_Tick(sender As Object, e As EventArgs) Handles tmrPick3.Tick
        Try
            '==================================
            Dim Nilai As Integer
            Dim StartAddress As Integer = Val(txtOFF3.Text)
            Dim ReadValue() As Integer = modbusClient2.ReadHoldingRegisters(StartAddress, 1)
            Nilai = ReadValue(0)

            If Nilai = 1 Then
                Dim AddMW As Integer = Val(txtMW3.Text)
                Dim ValMW(0) As Integer
                ValMW(0) = 0
                modbusClient2.WriteMultipleRegisters(AddMW, ValMW)

                Dim OFFMW As Integer = Val(txtOFF3.Text)
                Dim VALOFF(0) As Integer
                VALOFF(0) = 0
                modbusClient2.WriteMultipleRegisters(OFFMW, VALOFF)

                Call SEQUENCE_PICKUPBENCH3()

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub TmrPick4_Tick(sender As Object, e As EventArgs) Handles tmrPick4.Tick
        Try
            '==================================
            Dim Nilai5001 As Integer
            Dim Address5001 As Integer = 5001
            Dim Read5001() As Integer = modbusClient2.ReadHoldingRegisters(Address5001, 1)
            Nilai5001 = Read5001(0)

            If Nilai5001 = 1 Then

                Dim Nilai As Integer
                Dim StartAddress As Integer = Val(txtOFF4.Text)
                Dim ReadValue() As Integer = modbusClient2.ReadHoldingRegisters(StartAddress, 1)
                Nilai = ReadValue(0)

                If Nilai = 1 Then

                    Dim AddMW As Integer = Val(txtMW4.Text)
                    Dim ValMW(0) As Integer
                    ValMW(0) = 0
                    modbusClient2.WriteMultipleRegisters(AddMW, ValMW)

                    Dim OFFMW As Integer = Val(txtOFF4.Text)
                    Dim VALOFF(0) As Integer
                    VALOFF(0) = 0
                    modbusClient2.WriteMultipleRegisters(OFFMW, VALOFF)

                    Call SEQUENCE_PICKUPBENCH4()

                End If
            Else
                'Exit Sub
                Dim Nilai As Integer
                Dim StartAddress As Integer = Val(txtOFF4.Text)
                Dim ReadValue() As Integer = modbusClient2.ReadHoldingRegisters(StartAddress, 1)
                Nilai = ReadValue(0)

                If Nilai = 1 Then

                    Dim AddMW As Integer = Val(txtMW4.Text)
                    Dim ValMW(0) As Integer
                    ValMW(0) = 0
                    modbusClient2.WriteMultipleRegisters(AddMW, ValMW)

                    Dim OFFMW As Integer = Val(txtOFF4.Text)
                    Dim VALOFF(0) As Integer
                    VALOFF(0) = 0
                    modbusClient2.WriteMultipleRegisters(OFFMW, VALOFF)

                    'Call SEQUENCE_PICKUPBENCH4()

                End If

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub TmrScrew3_Tick(sender As Object, e As EventArgs) Handles tmrScrew3.Tick
        Try

            '==================================
            Dim Nilai As Integer
            Dim StartAddress As Integer = Val(txtOFF3.Text)
            Dim ReadValue() As Integer = modbusScrewing.ReadHoldingRegisters(StartAddress, 1)
            Nilai = ReadValue(0)

            If Nilai = 1 Then
                Dim AddMW As Integer = Val(txtMW3.Text)
                Dim ValMW(0) As Integer
                ValMW(0) = 0
                modbusScrewing.WriteMultipleRegisters(AddMW, ValMW)

                Dim OFFMW As Integer = Val(txtOFF3.Text)
                Dim VALOFF(0) As Integer
                VALOFF(0) = 0
                modbusScrewing.WriteMultipleRegisters(OFFMW, VALOFF)

                'Call SEQUENCE_PICKUP()
                lblScrew4.Text = Nilai
                Call SEQUENCE_PICKUPBENCH3()
                'Else

            End If
            'lblSrew802.Text = ReadValue(0)
            '==================================
        Catch ex As Exception
            'MessageBox.Show(ex.Message, "Error 802", MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub TmrScrew4_Tick(sender As Object, e As EventArgs) Handles tmrScrew4.Tick
        Try
            '==================================
            Dim Nilai As Integer
            Dim StartAddress As Integer = Val(txtOFF4.Text)
            Dim ReadValue() As Integer = modbusScrewing.ReadHoldingRegisters(StartAddress, 1)
            Nilai = ReadValue(0)

            If Nilai = 1 Then
                Dim AddMW As Integer = Val(txtMW4.Text)
                Dim ValMW(0) As Integer
                ValMW(0) = 0
                modbusScrewing.WriteMultipleRegisters(AddMW, ValMW)

                Dim OFFMW As Integer = Val(txtOFF4.Text)
                Dim VALOFF(0) As Integer
                VALOFF(0) = 0
                modbusScrewing.WriteMultipleRegisters(OFFMW, VALOFF)

                'Call SEQUENCE_PICKUP()
                'Call SEQUENCE_PICKUPBENCH4()
                'Else

            End If
        Catch ex As Exception

        End Try
    End Sub

 
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If (e.KeyCode = Keys.Enter) Then
            ReceivedText2(TextBox1.Text)
        End If

    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If (e.KeyCode = Keys.Enter) Then
            ReceivedText3(TextBox2.Text)
        End If

    End Sub

    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If (e.KeyCode = Keys.Enter) Then
            ReceivedText4(TextBox3.Text)
        End If
    End Sub


#Region "SUB"

    Private Sub RESET_BECNH_4()
        modbusClient2.WriteSingleRegister(381, 0)
        modbusClient2.WriteSingleRegister(382, 0)
        modbusClient2.WriteSingleRegister(383, 0)
        modbusClient2.WriteSingleRegister(384, 0)
        modbusClient2.WriteSingleRegister(5000, 0)
    End Sub

    Private Sub RESET_BECNH_2()
        modbusClient1.WriteSingleRegister(370, 0)
        modbusClient1.WriteSingleRegister(371, 0)
        modbusClient1.WriteSingleRegister(372, 0)
        modbusClient1.WriteSingleRegister(373, 0)
    End Sub

    Private Sub RESET_BECNH_3()
        modbusClient2.WriteSingleRegister(374, 0)
        modbusClient2.WriteSingleRegister(375, 0)
        modbusClient2.WriteSingleRegister(376, 0)
        modbusClient2.WriteSingleRegister(377, 0)
        modbusClient2.WriteSingleRegister(378, 0)
        modbusClient2.WriteSingleRegister(379, 0)
        modbusClient2.WriteSingleRegister(380, 0)
        modbusClient2.WriteSingleRegister(392, 0)
        modbusClient2.WriteSingleRegister(393, 0)

    End Sub



#End Region


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
End Class