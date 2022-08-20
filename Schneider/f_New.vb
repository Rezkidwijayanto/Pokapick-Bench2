Imports System.Data
Imports System.Data.SqlClient
Imports EasyModbus
Imports System.Text
Imports System.Threading


'Imports SkinSoft

Public Class f_New
    Dim modBusClient As ModbusClient
    Dim modBusClient1 As ModbusClient

    Sub FormBlank()
        txtSO.Text = ""
        txtItem.Text = ""
        txtQty.Text = ""
        txtRef.Text = ""

        txtSO.Focus()

    End Sub

    Private Sub f_New_Load(sender As Object, e As EventArgs) Handles Me.Load

        Call FormBlank()
        Call ConnectPLC1()
        Call ConnectPLC2()

    End Sub
    Sub ConnectPLC1()
        Try
            modBusClient = New ModbusClient("192.168.1.106", 502)
            modBusClient.Connect()

        Catch ex As Exception

        End Try
    End Sub
    Sub ConnectPLC2()
        Try
            modBusClient1 = New ModbusClient("192.168.1.107", 502)
            modBusClient1.Connect()

        Catch ex As Exception

        End Try
    End Sub
    Sub LihatRefNo()
        Dim Conn As SqlConnection
        Dim SQl1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()
            SQl1 = Conn.CreateCommand
            SQl1.CommandText = "SELECT *FROM m_refmanual WHERE (Model='" & txtRef.Text & "')AND(PartSt='Y')"

            Reader = SQl1.ExecuteReader
            If Reader.Read Then
                'txtItem.Text = Reader("ItemNo").ToString
                'txtRef.Text = Reader("RefNo").ToString
                'txtQty.Text = Format(Reader("Qty"), "#,##0").ToString
                txtQty.Focus()

                Call BuatQRCode()

                Call dpStation1()
                Call dpStation2()
                Call dpStation3()
                Call dpStation4()

            Else
                MessageBox.Show("Reference / Model number not found!", "Error", MessageBoxButtons.OK)
            End If
            Conn.Close()
            Reader.Close()

        Catch ex As Exception

        End Try
    End Sub


    Sub KirimkeMW350()
        Try
            Dim StartAddress As Integer = 350
            Dim Values(0) As Integer
            Values(0) = 2
            modBusClient1.WriteMultipleRegisters(StartAddress, Values)
            modBusClient.WriteMultipleRegisters(StartAddress, Values)

        Catch ex As Exception

        End Try

    End Sub
    Private Sub txtSO_KeyDown(sender As Object, e As KeyEventArgs) Handles txtSO.KeyDown
        If e.KeyCode = Keys.Enter Then
            If txtSO.Text = "" Then
                MessageBox.Show("Masukan SO number", "Error", MessageBoxButtons.OK)
                txtSO.Focus()
            Else
                Call BuatQRCode()
                txtItem.Focus()
                Call BersihkanSerialNo()

                'Call LihatSO()
                'txtQty.Focus()
            End If
        End If
    End Sub
    Sub NoItemRunning()
        Dim Conn As SqlConnection
        Dim SQl1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()
            SQl1 = Conn.CreateCommand
            SQl1.CommandText = "SELECT TOP 1 *FROM m_refRunningItem WHERE (NoID=" & Val(lblNoID.Text) & ") ORDER BY Item DESC"
            Reader = SQl1.ExecuteReader
            If Reader.Read Then
                lblItemRunning.Text = Val(Reader("Item")) + 1
            Else
                lblItemRunning.Text = "1"
            End If
            Conn.Close()
            Reader.Close()

        Catch ex As Exception

        End Try
    End Sub
    Sub MasukItemRunning()
        Dim Conn As SqlConnection
        Dim SQl1 As New SqlCommand
        Dim Tanggal = Format(Now, "yyyy/MM/dd HH:mm:ss").ToString
        Try
            'Call NoItem()
            Call NoItemRunning()

            Conn = GetConnect()
            Conn.Open()
            SQl1 = Conn.CreateCommand
            SQl1.CommandText = "INSERT INTO m_refRunningItem" _
                               & "(NoID,Item,SONo,QRCode,Model,Qty,ScanQrSt1,ScanQRSt2,ScanQRSt3,ScanQRSt4,iDate,QtyRun)" _
                               & "VALUES (" & Val(lblNoID.Text) & "," & Val(lblItemRunning.Text) & ",'" & txtSO.Text & "','" & txtQrCode.Text & "','" & txtRef.Text & "'," & Val(txtQty.Text) & ",0,0,0,0,'" & Tanggal & "',0)"
            SQl1.ExecuteReader()
            Conn.Close()

        Catch ex As Exception

        End Try
    End Sub
    Sub BersihkanScan()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "UPDATE m_refPartList SET Scan=0 WHERE (Model='" & txtRef.Text & "')"
            SQL1.ExecuteReader()
            Conn.Close()

            f_utama.lblLabel.Text = "SCAN PART LIST..."
        Catch ex As Exception

        End Try
    End Sub
    Sub MasukQtyMW()

        Dim StartAddress As Integer = 300
        Dim Values(0) As Integer
        Values(0) = Val(txtQty.Text)
        modBusClient.WriteMultipleRegisters(StartAddress, Values)

    End Sub
    Sub MasukQtyMW1()

        Dim StartAddress As Integer = 300
        Dim Values(0) As Integer
        Values(0) = Val(txtQty.Text)
        modBusClient1.WriteMultipleRegisters(StartAddress, Values)

    End Sub
    Sub MasukQRCodeMW()
        Dim cb() As Byte
        txtConvert.Text = ""
        If txtMWSerial.Text <> "" Then
            cb = ASCIIEncoding.ASCII.GetBytes(txtMWSerial.Text)
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
                    modBusClient.WriteMultipleRegisters(StartAddress, MWval_reff)
                Catch ex As Exception

                End Try
            Next

        End If
    End Sub
    Sub MasukModelMW()

        Dim cb() As Byte
        txtConvert.Text = ""
        If txtRef.Text <> "" Then
            cb = ASCIIEncoding.ASCII.GetBytes(txtRef.Text)
            Dim MWval_reff(19) As Integer 'jumlah address hingga 20
            For x As Integer = 0 To cb.Length - 1
                txtConvert.AppendText(cb(x).ToString & "")
                'txtFrom.Text = ""
                'txtTo.AppendText(cb(x).ToString & "")
                'Dim arrayRef() as Char=cb(x).to
                'MessageBox.Show(cb(x).ToString)
                MWval_reff(x) = Convert.ToInt64(cb(x))
                Try
                    Dim StartAddress As Integer = 400 'mulai dari MW400
                    modBusClient.WriteMultipleRegisters(StartAddress, MWval_reff)
                Catch ex As Exception

                End Try
            Next

        End If

    End Sub
    Sub MasukModelMW1()

        Dim cb() As Byte
        txtConvert.Text = ""
        If txtRef.Text <> "" Then
            cb = ASCIIEncoding.ASCII.GetBytes(txtRef.Text)
            Dim MWval_reff(19) As Integer 'jumlah address hingga 20
            For x As Integer = 0 To cb.Length - 1
                txtConvert.AppendText(cb(x).ToString & "")
                'txtFrom.Text = ""
                'txtTo.AppendText(cb(x).ToString & "")
                'Dim arrayRef() as Char=cb(x).to
                'MessageBox.Show(cb(x).ToString)
                MWval_reff(x) = Convert.ToInt64(cb(x))
                Try
                    Dim StartAddress As Integer = 400 'mulai dari MW400
                    modBusClient1.WriteMultipleRegisters(StartAddress, MWval_reff)
                Catch ex As Exception

                End Try
            Next

        End If

    End Sub
    Sub NoID()
        Dim Conn As SqlConnection
        Dim SQl1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()

            SQl1 = Conn.CreateCommand
            SQl1.CommandText = "SELECT TOP 1 *FROM m_refRunning ORDER BY NoID DESC"
            Reader = SQl1.ExecuteReader
            If Reader.Read Then
                lblNoID.Text = Val(Reader("NoID")) + 1
            Else
                lblNoID.Text = "10001"
            End If

            Conn.Close()
            Reader.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString)
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
            SQL1.CommandText = "SELECT TOP 1 *FROM m_refRunning WHERE (NoID=" & Val(lblNoID.Text) & ") ORDER BY Item DESC"
            Reader = SQL1.ExecuteReader
            If Reader.Read Then
                lblItem.Text = Val(Reader("Item")) + 1
            Else
                lblItem.Text = "1"
            End If
            Conn.Close()
            Reader.Close()

        Catch ex As Exception

        End Try
    End Sub
    Sub MasukRunning()
        Dim Conn As SqlConnection
        Dim SQl1 As New SqlCommand
        Dim Tanggal = Format(Now, "yyyy/MM/dd HH:mm:ss").ToString
        Dim Jam = Format(Now, "HH:mm:ss").ToString

        Try
            Call NoID()
            Call NoItem()

            Conn = GetConnect()
            Conn.Open()

            SQl1 = Conn.CreateCommand
            SQl1.CommandText = "INSERT INTO m_refrunning" _
                             & "(NoID,Item,SONo,ItemNo,Model,Qty)" _
                             & "VALUES (" & Val(lblNoID.Text) & "," & Val(lblItem.Text) & ",'" & txtSO.Text & "','" & txtItem.Text & "','" & txtRef.Text & "'," & Val(txtQty.Text) & ")"
            SQl1.ExecuteReader()
            Conn.Close()

            Call Update_m_running()

            'Call MasukQtyMW()
            'Call MasukModelMW()
            Call KirimkeMW350()

            Call MasukItemRunning()
            Call NomorIDkeMW()
            Call NoIDtoMW()
            Call StatusConnectToMW()

            f_utama.lblSequence.Text = "1"
            f_utama.lblNoID.Text = lblNoID.Text
            f_utama.tmrSequence.Enabled = True


        Catch ex As Exception
            'MessageBox.Show(ex.Message.ToString)
        End Try

    End Sub
    Sub NoIDtoMW()
        Try
            ' MW 10001 untuk Nomor ID
            Dim StartAddress As Integer = 10001
            Dim Values(0) As Integer
            Values(0) = Val(lblNoID.Text)
            modBusClient.WriteMultipleRegisters(StartAddress, Values)
        Catch ex As Exception

        End Try
    End Sub
    Sub NomorIDkeMW()
        Try
            ' MW 10001 untuk status connect
            Dim StartAddress As Integer = 10003
            Dim Values(0) As Integer
            Values(0) = Val(lblNoID.Text)
            modBusClient.WriteMultipleRegisters(StartAddress, Values)
        Catch ex As Exception

        End Try

    End Sub
    Sub StatusConnectToMW()
        Try
            Dim StartsAddress As Integer = 10002
            Dim Values(0) As Integer
            Values(0) = 1
            modBusClient.WriteMultipleRegisters(StartsAddress, Values)
        Catch ex As Exception

        End Try

    End Sub
    Sub Update_m_running()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim SQl2 As New SqlCommand
        Dim Tanggal = Format(Now, "yyyy/MM/dd HH:mm:ss").ToString
        Dim Jam = Format(Now, "HH:mm:ss").ToString

        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "UPDATE m_refrunning SET Voltage=(m_refmanual.Voltage),Layer=(m_refmanual.Layer),Buzzer=(m_refmanual.Buzzer),Tube=(m_refmanual.Tube),Foot=(m_refmanual.Foot),PCBA=(m_refmanual.PCBA) FROM m_refrunning,m_refmanual WHERE (m_refrunning.model=m_refmanual.Model)AND(m_refrunning.NoID=" & Val(lblNoID.Text) & ")"
            SQL1.ExecuteReader()
            Conn.Close()

            Conn.Open()
            SQl2 = Conn.CreateCommand
            SQl2.CommandText = "UPDATE m_refrunning SET iDate='" & Tanggal & "',iTime='" & Jam & "',Status='R' WHERE (NoID=" & Val(lblNoID.Text) & ")"
            SQl2.ExecuteReader()
            Conn.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Update m_refrunning", MessageBoxButtons.OK)
        End Try
    End Sub
    Sub BuatQRCode()

        Dim Conn As SqlConnection
        Dim SQl1 As New SqlCommand
        Dim Reader As SqlDataReader
        Dim Awal As String = "8B"
        Dim Model As String = txtRef.Text
        Dim ItemNo As String = txtItem.Text
        Dim Tahun = Format(Now, "yy").ToString
        Dim Week = DatePart("ww", Now, vbSunday)


        Dim NomorBarcode As String

        Try
            Conn = GetConnect()
            Conn.Open()
            SQl1 = Conn.CreateCommand
            SQl1.CommandText = "SELECT RIGHT(QRCode,5) AS QRCode FROM m_refRunningItem WHERE (SONo='" & txtSO.Text & "') ORDER BY RIGHT(QRCode,5) DESC"
            Reader = SQl1.ExecuteReader
            If Reader.Read Then
                txtItem.Text = Format(Reader("QRCode") + 1, "00000")
                NomorBarcode = Awal & "," & Model & "," & Tahun & Week & "," & txtItem.Text
                txtQrCode.Text = NomorBarcode

            Else
                txtItem.Text = "00001"
                NomorBarcode = Awal & "," & Model & "," & Tahun & Week & "," & txtItem.Text
                txtQrCode.Text = NomorBarcode

            End If
            Conn.Close()
            Reader.Close()

            'NomorBarcode = Awal & "," & Model & "," & Tahun & Week & "," & ItemNo

            txtMWSerial.Text = NomorBarcode
            f_utama.lblQr.Text = NomorBarcode
        Catch ex As Exception

        End Try

    End Sub
    Private Sub BtnSubmit_Click(sender As Object, e As EventArgs) Handles btnSubmit.Click
        'If txtSO.Text = "" Or txtRef.Text = "" Or txtItem.Text = "" Or txtQty.Text = "" Then
        'MessageBox.Show("Masukan data reference", "Error", MessageBoxButtons.OK)
        'Else

        'Dim Utama As New f_utama
        'Utama.lblSO.Text = txtSO.Text
        f_utama.lblRef.Text = txtRef.Text
        f_utama.lblSO.Text = txtSO.Text
        f_utama.lblItem.Text = txtItem.Text
        f_utama.lblQty.Text = txtQty.Text
        f_utama.lblOperator.Text = lblUser.Text
        f_utama.lblDept.Text = lblDept.Text
        f_utama.lblTarget.Text = txtQty.Text

        Call BersihkanScan()
        Call f_utama.PanggilRef()
        'Call f_utama.BuatNoBarcode()
        Call BuatQRCode()

        Call MasukRunning()

        Me.Close()
        'Call f_utama.PanggilRef()

        f_utama.ShowDialog()

        'f_utama.Show()
        'Utama.ShowDialog()
        'Me.Hide()
        'End If


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
        f_utama.ShowDialog()

    End Sub

    Private Sub txtItem_KeyDown(sender As Object, e As KeyEventArgs) Handles txtItem.KeyDown
        If e.KeyCode = Keys.Enter Then
            txtRef.Focus()
        End If
    End Sub

    Private Sub txtRef_KeyDown(sender As Object, e As KeyEventArgs) Handles txtRef.KeyDown
        If e.KeyCode = Keys.Enter Then

            Call LihatRefNo()
            Call BersihkanMWRefBench1()
            Call BersihkanMWRefBench2()
            'Call BersihkanMWRefBench3()
            'Call BersihkanMWRefBench4()

            Call BersihkanStatusBench1()
            Call BersihkanStatusBench2()
            Call BersihkanStatusBench3()
            Call BersihkanStatusBench4()

            Call BersihkanPart1Bench1()
            Call BersihkanPart2Bench1()
            Call BersihkanPart3Bench1()
            Call BersihkanPart4Bench1()
            Call BersihkanPart5Bench1()
            Call BersihkanPart6Bench1()
            Call BersihkanPart7Bench1()

            Call BersihkanPart1Bench2()
            Call BersihkanPart2Bench2()
            Call BersihkanPart3Bench2()
            Call BersihkanPart4Bench2()

            Call BersihkanPart1Bench3()
            Call BersihkanPart2Bench3()
            Call BersihkanPart3Bench3()
            Call BersihkanPart4Bench3()
            Call BersihkanPart5Bench3()
            Call BersihkanPart6Bench3()

            Call BersihkanPart1Bench4()
            Call BersihkanPart2Bench4()
            Call BersihkanPart3Bench4()
            Call BersihkanPart4Bench4()
            Call BersihkanPart5Bench4()
            Call BersihkanPart6Bench4()

            Call BersihkanPokaPick()
            Call BersihkanQtyRunHMI()

            Call MasukModelMW()
            Call MasukModelMW1()

            'Call MasukQRCodeMW()


        End If
    End Sub
    Sub BersihkanSerialNo()
        Try
            Dim StartAddress As Integer = 500
            Dim Values(19) As Integer
            Values(19) = 0
            modBusClient.WriteMultipleRegisters(StartAddress, Values)

        Catch ex As Exception

        End Try

    End Sub
    Sub BersihkanStatusBench1()
        Try
            'Status Progress
            Dim StartAddress As Integer = 1700
            Dim Values(0) As Integer
            Values(0) = 0
            modBusClient.WriteMultipleRegisters(StartAddress, Values)

            'Status Fail
            Dim StartAddress1 As Integer = 1701
            Dim Values1(0) As Integer
            Values1(0) = 0
            modBusClient.WriteMultipleRegisters(StartAddress1, Values1)

            'Status Complete
            Dim StartAddress2 As Integer = 1702
            Dim Values2(0) As Integer
            Values2(0) = 0
            modBusClient.WriteMultipleRegisters(StartAddress2, Values2)
        Catch ex As Exception

        End Try

    End Sub
    Sub BersihkanStatusBench2()
        Try
            'Status Progress
            Dim StartAddress As Integer = 2700
            Dim Values(0) As Integer
            Values(0) = 0
            modBusClient1.WriteMultipleRegisters(StartAddress, Values)

            'Status Fail
            Dim StartAddress1 As Integer = 2702
            Dim Values1(0) As Integer
            Values1(0) = 0
            modBusClient1.WriteMultipleRegisters(StartAddress1, Values1)

            'Status Complete
            Dim StartAddress2 As Integer = 2703
            Dim Values2(0) As Integer
            Values2(0) = 0
            modBusClient1.WriteMultipleRegisters(StartAddress2, Values2)
        Catch ex As Exception

        End Try

    End Sub
    Sub BersihkanStatusBench3()
        Try
            'Status Progress
            Dim StartAddress As Integer = 3700
            Dim Values(0) As Integer
            Values(0) = 0
            modBusClient1.WriteMultipleRegisters(StartAddress, Values)

            'Status Fail
            Dim StartAddress1 As Integer = 3702
            Dim Values1(0) As Integer
            Values1(0) = 0
            modBusClient1.WriteMultipleRegisters(StartAddress1, Values1)

            'Status Complete
            Dim StartAddress2 As Integer = 3703
            Dim Values2(0) As Integer
            Values2(0) = 0
            modBusClient1.WriteMultipleRegisters(StartAddress2, Values2)
        Catch ex As Exception

        End Try

    End Sub
    Sub BersihkanStatusBench4()
        Try
            'Status Progress
            Dim StartAddress As Integer = 4700
            Dim Values(0) As Integer
            Values(0) = 0
            modBusClient1.WriteMultipleRegisters(StartAddress, Values)

            'Status Fail
            Dim StartAddress1 As Integer = 4702
            Dim Values1(0) As Integer
            Values1(0) = 0
            modBusClient1.WriteMultipleRegisters(StartAddress1, Values1)

            'Status Complete
            Dim StartAddress2 As Integer = 4703
            Dim Values2(0) As Integer
            Values2(0) = 0
            modBusClient1.WriteMultipleRegisters(StartAddress2, Values2)
        Catch ex As Exception

        End Try

    End Sub
    Private Sub txtQty_KeyDown(sender As Object, e As KeyEventArgs) Handles txtQty.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then
                If txtQty.Text = "" Then
                    MessageBox.Show("Entry Qty Product", "Error", MessageBoxButtons.OK)
                Else
                    f_utama.lblRef.Text = txtRef.Text
                    f_utama.lblSO.Text = txtSO.Text
                    f_utama.lblItem.Text = txtItem.Text
                    f_utama.lblQty.Text = txtQty.Text
                    f_utama.lblOperator.Text = lblUser.Text
                    f_utama.lblDept.Text = lblDept.Text
                    f_utama.lblTarget.Text = txtQty.Text
                    'f_utama.lblNoID.Text = lblNoID.Text

                    '==============================================================
                    ' Bersihkan MW Pokapick
                    Dim StartAddress As Integer = 360
                    Dim Values(25) As Integer
                    Values(25) = 0
                    modBusClient.WriteMultipleRegisters(StartAddress, Values)
                    '==============================================================

                    Call BersihkanScan()
                    Call f_utama.PanggilRef()
                    Call BuatQRCode()
                    Call MasukRunning()
                    Call MasukQtyMW()
                    Call MasukQtyMW1()

                    Me.Close()
                    'Call f_utama.PanggilRef()

                    f_utama.ShowDialog()

                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString)
        End Try
    End Sub
    Sub BersihkanPart1Bench1()
        Try
            Dim StartAddress As Integer = 600
            Dim Values(29) As Integer
            Values(29) = 0
            modBusClient.WriteMultipleRegisters(StartAddress, Values)

        Catch ex As Exception

        End Try

    End Sub
    Sub BersihkanPart1Bench2()
        Try
            Dim StartAddress As Integer = 1800
            Dim Values(29) As Integer
            Values(29) = 0
            modBusClient1.WriteMultipleRegisters(StartAddress, Values)

        Catch ex As Exception

        End Try
    End Sub
    Sub BersihkanPart2Bench2()
        Try
            Dim StartAddress As Integer = 1900
            Dim Values(29) As Integer
            Values(29) = 0
            modBusClient1.WriteMultipleRegisters(StartAddress, Values)

        Catch ex As Exception

        End Try
    End Sub
    Sub BersihkanPart3Bench2()
        Try
            Dim StartAddress As Integer = 2000
            Dim Values(29) As Integer
            Values(29) = 0
            modBusClient1.WriteMultipleRegisters(StartAddress, Values)

        Catch ex As Exception

        End Try
    End Sub
    Sub BersihkanPart4Bench2()
        Try
            Dim StartAddress As Integer = 21000
            Dim Values(29) As Integer
            Values(29) = 0
            modBusClient1.WriteMultipleRegisters(StartAddress, Values)


        Catch ex As Exception

        End Try
    End Sub
    Sub BersihkanPart5Bench2()
        Try
            Dim StartAddress As Integer = 22000
            Dim Values(29) As Integer
            Values(29) = 0
            modBusClient1.WriteMultipleRegisters(StartAddress, Values)

        Catch ex As Exception

        End Try
    End Sub
    Sub BersihkanPart2Bench1()
        Try
            Dim StartAddress As Integer = 700
            Dim Values(29) As Integer
            Values(29) = Val(txtMWSerial.Text)
            modBusClient.WriteMultipleRegisters(StartAddress, Values)

        Catch ex As Exception

        End Try
    End Sub
    Sub BersihkanPart3Bench1()
        Try
            Dim StartAddress As Integer = 800
            Dim Values(29) As Integer
            Values(29) = Val(txtMWSerial.Text)
            modBusClient.WriteMultipleRegisters(StartAddress, Values)

        Catch ex As Exception

        End Try
    End Sub
    Sub BersihkanPart4Bench1()
        Try
            Dim StartAddress As Integer = 900
            Dim Values(29) As Integer
            Values(29) = Val(txtMWSerial.Text)
            modBusClient.WriteMultipleRegisters(StartAddress, Values)

        Catch ex As Exception

        End Try
    End Sub
    Sub BersihkanPart5Bench1()
        Try
            Dim StartAddress As Integer = 1000
            Dim Values(29) As Integer
            Values(29) = Val(txtMWSerial.Text)
            modBusClient.WriteMultipleRegisters(StartAddress, Values)

        Catch ex As Exception

        End Try
    End Sub
    Sub BersihkanPart6Bench1()
        Try
            Dim StartAddress As Integer = 1100
            Dim Values(29) As Integer
            Values(29) = Val(txtMWSerial.Text)
            modBusClient.WriteMultipleRegisters(StartAddress, Values)

        Catch ex As Exception

        End Try
    End Sub
    Sub BersihkanPart7Bench1()
        Try
            Dim StartAddress As Integer = 1200
            Dim Values(29) As Integer
            Values(29) = Val(txtMWSerial.Text)
            modBusClient.WriteMultipleRegisters(StartAddress, Values)

        Catch ex As Exception

        End Try
    End Sub
    Sub BersihkanPart1Bench3()
        Try
            Dim StartAddress As Integer = 2800
            Dim Values(29) As Integer
            Values(29) = 0
            modBusClient1.WriteMultipleRegisters(StartAddress, Values)

        Catch ex As Exception

        End Try
    End Sub
    Sub BersihkanPart2Bench3()
        Try
            Dim StartAddress As Integer = 2900
            Dim Values(29) As Integer
            Values(29) = 0
            modBusClient1.WriteMultipleRegisters(StartAddress, Values)

        Catch ex As Exception

        End Try
    End Sub
    Sub BersihkanPart3Bench3()
        Try
            Dim StartAddress As Integer = 3000
            Dim Values(29) As Integer
            Values(29) = 0
            modBusClient1.WriteMultipleRegisters(StartAddress, Values)

        Catch ex As Exception

        End Try
    End Sub
    Sub BersihkanPart4Bench3()
        Try
            Dim StartAddress As Integer = 3100
            Dim Values(29) As Integer
            Values(29) = 0
            modBusClient1.WriteMultipleRegisters(StartAddress, Values)

        Catch ex As Exception

        End Try
    End Sub
    Sub BersihkanPart5Bench3()
        Try
            Dim StartAddress As Integer = 3200
            Dim Values(29) As Integer
            Values(29) = 0
            modBusClient1.WriteMultipleRegisters(StartAddress, Values)

        Catch ex As Exception

        End Try
    End Sub
    Sub BersihkanPart6Bench3()
        Try
            Dim StartAddress As Integer = 3300
            Dim Values(29) As Integer
            Values(29) = 0
            modBusClient1.WriteMultipleRegisters(StartAddress, Values)

        Catch ex As Exception

        End Try
    End Sub
    Sub BersihkanPart1Bench4()
        Try
            Dim StartAddress As Integer = 3800
            Dim Values(29) As Integer
            Values(29) = 0
            modBusClient1.WriteMultipleRegisters(StartAddress, Values)

        Catch ex As Exception

        End Try
    End Sub
    Sub BersihkanPart2Bench4()
        Try
            Dim StartAddress As Integer = 3900
            Dim Values(29) As Integer
            Values(29) = 0
            modBusClient1.WriteMultipleRegisters(StartAddress, Values)

        Catch ex As Exception

        End Try
    End Sub
    Sub BersihkanPart3Bench4()
        Try
            Dim StartAddress As Integer = 4000
            Dim Values(29) As Integer
            Values(29) = 0
            modBusClient1.WriteMultipleRegisters(StartAddress, Values)

        Catch ex As Exception

        End Try
    End Sub
    Sub BersihkanPart4Bench4()
        Try
            Dim StartAddress As Integer = 4100
            Dim Values(29) As Integer
            Values(29) = 0
            modBusClient1.WriteMultipleRegisters(StartAddress, Values)

        Catch ex As Exception

        End Try
    End Sub
    Sub BersihkanPart5Bench4()
        Try
            Dim StartAddress As Integer = 4200
            Dim Values(29) As Integer
            Values(29) = 0
            modBusClient1.WriteMultipleRegisters(StartAddress, Values)

        Catch ex As Exception

        End Try
    End Sub
    Sub BersihkanPart6Bench4()
        Try
            Dim StartAddress As Integer = 4300
            Dim Values(29) As Integer
            Values(29) = 0
            modBusClient1.WriteMultipleRegisters(StartAddress, Values)

        Catch ex As Exception

        End Try
    End Sub
    Sub BersihkanPokaPick()

        Try
            Dim StartAddress As Integer = 360
            Dim Values(25) As Integer
            Values(25) = 0
            modBusClient.WriteMultipleRegisters(StartAddress, Values)

        Catch ex As Exception

        End Try
    End Sub
    'End Sub
    Sub BersihkanMWRefBench1()
        Try

            Dim StartAddress As Integer = 400
            Dim Values(0) As Integer
            Values(0) = Val(txtMWSerial.Text)
            modBusClient.WriteMultipleRegisters(StartAddress, Values)
        Catch ex As Exception

        End Try

    End Sub
    Sub BersihkanQtyRunHMI()
        Try
            Dim StartAddress As Integer = 310
            Dim StartAddress1 As Integer = 311
            'Dim StartAddress2 As Integer = 312
            'Dim StartAddress3 As Integer = 313

            Dim Values(0) As Integer
            Values(0) = 0
            modBusClient.WriteMultipleRegisters(StartAddress, Values)

            Dim Values1(3) As Integer
            Values1(3) = 0
            modBusClient1.WriteMultipleRegisters(StartAddress1, Values1)

        Catch ex As Exception

        End Try
    End Sub
    Sub BersihkanMWRefBench2()

        Try
            Dim StartAddress As Integer = 400
            Dim Values(0) As Integer
            Values(0) = Val(txtMWSerial.Text)
            modBusClient1.WriteMultipleRegisters(StartAddress, Values)
        Catch ex As Exception

        End Try

    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'Call MasukModelMW()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        txtSO.Text = ""
        txtItem.Text = ""
        txtQty.Text = ""
        txtRef.Text = ""
        txtSO.Focus()

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
            SQL1.CommandText = "SELECT *FROM m_refPartList WHERE (Model='" & txtRef.Text & "')AND(Station=1) ORDER BY Item ASC"
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
                    lSingleItem = LvSt1.Items.Add(.Item("Item").ToString)
                    lSingleItem.SubItems.Add(.Item("Model").ToString)
                    'lSingleItem.SubItems.Add(.Item("PartNo").ToString)
                    'lSingleItem.SubItems.Add(.Item("Scan").ToString)
                    'lSingleItem.SubItems.Add(.Item("PokNo").ToString)

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
            'Call TotalOrder()

            'Call TotalItem()

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
            SQL1.CommandText = "SELECT *FROM m_refPartList WHERE (Model='" & txtRef.Text & "')AND(Station=2) ORDER BY Item ASC"
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
                    lSingleItem = LvSt1.Items.Add(.Item("Item").ToString)
                    lSingleItem.SubItems.Add(.Item("Model").ToString)
                    'lSingleItem.SubItems.Add(.Item("PartNo").ToString)
                    'lSingleItem.SubItems.Add(.Item("Scan").ToString)
                    'lSingleItem.SubItems.Add(.Item("PokNo").ToString)

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
            SQL1.CommandText = "SELECT *FROM m_refPartList WHERE (Model='" & txtRef.Text & "')AND(Station=3) ORDER BY Item ASC"
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
                    lSingleItem = LvSt3.Items.Add(.Item("Item").ToString)
                    lSingleItem.SubItems.Add(.Item("Model").ToString)
                    'lSingleItem.SubItems.Add(.Item("PartNo").ToString)
                    'lSingleItem.SubItems.Add(.Item("Scan").ToString)
                    'lSingleItem.SubItems.Add(.Item("PokNo").ToString)

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
            SQL1.CommandText = "SELECT *FROM m_refPartList WHERE (Model='" & txtRef.Text & "')AND(Station=4) ORDER BY Item ASC"
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
                    lSingleItem = LvSt4.Items.Add(.Item("Item").ToString)
                    lSingleItem.SubItems.Add(.Item("Model").ToString)
                    'lSingleItem.SubItems.Add(.Item("PartNo").ToString)
                    'lSingleItem.SubItems.Add(.Item("Scan").ToString)
                    'lSingleItem.SubItems.Add(.Item("PokNo").ToString)

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
            'Call TotalOrder()

            'Call TotalItem()

            conn.Close()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtQty_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtQty.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub
End Class