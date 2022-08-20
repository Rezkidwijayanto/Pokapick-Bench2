Imports System
Imports System.Data
Imports System.Data.SqlClient

Public Class f_reference
    Sub NoID()
        Dim Conn As SqlConnection
        Dim SQl1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()
            SQl1 = Conn.CreateCommand
            SQl1.CommandText = "SELECT TOP 1 *FROM m_refmanual ORDER BY NoID DESC"
            Reader = SQl1.ExecuteReader
            If Reader.Read Then
                lblID.Text = Reader("NoID") + 1
            Else
                lblID.Text = "2001001001"
            End If
            Reader.Close()
            Conn.Close()

        Catch ex As Exception

        End Try
    End Sub

    Private Sub f_reference_Load(sender As Object, e As EventArgs) Handles Me.Load

        'Call NoID()
        Call FormBlank()


    End Sub
    Sub FormBlank()
        lblID.Text = "#SSID"
        'txtSO.Text = ""
        'txtItem.Text = ""
        'txtRef.Text = ""
        'txtQty.Text = ""
        'txtType.Text = ""
        'rb24.Checked = True
        'txtVolt.Text = ""
        'txtLayer.Text = ""
        'txtBuzzer.Text = ""
        'txtFoot.Text = ""
        'txtTube.Text = ""
        'txtShip.Text = ""
        txtLay1.Text = ""
        txtLay2.Text = ""
        txtLay3.Text = ""
        txtLay4.Text = ""
        txtBuzTest.Text = ""
        txtVoltTest.Text = ""
        txtLay1Max.Text = ""
        txtLay2Max.Text = ""
        txtLay3Max.Text = ""
        txtLay4Max.Text = ""
        txtBuzMax.Text = ""
        txtSoundMax.Text = ""

        Call selectRef()

        Call TotalItem()
        Call dpItem()


    End Sub
    Sub Baru()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Tanggal = Format(Now, "yyyy/MM/dd HH:mm:ss").ToString

        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "INSERT INTO m_refmanual" _
                             & "(NoID,iDate,Status)" _
                             & "VALUES (" & Val(lblID.Text) & ",'" & Tanggal & "','O')"
            SQL1.ExecuteReader()
            Conn.Close()

        Catch ex As Exception

        End Try
    End Sub

    Private Sub BtnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        'Call NoID()
        'Call Baru()
        cmbRef.Focus()
        'txtSO.Focus()

    End Sub

    Private Sub BtnUndo_Click(sender As Object, e As EventArgs) Handles btnUndo.Click
        Call FormBlank()

    End Sub
    Sub SetDNA()
        Dim Conn As SqlConnection
        Dim SQl1 As New SqlCommand
        Try
            If ckDNA.Checked = True Then
                Conn = GetConnect()
                Conn.Open()
                SQl1 = Conn.CreateCommand
                SQl1.CommandText = "UPDATE m_refmanual SET DNA='Y' WHERE (NoID=" & Val(lblID.Text) & ")"
                SQl1.ExecuteReader()
                Conn.Close()
            Else
                Conn = GetConnect()
                Conn.Open()
                SQl1 = Conn.CreateCommand
                SQl1.CommandText = "UPDATE m_refmanual SET DNA='N' WHERE (NoID=" & Val(lblID.Text) & ")"
                SQl1.ExecuteReader()
                Conn.Close()

            End If
        Catch ex As Exception

        End Try
    End Sub
    Sub SetVoltage()
        Dim Conn As SqlConnection
        Dim SQl1 As New SqlCommand
        Try
            If rb24.Checked = True Then
                Conn = GetConnect()
                Conn.Open()
                SQl1 = Conn.CreateCommand
                SQl1.CommandText = "UPDATE m_refmanual SET Voltage='" & rb24.Text & "',ACDC='DC' WHERE (NoID=" & Val(lblID.Text) & ")"
                SQl1.ExecuteReader()
                Conn.Close()
            ElseIf rb120.Checked = True Then
                Conn = GetConnect()
                Conn.Open()
                SQl1 = Conn.CreateCommand
                SQl1.CommandText = "UPDATE m_refmanual SET Voltage='" & rb120.Text & "',ACDC='AC' WHERE (NoID=" & Val(lblID.Text) & ")"
                SQl1.ExecuteReader()
                Conn.Close()
            ElseIf rb220.Checked = True Then
                Conn = GetConnect()
                Conn.Open()
                SQl1 = Conn.CreateCommand
                SQl1.CommandText = "UPDATE m_refmanual SET Voltage='" & rb220.Text & "',ACDC='AC' WHERE (NoID=" & Val(lblID.Text) & ")"
                SQl1.ExecuteReader()
                Conn.Close()


            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Voltage Error", MessageBoxButtons.OK)
        End Try
    End Sub
    Sub SetLayer()
        Dim Conn As SqlConnection
        Dim SQl1 As New SqlCommand
        Try
            If rb1Layer.Checked = True Then
                Conn = GetConnect()
                Conn.Open()
                SQl1 = Conn.CreateCommand
                SQl1.CommandText = "UPDATE m_refmanual SET Layer='" & rb1Layer.Text & "' WHERE (NoID=" & Val(lblID.Text) & ")"
                SQl1.ExecuteReader()
                Conn.Close()
            ElseIf rb2Layer.Checked = True Then
                Conn = GetConnect()
                Conn.Open()
                SQl1 = Conn.CreateCommand
                SQl1.CommandText = "UPDATE m_refmanual SET Layer='" & rb2Layer.Text & "' WHERE (NoID=" & Val(lblID.Text) & ")"
                SQl1.ExecuteReader()
                Conn.Close()
            ElseIf rb3Layer.Checked = True Then
                Conn = GetConnect()
                Conn.Open()
                SQl1 = Conn.CreateCommand
                SQl1.CommandText = "UPDATE m_refmanual SET Layer='" & rb3Layer.Text & "' WHERE (NoID=" & Val(lblID.Text) & ")"
                SQl1.ExecuteReader()
                Conn.Close()
            ElseIf rb4Layer.Checked = True Then
                Conn = GetConnect()
                Conn.Open()
                SQl1 = Conn.CreateCommand
                SQl1.CommandText = "UPDATE m_refmanual SET Layer='" & rb4Layer.Text & "' WHERE (NoID=" & Val(lblID.Text) & ")"
                SQl1.ExecuteReader()
                Conn.Close()


            End If
        Catch ex As Exception

        End Try
    End Sub
    Sub SetBuzzer()
        Dim Conn As SqlConnection
        Dim SQl1 As New SqlCommand
        Try
            If ckBuzzer.Checked = True Then
                Conn = GetConnect()
                Conn.Open()
                SQl1 = Conn.CreateCommand
                SQl1.CommandText = "UPDATE m_refmanual SET Buzzer='ON' WHERE (NoID=" & Val(lblID.Text) & ")"
                SQl1.ExecuteReader()
                Conn.Close()
            Else
                Conn = GetConnect()
                Conn.Open()
                SQl1 = Conn.CreateCommand
                SQl1.CommandText = "UPDATE m_refmanual SET Buzzer='OFF' WHERE (NoID=" & Val(lblID.Text) & ")"
                SQl1.ExecuteReader()
                Conn.Close()


            End If
        Catch ex As Exception

        End Try
    End Sub
    Sub SetTube()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Try
            If rbTubeShort.Checked = True Then
                Conn = GetConnect()
                Conn.Open()
                SQL1 = Conn.CreateCommand
                SQL1.CommandText = "UPDATE m_refmanual SET Tube='" & rbTubeShort.Text & "' WHERE (NoID=" & Val(lblID.Text) & ")"
                SQL1.ExecuteReader()
                Conn.Close()
            ElseIf rbTubeMed.Checked = True Then
                Conn = GetConnect()
                Conn.Open()
                SQL1 = Conn.CreateCommand
                SQL1.CommandText = "UPDATE m_refmanual SET Tube='" & rbTubeMed.Text & "' WHERE (NoID=" & Val(lblID.Text) & ")"
                SQL1.ExecuteReader()
                Conn.Close()
            ElseIf rbTubeLong.Checked = True Then
                Conn = GetConnect()
                Conn.Open()
                SQL1 = Conn.CreateCommand
                SQL1.CommandText = "UPDATE m_refmanual SET Tube='" & rbTubeLong.Text & "' WHERE (NoID=" & Val(lblID.Text) & ")"
                SQL1.ExecuteReader()
                Conn.Close()

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Set Tube", MessageBoxButtons.OK)
        End Try
    End Sub
    Sub SetFoot()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Try
            If rbFootDirect.Checked = True Then
                Conn = GetConnect()
                Conn.Open()
                SQL1 = Conn.CreateCommand
                SQL1.CommandText = "UPDATE m_refmanual SET Foot='" & rbFootDirect.Text & "' WHERE (NoID=" & Val(lblID.Text) & ")"
                SQL1.ExecuteReader()
                Conn.Close()
            ElseIf rbBase.Checked = True Then
                Conn = GetConnect()
                Conn.Open()
                SQL1 = Conn.CreateCommand
                SQL1.CommandText = "UPDATE m_refmanual SET Foot='" & rbBase.Text & "' WHERE (NoID=" & Val(lblID.Text) & ")"
                SQL1.ExecuteReader()
                Conn.Close()
            ElseIf rbTubeNut.Checked = True Then
                Conn = GetConnect()
                Conn.Open()
                SQL1 = Conn.CreateCommand
                SQL1.CommandText = "UPDATE m_refmanual SET Foot='" & rbTubeNut.Text & "' WHERE (NoID=" & Val(lblID.Text) & ")"
                SQL1.ExecuteReader()
                Conn.Close()
            ElseIf rbTubeBrace.Checked = True Then
                Conn = GetConnect()
                Conn.Open()
                SQL1 = Conn.CreateCommand
                SQL1.CommandText = "UPDATE m_refmanual SET Foot='" & rbTubeBrace.Text & "' WHERE (NoID=" & Val(lblID.Text) & ")"
                SQL1.ExecuteReader()
                Conn.Close()
            ElseIf rbFoldBracket.Checked = True Then
                Conn = GetConnect()
                Conn.Open()
                SQL1 = Conn.CreateCommand
                SQL1.CommandText = "UPDATE m_refmanual SET Foot='" & rbFoldBracket.Text & "' WHERE (NoID=" & Val(lblID.Text) & ")"
                SQL1.ExecuteReader()
                Conn.Close()
            ElseIf rbFoldFolder.Checked = True Then
                Conn = GetConnect()
                Conn.Open()
                SQL1 = Conn.CreateCommand
                SQL1.CommandText = "UPDATE m_refmanual SET Foot='" & rbFoldFolder.Text & "' WHERE (NoID=" & Val(lblID.Text) & ")"
                SQL1.ExecuteReader()
                Conn.Close()
            ElseIf rbFoldUSB.Checked = True Then
                Conn = GetConnect()
                Conn.Open()
                SQL1 = Conn.CreateCommand
                SQL1.CommandText = "UPDATE m_refmanual SET Foot='" & rbFoldUSB.Text & "' WHERE (NoID=" & Val(lblID.Text) & ")"
                SQL1.ExecuteReader()
                Conn.Close()


            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Foot Error", MessageBoxButtons.OK)
        End Try
    End Sub
    Sub SetIRTest()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Try
            If rbIR1.Checked = True Then
                Conn = GetConnect()
                Conn.Open()
                SQL1 = Conn.CreateCommand
                SQL1.CommandText = "UPDATE m_refmanual SET IR='" & rbIR1.Text & "' WHERE (NoID=" & Val(lblID.Text) & ")"
                SQL1.ExecuteReader()
                Conn.Close()
            ElseIf rbIR2.Checked = True Then
                Conn = GetConnect()
                Conn.Open()
                SQL1 = Conn.CreateCommand
                SQL1.CommandText = "UPDATE m_refmanual SET IR='" & rbIR2.Text & "' WHERE (NoID=" & Val(lblID.Text) & ")"
                SQL1.ExecuteReader()
                Conn.Close()
            ElseIf rbNotest.Checked = True Then
                Conn = GetConnect()
                Conn.Open()
                SQL1 = Conn.CreateCommand
                SQL1.CommandText = "UPDATE m_refmanual SET IR='" & rbNotest.Text & "' WHERE (NoID=" & Val(lblID.Text) & ")"
                SQL1.ExecuteReader()
                Conn.Close()
            ElseIf rbBoth.Checked = True Then
                Conn = GetConnect()
                Conn.Open()
                SQL1 = Conn.CreateCommand
                SQL1.CommandText = "UPDATE m_refmanual SET IR='" & rbBoth.Text & "' WHERE (NoID=" & Val(lblID.Text) & ")"
                SQL1.ExecuteReader()
                Conn.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "IR Test", MessageBoxButtons.OK)
        End Try
    End Sub
    Sub Updateref()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Try
            Dim Lay1Min = Format(CType(txtLay1.Text, Decimal))
            Dim Lay1Max = Format(CType(txtLay1Max.Text, Decimal))
            Dim Lay2Min = Format(CType(txtLay2.Text, Decimal))
            Dim Lay2Max = Format(CType(txtLay2Max.Text, Decimal))
            Dim Lay3Min = Format(CType(txtLay3.Text, Decimal))
            Dim Lay3Max = Format(CType(txtLay3Max.Text, Decimal))
            Dim Lay4Min = Format(CType(txtLay4.Text, Decimal))
            Dim Lay4Max = Format(CType(txtLay4Max.Text, Decimal))
            Dim BuzzMin = Format(CType(txtBuzTest.Text, Decimal))
            Dim BuzzMax = Format(CType(txtBuzMax.Text, Decimal))
            Dim SoundMin = Format(CType(txtVoltTest.Text, Decimal))
            Dim SoundMax = Format(CType(txtSoundMax.Text, Decimal))

            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "UPDATE m_refmanual SET Layer1Min=" & Lay1Min & ",Layer1Max=" & Lay1Max & ",Layer2Min=" & Lay2Min & ",Layer2Max=" & Lay2Max & ",Layer3Min=" & Lay3Min & ",Layer3Max=" & Lay3Max & ",Layer4Min=" & Lay4Min & ",Layer4Max=" & Lay4Max & ",BuzTestMin=" & BuzzMin & ",BuzTestMax=" & BuzzMax & ",SoundMin=" & SoundMin & ",SoundMax=" & SoundMax & ",Status='C',PCBA='" & txtPCBA.Text & "'" _
                             & "WHERE (NoID=" & Val(lblID.Text) & ")"
            SQL1.ExecuteReader()
            Conn.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error Save", MessageBoxButtons.OK)
        End Try
    End Sub
    Sub Simpan()
        Call SetDNA()
        Call SetVoltage()
        Call SetLayer()
        Call SetBuzzer()
        Call SetTube()
        Call SetFoot()
        Call SetIRTest()
        Call Updateref()

        Call dpItem()

    End Sub


    Sub dpItem()
        Try
            Dim conn As SqlConnection
            Dim SQL1 As New SqlCommand
            Dim Pilih As New SqlDataAdapter
            Dim datWarga As New DataSet
            Dim datTable As New DataTable

            conn = GetConnect()
            conn.Open()

            SQL1 = conn.CreateCommand
            SQL1.CommandText = "SELECT *FROM m_refmanual WHERE (Status='C') ORDER BY NoID ASC"
            Pilih.SelectCommand = SQL1
            Pilih.Fill(datWarga, "m_refmanual")
            datTable = datWarga.Tables("m_refmanual")

            Dim i As Integer
            Dim x As Integer

            'Displaydata()
            Lv1.Items.Clear()
            For i = 0 To (datTable.Rows.Count - 1)
                With datTable.Rows(i)
                    Dim lSingleItem As ListViewItem
                    'Dim Tot = FormatCurrency("Amount")
                    lSingleItem = Lv1.Items.Add(.Item("NoID").ToString)
                    lSingleItem.SubItems.Add(.Item("Model").ToString)
                    lSingleItem.SubItems.Add(.Item("ACDC").ToString)
                    lSingleItem.SubItems.Add(.Item("Layer").ToString)
                    'lSingleItem.SubItems(3).Text = CDbl(lSingleItem.SubItems(3).Text).ToString("N2")
                    lSingleItem.SubItems.Add(.Item("Buzzer").ToString)
                    'lSingleItem.SubItems(4).Text = CDbl(lSingleItem.SubItems(4).Text).ToString("N2")
                    lSingleItem.SubItems.Add(.Item("Tube").ToString)
                    lSingleItem.SubItems.Add(.Item("Foot").ToString)
                    lSingleItem.SubItems.Add(.Item("Voltage").ToString)
                End With

                For x = 0 To Lv1.Items.Count - 1
                    If (x Mod 2) = 0 Then
                        Lv1.Items(x).BackColor = Color.MintCream

                    End If
                Next x

            Next i
            'Call TotalOrder()

            Call TotalItem()

            conn.Close()
        Catch ex As Exception

        End Try
    End Sub
    Sub selectRef()
        'On Error Resume Next
        Try
            Dim conn As New SqlConnection

            conn = GetConnect()
            conn.Open()

            Dim strSQL As String = "SELECT * FROM m_refmaster ORDER BY NoID ASC"
            Dim da As New SqlDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "m_refmaster")

            With cmbRef
                .DataSource = ds.Tables("m_refmaster")
                .DisplayMember = "refNO"
                .ValueMember = "NoID"
                .SelectedIndex = 0
            End With

            conn.Close()
        Catch ex As Exception

        End Try


    End Sub
    Sub DpCari()
        Try
            Dim conn As SqlConnection
            Dim SQL1 As New SqlCommand
            Dim Pilih As New SqlDataAdapter
            Dim datWarga As New DataSet
            Dim datTable As New DataTable

            conn = GetConnect()
            conn.Open()

            SQL1 = conn.CreateCommand
            SQL1.CommandText = "SELECT *FROM m_refmanual WHERE (SONo Like '%" & Trim(txtCari.Text) & "%' OR Model Like '%" & Trim(txtCari.Text) & "%') AND (Status='C') ORDER BY NoID ASC"
            Pilih.SelectCommand = SQL1
            Pilih.Fill(datWarga, "m_refmanual")
            datTable = datWarga.Tables("m_refmanual")

            Dim i As Integer
            Dim x As Integer

            'Displaydata()
            Lv1.Items.Clear()
            For i = 0 To (datTable.Rows.Count - 1)
                With datTable.Rows(i)
                    Dim lSingleItem As ListViewItem
                    'Dim Tot = FormatCurrency("Amount")
                    lSingleItem = Lv1.Items.Add(.Item("NoID").ToString)
                    lSingleItem.SubItems.Add(.Item("SONo").ToString)
                    lSingleItem.SubItems.Add(.Item("Model").ToString)
                    lSingleItem.SubItems.Add(.Item("ItemNo").ToString)
                    'lSingleItem.SubItems(3).Text = CDbl(lSingleItem.SubItems(3).Text).ToString("N2")
                    lSingleItem.SubItems.Add(.Item("Qty").ToString)
                    lSingleItem.SubItems(4).Text = CDbl(lSingleItem.SubItems(4).Text).ToString("N2")
                    lSingleItem.SubItems.Add(.Item("Type").ToString)
                End With

                For x = 0 To Lv1.Items.Count - 1
                    If (x Mod 2) = 0 Then
                        Lv1.Items(x).BackColor = Color.MintCream

                    End If
                Next x

            Next i
            'Call TotalOrder()

            Call TotalItem()

            conn.Close()
        Catch ex As Exception

        End Try
    End Sub
    Sub TotalItem()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader
        Dim tTotal As Double

        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "SELECT COUNT(SONo) AS tSO FROM m_refmanual WHERE (Status='C')"
            Reader = SQL1.ExecuteReader
            If Reader.Read Then
                tTotal = Format(Reader("tSO"), "#,##0").ToString
                lblTotal.Text = tTotal
            Else
                lblTotal.Text = "0"
            End If
            Conn.Close()
            Reader.Close()

        Catch ex As Exception

        End Try
    End Sub

    Private Sub BtnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close()

    End Sub

    Private Sub txtCari_KeyDown(sender As Object, e As KeyEventArgs) Handles txtCari.KeyDown
        If e.KeyCode = Keys.Enter Then
            Call DpCari()
        ElseIf e.KeyCode = Keys.Escape Then
            'txtSO.Focus()
            cmbRef.Focus()
        End If

    End Sub

    Private Sub BtnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click

        Call Simpan()

        'End If
    End Sub

    Private Sub Lv1_Click(sender As Object, e As EventArgs) Handles Lv1.Click
        Try
            Dim I As Integer
            For I = 0 To Lv1.SelectedItems.Count - 1
                lblID.Text = Lv1.SelectedItems(I).SubItems(0).Text

                Call Panggil()

            Next
        Catch ex As Exception

        End Try
    End Sub
    Sub Panggil()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "SELECT *FROM m_refmanual WHERE (NoID=" & Val(lblID.Text) & ")AND(Status='C')"
            Reader = SQL1.ExecuteReader
            If Reader.Read Then
                'txtSO.Text = Reader("SONo").ToString
                'txtRef.Text = Reader("Model").ToString
                cmbRef.Text = Reader("Model").ToString
                'txtItem.Text = Reader("ItemNo").ToString
                'txtQty.Text = Format(Reader("Qty"), "#,##0").ToString
                'txtType.Text = Reader("Type").ToString
                'txtVolt.Text = Reader("Voltage").ToString
                'txtLayer.Text = Reader("Layer").ToString
                'txtBuzzer.Text = Reader("Buzzer").ToString
                'txtFoot.Text = Reader("Foot").ToString
                'txtTube.Text = Reader("Tube").ToString
                'txtShip.Text = Reader("Shipto").ToString
                txtLay1.Text = Format(Reader("Layer1Min"), "#,##0").ToString
                txtLay1Max.Text = Format(Reader("Layer1Max"), "#,##0").ToString
                txtLay2.Text = Format(Reader("Layer2Min"), "#,##0").ToString
                txtLay2Max.Text = Format(Reader("Layer2Max"), "#,##0").ToString
                txtLay3.Text = Format(Reader("Layer3Min"), "#,##0").ToString
                txtLay3Max.Text = Format(Reader("Layer3Max"), "#,##0").ToString
                txtLay4.Text = Format(Reader("Layer4Min"), "#,##0").ToString
                txtLay4Max.Text = Format(Reader("Layer4Max"), "#,##0").ToString
                txtBuzTest.Text = Format(Reader("BuzTestMin"), "#,##0").ToString
                txtBuzMax.Text = Format(Reader("BuzTestMax"), "#,##0").ToString
                txtVoltTest.Text = Format(Reader("SoundMin"), "#,##0").ToString
                txtSoundMax.Text = Format(Reader("SoundMax"), "#,##0").ToString

                lblVolt.Text = Reader("Voltage").ToString
                lblLayer.Text = Reader("Layer").ToString
                lblBuzzer.Text = Reader("Buzzer").ToString
                lblFoot.Text = Reader("Foot").ToString
                lblTube.Text = Reader("Tube").ToString


                If Reader("Voltage") = "24VDC" Then
                    rb24.Checked = True
                ElseIf Reader("Voltage") = "120VAC" Then
                    rb120.Checked = True
                ElseIf Reader("Voltage") = "220VAC" Then
                    rb220.Checked = True
                End If

                If Reader("Layer") = "1 Layer" Then
                    rb1Layer.Checked = True
                ElseIf ("Layer") = "2 Layer" Then
                    rb2Layer.Checked = True
                ElseIf Reader("Layer") = "3 Layer" Then
                    rb3Layer.Checked = True
                ElseIf Reader("Layer") = "4 Layer" Then
                    rb4Layer.Checked = True
                End If

                'Else
                'MessageBox.Show("Data Not found!", "Error", MessageBoxButtons.OK)
            End If
            Conn.Close()
            Reader.Close()

        Catch ex As Exception

        End Try
    End Sub
    Sub Hapus()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "DELETE FROM m_refmanual WHERE (NoID=" & Val(lblID.Text) & ")"
            SQL1.ExecuteReader()

            Call dpItem()
            Call FormBlank()

        Catch ex As Exception

        End Try
    End Sub

    Private Sub BtnDel_Click(sender As Object, e As EventArgs) Handles btnDel.Click
        If lblID.Text = "#SSID" Then
            MessageBox.Show("No data to delete!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            'txtSO.Focus()
            cmbRef.Focus()
        Else
            If MessageBox.Show("Are you sure want to remove Reference?", "Remove", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                Call Hapus()

            End If
        End If
    End Sub

    'Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles rbTubeNut.CheckedChanged

    'End Sub

    Private Sub CmbRef_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbRef.SelectedIndexChanged
        Dim Conn As SqlConnection
        Dim SQl1 As New SqlCommand
        Dim Reader As SqlDataReader

        Try
            Conn = GetConnect()
            Conn.Open()
            'Dim Reader As SqlDataReader
            SQl1 = Conn.CreateCommand
            SQl1.CommandText = "SELECT NoID,Model,PCBA FROM m_refmanual WHERE (Model='" & cmbRef.Text & "')"
            Reader = SQl1.ExecuteReader
            If Reader.Read Then
                lblID.Text = Val(Reader("NoID"))
                txtPCBA.Text = Reader("PCBA").ToString

                Call Panggil()
                Call PanggilDNA()
                Call PanggilBuzzer()
                Call PanggilTube()
                Call PanggilFoot()
                Call PanggilIR()

            Else
                lblID.Text = "Unknow"
                txtPCBA.Text = ""

            End If
            Conn.Close()
            Reader.Close()

        Catch ex As Exception

        End Try
    End Sub
    Sub PanggilDNA()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "SELECT *FROM m_refmanual WHERE (NoID=" & Val(lblID.Text) & ")"
            Reader = SQL1.ExecuteReader
            If Reader.Read Then
                If Reader("DNA") = "Y" Then
                    ckDNA.Checked = True
                Else
                    ckDNA.Checked = False
                End If
            End If
            Conn.Close()
            Reader.Close()

        Catch ex As Exception

        End Try
    End Sub
    Sub PanggilIR()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "SELECT *FROM m_refmanual WHERE (NoID=" & Val(lblID.Text) & ")"
            Reader = SQL1.ExecuteReader
            If Reader.Read Then
                If Reader("IR") = "IR1 TEST ONLY" Then
                    'ckDNA.Checked = True
                    rbIR1.Checked = True
                ElseIf Reader("IR") = "IR2 TEST ONLY" Then
                    'ckDNA.Checked = False
                    rbIR2.Checked = True
                ElseIf Reader("IR") = "BOTH TEST" Then
                    rbBoth.Checked = True
                ElseIf Reader("IR") = "NO TEST" Then
                    rbNotest.Checked = True
                End If
            End If
            Conn.Close()
            Reader.Close()

        Catch ex As Exception

        End Try
    End Sub
    Sub PanggilBuzzer()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "SELECT *FROM m_refmanual WHERE (NoID=" & Val(lblID.Text) & ")"
            Reader = SQL1.ExecuteReader
            If Reader.Read Then
                If Reader("Buzzer") = "ON" Then
                    ckBuzzer.Checked = True
                Else
                    ckBuzzer.Checked = False
                End If
            End If
            Conn.Close()
            Reader.Close()

        Catch ex As Exception

        End Try
    End Sub
    Sub PanggilTube()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "SELECT *FROM m_refmanual WHERE (NoID=" & Val(lblID.Text) & ")"
            Reader = SQL1.ExecuteReader
            If Reader.Read Then
                If Reader("Tube") = "Long" Then
                    'ckBuzzer.Checked = True
                    rbTubeLong.Checked = True
                ElseIf Reader("Tube") = "Medium" Then
                    'ckBuzzer.Checked = False
                    rbTubeMed.Checked = True
                ElseIf Reader("Tube") = "Short" Then
                    rbTubeShort.Checked = True
                End If

            End If
            Conn.Close()
            Reader.Close()

        Catch ex As Exception

        End Try
    End Sub
    Sub PanggilFoot()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "SELECT *FROM m_refmanual WHERE (NoID=" & Val(lblID.Text) & ")"
            Reader = SQL1.ExecuteReader
            If Reader.Read Then
                If Reader("Foot") = "Direct" Then
                    'ckBuzzer.Checked = True
                    'rbTubeLong.Checked = True
                    rbFootDirect.Checked = True
                ElseIf Reader("Foot") = "Base" Then
                    'ckBuzzer.Checked = False
                    'rbTubeMed.Checked = True
                    rbBase.Checked = True
                ElseIf Reader("Foot") = "Tube + Nut" Then
                    'rbTubeShort.Checked = True
                    rbTubeNut.Checked = True
                ElseIf Reader("Foot") = "Tube + Brace" Then
                    rbTubeBrace.Checked = True
                ElseIf Reader("Foot") = "Foldable Bracked" Then
                    rbFoldBracket.Checked = True
                ElseIf Reader("Foot") = "Foldable Folder" Then
                    rbFoldFolder.Checked = True
                ElseIf Reader("Foot") = "Foldable USB" Then
                    rbFoldUSB.Checked = True
                End If

            End If
            Conn.Close()
            Reader.Close()

        Catch ex As Exception

        End Try
    End Sub
End Class