Imports System
Imports System.Data
Imports System.Data.SqlClient

Public Class f_master
    Dim x As String

    Sub FormBlank()

        LvSt1.Items.Clear()
        LvSt2.Items.Clear()
        LvSt3.Items.Clear()
        LvSt4.Items.Clear()
        txtPartNo.Text = ""
        txtStation.Text = ""
        txtPokNo.Text = ""

        LvSt1.Items.Clear()
        LvSt2.Items.Clear()
        LvSt3.Items.Clear()
        LvSt4.Items.Clear()

        Call dpList()

    End Sub
    Sub UpdateNoMW()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "UPDATE m_refPartList SET MW=(m_refMW.WM) FROM m_refPartList,m_refMW WHERE(m_refPartList.PokNo=m_refMW.PP)AND(m_refPartList.PartNo='" & txtPartNo.Text & "')"
            SQL1.ExecuteReader()
            Conn.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString)
        End Try
    End Sub
    Private Sub f_master_Load(sender As Object, e As EventArgs) Handles Me.Load
        'lblDate.Text = Format(Now, "dd/MM/yyyy").ToString

        'Call FormBlank()
        'Call dpItem()
        Call FormBlank()
        Call dpList()
        Call selectRef()

    End Sub
    Sub selectRef()
        'On Error Resume Next
        Try
            Dim conn As New SqlConnection

            conn = GetConnect()
            conn.Open()

            Dim strSQL As String = "SELECT * FROM m_refmanual WHERE (Status='C') ORDER BY Model ASC"
            Dim da As New SqlDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "m_refmanual")

            With cmbRef
                .DataSource = ds.Tables("m_refmanual")
                .DisplayMember = "Model"
                .ValueMember = "NoID"
                .SelectedIndex = 0
            End With

            conn.Close()
        Catch ex As Exception

        End Try


    End Sub
    Sub dpList()
        Try
            Dim conn As SqlConnection
            Dim SQL1 As New SqlCommand
            Dim Pilih As New SqlDataAdapter
            Dim datWarga As New DataSet
            Dim datTable As New DataTable

            conn = GetConnect()
            conn.Open()

            SQL1 = conn.CreateCommand
            SQL1.CommandText = "SELECT *FROM m_refmanual WHERE (PartSt='Y') ORDER BY Model ASC"
            Pilih.SelectCommand = SQL1
            Pilih.Fill(datWarga, "m_refmanual")
            datTable = datWarga.Tables("m_refmanual")

            Dim i As Integer
            Dim x As Integer

            'Displaydata()
            LvList.Items.Clear()
            For i = 0 To (datTable.Rows.Count - 1)
                With datTable.Rows(i)
                    Dim lSingleItem As ListViewItem
                    'Dim Tot = FormatCurrency("Amount")
                    lSingleItem = LvList.Items.Add(.Item("Model").ToString)
                    lSingleItem.SubItems.Add(.Item("Voltage").ToString)
                    lSingleItem.SubItems.Add(.Item("Layer").ToString)
                    lSingleItem.SubItems.Add(.Item("Buzzer").ToString)
                    'lSingleItem.SubItems(3).Text = CDbl(lSingleItem.SubItems(3).Text).ToString("N2")
                    lSingleItem.SubItems.Add(.Item("Tube").ToString)
                    'lSingleItem.SubItems(4).Text = CDbl(lSingleItem.SubItems(4).Text).ToString("N2")
                    'lSingleItem.SubItems.Add(.Item("Type").ToString)
                End With

                For x = 0 To LvList.Items.Count - 1
                    If (x Mod 2) = 0 Then
                        LvList.Items(x).BackColor = Color.MintCream

                    End If
                Next x

            Next i
            conn.Close()
        Catch ex As Exception

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
            SQL1.CommandText = "SELECT *FROM m_refPartList WHERE (Model='" & cmbRef.Text & "')AND(Station=1) ORDER BY Item ASC"
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
                    lSingleItem.SubItems.Add(.Item("PartNo").ToString)
                    'lSingleItem = LvSt1.Items.Add(.Item("Item").ToString)
                    'lSingleItem.SubItems.Add(.Item("RefNo").ToString)
                    'lSingleItem.SubItems.Add(.Item("ItemNo").ToString)
                    'lSingleItem.SubItems(3).Text = CDbl(lSingleItem.SubItems(3).Text).ToString("N2")
                    'lSingleItem.SubItems.Add(.Item("Qty").ToString)
                    'lSingleItem.SubItems(4).Text = CDbl(lSingleItem.SubItems(4).Text).ToString("N2")
                    'lSingleItem.SubItems.Add(.Item("Type").ToString)
                End With

                For x = 0 To LvSt1.Items.Count - 1
                    If (x Mod 2) = 0 Then
                        LvSt1.Items(x).BackColor = Color.MintCream

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
            SQL1.CommandText = "SELECT *FROM m_refPartList WHERE (Model='" & cmbRef.Text & "')AND(Station=2) ORDER BY Item ASC"
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
                    lSingleItem = LvSt2.Items.Add(.Item("Item").ToString)
                    lSingleItem.SubItems.Add(.Item("PartNo").ToString)
                    'lSingleItem.SubItems.Add(.Item("RefNo").ToString)
                    'lSingleItem.SubItems.Add(.Item("ItemNo").ToString)
                    'lSingleItem.SubItems(3).Text = CDbl(lSingleItem.SubItems(3).Text).ToString("N2")
                    'lSingleItem.SubItems.Add(.Item("Qty").ToString)
                    'lSingleItem.SubItems(4).Text = CDbl(lSingleItem.SubItems(4).Text).ToString("N2")
                    'lSingleItem.SubItems.Add(.Item("Type").ToString)
                End With

                For x = 0 To LvSt2.Items.Count - 1
                    If (x Mod 2) = 0 Then
                        LvSt2.Items(x).BackColor = Color.MintCream

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
            SQL1.CommandText = "SELECT *FROM m_refPartList WHERE (Model='" & cmbRef.Text & "')AND(Station=3) ORDER BY Item ASC"
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
                    lSingleItem.SubItems.Add(.Item("PartNo").ToString)
                    'lSingleItem.SubItems.Add(.Item("RefNo").ToString)
                    'lSingleItem.SubItems.Add(.Item("ItemNo").ToString)
                    'lSingleItem.SubItems(3).Text = CDbl(lSingleItem.SubItems(3).Text).ToString("N2")
                    'lSingleItem.SubItems.Add(.Item("Qty").ToString)
                    'lSingleItem.SubItems(4).Text = CDbl(lSingleItem.SubItems(4).Text).ToString("N2")
                    'lSingleItem.SubItems.Add(.Item("Type").ToString)
                End With

                For x = 0 To LvSt3.Items.Count - 1
                    If (x Mod 2) = 0 Then
                        LvSt3.Items(x).BackColor = Color.MintCream

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
            SQL1.CommandText = "SELECT *FROM m_refPartList WHERE (Model='" & cmbRef.Text & "')AND(Station=4) ORDER BY Item ASC"
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
                    lSingleItem.SubItems.Add(.Item("PartNo").ToString)
                    'lSingleItem.SubItems.Add(.Item("RefNo").ToString)
                    'lSingleItem.SubItems.Add(.Item("ItemNo").ToString)
                    'lSingleItem.SubItems(3).Text = CDbl(lSingleItem.SubItems(3).Text).ToString("N2")
                    'lSingleItem.SubItems.Add(.Item("Qty").ToString)
                    'lSingleItem.SubItems(4).Text = CDbl(lSingleItem.SubItems(4).Text).ToString("N2")
                    'lSingleItem.SubItems.Add(.Item("Type").ToString)
                End With

                For x = 0 To LvSt4.Items.Count - 1
                    If (x Mod 2) = 0 Then
                        LvSt4.Items(x).BackColor = Color.MintCream

                    End If
                Next x

            Next i
            'Call TotalOrder()

            'Call TotalItem()

            conn.Close()
        Catch ex As Exception

        End Try
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        'lblTime.Text = Format(Now, "HH:mm:ss").ToString

    End Sub
    Sub NoItem()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "SELECT TOP 1 Item FROM m_refPartList WHERE (Model='" & cmbRef.Text & "') ORDER BY Item DESC"
            Reader = SQL1.ExecuteReader
            If Reader.Read Then
                'Dim addno As Double
                'addno = Reader("Item") + 1
                lblItem.Text = Reader("Item") + 1
            Else
                lblItem.Text = "1"
                'lblItem.Text = "1"
            End If
            Conn.Close()
            Reader.Close()

        Catch ex As Exception

        End Try
    End Sub
    Sub NoItemStation()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "SELECT TOP 1 ItemStation FROM m_refPartList WHERE (Model='" & cmbRef.Text & "')AND(Station=" & Val(txtStation.Text) & ") ORDER BY Item DESC"
            Reader = SQL1.ExecuteReader
            If Reader.Read Then
                'Dim addno As Double
                'addno = Reader("Item") + 1
                lblItemStation.Text = Reader("ItemStation") + 1
            Else
                lblItemStation.Text = "1"
                'lblItem.Text = "1"
            End If
            Conn.Close()
            Reader.Close()

        Catch ex As Exception

        End Try
    End Sub
    Sub MasukItem()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Try
            Call NoItem()
            Call NoItemStation()

            Conn = GetConnect()
            Conn.Open()

            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "INSERT INTO m_refPartList" _
                            & "(Model,Item,Station,PartNo,PokNo,Scan,ItemStation)" _
                            & "VALUES ('" & cmbRef.Text & "'," & Val(lblItem.Text) & "," & Val(txtStation.Text) & ",'" & txtPartNo.Text & "'," & Val(txtPokNo.Text) & ",0," & Val(lblItemStation.Text) & ")"
            SQL1.ExecuteReader()
            Conn.Close()

            Call UpdateNoMW()

            Call dpStation1()
            Call dpStation2()
            Call dpStation3()
            Call dpStation4()

        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString)
        End Try
    End Sub

    Private Sub txtPartNo_KeyDown(sender As Object, e As KeyEventArgs) Handles txtPartNo.KeyDown
        If e.KeyCode = Keys.Enter Then
            If txtStation.Text = "" Or txtPartNo.Text = "" Then
                MessageBox.Show("Complete Station number and Part Number", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Else
                'Call MasukItem()
                txtPokNo.Focus()
            End If
        End If
    End Sub

    Private Sub txtPokNo_KeyDown(sender As Object, e As KeyEventArgs) Handles txtPokNo.KeyDown
        If e.KeyCode = Keys.Enter Then
            If txtPokNo.Text = "" Then
                MessageBox.Show("Complete Station Number, Part No and Pokapick Number", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Else
                Call MasukItem()
                txtPartNo.Text = ""
                txtStation.Text = ""
                txtPokNo.Text = ""

                txtStation.Focus()

            End If
        End If
    End Sub
    Sub Panggil()
        Dim Conn As SqlConnection
        Dim SQl1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()

            SQl1 = Conn.CreateCommand
            SQl1.CommandText = "SELECT *FROM m_refmanual WHERE (Model='" & cmbRef.Text & "')"
            Reader = SQl1.ExecuteReader
            If Reader.Read Then
                lblVolt.Text = Reader("Voltage").ToString
                lblLayer.Text = Reader("Layer").ToString
                lblBuzz.Text = Reader("Buzzer").ToString
                lblTube.Text = Reader("Tube").ToString
                lblFoot.Text = Reader("Foot").ToString

                Call dpStation1()
                Call dpStation2()
                Call dpStation3()
                Call dpStation4()

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString)
        End Try
    End Sub
    Sub UpdatePart()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "UPDATE m_refmanual SET PartSt='Y' WHERE (Model='" & cmbRef.Text & "')"
            SQL1.ExecuteReader()
            Conn.Close()

            Call dpList()
            Call FormBlank()

        Catch ex As Exception

        End Try
    End Sub
    Private Sub CmbRef_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbRef.SelectedIndexChanged

        Call Panggil()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()

    End Sub

    Private Sub txtStation_KeyDown(sender As Object, e As KeyEventArgs) Handles txtStation.KeyDown
        If e.KeyCode = Keys.Enter Then
            If txtStation.Text = "" Then
                MessageBox.Show("Entry Station Number", "Error", MessageBoxButtons.OK)
            Else
                txtPartNo.Focus()
            End If
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Call UpdatePart()
        Call FormBlank()
        cmbRef.Focus()

    End Sub

    Private Sub LvSt1_KeyDown(sender As Object, e As KeyEventArgs) Handles LvSt1.KeyDown
        If e.KeyCode = Keys.Delete Then
            Dim i As Integer
            For i = 0 To LvSt1.SelectedItems.Count - 1
                lbldel.Text = LvSt1.SelectedItems(i).SubItems(0).Text
                Dim Conn As SqlConnection
                Dim SQL1 As New SqlCommand
                Try
                    Conn = GetConnect()
                    Conn.Open()
                    SQL1 = Conn.CreateCommand
                    SQL1.CommandText = "DELETE FROM m_refPartList WHERE (Model='" & cmbRef.Text & "')AND(PartNo='" & LvSt1.SelectedItems(i).SubItems(1).Text & "')"
                    SQL1.ExecuteReader()

                    Call dpStation1()
                    txtStation.Focus()

                Catch ex As Exception

                End Try
            Next
        End If
    End Sub

    Private Sub LvSt1_DoubleClick(sender As Object, e As EventArgs) Handles LvSt1.DoubleClick

    End Sub

    Private Sub LvSt1_Click(sender As Object, e As EventArgs) Handles LvSt1.Click
        Dim i As Integer
        For i = 0 To LvSt1.SelectedItems.Count - 1
            lbldel.Text = LvSt1.SelectedItems(i).SubItems(0).Text

        Next
    End Sub

    Private Sub LvSt1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles LvSt1.SelectedIndexChanged

    End Sub

    Private Sub LvSt2_KeyDown(sender As Object, e As KeyEventArgs) Handles LvSt2.KeyDown
        If e.KeyCode = Keys.Delete Then
            Dim i As Integer
            For i = 0 To LvSt2.SelectedItems.Count - 1
                lbldel.Text = LvSt2.SelectedItems(i).SubItems(0).Text
                Dim Conn As SqlConnection
                Dim SQL1 As New SqlCommand
                Try
                    Conn = GetConnect()
                    Conn.Open()
                    SQL1 = Conn.CreateCommand
                    SQL1.CommandText = "DELETE FROM m_refPartList WHERE (Model='" & cmbRef.Text & "')AND(PartNo='" & LvSt2.SelectedItems(i).SubItems(1).Text & "')"
                    SQL1.ExecuteReader()

                    Call dpStation2()
                    txtStation.Focus()

                Catch ex As Exception

                End Try
            Next
        End If
    End Sub

    Private Sub LvSt3_KeyDown(sender As Object, e As KeyEventArgs) Handles LvSt3.KeyDown
        If e.KeyCode = Keys.Delete Then
            Dim i As Integer
            For i = 0 To LvSt3.SelectedItems.Count - 1
                lbldel.Text = LvSt3.SelectedItems(i).SubItems(0).Text
                Dim Conn As SqlConnection
                Dim SQL1 As New SqlCommand
                Try
                    Conn = GetConnect()
                    Conn.Open()
                    SQL1 = Conn.CreateCommand
                    SQL1.CommandText = "DELETE FROM m_refPartList WHERE (Model='" & cmbRef.Text & "')AND(PartNo='" & LvSt3.SelectedItems(i).SubItems(1).Text & "')"
                    SQL1.ExecuteReader()

                    Call dpStation3()
                    txtStation.Focus()

                Catch ex As Exception

                End Try
            Next
        End If
    End Sub

    Private Sub LvSt4_KeyDown(sender As Object, e As KeyEventArgs) Handles LvSt4.KeyDown
        If e.KeyCode = Keys.Delete Then
            Dim i As Integer
            For i = 0 To LvSt4.SelectedItems.Count - 1
                lbldel.Text = LvSt4.SelectedItems(i).SubItems(0).Text
                Dim Conn As SqlConnection
                Dim SQL1 As New SqlCommand
                Try
                    Conn = GetConnect()
                    Conn.Open()
                    SQL1 = Conn.CreateCommand
                    SQL1.CommandText = "DELETE FROM m_refPartList WHERE (Model='" & cmbRef.Text & "')AND(PartNo='" & LvSt4.SelectedItems(i).SubItems(1).Text & "')"
                    SQL1.ExecuteReader()

                    Call dpStation4()
                    txtStation.Focus()

                Catch ex As Exception

                End Try
            Next
        End If
    End Sub
End Class