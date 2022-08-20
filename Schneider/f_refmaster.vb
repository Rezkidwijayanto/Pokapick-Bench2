Imports System.Data
Imports System.Data.SqlClient

Public Class f_refmaster
    Private Sub BtnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close()

    End Sub

    Private Sub f_refmaster_Load(sender As Object, e As EventArgs) Handles Me.Load
        txtCari.Text = ""
        txtRef.Text = ""

        txtRef.Focus()
        Call dpItem()

    End Sub
    Sub NoID()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "SELECT TOP 1 *FROM m_refmaster ORDER BY NoID DESC"
            Reader = SQL1.ExecuteReader
            If Reader.Read Then
                lblNoID.Text = Val(Reader("NoID")) + 1
            Else
                lblNoID.Text = "1"
            End If
            Conn.Close()
            Reader.Close()

        Catch ex As Exception

        End Try
    End Sub
    Sub SimpanReference()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Try
            Call NoID()

            Conn = GetConnect()
            Conn.Open()
            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "INSERT INTO m_refmaster" _
                            & "(NoID,refNO)" _
                            & "VALUES (" & Val(lblNoID.Text) & ",'" & txtRef.Text & "')"
            SQL1.ExecuteReader()

            Call SimpantoMaster()

            Call dpItem()

            txtRef.Text = ""
            txtRef.Focus()

        Catch ex As Exception

        End Try
    End Sub
    Sub SimpantoMaster()
        Dim Conn As SqlConnection
        Dim SQl1 As New SqlCommand
        Dim Tanggal = Format(Now, "yyyy/MM/dd HH:mm:ss").ToString

        Try
            Conn = GetConnect()
            Conn.Open()
            SQl1 = Conn.CreateCommand
            SQl1.CommandText = "INSERT INTO m_refmanual" _
                            & "(NoID,Model,iDate,Status)" _
                            & "VALUES (" & Val(lblNoID.Text) & ",'" & txtRef.Text & "','" & Tanggal & "','O')"
            SQl1.ExecuteReader()
            Conn.Close()

        Catch ex As Exception

        End Try
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
            SQL1.CommandText = "SELECT *FROM m_refmaster ORDER BY NoID ASC"
            Pilih.SelectCommand = SQL1
            Pilih.Fill(datWarga, "m_refmaster")
            datTable = datWarga.Tables("m_refmaster")

            Dim i As Integer
            Dim x As Integer

            'Displaydata()
            Lv1.Items.Clear()
            For i = 0 To (datTable.Rows.Count - 1)
                With datTable.Rows(i)
                    Dim lSingleItem As ListViewItem
                    'Dim Tot = FormatCurrency("Amount")
                    lSingleItem = Lv1.Items.Add(.Item("NoID").ToString)
                    lSingleItem.SubItems.Add(.Item("refNO").ToString)
                    'lSingleItem.SubItems.Add(.Item("Model").ToString)
                    'lSingleItem.SubItems.Add(.Item("ItemNo").ToString)
                    'lSingleItem.SubItems(3).Text = CDbl(lSingleItem.SubItems(3).Text).ToString("N2")
                    'lSingleItem.SubItems.Add(.Item("Qty").ToString)
                    'lSingleItem.SubItems(4).Text = CDbl(lSingleItem.SubItems(4).Text).ToString("N2")
                    'lSingleItem.SubItems.Add(.Item("Type").ToString)
                End With

                For x = 0 To Lv1.Items.Count - 1
                    If (x Mod 2) = 0 Then
                        Lv1.Items(x).BackColor = Color.MintCream

                    End If
                Next x

            Next i
            'Call TotalOrder()

            'Call TotalItem()

            conn.Close()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Lv1_KeyDown(sender As Object, e As KeyEventArgs) Handles Lv1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim i As Integer
            For i = 0 To Lv1.SelectedItems.Count - 1
                Dim Conn As SqlConnection
                Dim SQl1 As New SqlCommand
                Try
                    Conn = GetConnect()
                    Conn.Open()
                    SQl1 = Conn.CreateCommand
                    SQl1.CommandText = "DELETE FROM m_refmaster WHERE (NoID=" & Val(Lv1.SelectedItems(i).SubItems(0).Text) & ")"
                    SQl1.ExecuteReader()
                    Conn.Close()

                    Call Deletem_refmanual()
                    Call dpItem()
                    txtRef.Text = ""
                    txtRef.Focus()

                Catch ex As Exception

                End Try
            Next
        End If
        If e.KeyCode = Keys.Escape Then
            txtRef.Text = ""
            txtRef.Focus()
        End If

    End Sub
    Sub Deletem_refmanual()
        Dim Conn As SqlConnection
        Dim SQl1 As New SqlCommand
        Try
            Dim i As Integer
            For i = 0 To Lv1.SelectedItems.Count - 1
                Conn = GetConnect()
                Conn.Open()
                SQl1 = Conn.CreateCommand
                SQl1.CommandText = "DELETE FROM m_refmanual WHERE (NoID=" & Val(Lv1.SelectedItems(i).SubItems(0).Text) & ")"
                SQl1.ExecuteReader()
                Conn.Close()

            Next
        Catch ex As Exception

        End Try
    End Sub
    Private Sub BtnDel_Click(sender As Object, e As EventArgs) Handles btnDel.Click
        Dim i As Integer
        For i = 0 To Lv1.SelectedItems.Count - 1
            Dim Conn As SqlConnection
            Dim SQl1 As New SqlCommand
            Try
                Conn = GetConnect()
                Conn.Open()
                SQl1 = Conn.CreateCommand
                SQl1.CommandText = "DELETE FROM m_refmaster WHERE (NoID=" & Val(Lv1.SelectedItems(i).SubItems(0).Text) & ")"
                SQl1.ExecuteReader()
                Conn.Close()

                Call dpItem()
                txtRef.Text = ""
                txtRef.Focus()

            Catch ex As Exception

            End Try
        Next
    End Sub
    Sub CekSimpan()
        Dim Conn As SqlConnection
        Dim SQl1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()
            SQl1 = Conn.CreateCommand
            SQl1.CommandText = "SELECT *FROM m_refmaster WHERE (refNO='" & txtRef.Text & "')"
            Reader = SQl1.ExecuteReader
            If Reader.Read Then
                MessageBox.Show("Nomor Reference sudah terdaftar!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtRef.Focus()
                Exit Sub
            Else
                Call SimpanReference()
            End If
            Conn.Close()
            Reader.Close()

        Catch ex As Exception

        End Try
    End Sub
    Private Sub BtnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If txtRef.Text = "" Then
            MessageBox.Show("Lengkapi nomor reference!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            'Call SimpanReference()
            Call CekSimpan()

        End If
    End Sub
End Class