Imports System
Imports System.ComponentModel
Imports System.Data.SqlClient

Public Class f_Open
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
            SQL1.CommandText = "SELECT *FROM m_refrunning WHERE (Status='S')" 'AND(Station=1) ORDER BY Item ASC"
            Pilih.SelectCommand = SQL1
            Pilih.Fill(datWarga, "m_refrunning")
            datTable = datWarga.Tables("m_refrunning")

            Dim i As Integer
            Dim x As Integer

            'Displaydata()
            Lv1.Items.Clear()
            For i = 0 To (datTable.Rows.Count - 1)
                With datTable.Rows(i)
                    Dim lSingleItem As ListViewItem
                    'Dim Tot = FormatCurrency("Amount")
                    lSingleItem = Lv1.Items.Add(.Item("iDate").ToString)
                    lSingleItem.SubItems(0).Text = CDate(lSingleItem.SubItems(0).Text).ToString("dd/MM/yyyy HH:mm:ss")
                    lSingleItem.SubItems.Add(.Item("SONo").ToString)
                    lSingleItem.SubItems.Add(.Item("Model").ToString)
                    lSingleItem.SubItems.Add(.Item("Qty").ToString)
                    lSingleItem.SubItems.Add(.Item("RunQty").ToString)
                    lSingleItem.SubItems.Add(.Item("BarCode").ToString)
                    lSingleItem.SubItems.Add(.Item("Voltage").ToString)
                    'lSingleItem.SubItems(3).Text = CDbl(lSingleItem.SubItems(3).Text).ToString("N2")
                    lSingleItem.SubItems.Add(.Item("Layer").ToString)
                    'lSingleItem.SubItems(4).Text = CDbl(lSingleItem.SubItems(4).Text).ToString("N2")
                    lSingleItem.SubItems.Add(.Item("Buzzer").ToString)
                    lSingleItem.SubItems.Add(.Item("NoID").ToString)

                End With



            Next i

            'Call TotalSt1Masterial()
            'Call St1MasterialScan()

            'Call CekScanStatusSt1()

            conn.Close()
        Catch ex As Exception
            'MessageBox.Show(ex.Message.ToString)
        End Try
    End Sub
    Sub dpCari()
        Try
            Dim conn As SqlConnection
            Dim SQL1 As New SqlCommand
            Dim Pilih As New SqlDataAdapter
            Dim datWarga As New DataSet
            Dim datTable As New DataTable

            conn = GetConnect()
            conn.Open()

            SQL1 = conn.CreateCommand
            SQL1.CommandText = "SELECT *FROM m_refrunning WHERE (Model Like '%" & Trim(txtSearch.Text) & "%' OR SONo Like '%" & Trim(txtSearch.Text) & "%')AND (Status='S')" 'AND(Station=1) ORDER BY Item ASC"
            Pilih.SelectCommand = SQL1
            Pilih.Fill(datWarga, "m_refrunning")
            datTable = datWarga.Tables("m_refrunning")

            Dim i As Integer
            Dim x As Integer

            'Displaydata()
            Lv1.Items.Clear()
            For i = 0 To (datTable.Rows.Count - 1)
                With datTable.Rows(i)
                    Dim lSingleItem As ListViewItem
                    'Dim Tot = FormatCurrency("Amount")
                    lSingleItem = Lv1.Items.Add(.Item("iDate").ToString)
                    lSingleItem.SubItems(0).Text = CDate(lSingleItem.SubItems(0).Text).ToString("dd/MM/yyyy HH:mm:ss")
                    lSingleItem.SubItems.Add(.Item("SONo").ToString)
                    lSingleItem.SubItems.Add(.Item("Model").ToString)
                    lSingleItem.SubItems.Add(.Item("Qty").ToString)
                    lSingleItem.SubItems.Add(.Item("RunQty").ToString)
                    lSingleItem.SubItems.Add(.Item("BarCode").ToString)
                    lSingleItem.SubItems.Add(.Item("Voltage").ToString)
                    'lSingleItem.SubItems(3).Text = CDbl(lSingleItem.SubItems(3).Text).ToString("N2")
                    lSingleItem.SubItems.Add(.Item("Layer").ToString)
                    'lSingleItem.SubItems(4).Text = CDbl(lSingleItem.SubItems(4).Text).ToString("N2")
                    lSingleItem.SubItems.Add(.Item("Buzzer").ToString)
                    lSingleItem.SubItems.Add(.Item("NoID").ToString)

                End With



            Next i

            'Call TotalSt1Masterial()
            'Call St1MasterialScan()

            'Call CekScanStatusSt1()

            conn.Close()
        Catch ex As Exception
            'MessageBox.Show(ex.Message.ToString)
        End Try
    End Sub

    Private Sub f_Open_Load(sender As Object, e As EventArgs) Handles Me.Load
        ResizeFormClass.SubResize(Me, (Me.Width / 1280) * 100, (Me.Height / 720) * 100)
        Call dpStation1()

    End Sub

    Private Sub txtSearch_KeyDown(sender As Object, e As KeyEventArgs) Handles txtSearch.KeyDown
        If e.KeyCode = Keys.Enter Then
            If txtSearch.Text <> "" Then
                Call dpCari()
            End If
        End If
        If e.KeyCode = Keys.Down Then
            Lv1.Focus()
        End If
    End Sub

    Private Sub f_Open_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing


    End Sub

    Private Sub Lv1_DoubleClick(sender As Object, e As EventArgs) Handles Lv1.DoubleClick
        Try
            Dim i As Integer
            For i = 0 To Lv1.SelectedItems.Count - 1
                lblNoID.Text = Lv1.SelectedItems(i).SubItems(9).Text
                f_utama.lblNoID.Text = lblNoID.Text

                Me.Close()
                f_utama.ShowDialog()

            Next
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Lv1_Click(sender As Object, e As EventArgs) Handles Lv1.Click
        Try
            Dim i As Integer
            For i = 0 To Lv1.SelectedItems.Count - 1
                lblNoID.Text = Lv1.SelectedItems(i).SubItems(9).Text
            Next
        Catch ex As Exception

        End Try
    End Sub

    Private Sub f_Open_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Private Sub f_Open_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        Me.Close()
        f_utama.ShowDialog()
    End Sub
End Class