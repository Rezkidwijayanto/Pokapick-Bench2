Imports System.Data
Imports System.Data.SqlClient

Public Class FrmLogin
    Sub LihatCode()
        Dim Conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader
        Try
            Conn = GetConnect()
            Conn.Open()

            SQL1 = Conn.CreateCommand
            SQL1.CommandText = "SELECT *FROM s_users WHERE (IDCode='" & txtUser.Text & "')"
            Reader = SQL1.ExecuteReader
            If Reader.Read Then
                Me.Hide()

                Dim utama As New f_utama
                utama.lblOperator.Text = Reader("FullName").ToString
                utama.lblDept.Text = Reader("Dept").ToString
                Call utama.FormBlank()


                utama.ShowDialog()

            Else
                Call Masuk()

            End If

        Catch ex As Exception

        End Try
    End Sub
    Sub Masuk()
        Dim conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader

        Try
            conn = GetConnect()
            conn.Open()

            SQL1 = conn.CreateCommand
            SQL1.CommandText = "SELECT *FROM s_users WHERE (UserID='" & txtUser.Text & "')AND(Pass='" & txtPass.Text & "')"
            Reader = SQL1.ExecuteReader

            If Reader.Read = True Then

                'Mulai.lblUser.Text = txtUser.Text
                'Mulai.ShowDialog()
                'Me.Close()
                'Mulai.ShowDialog()
                'Form1.lblUser.Text = txtUser.Text
                'Form1.lblLevel.Text = Reader("uLevel").ToString

                'Me.Hide()
                'frmLogin.Hide()

                Dim utama As New f_utama
                utama.lblOperator.Text = Reader("FullName").ToString
                utama.lblDept.Text = Reader("Dept").ToString
                'Call utama.FormBlank()

                utama.ShowDialog()
                'f_utama.lblOperator.Text = Reader("FullName").ToString
                'f_utama.lblDept.Text = Reader("uLevel").ToString
                'f_utama.ShowDialog()
                'Close()
                'Form1.ShowDialog()

            Else
                MessageBox.Show("Invalid User ID and Password!")
                txtUser.Text = ""
                txtPass.Text = ""
                txtUser.Focus()
            End If
            'Me.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        End

    End Sub

    Private Sub frmLogin_Load(sender As Object, e As EventArgs) Handles Me.Load
        txtUser.Text = ""
        txtPass.Text = ""

        txtUser.Focus()

    End Sub

    Private Sub txtUser_KeyDown(sender As Object, e As KeyEventArgs) Handles txtUser.KeyDown
        If e.KeyCode = Keys.Enter Then
            If txtUser.Text = "" Then
                MessageBox.Show("SCAN ID / masukan user dan password", "Error", MessageBoxButtons.OK)
                txtUser.Focus()
            Else
                Call LihatCode()

            End If
        End If
    End Sub

    Private Sub txtPass_KeyDown(sender As Object, e As KeyEventArgs) Handles txtPass.KeyDown
        If e.KeyCode = Keys.Enter Then
            Call Masuk()

        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

    End Sub
End Class
