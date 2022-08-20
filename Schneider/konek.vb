Imports System.Data.SqlClient
Imports System.Data.Sql
Imports System.Net
Imports System.IO

Module Konek
    Public conn As SqlConnection
    Public Tgl As DateTime
    'Public MikrotikUser As String

    Public Function GetConnect()
        Dim TW As New StreamReader("String.ini")
        Dim koneksi As String
        koneksi = TW.ReadLine
        conn = New SqlConnection(koneksi)
        'conn = New SqlConnection("server = QONOD_1;database = Timbangan;Trusted_Connection = yes")

        Return conn
    End Function
    Public MkUser As String

    Public Function MikrotikUser() As String
        Dim SQL1 As New SqlCommand
        Dim Conn As SqlConnection
        Dim Reader As SqlDataReader

        Conn = GetConnect()
        Conn.Open()
        SQL1 = Conn.CreateCommand
        SQL1.CommandText = "SELECT *FROM s_Config WHERE (NoID=1)"
        Reader = SQL1.ExecuteReader
        If Reader.Read Then
            Dim uName As String
            uName = Reader("MikrotikUser").ToString
            MkUser = New String(uName)

        End If
        Conn.Close()
        Reader.Close()

        Return MkUser

    End Function
    Public MkPass As String
    Public Function MikrotikPass() As String
        Dim SQL1 As New SqlCommand
        Dim Conn As SqlConnection
        Dim Reader As SqlDataReader

        Conn = GetConnect()
        Conn.Open()
        SQL1 = Conn.CreateCommand
        SQL1.CommandText = "SELECT *FROM s_Config WHERE (NoID=1)"
        Reader = SQL1.ExecuteReader
        If Reader.Read Then
            Dim uPass As String
            uPass = Reader("MikrotikPass").ToString
            MkPass = New String(uPass)

        End If
        Conn.Close()
        Reader.Close()

        Return MkPass

    End Function
    Public IPAdd As String
    Public Function IPMikrotik() As String
        Dim SQL1 As New SqlCommand
        Dim Conn As SqlConnection
        Dim Reader As SqlDataReader

        Conn = GetConnect()
        Conn.Open()
        SQL1 = Conn.CreateCommand
        SQL1.CommandText = "SELECT *FROM s_Config WHERE (NoID=1)"
        Reader = SQL1.ExecuteReader
        If Reader.Read Then
            Dim uIP As String
            uIP = Reader("IPAddress").ToString
            IPAdd = New String(uIP)

        End If
        Conn.Close()
        Reader.Close()

        Return IPAdd

    End Function
    Public KitPrint As String
    Public Function KitchenPrinter() As String
        Dim SQL1 As New SqlCommand
        Dim Conn As SqlConnection
        Dim Reader As SqlDataReader

        Conn = GetConnect()
        Conn.Open()
        SQL1 = Conn.CreateCommand
        SQL1.CommandText = "SELECT *FROM s_Config WHERE (NoID=1)"
        Reader = SQL1.ExecuteReader
        If Reader.Read Then
            Dim ktPrint As String
            ktPrint = Reader("KitPrint").ToString
            KitPrint = New String(ktPrint)

        End If
        Conn.Close()
        Reader.Close()

        Return KitPrint

    End Function
    Public CsPrint As String
    Public Function CashierPrinter() As String
        Dim SQL1 As New SqlCommand
        Dim Conn As SqlConnection
        Dim Reader As SqlDataReader

        Conn = GetConnect()
        Conn.Open()
        SQL1 = Conn.CreateCommand
        SQL1.CommandText = "SELECT *FROM s_Config WHERE (NoID=1)"
        Reader = SQL1.ExecuteReader
        If Reader.Read Then
            Dim CashPrint As String
            CashPrint = Reader("CashPrint").ToString
            CsPrint = New String(CashPrint)

        End If
        Conn.Close()
        Reader.Close()

        Return CsPrint

    End Function
    Public BrPrint As String
    Public Function BarPrinter() As String
        Dim SQL1 As New SqlCommand
        Dim Conn As SqlConnection
        Dim Reader As SqlDataReader

        Conn = GetConnect()
        Conn.Open()
        SQL1 = Conn.CreateCommand
        SQL1.CommandText = "SELECT *FROM s_Config WHERE (NoID=1)"
        Reader = SQL1.ExecuteReader
        If Reader.Read Then
            Dim BarPrint As String
            BarPrint = Reader("BarPrint").ToString
            BrPrint = New String(BarPrint)

        End If
        Conn.Close()
        Reader.Close()

        Return BrPrint

    End Function
    Public nOutlet As String
    Public Function NamaOutlet() As String
        Dim SQL1 As New SqlCommand
        Dim COnn As SqlConnection
        Dim Reader As SqlDataReader

        COnn = GetConnect()
        COnn.Open()
        SQL1 = COnn.CreateCommand
        SQL1.CommandText = "SELECT *FROM s_Config WHERE (NoID=1)"
        Reader = SQL1.ExecuteReader
        If Reader.Read Then
            Dim Outlet As String
            Outlet = Reader("Company").ToString
            nOutlet = New String(Outlet)

        End If
        COnn.Close()
        Reader.Close()

        Return nOutlet

    End Function
    Public Function Tgl1()
        Dim conn As SqlConnection
        Dim SQL1 As New SqlCommand
        Dim Reader As SqlDataReader
        'Dim Tanggal As String

        conn = GetConnect()
        conn.Open()
        SQL1 = conn.CreateCommand
        'SQL1.CommandText = "SELECT sysdate FROM DUAL"
        SQL1.CommandText = "SELECT getdate() as ServerTime"
        Reader = SQL1.ExecuteReader
        Tgl = Reader.ToString

        Return Tgl

    End Function
End Module
