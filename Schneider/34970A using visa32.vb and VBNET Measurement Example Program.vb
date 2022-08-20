Option Strict Off
Option Explicit On 

Public Class VISAExample
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents List1 As System.Windows.Forms.ListBox
    Friend WithEvents GetReadings As System.Windows.Forms.Button
    Friend WithEvents EndProg As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ioType As System.Windows.Forms.TextBox
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents SelectIO As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(VISAExample))
        Me.List1 = New System.Windows.Forms.ListBox()
        Me.GetReadings = New System.Windows.Forms.Button()
        Me.EndProg = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ioType = New System.Windows.Forms.TextBox()
        Me.SelectIO = New System.Windows.Forms.Button()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.SuspendLayout()
        '
        'List1
        '
        Me.List1.Location = New System.Drawing.Point(16, 136)
        Me.List1.Name = "List1"
        Me.List1.Size = New System.Drawing.Size(304, 329)
        Me.List1.TabIndex = 0
        '
        'GetReadings
        '
        Me.GetReadings.Location = New System.Drawing.Point(120, 96)
        Me.GetReadings.Name = "GetReadings"
        Me.GetReadings.Size = New System.Drawing.Size(96, 32)
        Me.GetReadings.TabIndex = 0
        Me.GetReadings.Text = "Get Readings"
        '
        'EndProg
        '
        Me.EndProg.Location = New System.Drawing.Point(120, 477)
        Me.EndProg.Name = "EndProg"
        Me.EndProg.Size = New System.Drawing.Size(96, 32)
        Me.EndProg.TabIndex = 1
        Me.EndProg.Text = "Exit"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(44, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(248, 40)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Enter/select Instrument Address (e.g., GPIB0::9, ASRL1, etc.), click on ""Select I" & _
            "/O"", and then click on ""Get Readings""."
        '
        'ioType
        '
        Me.ioType.Location = New System.Drawing.Point(80, 48)
        Me.ioType.Name = "ioType"
        Me.ioType.Size = New System.Drawing.Size(64, 20)
        Me.ioType.TabIndex = 3
        Me.ioType.Text = "USB0::2391::8199::MY57010982::0"
        '
        'SelectIO
        '
        Me.SelectIO.Location = New System.Drawing.Point(160, 40)
        Me.SelectIO.Name = "SelectIO"
        Me.SelectIO.Size = New System.Drawing.Size(96, 32)
        Me.SelectIO.TabIndex = 4
        Me.SelectIO.Text = "Select I/O"
        '
        'Timer1
        '
        '
        'VISAExample
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(340, 531)
        Me.Controls.Add(Me.SelectIO)
        Me.Controls.Add(Me.ioType)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.EndProg)
        Me.Controls.Add(Me.GetReadings)
        Me.Controls.Add(Me.List1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "VISAExample"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "VISA Example"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    '"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    '"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    '34970A/34972A using visa32.vb and VBNET Measurement Example Program - This program sets the
    ' 34970A/34972A for 2 simple scans that make both DCV and type J thermocouple temperature
    ' measurements. The program also checks the instrument and module identification to
    ' make sure there is communication between the 34970A/34972A and the computer, and the
    ' correct module is installed. The returned data is entered into a list box.
    '
    ' Additonally, if the instrument is a 34972A, logging to a USB memory stick is enabled.
    ' With logging enabled, readings will be saved to internal memory and to USB memory if
    ' USB memory is present.  If USB memory stick is not present, then only internal memory
    ' will be filled with the readings.
    '
    ' This program requires a 34901A module
    '
    ' Included is an error checking routine to make sure the SCPI commands executed have
    ' the correct syntax.
    '
    ' The program requires that VISA is installed in your computer. VISA comes with the
    ' Agilent I/O Library.
    '
    ' Include the visa32.vb file. This file comes with VISA version M.01.01.041 or 
    ' above. This file is different from the visa32.bas since the visa32.vb was developed 
    ' for VB.NET (the visa32.bas is for VB version 6 and below).  The file may be found in the Include
    ' directory of where Visa is installed.  On the most recent IO Libraries, this would be in:
    '   C:\Program Files\IVI Foundation\VISA\WinNT\include
    '
    ' The visa32.vb file is also included in the directory of this example.
    '
    '
    '
    ' The program was developed in Microsoft® Visual Basic .NET version 7.0
    '
    ' The program can use either the user selectable GPIB or RS-232 interface. If
    ' selecting RS-232, the 34970A must be set to the following RS-232 parameters:
    '   Baud Rate: 115200
    '   Parity: None
    '   Data bits: 8
    '   Start bits: 1
    '   Stop bits: 1
    '   Flow control: XON/XOFF
    '
    '"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    ' Copyright © 2004,2009 Agilent Technologies Inc. All rights reserved.
    '
    ' You have a royalty-free right to use, modify, reproduce and distribute this
    ' example files (and/or any modified version) in any way you find useful, provided
    ' that you agree that Agilent has no warranty, obligations or liability for any
    ' Sample Application Files.
    '
    ' Agilent Technologies will not modify the program to provide added
    ' functionality or construct procedures to meet your specific needs.
    '"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

    Dim videfaultRM As Integer = 0  ' Resource manager session returned by viOpenDefaultRM(videfaultRM)
    Dim vi As Integer = 0           ' Session identifier of devices
    Dim errorStatus As Integer      ' VISA function status return code

    Dim connected As Boolean = False  ' Used to determine if there is connection with the instrument

    Dim ReturnedData As String ' Used to read returned data
    Dim NumRdgs As Long        ' Used for the number of readings taken
    Dim TotTime As Double      ' Used to calculate total measurement time
    Dim TrigCount As Integer    ' Used to determine number of scans
    Dim NumChan As Long         ' Used to determine the number of channels scanned

    '""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    Private Sub VISAExample_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        '"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
        'Loads the form.
        '"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
        List1.Items.Add("Enter/select instrument address, if needed,")
        List1.Items.Add("click on " + Chr(34) + "Select I/O" + Chr(34) + " to select the adress,")
        List1.Items.Add("and click on " + Chr(34) + "Get Readings" + Chr(34) + " to trigger instrument.")
        List1.Items.Add("Measurements will take some time.")
    End Sub

    '""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    Public Sub RunProgram()
        Try
            ' Call the function that opens communication with instrument
            If connected = False Then
                If Not OpenPort() Then
                    Exit Sub
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.ToString)
            If connected Then
                End_Prog()
            End If
            End
        End Try

    End Sub

    '""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    Private Sub Setup(ByVal chStart As Integer, ByVal chEnd As Integer)
        '""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
        ' This sub performs the instrument setup.
        '""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

        Dim DelayVal As Double
        Dim TrigTime As Double

        ' Check for exceptions
        Try
            ' Reset instrument to turn-on condition
            SendCmd("*RST")

            ' Configure for temperature measurements
            '  Select channels 101 to 110
            '  Type J thermocouple measurement
            '  5.5 digit (selected by *RST)
            ' SendCmd("CONFigure:TEMPerature TCouple, J, (@" & chStart & ":" & chEnd & ")")

            ' Select the temperature unit (C = Celcius)
            'SendCmd("UNIT:TEMPerature C, (@" & chStart & ":" & chEnd & ")")

            ' Set the reference temperature type (internal)
            'SendCmd("SENSe:TEMPerature:TRANSducer:TCouple:RJUNction:TYPE INTernal, (@" & chStart & ":" & chEnd & ")")

            ' Configure for voltage readings:
            '   Select channels 111 to 120
            '   DC volts
            '   10 V range
            '   5.5 digit (selected by *RST)
            SendCmd("CONFigure:VOLTage:DC 5, (@" & chStart & ":" & chEnd & ")") '>>>>>>>>>>>>>>>>>>>>>>>>> SKIP

            ' Set the NPLC value for channels 111 to 120
            SendCmd("SENSe:VOLTage:NPLC 1,(@" & chStart & ":" & chEnd & ")")  '>>>>>>>>>>>>>>>>>>>>>>>>> SKIP

            ' Select the scan list for channels 101 to 120 (all configured channels)
            SendCmd("ROUTe:SCAN (@" & chStart & ":" & chEnd & ")")

            ' Set the measurement delay between the channels
            SendCmd("ROUTe:CHANnel:DELay 0, (@" & chStart & ":" & chEnd & ")")

            ' Set number of sweeps to 2; use your own value
            SendCmd("TRIGger:COUNt 1")

            ' Set the trigger mode to TIMER (timed trigger); use your own type
            SendCmd("TRIGger:SOURce TIMer")

            ' Set the trigger time to 10 seconds (i.e., time between scans); use your own value
            SendCmd("TRIGger:TIMer 0")

            ' Format the reading time to show the time value from the start of the scan
            SendCmd("FORMat:READing:TIME:TYPE RELative")

            ' Add time stamp to reading using the selected time format
            SendCmd("FORMat:READing:TIME ON")

            ' Add the channel number to reading
            SendCmd("FORMat:READing:CHANnel ON")

            ' Wait for instrument to setup
            SendCmd("*OPC?")
            ReturnedData = GetData()

            ' Gets the number of channels to be scanned; used to determine the number of readings
            SendCmd("ROUTe:SCAN:SIZE?")
            NumChan = Val(GetData())

            ' Gets the number of triggers; used to determine the number of readings
            SendCmd("TRIGger:COUNt?")
            TrigCount = Val(GetData())

            ' Get the delay; for future use
            SendCmd("ROUTe:CHANnel:DELay? (@" & chStart & ")")
            DelayVal = Val(GetData())

            ' Get the trigger time
            SendCmd("TRIGger:TIMer?")
            TrigTime = Val(GetData())

            ' Calculate total number of readings
            NumRdgs = NumChan * TrigCount

            ' Calculate total time
            TotTime = (TrigTime * TrigCount) - TrigTime + (NumChan * DelayVal)

            'Check for errors
            Call Check_Error("Setup")

        Catch ex As Exception
            MsgBox(ex.ToString)
            If connected Then
                End_Prog()
            End If
            End
        End Try

    End Sub

    Private Sub SetupTemp(ByVal chStart As Integer, ByVal chEnd As Integer)
        '""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
        ' This sub performs the instrument setup.
        '""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

        Dim DelayVal As Double
        Dim TrigTime As Double

        ' Check for exceptions
        Try
            ' Reset instrument to turn-on condition
            SendCmd("*RST")

            ' Configure for temperature measurements
            '  Select channels 101 to 110
            '  Type J thermocouple measurement
            '  5.5 digit (selected by *RST)
            SendCmd("CONFigure:TEMPerature TCouple, J, (@" & chStart & ":" & chEnd & ")")

            ' Select the temperature unit (C = Celcius)
            SendCmd("UNIT:TEMPerature C, (@" & chStart & ":" & chEnd & ")")

            ' Set the reference temperature type (internal)
            SendCmd("SENSe:TEMPerature:TRANSducer:TCouple:RJUNction:TYPE INTernal, (@" & chStart & ":" & chEnd & ")")

            ' Configure for voltage readings:
            '   Select channels 111 to 120
            '   DC volts
            '   10 V range
            '   5.5 digit (selected by *RST)
            'SendCmd("CONFigure:VOLTage:DC 10, (@211:220)") >>>>>>>>>>>>>>>>>>>>>>>>> SKIP

            ' Set the NPLC value for channels 111 to 120
            'SendCmd("SENSe:VOLTage:NPLC 1,(@211:220)")  >>>>>>>>>>>>>>>>>>>>>>>>> SKIP

            ' Select the scan list for channels 101 to 120 (all configured channels)
            SendCmd("ROUTe:SCAN (@" & chStart & ":" & chEnd & ")")

            ' Set the measurement delay between the channels
            SendCmd("ROUTe:CHANnel:DELay 0, (@" & chStart & ":" & chEnd & ")")

            ' Set number of sweeps to 2; use your own value
            SendCmd("TRIGger:COUNt 1")

            ' Set the trigger mode to TIMER (timed trigger); use your own type
            SendCmd("TRIGger:SOURce TIMer")

            ' Set the trigger time to 10 seconds (i.e., time between scans); use your own value
            SendCmd("TRIGger:TIMer 0")

            ' Format the reading time to show the time value from the start of the scan
            SendCmd("FORMat:READing:TIME:TYPE RELative")

            ' Add time stamp to reading using the selected time format
            SendCmd("FORMat:READing:TIME ON")

            ' Add the channel number to reading
            SendCmd("FORMat:READing:CHANnel ON")

            ' Wait for instrument to setup
            SendCmd("*OPC?")
            ReturnedData = GetData()

            ' Gets the number of channels to be scanned; used to determine the number of readings
            SendCmd("ROUTe:SCAN:SIZE?")
            NumChan = Val(GetData())

            ' Gets the number of triggers; used to determine the number of readings
            SendCmd("TRIGger:COUNt?")
            TrigCount = Val(GetData())

            ' Get the delay; for future use
            SendCmd("ROUTe:CHANnel:DELay? (@" & chStart & ")")
            DelayVal = Val(GetData())

            ' Get the trigger time
            SendCmd("TRIGger:TIMer?")
            TrigTime = Val(GetData())

            ' Calculate total number of readings
            NumRdgs = NumChan * TrigCount

            ' Calculate total time
            TotTime = (TrigTime * TrigCount) - TrigTime + (NumChan * DelayVal)

            'Check for errors
            Call Check_Error("Setup")

        Catch ex As Exception
            MsgBox(ex.ToString)
            If connected Then
                End_Prog()
            End If
            End
        End Try

    End Sub


    Public valR(30) As Integer
    Public valTemperature(5) As Integer

    '""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    Sub Readings()
        '""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
        ' This sub triggers the instrument and takes readings.
        '""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

        Dim rdgs As String ' Returns a reading from memory
        Dim readMsg As String ' Stores data to show in ListBox
        Dim Dateval As String ' Returns the date value
        Dim Timeval As String ' Returns the time value
        Dim reading As String  ' Used to store reading
        Dim timestamp As String ' Used to store time stamp
        Dim channelnum As String ' Used to store channel number
        Dim I As Long ' Used for a loop

        ' Check for exceptions
        Try
            ' Trigger instrument
            SendCmd("INITiate") ' >>>>>>>>>>>>>>> SKIP

            ' Get the date at which the scan was started
            'SendCmd("SYSTem:DATE?") >>>>>>>>>>>>>>> SKIP
            'Dateval = GetData() >>>>>>>>>>>>>>> SKIP

            ' Get the time at which the scan was started
            'SendCmd("SYSTem:TIME?") >>>>>>>>>>>>>>> SKIP
            'Timeval = GetData() >>>>>>>>>>>>>>> SKIP

            ' Wait until instrument is finished taken readings. The instrument is queried until
            ' all channels are measured.
            Do
                SendCmd("DATA:POINTS?")
                ReturnedData = GetData()
                I = Val(ReturnedData)
            Loop Until I = NumRdgs

            List1.Items.Clear()
            List1.Items.Add("Enter/select instrument address, if needed;")
            List1.Items.Add("Click on " + Chr(34) + "Get Readings" + Chr(34) + _
            " to trigger instrument.")
            List1.Items.Add("Measurements will take some time.")

            List1.Items.Add("")
            'List1.Items.Add("Start Date (yyyy,mm,dd): " + Dateval)
            'List1.Items.Add("Start Time (hh,mm,ss): " + Timeval)
            List1.Items.Add("Rdng#" + Chr(9) + "Channel" + Chr(9) + "Value" + _
            Chr(9) + Chr(9) + "Time")
            List1.Refresh()

            ' Check for errors
            Call Check_Error("Readings")

            ' Take readings out of memory one reading at a time. The "FETCh?" can also be used.
            ' It reads all readings in memory, but leaves the readings in memory. The
            ' "DATA:REMove?" command removes and erases the readings in memory.
        Catch ex As Exception
            MsgBox(ex.ToString)
            If connected Then
                End_Prog()
            End If
            End
        End Try

    End Sub
    Public Sub ReadingsTemp()
        '""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
        ' This sub triggers the instrument and takes readings.
        '""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

        Dim rdgs As String ' Returns a reading from memory
        Dim readMsg As String ' Stores data to show in ListBox
        Dim Dateval As String ' Returns the date value
        Dim Timeval As String ' Returns the time value
        Dim reading As String  ' Used to store reading
        Dim timestamp As String ' Used to store time stamp
        Dim channelnum As String ' Used to store channel number
        Dim I As Long ' Used for a loop

        ' Check for exceptions
        'Try
        ' Trigger instrument
        SendCmd("INITiate") ' >>>>>>>>>>>>>>> SKIP

        ' Get the date at which the scan was started
        'SendCmd("SYSTem:DATE?") >>>>>>>>>>>>>>> SKIP
        'Dateval = GetData() >>>>>>>>>>>>>>> SKIP

        ' Get the time at which the scan was started
        'SendCmd("SYSTem:TIME?") >>>>>>>>>>>>>>> SKIP
        'Timeval = GetData() >>>>>>>>>>>>>>> SKIP

        ' Wait until instrument is finished taken readings. The instrument is queried until
        ' all channels are measured.
        Do
            SendCmd("DATA:POINTS?")
            ReturnedData = GetData()
            I = Val(ReturnedData)
        Loop Until I = NumRdgs

        List1.Items.Clear()
        List1.Items.Add("Enter/select instrument address, if needed;")
        List1.Items.Add("Click on " + Chr(34) + "Get Readings" + Chr(34) + _
        " to trigger instrument.")
        List1.Items.Add("Measurements will take some time.")

        List1.Items.Add("")
        'List1.Items.Add("Start Date (yyyy,mm,dd): " + Dateval)
        'List1.Items.Add("Start Time (hh,mm,ss): " + Timeval)
        List1.Items.Add("Rdng#" + Chr(9) + "Channel" + Chr(9) + "Value" + _
        Chr(9) + Chr(9) + "Time")
        List1.Refresh()

        ' Check for errors
        Call Check_Error("Readings")

        ' Take readings out of memory one reading at a time. The "FETCh?" can also be used.
        ' It reads all readings in memory, but leaves the readings in memory. The
        ' "DATA:REMove?" command removes and erases the readings in memory.
        For I = 1 To 1

            ' Get reading value one at a time
            SendCmd("DATA:REMove? 1")
            rdgs = GetData()

            ' Get reading
            reading = Mid(rdgs, 1, InStr(rdgs, ",") - 1)
            rdgs = Mid(rdgs, InStr(rdgs, ",") + 1, Len(rdgs))

            ' Get time stamp and remove leading zeros
            timestamp = Mid(rdgs, 1, InStr(rdgs, ",") - 1)
            rdgs = Mid(rdgs, InStr(rdgs, ",") + 1, Len(rdgs))

            ' Get channel number
            channelnum = rdgs

            List1.Items.Add(Str(I) + Chr(9) + channelnum + Chr(9) + reading + _
            Chr(9) + timestamp)
            valTemperature(I) = Mid(reading, 1, 6) * 100

        Next I

        'Catch ex As Exception
        '    MsgBox(ex.ToString)
        '    If connected Then
        '        End_Prog()
        '    End If
        '    End
        'End Try

    End Sub

    Function readCurrent() As Double

        SendCmd("SENS:FUNC 'CURR:DC'")
        SendCmd(":READ?")
        Return GetData()

    End Function

    '"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    Function OpenPort() As Boolean
        '"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
        ' This function opens a port (the communication between the instrument and
        ' computer).
        '"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

        ' Check for exceptions
        Try
            ' If port is open, close it
            If connected Then
                errorStatus = visa32.viClose(vi)
            End If

            ' Open the Visa session
            errorStatus = visa32.viOpenDefaultRM(videfaultRM)

            'USB0::2391::8199::MY57010982::0
            'TCPIP0::169.254.9.72::inst0
            ' Open communication to the instrument
            'errorStatus = visa32.viOpen(videfaultRM, "GPIB1::9::INSTR", 0, 0, vi)
            errorStatus = visa32.viOpen(videfaultRM, UCase$(ioType.Text), 0, 0, vi)

            ' If an error occurs, give a message
            If errorStatus < VI_SUCCESS Then
                'MsgBox("Unable to Open instrument. Likely cause is bad address. Correct and press 'Select I/O' button", MsgBoxStyle.Exclamation, "I/O Open Error")
                ioType.Text = "GPIB0::9"
                ioType.Refresh()
                connected = False
                OpenPort = False
                Exit Function
            End If

            ' Set timeout in milliseconds; set the timeout for your requirements
            'errorStatus = viSetAttribute(vi, VI_ATTR_TMO_VALUE, 4000)

            'If InStr(UCase$(ioType.Text), "ASRL") Then

            '    ' Set the RS-232 parameters; refer to the 34970A and VISA documentation
            '    ' to change the settings. Make sure the instrument and the following
            '    ' settings agree.
            '    errorStatus = viSetAttribute(vi, VI_ATTR_ASRL_BAUD, 115200)
            '    errorStatus = viSetAttribute(vi, VI_ATTR_ASRL_DATA_BITS, 8)
            '    errorStatus = viSetAttribute(vi, VI_ATTR_ASRL_PARITY, VI_ASRL_PAR_NONE)
            '    errorStatus = viSetAttribute(vi, VI_ATTR_ASRL_STOP_BITS, VI_ASRL_STOP_ONE)
            '    errorStatus = viSetAttribute(vi, VI_ATTR_ASRL_FLOW_CNTRL, VI_ASRL_FLOW_XON_XOFF)

            '    ' Set the instrument to remote
            '    SendCmd("SYSTem:REMote")
            'End If

            ' Check and make sure the correct instrument is addressed
            Return True
            Exit Function
            List1.Items.Clear()

            If (InStr(ReturnedData, "34970A") = 0 And InStr(ReturnedData, "34972A") = 0) Then
                MsgBox("Incorrect instrument addressed; use the correct address.")
                ioType.Text = "GPIB0::9"
                ioType.Refresh()
                connected = False
                OpenPort = False
                Exit Function
            End If

            If (InStr(ReturnedData, "34972A") > 0) Then
                'Enable logging to USB
                SendCmd("MMEM:LOG ON")
                'Make sure data separator is COMMA
                SendCmd("MMEM:FORM:READ:CSEP COMMA")
                List1.Items.Add("34972 detected.  Enabling logging to USB memory.  Readings")
                List1.Items.Add("will go to internal memory and USB memory (if attached).")
            Else
                List1.Items.Add("Instrument ID is:")
                List1.Items.Add(ReturnedData)
            End If

            ' Check and make sure the 34901A Module is installed in slot 100;
            ' Exit program if not correct
            SendCmd("SYSTem:CTYPe? 100")
            ReturnedData = GetData()

            If InStr(ReturnedData, "34908A") = 0 Then
                MsgBox("Incorrect Module Installed in slot 100!")
                End_Prog()
            End If

            ' Check if the DMM is installed
            SendCmd("INSTrument:DMM:INSTalled?")
            ReturnedData = GetData()
            'If not installed, stop programming the 34970A
            If Val(ReturnedData) = 0 Then
                MsgBox("DMM not installed; unable to make measurements.")
                End_Prog()
            End If

            ' Check if the DMM is enabled; enable if not enabled
            SendCmd("INSTrument:DMM?")
            ReturnedData = GetData()
            If Val(ReturnedData) = 0 Then
                SendCmd("INSTrument:DMM ON")
            End If

            connected = True
            OpenPort = True

        Catch ex As Exception
            MsgBox(ex.ToString)
            If connected Then
                End_Prog()
            End If
            End
        End Try

    End Function

    '""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    Private Sub SendCmd(ByVal SCPICmd As String)
        '""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
        ' This routine will send a SCPI command string to the instrument. If the
        ' command contains a question mark (i.e., is a query command), you must
        ' read the response with the 'GetData' function.
        '""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

        ' Check for exceptions
        Try
            ' Write the command to the instrument (terminated by a linefeed; vbLf is ASCII character 10)
            errorStatus = visa32.viPrintf(vi, SCPICmd & vbLf)

            ' Check for error
            If (errorStatus < visa32.VI_SUCCESS) Then
                MsgBox("I/O Error")

                ' Close the device session
                errorStatus = visa32.viClose(vi)

                End
            End If

        Catch ex As Exception
            MsgBox(ex.ToString)
            If connected Then
                End_Prog()
            End If
            End
        End Try

    End Sub

    '""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    Function GetData() As String
        '""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
        ' This function reads the string returned by the instrument
        '""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

        Dim rdgs As System.Text.StringBuilder = New System.Text.StringBuilder(2048)
        Dim StrVal As String

        ' Check for exceptions
        Try
            ' Return the reading
            errorStatus = visa32.viScanf(vi, "%2048t", rdgs)

            ' Check for error
            If (errorStatus < visa32.VI_SUCCESS) Then
                MsgBox("I/O Error")

                ' Close the device session
                errorStatus = visa32.viClose(vi)

                End
            End If

            ' Store reading in StrVal
            StrVal = rdgs.ToString

            ' Strip out the line feed (vbLf is ASCII character 10) 
            If InStr(StrVal, vbLf) Then
                StrVal = StrVal.Remove((InStr(StrVal, vbLf) - 1), (Len(StrVal) - InStr(StrVal, vbLf) + 1))
            End If

            ' Strip out the carriage return (vbCr is ASCII character 13) 
            If InStr(StrVal, vbCr) Then
                StrVal = StrVal.Remove((InStr(StrVal, vbCr) - 1), (Len(StrVal) - InStr(StrVal, vbCr) + 1))
            End If

            ' Return the data
            GetData = StrVal

        Catch ex As Exception
            MsgBox(ex.ToString)
            If connected Then
                End_Prog()
            End If
            End
        End Try

    End Function

    '""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    Sub Check_Error(ByVal msg As String)
        '""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
        ' Checks for syntax and other errors.
        '""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

        Dim err_code As Integer
        Dim err_msg As String
        Dim TempCheck As Integer
        Dim valpos As Integer

        ' Check for exceptions
        Try
            ' check for initial error
            SendCmd("SYSTem:ERRor?")
            err_msg = GetData()

            ' If error found, check for more errors and exit program
            err_code = Val(err_msg)

            TempCheck = 0
            While err_code <> 0
                TempCheck = 1

                msg = "Error in: " + msg + Chr(10)
                msg = msg + "Error Number: " + Str$(err_code) + Chr(10) + "Error Message: " + err_msg

                MsgBox(msg)

                ' check for more errors
                SendCmd("SYSTem:ERRor?")
                err_msg = GetData()
                err_code = Val(err_msg)
            End While

            If TempCheck <> 0 Then

                ' Send a device clear
                SendCmd("*CLS")

                ' Close instrument session
                errorStatus = viClose(vi)

                ' Close the session
                errorStatus = viClose(videfaultRM)

                ' end the program
                End
            End If

        Catch ex As Exception
            MsgBox(ex.ToString)
            If connected Then
                End_Prog()
            End If
            End
        End Try

    End Sub

    '""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    Sub End_Prog()
        '""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
        ' Closes the session and ends the program.
        '""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
        If connected Then

            ' Abort a scan
            SendCmd("ABORt")

            ' Send a device clear
            SendCmd("*CLS")

            ' Close instrument session
            errorStatus = viClose(vi)

            ' Close the session
            errorStatus = viClose(videfaultRM)
        End If

        End

    End Sub

    '""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    Private Sub SelectIO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SelectIO.Click
        '"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
        ' Button that selects the I/O and creates an instrument session.
        '"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

        OpenPort()

    End Sub

    '""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    Private Overloads Sub GetReadings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GetReadings.Click
        '""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
        ' Call sub routine to trigger instrument and get readings.
        '""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""


    End Sub

    '""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    Private Sub EndProg_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EndProg.Click
        '""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
        ' Calls sub to close the session and end the program.
        '""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

        ' Call sub
        End_Prog()

    End Sub

    '""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    Private Sub ioType_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ioType.TextChanged

    End Sub

    '""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""


    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        '        ReadingsTemp()

    End Sub
End Class
