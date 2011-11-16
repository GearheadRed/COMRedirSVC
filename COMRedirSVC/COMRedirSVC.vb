' Copyright 2011 Branden Coates

' Licensed under the Apache License, Version 2.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at

' http://www.apache.org/licenses/LICENSE-2.0

' Unless required by applicable law or agreed to in writing, software
' distributed under the License is distributed on an "AS IS" BASIS,
' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
' See the License for the specific language governing permissions and
' limitations under the License.

' Contact:    Branden Coates
'             branden.c.coates@gmail.com

Imports System.Timers
Imports System.ServiceProcess
Imports System.Text
Imports System.IO.File

Public Class COMRedirSVC
    'Initialize variables
    Friend WithEvents EventLog1 As System.Diagnostics.EventLog
    Dim WithEvents serialPort As New IO.Ports.SerialPort
    Private ReceiveBuffer As New StringBuilder(32768)
    Public Property PortName As String
    Dim serviceStarted As Boolean
    Dim COMTimer As New System.Timers.Timer()

    'Initialize values and start service
    Protected Overrides Sub OnStart(ByVal args() As String)
        COMTimer.Interval = 100
        COMTimer.Enabled = True
        COMTimer.AutoReset = True

        serialPort.PortName = "COM1"
        serialPort.BaudRate = 9600
        serialPort.DataBits = 8
        serialPort.Parity = IO.Ports.Parity.None
        serialPort.StopBits = IO.Ports.StopBits.One
        serialPort.RtsEnable = True
        serialPort.DtrEnable = True
        'Attempt opening the serial port, if fail log error(most likely com in use).
        Try
            serialPort.Open()
        Catch ex As Exception
            EventLog1.WriteEntry("COMRedirSVC could not bind to COM1.", EventLogEntryType.Error, eventID:=1)
        End Try
        serialPort.Write("Ready." & vbNewLine & vbNewLine)
        AddHandler COMTimer.Elapsed, AddressOf Monitor_Comm
    End Sub

    'Monitor serial port for data in buffer(checks every 100 miliseconds)
    Private Sub Monitor_Comm(ByVal source As Object, ByVal e As ElapsedEventArgs)
        If serialPort.IsOpen = True Then
            If serialPort.BytesToRead > 0 Then SerialPort_OnComm()
        End If
    End Sub

    'If data in buffer from the Monitor_Comm() sub then check the buffer for backspaces, enters, and specialized commands
    Private Sub SerialPort_OnComm()
        ReceiveBuffer.Append(serialPort.ReadExisting)
        serialPort.DiscardInBuffer()
        'Looking for backspace or enter keypress and specialized commands
        If ReceiveBuffer.ToString.EndsWith(Chr(127)) And ReceiveBuffer.Length - 1 > 0 Then
            ReceiveBuffer.Remove(ReceiveBuffer.Length - 2, 2)
        ElseIf ReceiveBuffer.ToString.EndsWith(Chr(27)) Then
            serialPort.Write(vbCr & Space(ReceiveBuffer.Length + 1) & vbCr)
            ReceiveBuffer.Length = 0
        ElseIf ReceiveBuffer.ToString.EndsWith(Chr(13)) Then
            If ReceiveBuffer.ToString.Contains("Shrink C:\") Then
                Shrink()
            ElseIf ReceiveBuffer.ToString.Contains("Extend C:\") Then
                Extend()
            Else
                ExecuteCommand()
            End If
        End If
        'Output current receivebuffer data so user knows what they have entered
        serialPort.Write(vbCr & Space(ReceiveBuffer.Length + 1) & vbCr & ReceiveBuffer.ToString)
    End Sub

    'Execute the command in the receive buffer
    Private Sub ExecuteCommand()
        If ReceiveBuffer.Length > 0.8 * ReceiveBuffer.Capacity Then
            ReceiveBuffer.Remove(0, CInt(0.6 * ReceiveBuffer.Capacity))
        End If
        serialPort.Write(vbNewLine & "Executing: " & ReceiveBuffer.ToString & vbNewLine)
        Dim execProcess As New Process()
        execProcess.StartInfo.UseShellExecute = False
        execProcess.StartInfo.RedirectStandardInput = True
        execProcess.StartInfo.RedirectStandardOutput = True
        execProcess.StartInfo.RedirectStandardError = True
        execProcess.StartInfo.CreateNoWindow = True
        execProcess.StartInfo.FileName = "cmd.exe"
        execProcess.StartInfo.Arguments = "/s /c """ & ReceiveBuffer.ToString & """"
        Try
            execProcess.Start()
        Catch ex As Exception
            serialPort.Write("Execution Error: (" & Err.Number & ") " & Err.Source & " - " & Err.Description & vbNewLine)
        End Try
        'Catch and return executed command results
        Do Until execProcess.StandardOutput.EndOfStream
            serialPort.Write(execProcess.StandardOutput.ReadLine & vbNewLine)
        Loop
        Do Until execProcess.StandardError.EndOfStream
            serialPort.Write("Command Error: " & execProcess.StandardError.ReadLine & vbNewLine)
        Loop
        serialPort.Write("Execution Completed." & vbNewLine & vbNewLine)
        ReceiveBuffer.Length = 0
    End Sub

    'Specialized process to facilitate shrinking the C:\ partition
    Private Sub Shrink()
        If ReceiveBuffer.Length > 0.8 * ReceiveBuffer.Capacity Then
            ReceiveBuffer.Remove(0, CInt(0.6 * ReceiveBuffer.Capacity))
        End If
        'Delete the temp script if existing and create new shrink diskpart script
        Delete("C:\Windows\Temp\COMRedirSVCShrink.txt")
        Dim ShrinkWrite As System.IO.StreamWriter
        ShrinkWrite = IO.File.CreateText("C:\Windows\Temp\COMRedirSVCShrink.txt")
        ShrinkWrite.WriteLine("select disk 0")
        ShrinkWrite.WriteLine("select partition 2")
        ShrinkWrite.WriteLine("shrink")
        ShrinkWrite.WriteLine("detail disk")
        ShrinkWrite.WriteLine("Exit")
        ShrinkWrite.Close()
        'Execute diskpart using the temporary script and shrink the partition
        serialPort.Write(vbNewLine & "Shrinking C:\" & vbNewLine)
        Dim ShrinkProcess As New Process()
        ShrinkProcess.StartInfo.UseShellExecute = False
        ShrinkProcess.StartInfo.RedirectStandardInput = True
        ShrinkProcess.StartInfo.RedirectStandardOutput = True
        ShrinkProcess.StartInfo.RedirectStandardError = True
        ShrinkProcess.StartInfo.CreateNoWindow = True
        ShrinkProcess.StartInfo.FileName = "diskpart.exe"
        ShrinkProcess.StartInfo.Arguments = "/s C:\Windows\Temp\COMRedirSVCShrink.txt"
        Try
            ShrinkProcess.Start()
        Catch ex As Exception
            serialPort.Write("Execution Error: (" & Err.Number & ") " & Err.Source & " - " & Err.Description & vbNewLine)
        End Try
        'Catch and return diskpart results
        Do Until ShrinkProcess.StandardOutput.EndOfStream
            serialPort.Write(ShrinkProcess.StandardOutput.ReadLine & vbNewLine)
        Loop
        Do Until ShrinkProcess.StandardError.EndOfStream
            serialPort.Write("Shrink Error: " & ShrinkProcess.StandardError.ReadLine & vbNewLine)
        Loop
        serialPort.Write("Disk Shrink Completed." & vbNewLine & vbNewLine)
        ReceiveBuffer.Length = 0
        'Delete the temp script
        Delete("C:\Windows\Temp\COMRedirSVCShrink.txt")
    End Sub

    Private Sub Extend()
        If ReceiveBuffer.Length > 0.8 * ReceiveBuffer.Capacity Then
            ReceiveBuffer.Remove(0, CInt(0.6 * ReceiveBuffer.Capacity))
        End If
        'delete the temp script if existing and create new extend diskpart script
        Delete("C:\Windows\Temp\COMRedirSVCExtend.txt")
        Dim ExtendWrite As System.IO.StreamWriter
        ExtendWrite = IO.File.CreateText("C:\Windows\Temp\COMRedirSVCExtend.txt")
        ExtendWrite.WriteLine("select disk 0")
        ExtendWrite.WriteLine("select partition 2")
        ExtendWrite.WriteLine("extend")
        ExtendWrite.WriteLine("detail disk")
        ExtendWrite.WriteLine("Exit")
        ExtendWrite.Close()
        'Execute diskpart using the temporary script and extend the partition
        serialPort.Write(vbNewLine & "Extending C:\" & vbNewLine)
        Dim ExtendProcess As New Process()
        ExtendProcess.StartInfo.UseShellExecute = False
        ExtendProcess.StartInfo.RedirectStandardInput = True
        ExtendProcess.StartInfo.RedirectStandardOutput = True
        ExtendProcess.StartInfo.RedirectStandardError = True
        ExtendProcess.StartInfo.CreateNoWindow = True
        ExtendProcess.StartInfo.FileName = "diskpart.exe"
        ExtendProcess.StartInfo.Arguments = "/s C:\Windows\Temp\COMRedirSVCExtend.txt"
        Try
            ExtendProcess.Start()
        Catch ex As Exception
            serialPort.Write("Execution Error: (" & Err.Number & ") " & Err.Source & " - " & Err.Description & vbNewLine)
        End Try
        'Catch and return diskpart results
        Do Until ExtendProcess.StandardOutput.EndOfStream
            serialPort.Write(ExtendProcess.StandardOutput.ReadLine & vbNewLine)
        Loop
        Do Until ExtendProcess.StandardError.EndOfStream
            serialPort.Write("Extend Error: " & ExtendProcess.StandardError.ReadLine & vbNewLine)
        Loop
        serialPort.Write("Disk Extend Completed." & vbNewLine & vbNewLine)
        ReceiveBuffer.Length = 0
        'Delete the temp script
        Delete("C:\Windows\Temp\COMRedirSVCExtend.txt")
    End Sub

    'Just error logging to the Windows Application log
    Public Sub New()
        MyBase.New()
        InitializeComponent()
        If Not System.Diagnostics.EventLog.SourceExists("COMRedirSVC") Then
            System.Diagnostics.EventLog.CreateEventSource("COMRedirSVC", "Appication")
        End If
        EventLog1.Source = "COMRedirSVC"
        EventLog1.Log = "Application"
    End Sub

    'Cleanup and stop the service
    Protected Overrides Sub OnStop()
        COMTimer.Enabled = False
        If serialPort.IsOpen = True Then serialPort.Close()
    End Sub

End Class
