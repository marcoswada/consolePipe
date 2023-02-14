Imports System.IO.Ports
Imports System.Text
class clsTest
    private pComPort As String
    const timeOut = 1000
    const progName = "readdyp"

    Public Sub New()
        autodetect
    End Sub

    Property comPort As String
        Get
            comPort = pcomPort
        End Get
        Set (cPort As String)
            pComPort = cPort
        End Set
    End Property

    Private Sub autodetect()
        dim readValue = my.computer.registry.getValue("HKEY_CURRENT_USER\SOFTWARE\" & progName, "ComPort", Nothing)
        if readValue IsNot Nothing Then
            pComPort = readValue
            if testPort() Then
                Exit Sub
            End If
        End If
        For Each sp as String In My.Computer.Ports.SerialPortNames
            pcomPort=sp
            if testPort() Then
                my.computer.registry.currentuser.CreateSubKey("SOFTWARE\" & progName)
                my.computer.registry.setValue("HKEY_CURRENT_USER\SOFTWARE\" & progName, "ComPort", pComPort)
                Exit Sub
            End If
        Next
        pComPort=""
    End Sub

    Private  Function sendCommand(xcmd as string) As String
        Dim comPort As IO.Ports.SerialPort = Nothing
        Dim ret As String = ""
        If xcmd="" then xcmd="Q"
        if pComPort="" Then
            MsgBox ("???")
            sendCommand ="-1"
            Exit Function
        End If
        try
            comPort = My.Computer.Ports.OpenSerialPort(pComPort)
            comPort.WriteTimeOut = timeOut
            comPort.baudRate = 38400
            comPort.parity = Parity.None
            comPort.dataBits = 8
            comPort.stopBits = StopBits.One
            comPort.Handshake = Handshake.None
            comPort.rtsEnable = True
            comPort.dtrEnable = True
            comPort.Write(xcmd & vbCrLf)
            comPort.ReadTimeout = timeOut
            ret=comPort.ReadLine()
        Catch ex As System.Exception
            ret="-1"
        Finally
            If pComPort IsNot Nothing Then comPort.Close()
        End Try
        sendCommand = ret
    End Function

    Readonly Property Distance As Single
        Get
            Dim s As String
            If pComPort <> "" Then
                Dim t As Integer
                s = sendCommand("Q")
                t = 0
                if Val(s)<0 Then
                    s="999.99 cm"
                End If
                Distance = Val (Left(s, Len(s)-3))
            Else
                Distance = 999.99
            End If
        End Get
    End Property
    Readonly Property RemainingTime As Long
        Get
            Dim s As String
            If pComPort <>"" Then
                s=sendCommand("T")
                RemainingTime = Val(Left(s, len(s)-3))
            Else
                RemainingTime = "COM Port not open"
            End If
        End Get
    End Property

    Public Function setTimer(t As Long) As String
        Dim s As String
        If (pcomPort <> "") Then 
            s=sendCommand("S" & trim(str(t)))
            setTimer = s        
        Else
            setTimer = "COM Port not open"
        End If
    End Function

    Private  Function testPort() As Boolean
        Dim ret As String = ""
        Dim retVal as Boolean = False
        If pComPort<>"" Then
            ret=sendCommand("Q")
            if UCase(Left(Right(ret,3),2))="CM" Then
                retVal=True
            End If
        End if  
        testPort = retVal 
    End Function
End Class   

Module Program

    Sub Main()
        dim d as new clsTest
        dim dArgs() as String
        dArgs=Environment.GetCommandLineArgs()
        dim s As String = "Q"
        if Ubound(dArgs) > 0 Then
            s=UCase(dArgs(1))
        End If
        If (s="Q") Then
            console.writeLine(d.Distance)
        ElseIf (s="T") Then
            console.writeLine(d.RemainingTime)
        ElseIf (val(s)>0) then
            console.writeLine(d.setTimer(val(s)))
        End If
    End Sub
End Module