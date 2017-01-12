Attribute VB_Name = "Module3"
Public Const COMMAND_ERROR = -1
Public Const RECV_ERROR = -1
Public Const NO_ERROR = 0
Dim x, count, check As Integer
Dim Version As Integer


Public socketId As Long

'Global Variables for WINSOCK
Global state As Integer

Sub CloseConnection()


    x = closesocket(socketId)
    
    If x = SOCKET_ERROR Then
        MsgBox ("ERROR: closesocket = ")
        Exit Sub
    End If

End Sub



'Code for this module was obtained from agilent excel lan control document
'link :http://cp.literature.agilent.com/litweb/pdf/16000-95012.pdf
'The code was modified to be compatiable with a 64 bit OS and minor changes were made as required
'The code in the module is used to call Windows functions to communicate with the an equipment over lan

Sub EndIt()

    'Shutdown Winsock DLL
    x = WSACleanup()

End Sub

Sub StartIt()

    Dim StartUpInfo As WSAData
    
    'Version 1.1 (1*256 + 1) = 257
    'version 2.0 (2*256 + 0) = 512
    
    'Get WinSock version
    
   ' Version = ActiveCell.FormulaR1C1
    Version = 257
    'Initialize Winsock DLL
    x = WSAStartup(Version, StartUpInfo)

End Sub
 
Function OpenSocket(ByVal Hostname As String, ByVal PortNumber As Integer) As Integer
   
    Dim I_SocketAddress As sockaddr_in
    Dim ipAddress As Long
    
    ipAddress = inet_addr(Hostname)

    'Create a new socket
    socketId = socket(AF_INET, SOCK_STREAM, 0)
    If socketId = SOCKET_ERROR Then
        MsgBox ("ERROR: socket = " + CStr(socketId))
        OpenSocket = COMMAND_ERROR
        Exit Function
    End If

    'Open a connection to a server

    I_SocketAddress.sin_family = AF_INET
    I_SocketAddress.sin_port = htons(PortNumber)
    I_SocketAddress.sin_addr = ipAddress
    I_SocketAddress.sin_zero = "00000000"

    x = connect(socketId, I_SocketAddress, Len(I_SocketAddress))
    If socketId = SOCKET_ERROR Then
        MsgBox ("ERROR: connect = " + CStr(x))
        OpenSocket = COMMAND_ERROR
        Exit Function
    End If

    OpenSocket = socketId

End Function

Function SendCommand(ByVal command As String) As Integer

    Dim strSend As String
    
    strSend = command + vbCrLf
    
    count = send(socketId, ByVal strSend, Len(strSend), 0)
    
    If count = SOCKET_ERROR Then
        MsgBox ("ERROR: send = " + CStr(count))
        SendCommand = COMMAND_ERROR
        Exit Function
    End If
    
    SendCommand = NO_ERROR

End Function

Function RecvAscii(databuf As String, ByVal maxLength As Integer) As Integer

    Dim c As String * 1
    Dim length As Integer
    
    check = setsockopt(socketId, SOL_SOCKET, SO_RCVTIMEO, 10000, 4)
    If (check = SOCKET_ERROR) Then
      MsgBox "Error setting SO_RCVTimeo option: "
     Exit Function
    End If
    
    
    databuf = ""
    While length < maxLength
        DoEvents
        count = recv(socketId, c, 1, 0)
        If count < 1 Then
            RecvAscii = RECV_ERROR
            databuf = vbNullChar
            Exit Function
        End If
        If c = vbCr Then  '& length > 1
        databuf = databuf + c
        Exit Function
        End If
        
        If c = vbCr Then
           databuf = databuf + vbNullChar
           RecvAscii = NO_ERROR
           Exit Function
        End If
        
        length = length + count
        databuf = databuf + c
    Wend
    
    RecvAscii = RECV_ERROR
    
End Function

Function RecvAryReal(databuf() As Double) As Long
    ' receive DOS format 64bit binary data

    Dim buf As String * 20
    Dim size As Long
    Dim length As Long
    Dim count As Long
    Dim recvBuf(25616) As Byte
    
    ' receive header info "#6NNNNNN"
    x = recv(socketId, buf, 8, 0)
    
    size = Val(Mid$(buf, 3, 6))
    
    count = 0
    length = 0
    Do While length < size
        DoEvents
        count = recvB(socketId, recvBuf(length), size - length, 0)
        If (count > 0) Then
            length = length + count
        End If
    Loop
    
    ' receive ending LF
    count = recv(socketId, buf, 1, 0)
    
    ' copy recieved data to Single type array dataBuf()
    CopyMemory databuf(LBound(databuf)), recvBuf(0), length
    'dataBuf = recvBuf
    
    RecvAryReal = length / 8
    
End Function

