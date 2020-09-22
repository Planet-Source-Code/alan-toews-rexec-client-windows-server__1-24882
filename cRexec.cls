VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Rexec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents wsOut As Winsock
Attribute wsOut.VB_VarHelpID = -1
Private WithEvents wsIn As Winsock
Attribute wsIn.VB_VarHelpID = -1
Private bConnected As Boolean
Private bError As Boolean

Public Event Response(Description As String)
Public Event SocketResponse(Description As String)
Public Event Error(Description As String)


Public UserID As String
Public Password As String
Public Address As String
Public Port As Long

Public Property Get Connected() As Boolean
    Connected = bConnected
End Property
Public Sub Execute(cmd As String, _
                  Optional sServer As String = "", _
                  Optional lPort As Long = 0, _
                  Optional sUid As String = "", _
                  Optional sPwd As String = "")
    
    'Close the sockets if they're open
    CloseSock wsIn
    CloseSock wsOut
    
    'set our flags to their start values
    bConnected = False
    bError = False
    
    'check if the defaults are being
    'overridden by the sub call
    'if they aren't then use the default values
    If sServer = "" Then sServer = Address
    If lPort = 0 Then lPort = Port
    If sUid = "" Then sUid = UserID
    If sPwd = "" Then sPwd = Password
    
    'set up our port information
    wsIn.LocalPort = 0
    wsOut.RemotePort = lPort
    
    'make sure our protocols are set correctly
    wsIn.Protocol = sckTCPProtocol
    wsOut.Protocol = sckTCPProtocol
    
    'begin listening on our inbound port
    wsIn.Listen
    
    'connect to the rexec server
    wsOut.RemoteHost = sServer
    wsOut.Connect
    
    'wait for the socket to connect before sending data
    Do While Not bConnected And Not bError
        DoEvents
    Loop
    
    'if there is no error, then we are ready to go.
    If Not bError Then
        wsOut.SendData CStr(wsIn.LocalPort) & vbNullChar
        wsOut.SendData Trim(sUid) & vbNullChar
        wsOut.SendData Trim(sPwd) & vbNullChar
        wsOut.SendData Trim(cmd) & vbNullChar
    End If
End Sub

Private Sub CloseSock(ByRef ws As Winsock)
    'make sure that the requested socket closes
    Do While ws.State <> sckClosed
        ws.Close
        DoEvents
    Loop
End Sub
Private Sub Class_Initialize()
    'establish our default values
    Address = "127.0.0.1"
    Port = 512
    UserID = "user"
    Password = "password"
        
    'initialize out winsock controls
    Set wsIn = New Winsock
    Set wsOut = New Winsock
    
    
End Sub

Private Sub wsIn_Close()
    'Mark socket closure
    RaiseEvent SocketResponse("Incoming Socket Closed.")
End Sub


Private Sub wsIn_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    RaiseEvent Error(Description)
End Sub

Private Sub wsOut_Close()
    'Mark socket closure
    bConnected = False
    RaiseEvent SocketResponse("Send Socket Closed.")
End Sub


Private Sub wsIn_ConnectionRequest(ByVal requestID As Long)
    'Rexecd tries to connect here to send stderr
    RaiseEvent SocketResponse("Incoming Connection Request on Send Socket.")
    CloseSock wsIn
    wsIn.Accept requestID
End Sub


Private Sub wsIn_DataArrival(ByVal bytesTotal As Long)
    'If we receive inbound data on the outbound port
    'then the server is sending an error message
    Dim DataRcvd As String
    wsIn.GetData DataRcvd
    
    'check if the data is valid.
    If bytesTotal < 2 Or wsIn.State <> sckConnected Then Exit Sub
    
    'raise socket event
    RaiseEvent SocketResponse("Incoming Response on Send Socket.")
    
    'fix the linefeeds. turn the unix vbLf into the windows vbCrLf
    DataRcvd = Replace(Replace(DataRcvd, vbLf, vbCrLf), vbCr & vbCrLf, vbCrLf)
    If Trim(DataRcvd) <> "" Then RaiseEvent Error(DataRcvd)
End Sub

Private Sub wsOut_Connect()
    'The main loop is waiting for this event
    RaiseEvent SocketResponse("Receive Socket Connected.")
    bConnected = True
End Sub

Private Sub wsOut_DataArrival(ByVal bytesTotal As Long)
    'Receive output
    Dim DataRcvd As String
    wsOut.GetData DataRcvd
    If CBool(InStr(1, DataRcvd, vbLf)) And Not CBool(InStr(1, DataRcvd, vbCr)) Then _
        DataRcvd = Replace(Replace(Replace(DataRcvd, vbCr, vbLf), vbLf & vbLf, vbLf), vbLf, vbCrLf)
    RaiseEvent Response(DataRcvd)
End Sub

Private Sub wsOut_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    bError = True
    RaiseEvent Error(Description)
End Sub
