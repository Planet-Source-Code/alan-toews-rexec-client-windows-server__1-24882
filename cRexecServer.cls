VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RexecServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents oIn As Winsock
Attribute oIn.VB_VarHelpID = -1
Private WithEvents oOut As Winsock
Attribute oOut.VB_VarHelpID = -1
Private oFs As FileIO

Private sLogin As String
Private sPassword As String
Private bEnabled As Boolean
Private iState As Integer
Private sServer As String
Private bWaiting As Boolean
Private lPort As Long

Public Event SocketEvent(Description As String)

Public Property Get Port() As Long
    Port = oIn.LocalPort
End Property
Public Property Let Port(NewVal As Long)
    lPort = NewVal
    oIn.LocalPort = lPort
End Property
Public Property Get Login() As String
    Login = sLogin
End Property
Public Property Let Login(NewVal As String)
    sLogin = NewVal
    'save the setting to the registry
    SaveSetting App.Title, "settings\login\", "login", oFs.Compress(oFs.Obscure(sLogin))
End Property
Public Property Get Password() As String
    Password = spasswprd
End Property
Public Property Let Password(NewVal As String)
    sPassword = NewVal
    'save the setting to the registry
    SaveSetting App.Title, "settings\login\", "password", oFs.Compress(oFs.Obscure(sPassword))
End Property
Public Property Get Enabled() As Boolean
    Enabled = bEnabled
End Property
Public Property Let Enabled(NewVal As Boolean)
    bEnabled = NewVal
    If bEnabled Then
        CloseSocket oIn
        oIn.LocalPort = lPort
        oIn.Listen
    Else
        CloseSocket oIn
    End If
End Property
Private Sub Class_Initialize()
    'set things up
    Set oIn = New Winsock
    Set oOut = New Winsock
    Set oFs = New FileIO
    
    'set defaults
    lPort = 512
    oIn.LocalPort = lPort
    sLogin = oFs.Recover(oFs.Decompress(GetSetting(App.Title, "settings\login\", "login", oFs.Compress(oFs.Obscure("user")))))
    sPassword = oFs.Recover(oFs.Decompress(GetSetting(App.Title, "settings\login\", "password", oFs.Compress(oFs.Obscure("password")))))
End Sub

Private Function CloseSocket(ws As Winsock)
    Do
        ws.Close
        DoEvents
    Loop While ws.State <> sckClosed
    CloseSocket = True
End Function

Private Sub oIn_Close()
    RaiseEvent SocketEvent("Server control socket closed.")
    CloseSocket oIn
    oIn.Listen
End Sub

Private Sub oIn_Connect()
    iState = 1
    RaiseEvent SocketEvent("Server control socket connected.")
End Sub

Private Sub oIn_ConnectionRequest(ByVal requestID As Long)
    
    CloseSocket oIn
    
    oIn.Accept requestID
    RaiseEvent SocketEvent("Server control socket connection Request.")
    
End Sub

Private Sub oIn_DataArrival(ByVal bytesTotal As Long)
    Dim InData As String
    Dim InCmds As Variant, x As Integer
    
    If oIn.State = sckConnected Then
        oIn.GetData InData
        
        InCmds = Split(InData, vbNullChar)
        For x = LBound(InCmds) To UBound(InCmds)
            If iState = -1 Then iState = 0
            Interpret Trim(InCmds(x))
        Next
                
    End If
End Sub

Private Sub Interpret(ByVal sInput As String)
    Dim tmpIn As String, path As String
    If oIn.State = sckConnected And iState = 0 Then iState = 1
    Select Case iState
        Case -1         'error occured.
            
        Case 1          'getting return port
            'connect to the remote port.
            CloseSocket oOut
            If IsNumeric(sInput) Then
                oOut.RemoteHost = oIn.RemoteHostIP
                oOut.RemotePort = CStr(sInput)
            
                'move on to the next stage
                If sInput > 0 And sInput < 65536 Then
                    iState = 2
                Else
                    oIn.SendData "Bad Port Number."
                    sInput = -1
                End If
            End If
        Case 2          'getting username
            If sInput = Login Then
                iState = 3
            Else
                oIn.SendData "Login Failed!"
                iState = -1
            End If
        Case 3          'getting password
            If sInput = sPassword Then
                iState = 4
            Else
                oIn.SendData "Password Failed!"
                iState = -1
            End If
        Case 4          'getting command
            iState = 0
            'connect to client
            'CloseSocket oOut
            oOut.Connect
            
            'wait for a connection or an error
            Do While oOut.State <> sckError And oOut.State <> sckConnected
                DoEvents
            Loop
            
            'if we had an error, then exit
            If oOut.State <> sckConnected Then
                Exit Sub
            End If
            
            'run the command received, and wait till it finishes.
            If oFs.RunAndWait("command.com", , _
                              "/C " & sInput & " >c:\cmdTemp.tmp", _
                              oFs.GetSpecialfolder(CSIDL_SYSTEM_32), _
                              START_HIDDEN) Then
                
                'open the output file
                
                tmpIn = oFs.Load("c:\cmdTemp.tmp")
                oFs.Delete "c:\cmdTemp.tmp"
                
                'reply to the client
                If oOut.State = sckConnected Then
                    bWaiting = True
                    oOut.SendData tmpIn & vbNullChar
                End If
                
                'wait for send to complete
                Do While bWaiting
                    DoEvents
                Loop
                
            Else
                'reply to the client
                If oOut.State = sckConnected Then oOut.SendData "Command Failed to execute." & vbNullChar
            End If
            'wait for reply to finish, so the port doesn't close too early
            Do While oOut.State <> sckError And oOut.State <> sckConnected
                DoEvents
            Loop
            
            'close the sockets behind us
            CloseSocket oOut
            CloseSocket oIn
            'prepare for next command
            Class_Initialize
            Enabled = False
            Enabled = True
            
    End Select

End Sub

Private Sub oIn_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    CloseSocket oIn
    oIn.Listen
End Sub

Private Sub oIn_SendComplete()
    RaiseEvent SocketEvent("Server control socket send complete.")
End Sub

Private Sub oOut_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    If oIn.State = sckConnected Then oIn.SendData Description
End Sub

Private Sub oOut_SendComplete()
    bWaiting = False
End Sub


