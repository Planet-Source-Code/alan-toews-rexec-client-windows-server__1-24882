VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "REXEC Test Form"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   7245
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   255
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   5160
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   255
      Index           =   2
      Left            =   3180
      TabIndex        =   4
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   255
      Index           =   1
      Left            =   2100
      TabIndex        =   3
      Top             =   0
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   2
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "Type your command and press enter to run it."
      Top             =   540
      Width           =   7215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   840
      Width           =   7215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Email Me"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3420
      MouseIcon       =   "Form1.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   300
      Width           =   645
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "View My Other PSC Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   5400
      MouseIcon       =   "Form1.frx":08CA
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   300
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Command Line"
      Height          =   195
      Index           =   4
      Left            =   0
      TabIndex        =   10
      Top             =   240
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   195
      Index           =   3
      Left            =   4440
      TabIndex        =   9
      Top             =   60
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Login"
      Height          =   195
      Index           =   2
      Left            =   2760
      TabIndex        =   8
      Top             =   60
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Port"
      Height          =   195
      Index           =   1
      Left            =   1740
      TabIndex        =   7
      Top             =   60
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Server"
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   60
      Width           =   465
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents rx As Remote.Rexec
Attribute rx.VB_VarHelpID = -1
Dim WithEvents rxS As Remote.RexecServer
Attribute rxS.VB_VarHelpID = -1
Private Sub Form_Load()
    Set rx = New Rexec
    Set rxS = New RexecServer
    
    'get defaults from client
    Text3(0).Text = rx.Address
    Text3(1).Text = rx.Port
    Text3(2).Text = rx.UserID
    Text3(3).Text = rx.Password
    
    'setup server login to match client defaults
    rxS.Login = Text3(2).Text
    rxS.Password = Text3(3).Text
    
    rxS.Enabled = True
End Sub

Private Sub Label2_Click()
    Dim oFs As FileIO
    Set oFs = New FileIO
    oFs.RunProgram "http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?lngWId=1&blnAuthorSearch=TRUE&lngAuthorId=136063&strAuthorName=Alan%20Toews&txtMaxNumberOfEntriesPerPage=25"
    Set oFs = Nothing
End Sub

Private Sub Label3_Click()
Dim oFs As FileIO
    Set oFs = New FileIO
    oFs.RunProgram "mailto:actoews@hotmail.com"
    Set oFs = Nothing
End Sub

Private Sub rx_Error(Description As String)
    'display any error messages
    Text1.Text = Text1.Text & Description & vbCrLf
End Sub

Private Sub rx_Response(Description As String)
    'display the server response
    Text1.Text = Text1.Text & Description & vbCrLf
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        'set server info
        rx.Address = Text3(0).Text
        rx.Port = CInt(Text3(1).Text)
        rx.UserID = Text3(2).Text
        rx.Password = Text3(3).Text
        'clear the response window
        Text1.Text = ""
        'send the command
        rx.Execute Text2.Text
        'clear the keypress
        KeyAscii = 0
        Text2.Text = ""
    End If
End Sub

Private Sub Text3_GotFocus(Index As Integer)
    Text3(Index).SelStart = 0: Text3(Index).SelLength = Len(Text3(Index).Text)
End Sub
