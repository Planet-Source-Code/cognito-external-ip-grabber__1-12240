VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "External IP Grabber"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   5625
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdVerbose 
      Caption         =   "Verbose"
      Height          =   375
      Left            =   3960
      TabIndex        =   8
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtMsg 
      Height          =   2775
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   2040
      Width           =   4815
   End
   Begin VB.TextBox txtServ 
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   360
      Width           =   2535
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   4560
      TabIndex        =   3
      Top             =   360
      Width           =   615
   End
   Begin VB.CommandButton cmdConn 
      Caption         =   "Grab It"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   4815
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Top             =   1440
      Width           =   1455
   End
   Begin MSWinsockLib.Winsock ws 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblServ 
      Caption         =   "IRC server:"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   360
      Width           =   855
   End
   Begin VB.Label lblPort 
      Caption         =   "port:"
      Height          =   255
      Left            =   4080
      TabIndex        =   4
      Top             =   360
      Width           =   375
   End
   Begin VB.Label lblIP 
      Caption         =   "Your external ip address:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Serv As String
Dim Port As Integer
Dim ServMsg As String
Dim Verbose As Boolean

Const Nick As String = "RebBarGPI"      'Meaningless NICK ("IPGraBbeR" backwards)
Const User As String = "guest"          'whatever...
Const RealName As String = "John Doe"
Const ServInit As String = "glass.dal.net"
Const PortInit As String = "6667"
Const HeightSmall As Integer = 2400     'Normal height of form
Const HeightLarge As Integer = 5550     'Height for verbose option

Private Sub cmdConn_Click()
   Serv = txtServ.Text
   Port = CInt(txtPort.Text)
   ServMsg = ""
   ws.Close
   ws.Connect Serv, Port
End Sub

Private Sub cmdVerbose_Click()
   Verbose = Not Verbose
   If Verbose Then
      Me.Height = HeightLarge
   Else
      Me.Height = HeightSmall
   End If
End Sub

Private Sub Form_Load()
   'Initialize some defaults
   txtServ.Text = ServInit
   txtPort.Text = PortInit
   Verbose = False
   Me.Height = HeightSmall
End Sub

Private Sub ws_Connect()
   'The necessary IRC protocol
   ws.SendData "NICK " & Nick & vbCrLf
   ws.SendData "USER " & User & " " & ws.LocalIP & " " & ws.RemoteHostIP & " :" & RealName & vbCrLf
End Sub

Private Sub ws_DataArrival(ByVal bytesTotal As Long)
   Dim Buf As String
   Dim Pos001, PosAt, PosCr As Integer
   Dim Host As String
   Dim IP As String
   
   ws.GetData Buf, vbString
   ServMsg = ServMsg & Buf
   txtMsg.Text = ServMsg
   
   'Parse ServMsg for hostname
   'See RFC 2812 for more info
   Pos001 = InStr(ServMsg, "001")       'expecting "001" welcome message
   If Pos001 = 0 Then Exit Sub
   PosAt = InStr(Pos001, ServMsg, "@")  'hostname follows "@"
   If PosAt = 0 Then Exit Sub
   PosCr = InStr(PosAt, ServMsg, vbCr)  'hostname is just before EOL
   If PosCr = 0 Then Exit Sub
   Host = Mid(ServMsg, PosAt + 1, PosCr - 1 - PosAt)
   
   'Now lookup ip from hostname
   If SocketsInitialize() Then
      IP = GetIPFromHostName(Host)
      SocketsCleanup
   Else
      MsgBox "Windows Sockets for 32 bit Windows " & _
             "environments is not successfully responding."
   End If
   
   txtIP.Text = IP
On Error Resume Next
   ws.SendData "QUIT" & vbCrLf
End Sub
