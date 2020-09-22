Attribute VB_Name = "Module1"
Option Explicit

Public Const IP_SUCCESS As Long = 0
Public Const MAX_WSADescription As Long = 256
Public Const MAX_WSASYSStatus As Long = 128
Public Const WS_VERSION_REQD As Long = &H101
Public Const WS_VERSION_MAJOR As Long = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR As Long = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD As Long = 1
Public Const SOCKET_ERROR As Long = -1

Public Type WSADATA
   wVersion As Integer
   wHighVersion As Integer
   szDescription(0 To MAX_WSADescription) As Byte
   szSystemStatus(0 To MAX_WSASYSStatus) As Byte
   wMaxSockets As Long
   wMaxUDPDG As Long
   dwVendorInfo As Long
End Type

Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal hostname As String) As Long
  
Public Declare Function WSAStartup Lib "WSOCK32.DLL" _
   (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
    
Public Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
   (xDest As Any, xSource As Any, ByVal nbytes As Long)

Private Declare Function lstrlenA Lib "kernel32" (lpString As Any) As Long


Public Function SocketsInitialize() As Boolean
   Dim WSAD As WSADATA
   Dim success As Long
   SocketsInitialize = WSAStartup(WS_VERSION_REQD, WSAD) = IP_SUCCESS
End Function

Public Sub SocketsCleanup()
   If WSACleanup() <> 0 Then
       MsgBox "Windows Sockets error occurred in Cleanup.", vbExclamation
   End If
End Sub

Public Function GetIPFromHostName(ByVal sHostName As String) As String
   Dim nbytes As Long
   Dim ptrHosent As Long
   Dim ptrName As Long
   Dim ptrAddress As Long
   Dim ptrIPAddress As Long
   Dim sAddress As String
   
   sAddress = Space$(4)
   ptrHosent = gethostbyname(sHostName & vbNullChar)
   If ptrHosent <> 0 Then
      ptrAddress = ptrHosent + 12
      CopyMemory ptrAddress, ByVal ptrAddress, 4
      CopyMemory ptrIPAddress, ByVal ptrAddress, 4
      CopyMemory ByVal sAddress, ByVal ptrIPAddress, 4
      GetIPFromHostName = IPToText(sAddress)
   End If
End Function

Private Function IPToText(ByVal IPAddress As String) As String
   IPToText = CStr(Asc(IPAddress)) & "." & _
              CStr(Asc(Mid$(IPAddress, 2, 1))) & "." & _
              CStr(Asc(Mid$(IPAddress, 3, 1))) & "." & _
              CStr(Asc(Mid$(IPAddress, 4, 1)))
End Function


