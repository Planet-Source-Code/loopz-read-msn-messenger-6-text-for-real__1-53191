VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Messenger"
   ClientHeight    =   3900
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   6255
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop Logging"
      Enabled         =   0   'False
      Height          =   435
      Left            =   1560
      TabIndex        =   3
      Top             =   3360
      Width           =   1275
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start Logging"
      Height          =   435
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   1275
   End
   Begin VB.TextBox Text1 
      Height          =   3015
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   6015
   End
   Begin VB.ComboBox cmbInterface 
      Height          =   315
      Left            =   3480
      TabIndex        =   0
      Text            =   "Interface List"
      Top             =   3360
      Width           =   2640
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Main"
      Visible         =   0   'False
      Begin VB.Menu mnuRemNode 
         Caption         =   "Remove Node"
      End
      Begin VB.Menu mnuExpandNode 
         Caption         =   "Expand Node"
      End
      Begin VB.Menu mnuCollapseNode 
         Caption         =   "Collapse Node"
      End
      Begin VB.Menu mnuViewData 
         Caption         =   "View Data"
      End
   End
   Begin VB.Menu mnuPopup2 
      Caption         =   "Main"
      Visible         =   0   'False
      Begin VB.Menu mnuExpandNode2 
         Caption         =   "Expand Node"
      End
      Begin VB.Menu mnuCollapseNode2 
         Caption         =   "Collapse Node"
      End
   End
   Begin VB.Menu mnuPopup3 
      Caption         =   "Main"
      Visible         =   0   'False
      Begin VB.Menu mnuRemoveNode3 
         Caption         =   "Remove Node"
      End
      Begin VB.Menu mnuExpandNode3 
         Caption         =   "Expand Node"
      End
      Begin VB.Menu mnuCollapseNode3 
         Caption         =   "Collapse Node"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ProtocolBuilder         As clsProtocolInterface

Private WithEvents TCPDriver    As clsTCPProtocol
Attribute TCPDriver.VB_VarHelpID = -1
Private WithEvents UDPDriver    As clsUDPProtocol
Attribute UDPDriver.VB_VarHelpID = -1
Private WithEvents ICMPDriver   As clsICMPProtocol
Attribute ICMPDriver.VB_VarHelpID = -1

Private BytesRecievedPackets    As Long
Private BytesRecieved           As Long
Private TCPPackets              As Long
Private UDPPackets              As Long
Private ICMPPackets             As Long
Private TCPLog                  As Integer
Private UDPLog                  As Integer
Private ICMPLog                 As Integer

Private Sub cmdStart_Click()
    If ProtocolBuilder.CreateRawSocket(Left$(cmbInterface.Text, InStr(1, cmbInterface, " ")), 7000, Me.hWnd) <> 0 Then
        cmdStart.Enabled = Not cmdStart.Enabled
        cmdStop.Enabled = Not cmdStop.Enabled
    End If
End Sub

Private Sub cmdStop_Click()
    cmdStart.Enabled = Not cmdStart.Enabled
    cmdStop.Enabled = Not cmdStop.Enabled
    ProtocolBuilder.CloseRawSocket
End Sub

Private Sub Form_Load()

  Dim str() As String, i As Integer

    Set ProtocolBuilder = New clsProtocolInterface
    Set TCPDriver = New clsTCPProtocol
    Set UDPDriver = New clsUDPProtocol
    Set ICMPDriver = New clsICMPProtocol

    ProtocolBuilder.AddinProtocol TCPDriver, "TCP", IPPROTO_TCP
    ProtocolBuilder.AddinProtocol UDPDriver, "UDP", IPPROTO_UDP
    ProtocolBuilder.AddinProtocol ICMPDriver, "ICMP", IPPROTO_ICMP

    str = Split(EnumNetworkInterfaces(), ";")
        
    For i = 0 To UBound(str)
        If str(i) <> "127.0.0.1" Then
            cmbInterface.AddItem str(i) & " [" & GetHostNameByAddr(inet_addr(str(i))) & "]"
        End If
    Next
    
    cmbInterface.Text = cmbInterface.List(0)

    cmdStart_Click
End Sub


Private Sub Form_Unload(Cancel As Integer)
    ProtocolBuilder.CloseRawSocket
    Set ProtocolBuilder = Nothing
    
    Set ICMPDriver = Nothing
    Set UDPDriver = Nothing
    Set TCPDriver = Nothing
End Sub


Private Sub TCPDriver_RecievedPacket(IPHeader As clsIPHeader, TCPProtocol As clsTCPProtocol, Data As String)

  Dim Parent    As Node
  Dim IPH       As Node
  Dim TCPH      As Node
  Dim Flags     As Node
  
  Dim strHeader As String
  Dim strData   As String
    
    strData = Space(40 - Len(strHeader)) & IIf(Len(Data) <= 35, Data, Left$(Data, 35) & "...")

    Text1.Text = Data
    
    TCPPackets = TCPPackets + 1
    BytesRecieved = BytesRecieved + LenB(Data)
    BytesRecievedPackets = BytesRecievedPackets + LenB(Data) + 40

End Sub


Private Sub Text1_Change()
Open "C:\log.txt" For Append As #1
Print #1, Text1
Close #1

If FoundInText(Text1, "SOMETHING") Then
MsgBox "FOUND."
    'here you can do anything you want...
    'just see if the text contains the data, then make it do something.
    
    'note that you'd better put the entire data-set you want to get,
    'like, for example:
    
    'MSG [data....]
    'MIME-Version: 1.0
    'Content-Type: text/plain; charset=UTF-8
    'X-MMS-IM-Format: [data....]
    '
    '[text of the message you received]
    
End If

End Sub

Private Sub UDPDriver_RecievedPacket(IPHeader As clsIPHeader, UDPProtocol As clsUDPProtocol, Data As String)

  Dim Parent    As Node
  Dim IPH       As Node
  Dim UDPH      As Node

  Dim strHeader As String
  Dim strData   As String
    
    strData = Space(40 - Len(strHeader)) & IIf(Len(Data) <= 35, Data, Left$(Data, 35) & "...")

    Text1.Text = Data
    
    UDPPackets = UDPPackets + 1
    BytesRecieved = BytesRecieved + LenB(Data)
    BytesRecievedPackets = BytesRecievedPackets + LenB(Data) + 28

End Sub

Private Function FoundInText(Text, StringToFind) As Boolean
If Replace(Text, StringToFind, "") = Text Then
FoundInText = False
Else
FoundInText = True
End If
End Function
