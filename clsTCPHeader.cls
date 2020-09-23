VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTCPProtocol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Type TCPHeader
    src_portno                  As Integer
    dst_portno                  As Integer
    Sequenceno                  As Long
    Acknowledgeno               As Long
    DataOffset                  As Byte
    flag                        As Byte
    Windows                     As Integer
    Checksum                    As Integer
    UrgentPointer               As Integer
End Type


Public Enum TCPFlag
             TCPF_FIN = 1
             TCPF_SYN = 2
             TCPF_RST = 4
             TCPF_PSH = 8
             TCPF_ACK = 16
             TCPF_URG = 32
End Enum


Private m_src_portno            As Integer
Private m_dst_portno            As Integer
Private m_Sequenceno            As Double
Private m_Acknowledgeno         As Double
Private m_DataOffset            As Byte
Private m_flag                  As Byte
Private m_Windows               As Integer
Private m_checksum              As Long
Private m_UrgentPointer         As Integer

Public ProtocolInterface        As clsProtocolInterface


Public Event RecievedPacket(IPHeader As clsIPHeader, TCPProtocol As clsTCPProtocol, Data As String)


Public Sub PacketArrived(IPH As clsIPHeader, ReadBuffer() As Byte, BytesRecieved As Long)

  Dim TCPH As TCPHeader

  'Payload
  Dim DataOffset    As Integer
  Dim PayloadSize   As Long
  Dim bPayload()    As Byte
  Dim strPayload    As String
  
    CopyMemory TCPH, ByVal VarPtr(ReadBuffer(0)) + 20, LenB(TCPH)
    
    DataOffset = 20 + (TCPH.DataOffset \ &H10) * 4
    PayloadSize = BytesRecieved - DataOffset
    
    If PayloadSize > 0 Then
        ReDim bPayload(0 To PayloadSize - 1)
        CopyMemory ByVal VarPtr(bPayload(0)), ByVal VarPtr(ReadBuffer(DataOffset)), PayloadSize
        strPayload = StrConv(bPayload, vbUnicode)
    End If
    
    AckNumber = LongToUnsigned(ntohl(TCPH.Acknowledgeno))
    Checksum = IntegerToUnsigned(ntohs(TCPH.Checksum))
    DataOffset = DataOffset
    DestPort = ntohs(TCPH.dst_portno)
    Flags = TCPH.flag
    SequenceNumber = LongToUnsigned(ntohl(TCPH.Sequenceno))
    SourcePort = ntohs(TCPH.src_portno)
    UrgentPointer = ntohs(TCPH.UrgentPointer)
    Windows = ntohs(TCPH.Windows)

    RaiseEvent RecievedPacket(IPH, Me, strPayload)
    
End Sub


Public Sub SendPacket(strDestination As String, DestPort As Long, IPHeader As clsIPHeader, Data As String)

  Dim IPH        As IPHeader
  Dim TCPH       As TCPHeader
  Dim Packet()   As Byte
  Dim bData()    As Byte
    
    With IPH
        .ip_checksum = htons(UnsignedToInteger(IPHeader.Checksum))
        .ip_destaddr = IPHeader.DestAddress
        .ip_id = htons(UnsignedToInteger(IPHeader.ID))
        .ip_offset = htons(IPHeader.Offset)
        .ip_protocol = IPHeader.Protocol
        .ip_srcaddr = IPHeader.SourceAddress
        .ip_tos = IPHeader.TypeOfService
        .ip_totallength = htons(IPHeader.Checksum)
        .ip_ttl = IPHeader.TypeOfService
        .ip_verlen = MakeByte(IPHeader.HeaderLength, IPHeader.Version)
    End With
    
    With TCPH
        .Acknowledgeno = htonl(UnsignedToLong(AckNumber))
        .Checksum = htons(UnsignedToInteger(Checksum))
        .DataOffset = DataOffset
        .dst_portno = htons(DestPort)
        .flag = Flags
        .Sequenceno = htonl(UnsignedToLong(SequenceNumber))
        .src_portno = htons(SourcePort)
        .UrgentPointer = htons(UrgentPointer)
        .Windows = htons(Windows)
    End With
    
    bData = StrConv(Data, vbFromUnicode)
    
    ReDim Packet(LenB(IPH) + LenB(TCPH) + UBound(bData))
    
    CopyMemory ByVal VarPtr(Packet(0)), ByVal VarPtr(IPH), LenB(IPH)
    CopyMemory ByVal VarPtr(Packet(0)) + LenB(IPH), ByVal VarPtr(TCPH), LenB(TCPH)
    CopyMemory ByVal VarPtr(Packet(0)) + LenB(IPH) + LenB(TCPH), ByVal VarPtr(bData(0)), UBound(bData) + 1
    
    ProtocolInterface.SendData strDestination, DestPort, Packet
    
End Sub


Public Property Get SourcePort() As ServicePort
    SourcePort = m_src_portno
End Property

Public Property Let SourcePort(val As ServicePort)
    m_src_portno = val
End Property



Public Property Get DestPort() As ServicePort
    DestPort = m_dst_portno
End Property

Public Property Let DestPort(val As ServicePort)
    m_dst_portno = val
End Property



Public Property Get SequenceNumber() As Double
    SequenceNumber = m_Sequenceno
End Property

Public Property Let SequenceNumber(val As Double)
    m_Sequenceno = val
End Property



Public Property Get AckNumber() As Double
    AckNumber = m_Acknowledgeno
End Property

Public Property Let AckNumber(val As Double)
    m_Acknowledgeno = val
End Property



Public Property Get DataOffset() As Byte
    DataOffset = m_DataOffset
End Property

Public Property Let DataOffset(val As Byte)
    m_DataOffset = val
End Property



Public Property Get Flags() As Byte
    Flags = m_flag
End Property

Public Property Let Flags(val As Byte)
    m_flag = val
End Property



Public Property Get Windows() As Integer
    Windows = m_Windows
End Property

Public Property Let Windows(val As Integer)
    m_Windows = val
End Property



Public Property Get Checksum() As Long
    Checksum = m_checksum
End Property

Public Property Let Checksum(val As Long)
    m_checksum = val
End Property



Public Property Get UrgentPointer() As Integer
    UrgentPointer = m_UrgentPointer
End Property

Public Property Let UrgentPointer(val As Integer)
    m_UrgentPointer = val
End Property



Public Function IsFlagSet(aFlag As TCPFlag) As Boolean
    IsFlagSet = m_flag And aFlag
End Function



Public Function GetTCPFlagStr() As String
    GetTCPFlagStr = GetTCPFlagStr & IIf(m_flag And TCPF_URG, "U", "x")      ' URG
    GetTCPFlagStr = GetTCPFlagStr & IIf(m_flag And TCPF_ACK, "A", "x")      ' ACK
    GetTCPFlagStr = GetTCPFlagStr & IIf(m_flag And TCPF_PSH, "P", "x")      ' PSH
    GetTCPFlagStr = GetTCPFlagStr & IIf(m_flag And TCPF_RST, "R", "x")      ' RST
    GetTCPFlagStr = GetTCPFlagStr & IIf(m_flag And TCPF_SYN, "S", "x")      ' SYN
    GetTCPFlagStr = GetTCPFlagStr & IIf(m_flag And TCPF_FIN, "F", "x")      ' FIN
End Function

