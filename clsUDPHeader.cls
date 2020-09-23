VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUDPProtocol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type UDPHeader
    src_portno                  As Integer
    dst_portno                  As Integer
    udp_length                  As Integer
    udp_checksum                As Integer
End Type


Private m_src_portno            As Integer
Private m_dst_portno            As Integer
Private m_udp_length            As Integer
Private m_udp_checksum          As Long

Public ProtocolInterface        As clsProtocolInterface

Public Event RecievedPacket(IPHeader As clsIPHeader, UDPProtocol As clsUDPProtocol, Data As String)


Public Sub PacketArrived(IPH As clsIPHeader, ReadBuffer() As Byte, BytesRecieved As Long)

  Dim UDPH As UDPHeader

  'Payload
  Dim DataOffset    As Integer
  Dim PayloadSize   As Long
  Dim bPayload()    As Byte
  Dim strPayload    As String
  
    CopyMemory UDPH, ByVal VarPtr(ReadBuffer(0)) + 20, LenB(UDPH)
    
    DataOffset = 20 + LenB(UDPH)
    PayloadSize = BytesRecieved - DataOffset
    
    If PayloadSize > 0 Then
        ReDim bPayload(0 To PayloadSize - 1)
        CopyMemory ByVal VarPtr(bPayload(0)), ByVal VarPtr(ReadBuffer(DataOffset)), PayloadSize
        strPayload = StrConv(bPayload, vbUnicode)
    End If
    
    Checksum = IntegerToUnsigned(ntohs(UDPH.udp_checksum))
    DestPort = ntohs(UDPH.dst_portno)
    HeaderLength = ntohs(UDPH.udp_length)
    SourcePort = ntohs(UDPH.src_portno)
    
    RaiseEvent RecievedPacket(IPH, Me, strPayload)
    
End Sub


Public Sub SendPacket(strDestination As String, DestPort As Long, IPHeader As clsIPHeader, Data As String)

  Dim IPH        As IPHeader
  Dim UDPH       As UDPHeader
  Dim Packet()   As Byte
  Dim bData()    As Byte
    
    With IPH
        .ip_checksum = htons(IPHeader.Checksum)
        .ip_destaddr = IPHeader.DestAddress
        .ip_id = htons(IPHeader.ID)
        .ip_offset = htons(IPHeader.Offset)
        .ip_protocol = IPHeader.Protocol
        .ip_srcaddr = IPHeader.SourceAddress
        .ip_tos = IPHeader.TypeOfService
        .ip_totallength = htons(IPHeader.Checksum)
        .ip_ttl = IPHeader.TypeOfService
        .ip_verlen = MakeByte(IPHeader.HeaderLength, IPHeader.Version)
    End With
    
    With UDPH
        .dst_portno = htons(DestPort)
        .src_portno = htons(SourcePort)
        .udp_checksum = htons(Checksum)
        .udp_length = htons(HeaderLength)
    End With
    
    bData = StrConv(Data, vbFromUnicode)
    
    ReDim Packet(LenB(IPH) + LenB(UDPH) + UBound(bData))
    
    CopyMemory ByVal VarPtr(Packet(0)), ByVal VarPtr(IPH), LenB(IPH)
    CopyMemory ByVal VarPtr(Packet(0)) + LenB(IPH), ByVal VarPtr(UDPH), LenB(UDPH)
    CopyMemory ByVal VarPtr(Packet(0)) + LenB(IPH) + LenB(UDPH), ByVal VarPtr(bData(0)), UBound(bData) + 1
    
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



Public Property Get HeaderLength() As Integer
    HeaderLength = m_udp_length
End Property

Public Property Let HeaderLength(val As Integer)
    m_udp_length = val
End Property



Public Property Get Checksum() As Long
    Checksum = m_udp_checksum
End Property

Public Property Let Checksum(val As Long)
    m_udp_checksum = val
End Property
