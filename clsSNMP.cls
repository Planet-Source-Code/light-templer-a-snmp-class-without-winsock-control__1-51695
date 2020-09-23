VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSNMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

'
'   clsSNMP.cls
'

'   Started:    10.02.2004
'   By:         LightTempler
'
'   Last edit:  12.02.2004

'   Original source by  :   Randol Tigrett
'   Released on         :   http://www.freevbcode.com/
'   Released at         :   feb 2002

'   clsSocket class     :   Oleg Gdalevich, plz have a look to the class for all details.

'   Changes by Light Templer:
'
'       - Added  'Option explicit'  ;)
'       - Handled, what followed after this ... ;(
'       - Removed unused var / code
'       - Some improvements, sharper scopes for subs/funcs/vars, code shortening,
'         handling for 'empty comunity string' and 'proper IP adress display'
'       - Whole code reformated by hand, some vars renamed to a usual scheme
'       - Enums implemented for PDUget / PDUgetNext
'       - Substituded the WinSock control with the VBIP winsock class 'cSocket'.
'       - Conversions to use it with VB5, too.
'       - Bugfix clsSocket (got GPFs on CopyMem when host ip adress cannot be resolved to a name...)
'
'       - Taken from a form/bas (on a "proof-of-concept"-level of code) into a (re-)useable class.
'
'       - Some tests against our servers and printers - no problems so far.
'       - Added support for an (not so important) 'Get SNMPlevel()'
'       - Improvements to the demo form.
'       - ...

' ### MANY thx, Randol, for this source! SNMP with VB is really a rare theme ;) ###

' Thx for any hints and votes on PSC !
'
' My email adress: schwepps_bitterlemon@gmx.de

' Kind regards
' Light Templer (LiTe)
'

Option Explicit


' *******************************
' *            EVENTS           *
' *******************************
Public Event Error(sErrMsg As String)
Public Event Result(sValue As String)



' *************************************
' *            CONSTANTS              *
' *************************************
Private Const DEFAULT_COMUNITY_NAME = "public"



' *************************************
' *        PRIVATE ENUMS              *
' *************************************
Private Enum enPDU
    pduGET = &HA0
    pduGETNEXT = &HA1
End Enum


' *************************************
' *            PRIVATES               *
' *************************************

' SNMP handling stuff
Private snmpBinary(255) As Byte

Private Type verSNMP
   byteType     As Byte
   packetLen    As Byte
   packetData() As Byte
End Type

Private Type commSNMP
   byteType     As Byte
   packetLen    As Byte
   packetData() As Byte
End Type

Private Type contextSNMP
   byteType     As Byte
   packetLen    As Byte
End Type

Private Type requestSNMP
   byteType     As Byte
   packetLen    As Byte
   packetData() As Byte
End Type

Private Type errorSNMP
   byteType     As Byte
   packetLen    As Byte
   packetData() As Byte
End Type

Private Type indexSNMP
   byteType     As Byte
   packetLen    As Byte
   packetData() As Byte
End Type

Private Type struct1SNMP
   byteType     As Byte
   packetLen    As Byte
End Type

Private Type struct2SNMP
   byteType     As Byte
   packetLen    As Byte
End Type

Private Type objectSNMP
   byteType     As Byte
   packetLen    As Byte
   packetData() As Byte
End Type

Private Type endSNMP
   byte01       As Byte
   byte02       As Byte
End Type

Private Type snmpPacket
   UniSeq       As Byte
   packetLenth  As Byte             ' Length of Packet minus 2
   verSNMP      As verSNMP
   commSNMP     As commSNMP
   contextSNMP  As contextSNMP
   requestSNMP  As requestSNMP
   errorSNMP    As errorSNMP
   indexSNMP    As indexSNMP
   struct1SNMP  As struct1SNMP
   struct2SNMP  As struct2SNMP
   objectSNMP   As objectSNMP
   endSNMP      As endSNMP
End Type

Private Type snmpData
   UniSeq01     As Byte
   UniSeq02     As Byte
   packetLenth  As Byte             ' Length of Packet minus 2
   verSNMP      As verSNMP
   commSNMP     As commSNMP
   contextSNMP  As contextSNMP
   requestSNMP  As requestSNMP
   errorSNMP    As errorSNMP
   indexSNMP    As indexSNMP
   struct1SNMP  As struct1SNMP
   struct2SNMP  As struct2SNMP
   objectSNMP   As objectSNMP
   endSNMP      As endSNMP
End Type


' class stuff
Private Type tpMvar
    DestHost        As String
    ComunityName    As String
    Oid             As String
    sSNMPresult     As String
    SNMPversion     As Long
End Type
Private mvar As tpMvar

Private WithEvents oSocket As clsSocket
Attribute oSocket.VB_VarHelpID = -1
'
'
'


' *************************************
' *            INIT/TERM              *
' *************************************
Private Sub Class_Initialize()

    On Local Error GoTo error_handler

    mvar.ComunityName = DEFAULT_COMUNITY_NAME
    Set oSocket = New clsSocket
    With oSocket
        .Protocol = sckUDPProtocol
        .LocalPort = 0
        .RemotePort = 161
    End With

    Exit Sub


error_handler:
    
    RaiseEvent Error("[" + Err.Description + "] in 'Class_Initialize()")
    
End Sub

Private Sub Class_Terminate()

    Set oSocket = Nothing

End Sub



' *************************************
' *         PUBLIC FUNCTIONS          *
' *************************************
Public Sub SNMPGet(Optional sOid As String = "")
    
    If sOid <> "" Then
        mvar.Oid = sOid
    End If
    If mvar.DestHost = "" Or mvar.Oid = "" Then
        RaiseEvent Error("[Missing parameter (DestHost or Oid)!] in 'SNMPGet()'")
        
        Exit Sub
    End If
    
    SNMPsendRequest pduGET
    
End Sub

Public Sub SNMPGetNext()
 
    If mvar.DestHost = "" Or mvar.Oid = "" Then
        RaiseEvent Error("[Missing parameter (DestHost or Oid)!] in 'SNMPGetNext()'")
        
        Exit Sub
    End If
    
    SNMPsendRequest pduGETNEXT
 
End Sub




' *************************************
' *         PRIVATE FUNCTIONS         *
' *************************************
Private Sub SNMPsendRequest(pdu As enPDU)
    
    Dim xsnmp    As snmpPacket
    Dim OIDArray As Variant
    Dim OIDsize  As Long
    Dim i        As Long
    
    On Local Error GoTo error_handler
    
    With mvar
        xsnmp.UniSeq = &H30
        xsnmp.verSNMP.byteType = 2
        xsnmp.verSNMP.packetLen = 1
        ReDim xsnmp.verSNMP.packetData(0)
        xsnmp.verSNMP.packetData(0) = 0
        
        xsnmp.commSNMP.byteType = 4
        If .ComunityName <> "" Then
            ReDim xsnmp.commSNMP.packetData(CByte(Len(Trim$(.ComunityName)) - 1))
            For i = 0 To UBound(xsnmp.commSNMP.packetData)
                xsnmp.commSNMP.packetData(i) = Asc(Mid(Trim$(.ComunityName), i + 1, 1))
            Next i
            xsnmp.commSNMP.packetLen = CByte(UBound(xsnmp.commSNMP.packetData) + 1)
        End If
                
        xsnmp.contextSNMP.byteType = CByte(pdu)
        
        xsnmp.requestSNMP.byteType = 2
        ReDim xsnmp.requestSNMP.packetData(0)
        xsnmp.requestSNMP.packetData(0) = 1
        xsnmp.requestSNMP.packetLen = 1
        
        xsnmp.errorSNMP.byteType = 2
        ReDim xsnmp.errorSNMP.packetData(0)
        xsnmp.errorSNMP.packetData(0) = 0
        xsnmp.errorSNMP.packetLen = 1
        
        xsnmp.indexSNMP.byteType = 2
        ReDim xsnmp.indexSNMP.packetData(0)
        xsnmp.indexSNMP.packetData(0) = 0
        xsnmp.indexSNMP.packetLen = 1
        
        xsnmp.struct1SNMP.byteType = &H30
        xsnmp.struct2SNMP.byteType = &H30
        
        xsnmp.objectSNMP.byteType = 6
        
        OIDArray = Split(Trim(.Oid), ".", , vbBinaryCompare)
        OIDsize = UBound(OIDArray)
        If OIDArray(UBound(OIDArray)) = "" Then
            OIDsize = OIDsize - 1
        End If
            
        ReDim xsnmp.objectSNMP.packetData(OIDsize - 1)
        xsnmp.objectSNMP.packetData(0) = &H2B
        For i = 2 To OIDsize
            xsnmp.objectSNMP.packetData(i - 1) = OIDArray(i)
        Next i
        
        xsnmp.endSNMP.byte01 = 5
        xsnmp.endSNMP.byte02 = 0
        xsnmp.objectSNMP.packetLen = OIDsize
        xsnmp.struct2SNMP.packetLen = xsnmp.objectSNMP.packetLen + 4
        xsnmp.struct1SNMP.packetLen = xsnmp.struct2SNMP.packetLen + 2
        xsnmp.contextSNMP.packetLen = xsnmp.struct1SNMP.packetLen + 11
        
        Call convertBinArray(xsnmp)
        
        ' Now send the request!
        oSocket.RemoteHost = .DestHost
        oSocket.SendData snmpBinary()
    End With
    
    Exit Sub


error_handler:
    
    RaiseEvent Error("[" + Err.Description + "] in 'SNMPsendRequest()'")
    
End Sub


Private Sub convertBinArray(xsnmp As snmpPacket)
    ' Prepare for sending SNMP request
   
    Dim i       As Long
    Dim lPos    As Long
    

    snmpBinary(0) = xsnmp.UniSeq
    
    snmpBinary(2) = xsnmp.verSNMP.byteType
    snmpBinary(3) = xsnmp.verSNMP.packetLen
    snmpBinary(4) = xsnmp.verSNMP.packetData(0)
    
    snmpBinary(5) = xsnmp.commSNMP.byteType
    snmpBinary(6) = xsnmp.commSNMP.packetLen
    
    For i = 0 To xsnmp.commSNMP.packetLen - 1
        snmpBinary(7 + i) = xsnmp.commSNMP.packetData(i)
    Next i
    lPos = 7 + i
    
    snmpBinary(lPos) = xsnmp.contextSNMP.byteType
    snmpBinary(lPos + 1) = xsnmp.contextSNMP.packetLen
    lPos = lPos + 2
    
    snmpBinary(lPos) = xsnmp.requestSNMP.byteType
    snmpBinary(lPos + 1) = xsnmp.requestSNMP.packetLen
    snmpBinary(lPos + 2) = xsnmp.requestSNMP.packetData(0)
    lPos = lPos + 3
    
    snmpBinary(lPos) = xsnmp.errorSNMP.byteType
    snmpBinary(lPos + 1) = xsnmp.errorSNMP.packetLen
    snmpBinary(lPos + 2) = xsnmp.errorSNMP.packetData(0)
    lPos = lPos + 3
    
    snmpBinary(lPos) = xsnmp.indexSNMP.byteType
    snmpBinary(lPos + 1) = xsnmp.indexSNMP.packetLen
    snmpBinary(lPos + 2) = xsnmp.indexSNMP.packetData(0)
    lPos = lPos + 3
    
    snmpBinary(lPos) = xsnmp.struct1SNMP.byteType
    snmpBinary(lPos + 1) = xsnmp.struct1SNMP.packetLen
    lPos = lPos + 2
    
    snmpBinary(lPos) = xsnmp.struct2SNMP.byteType
    snmpBinary(lPos + 1) = xsnmp.struct2SNMP.packetLen
    lPos = lPos + 2
    
    snmpBinary(lPos) = xsnmp.objectSNMP.byteType
    snmpBinary(lPos + 1) = xsnmp.objectSNMP.packetLen
    lPos = lPos + 2
    
    For i = 0 To xsnmp.objectSNMP.packetLen - 1
        snmpBinary(lPos + i) = xsnmp.objectSNMP.packetData(i)
    Next i
    lPos = lPos + i
    
    snmpBinary(lPos) = xsnmp.endSNMP.byte01
    snmpBinary(lPos + 1) = xsnmp.endSNMP.byte02
    snmpBinary(1) = CByte(lPos)

End Sub

Private Sub oSocket_OnDataArrival(ByVal bytesTotal As Long)
    
    Dim sOIDData    As String
    Dim sOIDused    As String

    oSocket.GetData sOIDData
    Call convertSnmp(sOIDData, sOIDused)
    RaiseEvent Result(mvar.sSNMPresult)
        
End Sub


Private Sub oSocket_OnError(ByVal Number As Integer, _
                            Description As String, _
                            ByVal Scode As Long, _
                            ByVal Source As String, _
                            ByVal HelpFile As String, _
                            ByVal HelpContext As Long, _
                            CancelDisplay As Boolean)

    RaiseEvent Error("[Socket Error! - " & Description & "  - Source " & Source)

End Sub


Private Function convertSnmp(sSNMPdata As String, ByRef sTrueOID As String) As String
    
    Dim xbinary()   As Byte
    Dim lLen        As Long
    Dim i           As Long
    
    mvar.sSNMPresult = ""
            
    lLen = Len(sSNMPdata)
    ReDim xbinary(lLen)
    
    For i = 1 To lLen
        xbinary(i - 1) = Asc(Mid(sSNMPdata, i, 1))
    Next i
    
    If xbinary(0) = 48 And ((xbinary(1) And &HF0) = 128) And xbinary(2) <> 2 Then
        mvar.SNMPversion = 2
        Call Version02(xbinary)
    Else
        mvar.SNMPversion = 1
        Call Version01(xbinary)
    End If
        
End Function


Private Sub Version01(xbinary() As Byte)
   
    Dim lDataLen    As Long
    Dim lPos        As Long
    Dim i           As Long

    lDataLen = xbinary(1)
    lPos = 1
    lDataLen = xbinary(lPos + 1)
    lPos = lPos + 1
    lPos = lPos + lDataLen
    lPos = lPos + 1
    lDataLen = xbinary(lPos + 1)
    lPos = lDataLen + lPos + 3
    lDataLen = lPos + 1
    lPos = 1 + lDataLen
    lDataLen = xbinary(lPos)
    lPos = lDataLen + lPos + 2
    lPos = lPos + 3
    lDataLen = xbinary(lPos)
    lPos = lDataLen + lPos + 1
    
    For i = 1 To 2
        If xbinary(lPos) = &H30 Then
            lPos = lPos + 1
            lDataLen = 1
            lPos = lPos + lDataLen
        End If
    Next i
    
    FilterSNMPresult lPos, xbinary

End Sub


Private Sub Version02(xbinary() As Byte)

    Dim lDataLen    As Long
    Dim lPos        As Long
    Dim i           As Long
    
    lDataLen = xbinary(1) And &HF
    lPos = 1 + lDataLen
    lDataLen = xbinary(lPos + 1)
    lPos = lPos + 1
    lPos = lPos + lDataLen
    lPos = lPos + 1
    lDataLen = xbinary(lPos + 1)
    lPos = lDataLen + lPos + 3
    lDataLen = lPos + (xbinary(lPos) And &HF)
    lPos = 2 + lDataLen
    lDataLen = xbinary(lPos)
    lPos = lDataLen + lPos + 2
    lPos = lPos + 3
    lDataLen = xbinary(lPos)
    lPos = lDataLen + lPos + 1
    
    For i = 1 To 2
        If xbinary(lPos) = &H30 Then
            lPos = lPos + 1
            lDataLen = xbinary(lPos) And &HF
            lPos = lPos + lDataLen + 1
        End If
    Next i
    
    FilterSNMPresult lPos, xbinary

End Sub


Private Sub FilterSNMPresult(pos As Long, xbinary() As Byte)

    Dim OIDArray    As Variant
    Dim OIDString   As String
    Dim lDataLen    As Long
    Dim i           As Long
    
    On Local Error GoTo error_handler

    With mvar
        Select Case xbinary(pos)    ' Evaluate Byte
            
            Case &H2, &H41, &H42, &H43          ' Display Timer Ticks
                    pos = pos + 1
                    For i = 1 To xbinary(pos)
                        .sSNMPresult = .sSNMPresult & Hex(xbinary(pos + i))
                    Next i
                    .sSNMPresult = Format("&H" & .sSNMPresult)
                    
                    
            Case &H3                            ' Display Bits
                    pos = pos + 1
                    For i = 1 To xbinary(pos)
                        .sSNMPresult = .sSNMPresult & Hex(xbinary(pos + i))
                    Next i
                    .sSNMPresult = Val("&H" & .sSNMPresult)
                    
                    
            Case &H40                           ' Display IP Addr
                    pos = pos + 1
                    For i = 1 To xbinary(pos)
                        .sSNMPresult = .sSNMPresult & xbinary(pos + i) & "."
                    Next i
                    .sSNMPresult = Left$(.sSNMPresult, Len(.sSNMPresult) - 1)
            
                    
            Case &H4                            ' Display String
                    pos = pos + 1
                    .sSNMPresult = Space$(xbinary(pos))
                    For i = 1 To xbinary(pos)
                        Mid$(.sSNMPresult, i, 1) = Chr(xbinary(pos + i))
                    Next i
                    
                    
            Case &H5   ' Display NULL
                    pos = pos + 1
                    .sSNMPresult = "NULL "
                    
                    
            Case &H6   ' Process Object
                    Do While (pos <= (xbinary(1) - 2))
                        DoEvents
                        If xbinary(pos) = 6 Then
                            pos = pos + 1
                            lDataLen = xbinary(pos)
                            For i = 1 To lDataLen
                                OIDString = OIDString & Hex(xbinary(pos + i)) & "#"
                            Next i
                            OIDArray = Split(OIDString, "#", , vbBinaryCompare)
                            .Oid = "1.3"
                            For i = 1 To UBound(OIDArray) - 1
                                .Oid = .Oid & "." & Val("&H" & OIDArray(i))
                            Next i
                            pos = pos + lDataLen + 1
                            If xbinary(pos) = 6 Then
                                pos = pos + 2
                            End If
                            FilterSNMPresult pos, xbinary
                            pos = pos + xbinary(pos) + 1
                        End If
                        pos = 1 + pos
                    Loop
                    
                    
            Case Else   ' Other values.
                    .sSNMPresult = "{Unknown SNMP value type} - Value:  &H" & Val("&H" & xbinary(pos))
                    
                    
        End Select
    End With

    Exit Sub


error_handler:
    
    RaiseEvent Error("[" + Err.Description + "] in 'FilterSNMPresult()")
     
End Sub


' *************************************
' *           PROPERTIES              *
' *************************************
Public Property Let Oid(ByVal sNewOid As String)
    
    mvar.Oid = sNewOid
 
End Property

Public Property Get Oid() As String
    
    Oid = mvar.Oid
 
End Property


Public Property Let ComunityName(ByVal sNewComunityName As String)
    
    mvar.ComunityName = sNewComunityName
 
End Property

Public Property Get ComunityName() As String
    
    ComunityName = mvar.ComunityName
 
End Property

Public Property Get SNMPversion() As Long
    
    SNMPversion = mvar.SNMPversion
 
End Property


Public Property Let DestHost(ByVal sNewDestHost As String)
    
    mvar.DestHost = sNewDestHost
 
End Property

Public Property Get DestHost() As String
    
    DestHost = mvar.DestHost
 
End Property


' #*#
