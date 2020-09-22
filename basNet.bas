Attribute VB_Name = "basNet"
'We need 2 NETRESORSE structures - one for get info
'from WNetEnumResourse, second (NETRESOURCE_STRING) for
'passing appropriate data to WNetOpenEnum
Private Type NETRESOURCE
   dwScope As Long
   dwType As Long
   dwDisplayType As Long
   dwUsage As Long
   lpLocalName As Long
   lpRemoteName As Long
   lpComment As Long
   lpProvider As Long
End Type

Private Type NETRESOURCE_STRING
   dwScope As Long
   dwType As Long
   dwDisplayType As Long
   dwUsage As Long
   lpLocalName As String
   lpRemoteName As String
   lpComment As String
   lpProvider As String
End Type

Type NetInfo       'UDT to store Data and use it for
   dwScope As Long 'filling ListView
   dwType As Long
   dwDisplayType As Long
   dwUsage As Long
   LocalName As String
   RemoteName As String
   Comment As String
   Provider As String
End Type

Private Declare Function WNetOpenEnum Lib "mpr.dll" Alias "WNetOpenEnumA" (ByVal dwScope As Long, ByVal dwType As Long, ByVal dwUsage As Long, lpNetResource As Any, lphEnum As Long) As Long
Private Declare Function WNetEnumResource Lib "mpr.dll" Alias "WNetEnumResourceA" (ByVal hEnum As Long, lpcCount As Long, ByVal lpBuffer As Long, lpBufferSize As Long) As Long
Private Declare Function WNetCloseEnum Lib "mpr.dll" (ByVal hEnum As Long) As Long

Const RESOURCE_CONNECTED = &H1
Const RESOURCE_GLOBALNET = &H2
Const RESOURCE_REMEMBERED = &H3
Const RESOURCE_RECENT = &H4
Const RESOURCE_CONTEXT = &H5

Public Const RESOURCETYPE_ANY = &H0
Public Const RESOURCETYPE_DISK = &H1
Public Const RESOURCETYPE_PRINT = &H2
Public Const RESOURCETYPE_RESERVED = &H8
Public Const RESOURCETYPE_UNKNOWN = &HFFFF

Public Const RESOURCEUSAGE_CONNECTABLE = &H1
Public Const RESOURCEUSAGE_CONTAINER = &H2
Public Const RESOURCEUSAGE_NOLOCALDEVICE = &H4
Public Const RESOURCEUSAGE_SIBLING = &H8
Public Const RESOURCEUSAGE_ALL = RESOURCEUSAGE_CONNECTABLE Or RESOURCEUSAGE_CONTAINER
Public Const RESOURCEUSAGE_RESERVED = &H80000000


Public Const RESOURCEDISPLAYTYPE_GENERIC = &H0
Public Const RESOURCEDISPLAYTYPE_DOMAIN = &H1
Public Const RESOURCEDISPLAYTYPE_SERVER = &H2
Public Const RESOURCEDISPLAYTYPE_SHARE = &H3
Public Const RESOURCEDISPLAYTYPE_FILE = &H4
Public Const RESOURCEDISPLAYTYPE_GROUP = &H5
Public Const RESOURCEDISPLAYTYPE_NETWORK = &H6
Public Const RESOURCEDISPLAYTYPE_ROOT = &H7
Public Const RESOURCEDISPLAYTYPE_SHAREADMIN = &H8
Public Const RESOURCEDISPLAYTYPE_DIRECTORY = &H9
Public Const RESOURCEDISPLAYTYPE_TREE = &HA

Private Const GMEM_FIXED = &H0
Private Const GMEM_ZEROINIT = &H40
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

Public NI() As NetInfo
Public NetCount As Long
Private NR As NETRESOURCE
Private NRS As NETRESOURCE_STRING
Public ShareName As String

Private Sub FillNRS(Index As Long)
   NRS.dwDisplayType = NI(Index).dwDisplayType
   NRS.dwScope = NI(Index).dwScope
   NRS.dwType = NI(Index).dwType
   NRS.dwUsage = NI(Index).dwUsage
   NRS.lpComment = NI(Index).Comment & Chr$(0)
   NRS.lpLocalName = NI(Index).LocalName & Chr$(0)
   NRS.lpProvider = NI(Index).Provider & Chr$(0)
   NRS.lpRemoteName = NI(Index).RemoteName & Chr$(0)
End Sub

Private Sub ClearNr()
  NR.dwDisplayType = 0&
  NR.dwScope = 0&
  NR.dwType = 0&
  NR.dwUsage = 0&
  NR.lpComment = 0&
  NR.lpLocalName = 0&
  NR.lpProvider = 0&
  NR.lpRemoteName = 0&
End Sub

Private Sub ClearNrs()
  NRS.dwDisplayType = 0&
  NRS.dwScope = 0&
  NRS.dwType = 0&
  NRS.dwUsage = 0&
  NRS.lpComment = Chr$(0)
  NRS.lpLocalName = Chr$(0)
  NRS.lpProvider = Chr$(0)
  NRS.lpRemoteName = Chr$(0)
End Sub

Private Sub FillInfo(Index As Long)
  NI(Index).dwScope = NR.dwScope
  NI(Index).dwDisplayType = NR.dwDisplayType
  NI(Index).dwType = NR.dwType
  NI(Index).dwUsage = NR.dwUsage
  NI(Index).RemoteName = PointerToString(NR.lpRemoteName)
  NI(Index).LocalName = PointerToString(NR.lpLocalName)
  NI(Index).Comment = PointerToString(NR.lpComment)
  NI(Index).Provider = PointerToString(NR.lpProvider)
End Sub

Public Function NetEnumChild(sRN As String) As Long
  Dim hEnum As Long, lpBuff As Long
  Dim cbBuff As Long, cCount As Long
  Dim p As Long, res As Long, i As Long

  On Error GoTo ErrorHandler
  NRS.lpRemoteName = sRN & Chr$(0)
  cbBuff = 16384
  cCount = &HFFFFFFFF

res = WNetOpenEnum(RESOURCE_GLOBALNET, RESOURCETYPE_ANY, RESOURCEUSAGE_CONTAINER, NRS, hEnum)
'If no error ocuured, res = 0, else ret > 0
'If you need, you can handle more errors
If res = 67 Then MsgBox "Sorry, this domain is not availiable now", vbExclamation, "Network Error"
If res = 1244 Then MsgBox "Password reguired", vbExclamation, "Network Error"
If res = 0 Then
   lpBuff = GlobalAlloc(GPTR, cbBuff)
   res = WNetEnumResource(hEnum, cCount, lpBuff, cbBuff)
   If res = 0 Then
      ReDim NI(cCount)
      p = lpBuff
      For i = 1 To cCount
         CopyMemory NR, ByVal p, LenB(NR)
         FillInfo i
         p = p + LenB(NR)
      Next i
   End If
ErrorHandler:
On Error Resume Next
   If lpBuff <> 0 Then GlobalFree (lpBuff)
   WNetCloseEnum (hEnum)
End If
NetEnumChild = cCount
End Function

Public Sub NetEnumLocal()
'This function enums neighborhood
  Dim hEnum As Long, lpBuff As Long
  Dim cbBuff As Long, cCount As Long
  Dim p As Long, res As Long, i As Long

  On Error GoTo ErrorHandler
  ClearNr
  cbBuff = 16384
  cCount = &HFFFFFFFF
res = WNetOpenEnum(RESOURCE_CONTEXT, RESOURCETYPE_ANY, RESOURCEUSAGE_CONTAINER, NR, hEnum)
If res = 0 Then
   lpBuff = GlobalAlloc(GPTR, cbBuff)
   res = WNetEnumResource(hEnum, cCount, lpBuff, cbBuff)
   If res = 0 Then
      ReDim NI(cCount)
      p = lpBuff
      For i = 1 To cCount
         CopyMemory NR, ByVal p, LenB(NR)
         FillInfo i
         p = p + LenB(NR)
      Next i
   End If
ErrorHandler:
On Error Resume Next
   If lpBuff <> 0 Then GlobalFree (lpBuff)
   WNetCloseEnum (hEnum)
End If
End Sub

