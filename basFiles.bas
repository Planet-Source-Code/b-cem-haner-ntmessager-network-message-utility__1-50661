Attribute VB_Name = "basFiles"
Option Explicit

Public Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Public Type SYSTEMTIME
  wYear             As Integer
  wMonth            As Integer
  wDayOfWeek        As Integer
  wDay              As Integer
  wHour             As Integer
  wMinute           As Integer
  wSecond           As Integer
  wMilliseconds     As Long
End Type
Private Const MAX_PATH = 260
Public Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type

Private Type SHFILEINFO
   hIcon As Long
   iIcon As Long
   dwAttributes As Long
   szDisplayName As String * MAX_PATH
   szTypeName As String * 80
End Type

Public Type FILE_INFO 'UDT to store Files_Info
   sFileName As String
   sFileSize As String
   sFileTime As String
   sFileRoot As String
   sFileNameWithExt As String
   sFileDescription As String
   sFileDate As String
   bIsFolder As Boolean
   hSmallIcon As Long
   hLargeIcon As Long
   nIcon As Long
End Type

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CopyPointer2String Lib "kernel32" Alias "lstrcpyA" (ByVal NewString As String, ByVal OldString As Long) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

Private Const MAXDWORD = &HFFFFFFF
Private Const INVALID_HANDLE_VALUE = -1
'File Attribute Const
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
'SHFile Info Const
Private Const SHGFI_ICON = &H100
Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_TYPENAME = &H400
Private Const SHGFI_ATTRIBUTES = &H800
Private Const SHGFI_ICONLOCATION = &H1000
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000
Private Const SHGFI_LINKOVERLAY = &H8000
Private Const SHGFI_SELECTED = &H10000
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1
Private Const SHGFI_OPENICON = &H2
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_PIDL = &H8
Private Const SHGFI_USEFILEATTRIBUTES = &H10
Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE Or SHGFI_ICON

Public FI() As FILE_INFO
   
Function GetHiWord(dw As Long) As Long
  If dw And &H80000000 Then
     GetHiWord = (dw \ 65535) - 1
  Else
     GetHiWord = dw \ 65535
  End If
End Function

Function GetLoWord(dw As Long) As Long
   If dw And &H8000& Then
      GetLoWord = &H8000 Or (dw And &H7FFF&)
   Else
      GetLoWord = dw And &HFFFF&
   End If
End Function

Public Function PointerToString(p As Long) As String
   Dim s As String
   s = String(255, Chr$(0))
   CopyPointer2String s, p
   PointerToString = Left(s, InStr(s, Chr$(0)) - 1)
End Function

Private Function GetPointerToString(lpString As Long, nBytes As Long) As String
   Dim Buffer As String
   If nBytes Then
      Buffer = Space$(nBytes)
      CopyMemory ByVal Buffer, ByVal lpString, nBytes
      GetPointerToString = Buffer
   End If
End Function

Public Function NetEnumFolders(sPath As String) As Long
'Enumerate Folders and files
   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long, nRet As Long
   Dim nSize As Long, cCount As Long
   Dim sTmp As String
   Dim shfi As SHFILEINFO
   ReDim FI(1000)
   hFile = FindFirstFile(sPath & "\*.*", WFD)
   If hFile <> INVALID_HANDLE_VALUE Then
      Do
         sTmp = TrimNull(WFD.cFileName)
         If sTmp <> "." And sTmp <> ".." Then
            nSize = nSize + (WFD.nFileSizeHigh * (MAXDWORD + 1)) + WFD.nFileSizeLow
            cCount = cCount + 1
            nRet = SHGetFileInfo(sPath & "\" & sTmp, 0&, shfi, Len(shfi), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
 'Store Files and Folders info for filling ListView
            FI(cCount).sFileName = TrimNull(shfi.szDisplayName)
            FI(cCount).sFileDate = vbGetFileDate$(WFD.ftCreationTime)
            FI(cCount).sFileDescription = TrimNull(shfi.szTypeName)
            FI(cCount).hLargeIcon = shfi.hIcon
            FI(cCount).sFileNameWithExt = sTmp
            nRet = SHGetFileInfo(sPath & "\" & sTmp, 0&, shfi, Len(shfi), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
            FI(cCount).hSmallIcon = shfi.hIcon
            FI(cCount).bIsFolder = (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY)
            FI(cCount).nIcon = shfi.iIcon
            If FI(cCount).bIsFolder Then
               FI(cCount).sFileSize = ""
            Else
               FI(cCount).sFileSize = GetFileSizeStr(WFD.nFileSizeHigh + WFD.nFileSizeLow)
            End If
         End If
     Loop While FindNextFile(hFile, WFD)
     hFile = FindClose(hFile)
     ReDim Preserve FI(cCount)
   End If
   NetEnumFolders = cCount
End Function

Private Function GetFileSizeStr(fsize As Long) As String
  If fsize < 1000 Then
     GetFileSizeStr = Format$(fsize, "###,###,###") & " b"
  Else
     GetFileSizeStr = Format$(fsize / 1024 + 0.5, "###,###,###") & " kb"
  End If
End Function

Public Function TrimNull(startstr As String) As String
   Dim pos As Integer
   pos = InStr(startstr, Chr$(0))
   If pos Then
      TrimNull = Left$(startstr, pos - 1)
      Exit Function
   End If
   TrimNull = startstr
End Function

Private Function vbGetFileDate(CT As FILETIME) As String
    Dim ST As SYSTEMTIME
    Dim r As Long
    Dim ds As Single
    r = FileTimeToSystemTime(CT, ST)
    If r Then
       ds = DateSerial(ST.wYear, ST.wMonth, ST.wDay)
       vbGetFileDate$ = Format$(ds, "Short Date")
    Else
       vbGetFileDate$ = ""
    End If
End Function
