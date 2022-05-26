Attribute VB_Name = "MTime"
Option Explicit

 
' alle benötigten API-Funktionen
Private Declare Function CreateFile Lib "kernel32" _
  Alias "CreateFileA" ( _
  ByVal lpFilename As String, _
  ByVal dwDesiredAccess As Long, _
  ByVal dwShareMode As Long, _
  lpSecurityAttributes As Any, _
  ByVal dwCreationDisposition As Long, _
  ByVal dwFlagsAndAttributes As Long, _
  ByVal hTemplateFile As Long) As Long
 
Private Declare Function GetFileTime Lib "kernel32" ( _
  ByVal hFile As Long, _
  lpCreationTime As FILETIME, _
  lpLastAccessTime As FILETIME, _
  lpLastWriteTime As FILETIME) As Long
 
Private Declare Function SetFileTime Lib "kernel32" ( _
  ByVal hFile As Long, _
  lpCreationTime As FILETIME, _
  lpLastAccessTime As FILETIME, _
  lpLastWriteTime As FILETIME) As Long
 
Private Declare Function CloseHandle Lib "kernel32" ( _
  ByVal hObject As Long) As Long
 
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" ( _
  lpFileTime As FILETIME, _
  lpLocalFileTime As FILETIME) As Long
 
Private Declare Function FileTimeToSystemTime Lib "kernel32" ( _
  lpFileTime As FILETIME, _
  lpSystemTime As SYSTEMTIME) As Long
 
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" ( _
  lpLocalFileTime As FILETIME, _
  lpFileTime As FILETIME) As Long
 
Private Declare Function SystemTimeToFileTime Lib "kernel32" ( _
  lpSystemTime As SYSTEMTIME, _
  lpFileTime As FILETIME) As Long
 
Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type
 
Private Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type
 
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const OPEN_EXISTING = 3
Private Const FILE_FLAG_BACKUP_SEMANTICS = &H2000000


Public Function ReadFolderTime(ByVal sFolder As String, _
  tCreation As Date, tLastAccess As Date, _
  tLastWrite As Date) As Boolean
 
  ' Datum/Zeitwert eines Ordners ermitteln
  Dim fHandle As Long
 
  Dim ftCreation As FILETIME
  Dim ftLastAccess As FILETIME
  Dim ftLastWrite As FILETIME
  Dim LocalFileTime As FILETIME
  Dim LocalSystemTime As SYSTEMTIME
 
  ReadFolderTime = False
 
  ' ggf. abschließenden Backslash hinzufügen
  If Right$(sFolder, 1) <> "\" Then sFolder = sFolder & "\"
 
  ' Verzeichnishandle ermitteln
  fHandle = CreateFile(sFolder, GENERIC_READ Or GENERIC_WRITE, _
    FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, _
    OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0&)
 
  If fHandle <> -1 Then
    ' Zeitinformationen auslesen
    If GetFileTime(fHandle, ftCreation, ftLastAccess, _
      ftLastWrite) <> 0 Then
 
      ' Erstellungsdatum
      FileTimeToLocalFileTime ftCreation, LocalFileTime
      FileTimeToSystemTime LocalFileTime, LocalSystemTime
      With LocalSystemTime
        tCreation = CDate(Format$(.wDay) & "." & _
          Format$(.wMonth) & "." & Format$(.wYear) & " " & _
          Format$(.wHour) & ":" & Format$(.wMinute, "00") & _
          ":" & Format$(.wSecond, "00"))
      End With
 
      ' Letzter Zugriff
      FileTimeToLocalFileTime ftLastAccess, LocalFileTime
      FileTimeToSystemTime LocalFileTime, LocalSystemTime
      With LocalSystemTime
        tLastAccess = CDate(Format$(.wDay) & "." & _
          Format$(.wMonth) & "." & Format$(.wYear) & " " & _
          Format$(.wHour) & ":" & Format$(.wMinute, "00") & _
          ":" & Format$(.wSecond, "00"))
      End With
 
      ' Letzte Änderung
      FileTimeToLocalFileTime ftLastWrite, LocalFileTime
      FileTimeToSystemTime LocalFileTime, LocalSystemTime
      With LocalSystemTime
        tLastWrite = CDate(Format$(.wDay) & "." & _
          Format$(.wMonth) & "." & Format$(.wYear) & " " & _
          Format$(.wHour) & ":" & Format$(.wMinute, "00") & _
          ":" & Format$(.wSecond, "00"))
      End With
 
      ReadFolderTime = True
    End If
 
    ' Verzeichnishandle schließen
    CloseHandle fHandle
  End If
End Function

Public Function WriteFolderTime(ByVal sFolder As String, _
  ByVal tCreation As Date, ByVal tLastAccess As Date, _
  ByVal tLastWrite As Date) As Boolean
 
  ' Datum/Zeitwert eines Ordners ändern
  Dim fHandle As Long
  Dim ftCreation As FILETIME
  Dim ftLastAccess As FILETIME
  Dim ftLastWrite As FILETIME
  Dim LocalFileTime As FILETIME
  Dim LocalSystemTime As SYSTEMTIME
 
  WriteFolderTime = False
 
  ' ggf. abschließenden Backslash hinzufügen
  If Right$(sFolder, 1) <> "\" Then sFolder = sFolder & "\"
 
  ' Verzeichnishandle ermitteln
  fHandle = CreateFile(sFolder, GENERIC_READ Or GENERIC_WRITE, _
    FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, _
    OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0&)
 
  If fHandle <> -1 Then
    ' Erstellungsdatum
    With LocalSystemTime
      .wDay = Day(tCreation)
      .wDayOfWeek = Weekday(tCreation)
      .wMonth = Month(tCreation)
      .wYear = Year(tCreation)
      .wHour = Hour(tCreation)
      .wMinute = Minute(tCreation)
      .wSecond = Second(tCreation)
    End With
    SystemTimeToFileTime LocalSystemTime, LocalFileTime
    LocalFileTimeToFileTime LocalFileTime, ftCreation
 
    ' Letzter Zugriff
    With LocalSystemTime
      .wDay = Day(tLastAccess)
      .wDayOfWeek = Weekday(tLastAccess)
      .wMonth = Month(tLastAccess)
      .wYear = Year(tLastAccess)
      .wHour = Hour(tLastAccess)
      .wMinute = Minute(tLastAccess)
      .wSecond = Second(tLastAccess)
    End With
    SystemTimeToFileTime LocalSystemTime, LocalFileTime
    LocalFileTimeToFileTime LocalFileTime, ftLastAccess
 
    ' Letzte Änderung
    With LocalSystemTime
      .wDay = Day(tLastWrite)
      .wDayOfWeek = Weekday(tLastWrite)
      .wMonth = Month(tLastWrite)
      .wYear = Year(tLastWrite)
      .wHour = Hour(tLastWrite)
      .wMinute = Minute(tLastWrite)
      .wSecond = Second(tLastWrite)
    End With
    SystemTimeToFileTime LocalSystemTime, LocalFileTime
    LocalFileTimeToFileTime LocalFileTime, ftLastWrite
 
    ' Datumswerte neu setzen
    If SetFileTime(fHandle, ftCreation, ftLastAccess, _
      ftLastWrite) <> 0 Then
 
      WriteFolderTime = True
    End If
 
    ' Verzeichnishandle schließen
    CloseHandle fHandle
  End If
End Function

