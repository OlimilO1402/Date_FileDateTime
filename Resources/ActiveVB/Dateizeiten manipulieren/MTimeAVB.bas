Attribute VB_Name = "MTimeAVB"
'Dieser Source stammt von http://www.activevb.de
'und kann frei verwendet werden. Für eventuelle Schäden
'wird nicht gehaftet.

'Um Fehler oder Fragen zu klären, nutzen Sie bitte unser Forum.
'Ansonsten viel Spaß und Erfolg mit diesem Source !

Option Explicit

Private Declare Function FileTimeToLocalFileTime Lib "kernel32" ( _
    lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long

Private Declare Function CreateFileW Lib "kernel32.dll" ( _
    ByVal lpFileName As String, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    ByVal lpSecurityAttributes As Long, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Long) As Long
        
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
        
Private Declare Function FileTimeToSystemTime Lib "kernel32" ( _
    lpFileTime As FILETIME, _
    lpSystemTime As SYSTEMTIME) As Long

Private Declare Function SystemTimeToFileTime Lib "kernel32" ( _
    lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long

Private Declare Function LocalFileTimeToFileTime Lib "kernel32" ( _
    lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long

Private Const OPEN_EXISTING   As Long = 3
Private Const GENERIC_READ    As Long = &H80000000
Private Const GENERIC_WRITE   As Long = &H40000000
Private Const OFS_MAXPATHNAME As Long = 128&

Private Type OFSTRUCT
    cBytes         As Byte
    fFixedDisk     As Byte
    nErrCode       As Integer
    Reserved1      As Integer
    Reserved2      As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type

Private Type FILETIME
    dwLowDateTime  As Long
    dwHighDateTime As Long
End Type

Private Type SYSTEMTIME
    wYear          As Integer
    wMonth         As Integer
    wDayOfWeek     As Integer
    wDay           As Integer
    wHour          As Integer
    wMinute        As Integer
    wSecond        As Integer
    wMilliseconds  As Integer
End Type

Private Sub GetAllFileTimes(FileName As String, ct_out As String, lat_out As String, lwt_out As String)
    
    Dim OFS As OFSTRUCT: OFS.cBytes = LenB(OFS)
    
    Dim hFile As Long: hFile = OpenFile(TestFile, 0)
    If hFile <= 0 Then Exit Sub
    
    Dim DTformat As String: DTformat = "dd.mm.yyyy hh:mm:ss"
    
    Dim cTime As FILETIME, lTime As FILETIME, lwTime As FILETIME
    Dim STime As SYSTEMTIME
    
    Call GetFileTime(hFile, cTime, lTime, lwTime)
    
    Call FileTimeToSystemTime(cTime, STime)
    ct_out = Format$(CalcFTime(STime), DTformat)
    
    Call FileTimeToSystemTime(lTime, STime)
    lat_out = Format$(CalcFTime(STime), DTformat)
    
    Call FileTimeToSystemTime(lwTime, STime)
    lwt_out = Format$(CalcFTime(STime), DTformat)
    
    Call CloseHandle(hFile)
End Sub

'Funktion müßte eigentlich heißen SystemTime_ToDate
Private Function CalcFTime(STime As SYSTEMTIME) As Date
    With STime
        
        Dim Da As String: Da = .wDay:        If Len(Da) < 2 Then Da = "0" & Da
        Dim Mo As String: Mo = .wMonth:      If Len(Mo) < 2 Then Mo = "0" & Mo
        Dim Ye As String: Ye = CStr(.wYear)
        
        Dim Datum As String: Datum = Da & "." & Mo & "." & Ye
    
        Dim mm As String: mm = Trim$(CStr(.wMinute))
        If Len(mm) < 2 Then mm = "0" & mm
    
        Dim ss As String: ss = Trim$(CStr(.wSecond))
        If Len(ss) < 2 Then ss = "0" & ss
        
        Dim Zeit As String: Zeit = .wHour & ":" & mm & ":" & ss
        Dim DT As Date: DT = CDate(Datum & " " & Zeit)
        
        CalcFTime = DT
    End With
End Function

Private Function CalcNewfTime(Datum As String) As FILETIME
    Dim SysT As SYSTEMTIME, FT As FILETIME
    Dim FTL As FILETIME
     
    With SysT
        .wDay = CInt(Left$(Datum, 2))
        .wMonth = CInt(Mid$(Datum, 4, 2))
        .wYear = CInt(Mid$(Datum, 7, 4))
        .wHour = CInt(Mid$(Datum, 12, 2))
        .wMinute = CInt(Mid$(Datum, 15, 2))
        .wSecond = CInt(Mid$(Datum, 18, 2))
    End With
    
    Call SystemTimeToFileTime(SysT, FT)
    
    'Update am  18 August 2003:
    'Nun sollten die Fehler mit der Zeitverschiebung verschwunden sein
    Call LocalFileTimeToFileTime(FT, FTL)
    CalcNewfTime = FTL
End Function

Private Function GetRandomDate(Base As Date) As Date
    Dim aa As String
    
    Do
        aa = ""
        aa = CStr(Int(28 * Rnd) + 1) & "." & _
           CStr(Int(12 * Rnd) + 1) & "." & _
           CStr(Int(10 * Rnd) + 1998) & " " & _
           CStr(Int(24 * Rnd)) & ":" & _
           CStr(Int(60 * Rnd)) & ":" & _
           CStr(Int(60 * Rnd))
    Loop While CDate(aa) < Base
    GetRandomDate = CDate(aa)
End Function

Public Function OpenFile(FileName As String, DesiredAccess As Long) As Long
    Dim dwDesiredAccess As Long
    
    If DesiredAccess = 0 Then dwDesiredAccess = GENERIC_READ
    If DesiredAccess = 1 Then dwDesiredAccess = GENERIC_WRITE
    If dwDesiredAccess = 0 Then Exit Function
    
    OpenFile = CreateFile(FileName, dwDesiredAccess, 0, 0, OPEN_EXISTING, 0, 0)
End Function

