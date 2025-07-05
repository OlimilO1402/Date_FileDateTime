VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "FileDateTime"
   ClientHeight    =   2175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10500
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manuell
   ScaleHeight     =   2175
   ScaleWidth      =   10500
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnWriteBackAllFileDates 
      Caption         =   "Write All Dates"
      Height          =   375
      Left            =   3240
      TabIndex        =   13
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton BtnOpenFolder 
      Caption         =   "Open Folder"
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton BtnOpenFile 
      Caption         =   "Open File"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Letzte Speicherung (=Änderungsdatum)"
      Height          =   255
      Left            =   3960
      TabIndex        =   12
      Top             =   1680
      Width           =   3465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Letzter Zugriff"
      Height          =   255
      Left            =   3960
      TabIndex        =   11
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Erstelldatum"
      Height          =   255
      Left            =   3960
      TabIndex        =   10
      Top             =   960
      Width           =   1080
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "File- Or Foldername:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   1785
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Last Write-Time:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   1410
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Last Access-Time:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Creation-Time:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   1275
   End
   Begin VB.Label LblLWriteTime 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "            "
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Top             =   1680
      Width           =   750
   End
   Begin VB.Label LblLAccessTime 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "            "
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Top             =   1320
      Width           =   750
   End
   Begin VB.Label LblCreationTime 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "            "
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Top             =   960
      Width           =   750
   End
   Begin VB.Label LblPathFileName 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "                                                                      "
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Top             =   600
      Width           =   4230
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_PDT As PFNDateTime

Private Sub Form_Load()
    MTime.Init
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not Data.GetFormat(ClipBoardConstants.vbCFFiles) Then Exit Sub
    If Data.Files.Count = 0 Then Exit Sub
    Dim PFN As PathFileName: Set PFN = MNew.PathFileName(Data.Files.Item(1))
    Set m_PDT = MNew.PFNDateTime(PFN)
    m_PDT.CClose
    UpdateView
End Sub

Private Sub UpdateView()
    With m_PDT
        LblPathFileName.Caption = .PathFileName.Value
        LblCreationTime.Caption = .CreationTime
        LblLAccessTime.Caption = .LastAccessTime
        LblLWriteTime.Caption = .LastWriteTime
    End With
End Sub

Private Sub BtnOpenFile_Click()
    Dim OFD As OpenFileDialog: Set OFD = New OpenFileDialog
    If OFD.ShowDialog(Me) = vbCancel Then Exit Sub
    Dim PFN As PathFileName: Set PFN = MNew.PathFileName(OFD.FileName)
    Set m_PDT = MNew.PFNDateTime(PFN)
    m_PDT.CClose
    UpdateView
End Sub

Private Sub BtnOpenFolder_Click()
    Dim OFD As OpenFolderDialog: Set OFD = New OpenFolderDialog
    If OFD.ShowDialog(Me.hwnd) = vbCancel Then Exit Sub
    Dim PFN As PathFileName: Set PFN = MNew.PathFileName(OFD.Folder)
    Set m_PDT = MNew.PFNDateTime(PFN)
    UpdateView
End Sub

Private Sub BtnWriteBackAllFileDates_Click()
    Dim s As String
    s = Trim(LblCreationTime.Caption):   If Len(s) <> 0 Then m_PDT.CreationTime = CDate(s)
    s = Trim(LblLAccessTime.Caption):    If Len(s) <> 0 Then m_PDT.LastAccessTime = CDate(s)
    s = Trim(LblLWriteTime.Caption):     If Len(s) <> 0 Then m_PDT.LastWriteTime = CDate(s)
    UpdateView
End Sub

Private Sub LblCreationTime_DblClick()
    Dim s As String: s = Trim(LblCreationTime.Caption)
    If Len(s) = 0 Then
        MsgBox "Nothing to edit, open a file or folder first!"
        Exit Sub
    End If
    s = InputBox("Edit Creation-Time:", "Edit DateTime-Value", s)
    If StrPtr(s) = 0 Then Exit Sub
    m_PDT.CreationTime = CDate(s)
    UpdateView
End Sub

Private Sub LblLAccessTime_DblClick()
    Dim s As String: s = Trim(LblLAccessTime.Caption)
    If Len(s) = 0 Then
        MsgBox "Nothing to edit, open a file or folder first!"
        Exit Sub
    End If
    s = InputBox("Edit Last Access-Time:", "Edit DateTime-Value", s)
    If StrPtr(s) = 0 Then Exit Sub
    m_PDT.LastAccessTime = CDate(s)
    UpdateView
End Sub

Private Sub LblLWriteTime_DblClick()
    Dim s As String: s = Trim(LblLWriteTime.Caption)
    If Len(s) = 0 Then
        MsgBox "Nothing to edit, open a file or folder first!"
        Exit Sub
    End If
    s = InputBox("Edit Last Write-Time:", "Edit DateTime-Value", s)
    If StrPtr(s) = 0 Then Exit Sub
    m_PDT.LastWriteTime = CDate(s)
    UpdateView
End Sub
