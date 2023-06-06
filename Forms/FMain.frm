VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "FileDateTime"
   ClientHeight    =   2775
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   ScaleHeight     =   2775
   ScaleWidth      =   10500
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnOpenFolder 
      Caption         =   "Open Folder"
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton BtnOpenFile 
      Caption         =   "Open File"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Letzte Speicherung (=Änderungsdatum)"
      Height          =   195
      Left            =   3360
      TabIndex        =   12
      Top             =   1680
      Width           =   2805
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Letzter Zugriff"
      Height          =   195
      Left            =   3360
      TabIndex        =   11
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Erstelldatum"
      Height          =   195
      Left            =   3360
      TabIndex        =   10
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "File- Or Foldername:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   1410
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Last Write-Time:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   1155
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Last Access-Time:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Creation-Time:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   1020
   End
   Begin VB.Label LblLWriteTime 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "            "
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   1680
      Width           =   600
   End
   Begin VB.Label LblLAccessTime 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "            "
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   1320
      Width           =   600
   End
   Begin VB.Label LblCreationTime 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "            "
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   960
      Width           =   600
   End
   Begin VB.Label LblPathFileName 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "                                                                      "
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   600
      Width           =   3180
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
    UpdateView
End Sub

Private Sub BtnOpenFolder_Click()
    Dim OFD As OpenFolderDialog: Set OFD = New OpenFolderDialog
    If OFD.ShowDialog(Me.hwnd) = vbCancel Then Exit Sub
    Dim PFN As PathFileName: Set PFN = MNew.PathFileName(OFD.Folder)
    Set m_PDT = MNew.PFNDateTime(PFN)
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
