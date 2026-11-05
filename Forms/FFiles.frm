VERSION 5.00
Begin VB.Form FFiles 
   Caption         =   "Form1"
   ClientHeight    =   8025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10020
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   10020
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnSavePFNDateTime 
      Caption         =   "Save Datetimes"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.ListBox LBFiles 
      Height          =   7035
      Left            =   0
      OLEDragMode     =   1  'Automatisch
      OLEDropMode     =   1  'Manuell
      TabIndex        =   0
      Top             =   480
      Width           =   9615
   End
End
Attribute VB_Name = "FFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Files As Collection 'Of PFNDateTime

Private Sub LBFiles_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Data.GetFormat(ClipBoardConstants.vbCFFiles) Then
        Set m_Files = New Collection
        Dim i As Long ', c as Long:  = data.Files.Count
        Dim s As String
        For i = 1 To Data.Files.Count
            s = Data.Files(i)
            Dim pdt As PFNDateTime: Set pdt = MNew.PFNDateTime(MNew.PathFileName(s))
            m_Files.Add pdt, pdt.PathFileName.Value
        Next
    End If
    UpdateView
End Sub

Sub UpdateView()
    LBFiles.Clear
    Dim pdt As PFNDateTime
    For Each pdt In m_Files
        LBFiles.AddItem pdt.CreationTime & " | " & pdt.LastAccessTime & " | " & pdt.LastWriteTime & " | " & pdt.PathFileName.Length & " | " & pdt.PathFileName.FileName
    Next
End Sub

