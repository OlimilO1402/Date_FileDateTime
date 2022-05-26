VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   6120
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnOpenFile 
      Caption         =   "Open File"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton BtnOpenFolder 
      Caption         =   "Open Folder"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   1455
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
      TabIndex        =   12
      Top             =   600
      Width           =   3180
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
      TabIndex        =   11
      Top             =   960
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
      TabIndex        =   10
      Top             =   1320
      Width           =   600
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
      TabIndex        =   9
      Top             =   1680
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Creation-Time:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   1020
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Last Access-Time:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   1305
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Last Write-Time:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1155
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "File- Or Foldername:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1410
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Erstelldatum"
      Height          =   195
      Left            =   3360
      TabIndex        =   4
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Letzter Zugriff"
      Height          =   195
      Left            =   3360
      TabIndex        =   3
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Letzte Speicherung"
      Height          =   195
      Left            =   3360
      TabIndex        =   2
      Top             =   1680
      Width           =   1380
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
End Sub

Private Sub BtnOpenFolder_Click()
    
    Dim OFD As OpenFolderDialog: Set OFD = New OpenFolderDialog
    If OFD.ShowDialog = vbCancel Then Exit Sub
    
    Dim tCreation As Date   ' Erstellt am
    Dim tLastAccess As Date ' Letzter Zugriff
    Dim tLastWrite As Date  ' Letzte Änderung
    
    ' Ordner
    Dim sFolder As String
    'sFolder = "C:\TestDir\"
    sFolder = OFD.Folder
    
    LblPathFileName.Caption = sFolder
    
    ' Zeitangaben lesen
    If MTime.ReadFolderTime(sFolder, tCreation, tLastAccess, tLastWrite) Then
         
         LblCreationTime.Caption = tCreation
         LblLAccessTime.Caption = tLastAccess
         LblLWriteTime.Caption = tLastWrite
         
         ' Erstellungsdatum ändern
         'tCreation = CDate("29.08.2002 17:35:41")
         
         ' Datum "Letzter Zugriff" ändern
         'tLastAccess = CDate("29.08.2002 17:35:41")
         
         ' Datum "Letzter Änderung" ändern
         'tLastWrite = CDate("29.08.2002 17:35:41")
         
         ' Zeitangaben setzen
         'WriteFileTime Datei, tCreation, tLastAccess, tLastWrite
    End If
    
End Sub
