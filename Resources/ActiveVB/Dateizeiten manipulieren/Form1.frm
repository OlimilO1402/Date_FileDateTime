VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "www.activevb.de"
   ClientHeight    =   5385
   ClientLeft      =   1680
   ClientTop       =   1530
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   5385
   ScaleWidth      =   4440
   Begin VB.Frame Frame2 
      Caption         =   "Neue Dateizeiten"
      Height          =   2175
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   3735
      Begin VB.CommandButton Command2 
         Caption         =   "Zuweisen"
         Height          =   375
         Left            =   2520
         TabIndex        =   15
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Mischen"
         Height          =   375
         Left            =   1440
         TabIndex        =   8
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Label9"
         Height          =   255
         Left            =   1560
         TabIndex        =   11
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Label8"
         Height          =   255
         Left            =   1560
         TabIndex        =   10
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Label7"
         Height          =   255
         Left            =   1560
         TabIndex        =   9
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Letzte Änderung"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Letzter Zugriff"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label10 
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Erstellung"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Alte Dateizeiten"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Label1"
         Height          =   255
         Left            =   1560
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Label2"
         Height          =   255
         Left            =   1560
         TabIndex        =   2
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Label3"
         Height          =   255
         Left            =   1560
         TabIndex        =   1
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Erstellung"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Letzter Zugriff"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Letzte Änderung"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TestFile As String

Private Sub Form_Load()
    Dim FN As Integer
    
    TestFile = App.Path & "\Test.txt"
    
    'Evt. Test-Datei erstellen
    If Len(Dir$(TestFile, vbNormal)) = 0 Then
        FN = FreeFile
        Open TestFile For Output As #FN
        Close FN
    End If
    
    Call GetFTime
    Command2.Enabled = False
End Sub

Private Sub Command1_Click()
    Dim DTformat As String: DTformat = "dd.mm.yyyy hh:mm:ss" 'is ohnehin default
    Dim Base As Date
    
    Base = GetRandomDate(Base)
    Label7.Caption = Base 'Format$(Base, DTformat)
    
    Base = GetRandomDate(Base)
    Label8.Caption = Base 'Format$(Base, DTformat)
    
    Base = GetRandomDate(Base)
    Label9.Caption = Base 'Format$(Base, DTformat)
    
    Command2.Enabled = True
    
End Sub

Private Sub Command2_Click()
    Dim hFile As Long, DTformat As String
    Dim STime As SYSTEMTIME
    Dim OFS As OFSTRUCT
    Dim cTime As FILETIME
    Dim lTime As FILETIME
    Dim lwTime As FILETIME
    
    'Änderung 08.03.2003:
    'Die Dateiattribute werden nun zwischenzeitlich geändert,
    'um Fehlern vorzubeugen
    Dim fAttr As VbFileAttribute
    
    fAttr = GetAttr(TestFile)
    SetAttr TestFile, vbNormal
  
    OFS.cBytes = Len(OFS)
    ' Update von Kai: OpenFile gilt als "veraltet"
    'hFile = OpenFile(TestFile, OFS, OF_WRITE)
    hFile = OpenFile(TestFile, 1)
    
    If hFile > 0 Then
        DTformat = "dd.mm.yyyy hh:mm:ss"
        Call GetFileTime(hFile, cTime, lTime, lwTime)
        
        Call FileTimeToLocalFileTime(cTime, cTime)
        Call FileTimeToSystemTime(cTime, STime)
        Label1.Caption = Format$(CalcFTime(STime), DTformat)
        
        Call FileTimeToLocalFileTime(lTime, lTime)
        Call FileTimeToSystemTime(lTime, STime)
        Label2.Caption = Format$(CalcFTime(STime), DTformat)
        
        Call FileTimeToLocalFileTime(lwTime, lwTime)
        Call FileTimeToSystemTime(lwTime, STime)
        Label3.Caption = Format$(CalcFTime(STime), DTformat)
        
        Call CloseHandle(hFile)
    End If
    
    Call GetFTime
    Command2.Enabled = False
    
    SetAttr TestFile, fAttr
End Sub
