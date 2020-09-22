VERSION 5.00
Begin VB.Form Frmhelp 
   BackColor       =   &H00000000&
   Caption         =   "Help Area"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7500
   LinkTopic       =   "Form3"
   ScaleHeight     =   5655
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00BC5F2C&
      Height          =   4695
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   7335
      Begin VB.Frame Frame2 
         BackColor       =   &H00404040&
         Height          =   4695
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   7335
         Begin VB.TextBox Text2 
            Height          =   2655
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   7
            Text            =   "form3.frx":0000
            Top             =   480
            Width           =   2895
         End
         Begin VB.TextBox Text1 
            Height          =   2655
            Left            =   3120
            MultiLine       =   -1  'True
            TabIndex        =   6
            Text            =   "form3.frx":004B
            Top             =   480
            Width           =   3975
         End
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00000000&
         ForeColor       =   &H80000009&
         Height          =   4155
         ItemData        =   "form3.frx":0090
         Left            =   120
         List            =   "form3.frx":00B2
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   4455
         Left            =   2280
         TabIndex        =   4
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Samples"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command list"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C00000&
      Height          =   615
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "Frmhelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Frame1.Visible = True
Frame2.Visible = False
End Sub

Private Sub Command2_Click()
Frame1.Visible = False
Frame2.Visible = True
End Sub

Private Sub List1_Click()
If List1.ListIndex = 0 Then ' Move

Label1.Caption = "Location file is at now          Location you want file" + Chr(13) + "Move c:\windows\desktop\filename c:\windows\filename " + Chr(13)
ElseIf List1.ListIndex = 1 Then ' Copy
Label1.Caption = "Location file is at now          Location you want copy file" + Chr(13) + "Copy c:\windows\desktop\filename c:\windows\filename " + Chr(13)
ElseIf List1.ListIndex = 2 Then ' Pause
Label1.Caption = "Pause this will stop that batch file"
ElseIf List1.ListIndex = 3 Then ' Echo on
Label1.Caption = "Echo on this will make it so you can see what your batch is doing"
ElseIf List1.ListIndex = 4 Then 'echo off
Label1.Caption = "Echo on this will make it so you can't  see what your batch is doing"
ElseIf List1.ListIndex = 5 Then ' del file
Label1.Caption = "Del c:\windows\desktop\filename.txt or any other"
ElseIf List1.ListIndex = 6 Then ' Make Dirctory
Label1.Caption = "MD c:\windows\desktop\name of dirc" + Chr(13) + "This will Create the dirctory" + Chr(13)
ElseIf List1.ListIndex = 7 Then ' Remove Dirctory
Label1.Caption = "RD c:\windows\desktop\name of dirc" + Chr(13) + "This will Delete the dirctory" + Chr(13)
ElseIf List1.ListIndex = 8 Then ' Run Exe
Label1.Caption = "Run c:\windows\desktop\filenam.exe"
ElseIf List1.ListIndex = 9 Then ' exit batch
Label1.Caption = "Exit batch file"
End If
End Sub
