VERSION 5.00
Begin VB.Form delfile 
   BackColor       =   &H00800000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Upgrade "
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4350
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FF0000&
      Caption         =   "Save Don't show this again"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   1440
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00800000&
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   480
      TabIndex        =   4
      Text            =   "Virtual123"
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF0000&
      Caption         =   "OK"
      Height          =   375
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   480
      PasswordChar    =   "i"
      TabIndex        =   1
      Text            =   "A1S2"
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "OzHandicraft BT"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"delfile.frx":0000
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "delfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
bla
End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Text1_Change()
If Text1.text = "test" Then
Check1.Visible = True
Label4.Caption = "Status: Correct"
Else
Label4.Caption = "Sorry Wrong Pass"
End If
End Sub

Private Sub Text2_Change()
Text1_Change
End Sub

