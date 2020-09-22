VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2265
   ClientLeft      =   9270
   ClientTop       =   1260
   ClientWidth     =   2580
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   2580
   Begin VB.CommandButton Command3 
      Caption         =   "Words From Spawn"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Help "
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "About"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   1935
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   2055
   End
   Begin VB.Menu crap 
      Caption         =   "test"
      Visible         =   0   'False
      Begin VB.Menu dd 
         Caption         =   "&Undo"
         Enabled         =   0   'False
      End
      Begin VB.Menu space 
         Caption         =   "-"
      End
      Begin VB.Menu cut1 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu copy1 
         Caption         =   "&Copy"
      End
      Begin VB.Menu paste1 
         Caption         =   "&Paste"
      End
      Begin VB.Menu delete1 
         Caption         =   "&Delete"
      End
      Begin VB.Menu space1 
         Caption         =   "-"
      End
      Begin VB.Menu sa 
         Caption         =   "Select &All"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmAbout.Show

End Sub

Private Sub Command2_Click()
Frmhelp.Show
End Sub

Private Sub Command3_Click()
frmSplash1.Show
frmSplash1.Height = 6090
frmSplash1.Label3.Visible = True
frmSplash1.Label1.Visible = False
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Button = 2 Then Me.PopupMenu crap, vbPopupMenuRightButton
End Sub

