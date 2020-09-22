VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H80000006&
   Caption         =   "Easy Batch creator Version 1.00"
   ClientHeight    =   6135
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9735
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0442
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3360
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":F6300
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":F6754
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":F6BA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":F6FFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":F7450
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":F78A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":F7CF8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   1005
      ButtonWidth     =   3254
      ButtonHeight    =   953
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Disclaimer"
            Object.ToolTipText     =   "Disclaimer"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Create Batch"
            Object.ToolTipText     =   "Create Batch"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Help "
            Object.ToolTipText     =   "Help "
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Upgrade"
            Object.ToolTipText     =   "Upgrade"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Words From The Creator"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
      MousePointer    =   10
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6840
      Top             =   0
   End
   Begin VB.Menu crap 
      Caption         =   "Options"
      Visible         =   0   'False
      Begin VB.Menu ab 
         Caption         =   "About"
      End
      Begin VB.Menu dis 
         Caption         =   "Disclaimer"
      End
      Begin VB.Menu batcreat 
         Caption         =   "Bat Creator"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim one




Private Sub ab_Click()
frmAbout.Show
Unload frmSplash1
Unload frmSplash
Unload delfile
Unload Form1
End Sub

Private Sub batcreat_Click()
Form1.Show
Unload frmSplash
Unload frmSplash1
Unload frmSplash
Unload frmAbout
Unload delfile

End Sub

Private Sub dis_Click()
frmSplash.Show
Unload frmSplash1
Unload frmAbout
Unload delfile
Unload Form1
End Sub

Private Sub Timer1_Timer()
If one = 1 Then
SendKeys "{enter}"
Timer1.Enabled = False
ElseIf one = 2 Then
'green Label1


End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
    Select Case Button.Index
        Case "1" ' disclaimer
          dis_Click
        Case "2" 'batch creator
           batcreat_Click
        Case "3" 'help
           Frmhelp.Show
        Unload frmSplash1
        Unload frmSplash
        Unload frmAbout
        Unload delfile
        Unload Form1
    
        Case "4" 'upgrade
        delfile.Show
        Unload frmSplash
        Unload frmAbout
        Unload frmSplash1
        Unload Form1
     
    Case "5" ' words from spawn
    frmSplash1.Show
    'frmSplash1.Height = 6090
   ' frmSplash1.Label3.Visible = True
   ' frmSplash1.Label1.Visible = False
    End Select
    Exit Sub
crap:     MsgBox "The file Does not Exist", vbCritical, "invalid file"
End Sub

