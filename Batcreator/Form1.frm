VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spx Batch file creator"
   ClientHeight    =   5505
   ClientLeft      =   2220
   ClientTop       =   4275
   ClientWidth     =   9120
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   9120
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4560
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0442
            Key             =   "echo text"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0896
            Key             =   "Creat Dir"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0CEA
            Key             =   "Deldir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":113E
            Key             =   "pause"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1592
            Key             =   "Del File"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":19E6
            Key             =   "run exe"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1E3A
            Key             =   "Echo On"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":228E
            Key             =   "move"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":26E2
            Key             =   "Exit bat"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2B36
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2F8A
            Key             =   "Echo off"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   741
      ButtonWidth     =   609
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Echo Text"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Create Dir"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Del Dirc"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Pause"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Delete File"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Run exe"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Echo On"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Move "
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   6480
      TabIndex        =   31
      Top             =   4680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   375
      Left            =   6480
      TabIndex        =   30
      Top             =   4200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FF0000&
      Caption         =   "Create Batch"
      Height          =   855
      Left            =   3840
      TabIndex        =   26
      Top             =   4080
      Visible         =   0   'False
      Width           =   2535
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   120
         TabIndex        =   28
         Text            =   "test.bat"
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Create"
         Height          =   255
         Left            =   1560
         TabIndex        =   27
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label14 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "Name of batch:"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C00000&
      Caption         =   "Open txt"
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   3840
      TabIndex        =   21
      Top             =   4080
      Visible         =   0   'False
      Width           =   2535
      Begin VB.CommandButton Command5 
         Caption         =   "Open"
         Height          =   255
         Left            =   1560
         TabIndex        =   24
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   120
         TabIndex        =   22
         Text            =   "test.txt"
         Top             =   480
         Width           =   1335
      End
      Begin VB.Image Image1 
         Height          =   345
         Left            =   1680
         Picture         =   "Form1.frx":33DE
         Stretch         =   -1  'True
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label13 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "Name of file:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C00000&
      Caption         =   "Save Txt "
      Height          =   855
      Left            =   3840
      TabIndex        =   18
      Top             =   4080
      Visible         =   0   'False
      Width           =   2535
      Begin VB.CommandButton Command6 
         Caption         =   "Save"
         Height          =   255
         Left            =   1560
         TabIndex        =   25
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FF0000&
         Height          =   285
         Left            =   240
         TabIndex        =   19
         Text            =   "test.txt"
         Top             =   480
         Width           =   1215
      End
      Begin VB.Image Image2 
         Height          =   345
         Left            =   1680
         Picture         =   "Form1.frx":3820
         Stretch         =   -1  'True
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label11 
         BackColor       =   &H00800000&
         Caption         =   "Name of file:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSComDlg.CommonDialog CmDialog1 
      Left            =   0
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00008000&
      Caption         =   "additem"
      Height          =   255
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00400000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   240
      TabIndex        =   15
      Top             =   3720
      Width           =   6135
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00800000&
      ForeColor       =   &H80000005&
      Height          =   2790
      ItemData        =   "Form1.frx":3C62
      Left            =   240
      List            =   "Form1.frx":3C64
      TabIndex        =   14
      ToolTipText     =   "Double click to remove work you did "
      Top             =   480
      Width           =   6735
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FF0000&
      Caption         =   "Create batch"
      Height          =   255
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF0000&
      Caption         =   "Open txt"
      Height          =   255
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF0000&
      Caption         =   "Save txt"
      Height          =   255
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4560
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      Caption         =   "-Options-"
      ForeColor       =   &H00000000&
      Height          =   4215
      Left            =   7320
      TabIndex        =   0
      Top             =   480
      Width           =   1695
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-Exit bat-"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         MousePointer    =   2  'Cross
         TabIndex        =   10
         ToolTipText     =   "Make Batch auto exit when finished what it was doing"
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-Remove Dirc-"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         MousePointer    =   2  'Cross
         TabIndex        =   9
         ToolTipText     =   "Remove Folder "
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-Create Dirc-"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         MousePointer    =   2  'Cross
         TabIndex        =   8
         ToolTipText     =   "create a folder"
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-Pause-"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         MousePointer    =   2  'Cross
         TabIndex        =   7
         ToolTipText     =   "this is to pause the batch and press enter after"
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-Del file-"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         MousePointer    =   2  'Cross
         TabIndex        =   6
         ToolTipText     =   "this is to delete a file "
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-Echo Off-"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         MousePointer    =   2  'Cross
         TabIndex        =   5
         ToolTipText     =   "this is to make it so you can't see what the batch is doing"
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-Echo On-"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         MousePointer    =   2  'Cross
         TabIndex        =   4
         ToolTipText     =   "this is so you can see what the batch is doing "
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-Run Exe-"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         MousePointer    =   2  'Cross
         TabIndex        =   3
         ToolTipText     =   "Run software like auto open the software"
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-Copy-"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         MousePointer    =   2  'Cross
         TabIndex        =   2
         ToolTipText     =   "Copy is to copy a file to a different location"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-Move-"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         MousePointer    =   2  'Cross
         TabIndex        =   1
         ToolTipText     =   "Move is to move a file to a different location"
         Top             =   2400
         Width           =   1095
      End
   End
   Begin VB.Label Label12 
      BackColor       =   &H00000000&
      Caption         =   "Add Code Here and press Enter on you keyboard:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      MousePointer    =   3  'I-Beam
      TabIndex        =   17
      Top             =   3480
      Width           =   6135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   4440
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Frame3.Visible = False
Frame2.Visible = True
Frame4.Visible = False

'On Error Resume Next
'CmDialog1.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)"
'CmDialog1.FilterIndex = 1
'CmDialog1.Action = 2
'Open CmDialog1.FileName For Output As #1
'Print #1, List1.List(ListIndex)
'Close #1

End Sub

Private Sub Command2_Click()
Frame3.Visible = True
Frame2.Visible = False
Frame4.Visible = False
'CmDialog1.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)"
'CmDialog1.FilterIndex = 1
'CmDialog1.Action = 2
'Open CmDialog1.FileName For Output As #1
'Open CmDialog1.FileName For Input As 1
'Text1.text = Input$(LOF(1), 1)
'Close 1
End Sub

Private Sub Command3_Click()
Frame4.Visible = True
Frame3.Visible = False
Frame2.Visible = False
'MsgBox "not yet programmed"
 
End Sub

Private Sub Command4_Click()
If Text1.text = Text1.text Then
List1.AddItem Text1.text
End If
End Sub



Private Sub Command5_Click()
Call LoadListBox("c:\windows\desktop\" & (Text4.text), List1)
End Sub

Private Sub Command6_Click()
Call SaveListBox("c:\windows\desktop\" & (Text2.text), List1)
End Sub



Private Sub Command7_Click()
Call SaveListBox("c:\windows\desktop\" & Text3.text, List1)
End Sub

Private Sub Form_Load()

End Sub

Private Sub Image1_Click()
MsgBox "Reasons you might get an error" + Chr(13) + "1. you for got to put txt on the back of the the name of file ." + Chr(13) + "2.nothing in the text file may not save." + Chr(13) + "---------------------------------------------------" + Chr(13) + "Location the place is saved " + Chr(13) + "c:\windows\desktop\ " + Chr(13) + "you should beable to see it after you save the file" + Chr(13) + "Will be improving on the next version" + Chr(13)
End Sub

Private Sub Image2_Click()
MsgBox "Reasons you might get an error" + Chr(13) + "1. you for got to put txt on the back of the the name of file ." + Chr(13) + "2.nothing in the text file may not save." + Chr(13) + "---------------------------------------------------" + Chr(13) + "Location the place is saved " + Chr(13) + "c:\windows\desktop\ " + Chr(13) + "you should beable to see it after you save the file" + Chr(13) + "Will be improving on the next version" + Chr(13)
End Sub

Private Sub Label1_Click()
Text1.text = "Move " & " program name.exe,txt,ptf,bmp" & " C:\WINDOWS\Desktop"
green Label1
End Sub

Private Sub Label10_Click()
Text1.text = "@exit"
End Sub



Private Sub Label12_DblClick()
Form2.Show
End Sub

Private Sub Label2_Click() 'Copy
Text1.text = "Copy " & "program name.exe,txt,ptf,bmp" & " C:\WINDOWS\Desktop\loctation here.name file"
green Label2
End Sub

Private Sub Label3_Click()
Text1.text = "Run C:\Windows\bla.exe"
green Label3
End Sub

Private Sub Label4_Click()
Text1.text = "@Echo on"
green Label4
End Sub

Private Sub Label5_Click()
Text1.text = "@Echo off"
green Label5
End Sub

Private Sub Label6_Click()
Text1.text = "DEL location of file to delete"
green Label6
End Sub

Private Sub Label7_Click()

Text1.text = "pause"
green Label7
End Sub

Private Sub Label8_Click()
Text1.text = "MD File name"
green Label8
End Sub
Private Sub Label9_Click()
Text1.text = "RD file name"
green Label9
End Sub
Private Sub List1_DblClick()
List1.RemoveItem List1.ListIndex
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call Cmdsend_Click
End Sub
Private Sub Cmdsend_Click()
If Text1.text = Text1.text Then
List1.AddItem Text1.text
Text1.text = ""

End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'If person updates to version 2.00
On Error Resume Next
    Select Case Button.Index
        Case "1" ' Echo Text
     
        Case "2" 'Create Dirc
           
        Case "3" 'Remove Dirc
        Text1.text = "RD file name"
        Case "4" 'Pause
        
        Case "5" 'Delete File
        Text1.text = "DEL location of file to delete"
        Case "6" 'Run exe
        Text1.text = "Run C:\Windows\bla.exe"
        Case "7" 'Echo On
         Text1.text = "@Echo on"
        Case "8" 'Move
        Text1.text = "Move " & " program name.exe,txt,ptf,bmp" & " C:\WINDOWS\Desktop"
        Case "9" 'Exit Batch
        
        Case "10" 'Copy
        Text1.text = "Copy " & "program name.exe,txt,ptf,bmp" & " C:\WINDOWS\Desktop\loctation here.name file"
        
        Case "11" 'Echo off
        Text1.text = "@Echo off"
    End Select
    Exit Sub
End Sub
