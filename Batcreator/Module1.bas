Attribute VB_Name = "spawns"
Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long

Public Const MB_ABORTRETRYIGNORE = &H2&



Sub green(lbl As label)
If lbl.BackColor = &HFF0000 Then
lbl.BackColor = &HC00000
lbl.ForeColor = vbBlack
'Pause 2
'lbl.BackColor = vbBlack
'lbl.ForeColor = &H8000&
Else
lbl.BackColor = &HFF0000
lbl.ForeColor = vbWhite
End If
End Sub
Public Sub Pause(Duration As Long)
    Dim Current As Long
    Current = Timer
    Do Until Timer - Current >= Duration
        DoEvents
    Loop
End Sub
Sub LoadText(txtLoad As textbox, Path As String)
    Dim TextString As String
    On Error Resume Next
    Open Path$ For Input As #1
    TextString$ = Input(LOF(1), #1)
    Close #1
    txtLoad.text = TextString$
End Sub

Sub SaveText(txtSave As textbox, Path As String)
    Dim TextString As String
    On Error Resume Next
    TextString$ = txtSave.text
    Open Path$ For Output As #1
    Print #1, TextString$
    Close #1
End Sub
Sub bla()
Form1.Command8.Visible = True
delfile.Check1.Value = 1

End Sub
