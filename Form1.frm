VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mouse Record"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   2415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   2415
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2640
      Top             =   2160
   End
   Begin VB.Frame Frame1 
      Caption         =   "Record"
      Height          =   3255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2415
      Begin VB.CommandButton cmd_end 
         Caption         =   "&End"
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.ListBox list_mouse 
         Height          =   2205
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   2175
      End
      Begin VB.CommandButton cmd_begin 
         Caption         =   "&Begin"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2640
      Top             =   1560
   End
   Begin VB.TextBox txt_mousey 
      Height          =   285
      Left            =   2640
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txt_mousex 
      Height          =   285
      Left            =   2640
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2640
      Top             =   960
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   3120
      Width           =   2415
      Begin VB.CommandButton cmd_clear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   1560
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmd_stop 
         Caption         =   "&Stop"
         Height          =   375
         Left            =   840
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmd_play 
         Caption         =   "&Play"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuLoop 
         Caption         =   "How to Loop"
      End
      Begin VB.Menu mnuDash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'Mouse Record v1.0 by Tazrockon
'******************************************************
'If you use this program or any of its code as your
'own please give props to me. This program only
'currently records mouse positions and plays them
'back to the user. It doesn't record clicks.

Private Sub cmd_begin_Click()
Timer1.Enabled = True
End Sub

Private Sub cmd_clear_Click()
list_mouse.Clear
End Sub

Private Sub cmd_end_Click()
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
End Sub

Private Sub cmd_play_Click()
Timer1.Enabled = False
Timer2.Enabled = False
If list_mouse.ListCount > 0 Then
    list_mouse.ListIndex = -1
    Timer3.Enabled = True
Else
    MsgBox "You have not yet recorded anything to play back.", vbOKOnly + vbCritical, "Error"
End If
End Sub

Private Sub cmd_stop_Click()
Timer1.Enabled = False
Timer3.Enabled = False
End Sub

Private Sub mnuAbout_Click()
MsgBox "Mouse Record v1.0 was completed by Tazrockon in early January 2005.", vbOKOnly + vbInformation, "About Mouse Record"
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuLoop_Click()
MsgBox "To loop the playback of the mouse movements, after clicking Begin and recording the movements, click Play instead of clicking End.", vbOKOnly + vbInformation, "About Mouse Record"
End Sub

Private Sub Timer1_Timer()

'The following just adds the current mouse coords with
'the appropriate number of zeros to two text boxes
'which then are added to the list box. Making all the
'coords 4 digits long makes it easier to read them
'and play them back later. Mouse clicks are also
'detected here and are also added to the list.

Dim n As POINTAPI
Dim x As Long
Dim y As Long

GetCursorPos n
txt_mousex.Text = n.x
txt_mousey.Text = n.y

If Len(txt_mousex.Text) = 1 Then
    txt_mousex.Text = "000" & txt_mousex.Text
ElseIf Len(txt_mousex.Text) = 2 Then
    txt_mousex.Text = "00" & txt_mousex.Text
ElseIf Len(txt_mousex.Text) = 3 Then
    txt_mousex.Text = "0" & txt_mousex.Text
ElseIf Len(txt_mousex.Text) = "4" Then
    txt_mousex.Text = txt_mousex.Text
End If
 
If Len(txt_mousey.Text) = 1 Then
    txt_mousey.Text = "000" & txt_mousey.Text
ElseIf Len(txt_mousey.Text) = 2 Then
    txt_mousey.Text = "00" & txt_mousey.Text
ElseIf Len(txt_mousey.Text) = 3 Then
   txt_mousey.Text = "0" & txt_mousey.Text
ElseIf Len(txt_mousey.Text) = "4" Then
    txt_mousey.Text = txt_mousey.Text
End If

If GetAsyncKeyState(VK_LBUTTON) Then
    list_mouse.AddItem "LClick(" & txt_mousex.Text & "," & txt_mousey.Text & ")"
ElseIf GetAsyncKeyState(VK_RBUTTON) Then
    list_mouse.AddItem "RClick(" & txt_mousex.Text & "," & txt_mousey.Text & ")"
Else
    list_mouse.AddItem "Move(" & txt_mousex.Text & "," & txt_mousey.Text & ")"
End If

End Sub

Private Sub Timer3_Timer()

'Parsing the text of the current line in the list so
'that the program knows if it should left click, right
'click, or just move the mouse. Then it parses the lines
'again to find the x and y coords. The program ignores
'any extra zeros in the beginning of each coord. Once
'it has the coords it sets the mouse to that position
'and clicks if it should.

Dim lpt As String
Dim rpt As String
Dim whattodo As String

Timer1.Enabled = False
Timer2.Enabled = False

list_mouse.ListIndex = list_mouse.ListIndex + 1

whattodo = Left$(list_mouse.Text, 3)

If whattodo = "LCl" Then
    lpt = Mid$(list_mouse.Text, 8, 4)
    rpt = Mid$(list_mouse.Text, 13, 4)
    SetCursorPos lpt, rpt
    LeftClick
ElseIf whattodo = "RCl" Then
    lpt = Mid$(list_mouse.Text, 8, 4)
    rpt = Mid$(list_mouse.Text, 13, 4)
    SetCursorPos lpt, rpt
    RightClick
ElseIf whattodo = "Mov" Then
    lpt = Mid$(list_mouse.Text, 6, 4)
    rpt = Mid$(list_mouse.Text, 11, 4)
    SetCursorPos lpt, rpt
Else
    MsgBox "Unknown command on line " & list_mouse.ListIndex + 1 & ".", vbCritical + vbOKOnly, "Error"
End If

If list_mouse.ListIndex = list_mouse.ListCount - 1 Then
    Timer3.Enabled = False
    Exit Sub
End If
End Sub
