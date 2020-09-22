VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "BitBlt Example"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   354
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   4380
      Picture         =   "Form1.frx":57B2
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   1
      Top             =   60
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   3360
      Picture         =   "Form1.frx":B7F4
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Frame As Integer          'Tells if the frame is 1 or 2 in the current direction
Dim XPos As Integer, YPos As Integer    'Position holders for the position of the Wolf Man on the Form
Dim LeftX As Integer, TopY As Integer   'Tells where the frame of the animation is
Const Speed = 12  'Change this number to how fast you want him to move

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then    'The user pressed the Left arrow key
    If Frame = 0 Then
        Frame = 1
        LeftX = 0
        TopY = 96
        XPos = XPos - Speed
        MoveIt
    ElseIf Frame = 1 Then
        Frame = 0
        LeftX = 32
        TopY = 96
        XPos = XPos - Speed
        MoveIt
    End If
ElseIf KeyCode = vbKeyRight Then    'The user pressed the Right arrow key
    If Frame = 1 Then
        Frame = 0
        LeftX = 0
        TopY = 32
        XPos = XPos + Speed
        MoveIt
    ElseIf Frame = 0 Then
        Frame = 1
        LeftX = 32
        TopY = 32
        XPos = XPos + Speed
        MoveIt
    End If
ElseIf KeyCode = vbKeyUp Then    'The user pressed the Up arrow key
        If Frame = 1 Then
        Frame = 0
        LeftX = 0
        TopY = 0
        YPos = YPos - Speed
        MoveIt
    ElseIf Frame = 0 Then
        Frame = 1
        LeftX = 32
        TopY = 0
        YPos = YPos - Speed
        MoveIt
    End If
ElseIf KeyCode = vbKeyDown Then    'The user pressed the Down arrow key
    If Frame = 0 Then
        Frame = 1
        LeftX = 0
        TopY = 64
        YPos = YPos + Speed
        MoveIt
    ElseIf Frame = 1 Then
        Frame = 0
        LeftX = 32
        TopY = 64
        YPos = YPos + Speed
        MoveIt
    End If
End If
End Sub

Private Sub Form_Load()
XPos = Form1.ScaleWidth / 2    'Startup position on the form, is Center X
YPos = Form1.ScaleHeight / 2   'Startup position on the form, is Center Y
LeftX = 0                              'TopLeft X Coord of the Wolf Man Frame 1
TopY = 0                              'TopLeft Y Coord of the Wolf Man Frame 1
MoveIt
End Sub

'This sub updates and moves the Wolf Man
Sub MoveIt()
'Cls clears all the pixels from the screen.
Form1.Cls
'32 here is the length and width of each Wolf Man frame.  He is 32 x 32
Call BitBlt(Form1.hDC, XPos, YPos, 32, 32, picMask.hDC, LeftX, TopY, SRCAND)
Call BitBlt(Form1.hDC, XPos, YPos, 32, 32, picSprite.hDC, LeftX, TopY, SRCINVERT)
'Refresh must be always be written to tell the program you want to see all the picture information
Form1.Refresh
End Sub
