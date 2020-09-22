VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MaskMagic 1.0 - by, David Peace"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8640
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMain.frx":030A
   ScaleHeight     =   460
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   576
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbList 
      Height          =   315
      ItemData        =   "frmMain.frx":4614A
      Left            =   6000
      List            =   "frmMain.frx":46157
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   5520
      Width           =   2415
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   6000
      TabIndex        =   22
      Top             =   2040
      Width           =   2475
   End
   Begin VB.CommandButton Command3 
      Caption         =   "New Folder"
      Height          =   375
      Left            =   6480
      TabIndex        =   21
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Begin"
      Height          =   375
      Left            =   6000
      TabIndex        =   19
      Top             =   6120
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove Picture"
      Height          =   375
      Left            =   1500
      TabIndex        =   18
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Picture"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   3000
      Width           =   1275
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   255
      LargeChange     =   500
      Left            =   3060
      Max             =   0
      SmallChange     =   1000
      TabIndex        =   13
      Top             =   6480
      Width           =   2475
   End
   Begin VB.VScrollBar VScroll3 
      Height          =   2715
      LargeChange     =   500
      Left            =   5520
      Max             =   0
      SmallChange     =   1000
      TabIndex        =   12
      Top             =   3780
      Width           =   255
   End
   Begin VB.PictureBox picBack3 
      Height          =   2715
      Left            =   3060
      ScaleHeight     =   177
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   10
      Top             =   3780
      Width           =   2475
      Begin VB.PictureBox picSprite 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2715
         Left            =   0
         ScaleHeight     =   181
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   165
         TabIndex        =   11
         Top             =   0
         Width           =   2475
      End
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      LargeChange     =   500
      Left            =   120
      Max             =   0
      SmallChange     =   1000
      TabIndex        =   9
      Top             =   6480
      Width           =   2475
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   2715
      LargeChange     =   500
      Left            =   2580
      Max             =   0
      SmallChange     =   1000
      TabIndex        =   8
      Top             =   3780
      Width           =   255
   End
   Begin VB.PictureBox picBack2 
      Height          =   2715
      Left            =   120
      ScaleHeight     =   177
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   6
      Top             =   3780
      Width           =   2475
      Begin VB.PictureBox picMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2715
         Left            =   0
         ScaleHeight     =   181
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   165
         TabIndex        =   7
         Top             =   0
         Width           =   2475
      End
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   500
      Left            =   3060
      Max             =   0
      SmallChange     =   1000
      TabIndex        =   4
      Top             =   3120
      Width           =   2475
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2715
      LargeChange     =   500
      Left            =   5520
      Max             =   0
      SmallChange     =   1000
      TabIndex        =   3
      Top             =   420
      Width           =   255
   End
   Begin VB.PictureBox picBack 
      Height          =   2715
      Left            =   3060
      ScaleHeight     =   177
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   1
      Top             =   420
      Width           =   2475
      Begin VB.PictureBox picOriginal 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2715
         Left            =   0
         MousePointer    =   2  'Cross
         ScaleHeight     =   181
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   165
         TabIndex        =   5
         Top             =   0
         Width           =   2475
      End
   End
   Begin VB.ListBox lstPics 
      Height          =   2400
      Left            =   120
      TabIndex        =   0
      Top             =   420
      Width           =   2715
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   392
      X2              =   564
      Y1              =   104
      Y2              =   104
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C0C0&
      X1              =   392
      X2              =   564
      Y1              =   336
      Y2              =   336
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C0C0C0&
      X1              =   392
      X2              =   560
      Y1              =   16
      Y2              =   16
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0C0C0&
      X1              =   392
      X2              =   392
      Y1              =   448
      Y2              =   16
   End
   Begin VB.Label G 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "G: 255"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   7020
      TabIndex        =   26
      Top             =   1260
      Width           =   480
   End
   Begin VB.Label B 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "B: 255"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   7920
      TabIndex        =   25
      Top             =   1260
      Width           =   465
   End
   Begin VB.Label R 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R: 255"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6060
      TabIndex        =   24
      Top             =   1260
      Width           =   480
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Destination Directory:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6000
      TabIndex        =   23
      Top             =   1740
      Width           =   1515
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click Picture to Select BGColor"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3240
      TabIndex        =   20
      Top             =   180
      Width           =   2190
   End
   Begin VB.Shape shpBack 
      BackStyle       =   1  'Opaque
      Height          =   315
      Left            =   6840
      Shape           =   2  'Oval
      Top             =   780
      Width           =   855
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transparent\Background Color"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6120
      TabIndex        =   16
      Top             =   420
      Width           =   2205
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sprite:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3060
      TabIndex        =   15
      Top             =   3540
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mask:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   3540
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ToDo List:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   180
      Width           =   735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Directory As String
Dim Down As Boolean

Function Mask(PicSrc As PictureBox, picDEST As PictureBox, bColor As OLE_COLOR)
Dim looper As Integer
Dim looper2 As Integer
Dim bColor2 As OLE_COLOR
picDEST.Cls
For looper = 0 To PicSrc.Height
picDEST.Refresh
    For looper2 = 0 To PicSrc.Width
        If PicSrc.Point(looper2, looper) = bColor Then
            bColor2 = RGB(255, 255, 255)
        Else
            bColor2 = RGB(0, 0, 0)
        End If
        SetPixel picDEST.hdc, looper2, looper, bColor2
    Next looper2
Next looper
picDEST.Refresh
End Function

Function Sprite(PicSrc As PictureBox, picDEST As PictureBox, bColor As OLE_COLOR)
Dim looper As Integer
Dim looper2 As Integer
Dim bColor2 As OLE_COLOR
picDEST.Cls
For looper = 0 To PicSrc.Height
picDEST.Refresh
    For looper2 = 0 To PicSrc.Width
        If PicSrc.Point(looper2, looper) = bColor Then
            bColor2 = RGB(0, 0, 0)
        Else
            bColor2 = GetPixel(PicSrc.hdc, looper2, looper)
        End If
        SetPixel picDEST.hdc, looper2, looper, bColor2
    Next looper2
Next looper
picDEST.Refresh
End Function

Function TakeRGB(Colors As Long, Index As Integer) As Long
IndexColor = Colors
Red = IndexColor - Int(IndexColor / 256) * 256: IndexColor = (IndexColor - Red) / 256
Green = IndexColor - Int(IndexColor / 256) * 256: IndexColor = (IndexColor - Green) / 256
Blue = IndexColor
If Index = 1 Then TakeRGB = Red
If Index = 2 Then TakeRGB = Green
If Index = 3 Then TakeRGB = Blue
End Function

Public Sub SetScrollBars(Pic2 As PictureBox, Pic1 As PictureBox, Vert As VScrollBar, Horz As HScrollBar)
ScaleMode = 1
Pic2.ScaleMode = 1
Pic1.ScaleMode = 1
Vert.Min = 0
Vert.Max = (Pic2.ScaleHeight - Pic1.ScaleHeight) * -1
Vert.SmallChange = 100
Vert.LargeChange = Pic1.ScaleHeight / 4
Horz.Min = 0
Horz.Max = (Pic2.ScaleWidth - Pic1.ScaleWidth) * -1
Horz.SmallChange = 100
Horz.LargeChange = Pic1.ScaleWidth / 4
ScaleMode = 3
Pic2.ScaleMode = 3
Pic1.ScaleMode = 3
End Sub

Private Sub Command1_Click()
Dim FileName As SelectedFile
FileDialog.sDlgTitle = "Import Picture"
FileDialog.sFilter = "Bitmap Files (*.bmp)" & Chr(0) & "*.bmp" & Chr(0) & "All Files (*.*)" & Chr(0) & "*.*"
FileName = ShowOpen(hWnd)
If FileName.bCanceled = True Then GoTo errhandler
lstPics.AddItem FileName.sLastDirectory & FileName.sFiles(1)
lstPics.Selected(lstPics.ListCount - 1) = True
errhandler:
End Sub

Private Sub Command2_Click()
Dim lIndex As Integer
lIndex = lstPics.ListIndex
lstPics.RemoveItem lstPics.ListIndex
On Error Resume Next
lstPics.Selected(lIndex - 1) = True: picOriginal.Picture = LoadPicture(lstPics.List(lstPics.ListIndex))
lstPics.Selected(lIndex) = True: picOriginal.Picture = LoadPicture(lstPics.List(lstPics.ListIndex))
End Sub

Private Sub Command3_Click()
Dim Ans As String
Dim retval As Integer
Dim Notin As SECURITY_ATTRIBUTES
On Error GoTo errhandler
Ans = InputBox("Please name New Folder.", "New Folder", "New Folder")
retval = CreateDirectory(Directory & Ans, Notin)
Dir1.Refresh
Exit Sub
errhandler:
    MsgBox "Error '" & Err.Number & "' occurred" & vbCrLf & Err.Description, vbCritical, "Error."
End Sub

Private Sub Command4_Click()
Dim looper As Integer
Dim looper2 As Integer
Dim AllFiles As String
Dim Ans As Integer
Dim Pos As Integer
Dim FileTitle As String
If lstPics.ListCount = 0 Then MsgBox "Please insert pictures into the list on the left to begin.": Exit Sub
Ans = MsgBox("WARNING: Once begun, the process must be finished, and can not be cancelled." & vbCrLf & "Continue?", vbInformation + vbYesNo, "Ready?")
AllFiles = Date & " - " & Time & ":  Mask\Sprite Creation Begun." & vbCrLf
If Ans = vbNo Then Exit Sub
For looper = 1 To lstPics.ListCount
    lstPics.Selected(looper - 1) = True
    AllFiles = AllFiles & vbCrLf & Date & " - " & Time & ":  Picture File Loaded:" _
        & vbCrLf & lstPics.List(lstPics.ListIndex) & vbCrLf
    DoEvents
    For looper2 = 1 To Len(lstPics.List(lstPics.ListIndex))
        If Mid(lstPics.List(lstPics.ListIndex), looper2, 1) = "\" Then Pos = looper2
    Next looper2
    FileTitle = Right(lstPics.List(lstPics.ListIndex), Len(lstPics.List(lstPics.ListIndex)) - (Pos))
    FileTitle = Left(FileTitle, Len(FileTitle) - 4)
    
    If cmbList.Text = "Make Both" Or cmbList.Text = "Make Masks" Then
        Mask picOriginal, picMask, shpBack.BackColor
        SavePicture picMask.Image, Directory & "\" & FileTitle & "_Mask.bmp"
        AllFiles = AllFiles & vbCrLf & Date & " - " & Time & ":  Mask Made an Copied to:" & vbCrLf & Directory & _
            FileTitle & "_Mask.bmp" & vbCrLf
    End If
    If cmbList.Text = "Make Both" Or cmbList.Text = "Make Sprites" Then
        Sprite picOriginal, picSprite, shpBack.BackColor
        SavePicture picSprite.Image, Directory & "\" & FileTitle & "_Sprite.bmp"
        AllFiles = AllFiles & vbCrLf & Date & " - " & Time & ":  Sprite Made an Copied to:" & vbCrLf & Directory & _
            FileTitle & "_Sprite.bmp" & vbCrLf
    End If
Next looper
AllFiles = AllFiles & vbCrLf & Date & " - " & Time & ":  Text LOG Made and Copied to:" & vbCrLf & Directory & "Mask_Log.txt" & vbCrLf
AllFiles = AllFiles & vbCrLf & Date & " - " & Time & ":  Mask\Sprite Creation finished." & vbCrLf
Open Directory & "\Mask_Log.txt" For Append As #1
    Print #1, AllFiles
Close #1
MsgBox "Finished Successfully!" & vbCrLf & "All Masks\Sprites have been saved in the following directory:" _
        & vbCrLf & Directory, vbInformation, "Done Successfully!"
End Sub

Private Sub Dir1_Change()
Directory = Dir1.List(Dir1.ListIndex) & "\"
End Sub

Private Sub Form_Load()
shpBack.BackColor = RGB(255, 255, 255)
Directory = Dir1.List(Dir1.ListIndex) & "\"
cmbList.Text = "Make Both"
End Sub

Private Sub lstPics_Click()
picOriginal.Picture = LoadPicture(lstPics.List(lstPics.ListIndex))
picMask.Width = picOriginal.Width
picMask.Height = picOriginal.Height
picSprite.Width = picOriginal.Width
picSprite.Height = picOriginal.Height
SetScrollBars picBack, picOriginal, VScroll1, HScroll1
SetScrollBars picBack2, picMask, VScroll2, HScroll2
SetScrollBars picBack3, picSprite, VScroll3, HScroll3
picOriginal.Cls: picMask.Cls: picSprite.Cls
End Sub

Private Sub HScroll1_Scroll()
picBack.ScaleMode = 1
picOriginal.Left = -HScroll1.Value
picBack.ScaleMode = 3
End Sub

Private Sub HScroll1_Change()
picBack.ScaleMode = 1
picOriginal.Left = -HScroll1.Value
picBack.ScaleMode = 3
End Sub

Private Sub lstPics_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
lstPics.ToolTipText = lstPics.List(lstPics.ListIndex)
End Sub

Private Sub picOriginal_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Down = True
On Error Resume Next
shpBack.BackColor = picOriginal.Point(x, y)
R = "R: " & TakeRGB(shpBack.BackColor, 1)
G = "G: " & TakeRGB(shpBack.BackColor, 2)
B = "B: " & TakeRGB(shpBack.BackColor, 3)
End Sub

Private Sub picOriginal_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Down Then
On Error Resume Next
    shpBack.BackColor = picOriginal.Point(x, y)
    R = "R: " & TakeRGB(shpBack.BackColor, 1)
    G = "G: " & TakeRGB(shpBack.BackColor, 2)
    B = "B: " & TakeRGB(shpBack.BackColor, 3)
End If
End Sub

Private Sub picOriginal_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Down = False
End Sub

Private Sub VScroll1_Scroll()
picBack.ScaleMode = 1
picOriginal.Top = -VScroll1.Value
picBack.ScaleMode = 3
End Sub

Private Sub VScroll1_Change()
picBack.ScaleMode = 1
picOriginal.Top = -VScroll1.Value
picBack.ScaleMode = 3
End Sub

Private Sub HScroll2_Scroll()
picBack2.ScaleMode = 1
picMask.Left = -HScroll2.Value
picBack2.ScaleMode = 3
End Sub

Private Sub HScroll2_Change()
picBack2.ScaleMode = 1
picMask.Left = -HScroll2.Value
picBack2.ScaleMode = 3
End Sub

Private Sub VScroll2_Scroll()
picBack2.ScaleMode = 1
picMask.Top = -VScroll2.Value
picBack2.ScaleMode = 3
End Sub

Private Sub VScroll2_Change()
picBack2.ScaleMode = 1
picMask.Top = -VScroll2.Value
picBack2.ScaleMode = 3
End Sub

Private Sub HScroll3_Scroll()
picBack3.ScaleMode = 1
picSprite.Left = -HScroll3.Value
picBack3.ScaleMode = 3
End Sub

Private Sub HScroll3_Change()
picBack3.ScaleMode = 1
picSprite.Left = -HScroll3.Value
picBack3.ScaleMode = 3
End Sub

Private Sub VScroll3_Scroll()
picBack3.ScaleMode = 1
picSprite.Top = -VScroll3.Value
picBack3.ScaleMode = 3
End Sub

Private Sub VScroll3_Change()
picBack3.ScaleMode = 1
picSprite.Top = -VScroll3.Value
picBack3.ScaleMode = 3
End Sub

