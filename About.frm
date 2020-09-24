VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "IconForge"
   ClientHeight    =   4185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      FontTransparent =   0   'False
      Height          =   495
      Left            =   0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   409
      TabIndex        =   3
      ToolTipText     =   "ChromeSoft"
      Top             =   3720
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      FontTransparent =   0   'False
      Height          =   615
      Left            =   0
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   409
      TabIndex        =   2
      ToolTipText     =   "ChromeSoft"
      Top             =   3720
      Width           =   6135
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   3240
   End
   Begin VB.PictureBox picScroll 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   2535
      Left            =   480
      ScaleHeight     =   167
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   351
      TabIndex        =   0
      Top             =   720
      Width           =   5295
   End
   Begin VB.Label lblExit 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   360
      Left            =   5280
      TabIndex        =   4
      Top             =   3480
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   3210
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'TrueWavAnalyzer
'by: Paul Bryan in 2002
'
'Application Splash Screen
Option Explicit
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Byte
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Const DT_BOTTOM As Long = &H8
Const DT_CALCRECT As Long = &H400
Const DT_CENTER As Long = &H1
Const DT_EXPANDTABS As Long = &H40
Const DT_EXTERNALLEADING As Long = &H200
Const DT_LEFT As Long = &H0
Const DT_NOCLIP As Long = &H100
Const DT_NOPREFIX As Long = &H800
Const DT_RIGHT As Long = &H2
Const DT_SINGLELINE As Long = &H20
Const DT_TABSTOP As Long = &H80
Const DT_TOP As Long = &H0
Const DT_VCENTER As Long = &H4
Const DT_WORDBREAK As Long = &H10

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Const ScrollText As String = "Presented by:" & vbCrLf _
                             & vbCrLf & "ChromeSoft" & vbCrLf & _
                             vbCrLf & vbCrLf & _
                             "Author: Paul Bryan" & vbCrLf & _
                             "Copyright: 2002" & _
                             vbCrLf & "Email: pbryan@softhome.net" & _
                             vbCrLf & vbCrLf & "(Beta Test Version)" & _
                             vbCrLf & vbCrLf & _
                             "The Ultimate Wave Analyzer!"

Dim FX, FY As Integer
Dim Ak As Boolean
Dim EndingFlag As Boolean
Dim Started As Integer
Private Sub Form_Activate()
RunMain
End Sub
Public Sub StartFlash()
    Started = 1
    lblExit.Caption = "Start!"
    Me.Show
    Timer1.Enabled = True
    RunMain
End Sub
Private Sub Form_Load()
FX = 420
FY = 32
Picture1.Width = FX * Screen.TwipsPerPixelX + 4
Picture1.Height = FY * Screen.TwipsPerPixelY + 4
Picture2.Width = FX * Screen.TwipsPerPixelX + 4
Picture2.Height = FY * Screen.TwipsPerPixelY + 4
picScroll.ForeColor = vbYellow
picScroll.FontSize = 14
Picture1.Cls
Picture2.Cls
Randomize
End Sub
Private Sub RunMain()
Dim LastFrameTime As Long
Const IntervalTime As Long = 30
Dim rt As Long
Dim DrawingRect As RECT
Dim UpperX As Long, UpperY As Long
Dim RectHeight As Long
Static FlameArray() As Byte
ReDim FlameArray(0 To FX, 0 To FY) As Byte
Static FillVal As Byte
Static Frame As Integer
Static CurTime As Single
Static ProcDem As Byte
Static Temp As Single
Static Temp2 As Byte
Static X As Integer
Static Y As Integer
Static Color As Integer
Static Test As Byte
Static Uniformity As Byte
frmAbout.Refresh

Uniformity = 2
ProcDem = 1
LockWindowUpdate Picture2.hWnd
Temp = 256 / FY
FillVal = FY * 0.9
Label1.Caption = App.Title & "  v" & App.Major & "." & App.Minor & "." & App.Revision


rt = DrawText(picScroll.hdc, ScrollText, -1, DrawingRect, DT_CALCRECT)
If rt = 0 Then
    MsgBox "Error scrolling text", vbExclamation
    EndingFlag = True
Else
    DrawingRect.Top = picScroll.ScaleHeight
    DrawingRect.Left = 0
    DrawingRect.Right = picScroll.ScaleWidth
    RectHeight = DrawingRect.Bottom
    DrawingRect.Bottom = DrawingRect.Bottom + picScroll.ScaleHeight
End If
Do While Not EndingFlag
    
    If GetTickCount() - LastFrameTime > IntervalTime Then
                    
        picScroll.Cls
        
        DrawText picScroll.hdc, ScrollText, -1, DrawingRect, DT_CENTER Or DT_WORDBREAK
        DrawingRect.Top = DrawingRect.Top - 1
        DrawingRect.Bottom = DrawingRect.Bottom - 1
        If DrawingRect.Top < -(RectHeight) Then
            DrawingRect.Top = picScroll.ScaleHeight
            DrawingRect.Bottom = RectHeight + picScroll.ScaleHeight
        End If
        picScroll.Refresh
        LastFrameTime = GetTickCount()
    End If
    DoEvents
Frame = Frame + 1

If Frame Mod ProcDem = 0 Then DoEvents
For Y = FY To 4 Step -1
For X = 0 To FX Step 1
Temp2 = FlameArray(X, Y)
If Temp2 < Uniformity - 1 Then GoTo 1
Test = Int(Rnd * Uniformity)
FlameArray(X, Y) = Temp2 - Test
FlameArray(X, Y - Test) = FlameArray(X, Y)
Color = FlameArray(X, Y) * Temp
SetPixelV Picture2.hdc, X + (Rnd * 2), Y, RGB(Color + Color, Color, Color / 2)
1 Next X
Next Y

For X = 0 To FX
For Y = FillVal To FY
FlameArray(X, Y) = FY
Next Y
Next X
BitBlt Picture1.hdc, 0, 0, FX, FY, Picture2.hdc, 0, 0, vbSrcCopy

Loop
If Started = 1 Then MDIMain.Show
Unload Me
Set frmAbout = Nothing
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Started = 0 Then
lblExit.ForeColor = vbGreen
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    EndingFlag = True
End Sub
Private Sub lblExit_Click()
EndingFlag = True
End Sub
Private Sub lblExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblExit.ForeColor = vbRed
End Sub


Private Sub Timer1_Timer()
If Ak = False Then
    lblExit.ForeColor = vbBlack: Ak = True
    Else
    lblExit.ForeColor = vbGreen: Ak = False
    End If
End Sub
