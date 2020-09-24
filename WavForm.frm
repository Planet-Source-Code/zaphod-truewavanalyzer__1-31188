VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form WavForm 
   Caption         =   "Waveform Viewer"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10875
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   MousePointer    =   99  'Custom
   ScaleHeight     =   5355
   ScaleWidth      =   10875
   Begin VB.CommandButton Command8 
      Caption         =   "&Cancel"
      Height          =   255
      Left            =   5880
      TabIndex        =   51
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Analysis by FFT"
      Height          =   1455
      Left            =   3240
      TabIndex        =   44
      Top             =   3840
      Width           =   2535
      Begin VB.CommandButton Command5 
         Caption         =   "&Analyze"
         Height          =   255
         Left            =   720
         TabIndex        =   50
         Top             =   480
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Caption         =   "Channel"
         Height          =   615
         Left            =   120
         TabIndex        =   47
         Top             =   720
         Width           =   2295
         Begin VB.OptionButton Option2 
            Caption         =   "Right"
            Height          =   195
            Left            =   1320
            TabIndex        =   49
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Left"
            Height          =   195
            Left            =   120
            TabIndex        =   48
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   1560
         TabIndex        =   46
         Top             =   240
         Width           =   75
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Samples Selected:"
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   240
         Width           =   1320
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Restore"
      Height          =   375
      Left            =   9960
      TabIndex        =   40
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Zoom"
      Height          =   375
      Left            =   9960
      TabIndex        =   39
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Play Loop"
      Height          =   375
      Left            =   2160
      TabIndex        =   26
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Stop"
      Height          =   375
      Left            =   1560
      TabIndex        =   25
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Pa&use"
      Height          =   375
      Left            =   960
      TabIndex        =   24
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "P&lay"
      Height          =   375
      Left            =   360
      TabIndex        =   23
      Top             =   4080
      Width           =   615
   End
   Begin MCI.MMControl MMControl1 
      Height          =   330
      Left            =   6120
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      AutoEnable      =   0   'False
      Enabled         =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.PictureBox Picture1 
      Height          =   3000
      Left            =   460
      ScaleHeight     =   2940
      ScaleWidth      =   10035
      TabIndex        =   0
      Top             =   480
      Width           =   10100
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000001&
         Height          =   1240
         Left            =   0
         ScaleHeight     =   1215
         ScaleWidth      =   10005
         TabIndex        =   20
         Top             =   1520
         Visible         =   0   'False
         Width           =   10040
         Begin VB.PictureBox Picture6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1240
            Left            =   1800
            ScaleHeight     =   1245
            ScaleWidth      =   1815
            TabIndex        =   21
            Top             =   0
            Visible         =   0   'False
            Width           =   1815
            Begin VB.PictureBox Picture7 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000001&
               ForeColor       =   &H80000005&
               Height          =   1240
               Left            =   480
               ScaleHeight     =   1215
               ScaleWidth      =   585
               TabIndex        =   22
               Top             =   0
               Width           =   615
               Begin VB.Line Line4 
                  BorderColor     =   &H80000005&
                  Visible         =   0   'False
                  X1              =   240
                  X2              =   240
                  Y1              =   0
                  Y2              =   1200
               End
            End
         End
         Begin VB.Line Line2 
            Visible         =   0   'False
            X1              =   600
            X2              =   600
            Y1              =   0
            Y2              =   1200
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H8000000D&
         Height          =   2480
         Left            =   0
         ScaleHeight     =   2445
         ScaleWidth      =   10005
         TabIndex        =   1
         Top             =   240
         Width           =   10040
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2480
            Left            =   1200
            ScaleHeight     =   2475
            ScaleWidth      =   2175
            TabIndex        =   2
            Top             =   0
            Visible         =   0   'False
            Width           =   2175
            Begin VB.PictureBox Picture4 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H8000000D&
               ForeColor       =   &H80000009&
               Height          =   2480
               Left            =   600
               ScaleHeight     =   2445
               ScaleWidth      =   705
               TabIndex        =   3
               Top             =   -30
               Width           =   735
               Begin VB.Line Line3 
                  BorderColor     =   &H80000005&
                  Visible         =   0   'False
                  X1              =   360
                  X2              =   360
                  Y1              =   0
                  Y2              =   2400
               End
            End
         End
         Begin VB.Line Line1 
            Visible         =   0   'False
            X1              =   600
            X2              =   600
            Y1              =   0
            Y2              =   2400
         End
      End
   End
   Begin VB.Label Label17 
      Caption         =   "17"
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   3120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   "6/1"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   42
      Top             =   1440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label15 
      Caption         =   "15/1"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   41
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   8880
      TabIndex        =   38
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   8880
      TabIndex        =   37
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   8880
      TabIndex        =   36
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "Selection:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   7920
      TabIndex        =   35
      Top             =   3540
      Width           =   1455
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "End"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   7920
      TabIndex        =   34
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Beginning"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   6960
      TabIndex        =   33
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   7920
      TabIndex        =   32
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   6960
      TabIndex        =   31
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Samples"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   5880
      TabIndex        =   30
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   7920
      TabIndex        =   29
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   6960
      TabIndex        =   28
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Time (sec)"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   5880
      TabIndex        =   27
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label15 
      Caption         =   "15/0"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   19
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Label14"
      Height          =   195
      Left            =   480
      TabIndex        =   18
      Top             =   3840
      Width           =   570
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "13"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   2160
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "12"
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   1920
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label Label11 
      Caption         =   "11"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Label10"
      Height          =   195
      Left            =   480
      TabIndex        =   14
      Top             =   3480
      Width           =   570
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Label9"
      Height          =   195
      Left            =   1320
      TabIndex        =   13
      Top             =   3480
      Width           =   480
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3720
      TabIndex        =   12
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Frequency:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   11
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   6480
      TabIndex        =   10
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Num of Samples:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4800
      TabIndex        =   9
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "4"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   195
      Left            =   1200
      TabIndex        =   7
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   10095
      TabIndex        =   6
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   480
      TabIndex        =   4
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "WavForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'TrueWavAnalyzer
'by: Paul Bryan in 2002
'Allows for graphical isolation of sample ranges
'will analyze by frequency and decibal for up
'to 32768 samples (VB Single Precision Demension Max)
'Uses the FFT alogorythm
' I hope this helps, feel free to re-use this code.
Dim Xant As Integer
Dim LastWidth As Long, LastLeft As Long
Dim MovEsq As Boolean, MovDir As Boolean
Dim MvLiner As Integer ' Play Position Holder
Dim RepeatIt As Integer, PlayControls As Boolean, MMCPart As Boolean ' Play Switches
Dim IniPlay As Long, FimPlay As Long ' Begin And End Markers
Dim Uniand As Double
Dim FormInstance As Integer ' Allows for Multiple Files to be Opened
Dim CurFile As String ' File count open this Session

Public Sub LoadFileData(Instance As Integer, Fname As String)
    CurFile = Fname
    FormInstance = Instance
    Call InitStart
    'On Error GoTo Errhandler
    Dim ylec As Long, ydate As Date, ysg As Single
    Dim yint As Integer, ybt As Byte
    Dim LenData As Long, InData As Long
    Me.Caption = "Waveform View - " & "(" & Fname & ")"
    Label1.Caption = "File: " & Fname
    Open Fname For Binary Access Read As #1
    Label2.Caption = "Size: " & LOF(1) & " Bytes"
    Label3.Left = Label1.Left + Label1.Width + 500
    Label3.Caption = "Last modified: " & FileDateTime(Fname)
    Label4.Caption = 0 'IniPlay of present Zoom
    Label17.Caption = 0 '= selected Samples
    For N = 1 To 100
        X$ = Input(4, #1)
        If X$ = "fmt " Then Exit For
    Next N
    Get #1, , ylec 'is Bits
    Get #1, , yint 'is Channels
    Get #1, , yint 'is 1 if mono and 2 if stereo
    If yint = 2 Then
        Label9.Caption = "Stereo"
      ElseIf yint = 1 Then
        Label9.Caption = "Mono"
      Else
        Label9.Caption = "Error!"
        GoTo Errhandler
    End If
    Get #1, , ylec 'is the Sampling frequency of the file
    Label8.Caption = ylec
    Get #1, , ylec 'is a multiple of the sample frequency
    Get #1, , yint 'is the divisor of the number of bytes of
    'data which gives the number of Samples in the .wav
    ydiv = yint
    Label12.Caption = ydiv
    Get #1, , yint 'is the number of bits (8 or 16)
    If yint = 8 Or yint = 16 Then
        Label10.Caption = yint & " bits"
      Else
        Label10.Caption = "Error"
        GoTo Errhandler
    End If
GotTheData:
    For N = 1 To 100
        X$ = Input(1, #1)
        If X$ = "d" Then Exit For
    Next N
    Z$ = Input(3, #1)
    If Z$ <> "ata" Then
        If N > 90 Then GoTo Errhandler
        temp = Seek(1)
        Seek #1, temp - 3
        GoTo GotTheData
    End If
    Get #1, , ylec '= num of bytes of data, start reading data here.
    Label13.Caption = ylec
    LenData = ylec / ydiv
    Label6(0).Caption = LenData
    Label6(1).Caption = LenData
    ExeTemp = LenData / (Label8.Caption)
    Extemp = (Int(ExeTemp * 1000)) / 1000
    If ExeTemp - Extemp >= 0.0005 Then
        Extemp = Extemp + 0.001
    End If
    Label14.Caption = "Length: " & Extemp & " seconds"
    Label15(0).Caption = ExeTemp
    Label15(1).Caption = ExeTemp
    FimPlay = Int(ExeTemp * 1000)
    InData = Seek(1) 'Loc(1) + 1 is the number of the first snd data byte of the file.
    Label11.Caption = InData
    Call GraphWaver(InData, LenData)
    Close #1
    Call DrawTickMarks
    PlayControls = True
    If Label9.Caption = "Stereo" Then
        Frame2.Visible = True
        Frame1.Height = 1335
      Else
        Frame2.Visible = False
        Frame1.Height = 800
    End If
    MousePointer = 0
    Exit Sub
Errhandler:
    MsgBox "Error!!", vbOKOnly
    Close #1
    Call InitStart
  
    Exit Sub
End Sub
Sub GraphWaver(InData As Long, LenData As Long) ' Draws the Scope for Waveform View
    Dim yByte As Byte
    Dim yzero As Double, xmax As Double, xmult As Double, yselfat As Double
    Dim yint As Integer, ypos As Integer, ygraf As Integer
    Dim limsup As Integer, Nbits As Integer
    Dim ysel As Long
    Dim nmult As Double, xpos As Integer
    Nbits = Val(Label10.Caption)
    StMo = Label9.Caption
    If StMo = "Stereo" Then
        Picture2.Height = 1240
        Picture5.Visible = True
      Else
        Picture2.Height = 2480
        Picture5.Visible = False
    End If
    yselfat = LenData / Picture2.ScaleWidth
    xzero = 0
    yzero = Picture2.ScaleHeight / 2
    xmax = Picture2.ScaleWidth
    ymax = 128
    ymaxgraf = Picture2.ScaleHeight * 3 / 8
    ymult = ymaxgraf / ymax
    ypos = Int(yzero + 15 * 128)
    Picture2.Line (0, yzero)-(xmax, yzero)
    Picture2.Line (0, yzero - ymax * ymult)-(xmax, yzero - ymax * ymult), &H8000000F
    Picture2.Line (0, yzero + ymax * ymult)-(xmax, yzero + ymax * ymult), &H8000000F
    Picture4.Line (0, yzero)-(xmax, yzero)
    Picture4.Line (0, yzero - ymax * ymult)-(xmax, yzero - ymax * ymult), &H8000000F
    Picture4.Line (0, yzero + ymax * ymult)-(xmax, yzero + ymax * ymult), &H8000000F
    If StMo = "Stereo" Then GoTo Stereo8
    If Nbits = 16 Then GoTo Mono16
Mono8:
    Get #1, InData, yByte
    ygraf = ypos - 15 * yByte
    Picture2.PSet (xzero, ygraf)
    Picture4.PSet (xzero, ygraf)
    If LenData <= Picture2.ScaleWidth Then
        nmult = Picture2.ScaleWidth / LenData
        For N = 1 To LenData - 1
            xpos = Int(N * nmult)
            Get #1, , yByte
            ygraf = ypos - 15 * yByte
            Picture2.Line -(xpos, ygraf)
            Picture4.Line -(xpos, ygraf)
        Next N
      Else
        For N = 1 To 9999
            ysel = InData + Int(N * yselfat)
            Get #1, ysel, yByte
            ygraf = (ypos - 15 * yByte)
            Picture2.Line -(N, ygraf)
            Picture4.Line -(N, ygraf)
        Next N
    End If
    GoTo Done
Mono16:
    Get #1, InData, yint
    ygraf = yzero - yint / 17
    Picture2.PSet (xzero, ygraf)
    Picture4.PSet (xzero, ygraf)
    If LenData <= Picture2.ScaleWidth Then
        nmult = Picture2.ScaleWidth / LenData
        For N = 1 To LenData - 1
            xpos = Int(N * nmult)
            Get #1, , yint
            ygraf = yzero - yint / 17
            Picture2.Line -(xpos, ygraf)
            Picture4.Line -(xpos, ygraf)
        Next N
      Else
        For N = 1 To 9999
            ysel = InData + 2 * Int(N * yselfat)
            Get #1, ysel, yint
            ygraf = yzero - yint / 17
            Picture2.Line -(N, ygraf)
            Picture4.Line -(N, ygraf)
        Next N
    End If
    GoTo Done
Stereo8:
    Picture2.CurrentX = 0
    Picture2.CurrentY = 0
    Picture2.Print "Left"
    Picture4.CurrentX = 0
    Picture4.CurrentY = 0
    Picture4.Print "Left"
    Picture5.Line (0, yzero)-(xmax, yzero)
    Picture5.Line (0, yzero - ymax * ymult)-(xmax, yzero - ymax * ymult), &H8000000F
    Picture5.Line (0, yzero + ymax * ymult)-(xmax, yzero + ymax * ymult), &H8000000F
    Picture5.CurrentX = 0
    Picture5.CurrentY = 0
    Picture5.Print "Right"
    Picture7.Line (0, yzero)-(xmax, yzero)
    Picture7.Line (0, yzero - ymax * ymult)-(xmax, yzero - ymax * ymult), &H8000000F
    Picture7.Line (0, yzero + ymax * ymult)-(xmax, yzero + ymax * ymult), &H8000000F
    Picture7.CurrentX = 0
    Picture7.CurrentY = 0
    Picture7.Print "Right"
    If Nbits = 16 Then GoTo Stereo16
    ypos = Int(yzero + 7 * 128)
    Get #1, InData, yByte 'left Channel
    ygraf = ypos - 7 * yByte '15 * yByte
    Picture2.PSet (xzero, ygraf)
    Picture4.PSet (xzero, ygraf)
    Get #1, , yByte 'right Channel
    ygraf = ypos - 7 * yByte
    Picture5.PSet (xzero, ygraf)
    Picture7.PSet (xzero, ygraf)
    If LenData <= Picture2.ScaleWidth Then
        nmult = Picture2.ScaleWidth / LenData
        For N = 1 To LenData - 1
            xpos = Int(N * nmult)
            Get #1, , yByte 'left Channel
            ygraf = ypos - 7 * yByte '15 * yByte
            Picture2.Line -(xpos, ygraf)
            Picture4.Line -(xpos, ygraf)
            Get #1, , yByte 'right Channel
            ygraf = ypos - 7 * yByte
            Picture5.Line -(xpos, ygraf)
            Picture7.Line -(xpos, ygraf)
        Next N
      Else
        For N = 1 To 9999
            ysel = InData + 2 * Int(N * yselfat)
            Get #1, ysel, yByte 'left Channel
            ygraf = ypos - 7 * yByte '15 * yByte
            Picture2.Line -(N, ygraf)
            Picture4.Line -(N, ygraf)
            Get #1, , yByte 'right Channel
            ygraf = ypos - 7 * yByte
            Picture5.Line -(N, ygraf)
            Picture7.Line -(N, ygraf)
        Next N
    End If
    GoTo Done

Stereo16:
    Get #1, InData, yint 'left Channel
    ygraf = yzero - yint / 35 '17
    Picture2.PSet (xzero, ygraf)
    Picture4.PSet (xzero, ygraf)
    Get #1, , yint 'right Channel
    ygraf = yzero - yint / 35
    Picture5.PSet (xzero, ygraf)
    Picture7.PSet (xzero, ygraf)
    If LenData <= Picture2.ScaleWidth Then
        nmult = Picture2.ScaleWidth / LenData
        For N = 1 To LenData - 1
            xpos = Int(N * nmult)
            Get #1, , yint 'left Channel
            ygraf = yzero - yint / 35 '17
            Picture2.Line -(xpos, ygraf)
            Picture4.Line -(xpos, ygraf)
            Get #1, , yint 'right Channel
            ygraf = yzero - yint / 35
            Picture5.Line -(xpos, ygraf)
            Picture7.Line -(xpos, ygraf)
        Next N
      Else
        For N = 1 To 9999
            ysel = InData + 4 * Int(N * yselfat)
            Get #1, ysel, yint 'left Channel
            ygraf = yzero - yint / 35 '17
            Picture2.Line -(N, ygraf)
            Picture4.Line -(N, ygraf)
            Get #1, , yint 'right Channel
            ygraf = yzero - yint / 35
            Picture5.Line -(N, ygraf)
            Picture7.Line -(N, ygraf)
        Next N
    End If
    
Done:

End Sub
Sub InitStart() ' Initialize
    Caption = ""
    Cls
    Me.Height = 5865
    Me.Width = 10995
    Me.Left = 500
    Me.Top = 1900
    IniPlay = 0
    RepeatIt = 0
    PlayControls = False
    MMCPart = True
    Picture3.Width = 0
    Picture3.Visible = False
    Picture6.Width = 0
    Picture6.Visible = False
    Picture2.Cls: Picture5.Cls: Picture4.Cls: Picture7.Cls
    Label1.Caption = "": Label2.Caption = "":  Label3.Caption = "": Label6(0).Caption = ""
    Label6(1).Caption = "": Label8.Caption = "": Label9.Caption = "": Label10.Caption = ""
    Label11.Caption = "": Label12.Caption = "": Label13.Caption = "": Label14.Caption = ""
    Label19.Caption = ""
    For N = 1 To 6
        Label16(N).Caption = ""
    Next N

End Sub

Private Sub Command1_Click() 'Play
    If PlayControls = False Then
        Picture2.SetFocus
        Exit Sub
    End If
    RepeatIt = 0
    Call MMControl1_PlayClick(False)
    MMControl1.Command = "Play"
    Picture2.SetFocus
    Command1.Enabled = False: Command4.Enabled = False
End Sub

Private Sub Command2_Click() 'Pause
    If PlayControls = False Or MMCPart = True Then
        Picture2.SetFocus
        Exit Sub
    End If
    MMControl1.Command = "Pause"
    Picture2.SetFocus
    Command1.Enabled = False: Command4.Enabled = False
End Sub

Private Sub Command3_Click() 'Stop
    If PlayControls = False Then
        Picture2.SetFocus
        Exit Sub
    End If
    RepeatIt = 0
    MMControl1.Command = "Stop"
    Picture2.SetFocus
    Command1.Enabled = True: Command4.Enabled = True
End Sub

Private Sub Command4_Click() 'Play Loop
    If PlayControls = False Then
        Picture2.SetFocus
        Exit Sub
    End If
    Call MMControl1_PlayClick(False)
    RepeatIt = 1
    MMControl1.Command = "Play"
    Picture2.SetFocus
    Command1.Enabled = False: Command4.Enabled = False
End Sub

Private Sub Command5_Click() ' Send to FFT for Analysis
    Dim Nbits As Integer
    Dim StMo As String
    Dim InData As Long, LenData As Long, Freq As Long
    Dim InDataNov As Long, LenDataNov As Long
    Dim SampIni As Long, BytInic As Long
    Dim ydiv As Integer, yByte As Byte, yint As Integer
    
    If MMCPart = False Or Label8.Caption = "" Then
        Picture2.SetFocus
        Exit Sub
    End If
    If Label19.Caption = "" Then
        msg = "make a selection before Analyzing!"
        MsgBox msg, vbOKOnly
        Exit Sub
    End If
    MousePointer = vbHourglass
    Picture2.SetFocus
    Freq = Label8.Caption
    InData = Label11.Caption
    SampIni = Label16(4).Caption
    ydiv = Label12.Caption
    BytInic = SampIni * ydiv
    InDataNov = InData + BytInic
    LenDataNov = Label19.Caption
    Nbits = Val(Label10.Caption)
    StMo = Label9.Caption
    Open CurFile For Binary Access Read As #1
    Seek #1, InDataNov
    ReDim Y(1 To LenDataNov) As Double
    If StMo = "Stereo" Then GoTo Stereo8
    If Nbits = 16 Then GoTo Mono16
Mono8:
    For N = 1 To LenDataNov
        Get #1, , yByte
        Y(N) = CDbl(yByte - 128)
    Next N
    GoTo Fim
Mono16:
    For N = 1 To LenDataNov
        Get #1, , yint
        Y(N) = CDbl(yint)
    Next N
    GoTo Fim
Stereo8:
    If Nbits = 16 Then GoTo Stereo16
    If Option1.Value = True Then
        For N = 1 To LenDataNov
            Get #1, , yByte 'left Channel
            Y(N) = CDbl(yByte - 128)
            Get #1, , yByte 'right Channel
        Next N
      Else
        For N = 1 To LenDataNov
            Get #1, , yByte 'left Channel
            Get #1, , yByte 'right Channel
            Y(N) = CDbl(yByte - 128)
        Next N
    End If
    GoTo Fim
Stereo16:
    If Option1.Value = True Then
        For N = 1 To LenDataNov
            Get #1, , yint 'left Channel
            Y(N) = CDbl(yint)
            Get #1, , yint 'right Channel
        Next N
      Else
        For N = 1 To LenDataNov
            Get #1, , yint 'left Channel
            Get #1, , yint 'right Channel
            Y(N) = CDbl(yint)
        Next N
    End If
Fim:
    Close #1
    PlotFreq.FFTWave Y(), LenDataNov, Freq, Label16(3) & " Seconds", Val(Label19.Caption)
    MousePointer = 0

End Sub

Private Sub Command6_Click() ' Zoom on Selected
    Dim InData As Long, LenData As Long
    Dim InDataNov As Long, LenDataNov As Long
    Dim SampIni As Long, BytInic As Long
    Dim ydiv As Integer
    If MMCPart = False Or Label8.Caption = "" Then
        Picture2.SetFocus
        Exit Sub
    End If
    If Label16(4).Caption = "" Then
        msg = "No Zoom Selection Made"
        MsgBox msg, vbOKOnly
        Exit Sub
    End If
    MousePointer = vbHourglass
    Picture2.SetFocus
    Picture3.Width = 0
    Picture3.Visible = False
    Picture6.Width = 0
    Picture6.Visible = False
    Picture2.Cls: Picture5.Cls: Picture4.Cls: Picture7.Cls
    InData = Label11.Caption
    SampIni = Label16(4).Caption
    ydiv = Label12.Caption
    BytInic = SampIni * ydiv
    InDataNov = InData + BytInic
    LenDataNov = Label16(6).Caption + 1
    Open CurFile For Binary Access Read As #1
    Call GraphWaver(InDataNov, LenDataNov)
    Close #1
    Label4.Caption = IniPlay 'IniPlay of present zoom, without selection
    Label6(1).Caption = LenDataNov 'LenData of present zoom
    Label15(1).Caption = LenDataNov / Label8.Caption
    'Label8 contains file frequency
    'Label15(1) will be ExeTemp of actual zoom
    Label17.Caption = SampIni
    MousePointer = 0
End Sub

Private Sub Command7_Click() ' Restore The "Whole" Wav View
    Dim InData As Long, LenData As Long
    Dim ExeTemp As Double
    If MMCPart = False Then
        Picture2.SetFocus
        Exit Sub
    End If
    If Label16(4).Caption = "" Then
        Picture2.SetFocus
        Exit Sub
    End If
    MousePointer = vbHourglass
    Picture2.SetFocus
    IniPlay = 0
    Label4.Caption = 0
    Label15(1).Caption = Label15(0).Caption
    Label6(1).Caption = Label6(0).Caption
    Label17.Caption = 0
    RepeatIt = 0
    Picture3.Width = 0
    Picture3.Visible = False
    Picture6.Width = 0
    Picture6.Visible = False
    Picture2.Cls
    Picture5.Cls
    Picture4.Cls
    Picture7.Cls
    InData = Label11.Caption
    LenData = Label6(0).Caption
    Open CurFile For Binary Access Read As #1
    Call GraphWaver(InData, LenData)
    Close #1
    ExeTemp = Label15(1).Caption
    FimPlay = Int(ExeTemp * 1000)
    PlayControls = True
    Call Command8_Click
    MousePointer = 0
End Sub

Private Sub Command8_Click() ' Cancel
    Picture2.SetFocus
    If PlayControls = False Then Exit Sub
    If MMCPart = False Then Exit Sub
    Picture3.Width = 0
    Picture3.Visible = False
    Picture6.Width = 0
    Picture6.Visible = False
    IniPlay = Label4.Caption
    FimPlay = IniPlay + Label15(1).Caption * 1000
    If Label16(1) <> "" Then
        Label16(1).Caption = IniPlay / 1000
        Label16(2).Caption = FimPlay / 1000
        Label16(3).Caption = (FimPlay - IniPlay) / 1000
        Label16(4).Caption = Label17.Caption
        Label16(5).Caption = Val(Label6(1).Caption) + Val(Label17.Caption)
        'Data Length + Starting point of Zoom
        Label16(6).Caption = Label6(1).Caption 'Data Length of Zoom
        Call PosChange
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = MDIMain.Icon
    Me.Caption = "Waveform Viewer"
    Call InitStart
    Picture4.Width = Picture2.Width
    Picture7.Width = Picture2.Width
    MMControl1.Notify = False
    MMControl1.Wait = True
    MMControl1.Shareable = False ' Set properties needed by MCI to open.
    MMControl1.DeviceType = "WaveAudio"
    MMControl1.TimeFormat = mciFormatMilliseconds
    Option1.Value = True
    Frame2.Visible = False
    Frame1.Height = 800
End Sub

Private Sub Form_Unload(Cancel As Integer) ' Clean Up
    MMControl1.Command = "Close"
    Close #1
End Sub


Private Sub MMControl1_Done(NotifyCode As Integer)
    MMControl1.UpdateInterval = 0
    If RepeatIt = 1 Then ' Play Selection Again
        Call MMControl1_PlayClick(False)
        MMControl1.Command = "Play"
        Exit Sub
      Else ' Stop
        Command1.Enabled = True
        Command4.Enabled = True
    End If
    MMControl1.Command = "Close"
    Line1.Visible = False
    Line2.Visible = False
    Line3.Visible = False
    Line4.Visible = False
    MMCPart = True
End Sub

Private Sub MMControl1_PlayClick(Cancel As Integer) ' Play
    If RepeatIt = 1 Then ' Loop Mode
        MMControl1.UpdateInterval = 50
        MMControl1.From = IniPlay&
        MMControl1.To = FimPlay&
        'Track Playing Position
        Line1.X1 = MvLiner
        Line1.X2 = MvLiner
        Line2.X1 = MvLiner
        Line2.X2 = MvLiner
        Line3.X1 = MvLiner
        Line3.X2 = MvLiner
        Line4.X1 = MvLiner
        Line4.X2 = MvLiner
        Exit Sub
    End If
    ' Single Mode
    MMControl1.FileName = CurFile
    MMControl1.Command = "Open"
    MMControl1.From = IniPlay&
    MMControl1.To = FimPlay&
    MMControl1.UpdateInterval = 50
    ExeTemp = Label15(1).Caption
    Uniand = Picture2.ScaleWidth / (ExeTemp * 1000)
    MvLiner = Int((IniPlay - Label4.Caption) * Uniand)
'Track Playing Position
    Line1.X1 = MvLiner
    Line1.X2 = MvLiner
    Line2.X1 = MvLiner
    Line2.X2 = MvLiner
    Line3.X1 = MvLiner
    Line3.X2 = MvLiner
    Line4.X1 = MvLiner
    Line4.X2 = MvLiner
    If FimPlay - IniPlay > 500 Then
        If Picture3.Width > 100 Then
            Line3.Visible = True
            Line4.Visible = True
          Else
            Line1.Visible = True
            Line2.Visible = True
        End If
    End If
    MMCPart = False
End Sub

Private Sub MMControl1_StatusUpdate() ' Mark Play Position
    Z = Int((MMControl1.Position - Label4.Caption) * Uniand)
    Line1.X1 = Z
    Line1.X2 = Z
    Line2.X1 = Z
    Line2.X2 = Z
    Line3.X1 = Z
    Line3.X2 = Z
    Line4.X1 = Z
    Line4.X2 = Z
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If PlayControls = False Then Exit Sub
    If MMCPart = False Then Exit Sub
    If Button = vbRightButton Then
        If Picture3.Width < 40 Then Exit Sub
        If Picture3.Left - X < 50 And Picture3.Left - X > 0 Then
            Picture2.MousePointer = 9
            MovEsq = True
            LastWidth = Picture3.Width
            LastLeft = Picture3.Left
        End If
        If (X - Picture3.Left - Picture3.Width) < 50 And (X - Picture3.Left - Picture3.Width) > 0 Then
            Picture2.MousePointer = 9
            MovDir = True
        End If
    End If
    If Button = vbLeftButton Then
        Picture3.Width = 0
        Picture3.Left = X
        Picture3.Visible = True
        Picture4.Left = -X
        Picture6.Width = 0
        Picture6.Left = X
        Picture6.Visible = True
        Picture7.Left = -X
        IniPlay = Label4.Caption + X * Label15(1).Caption * 1000 / Picture2.ScaleWidth
        FimPlay = Label4.Caption + Label15(1).Caption * 1000
        Label16(1).Caption = IniPlay / 1000
        Label16(2).Caption = FimPlay / 1000
        Label16(3).Caption = (FimPlay - IniPlay) / 1000
        Label16(4).Caption = Label17.Caption + Int(X * (Label6(1).Caption - 1) / Picture2.ScaleWidth)
        Label16(5).Caption = Val(Label17.Caption) + Val(Label6(1).Caption) - 1
        Label16(6).Caption = Label16(5).Caption - Label16(4).Caption + 1
        Call PosChange
    End If
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Picture3.Visible = False Then Exit Sub
    If MMCPart = False Then Exit Sub
    If Button = vbLeftButton Then
        If X > Picture3.Left And X <= Picture2.ScaleWidth Then
            Picture3.Width = X - Picture3.Left
            Picture6.Width = X - Picture6.Left
            FimPlay = Label4.Caption + X * Label15(1).Caption * 1000 / Picture2.ScaleWidth
            Label16(2).Caption = FimPlay / 1000
            Label16(3).Caption = (FimPlay - IniPlay) / 1000
            Label16(5).Caption = Label17.Caption + Int(X * (Label6(1).Caption - 1) / Picture2.ScaleWidth)
            Label16(6).Caption = Label16(5).Caption - Label16(4).Caption + 1
            Call PosChange
        End If
    End If
    If Button = vbRightButton Then
        If X = Xant Then
            Picture4.Visible = True
            Picture7.Visible = True
            Exit Sub
          ElseIf MovEsq = True Then
            Call ShadeArea(X)
            Xant = X
            If X >= 0 Then
                IniPlay = Label4.Caption + Picture3.Left * Label15(1).Caption * 1000 / Picture2.ScaleWidth
                Label16(1).Caption = IniPlay / 1000
                Label16(3).Caption = (FimPlay - IniPlay) / 1000
                Label16(4).Caption = Label17.Caption + Int(Picture3.Left * (Label6(1).Caption - 1) / Picture2.ScaleWidth)
                Label16(6).Caption = Label16(5).Caption - Label16(4).Caption + 1
                Call PosChange
            End If
          ElseIf MovDir = True And X - Picture3.Left > 50 And X <= Picture2.ScaleWidth Then
            Picture3.Width = X - Picture3.Left
            Picture6.Width = Picture3.Width
            FimPlay = Label4.Caption + X * Label15(1).Caption * 1000 / Picture2.ScaleWidth
            Label16(2).Caption = FimPlay / 1000
            Label16(3).Caption = (FimPlay - IniPlay) / 1000
            Label16(5).Caption = Label17.Caption + Int(X * (Label6(1).Caption - 1) / Picture2.ScaleWidth)
            Label16(6).Caption = Label16(5).Caption - Label16(4).Caption + 1
            Call PosChange
        End If
    End If
End Sub
Sub ShadeArea(X)
    If LastWidth + LastLeft - X < 50 Then Exit Sub
    Picture4.Visible = False
    Picture7.Visible = False
    Picture3.Left = X
    Picture6.Left = X
    Picture4.Left = -X
    Picture7.Left = -X
    Picture3.Width = LastWidth + LastLeft - X
    Picture6.Width = Picture3.Width
    Picture4.Visible = True
    Picture7.Visible = True

End Sub
Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture2.MousePointer = 0
    MovEsq = False
    MovDir = False
End Sub

Private Sub Picture4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MMCPart = False Then Exit Sub
    If Button = vbRightButton Then
        If Picture3.Width < 100 Then
            Picture4.MousePointer = 9
            If X - Picture3.Left < Picture3.Width / 3 Then
                MovEsq = True
                LastWidth = Picture3.Width
                LastLeft = Picture3.Left
              Else
                MovDir = True
            End If
          ElseIf X - Picture3.Left < 50 Then
            Picture4.MousePointer = 9
            MovEsq = True
            LastWidth = Picture3.Width
            LastLeft = Picture3.Left
          ElseIf Picture3.Width + Picture3.Left - X < 100 Then
            Picture4.MousePointer = 9
            MovDir = True
        End If
    End If
    If Button = vbLeftButton Then
        Picture3.Width = 0
        Picture3.Left = X
        Picture3.Visible = True
        Picture4.Left = -X
        Picture6.Width = 0
        Picture6.Left = X
        Picture6.Visible = True
        Picture7.Left = -X
        IniPlay = Label4.Caption + X * Label15(1).Caption * 1000 / Picture2.ScaleWidth
        FimPlay = Label4.Caption + Label15(1).Caption * 1000
        Label16(1).Caption = IniPlay / 1000
        Label16(2).Caption = FimPlay / 1000
        Label16(3).Caption = (FimPlay - IniPlay) / 1000
        Label16(4).Caption = Label17.Caption + Int(X * (Label6(1).Caption - 1) / Picture2.ScaleWidth)
        Label16(5).Caption = Val(Label17.Caption) + Val(Label6(1).Caption) - 1
        Label16(6).Caption = Label16(5).Caption - Label16(4).Caption + 1
        Call PosChange
    End If

End Sub

Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MMCPart = False Then Exit Sub
    If Button = vbRightButton Then
        If X = Xant Then
            Picture4.Visible = True
            Picture7.Visible = True
            Exit Sub
          ElseIf MovEsq = True Then
            Call ShadeArea(X)
            Xant = X
            If X >= 0 Then
                IniPlay = Label4.Caption + Picture3.Left * Label15(1).Caption * 1000 / Picture2.ScaleWidth
                Label16(1).Caption = IniPlay / 1000
                Label16(3).Caption = (FimPlay - IniPlay) / 1000
                Label16(4).Caption = Label17.Caption + Int(Picture3.Left * (Label6(1).Caption - 1) / Picture2.ScaleWidth)
                Label16(6).Caption = Label16(5).Caption - Label16(4).Caption + 1
                Call PosChange
            End If
          ElseIf MovDir = True And X - Picture3.Left > 50 And X <= Picture2.ScaleWidth Then
            Picture3.Width = X - Picture3.Left
            Picture6.Width = Picture3.Width
            FimPlay = Label4.Caption + X * Label15(1).Caption * 1000 / Picture2.ScaleWidth
            Label16(2).Caption = FimPlay / 1000
            Label16(3).Caption = (FimPlay - IniPlay) / 1000
            Label16(5).Caption = Label17.Caption + Int(X * (Label6(1).Caption - 1) / Picture2.ScaleWidth)
            Label16(6).Caption = Label16(5).Caption - Label16(4).Caption + 1
            Call PosChange
        End If
    End If

End Sub

Private Sub Picture4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture4.MousePointer = 0
    MovDir = False
    MovEsq = False
End Sub

Private Sub Picture5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If PlayControls = False Then Exit Sub
    If MMCPart = False Then Exit Sub
    If Button = vbLeftButton Then
        Picture3.Width = 0
        Picture3.Left = X
        Picture3.Visible = True
        Picture4.Left = -X
        Picture6.Width = 0
        Picture6.Left = X
        Picture6.Visible = True
        Picture7.Left = -X
        IniPlay = Label4.Caption + X * Label15(1).Caption * 1000 / Picture2.ScaleWidth
        FimPlay = Label4.Caption + Label15(1).Caption * 1000
        Label16(1).Caption = IniPlay / 1000
        Label16(2).Caption = FimPlay / 1000
        Label16(3).Caption = (FimPlay - IniPlay) / 1000
        Label16(4).Caption = Label17.Caption + Int(X * (Label6(1).Caption - 1) / Picture2.ScaleWidth)
        Label16(5).Caption = Val(Label17.Caption) + Val(Label6(1).Caption) - 1
        Label16(6).Caption = Label16(5).Caption - Label16(4).Caption + 1
        Call PosChange
    End If
End Sub

Private Sub Picture5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Picture3.Visible = False Then Exit Sub
    If MMCPart = False Then Exit Sub
    If Button = vbLeftButton Then
        If X > Picture3.Left And X <= Picture2.ScaleWidth Then
            Picture3.Width = X - Picture3.Left
            Picture6.Width = X - Picture6.Left
            FimPlay = Label4.Caption + X * Label15(1).Caption * 1000 / Picture2.ScaleWidth
            Label16(2).Caption = FimPlay / 1000
            Label16(3).Caption = (FimPlay - IniPlay) / 1000
            Label16(5).Caption = Label17.Caption + Int(X * (Label6(1).Caption - 1) / Picture2.ScaleWidth)
            Label16(6).Caption = Label16(5).Caption - Label16(4).Caption + 1
            Call PosChange
        End If
    End If

End Sub
Sub DrawTickMarks()
   
    For N = 1 To 100
        Picture1.Line (N * 100, Picture2.Top)-(N * 100, Picture2.Top - 100)
    Next N
End Sub
Sub PosChange()
    Dim Npont As Long
    If Label16(6).Caption = "" Then
        Label19.Caption = ""
        Exit Sub
    End If
    Npont = Label16(6).Caption
    If Npont < 32 Then
        Label19.Caption = 32
        Exit Sub
      Else
        For N = 15 To 5 Step -1
            Pot2 = 2 ^ N
            If Npont > Pot2 Then
                Label19.Caption = Pot2
                Exit For
            End If
        Next N
    End If
End Sub
