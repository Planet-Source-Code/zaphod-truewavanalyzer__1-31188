VERSION 5.00
Begin VB.Form PlotFreq 
   AutoRedraw      =   -1  'True
   Caption         =   "Frequency Analyser"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6105
   ScaleWidth      =   11880
   Begin VB.PictureBox Picture1 
      Height          =   5415
      Left            =   0
      ScaleHeight     =   5355
      ScaleWidth      =   11835
      TabIndex        =   0
      Top             =   600
      Width           =   11895
      Begin VB.HScrollBar HScroll1 
         Height          =   240
         LargeChange     =   600
         Left            =   0
         Max             =   17700
         Min             =   -60
         SmallChange     =   90
         TabIndex        =   2
         Top             =   4920
         Width           =   11775
      End
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00400040&
         ForeColor       =   &H8000000E&
         Height          =   4960
         Left            =   0
         MousePointer    =   2  'Cross
         ScaleHeight     =   4905
         ScaleWidth      =   29445
         TabIndex        =   1
         Top             =   0
         Width           =   29500
         Begin VB.Line Line1 
            BorderColor     =   &H0000FFFF&
            X1              =   3720
            X2              =   3720
            Y1              =   0
            Y2              =   4920
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00400040&
            BackStyle       =   0  'Transparent
            Caption         =   "Frequency Analysis"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   9960
            TabIndex        =   3
            Top             =   0
            Width           =   2205
         End
      End
   End
   Begin VB.Label Label12 
      Caption         =   "57"
      Height          =   255
      Left            =   6360
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label11 
      Caption         =   "52"
      Height          =   255
      Left            =   5520
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   4440
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Label9"
      Height          =   195
      Left            =   3120
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Hz"
      Height          =   195
      Left            =   5040
      TabIndex        =   7
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Octave:"
      Height          =   195
      Left            =   8760
      TabIndex        =   6
      Top             =   360
      Width           =   570
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   5760
      TabIndex        =   5
      Top             =   360
      Width           =   45
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   4920
      TabIndex        =   4
      Top             =   360
      Width           =   45
   End
End
Attribute VB_Name = "PlotFreq"
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
Dim Ad$
Dim NotFilePos1 As Integer, NotFilePos2 As Integer
Dim Original(32768) As Double 'original data (before FFT)
Dim AfterFFT(32768) As Double 'data after FFT calculation
Dim yi(16384) As Double, yimax As Double 'imaginary
Dim yr(16384) As Double, yrmax As Double 'real
Dim ymod(16384) As Double, ymodmax As Double 'vector
Dim SampFreq As Long 'File Sampling Frequency

Sub FFTWave(Y() As Double, Npont As Long, Freq As Long, Sectime As String, SelSamp As Long)
    Me.Caption = "Frequency Analysis for first " & SelSamp & " Samples of " & Sectime & " Selected."
    
    Dim N As Long, g As Long
    N = Npont / 2
    'Store original data
    SampFreq = Freq
    For g = 1 To Npont
        Original(g) = Y(g)
    Next g
    RealFFT Y(), N, 1
    'Store FFT data
    For g = 1 To Npont
        AfterFFT(g) = Y(g)
    Next g
    GraphFFT Y(), N
    PlotFreq.SetFocus
End Sub
Sub RealFFT(Y() As Double, N As Long, Isign As Integer)
    Dim wr As Double, wi As Double, wpr As Double
    Dim PIsin As Double, TmpW As Double, CalcA As Double
    Dim c1 As Double, c2 As Double
    Dim PB As Long, Paul As Long, i As Long
    Dim i1 As Long, i2 As Long, i3 As Long, i4 As Long
    Dim wrs As Single, wis As Single
    Dim h1r As Double, h1i As Double
    Dim h2r As Double, h2i As Double
    PB = 2 * N
    CalcA = 3.14159265358979 / CDbl(N)
    c1 = 0.5
    If Isign = 1 Then
        c2 = -0.5
        PlotIt Y(), N, 1
      Else
        c2 = 0.5
        CalcA = -CalcA
    End If
    wpr = -2# * Sin(0.5 * CalcA) ^ 2
    PIsin = Sin(CalcA)
    wr = 1# + wpr
    wi = PIsin
    Paul = 2 * N + 3
    For i = 2 To N / 2 + 1
       i1 = 2 * i - 1
       i2 = i1 + 1
       i3 = Paul - i2
       i4 = i3 + 1
       wrs = CSng(wr)
       wis = CSng(wi)
       h1r = c1 * (Y(i1) + Y(i3))
       h1i = c1 * (Y(i2) - Y(i4))
       h2r = -c2 * (Y(i2) + Y(i4))
       h2i = c2 * (Y(i1) - Y(i3))
       Y(i1) = h1r + wrs * h2r - wis * h2i
       Y(i2) = h1i + wrs * h2i + wis * h2r
       Y(i3) = h1r - wrs * h2r + wis * h2i
       Y(i4) = -h1i + wrs * h2i + wis * h2r
       TmpW = wr
       wr = wr * wpr - wi * PIsin + wr
       wi = wi * wpr + TmpW * PIsin + wi
    Next i
    If Isign = 1 Then
        h1r = Y(1)
        Y(1) = h1r + Y(2)
        Y(2) = h1r - Y(2)
      Else
        h1r = Y(1)
        Y(1) = c1 * (h1r + Y(2))
        Y(2) = c1 * (h1r - Y(2))
        PlotIt Y(), N, -1
    End If
End Sub

Sub PlotIt(Y() As Double, PB As Long, Isign As Integer)
    Dim N As Long, i As Long, j As Long
    Dim m As Long, mmax As Long, istep As Long
    Dim TmpR As Double, TmpI As Double
    Dim wr As Double, wi As Double, wpr As Double
    Dim PIsin As Double, TmpW As Double, CalcA As Double
    N = 2 * PB
    j = 1
    For i = 1 To N Step 2
       If j > i Then
          TmpR = Y(j)
          TmpI = Y(j + 1)
          Y(j) = Y(i)
          Y(j + 1) = Y(i + 1)
          Y(i) = TmpR
          Y(i + 1) = TmpI
       End If
       m = N / 2
1:     If (m >= 2 And j > m) Then
          j = j - m
          m = m / 2
          GoTo 1
       End If
       j = j + m
    Next i
    mmax = 2
2:  If N > mmax Then
       istep = 2 * mmax
       CalcA = 6.28318530717959 / (Isign * mmax)
       wpr = -2 * Sin(0.5 * CalcA) ^ 2
       PIsin = Sin(CalcA)
       wr = 1
       wi = 0
       For m = 1 To mmax Step 2
          For i = m To N Step istep
             j = i + mmax
             TmpR = CSng(wr) * Y(j) - CSng(wi) * Y(j + 1)
             TmpI = CSng(wr) * Y(j + 1) + CSng(wi) * Y(j)
             Y(j) = Y(i) - TmpR
             Y(j + 1) = Y(i + 1) - TmpI
             Y(i) = Y(i) + TmpR
             Y(i + 1) = Y(i + 1) + TmpI
          Next i
          TmpW = wr
          wr = wr * wpr - wi * PIsin + wr
          wi = wi * wpr + TmpW * PIsin + wi
       Next m
       mmax = istep
       GoTo 2
    End If
End Sub
Sub GraphFFT(Y() As Double, CurSamp As Long)
    Dim g As Long
    'Separate real from imaginary; save; calculate vector; save;
    'and finally find maximum values for each case
    yimax = 0
    yrmax = 0
    ymodmax = 0
    For g = 0 To CurSamp - 1
        yr(g + 1) = Y(g * 2 + 1)
        If Abs(yr(g + 1)) > yrmax Then
            yrmax = Abs(yr(g + 1))
        End If
        yi(g + 1) = Y(g * 2 + 2)
        If Abs(yi(g + 1)) > yimax Then
            yimax = Abs(yi(g + 1))
        End If
        ymod(g + 1) = ((yr(g + 1)) ^ 2 + (yi(g + 1)) ^ 2) ^ (1 / 2)
        If ymod(g + 1) > ymodmax Then
            ymodmax = ymod(g + 1)
        End If
    Next g
    Call DrawRuler(CurSamp, False)
End Sub
Sub DrawRuler(CurSamp As Long, SoEsc As Boolean)
    Dim a As Integer, u As Integer, xmin As Integer
    Dim xzero As Double, x440 As Integer
    Dim yzero As Double, ymaxgraf As Double
    Dim xmult As Double, xmax As Integer
    Dim ymult As Double, N As Long, PaulBryan As Double
    Dim mpl As Double, xn As Integer
    
    a = 1
Rule:
    u = 0
    Picture2.Cls
    xmin = 0
    xzero = 0.964615822 'Hz
    x440 = 15900 'twips
    yzero = Picture2.Height * 2 / 3 - 500
    If a = -1 Then yzero = Picture2.Height * 1 / 3
    ymaxgraf = Picture2.Height / 8
    If a = -1 Then ymaxgraf = 0
    xmult = x440 / Log(440 / xzero)
    xmax = 7362 '150 twips for each logical note
    Picture2.Line (xmin, yzero)-(xmin + Picture2.Width, yzero), &H0&
    If SoEsc = True Then GoTo NumRuler
    ymult = (yzero - ymaxgraf) / ymodmax
    Picture2.PSet (xmin + u, yzero - (a * ymod(1)) * ymult)
    PaulBryan = CurSamp * 2 / SampFreq
    For N = 1 To CurSamp - 1
       Picture2.Line -(Log(N / (PaulBryan * xzero)) * xmult + u, yzero - (a * ymod(N + 1)) * ymult), &HFF00&
    Next N
NumRuler:
    mpl = x440 / Log(440 / xzero)
    Picture2.Line (xmin, yzero + 200)-(xmin + Picture2.Width, yzero + 200)
    For N = 1 To 50
        xn = Int(Log(N / xzero) * mpl + u)
        Picture2.Line (xn, yzero + 200)-(xn, yzero + 260)
        If N < 5 And N > 1 Then
            Picture2.PSet (xn - 100, yzero + 280), &H400040
            Picture2.Print N
        End If
    Next N
    For N = 60 To 500 Step 10
        xn = Int(Log(N / xzero) * mpl + u)
        Picture2.Line (xn, yzero + 200)-(xn, yzero + 260)
    Next N
    For N = 600 To 5000 Step 100
        xn = Int(Log(N / xzero) * mpl + u)
        Picture2.Line (xn, yzero + 200)-(xn, yzero + 260)
    Next N
    For N = 6000 To 50000 Step 1000
        xn = Int(Log(N / xzero) * mpl + u)
        Picture2.Line (xn, yzero + 200)-(xn, yzero + 260)
    Next N
    For N = 1 To 5 Step 4
        xn = Int(Log(N / xzero) * mpl + u)
        Picture2.Line (xn, yzero + 200)-(xn, yzero + 360)
        Picture2.Circle (xn, yzero + 360), 20
        Picture2.PSet (xn - 100, yzero + 400), &H400040
        Picture2.Print N
    Next N
    For N = 10 To 50 Step 10
        xn = Int(Log(N / xzero) * mpl + u)
        Picture2.Line (xn, yzero + 200)-(xn, yzero + 360)
        Picture2.Circle (xn, yzero + 360), 20
        Picture2.PSet (xn - 120, yzero + 400), &H400040
        Picture2.Print N
    Next N
    For N = 100 To 500 Step 100
        xn = Int(Log(N / xzero) * mpl + u)
        Picture2.Line (xn, yzero + 200)-(xn, yzero + 360)
        Picture2.Circle (xn, yzero + 360), 20
        Picture2.PSet (xn - 180, yzero + 400), &H400040
        Picture2.Print N
    Next N
    For N = 1000 To 5000 Step 1000
        xn = Int(Log(N / xzero) * mpl + u)
        Picture2.Line (xn, yzero + 200)-(xn, yzero + 360)
        Picture2.Circle (xn, yzero + 360), 20
        Picture2.PSet (xn - 180, yzero + 400), &H400040
        Picture2.Print N / 1000; " K"
    Next N
    For N = 10000 To 50000 Step 10000
        xn = Int(Log(N / xzero) * mpl + u)
        Picture2.Line (xn, yzero + 200)-(xn, yzero + 360)
        Picture2.Circle (xn, yzero + 360), 20
        Picture2.PSet (xn - 180, yzero + 400), &H400040
        Picture2.Print N / 1000; " K"
    Next N
    For N = 5 To 50 Step 5
        xn = Int(Log(N / xzero) * mpl + u)
        Picture2.Line (xn, yzero + 200)-(xn, yzero + 320)
    Next N
    For N = 50 To 500 Step 50
        xn = Int(Log(N / xzero) * mpl + u)
        Picture2.Line (xn, yzero + 200)-(xn, yzero + 320)
    Next N
    For N = 500 To 5000 Step 500
        xn = Int(Log(N / xzero) * mpl + u)
        Picture2.Line (xn, yzero + 200)-(xn, yzero + 320)
    Next N
    For N = 5000 To 50000 Step 5000
        xn = Int(Log(N / xzero) * mpl + u)
        Picture2.Line (xn, yzero + 200)-(xn, yzero + 320)
    Next N
    'Call DrawLines

End Sub
Sub DrawLines()
    yzero = Picture2.Height * 2 / 3 + 400
    For N = 0 To 29500 Step 150
        Picture2.Line (N, yzero)-(N, yzero + 280), &HFFFF&
    Next N
    Picture2.Line (15900, yzero + 280)-(15900, yzero - 100)
End Sub

Private Sub Form_Load()
    Me.Width = 11900
    Me.Height = 6270
    Me.Top = 300
    Me.Left = 0
    HScroll1.Value = 12500
    Me.Icon = MDIMain.Icon
    Call DrawRuler(0, True)
End Sub

Private Sub HScroll1_Change()
    Picture2.Left = -HScroll1.Value
End Sub

Private Sub MMControl1_Done(NotifyCode As Integer)
    MMControl1.Command = "Close"
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Pn As Integer, Pt As Double, freqf As Double
    Dim Freq As Single, PNot As Single, PNotBas As Single
    Dim Octave As Integer, PNotInt As Integer
    Dim Note As String
    
    'Post Frequency Under the mouse position
    Pn = (X - 15900) / 15
    Pt = 2 ^ (1 / 120)
    freqf = 440 * (Pt ^ Pn)
    Freq = Int(freqf * 1000) / 1000
    If freqf - Freq >= 0.0005 Then Freq = Freq + 0.001
    Label2.Caption = Freq
    'If corresponds exactly to a note (turn captions Blue)
    PNot = Pn / 10
    If Abs(PNot - Int(PNot)) < 0.001 Then
        Label2.ForeColor = vbBlue
        Label3.ForeColor = vbBlue
      Else
        Label2.ForeColor = vbBlack
        Label3.ForeColor = vbBlack
    End If
    'To which note it belongs
    'and to which octave it belongs
    PNotBas = PNot
    Octave = 5
    PNotInt = Int(PNotBas)
    If PNotBas - PNotInt >= 0.5 Then
        PNotInt = PNotInt + 1
    End If
    XNotPlay = PNotInt * 10 * 15 + 15900
    Label9.Caption = PNotInt + 69 'note played
    Do While PNotInt < 0
        PNotInt = PNotInt + 12
        Octave = Octave - 1
    Loop
    Do While PNotInt >= 12
        PNotInt = PNotInt - 12
        Octave = Octave + 1
    Loop
    If PNotInt < 3 Then 'It is A, A# or B of the next octave
        Octave = Octave - 1
    End If
    Select Case PNotInt
        Case 0
            Note = "(A)"
        Case 12
            Note = "(A)"
        Case 1
            Note = "(A #)   or   (B b)"
        Case 2
            Note = "(B)"
        Case 3
            Note = "(C)"
        Case 4
            Note = "(C #)   or   (D b)"
        Case 5
            Note = "(D)"
        Case 6
            Note = "(D #)   or   (E b)"
        Case 7
            Note = "(E)"
        Case 8
            Note = "(F)"
        Case 9
            Note = "(F #)   or   (G b)"
        Case 10
            Note = "(G)"
        Case 11
            Note = "(G #)   or   (A b)"
    End Select
Fim:
    Label3.Caption = Note
    Label4.Caption = "Octave: " & Octave
    Line1.X1 = X
    Line1.X2 = X
    Line1.Visible = True
End Sub


