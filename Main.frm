VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "VB Wav Analyzer by Paul Bryan in 2002"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5745
   Icon            =   "Main.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpenItem 
         Caption         =   "&Open (*.wav file)"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSepr 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExitItem 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "MDIMain"
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

Private Sub mnuAbout_Click()
    frmAbout.Show , Me
End Sub

Private Sub mnuOpenItem_Click()
    CommonDialog1.Filter = "Wave Files (*.wav)|*.wav"
    CommonDialog1.CancelError = True
    On Error GoTo Errhandler
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
        MousePointer = 11 'hourglass
        LoadNewFile (CommonDialog1.FileName)
        MousePointer = 0
        Exit Sub
    End If
Errhandler:
    Exit Sub

End Sub

Private Sub mnuExitItem_Click()
    End
End Sub
