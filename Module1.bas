Attribute VB_Name = "Module1"
'TrueWavAnalyzer
'by: Paul Bryan in 2002
'Allows for graphical isolation of sample ranges
'will analyze by frequency and decibal for up
'to 32768 samples (VB Single Precision Demension Max)
'Uses the FFT alogorythm

Public SCount As Integer ' Multiple Open Wave Files
Public Scope(255) As Form ' Session Filecount
Public Sub main()
    MDIMain.Show
    frmAbout.StartFlash
End Sub
Public Sub LoadNewFile(Fname As String) ' Open another file
        SCount = SCount + 1
        
        Set Scope(SCount) = New WavForm
        Scope(SCount).SetFocus
        Call Scope(SCount).LoadFileData(SCount, Fname)
    
        Exit Sub
End Sub
