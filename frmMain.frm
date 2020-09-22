VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Amplitude Modulation"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2715
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   2715
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMain 
      AutoSize        =   -1  'True
      Height          =   1035
      Left            =   960
      ScaleHeight     =   975
      ScaleWidth      =   1155
      TabIndex        =   4
      Top             =   60
      Width           =   1215
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   375
      Left            =   60
      TabIndex        =   3
      Top             =   1260
      Width           =   795
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   1320
      Top             =   1620
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin ComctlLib.Slider sldMain 
      Height          =   1155
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   2037
      _Version        =   327682
      Orientation     =   1
      Max             =   24
      SelStart        =   12
      Value           =   12
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1860
      TabIndex        =   1
      Top             =   1260
      Width           =   795
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Enabled         =   0   'False
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   1260
      Width           =   795
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'***********************************************************************
'This application was explicitly developed for
'PSC(Planet Source Code) Users as an Open Source Project.
'This code is the property of its author.
'
'If you compile this application you may not redistribute it.
'However, you may use any of this code in you're own application(s).
'
'Alex Smoljanovic, Salex Software (c) 2001-2003
'salex_software@shaw.ca
'***********************************************************************

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private DX8 As New DirectX8, DS8 As DirectSound8, WaveFormat As WAVEFORMATEX
Private sndBuffer As DirectSoundSecondaryBuffer8, sndBufferDesc As DSBUFFERDESC
'External Function Declarations...
Private CurRWPos As DSCURSORS 'Privately Scope Variable CurRWPos(Current Read Write Position) as DSCURSORS type structure
Private WAVFile$ 'Privately Scope Variable WAVFile as string data type
Dim byteAr(1) As Integer 'dimensionalize byteAr as a one dimensionalize fixed array with one element of integer data type
Dim arBMP(1 To 24) As String 'dimensionalize arBMP as a one dimensionalize fixed array with 24 elements of string data type
'general declarations...

Private Sub InitDS() 'Direct Sound Initialization sub routine
If Dir(WAVFile, vbNormal Or vbHidden) = "" Then Exit Sub
'If the file specified by variable WAVFile's value(and has normal or hidden file attributes) doesn't exist then exit this procedure...
If Not (DS8 Is Nothing) Then Set DS8 = Nothing 'if variable DS8 is currently initialized with an instance of the DirectSound8 class, then terminate its instance
 Set DS8 = DX8.DirectSoundCreate("") 'initialize DS8 with a new instance of DirectSound8 returned by DX8's DirectSoundCreate method...
  DS8.SetCooperativeLevel Me.hWnd, DSSCL_PRIORITY
  'SetCooperativeLevel method sets the cooperative level of the application for this sound device
  'the cooperative level is set by calling this method before its buffers can be played
  'MSDN Online documentation of DirectSound8 states the following:
  '"The recommended cooperative level is DSSCL_PRIORITY.
  'The hwnd parameter should be the top-level application window handle."
   sndBufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_GETCURRENTPOSITION2
   'sndBufferDesc variable of DSBUFFERDESC type structure's lFlags member specifies the desired Sound Buffer capabilites such as the control this application will have over the buffer
   'this member can specify any of the CONST_DSBCAPSFLAGS constants or can be zero
   'for more detail on this structure's(DSBUFFERDESC) members see the following
   'MSDN Online DSBUFFERDESC documentation at http://msdn.microsoft.com/library/en-us/dx8_vb/directx_vb/htm/_dx_dsbufferdesc_dxaudio_vb.asp?frame=false
    Set sndBuffer = DS8.CreateSoundBufferFromFile(WAVFile, sndBufferDesc)
    'initialize sndBuffer(DirectSoundSecondaryBuffer8 class) with the class instance returned by DS8's CreateSoundBufferFromFile method...
    'CreateSoundBufferFromFile method creates a secondary buffer and loads data from a file specified in the filename paramater into the buffer
     sndBuffer.GetFormat WaveFormat
     'initialize variable WaveFormat of WAVEFORMATEX type returned by the GetFormat method which retrieves a description of the wave format of the buffer
End Sub


Private Sub cmdPlay_Click()
On Error Resume Next 'on the event of an error resume execution of this procedure on the next line
Dim sample&, i&, tTC&, lastPos%: lastPos = 12: cmdPlay.Enabled = False: cmdStop.Enabled = True
'dimensionalize sample as long data type, i as long data type, tTC as long data type, lastPos as integer data type
'initialize lastPos to 12, update CommandButton cmdPlay, and cmdStop's enabled properties
 sndBuffer.Play DSBPLAY_DEFAULT 'sndBuffer's Play method causes the sound buffer to play from the play cursor
 'Public Enum CONST_DSBPLAYFLAGS
 ' DSBPLAY_DEFAULT = 0
 ' DSBPLAY_LOCHARDWARE = 2
 ' DSBPLAY_LOCSOFTWARE = 4
 ' DSBPLAY_LOOPING = 1
 ' DSBPLAY_TERMINATEBY_DISTANCE = 16 (&H10)
 ' DSBPLAY_TERMINATEBY_PRIORITY = 32 (&H20)
 ' DSBPLAY_TERMINATEBY_TIME = 8
 'End Enum
 'for a description on any of the members of this constant enumeration see the following online MSDN documentation
 'http://msdn.microsoft.com/library/en-us/dx8_vb/directx_vb/htm/_dx_const_dsbplayflags_dxaudio_vb.asp?frame=true
  Do While sndBuffer.GetStatus = DSBSTATUS_PLAYING
  'Do While loop, loop's while the status of the sound buffer evaluates to DSBSTATUS_PLAYING(Playing)
   DoEvents 'yield execution to other asynchronously processing procedures
    sndBuffer.GetCurrentPosition CurRWPos 'retrieve capture of the read write cursor
    'CurRWPos receives the position of the capture cursor and the read cursor, each value is an offset from the start of the buffer in bytes
    'see below for more information...
     Select Case WaveFormat.nBitsPerSample
     'select case statement, selecting the current sound files Bits Per Sample
      Case 8: 'if nBitsPerSample evaluates to 8 bits per sample then....
       sndBuffer.ReadBuffer CurRWPos.lPlay, 1, byteAr(0), DSBLOCK_DEFAULT
       'read one byte from the buffer at the current play position
        tTC = GetTickCount 'GetTickCount function retrieves the ammount of milliseconds which has elapsed since the system has started
         Do While (GetTickCount - tTC) <= 150
         'do while loop, loops until one-hundred-fifty milliseconds have elapsed
          If ((byteAr(0) / 255) * 24) > lastPos Then 'if the amplitude is inclining then...
          'element zero of array byteAr contains a single byte of the sound buffer
          'since this sample is only a (mono)single-channel sample
          'the elements value will be within the range of the byte data type range (0 - 255)
          'the amplitude of this sample can be determined by this value
           Set picMain.Picture = LoadPicture(arBMP(lastPos + Abs((((byteAr(0) / 255) * 24) - lastPos) * ((GetTickCount - tTC) / 150))))
           'determine the coinciding image path to be displayed based upon the amplitude where
           '255 = 100%, 0 = 0%;
           '0% = arBMP(0)[HeadBang 01.bmp], 100% = arBMP(24)[HeadBang 24.bmp]
            sldMain.Value = lastPos + Abs((((byteAr(0) / 255) * 24) - lastPos) * ((GetTickCount - tTC) / 150))
            '...
          Else
          'amplitude is declining...
           Set picMain.Picture = LoadPicture(arBMP(lastPos - Abs((lastPos - ((byteAr(0) / 255) * 24)) * ((GetTickCount - tTC) / 150))))
           '...
            sldMain.Value = lastPos - Abs((lastPos - ((byteAr(0) / 255) * 24)) * ((GetTickCount - tTC) / 150))
            '...
          End If
         Loop
          lastPos = ((byteAr(0) / 255) * 24)
          'update flag
      Case 16: 'multi-channel (16 bits-per-sample)
       sndBuffer.ReadBuffer CurRWPos.lPlay, 2, byteAr(0), DSBLOCK_DEFAULT '...
        sample = ((Abs(byteAr(0)) + 32768) / 65536) * 24
        'since the sample is multi-channel(stereo) two bytes will be read from
        'the sound buffer yet only the first byte will be processed.
        'the data returned by the method ReadBuffer in this case is actually returning data of signed integer data type
        'convert the data type of the data returned from signed integer data type (range: -32,768 to 32,767) to unsigned integer data type (range: 0 to 65,536)
        'initialize variable sample with the converted range of (0 - 65,536) to (0 - 24)
         tTC = GetTickCount 'return the ammount of milliseconds which has elapsed since the system started
          Do While (GetTickCount - tTC) <= 150
          'loop until 150 milliseconds elapses
           If sample > lastPos Then 'if amplitude is inclining then...
            Set picMain.Picture = LoadPicture(arBMP(lastPos + Abs((sample - lastPos) * ((GetTickCount - tTC) / 150))))
             sldMain.Value = lastPos + Abs((sample - lastPos) * ((GetTickCount - tTC) / 150))
           Else
            Set picMain.Picture = LoadPicture(arBMP(lastPos - Abs((lastPos - sample) * ((GetTickCount - tTC) / 150))))
             sldMain.Value = lastPos - Abs((lastPos - sample) * ((GetTickCount - tTC) / 150))
           End If
          Loop
           lastPos = sample
           'see documentation of 8-bit per sample sound amplitude modulation processing above(Case 8:)
     End Select
  Loop
   lastPos = 12 'initialize lastPos with the median(12) of the bitmap name range (0 - 24)
    cmdPlay.Enabled = True: cmdStop.Enabled = False
End Sub

Private Sub cmdStop_Click()
 sndBuffer.Stop 'call object sndBuffer's Stop method to stop audio output
  sndBuffer.SetCurrentPosition 0 'set the current position to the beggining of the file...
End Sub

Private Sub cmdLoad_Click()
On Error GoTo errh 'on the event of an error resume execution of this procedure at label errh
 cd.Filter = "WAV|*.wav|All Files|*.*"
 'initialize object cd's Filter property
 'the filter property of object cd(CommonDialog) specifies the file name pattern
 'syntax: File Desc|Pattern|File Desc 2|Pattern2;Pattern3
 '* = True Wildcard
 '? = Single Wildcard
 '# = Numerical character
 '[$-$] = Alphabetical character, ex: [A-C] (Case sensitive)
 '...
  cd.ShowOpen 'call object cd's ShowOpen method to show the file select(open) dialog
   WAVFile = cd.FileName 'initialize variable WAVFile to the file name of the file the user selected in the file open dialog
    InitDS 'see InitDS(Initiate DirectSound) procedure for more info...
     cmdPlay.Enabled = True 'initialize object cmdPlay's Enabled property to True to enable the window(Command Button)
     Exit Sub 'exit this procedure
errh: 'label errh
 If Err.Number = 32755 Then Exit Sub
 'if the number of the error raised was 32755 then the user cancelled the open file dialog, exit this procedure
  MsgBox "An un-expected error has occured." & vbCrLf & Err.Description, vbCritical, "Error"
  'inform user an error has occured
End Sub

Private Sub Form_Load()
Dim buffer$ 'declare buffer as string data type
   buffer = IIf(Right$(App.Path, 1) <> "\", App.Path & "\", App.Path) & "res bmp\"
   'perform IIf operation to conditionally return the path which specifies the 'res bmp' folder
   'IIf operation syntax: a = IIf(Expression to Evaluate, TruePart, FalsePart)
    arBMP(1) = buffer & "HeadBang 01.bmp"
     arBMP(2) = buffer & "HeadBang 02.bmp"
      arBMP(3) = buffer & "HeadBang 03.bmp"
       arBMP(4) = buffer & "HeadBang 04.bmp"
        arBMP(5) = buffer & "HeadBang 05.bmp"
         arBMP(6) = buffer & "HeadBang 06.bmp"
          arBMP(7) = buffer & "HeadBang 07.bmp"
           arBMP(8) = buffer & "HeadBang 08.bmp"
            arBMP(9) = buffer & "HeadBang 09.bmp"
             arBMP(10) = buffer & "HeadBang 10.bmp"
              arBMP(11) = buffer & "HeadBang 11.bmp"
               arBMP(12) = buffer & "HeadBang 12.bmp"
                arBMP(13) = buffer & "HeadBang 13.bmp"
                 arBMP(14) = buffer & "HeadBang 14.bmp"
                  arBMP(15) = buffer & "HeadBang 15.bmp"
                   arBMP(16) = buffer & "HeadBang 16.bmp"
                    arBMP(17) = buffer & "HeadBang 17.bmp"
                     arBMP(18) = buffer & "HeadBang 18.bmp"
                      arBMP(19) = buffer & "HeadBang 19.bmp"
                       arBMP(20) = buffer & "HeadBang 20.bmp"
                        arBMP(21) = buffer & "HeadBang 21.bmp"
                         arBMP(22) = buffer & "HeadBang 22.bmp"
                          arBMP(23) = buffer & "HeadBang 23.bmp"
                           arBMP(24) = buffer & "HeadBang 24.bmp"
                           'initialize the values of 24 elements in the
                           'one dimensional array arBMP
                           
                           'for more information on how this array will be used see cmdPlay_Click procedure
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Call cmdStop_Click 'see cmdStop_Click procedure for more information...
  Set DS8 = Nothing 'terminate the DirectSound8 class instance
   Set DX8 = Nothing 'terminate the DirectX8 class instance
    End 'unload all dialog resources from memory
End Sub

