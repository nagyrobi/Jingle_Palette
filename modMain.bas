Attribute VB_Name = "modMain"
Option Explicit
Option Private Module

Public Const GWL_WNDPROC = (-4)

'for>next cycles variable
Public Tme As Integer
Attribute Tme.VB_VarUserMemId = 1073741824
Public i As Integer
Attribute i.VB_VarUserMemId = 1073741825
Public j As Integer
Attribute j.VB_VarUserMemId = 1073741826
Public k As Integer
Attribute k.VB_VarUserMemId = 1073741827
Public l As Integer
Attribute l.VB_VarUserMemId = 1073741828
Public RunFirst As Boolean
Attribute RunFirst.VB_VarUserMemId = 1073741829
Public PaintFirst As Boolean
Attribute PaintFirst.VB_VarUserMemId = 1073741830
Public Thrd As Integer
Attribute Thrd.VB_VarUserMemId = 1073741831
Public OnTop As Boolean

'jingle Time Length and display
Public Sec As String
Attribute Sec.VB_VarUserMemId = 1073741832
Public min As String
Attribute min.VB_VarUserMemId = 1073741833
Public TimeForm As String
Attribute TimeForm.VB_VarUserMemId = 1073741834
Public DateForm As String
Attribute DateForm.VB_VarUserMemId = 1073741835
Public RemWarn As Integer
Attribute RemWarn.VB_VarUserMemId = 1073741836
Public RemColor As Long
Attribute RemColor.VB_VarUserMemId = 1073741837

'This variable keeps track of the filename information for opening and closing files.
Public FileName As String
Attribute FileName.VB_VarUserMemId = 1073741838
Public filemam As FileOpenConstants
Attribute filemam.VB_VarUserMemId = 1073741839
Public IniSet As New ImpulseRegistryAndINI
Attribute IniSet.VB_VarUserMemId = 1073741840
Public IniFile As String
Attribute IniFile.VB_VarUserMemId = 1073741841
Public PalSet As New ImpulseRegistryAndINI
Attribute PalSet.VB_VarUserMemId = 1073741842
Public PalFile As String
Attribute PalFile.VB_VarUserMemId = 1073741843
Public LangSet As New ImpulseRegistryAndINI
Attribute LangSet.VB_VarUserMemId = 1073741844
Public LangFile As String
Attribute LangFile.VB_VarUserMemId = 1073741845
Public Language As String
Attribute Language.VB_VarUserMemId = 1073741846

' Vu Meter
Public VuStrm As Long, VuVol As Long, VolIndex As Integer
Attribute VuStrm.VB_VarUserMemId = 1073741847
Attribute VuVol.VB_VarUserMemId = 1073741848
Attribute VolIndex.VB_VarUserMemId = 1073741849
Public Lft() As Long, Rht() As Long, VuLft As Long, VuRght As Long, Lvel As Double
Attribute Lft.VB_VarUserMemId = 1073741850
Public VuM() As Long
Attribute VuM.VB_VarUserMemId = 1073741855

'time announce
Public tmSound As String, tmSnStrm As Long
Attribute tmSound.VB_VarUserMemId = 1073741856
Attribute tmSnStrm.VB_VarUserMemId = 1073741857
Public tmJing As String, tmJnStrm As Long
Attribute tmJing.VB_VarUserMemId = 1073741858
Attribute tmJnStrm.VB_VarUserMemId = 1073741859
Public tmMin As Integer
Attribute tmMin.VB_VarUserMemId = 1073741860
Public tmSec As Integer
Attribute tmSec.VB_VarUserMemId = 1073741861
Public tmDel As Integer, tmDelDif As Long
Attribute tmDel.VB_VarUserMemId = 1073741862
Attribute tmDelDif.VB_VarUserMemId = 1073741863
Public tmStat As Boolean, tmStat1 As Boolean
Attribute tmStat.VB_VarUserMemId = 1073741864
Attribute tmStat1.VB_VarUserMemId = 1073741865
Public tmJnGo As Boolean

'internal player volume, other parameters
Public vol As Long
Attribute vol.VB_VarUserMemId = 1073741866
Public Cfrq As Long
Attribute Cfrq.VB_VarUserMemId = 1073741867
Public aMixtime As Integer
Attribute aMixtime.VB_VarUserMemId = 1073741868
Public aMixCount As Integer, aMixAir As Integer
Attribute aMixCount.VB_VarUserMemId = 1073741869
Attribute aMixAir.VB_VarUserMemId = 1073741870
Public aMixStrm As Long
Attribute aMixStrm.VB_VarUserMemId = 1073741871
Public aMixPrevStrm(1) As Long
Attribute aMixPrevStrm.VB_VarUserMemId = 1073741872
Public VolMix As Long
Attribute VolMix.VB_VarUserMemId = 1073741873
Public VolAmix(29) As Long
Attribute VolAmix.VB_VarUserMemId = 1073741874

'effects
Public floatable As Variant, rotdsp As Variant, fladsp As Variant
Attribute floatable.VB_VarUserMemId = 1073741875
Attribute rotdsp.VB_VarUserMemId = 1073741876
Attribute fladsp.VB_VarUserMemId = 1073741877

Public errCk As Integer
Attribute errCk.VB_VarUserMemId = 1073741878
Public LoopBit As Double
Attribute LoopBit.VB_VarUserMemId = 1073741879

'counters for timers and players
Public tmout As Integer, tm1 As Integer, tmf As Integer, tm2 As Integer
Attribute tmout.VB_VarUserMemId = 1073741880
Attribute tm1.VB_VarUserMemId = 1073741881
Attribute tmf.VB_VarUserMemId = 1073741882
Attribute tm2.VB_VarUserMemId = 1073741883
Public tmoutMnu As Integer, mnuOn As Boolean
Attribute tmoutMnu.VB_VarUserMemId = 1073741884
Attribute mnuOn.VB_VarUserMemId = 1073741885
Public flash As Boolean
Attribute flash.VB_VarUserMemId = 1073741886

Public exitBut As Boolean
Attribute exitBut.VB_VarUserMemId = 1073741887
Public ManiPulate As Boolean
Attribute ManiPulate.VB_VarUserMemId = 1073741888
Public ManiPulVol As Boolean
Attribute ManiPulVol.VB_VarUserMemId = 1073741889

'Variable with path, button, etc of jingle
Public Type JingCheck
    Path As String      'the path of the jingle
    'BtNr As Integer     'button number where it is on
    Color As ColorConstants
    OnAir As Boolean    'status if it's On Air or not
    Paused As Boolean
    Strm As Long
    volume As Long
    Loop As Boolean
    VuL As Long
    VuR As Long
    inDebt As Boolean
End Type
Public Jing(29) As JingCheck
Attribute Jing.VB_VarUserMemId = 1073741890
Public b_color(29) As ColorConstants
Attribute b_color.VB_VarUserMemId = 1073741891
Public SelPalette As String
Attribute SelPalette.VB_VarUserMemId = 1073741892
Public urlPrv As String
Attribute urlPrv.VB_VarUserMemId = 1073741893

'streams
Public StreamHandle As Long
Attribute StreamHandle.VB_VarUserMemId = 1073741894
Public Streams As Variant
Attribute Streams.VB_VarUserMemId = 1073741895
Public LiveStream As Long
Attribute LiveStream.VB_VarUserMemId = 1073741896
Public LiveOn As Boolean

'hardware
Public device As Long, devicePrev As Long
Attribute device.VB_VarUserMemId = 1073741897
Attribute devicePrev.VB_VarUserMemId = 1073741898
Public Sub ErrLog(Optional Numb As String, Optional Desc As String)
        '    frmAbout.LstErr.AddItem Date & " at " & Time & ", Code " & Numb & ": " & Desc
        '    Dim Ret As Long
        '    Ret = frmAbout.LstErr.ListCount
        '    frmAbout.LstErr.ListIndex = Ret - 1
        On Error GoTo Error_Routine

1       If errCk Then
2           Dim f As Integer
3           f = FreeFile
4           FileName = App.Path & "\Error.log"
5           Open FileName For Append As f       ' Open the filename for append.
6           Write #f, Date & " at " & time & ", Code " & Numb & ": " & Desc
7           Close f                             ' Close the file.
8       End If

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    err.Raise err.Number, "modMain.ErrLog(" & Erl & "):" & err.Source, err.Description
End Sub
Public Sub Main()

        On Error GoTo Error_Routine

1       Thrd = ThreadPriority
2       IniFile = App.Path & "\jp.ini"   ' the Settings INI file definition
3       If Dir$(IniFile) = vbNullString Then
4           Dim f As Integer
5           f = FreeFile
6           Open IniFile For Output As f
7           Close f
8           IniSet.Location = IniFile
9           IniSet.SectionKey = "Settings"
            'DefaultSettings
10          RunFirst = True
11      Else
12          IniSet.SectionKey = "Settings"
13          IniSet.Location = IniFile
14      End If

15      PalFile = App.Path & "\palette.ini"  'the Palette INI file definition
16      If Dir$(PalFile) = vbNullString Then
17          Dim e As Integer
18          e = FreeFile
19          Open PalFile For Output As e
20          Close e
21      End If
22      PalSet.Location = PalFile
        'frmSplash.vuLoad.Level = 10
23      ChDrive App.Path
24      ChDir App.Path

25      If RunFirst Then
26          DefaultSettings
27      End If

28      errCk = IniSet.Entry("ErrorLog")
29      ErrLog , "***********************************************************"
30      ErrLog , "Starting program..."
31      DoEvents

32      LoopBit = 0
        'check if 'bass21.dll' is exists
33      If Not FileExists(RPP(App.Path) & "bass21.dll") Then
34          ErrLog "bass21.dll", " bass21.dll does not exist!"
35          MsgBox "The file bass21.dll does not exist!", vbCritical, "bass21.dll"
36          End
37      End If

        'Check that BASS 2.1 was loaded
38      If BASS_GetVersion <> MakeLong(2, 1) Then
39          ErrLog "bass21.dll", " BASS version 2.1 was not loaded"
40          MsgBox "BASS module version 2.1 was not loaded", vbCritical, "bass21.dll"
41          End
42      End If
        'frmSplash.vuLoad.Level = 25
        '    Call BASS_SetConfig(BASS_CONFIG_UPDATEPERIOD, 5)

        '    'enable floating-point DSP
        '    Call BASS_SetConfig(BASS_CONFIG_FLOATDSP, BASSTRUE)
        '
        '
        '    'check for floating-point capability
        '    floatable = BASS_StreamCreate(44100, 2, BASS_SAMPLE_FLOAT, 0, 0)
        '    If (floatable) Then
        '        Call BASS_StreamFree(floatable)  'woohoo!
        '        floatable = BASS_SAMPLE_FLOAT
        '    End If

43      DoEvents

44      LangFile = App.Path & "\language.ini"
45      If Dir$(LangFile) = vbNullString Then
46          Dim l As Integer
47          l = FreeFile
48          Open LangFile For Output As l
49          Close l
50      End If
51      LangSet.Location = LangFile

52      aMixtime = IniSet.Entry("AutoMixTime")
53      TimeForm = IniSet.Entry("TimeFormat")
54      DateForm = IniSet.Entry("DateFormat")
55      RemWarn = IniSet.Entry("RemainWarn")
56      RemColor = IniSet.Entry("RemainColor")
57      tmSound = IniSet.Entry("TimeAnnouncer")
58      tmMin = IniSet.Entry("TimeAnnMin")
59      tmSec = IniSet.Entry("TimeAnnSec")
60      tmDel = IniSet.Entry("TimeAnnDel")
61      vol = IniSet.Entry("Volume")
62      device = IniSet.Entry("Device")
63      devicePrev = device
64      Language = IniSet.Entry("Language")
        OnTop = IniSet.Entry("AlwaysOnTop")
        
65      For k = 0 To 29
66          Jing(k).Color = vbButtonFace
67          b_color(k) = vbButtonFace
68      Next k
        'frmSplash.vuLoad.Level = 37
69      frmMain.Show

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "modMain.Main"
    Resume Exit_Routine
End Sub
Sub CloseFile(Section As String)

        On Error GoTo Error_Routine

        If Section = "" Then Exit Sub

1       ErrDisp "save", False
2       For i = 0 To 29
3           If Jing(i).Path <> vbNullString Then
4               PalSet.Entry("Path_" & i, , Section) = Jing(i).Path
5               PalSet.Entry("Volm_" & i, , Section) = Jing(i).volume
6               PalSet.Entry("Loop_" & i, , Section) = CByte(Jing(i).Loop)
7           End If
8       Next i
9       frmMain.lstPalRefresh
10      frmMain.lstPal.Text = Section

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    err.Raise err.Number, "modMain.CloseFile(" & Erl & "):" & err.Source, err.Description
End Sub
Sub OpenFile(Section As String)

        On Error GoTo Error_Routine

1       DoEvents
2       For i = 0 To 29
3           Call openNames(i, Section)
4       Next i
5       DoEvents
6       For i = 0 To 29
7           Call openStreams(i)
8       Next i
9       DoEvents
10      frmMain.ckAssign = 0

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    err.Raise err.Number, "modMain.OpenFile(" & Erl & "):" & err.Source, err.Description
End Sub

Public Function btCaption(jin As String, Optional btName As Control, Optional cLun As Integer = 60, Optional Slsh As String = "\") As String   ' FUNCTIE PENRTU NUMIREA BUTOANELOR CU NUMELE FISIERELOR

        On Error GoTo Error_Routine

1       Dim jnamext As String

2       If (Len(jin) < 8 And InStr(jin, ":") <> 0) Or jin = vbNullString Then     ' skip empty name error
3           btCaption = ""
4           Exit Function
5       Else
6           jnamext = Replace(FileNameOnly(jin), "_", " ", , , vbTextCompare)
7           btCaption = Left$(StrConv(StrConv(Trim$(Left$(jnamext, Len(jnamext) - 4)), 2), 3), cLun)
            'btName.Caption = btCaption
8       End If

Exit_Routine:
    Exit Function
Error_Routine:
    Debug.Assert False
    err.Raise err.Number, "modMain.btCaption(" & Erl & "):" & err.Source, err.Description
End Function

Public Sub autoMixer(Optional t As Integer, Optional remainsOnair As Integer)
        On Error GoTo Error_Routine

1       If frmMain.ckAutoMix Then
2           aMixCount = 0
3           frmMain.Timer4.Enabled = True
4           aMixStrm = aMixPrevStrm(t)
5           aMixAir = remainsOnair
            'VuStrm = Jing(aMixAir).Strm
6       Else
7           Exit Sub
8       End If

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    err.Raise err.Number, "modMain.autoMixer(" & Erl & "):" & err.Source, err.Description
End Sub

Public Sub btClic(btNum As Integer, Butt As CheckBox)
        ' If there is an error, display the error message below.
        On Error GoTo Error_Routine

1       If frmMain.ckAssign = 1 Then
2           Call clAssign(btNum, Butt)  'enter jingle assign mode

3       ElseIf frmMain.b_tmAnJin.value = vbChecked And Butt.Caption <> "" And frmMain.b_Jingle(0).BackColor <> vbButtonFace Then
4           Call clTmAn(btNum, Butt)    'choose the time announce jingle

5       ElseIf frmMain.ckAssVol.value = vbChecked And Butt.Caption <> "" Then
6           Call clVolSel(btNum, Butt)  'select individual volume level

7       ElseIf frmMain.ckLoopI.value = vbChecked And Butt.Caption <> "" Then
8           Call clLoopSet(btNum, Butt) 'select individual loop mode

9       ElseIf frmMain.ckTouch.value = vbChecked Then
10          Call clJingCheck(btNum, Butt)   'check if jingle is valid
11          Call clPlay(btNum, Butt)        'start playback]
12          Butt.BackColor = RGB(221, 255, 221) '&HC0FFC0
13      Else
14          Call clJingCheck(btNum, Butt)   'check if jingle is valid
15          Call clPlay(btNum, Butt)        'start playback
16      End If

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    err.Raise err.Number, "modMain.btClic(" & Erl & "):" & err.Source, err.Description
End Sub

Sub DefaultSettings()

        On Error GoTo Error_Routine

1       IniSet.SectionKey = "Settings"
2       ChDrive App.Path
3       ChDir App.Path
4       IniSet.Entry("Path") = App.Path
5       IniSet.Entry("Device") = 1
6       IniSet.Entry("Volume") = 100
7       IniSet.Entry("Refresh") = 100
8       IniSet.Entry("AutoMixTime") = 1500
9       IniSet.Entry("Touch") = 0
10      IniSet.Entry("Autorepeat") = 0
11      IniSet.Entry("AutoMix") = 0
12      IniSet.Entry("TimeFormat") = "HH:mm:ss"
13      IniSet.Entry("DateFormat") = "yyyy mmmm d"
14      IniSet.Entry("PaletteIndex") = "Demo"
15      IniSet.Entry("ErrorLog") = 0
16      IniSet.Entry("RemainWarn") = 3
17      IniSet.Entry("RemainColor") = &HC0&
18      IniSet.Entry("ActiveTab") = 0
19      IniSet.Entry("Language") = "English"
20      IniSet.Entry("WinWidth") = 12000
21      IniSet.Entry("WinHeight") = 8565
22      IniSet.Entry("WinState") = 0
        IniSet.Entry("AlwaysOnTop") = False


23      If Dir$("Time_Announce.wav") <> vbNullString Then
24          tmSound = App.Path & "\" & "Time_Announce.wav"
25          IniSet.Entry("TimeAnnouncer") = tmSound
26      Else
27          tmSound = ""
28          IniSet.Entry("TimeAnnouncer") = ""
29      End If
30      IniSet.Entry("TimeAnnMin") = "59"
31      IniSet.Entry("TimeAnnSec") = "55"
32      IniSet.Entry("TimeAnnDel") = "4"

33      ErrLog , "Default Registry Settings Loaded"

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    err.Raise err.Number, "modMain.DefaultSettings(" & Erl & "):" & err.Source, err.Description
End Sub

Public Sub ErrDisp(picture As String, Optional flsh As Boolean, Optional mesg As String)
        'in wingdings, use these: <         ý         ?          x              6
        ' these are             folppy   X framed  checked    X framed(oth)  hourglass
        On Error GoTo Error_Routine

1       tmout = 0
2       If flsh Then
3           flash = True
4       Else
5           flash = False
6       End If
7       frmMain.lblErr.Font = "Wingdings"
8       frmMain.lblErr.FontSize = 20
9       Select Case picture
            Case "save"
10              frmMain.lblErr.Caption = "<"
11          Case "wait"
12              frmMain.lblErr.Caption = "6"
13          Case "ok"
14              frmMain.lblErr.Caption = "?"
15          Case "no"
16              frmMain.lblErr.Caption = "x"
17      End Select
18      frmMain.lblErr.ToolTipText = mesg
19      If picture = "hide" Then
20          frmMain.lblErr.Visible = False
21      Else
22          frmMain.lblErr.Visible = True
23      End If

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    err.Raise err.Number, "modMain.ErrDisp(" & Erl & "):" & err.Source, err.Description
End Sub

Public Sub InitDev()
        'Start digital output
        On Error GoTo Error_Routine

1       If (BASS_Init(device, 44100, 0, frmMain.hwnd, 0) = 0) Then
2           ErrLog "Device: " & device, " Couldn't initialize Digital Output for the selected device. Using previously selected one."
3           MsgBox "Couldn't initialize Digital Output for the selected device." & vbCrLf & "Using previously selected one.", vbCritical, "Jingle Palette Error"
4           device = devicePrev
5           If (BASS_Init(device, 44100, 0, frmMain.hwnd, 0) = 0) Then
6               ErrLog "Device: " & device, " Couldn't initialize Digital Output for the previous device. Using first."
7               MsgBox "Couldn't initialize Digital Output for the previous device." & vbCrLf & "Using first.", vbCritical, "Jingle Palette Error"
8               device = 1
9               If (BASS_Init(device, 44100, 0, frmMain.hwnd, 0) = 0) Then
10                  ErrLog "Device: " & device, " Couldn't initialize Digital Output for the first device. Something's really wrong, check your system. Now running without sound."
11                  MsgBox "Couldn't initialize Digital Output for the first device." & vbCrLf & "Something's really wrong, check your system." & vbCrLf & "Now running without sound.", vbCritical, "Jingle Palette Error"
12                  Call BASS_Init(0, 44100, 0, frmMain.hwnd, 0)
13              End If
14          End If
15      End If

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    err.Raise err.Number, "modMain.InitDev(" & Erl & "):" & err.Source, err.Description
End Sub

Public Sub UrlAdd(urlText As String)
        On Error GoTo Error_Routine

1       If urlText = urlPrv Then Exit Sub
2       frmMain.cmbStream.AddItem urlText
3       Dim f As Integer
4       f = FreeFile
5       FileName = App.Path & "\urls.txt"
6       Open FileName For Append As f       ' Open the filename for output.
7       Write #f, urlText                   ' Write variables to the opened file.
8       Close f                             ' Close the file.
9       urlPrv = urlText

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    err.Raise err.Number, "modMain.UrlAdd(" & Erl & "):" & err.Source, err.Description
End Sub

Public Sub UrlLoad()
        On Error GoTo Error_Routine

1       If Dir$(App.Path & "\urls.txt") = vbNullString Then Exit Sub
2       Dim f As Integer, urLd As String
3       f = FreeFile
4       FileName = App.Path & "\urls.txt"
5       Open FileName For Input As f      ' Open the file selected in the File Open About dialog box.
6       Do Until EOF(f)
7           Input #f, urLd
8           frmMain.cmbStream.AddItem urLd
9       Loop
10      Close f    ' Close the file.
11      frmMain.cmbStream.ListIndex = frmMain.cmbStream.ListCount - 1
12      urlPrv = frmMain.cmbStream.Text

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    err.Raise err.Number, "modMain.UrlLoad(" & Erl & "):" & err.Source, err.Description
End Sub

' *********************************************
' Check for existance of any file by attempting to open it.
' *********************************************
Public Function FileExists(ByVal FileName As String) As Boolean
1       Dim f%: f = FreeFile
        On Error GoTo FileDoesntExist
2       Open FileName For Input As #f: Close #f
3       FileExists = True: Exit Function
4 FileDoesntExist:
5       FileExists = False
End Function

