Attribute VB_Name = "modClick"
Option Explicit
Option Private Module

Public VolTmp As String, LoopTmp As String
Attribute VolTmp.VB_VarUserMemId = 1073741824
Attribute LoopTmp.VB_VarUserMemId = 1073741825
Public butMenuIdx As Integer, butMenu As CheckBox
Attribute butMenuIdx.VB_VarUserMemId = 1073741826
Attribute butMenu.VB_VarUserMemId = 1073741827

Public Sub clAssign(btNum As Integer, Butt As CheckBox)

        On Error GoTo Error_Routine

1       If Jing(btNum).OnAir Then Exit Sub
        'JINGLE ASSIGNMENT TO BUTTON (if assign mode is ON)
2       frmMain.ckAssign = 0  'reset assign mode
3       On Error GoTo errc ' CancelError is True
4       frmMain.cmDlg.flags = cdlOFNHideReadOnly + cdlOFNFileMustExist
5       frmMain.cmDlg.Filter = LangSet.Entry("mDlgType", , Language) & " (*.wav;*.mp3;*.mp2;*.mp1;*.mpa;*.ogg)|*.wav;*.mp3;*.mp2;*.mp1;*.mpa;*.ogg"
6       frmMain.cmDlg.DialogTitle = LangSet.Entry("mDlgAsBu", , Language)
7       frmMain.cmDlg.InitDir = ""
8       frmMain.cmDlg.ShowOpen  'open dialog to select a file

9       If frmMain.cmDlg.FileName <> "" Then
10          Jing(btNum).Path = frmMain.cmDlg.FileName
11          frmMain.p_Jingle(btNum).Visible = False
12          Jing(btNum).Loop = False
13      End If

14      If Jing(btNum).Path <> "" Then Jing(btNum).Strm = BASS_StreamCreateFile(BASSFALSE, Jing(btNum).Path, 0, 0, 0)
15      If Jing(btNum).Strm = 0 Then
16          Call Error_("Can't create stream when assigning")
17      Else
18          BASS_ChannelPreBuf Jing(btNum).Strm
19          frmMain.b_Jingle(btNum).Caption = btCaption(Jing(btNum).Path, frmMain.b_Jingle(btNum))
20      End If

errc:
        'asignment over, cancelled
Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    err.Raise err.Number, "modClick.clAssign(" & Erl & "):" & err.Source, err.Description
End Sub

Public Sub clTmAn(btNum As Integer, Butt As CheckBox, Optional mnu As Boolean)

        'TIME ANNOUNCE JINGLE CHOOSING
        On Error GoTo Error_Routine

1       For k = 0 To 29
2           frmMain.b_Jingle(k).BackColor = Jing(k).Color
3       Next k
4       tmJing = Jing(btNum).Path
5       tmStat1 = True
6       frmMain.b_Jingle(btNum).value = vbChecked
7       frmMain.b_tmAnJin.Caption = btCaption(tmJing, frmMain.b_tmAnJin, 26) & " " & tmDel & " " & LangSet.Entry("mTaLater", , Language)
8       If mnu Then
9           frmMain.b_tmAnJin.value = vbChecked
10          frmMain.b_tmAnJin.BackColor = &HFFC0C0
11      End If
12      Call BASS_StreamFree(tmJnStrm)
13      tmJnStrm = BASS_StreamCreateFile(BASSFALSE, tmJing, 0, 0, 0)
14      If tmJnStrm = 0 Then
15          Call Error_("Can't create time announced jingle stream")
16      Else
17          BASS_ChannelPreBuf tmJnStrm
18          frmMain.b_tmAnJin.BackColor = &HFFC0C0
19      End If
20      If Butt.value = vbUnchecked Then
21          Butt.BackColor = b_color(btNum)
22      Else
23          Butt.value = vbUnchecked
24      End If

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    err.Raise err.Number, "modClick.clTmAn(" & Erl & "):" & err.Source, err.Description
End Sub

Public Sub clVolSel(btNum As Integer, Butt As CheckBox, Optional mnu As Boolean)

        'VOLUME SELECTION
        On Error GoTo Error_Routine

1       If mnu Then GoTo voll
2       If Butt.value = vbUnchecked Then
3           Butt.BackColor = b_color(btNum)
4       Else
5           Butt.value = vbUnchecked
6       End If
7 voll:
8       frmMain.lblVolexp.Caption = Butt.Caption & ":"
9       frmMain.ckAssVol.Visible = False
10      frmMain.ckAssVol.value = vbUnchecked
11      tmout = 0
12      frmMain.btAbout.Caption = LangSet.Entry("mVolSave", , Language)
13      frmMain.btAbout.MaskColor = RGB(255, 0, 255)
14      frmMain.btAbout.picture = LoadResPicture(118, vbResBitmap)
15      frmMain.SlideVolCh.pos = Jing(btNum).volume
16      frmMain.lblVol.Caption = Jing(btNum).volume - 100 & " dB"
17      VolIndex = btNum

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    err.Raise err.Number, "modClick.clVolSel(" & Erl & "):" & err.Source, err.Description
End Sub

Public Sub clLoopSet(btNum As Integer, Butt As CheckBox, Optional mnu As Boolean)

        'LOOP MODE SETTING
        On Error GoTo Error_Routine

1       If Jing(btNum).OnAir And mnu Then Exit Sub
2       If Butt.value = vbUnchecked Then
3           Butt.BackColor = b_color(btNum)
            'Butt.BackColor = &H80C0FF
4       Else
5           Butt.value = vbUnchecked
6       End If
7       If Jing(btNum).OnAir Then Exit Sub

8       frmMain.ckLoopI.value = vbUnchecked
9       Jing(btNum).Loop = Not Jing(btNum).Loop
10      frmMain.p_Jingle(btNum).Visible = Jing(btNum).Loop
11      PalSet.Entry("Loop_" & btNum, , frmMain.lstPal.Text) = CByte(Jing(btNum).Loop)

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    err.Raise err.Number, "modClick.clLoopSet(" & Erl & "):" & err.Source, err.Description
End Sub

Public Sub clJingCheck(btNum As Integer, Butt As CheckBox)

        On Error GoTo Error_Routine

1       If Jing(btNum).Path = vbNullString Then
            'NO EVENT IF NO JINGLE ASSIGNED, OR FILE MISSING
2           frmMain.b_Jingle(btNum).value = vbChecked
3           If frmMain.b_Jingle(btNum).Caption <> "" Then
4               ErrDisp "no", , "Button " & btNum + 1 & " Jingle is missing from HDD. Reassign this button"
5               ErrLog "Button " & btNum + 1, "Jingle is missing from HDD. Reassign this button"
6               frmMain.b_Jingle(btNum).Caption = ""
7           End If
8           Exit Sub 'nothing happens
9       End If
        '        If Dir(Jing(btNum).Path) = vbNullString Then
        '            'PROTECTION IN CASE OF FILE MISSING
        '            ErrDisp "no", , "Button " & btNum + 1 & " Jingle is missing from HDD. Reassign this button"
        '            ErrLog "Button " & btNum + 1, "Jingle is missing from HDD. Reassign this button"
        '            Jing(btNum).Path = ""
        '            frmMain.b_Jingle(btNum).value = 1
        '            frmMain.b_Jingle(btNum).Caption = ""
        '            Exit Sub
        '        End If

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    err.Raise err.Number, "modClick.clJingCheck(" & Erl & "):" & err.Source, err.Description
End Sub

Public Sub clPlay(btNum As Integer, Butt As CheckBox)

        On Error GoTo Error_Routine

1       If Jing(btNum).Strm <> 0 Then
2           BASS_ChannelSetAttributes Jing(btNum).Strm, -1, Jing(btNum).volume, -101 'reset volume
3           If BASS_ChannelIsActive(Jing(btNum).Strm) <> BASS_ACTIVE_PLAYING Then  'Jing(btNum).OnAir = False Then
4               Jing(btNum).OnAir = True   'set this True while on air
5               If Jing(btNum).Loop Then
6                   Call BASS_ChannelSetFlags(Jing(btNum).Strm, BASS_SAMPLE_LOOP)
7               Else
8                   Call BASS_ChannelSetFlags(Jing(btNum).Strm, 0)
9               End If
10              If BASS_ChannelPlay(Jing(btNum).Strm, BASSFALSE) = BASSFALSE Then
11                  Call Error_("Can't play jingle stream from the main button palette")
12              End If

13              If Butt.value = vbChecked Then Butt.value = vbUnchecked

14              VuStrm = Jing(btNum).Strm       'used for time countback
15              VuVol = Jing(btNum).volume

16              If j = 0 Then
17                  j = 1
18              Else
19                  j = 0
20              End If
21              aMixPrevStrm(j) = Jing(btNum).Strm
22              If aMixPrevStrm(0) <> aMixPrevStrm(1) Then Call autoMixer(Abs(j - 1), btNum)
23              frmMain.mnuPause.Caption = LangSet.Entry("mPause", , Language)

24          Else 'if click comes from the same button
25              If frmMain.ckAutoRep.value = 1 Then 'AUTOREPEAT = seek to 0
26                  Call BASS_ChannelSetPosition(Jing(btNum).Strm, 0)
27                  Jing(btNum).OnAir = True
28                  frmMain.b_Jingle(btNum).value = vbUnchecked
29              Else
30                  Jing(btNum).OnAir = False       'NO AUTOREPEAT = STOP
31                  Call BASS_ChannelStop(Jing(btNum).Strm)
32                  Call BASS_ChannelSetPosition(Jing(btNum).Strm, 0)
33              End If
34          End If
35      End If

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    err.Raise err.Number, "modClick.clPlay(" & Erl & "):" & err.Source, err.Description
End Sub

Public Sub clTouchStop(btNum As Integer, Butt As CheckBox)
        On Error GoTo Error_Routine

1       Jing(btNum).OnAir = False       'NO AUTOREPEAT = STOP
2       Call BASS_ChannelStop(Jing(btNum).Strm)
3       Call BASS_ChannelSetPosition(Jing(btNum).Strm, 0)
4       Butt.value = vbUnchecked
5       Butt.BackColor = vbButtonFace

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    err.Raise err.Number, "modClick.clTouchStop(" & Erl & "):" & err.Source, err.Description
End Sub

Public Sub openNames(Index As Integer, Section As String)
        On Error GoTo Error_Routine

1       If BASS_ChannelIsActive(Jing(Index).Strm) <> BASS_ACTIVE_PLAYING Then
2           Call BASS_StreamFree(Jing(Index).Strm)
3           Jing(Index).Path = PalSet.Entry("Path_" & Index, , Section)
4           If PalSet.EntryNotFound Then
5               Jing(Index).Path = ""
6               Jing(Index).volume = 100
7               Jing(Index).Loop = False
8           Else
9               VolTmp = PalSet.Entry("Volm_" & Index, , Section)
10              If PalSet.EntryNotFound Then
11                  Jing(Index).volume = 100
12              Else
13                  Jing(Index).volume = VolTmp
14              End If
15              LoopTmp = PalSet.Entry("Loop_" & Index, , Section)
16              If PalSet.EntryNotFound Then
17                  Jing(Index).Loop = False
18              Else
19                  Jing(Index).Loop = CBool(LoopTmp)
20              End If
21          End If
22          frmMain.b_Jingle(Index).Caption = ""
23          frmMain.b_Jingle(Index).BackColor = vbButtonFace
24          frmMain.b_Jingle(Index).value = vbUnchecked
25          frmMain.p_Jingle(Index).Visible = Jing(Index).Loop
26          frmMain.b_Jingle(Index).Caption = btCaption(Jing(Index).Path, frmMain.b_Jingle(Index)) '& Index
27      Else
28          Jing(Index).inDebt = True
29          frmMain.b_Jingle(Index).BackColor = &HFFFFFF
30      End If

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    err.Raise err.Number, "modClick.openNames(" & Erl & "):" & err.Source, err.Description
End Sub

Public Sub openStreams(Index As Integer)
        On Error GoTo Error_Routine

1       If Jing(Index).Path <> "" Then
2           If InStr(1, Jing(Index).Path, "\", vbTextCompare) = 0 Then
3               Jing(Index).Path = App.Path & "\" & Jing(Index).Path 'if there is no path for the file, assume working directory
4           End If

5           If Not Jing(Index).inDebt Then
6               Jing(Index).Strm = BASS_StreamCreateFile(BASSFALSE, Jing(Index).Path, 0, 0, 0)

7               If Jing(Index).Strm = 0 Then
8                   If BASS_ErrorGetCode() = 2 Then
9                       frmMain.b_Jingle(Index).Caption = ""
10                      Jing(Index).volume = 100
11                      Jing(Index).Loop = False
                        frmMain.p_Jingle(Index).Visible = False
12                  Else
13                      Call Error_("Can't create stream when opening palette, btNum = " & i)
14                  End If
15              Else
16                  BASS_ChannelPreBuf Jing(Index).Strm
17                  BASS_ChannelSetAttributes aMixStrm, -1, Jing(Index).volume, -101
18              End If

19          End If
20      Else
21          Jing(Index).Strm = 0
22      End If

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    err.Raise err.Number, "modClick.openStreams(" & Erl & "):" & err.Source, err.Description
End Sub
