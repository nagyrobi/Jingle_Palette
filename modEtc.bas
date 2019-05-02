Attribute VB_Name = "modEtc"
Option Explicit
      Public Const SWP_NOMOVE = 2
      Public Const SWP_NOSIZE = 1
      Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
      Public Const HWND_TOPMOST = -1
      Public Const HWND_NOTOPMOST = -2

      Declare Function SetWindowPos Lib "user32" _
            (ByVal hwnd As Long, _
            ByVal hWndInsertAfter As Long, _
            ByVal x As Long, _
            ByVal y As Long, _
            ByVal cx As Long, _
            ByVal cy As Long, _
            ByVal wFlags As Long) As Long

      Public Function SetTopMostWindow(hwnd As Long, Topmost As Boolean) _
         As Long

         If Topmost = True Then 'Make the window topmost
            SetTopMostWindow = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, _
               0, FLAGS)
         Else
            SetTopMostWindow = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, _
               0, 0, FLAGS)
            SetTopMostWindow = False
         End If
      End Function




Function BASS_GetStringVersion() As String
        'This function will return the string version
        'of the BASS DLL. For example the provided function within the DLL
        '"BASS_GetVersion" will return 393216, whereas this function works
        'out the actual version string as you would need to see it.
        On Error GoTo Error_Routine

1       BASS_GetStringVersion = Trim$(Str$(LoWord(BASS_GetVersion))) & "." & Trim$(Str$(HiWord(BASS_GetVersion)))

Exit_Routine:
    Exit Function
Error_Routine:
    Debug.Assert False
    err.Raise err.Number, "modBass.BASS_GetStringVersion(" & Erl & "):" & err.Source, err.Description
End Function




Public Function RPP(ByVal fp As String) As String
        On Error GoTo Error_Routine

1       RPP = IIf(Mid$(fp, Len(fp), 1) = "\", fp, fp & "\")

Exit_Routine:
    Exit Function
Error_Routine:
    Debug.Assert False
    err.Raise err.Number, "modBass.RPP(" & Erl & "):" & err.Source, err.Description
End Function

Public Function FileExists(ByVal FileName As String) As Boolean
        On Error GoTo Error_Routine

1       On Local Error Resume Next
2       FileExists = (Dir$(FileName) <> "")

Exit_Routine:
    Exit Function
Error_Routine:
    Debug.Assert False
    err.Raise err.Number, "modBass.FileExists(" & Erl & "):" & err.Source, err.Description
End Function

Public Function GetFileName(ByVal fp As String) As String
        On Error GoTo Error_Routine

1       GetFileName = Mid$(fp, InStrRev(fp, "\") + 1)

Exit_Routine:
    Exit Function
Error_Routine:
    Debug.Assert False
    err.Raise err.Number, "modBass.GetFileName(" & Erl & "):" & err.Source, err.Description
End Function

Public Sub Error_(ByVal es As String)
        'Call MsgBox(es & vbCrLf & "(error code: " & BASS_ErrorGetCode() & ")", vbExclamation, "Error")
        On Error GoTo Error_Routine

1       ErrLog "Bass, code: " & BASS_ErrorGetCode(), ", " & es
2       ErrDisp "no", False, "Bass, code: " & BASS_ErrorGetCode() & ", " & es

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    err.Raise err.Number, "modBass.Error_(" & Erl & "):" & err.Source, err.Description
End Sub

Public Sub ErrorLive_(ByVal es As String)
        On Error GoTo Error_Routine

1       Dim lvMes As String
2       Select Case BASS_ErrorGetCode()
            Case 20
3               lvMes = LangSet.Entry("mStrEr20", , Language)
4           Case 2
5               lvMes = LangSet.Entry("mStrEr2", , Language)
6           Case 32
7               lvMes = LangSet.Entry("mStrEr32", , Language)
8           Case 40
9               lvMes = LangSet.Entry("mStrEr40", , Language)
10          Case 41
11              lvMes = LangSet.Entry("mStrEr41", , Language)
12          Case 6
13              lvMes = LangSet.Entry("mStrEr6", , Language)
14          Case 1
15              lvMes = LangSet.Entry("mStrEr1", , Language)
16          Case 21
17              lvMes = LangSet.Entry("mStrEr21", , Language)
18          Case -1
19              lvMes = LangSet.Entry("mStrEr-1", , Language)
20          Case Else
21              lvMes = LangSet.Entry("mStrEr", , Language) & BASS_ErrorGetCode()
22      End Select
23      ErrLog "Bass, code: " & BASS_ErrorGetCode() & lvMes, ", " & es
24      frmMain.txtConDisp.Text = es & ": " & lvMes

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    err.Raise err.Number, "modBass.ErrorLive_(" & Erl & "):" & err.Source, err.Description
End Sub
