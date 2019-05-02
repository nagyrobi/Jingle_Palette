Attribute VB_Name = "modResize"

' * * * * * * * * * * Caution * * * * * * * * * * * * *
' Changes made to the functions contained herein can cause VB to crash!
' SAVE YOUR CHANGES BEFORE RUNNING THIS PROGRAM IN THE VB IDE!
' DO NOT ENTER BREAK MODE! DOING SO WILL CRASH VB!
' * * * * * * * * * * Caution * * * * * * * * * * * * *

Option Explicit
Public OldWindowProc As Long  ' Original window proc
Attribute OldWindowProc.VB_VarUserMemId = 1073741824

' Function to retrieve the address of the current Message-Handling routine
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Attribute GetWindowLong.VB_UserMemId = 1879048192
' Function to define the address of the Message-Handling routine
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Attribute SetWindowLong.VB_UserMemId = 1879048228
' Function to copy an object/variable/structure passed by reference onto a variable of your own
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Attribute CopyMemory.VB_UserMemId = 1879048264
' Function to execute a function residing at a specific memory address
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lparam As Long) As Long
Attribute CallWindowProc.VB_UserMemId = 1879048300

' This is the message constant
Public Const WM_GETMINMAXINFO = &H24

' This is a structure referenced by the MINMAXINFO structure
Type POINTAPI
    X As Long
    Y As Long
End Type

' This is the structure that is passed by reference (ie an address) to your message handler
' The key items in this structure are ptMinTrackSize and ptMaxTrackSize
Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type
Public Function SubClass1_WndMessage(ByVal hwnd As Long, ByVal Msg As Long, ByVal wp As Long, ByVal lp As Long) As Long

        ' Watch for the pertinent message to come in
        On Error GoTo Error_Routine

1       If Msg = WM_GETMINMAXINFO Then

2           Dim MinMax As MINMAXINFO

            ' This is necessary because the structure was passed by its address and there
            ' is currently no intrinsic way to use an address in Visual Basic
3           CopyMemory MinMax, ByVal lp, Len(MinMax)

            ' This is where you set the values of the MinX,MinY,MaxX, and MaxY
            ' The values placed in the structure must be in pixels. The values
            ' normally used in Visual Basic are in twips. The conversion is as follows:
            ' pixels = twips\twipsperpixel
4           MinMax.ptMinTrackSize.X = 12000 \ Screen.TwipsPerPixelX
5           MinMax.ptMinTrackSize.Y = 8565 \ Screen.TwipsPerPixelY
6           MinMax.ptMaxTrackSize.X = Screen.Width \ Screen.TwipsPerPixelX
7           MinMax.ptMaxTrackSize.Y = Screen.Height \ Screen.TwipsPerPixelY

            ' Here we copy the datastructure back up to the address passed in the parameters
            ' because Windows will look there for the information.
8           CopyMemory ByVal lp, MinMax, Len(MinMax)

            ' This message tells Windows that the message was handled successfully
9           SubClass1_WndMessage = 1
10          Exit Function

11      End If

        ' Here, we forward all irrelevant messages on to the default message handler.
12      SubClass1_WndMessage = CallWindowProc(OldWindowProc, hwnd, Msg, wp, lp)

Exit_Routine:
    Exit Function
Error_Routine:
    Debug.Assert False
    err.Raise err.Number, "modResize.SubClass1_WndMessage(" & Erl & "):" & err.Source, err.Description
End Function

Public Function IsNotIDE() As Boolean

        On Error GoTo errHandler
        'because debug statements are ignored when
        'the app is compiled, the next statment will
        'never be executed in the EXE.
1       Debug.Print 1 / 0
2       IsNotIDE = True
3       Exit Function
4 errHandler:
        'If we get an error then we are
        'running in IDE / Debug mode
5       IsNotIDE = False
End Function

