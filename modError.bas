Attribute VB_Name = "modError"
Option Explicit
Private m_sDesc As String
Attribute m_sDesc.VB_VarUserMemId = 1073741824
Private m_lNumb As Long
Attribute m_lNumb.VB_VarUserMemId = 1073741825
Private m_sSrc As String
Attribute m_sSrc.VB_VarUserMemId = 1073741826
Private Declare Sub APISleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
Attribute APISleep.VB_UserMemId = 1879048192
Public Sub ErrorSave()
        ' this is used when adding error hanlding to Class_Terminate
1       With err
2           m_sDesc = .Description
3           m_lNumb = .Number
4           m_sSrc = .Source
5       End With
End Sub
Public Sub ErrorRestore()
        ' this is used when adding error hanlding to Class_Terminate
1       With err
2           .Clear
3           .Description = m_sDesc
4           .Number = m_lNumb
5           .Source = m_sSrc
6       End With
End Sub
Public Sub ErrorLog(sFunctionName As String)
        ' an example of an error logging routine
1       Dim momEnt As String
2       Dim sError As String
3       Dim nError As Long
4       Dim sSource As String
5       Dim sText As String
6       Dim nErl As Long
7       With err
8           sError = .Description
9           nError = .Number
10          sSource = .Source
11      End With
12      momEnt = Now
13      nErl = Erl

        ' reset the saved error stuff!
14      m_sDesc = ""
15      m_lNumb = 0
16      m_sSrc = ""

        On Error Resume Next   ' don't put error stuff earlier - it will reset the error!

        ' here would be a good place to place any rollbacks etc.

17      sText = "An error has occurred.  " & _
            "Contact support with the following information:" & vbCrLf & momEnt & vbCrLf & vbCrLf & _
            vbTab & "Function: " & sFunctionName & vbCrLf & _
            vbTab & "Error Source: " & vbCrLf & vbTab & vbTab & _
            Replace(sSource, ":", vbCrLf & vbTab & vbTab) & vbCrLf & _
            vbTab & "Error Number: " & nError & vbCrLf & _
            vbTab & "Error Description: " & sError & vbCrLf

18      If nErl <> 0 Then
19          sText = sText & vbTab & "Last known line number: " & nErl & vbCrLf & vbCrLf & "THIS INFORMATION HAS BEEN COPIED TO THE CLIPBOARD." & vbCrLf & "_________________________________________________________"
20      End If

21      Clipboard.Clear
22      Clipboard.SetText sText

23      Debug.Print
24      Debug.Print "----------------------------------------"
25      Debug.Print sText

26      MsgBox sText, vbCritical
27      App.LogEvent sText

28      Call ErrSaveToFile(sText)

        ' reset the error
29      err.Clear

End Sub
Private Function InDebug() As Boolean
        ' sure fire way of determining if we are in development environment
        ' in EXE mode the debug.print is not executed
        ' static variable saves a tiny bit of time
        ' (especially if you "Break on All Errors")
        On Error Resume Next
1       Static bInDebug As Variant
2       If Not IsEmpty(bInDebug) Then
3           InDebug = bInDebug
4       Else
5           Debug.Print 1 / 0
6           bInDebug = err.Number <> 0
7           InDebug = bInDebug
8       End If
9       err.Clear
End Function

Public Function ErrSaveToFile(ErrReport As String, Optional ByVal ErrFileName As String = "Errors.txt") As Boolean
1       Dim F As Long, FName As String
        On Error GoTo errHandler
2       If ErrReport <> "" Then
3           F = FreeFile
4           FName = IIf(InStr(1, ErrFileName, "\") > 0, ErrFileName, App.Path & "\" & App.EXEName & "_" & ErrFileName)
5           Open ErrFileName For Append As #F
6           Print #F, ErrReport & vbNewLine & vbNewLine
7           Close #F
8       End If 'mErrReport...
9       ErrSaveToFile = True
10      Exit Function
11 errHandler:
        'nothing to do...
End Function

