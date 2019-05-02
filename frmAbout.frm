VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Jingle Palette"
   ClientHeight    =   2775
   ClientLeft      =   2490
   ClientTop       =   3630
   ClientWidth     =   8760
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1915.354
   ScaleMode       =   0  'User
   ScaleWidth      =   8226.091
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraAbt 
      Caption         =   "Information"
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      Begin VB.TextBox web 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   5760
         Locked          =   -1  'True
         MouseIcon       =   "frmAbout.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Text            =   "http://www.horvark.hu/jinglepalette"
         ToolTipText     =   "Click to visit the website"
         Top             =   720
         Width           =   2655
      End
      Begin VB.CommandButton cmdOK 
         Cancel          =   -1  'True
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   1035
         Left            =   7080
         MaskColor       =   &H00FF00FF&
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1440
         UseMaskColor    =   -1  'True
         Width           =   1260
      End
      Begin VB.TextBox email 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   6840
         Locked          =   -1  'True
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Text            =   "horvark@gmail.com"
         ToolTipText     =   "Double-click to send e-mail using your default email client"
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lblIntLang 
         Caption         =   "English interface by: Horváth Árkosi Róbert"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   8295
      End
      Begin VB.Label lblDescription 
         Alignment       =   1  'Right Justify
         Caption         =   "Email:"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   2
         Left            =   5760
         TabIndex        =   6
         Top             =   480
         Width           =   1005
      End
      Begin VB.Label lblCop 
         Alignment       =   1  'Right Justify
         Caption         =   "This program is using the"
         Height          =   240
         Left            =   120
         TabIndex        =   9
         Top             =   2280
         Width           =   2445
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Bass audio library, (c) 1999-2004 Ian Luck. All rights reserved."
         Height          =   195
         Left            =   2520
         TabIndex        =   15
         Top             =   2280
         Width           =   4455
      End
      Begin VB.Label lblCaci 
         Caption         =   "Nagy Attila"
         Height          =   255
         Left            =   1560
         TabIndex        =   14
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label lblRobi 
         Caption         =   "Horváth Árkosi Róbert"
         Height          =   255
         Left            =   1560
         TabIndex        =   13
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label lblDescription 
         Alignment       =   1  'Right Justify
         Caption         =   "Web:"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   3
         Left            =   4680
         TabIndex        =   11
         Top             =   720
         Width           =   1035
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   8400
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label lblThx 
         Caption         =   "Special thanks to:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblDescription 
         Caption         =   "Created by:"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1485
      End
      Begin VB.Label lblLic 
         Alignment       =   1  'Right Justify
         Caption         =   $"frmAbout.frx":0152
         Height          =   855
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   6855
      End
      Begin VB.Label lblDescription 
         Alignment       =   1  'Right Justify
         Caption         =   "This is an instant jingle player designed for radio studios."
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   3240
         TabIndex        =   3
         Top             =   240
         Width           =   5205
      End
      Begin VB.Label lblVersion 
         Caption         =   "Version"
         Height          =   225
         Left            =   1560
         TabIndex        =   2
         Top             =   240
         Width           =   1605
      End
      Begin VB.Label lblTitle 
         Caption         =   "AppTitle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1365
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function inStringsA(LangSec As String)
        On Error GoTo Error_Routine

1       Dim lnBuf As String
2       LangSet.SectionKey = LangSec

3       lnBuf = LangSet.Entry("alblLic")
4       If Not LangSet.EntryNotFound Then frmAbout.lblLic.Caption = lnBuf

5       lnBuf = LangSet.Entry("afrmAbout")
6       If Not LangSet.EntryNotFound Then frmAbout.Caption = lnBuf

7       lnBuf = LangSet.Entry("acmdOK")
8       If Not LangSet.EntryNotFound Then frmAbout.cmdOK.Caption = lnBuf

9       lnBuf = LangSet.Entry("alblCop")
10      If Not LangSet.EntryNotFound Then frmAbout.lblCop.Caption = lnBuf

11      lnBuf = LangSet.Entry("alblThx")
12      If Not LangSet.EntryNotFound Then frmAbout.lblThx.Caption = lnBuf

13      For i = 0 To 3
14          lnBuf = LangSet.Entry("alblDescription" & i)
15          If Not LangSet.EntryNotFound Then frmAbout.lblDescription(i).Caption = lnBuf
16      Next i

17      lnBuf = LangSet.Entry("afraAbt")
18      If Not LangSet.EntryNotFound Then fraAbt.Caption = lnBuf

19      lnBuf = LangSet.Entry("alblIntLang")
20      If Not LangSet.EntryNotFound Then lblIntLang.Caption = lnBuf

Exit_Routine:
    Exit Function
Error_Routine:
    Debug.Assert False
    err.Raise err.Number, "frmAbout.inStringsA(" & Erl & "):" & err.Source, err.Description
End Function

'Public Function outStringsA(LangSec As String)
'    LangFile = App.Path & "\language.ini"
'    If Dir$(LangFile) = vbNullString Then
'        Dim l As Integer
'        l = FreeFile
'        Open LangFile For Output As l
'        Close l
'    End If
'    LangSet.Location = LangFile
'    LangSet.SectionKey = LangSec
'    LangSet.Entry("alblLic") = lblLic.Caption
'    LangSet.Entry("afrmAbout") = frmAbout.Caption
'    LangSet.Entry("acmdOK") = cmdOK.Caption
'    LangSet.Entry("alblCop") = lblCop.Caption
'    LangSet.Entry("alblThx") = lblThx.Caption
'    For i = 0 To 3
'    LangSet.Entry("alblDescription" & i) = lblDescription(i).Caption
'    Next i
'End Function

Private Sub cmdOK_Click()
        On Error GoTo Error_Routine

1       Unload Me

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmAbout.cmdOK_Click"
    Resume Exit_Routine
End Sub
Private Sub email_DblClick()
        On Error GoTo Error_Routine

1       ExecuteDocument "mailto:" & email.Text & "?subject=" & lblTitle.Caption & " " & lblVersion.Caption, impShowNormal

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmAbout.email_DblClick"
    Resume Exit_Routine
End Sub

Private Sub Form_Load()
        'Me.Caption = "About " & App.Title
        On Error GoTo Error_Routine
        Call SetTopMostWindow(frmAbout.hwnd, OnTop)
'        frmAbout.Left = frmMain.Width '/ 2 - frmAbout.Width / 2

1       lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
2       lblTitle.Caption = App.Title
3       cmdOK.picture = LoadResPicture(118, vbResBitmap)
4       email.MouseIcon = LoadResPicture(101, vbResCursor)
5       web.MouseIcon = LoadResPicture(101, vbResCursor)
6       If Language <> "English" Then
7           Call inStringsA(Language)
8       End If

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmAbout.Form_Load"
    Resume Exit_Routine
End Sub



Private Sub web_Click()
        On Error GoTo Error_Routine

1       ExecuteDocument web.Text, impShowMaximized

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmAbout.web_Click"
    Resume Exit_Routine
End Sub
