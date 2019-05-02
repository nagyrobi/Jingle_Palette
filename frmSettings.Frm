VERSION 5.00
Object = "{F9B85A83-4DBD-11CF-94E5-0000C0571740}#1.0#0"; "NSLIDE32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSett 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuration Settings"
   ClientHeight    =   3495
   ClientLeft      =   3420
   ClientTop       =   3255
   ClientWidth     =   6120
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSettings.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3495
   ScaleMode       =   0  'User
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer 
      Interval        =   1
      Left            =   4680
      Top             =   2520
   End
   Begin TabDlg.SSTab TabSt 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5953
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Display "
      TabPicture(0)   =   "frmSettings.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblTDformat"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDDformat"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblRemWar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblRemCol"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblDispRef"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblAmx"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "UpDown1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "UpDown2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "UpDown3"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "TxtRefr"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "CmbTime"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "CmbDate"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "TxtWarTim"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "ColSel"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "TxtAmix"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Audio  "
      TabPicture(1)   =   "frmSettings.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblTasig"
      Tab(1).Control(1)=   "lblDevice"
      Tab(1).Control(2)=   "lblGenVol"
      Tab(1).Control(3)=   "SlideVol"
      Tab(1).Control(4)=   "btSound"
      Tab(1).Control(5)=   "CmbDev"
      Tab(1).Control(6)=   "txtVol"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Others  "
      TabPicture(2)   =   "frmSettings.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblProcPri"
      Tab(2).Control(1)=   "lblLang"
      Tab(2).Control(2)=   "CkErrLog"
      Tab(2).Control(3)=   "cmbThread"
      Tab(2).Control(4)=   "cmbLang"
      Tab(2).Control(5)=   "ckOnTop"
      Tab(2).ControlCount=   6
      Begin VB.CheckBox ckOnTop 
         Caption         =   "Always On Top of all the other windows"
         Height          =   255
         Left            =   -74880
         TabIndex        =   12
         Top             =   2160
         Width           =   4095
      End
      Begin VB.ComboBox cmbLang 
         Height          =   315
         Left            =   -72480
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtVol 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   -71280
         Locked          =   -1  'True
         MousePointer    =   99  'Custom
         TabIndex        =   28
         Text            =   "0 dB"
         ToolTipText     =   "Click to reset to 0dB"
         Top             =   360
         Width           =   615
      End
      Begin VB.ComboBox CmbDev 
         Height          =   315
         Left            =   -74880
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2760
         Width           =   3615
      End
      Begin VB.TextBox TxtAmix 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3360
         TabIndex        =   4
         Top             =   2400
         Width           =   600
      End
      Begin VB.PictureBox ColSel 
         Height          =   285
         Left            =   3360
         ScaleHeight     =   225
         ScaleWidth      =   795
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Click to select color, Right-click to reset to default"
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox TxtWarTim 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3360
         TabIndex        =   3
         Top             =   1440
         Width           =   600
      End
      Begin VB.ComboBox CmbDate 
         Height          =   315
         ItemData        =   "frmSettings.frx":0496
         Left            =   2520
         List            =   "frmSettings.frx":04A3
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   960
         Width           =   1695
      End
      Begin VB.ComboBox CmbTime 
         Height          =   315
         ItemData        =   "frmSettings.frx":04D0
         Left            =   2520
         List            =   "frmSettings.frx":04DA
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox TxtRefr 
         Alignment       =   1  'Right Justify
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   3360
         TabIndex        =   5
         Top             =   2880
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.ComboBox cmbThread 
         Height          =   315
         ItemData        =   "frmSettings.frx":04F6
         Left            =   -72480
         List            =   "frmSettings.frx":050F
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   480
         Width           =   1695
      End
      Begin VB.CheckBox CkErrLog 
         Caption         =   "Save all errors to ""Error.log"" file"
         Height          =   255
         Left            =   -74880
         TabIndex        =   11
         Top             =   1680
         Width           =   4095
      End
      Begin VB.CommandButton btSound 
         Height          =   435
         Left            =   -74880
         TabIndex        =   6
         ToolTipText     =   "Assigned time announce sound. Click to assign another"
         Top             =   1080
         Width           =   3615
      End
      Begin MSComCtl2.UpDown UpDown3 
         Height          =   285
         Left            =   3960
         TabIndex        =   18
         Top             =   2400
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   100
         BuddyControl    =   "CmbTime"
         BuddyDispid     =   196618
         OrigLeft        =   3480
         OrigTop         =   3120
         OrigRight       =   3735
         OrigBottom      =   3495
         Increment       =   10
         Max             =   3000
         Min             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   285
         Left            =   3960
         TabIndex        =   19
         Top             =   1440
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "TxtWarTim"
         BuddyDispid     =   196616
         OrigLeft        =   2760
         OrigTop         =   1320
         OrigRight       =   3015
         OrigBottom      =   1575
         Max             =   59
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   3960
         TabIndex        =   20
         Top             =   2880
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   50
         BuddyControl    =   "UpDown3"
         BuddyDispid     =   196637
         OrigLeft        =   2880
         OrigTop         =   960
         OrigRight       =   3135
         OrigBottom      =   1215
         Max             =   1000
         Min             =   50
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin SliderLib.Slider SlideVol 
         Height          =   2655
         Left            =   -71040
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Main Volume"
         Top             =   600
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   4683
         _StockProps     =   1
         BorderColor     =   255
         Min             =   0
         Max             =   100
         Pos             =   100
         OuterBevel      =   1
         InnerBevel      =   0
         OuterBevelWidth =   1
         InnerBevelWidth =   1
         BorderWidth     =   0
         ShadowColor     =   8421504
         HiliteColor     =   16777215
         Orientation     =   1
         TickMarks       =   30
         TickColor       =   0
         TickStyle       =   2
         CenterBar       =   4
      End
      Begin VB.Label lblLang 
         Caption         =   "Interface Language:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   30
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label lblGenVol 
         Alignment       =   1  'Right Justify
         Caption         =   "General Volume:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   29
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label lblDevice 
         Caption         =   "Output Device:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   27
         Top             =   2400
         Width           =   3615
      End
      Begin VB.Label lblAmx 
         Caption         =   "Auto Mix time (ms):"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   2400
         Width           =   3135
      End
      Begin VB.Label lblDispRef 
         Caption         =   "Display refresh time (ms):"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2880
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Label lblRemCol 
         Caption         =   "Remaining warning color:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1920
         Width           =   3135
      End
      Begin VB.Label lblRemWar 
         Caption         =   "Remaining warning time (s):"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1440
         Width           =   3135
      End
      Begin VB.Label lblDDformat 
         Caption         =   "Date display format:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label lblTDformat 
         Caption         =   "Time display format:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label lblProcPri 
         Caption         =   "Process Priority Thread:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   16
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label lblTasig 
         Caption         =   "Time announcing signal:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   15
         Top             =   720
         Width           =   3615
      End
   End
   Begin VB.CommandButton btOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4680
      TabIndex        =   13
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton btCancel 
      Caption         =   "Cancel"
      Height          =   1095
      Left            =   4680
      TabIndex        =   14
      Top             =   1320
      Width           =   1335
   End
End
Attribute VB_Name = "frmSett"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public Function outStringsS(LangSec As String)
'    LangFile = App.Path & "\language.ini"
'    If Dir$(LangFile) = vbNullString Then
'        Dim l As Integer
'        l = FreeFile
'        Open LangFile For Output As l
'        Close l
'    End If
'    LangSet.Location = LangFile
'    LangSet.SectionKey = LangSec
'    LangSet.Entry("sbtCancel") = btCancel.Caption
'    LangSet.Entry("sbtOK") = btOK.Caption
'    LangSet.Entry("sCkErrLog") = CkErrLog.Caption
'    LangSet.Entry("sfrmSett") = frmSett.Caption
'    LangSet.Entry("slblDispRef") = lblDispRef.Caption
'    LangSet.Entry("slblGenVol") = lblGenVol.Caption
'    LangSet.Entry("slblTasig") = lblTasig.Caption
'    LangSet.Entry("slblTDformat") = lblTDformat.Caption
'    LangSet.Entry("slblDDformat") = lblDDformat.Caption
'    LangSet.Entry("slblRemWar") = lblRemWar.Caption
'    LangSet.Entry("slblRemCol") = lblRemCol.Caption
'    LangSet.Entry("slblAmx") = lblAmx.Caption
'    LangSet.Entry("slblProcPri") = lblProcPri.Caption
'    LangSet.Entry("slblDevice") = lblDevice.Caption
'    LangSet.Entry("sTabSt0") = TabSt.TabCaption(0)
'    LangSet.Entry("sTabSt1") = TabSt.TabCaption(1)
'    LangSet.Entry("sTabSt2") = TabSt.TabCaption(2)
'
'End Function

Public Function inStringsS(LangSec As String)
        On Error GoTo Error_Routine

1       Dim lnBuf As String
2       LangSet.SectionKey = LangSec

3       lnBuf = LangSet.Entry("sbtCancel")
4       If Not LangSet.EntryNotFound Then frmSett.btCancel.Caption = lnBuf

5       lnBuf = LangSet.Entry("sbtOK")
6       If Not LangSet.EntryNotFound Then frmSett.btOK.Caption = lnBuf

7       lnBuf = LangSet.Entry("sCkErrLog")
8       If Not LangSet.EntryNotFound Then frmSett.CkErrLog.Caption = lnBuf

9       lnBuf = LangSet.Entry("sfrmSett")
10      If Not LangSet.EntryNotFound Then frmSett.Caption = lnBuf

11      lnBuf = LangSet.Entry("slblDispRef")
12      If Not LangSet.EntryNotFound Then frmSett.lblDispRef.Caption = lnBuf

13      lnBuf = LangSet.Entry("slblGenVol")
14      If Not LangSet.EntryNotFound Then frmSett.lblGenVol.Caption = lnBuf

15      lnBuf = LangSet.Entry("slblTasig")
16      If Not LangSet.EntryNotFound Then frmSett.lblTasig.Caption = lnBuf

17      lnBuf = LangSet.Entry("slblTDformat")
18      If Not LangSet.EntryNotFound Then frmSett.lblTDformat.Caption = lnBuf

19      lnBuf = LangSet.Entry("slblDDformat")
20      If Not LangSet.EntryNotFound Then frmSett.lblDDformat.Caption = lnBuf

21      lnBuf = LangSet.Entry("slblRemWar")
22      If Not LangSet.EntryNotFound Then frmSett.lblRemWar.Caption = lnBuf

23      lnBuf = LangSet.Entry("slblRemCol")
24      If Not LangSet.EntryNotFound Then frmSett.lblRemCol.Caption = lnBuf

25      lnBuf = LangSet.Entry("slblAmx")
26      If Not LangSet.EntryNotFound Then frmSett.lblAmx.Caption = lnBuf

27      lnBuf = LangSet.Entry("slblProcPri")
28      If Not LangSet.EntryNotFound Then frmSett.lblProcPri.Caption = lnBuf

29      lnBuf = LangSet.Entry("slblDevice")
30      If Not LangSet.EntryNotFound Then frmSett.lblDevice.Caption = lnBuf

31      lnBuf = LangSet.Entry("sTabSt0")
32      If Not LangSet.EntryNotFound Then frmSett.TabSt.TabCaption(0) = lnBuf

33      lnBuf = LangSet.Entry("sTabSt1")
34      If Not LangSet.EntryNotFound Then frmSett.TabSt.TabCaption(1) = lnBuf

35      lnBuf = LangSet.Entry("sTabSt2")
36      If Not LangSet.EntryNotFound Then frmSett.TabSt.TabCaption(2) = lnBuf

37      lnBuf = LangSet.Entry("slblLang")
38      If Not LangSet.EntryNotFound Then frmSett.lblLang.Caption = lnBuf

39      lnBuf = LangSet.Entry("ckOnTop")
40      If Not LangSet.EntryNotFound Then frmSett.ckOnTop.Caption = lnBuf

Exit_Routine:
    Exit Function
Error_Routine:
    Debug.Assert False
    err.Raise err.Number, "frmSett.inStringsS(" & Erl & "):" & err.Source, err.Description
End Function

Private Sub btCancel_Click()
        On Error GoTo Error_Routine

1       Unload Me

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmSett.btCancel_Click"
    Resume Exit_Routine
End Sub

Private Sub btOK_Click()

        On Error GoTo Error_Routine

1       With TxtAmix
2           If .Text > 3000 Then .Text = 3000
3           If .Text < 100 Then .Text = 100
4       End With
5       With TxtRefr
6           If .Text > 1000 Then .Text = 1000
7           If .Text < 50 Then .Text = 50
8       End With
9       With TxtWarTim
10          If .Text > 59 Then .Text = 59
11          If .Text < 0 Then .Text = 0
12      End With

13      DateForm = CmbDate.Text
14      TimeForm = CmbTime.Text
15      frmMain.lblTime.Caption = Format(time, TimeForm)
16      frmMain.lblDatex.Caption = Format(Date, DateForm)
17      RemWarn = TxtWarTim.Text
18      errCk = CkErrLog.value
19      Cfrq = TxtRefr.Text
20      aMixtime = TxtAmix.Text
21      frmMain.Timer4.Interval = aMixtime / 100
22      device = CmbDev.ListIndex + 1

        '    For i = 0 To 2
        '        frmMain.'MPlayer(i).ControlFreq Cfrq
        '    Next i

23      If device <> devicePrev Then
24          MsgBox LangSet.Entry("sMsdev", , Language), 64, "Jingle Palette"
25          BASS_Free   'stop everything and free all player resources
26          InitDev     'start digital output
27          OpenFile frmMain.lstPal.Text 'reload palette
28          devicePrev = device
29      End If

30      IniSet.Entry("Device") = (CmbDev.ListIndex + 1)
31      IniSet.Entry("Refresh") = TxtRefr.Text
32      IniSet.Entry("RemainWarn") = TxtWarTim.Text
33      IniSet.Entry("AutoMixTime") = TxtAmix.Text
34      If CmbTime.Text <> "" Then IniSet.Entry("TimeFormat") = CmbTime.Text
35      If CmbDate.Text <> "" Then IniSet.Entry("DateFormat") = CmbDate.Text
36      IniSet.Entry("RemainColor") = CStr(RemColor)
37      IniSet.Entry("ErrorLog") = CkErrLog.value

38      Language = cmbLang.Text
39      IniSet.Entry("Language") = Language

40      Select Case cmbThread.ListIndex
            Case 0
41              Thrd = -15
42          Case 1
43              Thrd = -2
44          Case 2
45              Thrd = -1
46          Case 3
47              Thrd = 0
48          Case 4
49              Thrd = 1
50          Case 5
51              Thrd = 2
52          Case 6
53              Thrd = 15
54      End Select
55      If Thrd = 15 Then
56          If MsgBox(LangSet.Entry("sMspr1", , Language) & vbCr & LangSet.Entry("sMspr2", , Language) & vbCr & LangSet.Entry("sMspr3", , Language), 292, LangSet.Entry("sMspr4", , Language)) = vbYes Then
57              ThreadPriority = Thrd
58          Else
59              Thrd = ThreadPriority
60          End If
61      Else
62          ThreadPriority = Thrd
63      End If


        If ckOnTop.value = vbChecked Then
            OnTop = True
        Else
            OnTop = False
        End If
        
        IniSet.Entry("AlwaysOnTop") = OnTop
        Call SetTopMostWindow(frmMain.hwnd, OnTop)

64      Unload Me

Exit_Routine:
    Exit Sub
    Unload Me
Error_Routine:
    Debug.Assert False
    ErrorLog "frmSett.btOK_Click"
    Resume Exit_Routine
End Sub
Private Sub CkErrLog_Click()
    '    If CkErrLog.value = vbUnchecked Then
    '        CkErrLog.BackColor = vbButtonFace       '&H8000000F&
    '    Else
    '        CkErrLog.BackColor = &HFFFF&
    '        'CkErrLog.value = vbUnchecked
    '    End If
    On Error GoTo Error_Routine

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmSett.CkErrLog_Click"
    Resume Exit_Routine
End Sub

Private Sub cmbLang_Click()
        On Error GoTo Error_Routine

1       If cmbLang.Visible Then
            'MsgBox "Changed to " & cmbLang.Text
2           Call inStringsS(cmbLang.Text)
3           Call frmMain.inStrings(cmbLang.Text)
            Language = cmbLang.Text
4       End If

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmSett.cmbLang_Click"
    Resume Exit_Routine
End Sub

Private Sub ColSel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

        On Error GoTo Error_Routine

1       If Button = vbRightButton Then
2           ColSel.BackColor = &HC0& 'reset to default color
3       Else
            On Error GoTo errc
4           frmMain.cmDlg.flags = &H2 Or &H1
5           frmMain.cmDlg.Color = RemColor
6           frmMain.cmDlg.ShowColor
7           ColSel.BackColor = frmMain.cmDlg.Color
8       End If
9       RemColor = ColSel.BackColor

errc:
    'cancel pressed on the color selection box
Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmSett.ColSel_MouseDown"
    Resume Exit_Routine
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        On Error GoTo Error_Routine

1       Select Case KeyCode
            Case vbKeyReturn
2               btOK_Click
3           Case vbKeyEscape
4               btCancel_Click
5       End Select

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmSett.Form_KeyDown"
    Resume Exit_Routine
End Sub

Private Sub Form_Load()

        On Error GoTo Error_Routine
        Call SetTopMostWindow(frmSett.hwnd, OnTop)

1       TabSt.Tab = 0
2       TxtRefr.Text = Cfrq
3       SelectListItemByString CmbDate, DateForm
4       SelectListItemByString CmbTime, TimeForm
5       TxtWarTim.Text = RemWarn
6       ColSel.BackColor = RemColor
7       CkErrLog.value = errCk
8       TxtAmix.Text = aMixtime
9       If Len(tmSound) > 0 Then btSound.Caption = btCaption(tmSound, btSound, 20)

10      Dim c As Integer
11      c = 1      'device 1 = 1st real device
12      While BASS_GetDeviceDescription(c)
13          CmbDev.AddItem VBStrFromAnsiPtr(BASS_GetDeviceDescription(c))
14          c = c + 1
15      Wend
16      If (CmbDev.ListCount) Then CmbDev.ListIndex = 0
        'device = CmbDev.ListIndex + 1
17      CmbDev.ListIndex = device - 1

18      txtVol.MouseIcon = LoadResPicture(101, vbResCursor)
19      SlideVol.ThumbBitmap = LoadResPicture(111, vbResBitmap)
20      SlideVol.pos = vol
21      txtVol.Text = SlideVol.pos - 100 & " dB"
22      Select Case Thrd
            Case -15
23              cmbThread.ListIndex = 0
24          Case -2
25              cmbThread.ListIndex = 1
26          Case -1
27              cmbThread.ListIndex = 2
28          Case 0
29              cmbThread.ListIndex = 3
30          Case 1
31              cmbThread.ListIndex = 4
32          Case 2
33              cmbThread.ListIndex = 5
34          Case 15
35              cmbThread.ListIndex = 6
36      End Select

37      Dim il As Long
38      LangSet.RegistrySubKeys.Update
39      cmbLang.Clear
40      For il = 1 To LangSet.RegistrySubKeys.Count
41          cmbLang.AddItem LangSet.RegistrySubKeys.Key(il)
42      Next il
43      cmbLang.Text = Language

44      If Language <> "English" Then
45          Call inStringsS(Language)
46      End If


        If OnTop Then
            ckOnTop.value = vbChecked
        Else
            ckOnTop.value = vbUnchecked
        End If


Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmSett.Form_Load"
    Resume Exit_Routine
End Sub

Private Sub btSound_Click()
        On Error GoTo Error_Routine

1       On Error GoTo errHandler ' CancelError is True
2       Dim initDirPrev As String
3       frmMain.cmDlg.flags = cdlOFNHideReadOnly + cdlOFNFileMustExist
4       frmMain.cmDlg.Filter = LangSet.Entry("mDlgType", , Language) & " (*.wav;*.mp3;*.mp2;*.mp1;*.mpa;*.ogg)|*.wav;*.mp3;*.mp2;*.mp1;*.mpa;*.ogg"
5       frmMain.cmDlg.DialogTitle = LangSet.Entry("mDlgAsTa", , Language)
6       initDirPrev = frmMain.cmDlg.InitDir
7       frmMain.cmDlg.InitDir = App.Path
8       frmMain.cmDlg.ShowOpen  'open dialog to select a file
9       tmSound = frmMain.cmDlg.FileName
10      btSound.Caption = btCaption(tmSound, btSound, 20)
errHandler:
12      frmMain.cmDlg.InitDir = initDirPrev
13      Exit Sub 'asignment over

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmSett.btSound_Click"
    Resume Exit_Routine
End Sub
Private Sub txtVol_Click()
        On Error GoTo Error_Routine

1       SlideVol.pos = 100
2       vol = 100
3       txtVol.Text = vol - 100 & " dB"
4       Call BASS_SetConfig(BASS_CONFIG_GVOL_STREAM, vol)

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmSett.txtVol_Click"
    Resume Exit_Routine
End Sub

Private Sub SlideVol_Scroll(ByVal value As Integer)

        On Error GoTo Error_Routine

1       vol = value
2       txtVol.Text = vol - 100 & " dB"
3       Call BASS_SetConfig(BASS_CONFIG_GVOL_STREAM, vol)

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmSett.SlideVol_Scroll"
    Resume Exit_Routine
End Sub

Private Sub Timer_Timer()

        On Error GoTo Error_Routine

1       If MouseIsOverObject(txtVol) Then
2           txtVol.ForeColor = &HFF&
3       Else
4           txtVol.ForeColor = vbButtonText
5       End If

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmSett.Timer_Timer"
    Resume Exit_Routine
End Sub

Private Sub TxtAmix_GotFocus()
        On Error GoTo Error_Routine

1       SelectAll TxtAmix

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmSett.TxtAmix_GotFocus"
    Resume Exit_Routine
End Sub

Private Sub TxtAmix_KeyDown(KeyCode As Integer, Shift As Integer)
        On Error Resume Next
1       With TxtAmix
2           If KeyCode = vbKeyUp Then .Text = .Text + 10
3           If KeyCode = vbKeyDown Then .Text = .Text - 10
4           If KeyCode = vbKeyRight Then .Text = .Text + 100
5           If KeyCode = vbKeyLeft Then .Text = .Text - 100
6       End With
7       KeyPressFilterDelKey TxtAmix, KeyCode, iskpNumbers
End Sub

Private Sub TxtAmix_KeyPress(KeyAscii As Integer)
        On Error GoTo Error_Routine

1       KeyAscii = KeyPressFilter(TxtAmix, KeyAscii, iskpNumbers)

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmSett.TxtAmix_KeyPress"
    Resume Exit_Routine
End Sub

Private Sub TxtAmix_KeyUp(KeyCode As Integer, Shift As Integer)
        On Error GoTo Error_Routine

1       KeyPressFilterDelKey TxtAmix, KeyCode, iskpNumbers

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmSett.TxtAmix_KeyUp"
    Resume Exit_Routine
End Sub

Private Sub TxtRefr_GotFocus()
        On Error GoTo Error_Routine

1       SelectAll TxtRefr

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmSett.TxtRefr_GotFocus"
    Resume Exit_Routine
End Sub

Private Sub TxtRefr_KeyDown(KeyCode As Integer, Shift As Integer)
        On Error Resume Next
1       With TxtRefr
2           If KeyCode = vbKeyUp Then .Text = .Text + 1
3           If KeyCode = vbKeyDown Then .Text = .Text - 1
4           If KeyCode = vbKeyRight Then .Text = .Text + 100
5           If KeyCode = vbKeyLeft Then .Text = .Text - 100
6       End With
7       KeyPressFilterDelKey TxtRefr, KeyCode, iskpNumbers
End Sub

Private Sub TxtRefr_KeyPress(KeyAscii As Integer)
        On Error GoTo Error_Routine

1       KeyAscii = KeyPressFilter(TxtRefr, KeyAscii, iskpNumbers)

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmSett.TxtRefr_KeyPress"
    Resume Exit_Routine
End Sub

Private Sub TxtRefr_KeyUp(KeyCode As Integer, Shift As Integer)
        On Error GoTo Error_Routine

1       KeyPressFilterDelKey TxtRefr, KeyCode, iskpNumbers

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmSett.TxtRefr_KeyUp"
    Resume Exit_Routine
End Sub

Private Sub TxtWarTim_GotFocus()
        On Error GoTo Error_Routine

1       SelectAll TxtWarTim

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmSett.TxtWarTim_GotFocus"
    Resume Exit_Routine
End Sub

Private Sub TxtWarTim_KeyDown(KeyCode As Integer, Shift As Integer)
        On Error Resume Next
1       With TxtWarTim
2           If KeyCode = vbKeyUp Then .Text = .Text + 1
3           If KeyCode = vbKeyDown Then .Text = .Text - 1
4           If KeyCode = vbKeyRight Then .Text = .Text + 10
5           If KeyCode = vbKeyLeft Then .Text = .Text - 10
6       End With
7       KeyPressFilterDelKey TxtWarTim, KeyCode, iskpNumbers
End Sub

Private Sub TxtWarTim_KeyPress(KeyAscii As Integer)
        On Error GoTo Error_Routine

1       KeyAscii = KeyPressFilter(TxtWarTim, KeyAscii, iskpNumbers)

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmSett.TxtWarTim_KeyPress"
    Resume Exit_Routine
End Sub

Private Sub TxtWarTim_KeyUp(KeyCode As Integer, Shift As Integer)
        On Error GoTo Error_Routine

1       KeyPressFilterDelKey TxtWarTim, KeyCode, iskpNumbers

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmSett.TxtWarTim_KeyUp"
    Resume Exit_Routine
End Sub
