VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F9B85A83-4DBD-11CF-94E5-0000C0571740}#1.0#0"; "NSLIDE32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3E2E5A60-5763-11D5-A29B-F938E4C62136}#10.3#0"; "LevelM.ocx"
Begin VB.Form frmMain 
   Caption         =   "Jingle Palette"
   ClientHeight    =   8055
   ClientLeft      =   255
   ClientTop       =   645
   ClientWidth     =   11880
   FillColor       =   &H00808080&
   Icon            =   "frmMainWindow.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8055
   ScaleWidth      =   11880
   Begin TabDlg.SSTab TabCt 
      Height          =   2535
      Left            =   120
      TabIndex        =   45
      Top             =   5400
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   4471
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   794
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Palettes"
      TabPicture(0)   =   "frmMainWindow.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ckAssign"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "btSave"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtSave"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "btNew"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lstPal"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Time Announce"
      TabPicture(1)   =   "frmMainWindow.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ln1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblDelay"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblSecond"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblMinute"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "b_tmAnJin"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "ckTmAnn"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "DTMinute"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "DTSecond"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "DTDelay"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Play Stream"
      TabPicture(2)   =   "frmMainWindow.frx":107A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblRemLoc"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "txtConDisp"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "ckMixByTa"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "ckPlayStream"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cmbStream"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Settings"
      TabPicture(3)   =   "frmMainWindow.frx":1096
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblVol"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "lblVole"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "lblVolexp"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "SlideVolCh"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "ckAssVol"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "btHelp"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "btSettings"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "btAbout"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "ckLoopI"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).ControlCount=   9
      Begin MSComCtl2.DTPicker DTDelay 
         Height          =   375
         Left            =   -74040
         TabIndex        =   112
         Top             =   1320
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "ss"
         Format          =   66060291
         UpDown          =   -1  'True
         CurrentDate     =   38018
      End
      Begin MSComCtl2.DTPicker DTSecond 
         Height          =   375
         Left            =   -74040
         TabIndex        =   111
         Top             =   600
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "ss"
         Format          =   66060291
         UpDown          =   -1  'True
         CurrentDate     =   38018
      End
      Begin MSComCtl2.DTPicker DTMinute 
         Height          =   375
         Left            =   -74040
         TabIndex        =   110
         Top             =   120
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "mm"
         Format          =   66060291
         UpDown          =   -1  'True
         CurrentDate     =   38018
      End
      Begin VB.CheckBox ckLoopI 
         Caption         =   "Assign/Clear &Loop Mode"
         Height          =   975
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton btAbout 
         Caption         =   "A&bout..."
         Height          =   975
         Left            =   -70920
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   960
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton btSettings 
         Caption         =   "Se&ttings..."
         Height          =   975
         Left            =   -73560
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   960
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton btHelp 
         Caption         =   "&Help..."
         Height          =   975
         Left            =   -72240
         MaskColor       =   &H00FF00FF&
         Style           =   1  'Graphical
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   960
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CheckBox ckAssVol 
         Caption         =   "Assign new Volume level to Jingle"
         Height          =   735
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   120
         Width           =   5175
      End
      Begin VB.ComboBox cmbStream 
         Height          =   315
         Left            =   -74880
         TabIndex        =   68
         TabStop         =   0   'False
         Text            =   "http://"
         ToolTipText     =   "Enter the location. Eg. http://yourserver.com/yourfile.mp3"
         Top             =   360
         Width           =   5175
      End
      Begin VB.CheckBox ckPlayStream 
         Caption         =   "Play Remote Location"
         Height          =   735
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   840
         Width           =   2055
      End
      Begin VB.CheckBox ckMixByTa 
         Caption         =   "Auto Mix by Time Announce"
         Height          =   735
         Left            =   -71295
         Style           =   1  'Graphical
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtConDisp 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFC0&
         Height          =   300
         Left            =   -74880
         Locked          =   -1  'True
         TabIndex        =   65
         TabStop         =   0   'False
         Text            =   "Not connected"
         Top             =   1680
         Width           =   5175
      End
      Begin VB.ListBox lstPal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1860
         ItemData        =   "frmMainWindow.frx":10B2
         Left            =   3000
         List            =   "frmMainWindow.frx":10B4
         Sorted          =   -1  'True
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   120
         Width           =   2295
      End
      Begin VB.CheckBox ckTmAnn 
         Caption         =   "Time A&nnounce"
         Height          =   855
         Left            =   -73200
         Style           =   1  'Graphical
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   120
         Width           =   3510
      End
      Begin VB.CheckBox b_tmAnJin 
         Caption         =   "Select jingle to be announ&ced..."
         Enabled         =   0   'False
         Height          =   735
         Left            =   -73200
         Style           =   1  'Graphical
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   1200
         Width           =   3510
      End
      Begin VB.CommandButton btNew 
         Caption         =   "Stop and &Empty Palette"
         Height          =   550
         Left            =   120
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   1415
         Width           =   2775
      End
      Begin VB.TextBox txtSave 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   165
         TabIndex        =   47
         TabStop         =   0   'False
         ToolTipText     =   "Enter a name for the palette and press 'Enter'. To cancel, press 'Esc'"
         Top             =   815
         Visible         =   0   'False
         Width           =   2670
      End
      Begin VB.CommandButton btSave 
         Caption         =   "&Save Current Palette"
         Height          =   550
         Left            =   120
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   755
         Width           =   2775
      End
      Begin VB.CheckBox ckAssign 
         Caption         =   "&Assign Jingle To Button"
         Height          =   550
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   95
         Width           =   2775
      End
      Begin SliderLib.Slider SlideVolCh 
         Height          =   375
         Left            =   -74880
         TabIndex        =   74
         TabStop         =   0   'False
         ToolTipText     =   "General Volume"
         Top             =   120
         Width           =   5175
         _Version        =   65536
         _ExtentX        =   9128
         _ExtentY        =   661
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
         Orientation     =   0
         TickMarks       =   60
         TickColor       =   0
         TickStyle       =   2
         CenterBar       =   4
      End
      Begin VB.Label lblVolexp 
         Caption         =   "Jingle"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -73155
         TabIndex        =   77
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label lblVole 
         Alignment       =   1  'Right Justify
         Caption         =   "The volume level for"
         Height          =   195
         Left            =   -74880
         TabIndex        =   76
         Top             =   600
         Width           =   1670
      End
      Begin VB.Label lblVol 
         Alignment       =   1  'Right Justify
         Caption         =   "0 dB"
         Height          =   255
         Left            =   -70320
         TabIndex        =   75
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblRemLoc 
         Caption         =   "Remote location URL:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   69
         Top             =   120
         Width           =   5175
      End
      Begin VB.Label lblMinute 
         Alignment       =   1  'Right Justify
         Caption         =   "Minute:"
         Height          =   375
         Left            =   -74880
         TabIndex        =   58
         Top             =   120
         Width           =   735
      End
      Begin VB.Label lblSecond 
         Alignment       =   1  'Right Justify
         Caption         =   "Second:"
         Height          =   375
         Left            =   -74880
         TabIndex        =   57
         Top             =   600
         Width           =   735
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblDelay 
         Alignment       =   1  'Right Justify
         Caption         =   "Jingle delay (s):"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -74925
         TabIndex        =   56
         Top             =   1320
         Width           =   765
      End
      Begin VB.Line ln1 
         X1              =   -74880
         X2              =   -69720
         Y1              =   1080
         Y2              =   1080
      End
   End
   Begin VB.Frame fraVu 
      Height          =   5295
      Left            =   11160
      TabIndex        =   42
      Top             =   0
      Width           =   615
      Begin levelm.LevelMeter VuL 
         Height          =   4845
         Left            =   60
         Top             =   165
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   8546
         Level           =   100
         MidLevel        =   60
         Horizontal      =   0   'False
         Reverse         =   0   'False
         PeakDelay       =   500
         Gradient        =   -1  'True
         Solid           =   0   'False
      End
      Begin levelm.LevelMeter VuR 
         Height          =   4845
         Left            =   305
         Top             =   160
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   8546
         Level           =   100
         MidLevel        =   60
         Horizontal      =   0   'False
         Reverse         =   0   'False
         PeakDelay       =   500
         Gradient        =   -1  'True
         Solid           =   0   'False
      End
      Begin VB.Label lblR 
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   44
         Top             =   5040
         Width           =   135
      End
      Begin VB.Label lblL 
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   43
         Top             =   5040
         Width           =   135
      End
   End
   Begin VB.Frame fraJingles 
      ClipControls    =   0   'False
      Height          =   5295
      Left            =   120
      TabIndex        =   30
      Top             =   0
      Width           =   10935
      Begin VB.PictureBox p_Jingle 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   0
         Left            =   180
         ScaleHeight     =   75
         ScaleWidth      =   225
         TabIndex        =   109
         Top             =   300
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox p_Jingle 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   1
         Left            =   2340
         ScaleHeight     =   75
         ScaleWidth      =   225
         TabIndex        =   108
         Top             =   300
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox p_Jingle 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   2
         Left            =   4500
         ScaleHeight     =   75
         ScaleWidth      =   225
         TabIndex        =   107
         Top             =   300
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox p_Jingle 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   3
         Left            =   6660
         ScaleHeight     =   75
         ScaleWidth      =   225
         TabIndex        =   106
         Top             =   300
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox p_Jingle 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   4
         Left            =   8820
         ScaleHeight     =   75
         ScaleWidth      =   225
         TabIndex        =   105
         Top             =   300
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox p_Jingle 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   5
         Left            =   180
         ScaleHeight     =   75
         ScaleWidth      =   225
         TabIndex        =   104
         Top             =   1140
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox p_Jingle 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   6
         Left            =   2340
         ScaleHeight     =   75
         ScaleWidth      =   225
         TabIndex        =   103
         Top             =   1140
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox p_Jingle 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   7
         Left            =   4500
         ScaleHeight     =   75
         ScaleWidth      =   225
         TabIndex        =   102
         Top             =   1140
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox p_Jingle 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   8
         Left            =   6660
         ScaleHeight     =   75
         ScaleWidth      =   225
         TabIndex        =   101
         Top             =   1140
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox p_Jingle 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   9
         Left            =   8820
         ScaleHeight     =   75
         ScaleWidth      =   225
         TabIndex        =   100
         Top             =   1140
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox p_Jingle 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   10
         Left            =   180
         ScaleHeight     =   75
         ScaleWidth      =   225
         TabIndex        =   99
         Top             =   1980
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox p_Jingle 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   11
         Left            =   2340
         ScaleHeight     =   75
         ScaleWidth      =   225
         TabIndex        =   98
         Top             =   1980
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox p_Jingle 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   12
         Left            =   4500
         ScaleHeight     =   75
         ScaleWidth      =   225
         TabIndex        =   97
         Top             =   1980
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox p_Jingle 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   13
         Left            =   6660
         ScaleHeight     =   75
         ScaleWidth      =   225
         TabIndex        =   96
         Top             =   1980
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox p_Jingle 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   14
         Left            =   8820
         ScaleHeight     =   75
         ScaleWidth      =   225
         TabIndex        =   95
         Top             =   1980
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox p_Jingle 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   15
         Left            =   180
         ScaleHeight     =   75
         ScaleWidth      =   225
         TabIndex        =   94
         Top             =   2820
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox p_Jingle 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   16
         Left            =   2340
         ScaleHeight     =   75
         ScaleWidth      =   225
         TabIndex        =   93
         Top             =   2820
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox p_Jingle 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   17
         Left            =   4500
         ScaleHeight     =   75
         ScaleWidth      =   225
         TabIndex        =   92
         Top             =   2820
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox p_Jingle 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   18
         Left            =   6660
         ScaleHeight     =   75
         ScaleWidth      =   225
         TabIndex        =   91
         Top             =   2820
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox p_Jingle 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   19
         Left            =   8820
         ScaleHeight     =   75
         ScaleWidth      =   225
         TabIndex        =   90
         Top             =   2820
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox p_Jingle 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   20
         Left            =   180
         ScaleHeight     =   75
         ScaleWidth      =   225
         TabIndex        =   89
         Top             =   3660
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox p_Jingle 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   21
         Left            =   2340
         ScaleHeight     =   75
         ScaleWidth      =   225
         TabIndex        =   88
         Top             =   3660
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox p_Jingle 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   22
         Left            =   4500
         ScaleHeight     =   75
         ScaleWidth      =   225
         TabIndex        =   87
         Top             =   3660
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox p_Jingle 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   23
         Left            =   6660
         ScaleHeight     =   75
         ScaleWidth      =   225
         TabIndex        =   86
         Top             =   3660
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox p_Jingle 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   24
         Left            =   8820
         ScaleHeight     =   75
         ScaleWidth      =   225
         TabIndex        =   85
         Top             =   3660
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox p_Jingle 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   25
         Left            =   180
         ScaleHeight     =   75
         ScaleWidth      =   225
         TabIndex        =   84
         Top             =   4500
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox p_Jingle 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   26
         Left            =   2340
         ScaleHeight     =   75
         ScaleWidth      =   225
         TabIndex        =   83
         Top             =   4500
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox p_Jingle 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   27
         Left            =   4500
         ScaleHeight     =   75
         ScaleWidth      =   225
         TabIndex        =   82
         Top             =   4500
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox p_Jingle 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   28
         Left            =   6660
         ScaleHeight     =   75
         ScaleWidth      =   225
         TabIndex        =   81
         Top             =   4500
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox p_Jingle 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   29
         Left            =   8820
         ScaleHeight     =   75
         ScaleWidth      =   225
         TabIndex        =   80
         Top             =   4500
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox b_Jingle 
         Height          =   735
         Index           =   0
         Left            =   120
         MaskColor       =   &H00000000&
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   2055
      End
      Begin VB.CheckBox b_Jingle 
         Height          =   735
         Index           =   29
         Left            =   8760
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   4440
         Width           =   2055
      End
      Begin VB.CheckBox b_Jingle 
         Height          =   735
         Index           =   28
         Left            =   6600
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   4440
         Width           =   2055
      End
      Begin VB.CheckBox b_Jingle 
         Height          =   735
         Index           =   27
         Left            =   4440
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   4440
         Width           =   2055
      End
      Begin VB.CheckBox b_Jingle 
         Height          =   735
         Index           =   26
         Left            =   2280
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   4440
         Width           =   2055
      End
      Begin VB.CheckBox b_Jingle 
         Height          =   735
         Index           =   25
         Left            =   120
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   4440
         Width           =   2055
      End
      Begin VB.CheckBox b_Jingle 
         Height          =   735
         Index           =   24
         Left            =   8760
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   3600
         Width           =   2055
      End
      Begin VB.CheckBox b_Jingle 
         Height          =   735
         Index           =   23
         Left            =   6600
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   3600
         Width           =   2055
      End
      Begin VB.CheckBox b_Jingle 
         Height          =   735
         Index           =   22
         Left            =   4440
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   3600
         Width           =   2055
      End
      Begin VB.CheckBox b_Jingle 
         Height          =   735
         Index           =   21
         Left            =   2280
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3600
         Width           =   2055
      End
      Begin VB.CheckBox b_Jingle 
         Height          =   735
         Index           =   20
         Left            =   120
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   3600
         Width           =   2055
      End
      Begin VB.CheckBox b_Jingle 
         Height          =   735
         Index           =   19
         Left            =   8760
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2760
         Width           =   2055
      End
      Begin VB.CheckBox b_Jingle 
         Height          =   735
         Index           =   18
         Left            =   6600
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2760
         Width           =   2055
      End
      Begin VB.CheckBox b_Jingle 
         Height          =   735
         Index           =   17
         Left            =   4440
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2760
         Width           =   2055
      End
      Begin VB.CheckBox b_Jingle 
         Height          =   735
         Index           =   16
         Left            =   2280
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2760
         Width           =   2055
      End
      Begin VB.CheckBox b_Jingle 
         Height          =   735
         Index           =   15
         Left            =   120
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2760
         Width           =   2055
      End
      Begin VB.CheckBox b_Jingle 
         Height          =   735
         Index           =   14
         Left            =   8760
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1920
         Width           =   2055
      End
      Begin VB.CheckBox b_Jingle 
         Height          =   735
         Index           =   13
         Left            =   6600
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1920
         Width           =   2055
      End
      Begin VB.CheckBox b_Jingle 
         Height          =   735
         Index           =   12
         Left            =   4440
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1920
         Width           =   2055
      End
      Begin VB.CheckBox b_Jingle 
         Height          =   735
         Index           =   11
         Left            =   2280
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1920
         Width           =   2055
      End
      Begin VB.CheckBox b_Jingle 
         Height          =   735
         Index           =   10
         Left            =   120
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1920
         Width           =   2055
      End
      Begin VB.CheckBox b_Jingle 
         Height          =   735
         Index           =   9
         Left            =   8760
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CheckBox b_Jingle 
         Height          =   735
         Index           =   8
         Left            =   6600
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CheckBox b_Jingle 
         Height          =   735
         Index           =   7
         Left            =   4440
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CheckBox b_Jingle 
         Height          =   735
         Index           =   6
         Left            =   2280
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CheckBox b_Jingle 
         Height          =   735
         Index           =   5
         Left            =   120
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CheckBox b_Jingle 
         Height          =   735
         Index           =   4
         Left            =   8760
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   2055
      End
      Begin VB.CheckBox b_Jingle 
         Height          =   735
         Index           =   3
         Left            =   6600
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   2055
      End
      Begin VB.CheckBox b_Jingle 
         Height          =   735
         Index           =   2
         Left            =   4440
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   2055
      End
      Begin VB.CheckBox b_Jingle 
         Height          =   735
         Index           =   1
         Left            =   2280
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin MSComDlg.CommonDialog cmDlg 
         Left            =   2040
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   2640
         Top             =   0
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   3120
         Top             =   0
      End
      Begin VB.Timer Timer3 
         Interval        =   40
         Left            =   3600
         Top             =   0
      End
      Begin VB.Timer Timer4 
         Enabled         =   0   'False
         Interval        =   20
         Left            =   4080
         Top             =   0
      End
      Begin VB.Timer Timer5 
         Interval        =   10
         Left            =   5280
         Top             =   0
      End
   End
   Begin VB.Frame fraPanel 
      Height          =   2655
      Left            =   5640
      TabIndex        =   31
      Top             =   5280
      Width           =   6135
      Begin VB.CommandButton btExit 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   4920
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   240
         Width           =   1110
      End
      Begin VB.CommandButton btDn 
         Caption         =   "Next Palette"
         Height          =   1095
         Left            =   120
         TabIndex        =   62
         TabStop         =   0   'False
         ToolTipText     =   "[Page Down]"
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton btUp 
         Caption         =   "Previous Palette"
         Height          =   1095
         Left            =   120
         TabIndex        =   61
         TabStop         =   0   'False
         ToolTipText     =   "[Page Up]"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox ckTouch 
         Caption         =   "Touc&h Play"
         Height          =   1095
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CheckBox ckAutoRep 
         Caption         =   "Auto&repeat"
         Height          =   1095
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CheckBox ckAutoMix 
         Caption         =   "Auto &Mix"
         Height          =   1095
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1095
      End
      Begin VB.PictureBox picDispBack 
         BackColor       =   &H00000000&
         Height          =   1095
         Left            =   1320
         ScaleHeight     =   1035
         ScaleWidth      =   3435
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   240
         Width           =   3495
         Begin VB.Label lblErr 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "Wingdings"
               Size            =   20.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H008080FF&
            Height          =   435
            Left            =   1510
            TabIndex        =   40
            ToolTipText     =   "Click me!"
            Top             =   10
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.Label lblWeek 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0FF&
            Height          =   255
            Left            =   1440
            TabIndex        =   41
            ToolTipText     =   "Current week number"
            Top             =   435
            Width           =   270
         End
         Begin VB.Label lblDatex 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0FF&
            Height          =   375
            Left            =   960
            TabIndex        =   36
            ToolTipText     =   "Current date"
            Top             =   720
            Width           =   2415
         End
         Begin VB.Shape ShRem 
            BorderColor     =   &H000000FF&
            BorderWidth     =   2
            Height          =   375
            Left            =   270
            Top             =   40
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.Label lblPalNext 
            BackColor       =   &H00004080&
            BackStyle       =   0  'Transparent
            Caption         =   "Next"
            ForeColor       =   &H00FFFFC0&
            Height          =   255
            Left            =   0
            TabIndex        =   64
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblPalPr 
            BackColor       =   &H00000080&
            BackStyle       =   0  'Transparent
            Caption         =   "Previous"
            ForeColor       =   &H00FFFFC0&
            Height          =   195
            Left            =   0
            TabIndex        =   63
            Top             =   405
            Width           =   1455
         End
         Begin VB.Label lblDate 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            Caption         =   "day"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0FF&
            Height          =   285
            Left            =   1680
            TabIndex        =   35
            ToolTipText     =   "Current day"
            Top             =   440
            Width           =   1695
         End
         Begin VB.Label lblTa 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "½"
            BeginProperty Font 
               Name            =   "Wingdings"
               Size            =   20.25
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   435
            Left            =   1160
            TabIndex        =   59
            ToolTipText     =   "Click me!"
            Top             =   0
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.Label txtDisPal 
            BackColor       =   &H00008000&
            BackStyle       =   0  'Transparent
            Caption         =   "Current"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFC0&
            Height          =   200
            Left            =   0
            TabIndex        =   50
            Top             =   620
            Width           =   1455
         End
         Begin VB.Label lblTime 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "00:00:00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0FF&
            Height          =   315
            Left            =   1680
            TabIndex        =   34
            ToolTipText     =   "Current time"
            Top             =   50
            Width           =   1695
         End
         Begin VB.Label lbl_gr 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "_"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   17.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFC0&
            Height          =   430
            Left            =   15
            TabIndex        =   37
            ToolTipText     =   "Time remaining of the last jingle played"
            Top             =   -150
            Width           =   280
         End
         Begin VB.Label lblPosMis 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   17.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFC0&
            Height          =   435
            Left            =   675
            TabIndex        =   39
            ToolTipText     =   "Time remaining of the last jingle played"
            Top             =   0
            Width           =   470
         End
         Begin VB.Label lblPosSec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   17.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFC0&
            Height          =   435
            Left            =   310
            TabIndex        =   38
            ToolTipText     =   "Time remaining of the last jingle played"
            Top             =   0
            Width           =   375
         End
      End
      Begin VB.CommandButton btStop 
         Caption         =   "STOP ALL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   4920
         MaskColor       =   &H8000000C&
         TabIndex        =   32
         TabStop         =   0   'False
         ToolTipText     =   "[Esc]"
         Top             =   1440
         Width           =   1095
      End
   End
   Begin VB.Menu mnujp 
      Caption         =   "Button"
      Visible         =   0   'False
      Begin VB.Menu mnuPlay 
         Caption         =   "Play"
      End
      Begin VB.Menu mnuPause 
         Caption         =   "Pause"
      End
      Begin VB.Menu mnuspc1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVol 
         Caption         =   "Assign Volume Level"
      End
      Begin VB.Menu mnuLoop 
         Caption         =   "Set/Clear Loop Mode"
      End
      Begin VB.Menu mnuTa 
         Caption         =   "Assign to Time Announce"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuspc2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAss 
         Caption         =   "Assign Jingle"
      End
      Begin VB.Menu mnuClr 
         Caption         =   "Clear Button"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SC_CLOSE As Long = &HF060&
Private Const MF_BYCOMMAND = &H0&
Private Declare Function DeleteMenu Lib "user32.dll" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32.dll" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Function inStrings(LangSec As String)
        On Error GoTo Error_Routine

1       Dim testlan As String, lnBuf As String
2       LangSet.SectionKey = LangSec
3       testlan = LangSet.Entry("mb_tmAnJin")
4       If LangSet.EntryNotFound Or testlan = "" Then
5           LangSet.SectionKey = "English"
6           Language = "English"
7       End If

8       lnBuf = LangSet.Entry("mb_tmAnJin")
9       If Not LangSet.EntryNotFound Then frmMain.b_tmAnJin.Caption = lnBuf

10      lnBuf = LangSet.Entry("mbtAbout")
11      If Not LangSet.EntryNotFound Then frmMain.btAbout.Caption = lnBuf

12      lnBuf = LangSet.Entry("mbtDn")
13      If Not LangSet.EntryNotFound Then frmMain.btDn.Caption = lnBuf

14      lnBuf = LangSet.Entry("mbtExit")
15      If Not LangSet.EntryNotFound Then frmMain.btExit.Caption = lnBuf

16      lnBuf = LangSet.Entry("mbtHelp")
17      If Not LangSet.EntryNotFound Then frmMain.btHelp.Caption = lnBuf

18      lnBuf = LangSet.Entry("mbtNew")
19      If Not LangSet.EntryNotFound Then frmMain.btNew.Caption = lnBuf

20      lnBuf = LangSet.Entry("mbtSave")
21      If Not LangSet.EntryNotFound Then frmMain.btSave.Caption = lnBuf

22      lnBuf = LangSet.Entry("mbtSettings")
23      If Not LangSet.EntryNotFound Then frmMain.btSettings.Caption = lnBuf

24      lnBuf = LangSet.Entry("mbtStop")
25      If Not LangSet.EntryNotFound Then frmMain.btStop.Caption = lnBuf

26      lnBuf = LangSet.Entry("mbtUp")
27      If Not LangSet.EntryNotFound Then frmMain.btUp.Caption = lnBuf

28      lnBuf = LangSet.Entry("mckAssign")
29      If Not LangSet.EntryNotFound Then frmMain.ckAssign.Caption = lnBuf

30      lnBuf = LangSet.Entry("mckAssVol")
31      If Not LangSet.EntryNotFound Then frmMain.ckAssVol.Caption = lnBuf

32      lnBuf = LangSet.Entry("mckAutoMix")
33      If Not LangSet.EntryNotFound Then frmMain.ckAutoMix.Caption = lnBuf

34      lnBuf = LangSet.Entry("mckAutoRep")
35      If Not LangSet.EntryNotFound Then frmMain.ckAutoRep.Caption = lnBuf

36      lnBuf = LangSet.Entry("mckLoopI")
37      If Not LangSet.EntryNotFound Then frmMain.ckLoopI.Caption = lnBuf

38      lnBuf = LangSet.Entry("mckMixByTa")
39      If Not LangSet.EntryNotFound Then frmMain.ckMixByTa.Caption = lnBuf

40      lnBuf = LangSet.Entry("mckPlayStream")
41      If Not LangSet.EntryNotFound Then frmMain.ckPlayStream.Caption = lnBuf

42      lnBuf = LangSet.Entry("mckTmAnn")
43      If Not LangSet.EntryNotFound Then frmMain.ckTmAnn.Caption = lnBuf

44      lnBuf = LangSet.Entry("mckTouch")
45      If Not LangSet.EntryNotFound Then frmMain.ckTouch.Caption = lnBuf

46      lnBuf = LangSet.Entry("mlblRemLoc")
47      If Not LangSet.EntryNotFound Then frmMain.lblRemLoc.Caption = lnBuf

48      lnBuf = LangSet.Entry("mlblL")
49      If Not LangSet.EntryNotFound Then frmMain.lblL.Caption = lnBuf

50      lnBuf = LangSet.Entry("mlblSecond")
51      If Not LangSet.EntryNotFound Then frmMain.lblSecond.Caption = lnBuf

52      lnBuf = LangSet.Entry("mlblMinute")
53      If Not LangSet.EntryNotFound Then frmMain.lblMinute.Caption = lnBuf

54      lnBuf = LangSet.Entry("mlblDelay")
55      If Not LangSet.EntryNotFound Then frmMain.lblDelay.Caption = lnBuf

56      lnBuf = LangSet.Entry("mlblR")
57      If Not LangSet.EntryNotFound Then frmMain.lblR.Caption = lnBuf

58      lnBuf = LangSet.Entry("mlblRemLoc")
59      If Not LangSet.EntryNotFound Then frmMain.lblRemLoc.Caption = lnBuf

60      lnBuf = LangSet.Entry("mlblVole")
61      If Not LangSet.EntryNotFound Then frmMain.lblVole.Caption = lnBuf

62      lnBuf = LangSet.Entry("mmnuAss")
63      If Not LangSet.EntryNotFound Then frmMain.mnuAss.Caption = lnBuf

64      lnBuf = LangSet.Entry("mmnuClr")
65      If Not LangSet.EntryNotFound Then frmMain.mnuClr.Caption = lnBuf

66      lnBuf = LangSet.Entry("mmnuLoop")
67      If Not LangSet.EntryNotFound Then frmMain.mnuLoop.Caption = lnBuf

68      lnBuf = LangSet.Entry("mmnuPause")
69      If Not LangSet.EntryNotFound Then frmMain.mnuPause.Caption = lnBuf

70      lnBuf = LangSet.Entry("mmnuPlay")
71      If Not LangSet.EntryNotFound Then frmMain.mnuPlay.Caption = lnBuf

72      lnBuf = LangSet.Entry("mmnuTa")
73      If Not LangSet.EntryNotFound Then frmMain.mnuTa.Caption = lnBuf

74      lnBuf = LangSet.Entry("mmnuVol")
75      If Not LangSet.EntryNotFound Then frmMain.mnuVol.Caption = lnBuf

76      lnBuf = LangSet.Entry("mTabCt0")
77      If Not LangSet.EntryNotFound Then frmMain.TabCt.TabCaption(0) = lnBuf

78      lnBuf = LangSet.Entry("mTabCt1")
79      If Not LangSet.EntryNotFound Then frmMain.TabCt.TabCaption(1) = lnBuf

80      lnBuf = LangSet.Entry("mTabCt2")
81      If Not LangSet.EntryNotFound Then frmMain.TabCt.TabCaption(2) = lnBuf

82      lnBuf = LangSet.Entry("mTabCt3")
83      If Not LangSet.EntryNotFound Then frmMain.TabCt.TabCaption(3) = lnBuf

Exit_Routine:
    Exit Function
Error_Routine:
    Debug.Assert False
    err.Raise err.Number, "frmMain.inStrings(" & Erl & "):" & err.Source, err.Description
End Function

Private Sub b_Jingle_Click(Index As Integer)

        On Error GoTo Error_Routine

1       If b_Jingle(Index).value = vbUnchecked Then
2           b_Jingle(Index).BackColor = b_color(Index) 'vbButtonFace
3       Else
4           b_color(Index) = b_Jingle(Index).BackColor

5           If Jing(Index).inDebt Then
6               b_Jingle(Index).BackColor = &HFFFFFF
7           Else
8               b_Jingle(Index).BackColor = &HC0FFC0
9           End If
10      End If

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    err.Raise err.Number, "frmMain.b_Jingle_Click(" & Erl & "):" & err.Source, err.Description
End Sub

Private Sub b_Jingle_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    'Call clAssign(Index, Source)
    On Error GoTo Error_Routine

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    err.Raise err.Number, "frmMain.b_Jingle_DragDrop(" & Erl & "):" & err.Source, err.Description
End Sub

Private Sub b_Jingle_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
        On Error GoTo Error_Routine

1       Call clAssign(Index, Source)

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    err.Raise err.Number, "frmMain.b_Jingle_DragOver(" & Erl & "):" & err.Source, err.Description
End Sub

Private Sub b_Jingle_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

        On Error GoTo Error_Routine

1       If Shift = 1 Or Shift = 2 Or Shift = 4 Then
2           Exit Sub
3       End If
4       If KeyCode = vbKeySpace And ckTouch.value = vbUnchecked Then
5           Call btClic(Index, b_Jingle(Index))
6       ElseIf KeyCode = vbKeySpace And ckTouch.value = vbChecked And Jing(Index).OnAir = False Then
7           Call btClic(Index, b_Jingle(Index))
8       End If
9       If KeyCode = vbKeyDelete And b_Jingle(Index).value = vbUnchecked Then
10          butMenuIdx = Index
11          Set butMenu = b_Jingle(Index)
12          Call mnuClr_Click
13      End If
14      Exit Sub
15 err:

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    err.Raise err.Number, "frmMain.b_Jingle_KeyDown(" & Erl & "):" & err.Source, err.Description
End Sub

Private Sub b_Jingle_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
        On Error GoTo Error_Routine

1       If frmMain.ckTouch.value = vbChecked And KeyCode = vbKeySpace Then
2           Call clTouchStop(Index, b_Jingle(Index))
3       End If

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    err.Raise err.Number, "frmMain.b_Jingle_KeyUp(" & Erl & "):" & err.Source, err.Description
End Sub

Private Sub b_Jingle_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

        On Error GoTo Error_Routine

1       If Button = 2 Or Button = 4 Then Exit Sub
        'If Shift = 1 Or Shift = 2 Or Shift = 4 Then Exit Sub
2       Call btClic(Index, b_Jingle(Index))
3       Exit Sub
4 err:

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    err.Raise err.Number, "frmMain.b_Jingle_MouseDown(" & Erl & "):" & err.Source, err.Description
End Sub
Private Sub b_Jingle_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

        On Error GoTo Error_Routine

1       If frmMain.ckTouch.value = vbChecked Then
2           Call clTouchStop(Index, b_Jingle(Index))
3       End If
4       If Button = 2 Then
5           Dim mnbold As Menu
6           tmoutMnu = 0
7           mnuOn = True

8           If b_Jingle(Index).Caption = vbNullString Then
                mnuAss.Visible = True
9               mnuLoop.Visible = False
10              mnuClr.Visible = False
11              mnuspc2.Visible = False
12              mnuspc1.Visible = False
13              mnuVol.Visible = False
14              mnuTa.Visible = False
15              mnuPlay.Visible = False
16              mnuPause.Visible = False
17          Else
18              mnuLoop.Visible = True
19              mnuClr.Visible = True
20              mnuspc2.Visible = True
21              mnuspc1.Visible = True
22              mnuVol.Visible = True
23              mnuTa.Visible = True
24              mnuPlay.Visible = True
25              mnuPause.Visible = True

26              If Jing(Index).OnAir Then
27                  mnuLoop.Visible = False
28                  mnuAss.Visible = False
29                  mnuClr.Visible = False
30                  mnuspc2.Visible = False
31                  mnuspc1.Visible = True
32                  mnuPause.Visible = True
33                  If ckAutoRep Then
34                      mnuPlay.Caption = LangSet.Entry("mAutorep", , Language)
35                  ElseIf ckTouch Then
36                      mnuPlay.Caption = LangSet.Entry("mTouch", , Language)
37                  Else
38                      mnuPlay.Caption = LangSet.Entry("mStop", , Language)
39                  End If
40              Else
41                  mnuAss.Visible = True
42                  mnuLoop.Visible = True
43                  mnuClr.Visible = True
44                  mnuspc2.Visible = True
45                  mnuspc1.Visible = True
46                  mnuPlay.Caption = LangSet.Entry("mPlay", , Language)
47                  mnuPause.Visible = False
48              End If

49          End If

50          If p_Jingle(Index).Visible Then
51              mnuLoop.Caption = LangSet.Entry("mLoopC", , Language)
52          Else
53              mnuLoop.Caption = LangSet.Entry("mLoopS", , Language)
54          End If

55          If ckTmAnn Then
56              mnuTa.Enabled = True
57          Else
58              mnuTa.Enabled = False
59          End If

60          butMenuIdx = Index
61          Set butMenu = b_Jingle(Index)
62          If mnuPlay.Visible Then
63              If BASS_ChannelIsActive(Jing(butMenuIdx).Strm) = BASS_ACTIVE_PAUSED Then
64                  PopupMenu mnujp, , , , mnuPause
65              Else
66                  PopupMenu mnujp, , , , mnuPlay
67              End If
68          Else
69              PopupMenu mnujp
70          End If
71      End If
72      Exit Sub
73 err:

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    err.Raise err.Number, "frmMain.b_Jingle_MouseUp(" & Erl & "):" & err.Source, err.Description
End Sub
Private Sub b_tmAnJin_Click()

        '    CkAutoRep.value = 0
        On Error GoTo Error_Routine

1       If b_tmAnJin.value = vbChecked And b_tmAnJin.Caption <> LangSet.Entry("mb_tmAnJin", , Language) Then
            'b_tmAnJin.value = vbUnchecked
2           Exit Sub
3       End If

4       If b_tmAnJin.value = vbUnchecked Then
5           b_tmAnJin.BackColor = vbButtonFace       '&H8000000F&
6           ManiPulate = False
7           b_tmAnJin.Caption = LangSet.Entry("mb_tmAnJin", , Language)
8           tmJing = ""
9           For i = 0 To 29
10              b_Jingle(i).BackColor = Jing(i).Color 'vbButtonFace
11          Next i
12          tmStat1 = False
13      Else
14          tmout = 0
15          ManiPulate = True
16          ckLoopI.value = vbUnchecked
17          ckAssign.value = vbUnchecked
18          ckAssVol.value = vbUnchecked
19          For i = 0 To 29
20              Jing(i).Color = b_Jingle(i).BackColor
21              b_tmAnJin.BackColor = &HFFC0C0
22              b_Jingle(i).BackColor = &HFFC0C0
23          Next i
24      End If
25      Exit Sub
26 err:

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    err.Raise err.Number, "frmMain.b_tmAnJin_Click(" & Erl & "):" & err.Source, err.Description
End Sub

Private Sub BtAbout_Click()
        On Error GoTo Error_Routine

1       If Not ckAssVol.Visible Then
2           btAbout.Caption = LangSet.Entry("mbtAbout", , Language)
3           btAbout.MaskColor = RGB(0, 0, 255)
4           btAbout.picture = LoadResPicture(126, vbResBitmap)
5           btAbout.BackColor = vbButtonFace
6           ckAssVol.value = vbUnchecked
7           ckAssVol.Visible = True
8           PalSet.Entry("Volm_" & VolIndex, , lstPal.Text) = Jing(VolIndex).volume
9           ckAssVol.SetFocus
10      Else
11          frmAbout.Show vbModal, Me
12      End If

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.BtAbout_Click"
    Resume Exit_Routine
End Sub

Private Sub evBtDn()

        On Error GoTo Error_Routine

1       If lstPal.ListIndex >= 0 And lstPal.ListIndex < lstPal.ListCount - 1 Then
2           lstPal.ListIndex = lstPal.ListIndex + 1
3       End If
4       If lstPal.ListIndex = -1 And lstPal.ListCount > 0 Then
5           lstPal.ListIndex = 0
6       End If
7       Exit Sub
8 err:

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    err.Raise err.Number, "frmMain.evBtDn(" & Erl & "):" & err.Source, err.Description
End Sub

Private Sub BtDn_KeyDown(KeyCode As Integer, Shift As Integer)
        On Error GoTo Error_Routine

1       evBtDn

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.BtDn_KeyDown"
    Resume Exit_Routine
End Sub

Private Sub BtDn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        On Error GoTo Error_Routine

1       evBtDn

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.BtDn_MouseDown"
    Resume Exit_Routine
End Sub

Private Sub btHelp_Click()
        On Error GoTo Error_Routine

1       ExecuteDocument App.Path & "\Jingle_Palette.chm", impShowMaximized

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.btHelp_Click"
    Resume Exit_Routine
End Sub

Private Sub evStop()

        On Error GoTo Error_Routine

1       btAbout.Caption = LangSet.Entry("mbtAbout", , Language)
2       ckAssVol.value = vbUnchecked
3       ckAssVol.Visible = True
4       ckAssign.value = vbUnchecked 'exit assign mode

5       For i = 0 To 29
6           Jing(i).OnAir = False
7           Call BASS_ChannelStop(Jing(i).Strm)
8           Call BASS_ChannelSetPosition(Jing(i).Strm, 0)
9       Next i

10      Call BASS_ChannelStop(tmSnStrm)
11      Call BASS_ChannelSetPosition(tmSnStrm, 0)
12      Call BASS_ChannelStop(tmJnStrm)
13      Call BASS_ChannelSetPosition(tmJnStrm, 0)

        '    BASS_Free   'stop everything and free all player resources
        '    InitDev     'start digital output
        '    OpenFile lstPal.Text 'reload palette

14      VuL.Level = 0
15      VuR.Level = 0
16      Exit Sub
17 err:

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    err.Raise err.Number, "frmMain.evStop(" & Erl & "):" & err.Source, err.Description
End Sub
Private Sub btStop_KeyDown(KeyCode As Integer, Shift As Integer)
        On Error GoTo Error_Routine

1       evStop

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.btStop_KeyDown"
    Resume Exit_Routine
End Sub

Private Sub btStop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        On Error GoTo Error_Routine

1       evStop

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.btStop_MouseDown"
    Resume Exit_Routine
End Sub

Private Sub BtUp_KeyDown(KeyCode As Integer, Shift As Integer)
        On Error GoTo Error_Routine

1       evBtUp

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.BtUp_KeyDown"
    Resume Exit_Routine
End Sub

Private Sub BtUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        On Error GoTo Error_Routine

1       evBtUp

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.BtUp_MouseDown"
    Resume Exit_Routine
End Sub

Private Sub evBtUp()

        On Error GoTo Error_Routine

1       If lstPal.ListIndex > 0 And lstPal.ListIndex <= lstPal.ListCount Then
2           lstPal.ListIndex = lstPal.ListIndex - 1
3       End If
4       If lstPal.ListIndex = -1 Then
5           lstPal.ListIndex = lstPal.ListCount - 1
6       End If
7       Exit Sub
8 err:

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    err.Raise err.Number, "frmMain.evBtUp(" & Erl & "):" & err.Source, err.Description
End Sub
Private Sub CkAssign_Click()

        '    CkAutoRep.value = 0
        On Error GoTo Error_Routine

1       If ckAssign.value = vbUnchecked Then
2           ckAssign.BackColor = vbButtonFace       '&H8000000F&
3           ManiPulate = False
4           For i = 0 To 29
5               If Jing(i).OnAir = True Then
6                   b_Jingle(i).value = vbChecked
7               End If
8               b_Jingle(i).BackColor = Jing(i).Color 'vbButtonFace
9           Next i
10      Else
11          tmout = 0
12          ckLoopI.value = vbUnchecked
13          b_tmAnJin.value = vbUnchecked
14          ckAssVol.value = vbUnchecked
15          ManiPulate = True
16          For i = 0 To 29
                'If Not Jing(i).OnAir Then
17              Jing(i).Color = b_Jingle(i).BackColor
18              b_Jingle(i).value = vbUnchecked
19              b_Jingle(i).BackColor = &HFFC0FF
                'End If
20          Next i
21      End If
22      Exit Sub
23 err:

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.CkAssign_Click"
    Resume Exit_Routine
End Sub

Private Sub btNew_Click()

        On Error GoTo Error_Routine

1       For i = 0 To 29
2           Jing(i).Path = ""       'Null the jingle variables
3           Call BASS_StreamFree(Jing(i).Strm)
4           frmMain.b_Jingle(i).Caption = btCaption(Jing(i).Path, frmMain.b_Jingle(i))
5       Next i
6       lstPal.ListIndex = -1
7       Exit Sub
8 err:

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.btNew_Click"
    Resume Exit_Routine
End Sub

Private Sub btsettings_Click()
        On Error GoTo Error_Routine

1       frmSett.Show vbModal, Me

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.btsettings_Click"
    Resume Exit_Routine
End Sub

Private Sub btExit_Click()
        On Error GoTo Error_Routine

1       exitBut = True
2       Unload frmMain

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.btExit_Click"
    Resume Exit_Routine
End Sub
Private Sub ckAssVol_Click()

        '    CkAutoRep.value = 0
        On Error GoTo Error_Routine

1       If ckAssVol.value = vbUnchecked Then
2           ckAssVol.BackColor = vbButtonFace       '&H8000000F&
3           ManiPulate = False
4           ManiPulVol = True
5           For i = 0 To 29
6               b_Jingle(i).BackColor = Jing(i).Color 'vbButtonFace
7           Next i
8           ckAssVol.BackColor = vbButtonFace
9       Else
10          tmout = 0
11          ManiPulate = True
12          ManiPulVol = True
13          ckLoopI.value = vbUnchecked
14          b_tmAnJin.value = vbUnchecked
15          ckAssign.value = vbUnchecked
16          For i = 0 To 29
17              Jing(i).Color = b_Jingle(i).BackColor
18              b_Jingle(i).BackColor = &HFFFF80   '&HFFFF00
19          Next i
20          ckAssVol.BackColor = &HFFFF80
21      End If
22      Exit Sub
23 err:

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.ckAssVol_Click"
    Resume Exit_Routine
End Sub

Private Sub CkAutoMix_Click()
        On Error GoTo Error_Routine

1       If ckAutoMix.value = vbUnchecked Then
2           ckAutoMix.BackColor = vbButtonFace       '&H8000000F&
3       Else
4           ckAutoMix.BackColor = &HFFFF00    '&HFFFF&
            'CkTouch.value = vbUnchecked
5       End If

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.CkAutoMix_Click"
    Resume Exit_Routine
End Sub

Private Sub CkAutoRep_Click()
        On Error GoTo Error_Routine

1       If ckAutoRep.value = vbUnchecked Then
2           ckAutoRep.BackColor = vbButtonFace       '&H8000000F&
3       Else
4           ckAutoRep.BackColor = &HFFFF&  '&HFFFF&
5           ckTouch.value = vbUnchecked
6       End If

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.CkAutoRep_Click"
    Resume Exit_Routine
End Sub

Private Sub CkTouch_Click()
        On Error GoTo Error_Routine

1       If ckTouch.value = vbUnchecked Then
2           ckTouch.BackColor = vbButtonFace       '&H8000000F&
3       Else
4           ckTouch.BackColor = &HFF00& ' &HFFFF00
5           ckAutoRep.value = vbUnchecked
6       End If
        'b_Jingle(Focused).SetFocus

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.CkTouch_Click"
    Resume Exit_Routine
End Sub

Private Sub ckLoopI_Click()
        On Error GoTo Error_Routine

1       If ckLoopI.value = vbUnchecked Then
2           ckLoopI.BackColor = vbButtonFace
3           ManiPulate = False
4           For i = 0 To 29
5               b_Jingle(i).BackColor = Jing(i).Color 'vbButtonFace
6           Next i
7       Else
8           tmout = 0
9           ManiPulate = True
10          ckAssign.value = vbUnchecked
11          b_tmAnJin.value = vbUnchecked
12          ckAssVol.value = vbUnchecked
13          For i = 0 To 29
14              Jing(i).Color = b_Jingle(i).BackColor
15              b_Jingle(i).BackColor = &H80C0FF
16          Next i
17          ckLoopI.BackColor = &H80C0FF
18      End If

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.ckLoopI_Click"
    Resume Exit_Routine
End Sub

Private Sub ckMixByTa_Click()
        On Error GoTo Error_Routine

1       If ckMixByTa.value = vbUnchecked Then
2           ckMixByTa.BackColor = vbButtonFace       '&H8000000F&
3       Else
4           ckMixByTa.BackColor = &HFFFF00    '&HFFFF&
5           If ckTmAnn.value = vbUnchecked Then
6               ckMixByTa.value = vbUnchecked
7           End If
8       End If

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.ckMixByTa_Click"
    Resume Exit_Routine
End Sub

Private Sub ckPlayStream_Click()
        On Error GoTo Error_Routine

1       If ckPlayStream.value = vbUnchecked Then
2           ckPlayStream.BackColor = vbButtonFace
3           cmbStream.Enabled = True
4           Call BASS_ChannelStop(LiveStream)
5           Call BASS_StreamFree(LiveStream)
6           ckMixByTa.value = vbUnchecked
7           Timer1.Enabled = True
8           tm1 = 0
            LiveOn = False
9       Else
            LiveOn = True
10          txtConDisp.Text = LangSet.Entry("mStrCon", , Language)
11          ckPlayStream.BackColor = &HC0FFC0
12          cmbStream.Enabled = False
13          DoEvents
14          If LCase$(Left$(cmbStream.Text, 7)) = "http://" Or LCase$(Left$(cmbStream.Text, 6)) = "ftp://" Then
15              LiveStream = BASS_StreamCreateURL(cmbStream.Text, 0, BASS_STREAM_AUTOFREE, 0, 0)
16              If LiveStream = 0 Then
17                  Call ErrorLive_(LangSet.Entry("mError", , Language))
18              Else
19                  Call BASS_ChannelPreBuf(LiveStream)
20                  If BASS_ChannelPlay(LiveStream, BASSFALSE) = BASSFALSE Then
21                      Call ErrorLive_("Can't play remote stream")
22                  End If
23                  txtConDisp.Text = LangSet.Entry("mStrPla1", , Language) & btCaption(cmbStream.Text, cmbStream, 20, "/") & LangSet.Entry("mStrPla2", , Language)
24                  'VuStrm = LiveStream
25                  UrlAdd cmbStream.Text
26              End If
27          Else
28              txtConDisp.Text = LangSet.Entry("mStrNoAd", , Language)
29          End If
30      End If

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.ckPlayStream_Click"
    Resume Exit_Routine
End Sub

Private Sub CkTmAnn_Click()

        '    CkAutoRep.value = 0
        On Error GoTo Error_Routine

1       If ckTmAnn.value = False Then
2           tmJing = ""
3           ckTmAnn.BackColor = vbButtonFace       '&H8000000F&
4           tmStat = False
5           ckTmAnn.Caption = LangSet.Entry("mckTmAnn", , Language)
6           b_tmAnJin.Enabled = False
7           DTDelay.Enabled = False
8           lblDelay.Enabled = False
9           lblTa.Visible = False
10          BASS_StreamFree (tmSnStrm)
11          BASS_StreamFree (tmJnStrm)
12          b_tmAnJin.Caption = LangSet.Entry("mb_tmAnJin", , Language)
13          b_tmAnJin.value = vbUnchecked
14          b_tmAnJin.BackColor = vbButtonFace
15          For i = 0 To 29
16              b_Jingle(i).BackColor = Jing(i).Color
17          Next i
18          ckMixByTa.value = vbUnchecked
19          Timer4.Interval = aMixtime / 100
20      Else
21          ckTmAnn.BackColor = &HFFC0C0
22          ckTmAnn.Caption = LangSet.Entry("mTaReady", , Language) & " " & Format(time, "HH") & ":" & Format(tmMin, "00") & ":" & Format(tmSec, "00")
23          b_tmAnJin.Enabled = True
24          DTDelay.Enabled = True
25          lblDelay.Enabled = True
26          lblTa.Visible = True

27          Call BASS_StreamFree(tmSnStrm)
28          tmSnStrm = BASS_StreamCreateFile(BASSFALSE, tmSound, 0, 0, 0)
29          If tmSnStrm = 0 Then
30              Call Error_("Can't create stream for time announce signal")
31              ckTmAnn.value = vbUnchecked
32          Else
33              Call BASS_ChannelPreBuf(tmSnStrm)
34              tmDelDif = BASS_ChannelBytes2Seconds(tmSnStrm, BASS_StreamGetLength(tmSnStrm))
35              If DTDelay.Second > tmDelDif Then tmDel = tmDelDif '- 1
                'tmDel = Int(tmDelDif) - 1
                'DTDelay.Second = tmDel

36          End If

37      End If

38      Exit Sub
39 err:

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.CkTmAnn_Click"
    Resume Exit_Routine
End Sub
Private Sub CmbStream_GotFocus()
        On Error GoTo Error_Routine

1       frmMain.KeyPreview = False

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.CmbStream_GotFocus"
    Resume Exit_Routine
End Sub

Private Sub CmbStream_LostFocus()
        On Error GoTo Error_Routine

1       frmMain.KeyPreview = True

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.CmbStream_LostFocus"
    Resume Exit_Routine
End Sub
Private Sub DTDelay_Change()
        On Error GoTo Error_Routine

1       If DTDelay.Second > tmDelDif Then DTDelay.Second = Format(Int(tmDelDif) - 1, "ss")
2       tmDel = DTDelay.Second
3       If frmMain.b_tmAnJin Then frmMain.b_tmAnJin.Caption = btCaption(tmJing, frmMain.b_tmAnJin, 26) & " " & tmDel & " " & LangSet.Entry("mTaLater", , Language)

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.DTDelay_Change"
    Resume Exit_Routine
End Sub

Private Sub DTMinute_Change()
        On Error GoTo Error_Routine

1       tmMin = DTMinute.Minute
2       If frmMain.ckTmAnn Then ckTmAnn.Caption = LangSet.Entry("mTaReady", , Language) & " " & Format(time, "HH") & ":" & Format(tmMin, "00") & ":" & Format(tmSec, "00")

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.DTMinute_Change"
    Resume Exit_Routine
End Sub

Private Sub DTSecond_Change()
        On Error GoTo Error_Routine

1       tmSec = DTSecond.Second
2       If frmMain.ckTmAnn Then ckTmAnn.Caption = LangSet.Entry("mTaReady", , Language) & " " & Format(time, "HH") & ":" & Format(tmMin, "00") & ":" & Format(tmSec, "00")

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.DTSecond_Change"
    Resume Exit_Routine
End Sub



Private Sub Form_Paint()
        On Error GoTo Error_Routine

1       If PaintFirst Then
2           PaintFirst = False
3           If RunFirst Then
4               SelectListItemByString lstPal, "Demo"  'if program runs for the 1st time and Demo palette is there, load it
5               ErrDisp "save", , "Default Settings Loaded"
6           Else
7               SelectListItemByString lstPal, IniSet.Entry("PaletteIndex")
8           End If
            'frmSplash.vuLoad.Level = 100
            'Unload frmSplash
9       End If

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.Form_Paint"
    Resume Exit_Routine
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        On Error GoTo Error_Routine

1       If UnloadMode = 0 And Not exitBut Then
2           Cancel = True
3           btExit_Click
4       End If

        'Cancel = True

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.Form_QueryUnload"
    Resume Exit_Routine
End Sub

Private Sub Form_Resize()
        'If frmMain.Width < 12000 Or frmMain.Height < 8565 Then
        '    If frmMain.WindowState = 0 Then frmMain.Width = 12000
        '    If frmMain.WindowState = 0 Then frmMain.Height = 8565
        '    Exit Sub
        'End If
        On Error GoTo Error_Routine
1       If Not frmMain.WindowState <> 1 Then Exit Sub

2       Dim bjinprevLeft As Long, bjinprevTop As Long
3       bjinprevLeft = 120
4       bjinprevTop = 240

5       fraJingles.Height = frmMain.Height - fraPanel.Height - 650
6       fraJingles.Width = frmMain.Width - fraVu.Width - 420

7       For i = 0 To 29
8           b_Jingle(i).Left = bjinprevLeft
9           p_Jingle(i).Left = bjinprevLeft + 60
10          bjinprevLeft = b_Jingle(i).Left + (fraJingles.Width - 170) / 5
11          If i = 4 Or i = 9 Or i = 14 Or i = 19 Or i = 24 Then
12              bjinprevLeft = 120
13          End If
14          If i >= 0 And i <= 4 Then
15              bjinprevTop = 240
16          ElseIf i >= 5 And i <= 9 Then
17              bjinprevTop = b_Jingle(0).top + (fraJingles.Height - 300) / 6
18          ElseIf i >= 10 And i <= 14 Then
19              bjinprevTop = b_Jingle(6).top + (fraJingles.Height - 300) / 6
20          ElseIf i >= 15 And i <= 19 Then
21              bjinprevTop = b_Jingle(11).top + (fraJingles.Height - 300) / 6
22          ElseIf i >= 20 And i <= 24 Then
23              bjinprevTop = b_Jingle(16).top + (fraJingles.Height - 300) / 6
24          ElseIf i >= 25 And i <= 29 Then
25              bjinprevTop = b_Jingle(21).top + (fraJingles.Height - 300) / 6
26          End If
27          b_Jingle(i).top = bjinprevTop
28          p_Jingle(i).top = bjinprevTop + 60
29          b_Jingle(i).Height = CInt((fraJingles.Height - 800) / 6)
30          b_Jingle(i).Width = CInt((fraJingles.Width - 600) / 5)
31          If b_Jingle(i).Width > 2000 And b_Jingle(i).Width <= 2500 Then
32              b_Jingle(i).FontSize = 8
33          ElseIf b_Jingle(i).Width > 2500 And b_Jingle(i).Width <= 3500 Then
34              b_Jingle(i).FontSize = 10
35          ElseIf b_Jingle(i).Width > 3500 And b_Jingle(i).Width <= 4000 Then
36              b_Jingle(i).FontSize = 12
37          End If
38      Next i

39      fraVu.Left = frmMain.Width - fraVu.Width - 190
40      fraVu.Height = frmMain.Height - fraPanel.Height - 650
41      VuL.Height = fraVu.Height - 400
42      VuR.Height = fraVu.Height - 400
43      lblL.top = fraVu.Height - 250
44      lblR.top = fraVu.Height - 250

45      lstPal.Width = (frmMain.Width / 2 - 600) / 2 - 400
46      lstPal.Left = (frmMain.Width / 2 - 600) - lstPal.Width - 120
47      TabCt.top = frmMain.Height - TabCt.Height - 670
48      TabCt.Width = frmMain.Width / 2 - 600
49      If TabCt.Tab = 0 Then
50          lstPal.Visible = True
51      Else
52          lstPal.Visible = False
53      End If
54      ckTmAnn.Width = TabCt.Width - 2000
55      b_tmAnJin.Width = TabCt.Width - 2000
56      ln1.X1 = 120
57      ln1.X2 = TabCt.Width - 230
58      If TabCt.Tab = 1 Then
59          ln1.Visible = True
60      Else
61          ln1.Visible = False
62      End If
63      cmbStream.Width = TabCt.Width - 230
64      txtConDisp.Width = TabCt.Width - 230
65      ckMixByTa.Left = TabCt.Width - ckMixByTa.Width - 140
66      ckPlayStream.Width = b_Jingle(0).Width
67      If TabCt.Tab = 2 Then
68          ckMixByTa.Visible = True
69      Else
70          ckMixByTa.Visible = False
71      End If
72      ckAssVol.Width = TabCt.Width - 280
73      SlideVolCh.Width = TabCt.Width - 280
74      lblVol.Left = TabCt.Width - lblVol.Width - 180
75      lblVolexp.Width = TabCt.Width - lblVole.Width - lblVol.Width - 240
76      ckLoopI.Width = TabCt.Width / 4 - 150
77      btSettings.Width = ckLoopI.Width
78      btHelp.Width = ckLoopI.Width
79      btAbout.Width = ckLoopI.Width
80      btSettings.Left = ckLoopI.Width + 220
81      btHelp.Left = btSettings.Left + btSettings.Width + 110
82      btAbout.Left = btHelp.Left + btHelp.Width + 110
83      If TabCt.Tab = 3 Then
84          lblVol.Visible = True
85          btSettings.Visible = True
86          btHelp.Visible = True
87          btAbout.Visible = True
88      Else
89          lblVol.Visible = False
90          btSettings.Visible = False
91          btHelp.Visible = False
92          btAbout.Visible = False
93      End If

94      ckAssign.Width = TabCt.Width / 2 + 70
95      btSave.Width = ckAssign.Width
96      btNew.Width = ckAssign.Width
97      txtSave.Width = btSave.Width - 80

98      fraPanel.Width = CInt(frmMain.Width / 2 + 200)
99      fraPanel.Left = frmMain.Width - fraPanel.Width - 190
100     fraPanel.top = frmMain.Height - fraPanel.Height - 650
101     ckAutoRep.Left = CInt((fraPanel.Width - 150) / 5 + 120)
102     picDispBack.Left = ckAutoRep.Left
103     ckTouch.Left = ckAutoRep.Left + CInt((fraPanel.Width - 150) / 5)
104     ckAutoMix.Left = ckTouch.Left + CInt((fraPanel.Width - 150) / 5)
105     btStop.Left = ckAutoMix.Left + CInt((fraPanel.Width - 150) / 5)
106     btExit.Left = btStop.Left
107     btDn.Width = CInt((fraPanel.Width - 700) / 5)
108     ckAutoRep.Width = btDn.Width
109     ckTouch.Width = btDn.Width
110     ckAutoMix.Width = btDn.Width
111     btStop.Width = btDn.Width
112     btUp.Width = btDn.Width
113     btExit.Width = btDn.Width
114     picDispBack.Width = fraPanel.Width - btExit.Width - btUp.Width - 490
115     lblTime.Left = picDispBack.Width - lblTime.Width - 100
116     lblDate.Left = picDispBack.Width - lblDate.Width - 100
117     lblDatex.Left = picDispBack.Width - lblDatex.Width - 100
118     lblWeek.Left = CInt(picDispBack.Width / 2 - 300)

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.Form_Resize"
    Resume Exit_Routine
End Sub
Private Sub lstpal_Click()

        On Error GoTo Error_Routine

1       ErrDisp "wait", False, "Loading palette"
2       SelPalette = lstPal.Text
3       OpenFile lstPal.Text

4       lblPalPr.Caption = lstPal.List(lstPal.ListIndex - 1)
5       If lblPalPr.Caption = "" Then lblPalPr.Caption = "<none>"
6       lblPalPr.ToolTipText = "Previous palette: " & lblPalPr.Caption
7       txtDisPal.Caption = lstPal.Text
8       txtDisPal.ToolTipText = "Currecntly selected palette: " & lstPal.Text
9       lblPalNext.Caption = lstPal.List(lstPal.ListIndex + 1)
10      If lblPalNext.Caption = "" Then lblPalNext.Caption = "<none>"
11      lblPalNext.ToolTipText = "Next palette: " & lblPalNext.Caption

12      ErrDisp "hide", , ""
13      txtSave.Text = ""
14      b_tmAnJin.value = vbUnchecked

        '    For i = 0 To 29
        '        b_Jingle(i).value = vbUnchecked
        '    Next i

15      btAbout.Caption = LangSet.Entry("mbtAbout", , Language)
16      ckAssVol.value = vbUnchecked
17      ckAssVol.Visible = True

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.lstpal_Click"
    Resume Exit_Routine
End Sub
Private Sub lstpal_KeyDown(KeyCode As Integer, Shift As Integer)

        On Error GoTo Error_Routine

1       Dim idx As Integer
2       idx = lstPal.ListIndex
3       If KeyCode = vbKeyDelete Then
4           If MsgBox(LangSet.Entry("mMsdp", , Language) & lstPal.Text & "'?", 292, "Jingle Palette") = vbYes Then
5               PalSet.DeleteSectionKey lstPal.Text
6               lstPalRefresh
7               If idx = 0 And lstPal.ListCount <> 0 Then
8                   lstPal.ListIndex = idx
9               ElseIf lstPal.ListCount = 0 Then
10                  lstPal.ListIndex = -1
11              Else
12                  lstPal.ListIndex = idx - 1
13              End If
14          End If
15          lstPal.SetFocus
16      End If

        '    If KeyCode = vbKeyPageUp Then
        '        BtUp_Click
        '    End If
        '
        '    If KeyCode = vbKeyPageDown Then
        '        BtDn_Click
        '    End If
17      Exit Sub
18 err:

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.lstpal_KeyDown"
    Resume Exit_Routine
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

        On Error GoTo Error_Routine

1       Select Case KeyCode

            Case vbKeyEscape
2               Call evStop
3           Case vbKeyPageUp
4               'Call evBtUp
5           Case vbKeyPageDown
6              ' Call evBtDn
7           Case vbKeyR
8               If ckAutoRep.value = vbUnchecked Then
9                   ckAutoRep.value = vbChecked
10              Else
11                  ckAutoRep.value = vbUnchecked
12              End If
13          Case vbKeyH
14              If ckTouch.value = vbUnchecked Then
15                  ckTouch.value = vbChecked
16              Else
17                  ckTouch.value = vbUnchecked
18              End If
19          Case vbKeyM
20              If ckAutoMix.value = vbUnchecked Then
21                  ckAutoMix.value = vbChecked
22              Else
23                  ckAutoMix.value = vbUnchecked
24              End If
25          Case vbKeyE
26              TabCt.Tab = 0
27              btNew_Click
28          Case vbKeyH
29              TabCt.Tab = 2
30              btHelp_Click
31          Case vbKeyN
32              TabCt.Tab = 1
33              If ckTmAnn.value = vbUnchecked Then
34                  ckTmAnn.value = vbChecked
35              Else
36                  ckTmAnn.value = vbUnchecked
37              End If
38          Case vbKeyA
39              TabCt.Tab = 0
40              If ckAssign.value = vbUnchecked Then
41                  ckAssign.value = vbChecked
42              Else
43                  ckAssign.value = vbUnchecked
44              End If
45          Case vbKeyS
46              TabCt.Tab = 0
47              btSave_Click
48          Case vbKeyL
49              TabCt.Tab = 3
                '        If ckLoopI.value = vbUnchecked Then
                '            ckLoopI.value = Not ckLoopI.value 'vbChecked
                '        Else
                '            ckLoopI.value = Not ckLoopI.value 'vbUnchecked
                '        End If
50          Case vbKeyC
51              TabCt.Tab = 1
52              If b_tmAnJin.Visible = True Then
53                  If b_tmAnJin.value = vbUnchecked Then
54                      b_tmAnJin.value = vbChecked
55                  Else
56                      b_tmAnJin.value = vbUnchecked
57                  End If
58              End If
59      End Select
60      Exit Sub
61 err:

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.Form_KeyDown"
    Resume Exit_Routine
End Sub
Public Sub Form_Load()
        ' First, we need to store the address of the existing Message Handler
        On Error GoTo Error_Routine

1       OldWindowProc = GetWindowLong(frmMain.hwnd, GWL_WNDPROC)
2       DeleteMenu GetSystemMenu(Me.hwnd, False), SC_CLOSE, MF_BYCOMMAND

3       Dim palInd As String, palPres As Boolean
4       If RunFirst Then
5           frmMain.top = Screen.Height / 2 - frmMain.Height / 2 - 230
6           frmMain.Left = Screen.Width / 2 - frmMain.Width / 2
7           IniSet.Entry("PosTop") = Screen.Height / 2 - frmMain.Height / 2 - 230
8           IniSet.Entry("PosLeft") = Screen.Width / 2 - frmMain.Width / 2
9       End If

10      frmMain.top = IniSet.Entry("PosTop")
11      frmMain.Left = IniSet.Entry("PosLeft")
12      If frmMain.top < 1 Or frmMain.top + frmMain.Height > Screen.Height - 1 Or frmMain.Left < 1 Or frmMain.Left + frmMain.Width > Screen.Width - 1 Then
13          frmMain.top = Screen.Height / 2 - frmMain.Height / 2 - 230
14          frmMain.Left = Screen.Width / 2 - frmMain.Width / 2
15      End If

16      frmMain.Width = IniSet.Entry("WinWidth")
17      frmMain.Height = IniSet.Entry("WinHeight")
18      frmMain.WindowState = IniSet.Entry("WinState")

        Call SetTopMostWindow(frmMain.hwnd, OnTop)

19      If Language <> "English" Then
20          Call inStrings(Language)
21      End If

        'frmSplash.vuLoad.Level = 50
22      lblTime.Caption = Format(time, TimeForm)
23      lblDatex.Caption = Format(Date, DateForm)

24      InitDev

25      DoEvents
        'frmSplash.vuLoad.Level = 75
26      ckTouch.value = IniSet.Entry("Touch")
27      ckAutoRep.value = IniSet.Entry("Autorepeat")
28      ckAutoMix.value = IniSet.Entry("AutoMix")
29      TabCt.Tab = IniSet.Entry("ActiveTab")
30      Timer4.Interval = aMixtime / 100

        '    If CkTouch.value = vbUnchecked Then
        '        LoopBit = 0
        '    Else
        '        LoopBit = BASS_SAMPLE_LOOP
        '    End If

31      SlideVolCh.ThumbBitmap = LoadResPicture(104, vbResBitmap)

32      btSettings.picture = LoadResPicture(125, vbResBitmap)
33      btSettings.MaskColor = RGB(0, 0, 253)
34      btHelp.picture = LoadResPicture(124, vbResBitmap)
35      btAbout.picture = LoadResPicture(126, vbResBitmap)

36      tmMin = IniSet.Entry("TimeAnnMin")
37      DTMinute.Minute = tmMin
38      tmSec = IniSet.Entry("TimeAnnSec")
39      DTSecond.Second = tmSec
40      tmDel = IniSet.Entry("TimeAnnDel")
41      DTDelay.Second = tmDel

42      Call lstPalRefresh
43      Call UrlLoad

44      DoEvents

45      rotdsp = 0
        'echdsp = 0
46      fladsp = 0
        'swpdsp = 0

47      PaintFirst = True
48      txtConDisp.Text = LangSet.Entry("mStrNoCo", , Language)
49      Tme = time

        ' Now we can tell windows to forward all messages to out own Message Handler
50      If IsNotIDE Then Call SetWindowLong(frmMain.hwnd, GWL_WNDPROC, AddressOf SubClass1_WndMessage)



        'Show
51      Exit Sub
52 err:
53      ErrDisp "no", , "main window load: " & err.Description

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.Form_Load"
    Resume Exit_Routine
End Sub

Private Sub Form_Unload(Cancel As Integer)

        ' We must return control of the messages back to windows before the program exits
        On Error GoTo Error_Routine

1       Call SetWindowLong(Me.hwnd, GWL_WNDPROC, OldWindowProc)
2       Unload frmAbout
3       Unload frmSett
4       IniSet.Entry("PosTop") = frmMain.top
5       IniSet.Entry("PosLeft") = frmMain.Left
6       If frmMain.WindowState = 0 Then IniSet.Entry("WinWidth") = frmMain.Width
7       If frmMain.WindowState = 0 Then IniSet.Entry("WinHeight") = frmMain.Height
8       If frmMain.WindowState <> 1 Then IniSet.Entry("WinState") = frmMain.WindowState
9       IniSet.Entry("Touch") = ckTouch.value
10      IniSet.Entry("AutoMix") = ckAutoMix.value
11      IniSet.Entry("Autorepeat") = ckAutoRep.value
12      If lstPal.Text <> vbNullString Then
13          IniSet.Entry("PaletteIndex") = lstPal.Text
14      End If
15      IniSet.Entry("Volume") = vol
16      IniSet.Entry("TimeAnnouncer") = tmSound
17      IniSet.Entry("TimeAnnMin") = tmMin
18      IniSet.Entry("TimeAnnSec") = tmSec
19      IniSet.Entry("TimeAnnDel") = tmDel
20      IniSet.Entry("ActiveTab") = TabCt.Tab
        IniSet.Entry("AlwaysOnTop") = OnTop

        'Stop digital output
21      Call BASS_Stop

        'Free the streams
22      For i = 0 To 29
23          Call BASS_StreamFree(Jing(i).Strm)
24      Next i
25      Call BASS_StreamFree(tmSnStrm)
26      Call BASS_StreamFree(tmJnStrm)
27      Call BASS_StreamFree(LiveStream)

        'Close digital sound system
28      Call BASS_Free

'29      Unload frmAbout
'30      Unload frmSett

31      ErrLog , "Program normal exit"
32      ErrLog , "***********************************************************"

33      End
34      Exit Sub
35 err:

36      End

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.Form_Unload"
    Resume Exit_Routine
End Sub
Private Sub LblDatex_Change()
        On Error GoTo Error_Routine

1       lblDate.Caption = Format(Date, "dddd")
2       lblWeek.Caption = Format(Date, "ww", vbMonday)

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.LblDatex_Change"
    Resume Exit_Routine
End Sub
Private Sub LblErr_Click()
        'in wingdings, use these: <         ý         ?          x              6
        ' these are             folppy   X framed  checked    X framed(oth)  hourglass
        On Error GoTo Error_Routine

1       lblErr.Visible = False

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.LblErr_Click"
    Resume Exit_Routine
End Sub
Public Sub lstPalRefresh()
        On Error GoTo Error_Routine

1       Dim iPal As Long
2       PalSet.RegistrySubKeys.Update
3       lstPal.Clear
4       For iPal = 1 To PalSet.RegistrySubKeys.Count
5           lstPal.AddItem PalSet.RegistrySubKeys.Key(iPal)
6       Next iPal

        If RunFirst And lstPal.ListCount > 0 Then
            lstPal.ListIndex = 0
        End If

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    err.Raise err.Number, "frmMain.lstPalRefresh(" & Erl & "):" & err.Source, err.Description
End Sub
Private Sub btSave_Click()

        On Error GoTo Error_Routine

1       ckAssign.value = vbUnchecked
2       txtSave.Visible = True
3       txtSave.SetFocus
4       frmMain.KeyPreview = False
5       ErrDisp "wait", True
6       Timer1.Enabled = True
7       tm1 = 0
8       Exit Sub
9 err:

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.btSave_Click"
    Resume Exit_Routine
End Sub
Private Sub mnuAss_Click()
        On Error GoTo Error_Routine

1       TabCt.Tab = 0
2       Call clAssign(butMenuIdx, butMenu)

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.mnuAss_Click"
    Resume Exit_Routine
End Sub

Private Sub mnuClr_Click()
        On Error GoTo Error_Routine

1       If butMenu.Caption = "" Then Exit Sub
2       If MsgBox(LangSet.Entry("mMsdb", , Language), 292, "Jingle Palette") = vbYes Then
3           butMenu.Caption = ""
4           p_Jingle(butMenuIdx).Visible = False
5           Call BASS_StreamFree(Jing(butMenuIdx).Strm)
6           Jing(butMenuIdx).Path = ""
7           Jing(butMenuIdx).Loop = False
8           Jing(butMenuIdx).Strm = 0
9           Jing(butMenuIdx).volume = 100
10          Jing(butMenuIdx).VuL = 0
11          Jing(butMenuIdx).VuR = 0
12      End If

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.mnuClr_Click"
    Resume Exit_Routine
End Sub
Private Sub mnuLoop_Click()
        On Error GoTo Error_Routine

1       Call clLoopSet(butMenuIdx, butMenu, True)

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.mnuLoop_Click"
    Resume Exit_Routine
End Sub

Private Sub mnuPause_Click()
        On Error GoTo Error_Routine

1       If Not BASS_ChannelIsActive(Jing(butMenuIdx).Strm) = BASS_ACTIVE_PAUSED Then
2           Call BASS_ChannelPause(Jing(butMenuIdx).Strm)
3           Jing(butMenuIdx).Paused = True
4           mnuPause.Caption = LangSet.Entry("mResume", , Language)
            'butMenu.FontItalic = True
5       ElseIf BASS_ChannelIsActive(Jing(butMenuIdx).Strm) = BASS_ACTIVE_PAUSED Then
6           Call clPlay(butMenuIdx, butMenu)
7           b_Jingle(butMenuIdx).value = vbChecked
8           Jing(butMenuIdx).Paused = False
9           mnuPause.Caption = LangSet.Entry("mPause", , Language)
            'butMenu.FontItalic = False
10      End If

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.mnuPause_Click"
    Resume Exit_Routine
End Sub
Private Sub mnuPlay_Click()
        On Error GoTo Error_Routine

1       If BASS_ChannelIsActive(Jing(butMenuIdx).Strm) = BASS_ACTIVE_PAUSED Then
2           Call BASS_ChannelStop(Jing(butMenuIdx).Strm)
3           Call BASS_ChannelSetPosition(Jing(butMenuIdx).Strm, 0)
4           mnuPause.Caption = LangSet.Entry("mPause", , Language)
5           butMenu.FontItalic = False
6       Else
7           Call clPlay(butMenuIdx, butMenu)
8           b_Jingle(butMenuIdx).value = vbChecked
9       End If

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.mnuPlay_Click"
    Resume Exit_Routine
End Sub

Private Sub mnuTa_Click()
        On Error GoTo Error_Routine

1       If butMenu.value = vbChecked Then
2           Call clTmAn(butMenuIdx, butMenu, True)
3           butMenu.value = vbChecked
4       Else
5           Call clTmAn(butMenuIdx, butMenu, True)
6       End If

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.mnuTa_Click"
    Resume Exit_Routine
End Sub

Private Sub mnuVol_Click()
        On Error GoTo Error_Routine

1       TabCt.Tab = 3
2       If butMenu.Caption <> "" Then Call clVolSel(butMenuIdx, butMenu, True)

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.mnuVol_Click"
    Resume Exit_Routine
End Sub

Private Sub SlideVolCh_Scroll(ByVal value As Integer)

        On Error GoTo Error_Routine

1       lblVol.Caption = value - 100 & " dB"
2       BASS_ChannelSetAttributes Jing(VolIndex).Strm, -1, value, -101
3       Jing(VolIndex).volume = value
4       VuVol = value
5       Exit Sub
6 err:

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.SlideVolCh_Scroll"
    Resume Exit_Routine
End Sub
Private Sub TabCt_Click(PreviousTab As Integer)
        On Error GoTo Error_Routine

1       If TabCt.Tab = 0 Then
2           lstPal.Visible = True
3       Else
4           lstPal.Visible = False
5       End If
6       If TabCt.Tab = 1 Then
7           ln1.Visible = True
8       Else
9           ln1.Visible = False
10      End If
11      If TabCt.Tab = 2 Then
12          ckMixByTa.Visible = True
13      Else
14          ckMixByTa.Visible = False
15      End If
16      If TabCt.Tab = 3 Then
17          lblVol.Visible = True
18          btSettings.Visible = True
19          btHelp.Visible = True
20          btAbout.Visible = True
21      Else
22          lblVol.Visible = False
23          btSettings.Visible = False
24          btHelp.Visible = False
25          btAbout.Visible = False
26      End If

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.TabCt_Click"
    Resume Exit_Routine
End Sub

Private Sub Timer1_Timer()
        ' FOR CONTROL FLASH
        On Error GoTo Error_Routine

1       If Timer2.Enabled = False Then
2           If Tme <> time Then
3               Timer2.Enabled = True
4               Timer1.Interval = 500
5           End If
6       End If

7       For i = 0 To 29
8           If BASS_ChannelIsActive(Jing(i).Strm) = BASS_ACTIVE_PAUSED Then
9               If tm1 = 0 Then
10                  tm1 = tm1 + 1
11                  b_Jingle(i).BackColor = &HC0FFC0
12              Else
13                  tm1 = 0
14                  b_Jingle(i).BackColor = &HC0FFFF
15              End If
16          End If
17      Next i

18      If mnuOn Then
19          If tmoutMnu = 2 Then
                'mnujp.MenuItems.Clear
20              mnuOn = False
                'MsgBox "menu?"
21          End If
22          tmoutMnu = tmoutMnu + 1
23      End If

24      If ckAssign.value = vbChecked Then
25          If tm1 = 0 Then
26              tm1 = tm1 + 1
27              ckAssign.BackColor = &HFF80FF
28          Else
29              tm1 = 0
30              ckAssign.BackColor = vbButtonFace
31          End If
32          If tmout = 10 Then
33              ckAssign.value = vbUnchecked
34          End If
35          tmout = tmout + 1
36      End If

37      If b_tmAnJin.value = vbChecked And b_tmAnJin.Caption = LangSet.Entry("mb_tmAnJin", , Language) Then
38          If tm1 = 0 Then
39              tm1 = tm1 + 1
40              b_tmAnJin.BackColor = &HFFC0C0
41          Else
42              tm1 = 0
43              b_tmAnJin.BackColor = vbButtonFace
44          End If
45          If tmout = 10 Then
46              b_tmAnJin.value = vbUnchecked
47          End If
48          tmout = tmout + 1

            'ElseIf b_tmAnJin.value = vbChecked And b_tmAnJin.Caption <> langset.entry("mb_tmAnJin",,language) Then
            '    b_tmAnJin.BackColor = &HFFC0C0
49      End If

50      If lblErr.Visible = True Then
51          If flash Then
52              If tmf = 0 Then
53                  tmf = tmf + 1
54                  lblErr.ForeColor = &H8080FF
55              Else
56                  tmf = 0
57                  lblErr.ForeColor = 0
58              End If
59          End If
60          If tmout = 10 Then
61              lblErr.Visible = False
62          End If
63          tmout = tmout + 1
64      End If

65      Timer1.Enabled = True

66      If txtSave.Visible = True Then
67          If tm1 = 10 And txtSave.Text = "" Then
68              txtSave.Visible = False
69              frmMain.KeyPreview = True
70          End If
71          If tm1 = 60 And txtSave.Text <> "" Then
72              txtSave.Visible = False
73              frmMain.KeyPreview = True
74          End If
75          tm1 = tm1 + 1
76      End If

77      If txtConDisp.Text <> LangSet.Entry("mStrNoCo", , Language) And ckPlayStream.value = vbUnchecked Then
78          If tm2 = 10 Then
79              txtConDisp.Text = LangSet.Entry("mStrNoCo", , Language)
80          End If
81          tm2 = tm2 + 1
82      End If

83      If txtConDisp.Text = LangSet.Entry("mStrPla", , Language) And ckPlayStream.value = vbUnchecked Then
84          txtConDisp.Text = LangSet.Entry("mStrNoCo", , Language)
85      End If

86      If Not ckAssVol.Visible And ManiPulVol Then   'btAbout.Caption = LangSet.Entry("mVolSave", , Language) Then
87          If tm1 = 0 Then
88              tm1 = tm1 + 1
89              btAbout.BackColor = &HFFFFC0   '&HFFFF80
90          Else
91              tm1 = 0
92              btAbout.BackColor = vbButtonFace
93          End If
94          If tmout = 30 Then
95              BtAbout_Click
96          End If
97          tmout = tmout + 1
98      End If

99      Exit Sub
100 err:

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.Timer1_Timer"
    Resume Exit_Routine
End Sub
Private Sub Timer2_Timer()
        ' FOR Time DISP AND TIME ANNOUNCE
        On Error GoTo Error_Routine

1       lblTime.Caption = Format(time, TimeForm)
2       lblDatex.Caption = Format(Date, DateForm)

3       If ckTmAnn Then
4           If Format(time, "hh:m") = Format(time, "hh:") & tmMin And Format(time, "s") = tmSec Then
5               If BASS_ChannelIsActive(LiveStream) <> 0 And ckMixByTa.value = vbChecked Then
6                   aMixCount = 0
7                   aMixStrm = LiveStream
8                   Timer4.Enabled = True
9                   Timer4.Interval = 40
10              End If
                tmJnGo = True
11              If BASS_ChannelPlay(tmSnStrm, BASSFALSE) = BASSFALSE Then
12                  Call Error_("Can't play time announce stream 1")
                    
13              End If
14              VuStrm = tmSnStrm
15          End If
        End If

16          If tmJnGo And b_tmAnJin And Format(time, "hh:m") = Format(time, "hh:") & tmMin And Format(time, "s") = tmSec + tmDel Then
17              aMixCount = 0
18              aMixStrm = tmSnStrm
19              Timer4.Enabled = True
20              Timer4.Interval = 10
21              If BASS_ChannelPlay(tmJnStrm, BASSFALSE) = BASSFALSE Then
22                  Call Error_("Can't play time announce stream 2")
23              End If
24              VuStrm = tmJnStrm
                tmJnGo = False
25          End If


        'If BASS_ChannelIsActive(tmSnStrm) = 0 And ckTmAnn Then
        '    ckTmAnn.Caption = "Time was announced"
        'End If

27      Exit Sub
28 err:

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.Timer2_Timer"
    Resume Exit_Routine
End Sub

Private Sub Timer3_Timer()
Dim VuCkL As Long
Dim VuCkR As Long

        ' FOR DISPLAYING
'        On Error GoTo Error_Routine

        On Error GoTo errh
1       Dim vnr As Integer
2       Dim vutm1L As Long, vutm1R As Long
3       Dim vutm2L As Long, vutm2R As Long
4       VuLft = 0
5       VuRght = 0
6       vnr = 0
7       For i = 0 To 29 'get the levels for each, icluding the volume level
8           If BASS_ChannelIsActive(Jing(i).Strm) = BASS_ACTIVE_PLAYING Then
9               Jing(i).VuL = LoWord(BASS_ChannelGetLevel(Jing(i).Strm)) * (Jing(i).volume / 100)
10              Jing(i).VuR = HiWord(BASS_ChannelGetLevel(Jing(i).Strm)) * (Jing(i).volume / 100)
11              vnr = vnr + 1
                'VuStrm = Jing(i).Strm
12          Else
13              Jing(i).VuL = 0
14              Jing(i).VuR = 0
15          End If
16      Next i

17      DoEvents
18      If vnr = 0 Then GoTo tmal

19      For i = 0 To 29 'get the highest one
20          If VuLft <= Jing(i).VuL Then
21              VuLft = Jing(i).VuL
22          End If
23          If VuRght <= Jing(i).VuR Then
24              VuRght = Jing(i).VuR
25          End If
26      Next i

27 tmal:
28      If BASS_ChannelIsActive(tmSnStrm) = BASS_ACTIVE_PLAYING Then
29          vutm1L = LoWord(BASS_ChannelGetLevel(tmSnStrm))
30          vutm1R = HiWord(BASS_ChannelGetLevel(tmSnStrm))
31          If VuLft <= vutm1L Then
32              VuLft = vutm1L
33          End If
34          If VuRght <= vutm1R Then
35              VuRght = vutm1R
36          End If
37      End If
38      If BASS_ChannelIsActive(tmJnStrm) = BASS_ACTIVE_PLAYING Then
39          vutm2L = LoWord(BASS_ChannelGetLevel(tmJnStrm))
40          vutm2R = HiWord(BASS_ChannelGetLevel(tmJnStrm))
41          If VuLft <= vutm2L Then
42              VuLft = vutm2L
43          End If
44          If VuRght <= vutm2R Then
45              VuRght = vutm2R
46          End If
47      End If
        If BASS_ChannelIsActive(LiveStream) = BASS_ACTIVE_PLAYING Then
            vutm2L = LoWord(BASS_ChannelGetLevel(LiveStream))
            vutm2R = HiWord(BASS_ChannelGetLevel(LiveStream))
            If VuLft <= vutm2L Then
                VuLft = vutm2L
            End If
            If VuRght <= vutm2R Then
                VuRght = vutm2R
            End If
        End If

48      VuL.Level = VuLft * (vol / 100) * 0.0030517578125 '(32767 - the max returned by the player dll, and
        VuR.Level = VuRght * (vol / 100) * 0.0030517578125 ' 100 - the max of the meter. This constant will bring things ok)

                'display in VU bar, including the general volume, with maximum check for meter
'49      If VuCkL <= 32767 Then
'            VuL.Level = VuCkL
'        Else
'            VuL.Level = 32767  'this is t
'        End If
'        If VuCkR <= 32767 Then
'            VuR.Level = VuCkR
'        Else
'            VuR.Level = 32767
'        End If
        

50      If VuStrm = 0 Or BASS_ChannelIsActive(VuStrm) <> BASS_ACTIVE_PLAYING Then Exit Sub

51      Sec = BASS_ChannelBytes2Seconds(VuStrm, BASS_StreamGetLength(VuStrm)) - BASS_ChannelBytes2Seconds(VuStrm, BASS_ChannelGetPosition(VuStrm))
52      min = Int(Sec / 60)

53      If min = 0 Then
54          lbl_gr.top = -150
55          lbl_gr.Caption = Chr(95)
56          lbl_gr.Alignment = 2
57      Else
58          lbl_gr.top = 0
59          lbl_gr.Caption = min & ":"
60          lbl_gr.Alignment = 1
61      End If
62      DoEvents
63      lblPosSec.Caption = Sec Mod 60 'Round((Sec), 0)
64      If Len(lblPosSec.Caption) = 1 Then lblPosSec.Caption = 0 & lblPosSec.Caption
65      lblPosMis.Caption = (Mid$(Round(Sec - Int(Sec), 2), 2, 3)) 'Mod 60

66      If Sec < RemWarn Then
67          ShRem.BorderColor = RemColor  '&H000000FF&  '>red
68          ShRem.Visible = True
69      Else
70          ShRem.BorderColor = &H0& 'black
71          ShRem.Visible = False
72      End If

73 errh:

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.Timer3_Timer"
    Resume Exit_Routine
End Sub

Private Sub Timer4_Timer()
        ' FOR AUTOMIXING
        '    aMixCount = aMixCount + 1
        '    BASS_ChannelGetAttributes aMixStrm, 0, VolMix, 0
        '    BASS_ChannelSetAttributes aMixStrm, -1, (VolMix - 10), -101
        '        If aMixCount = 10 Then
        '        Call BASS_ChannelStop(aMixStrm)
        '        Call BASS_ChannelSetPosition(aMixStrm, 0)
        ''        CountersReset aMixPlyr
        '        BASS_ChannelSetAttributes aMixStrm, -1, 100, -101
        ''        Jing(PlayEd(aMixPlyr).BtNr).OnAir = False
        ''        b_Jingle(PlayEd(aMixPlyr).BtNr).value = vbUnchecked
        '        Timer4.Enabled = False
        '        End If
        On Error GoTo Error_Routine

1       aMixCount = aMixCount + 1
2       For k = 0 To 29
3           If k <> aMixAir Then
4               BASS_ChannelGetAttributes Jing(k).Strm, 0, VolAmix(k), 0
5               BASS_ChannelSetAttributes Jing(k).Strm, -1, (VolAmix(k) - Jing(k).volume / 100), -101
                'If VolAmix(k) <= 0 Then
6               If aMixCount = 101 Then
7                   Call BASS_ChannelStop(Jing(k).Strm)
8                   Call BASS_ChannelSetPosition(Jing(k).Strm, 0)
9                   BASS_ChannelSetAttributes aMixStrm, -1, Jing(k).volume, -101
10                  Timer4.Enabled = False
11              End If
12          End If
13      Next k

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.Timer4_Timer"
    Resume Exit_Routine
End Sub

Private Sub Timer5_Timer()
        'for event retrieveing from bass
        On Error GoTo Error_Routine

1       For k = 0 To 29 'rise button when finished
2           If BASS_ChannelIsActive(Jing(k).Strm) = BASS_ACTIVE_STOPPED Then  'Or frmMain.CkTouch.value = vbChecked
3               b_Jingle(k).value = vbUnchecked
4               If Not ManiPulate And b_Jingle(k).BackColor <> vbButtonFace Then
5                   b_Jingle(k).BackColor = vbButtonFace
6               End If
7               Jing(k).OnAir = False
                j = 0
8               If Jing(k).inDebt Then
9                   Jing(k).inDebt = False
10                  Call openNames(k, lstPal.Text)
11                  Call openStreams(k)
                    'b_Jingle(k).BackColor = vbButtonFace
12              End If
            Else
                VuStrm = Jing(k).Strm
13          End If
14      Next k

15      If BASS_ChannelIsActive(VuStrm) = 0 Then
16          lblPosSec.Caption = "00"
17          lblPosMis.Caption = Right$(Format(0, "0.00"), 3)
18          ShRem.BorderColor = &H0&
19          ShRem.Visible = False
20          lbl_gr.Visible = False
21      Else
22          lbl_gr.Visible = True
23      End If

24      If BASS_ChannelIsActive(tmSnStrm) = 1 Then   'time announce rise button when finished
25          tmStat = True
26      End If
        'If BASS_ChannelIsActive(tmJnStrm) = 1 Then
        '    tmStat1 = True
        'End If
27      If tmStat And (BASS_ChannelIsActive(tmSnStrm) = 0 And BASS_ChannelIsActive(tmJnStrm) = 0) And tmJnGo = False Then
28          ckTmAnn.value = vbUnchecked
29          Call BASS_StreamFree(tmSnStrm)
30          Call BASS_StreamFree(tmJnStrm)
31      End If

32      If Not LiveOn Then 'BASS_ChannelIsActive(LiveStream) = 0 Then
33          ckPlayStream.value = vbUnchecked
34      End If

        'For i = 0 To 29
        '    If Not MouseIsOverObject(b_Jingle(i)) Then
        '        Call mnuCls_Click
        '    End If
        'Next i

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.Timer5_Timer"
    Resume Exit_Routine
End Sub
Private Sub TxtSave_KeyDown(KeyCode As Integer, Shift As Integer)

        On Error GoTo Error_Routine

1       tmout = 0
2       If KeyCode = vbKeyEscape Then
3           txtSave.Visible = False
4           frmMain.KeyPreview = True
5           Exit Sub
6       End If
7       If KeyCode = vbKeyReturn Then
8           If Not txtSave.Text = "" Then
9               Call CloseFile((Left$(StrConv(StrConv(txtSave.Text, vbLowerCase), vbProperCase), 15)))
10              txtSave.Visible = False
11              frmMain.KeyPreview = True
12              lstPal.SetFocus
13          End If
14      End If
15      Exit Sub
16 err:

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.TxtSave_KeyDown"
    Resume Exit_Routine
End Sub

Private Sub TxtSave_LostFocus()
        On Error GoTo Error_Routine

1       txtSave.Visible = False
2       frmMain.KeyPreview = True

Exit_Routine:
    Exit Sub
Error_Routine:
    Debug.Assert False
    ErrorLog "frmMain.TxtSave_LostFocus"
    Resume Exit_Routine
End Sub

