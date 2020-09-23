VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form NeoEphemerix_2001_Interface 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " NeoEphemerix 2001 v1 Beta 2 - Basic VSOP87 and FK5 Ephemeris Generator - NeoProgrammics"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "NeoEphemerix_2001_Interface.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   11910
   Begin VB.CheckBox Keep_On_Top_CheckBox 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Keep On Top "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   10080
      TabIndex        =   42
      TabStop         =   0   'False
      ToolTipText     =   " Check This Box to Keep This Window on Top of Others "
      Top             =   720
      Value           =   1  'Checked
      Width           =   1770
   End
   Begin VB.CommandButton Save_Work_Button 
      BackColor       =   &H00C0C0C0&
      Caption         =   "SAVE Work"
      Height          =   330
      Left            =   10035
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   7515
      Width           =   1050
   End
   Begin VB.CommandButton INFO_Button 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Info"
      Height          =   330
      Left            =   11385
      TabIndex        =   39
      TabStop         =   0   'False
      ToolTipText     =   " Some General Program Info "
      Top             =   7515
      Width           =   510
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H00C0C0C0&
      Caption         =   " ± TZ Adjust"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   600
      Left            =   2160
      TabIndex        =   38
      ToolTipText     =   " Time Zone Adjustment  (Local Time  +  TZ Adjust  =  UT) "
      Top             =   7965
      Width           =   1410
      Begin VB.TextBox The_TZ_Adjustment 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   90
         MaxLength       =   9
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "00:00:00"
         Top             =   225
         Width           =   1230
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Computation Mode"
      ForeColor       =   &H00000000&
      Height          =   1770
      Left            =   10035
      TabIndex        =   37
      ToolTipText     =   " Selects the Type of Computation Desired "
      Top             =   1080
      Width           =   1860
      Begin VB.OptionButton COMP_ALL_Mode 
         BackColor       =   &H00C0C0C0&
         Caption         =   " All Objects "
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   45
         TabIndex        =   12
         ToolTipText     =   " Compute Quick Ephemeris for All Objects "
         Top             =   1440
         Width           =   1770
      End
      Begin VB.OptionButton COMP_STATS_Mode 
         BackColor       =   &H00C0C0C0&
         Caption         =   " Set Date / Stats "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   45
         TabIndex        =   11
         ToolTipText     =   " Compute Basic Statistics for Current Interface Date/Time Settings "
         Top             =   1170
         Value           =   -1  'True
         Width           =   1770
      End
      Begin VB.OptionButton COMP_EQ_Mode 
         BackColor       =   &H00C0C0C0&
         Caption         =   " Geo Equatorial "
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   45
         TabIndex        =   10
         ToolTipText     =   " Compute Table of Apparent Equatorial FK5 Positions of Object for Interface Month "
         Top             =   900
         Width           =   1770
      End
      Begin VB.OptionButton COMP_EC_Mode 
         BackColor       =   &H00C0C0C0&
         Caption         =   " Geo Ecliptical "
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   45
         TabIndex        =   9
         ToolTipText     =   " Compute Table of Apparent Ecliptical FK5 Positions of Object for Interface Month "
         Top             =   585
         Width           =   1770
      End
      Begin VB.OptionButton COMP_HC_Mode 
         BackColor       =   &H00C0C0C0&
         Caption         =   " Heliocentric "
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   45
         TabIndex        =   8
         ToolTipText     =   " Compute VSOP87 Heliocentric Ephemeris of Object for Interface Month "
         Top             =   270
         Width           =   1770
      End
   End
   Begin MSComDlg.CommonDialog SAVE_Dialog 
      Left            =   9495
      Top             =   7380
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   ".txt"
      DialogTitle     =   " Ephemerix 2001 - SAVE Computations as a Text File"
      Filter          =   "Text File|*.txt"
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Longitude Mode"
      ForeColor       =   &H00000000&
      Height          =   915
      Left            =   10035
      TabIndex        =   34
      Top             =   4365
      Width           =   1860
      Begin VB.OptionButton Lng_DMS_Mode 
         BackColor       =   &H00C0C0C0&
         Caption         =   " °     '     "" "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   45
         TabIndex        =   17
         Top             =   270
         Value           =   -1  'True
         Width           =   1770
      End
      Begin VB.OptionButton Lng_DD_Mode 
         BackColor       =   &H00C0C0C0&
         Caption         =   " Deg.ddddd "
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   45
         TabIndex        =   18
         Top             =   585
         Width           =   1770
      End
   End
   Begin VB.Frame Dist_Units_Frame 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Distance Units"
      ForeColor       =   &H00000000&
      Height          =   1185
      Left            =   10035
      TabIndex        =   33
      Top             =   6255
      Width           =   1860
      Begin VB.OptionButton Dist_Units_MI_Mode 
         BackColor       =   &H00C0C0C0&
         Caption         =   " Statute Miles "
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   45
         TabIndex        =   23
         Top             =   855
         Width           =   1770
      End
      Begin VB.OptionButton Dist_Units_KM_Mode 
         BackColor       =   &H00C0C0C0&
         Caption         =   " Kilometers "
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   45
         TabIndex        =   22
         Top             =   585
         Width           =   1770
      End
      Begin VB.OptionButton Dist_Units_AU_Mode 
         BackColor       =   &H00C0C0C0&
         Caption         =   " Astronom. Units "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   45
         TabIndex        =   21
         Top             =   270
         Value           =   -1  'True
         Width           =   1770
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Decl / Lat / Eclip Mode"
      ForeColor       =   &H00000000&
      Height          =   915
      Left            =   10035
      TabIndex        =   32
      Top             =   5310
      Width           =   1860
      Begin VB.OptionButton Decl_Lat_DD_Mode 
         BackColor       =   &H00C0C0C0&
         Caption         =   " Deg.ddddd "
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   45
         TabIndex        =   20
         Top             =   585
         Width           =   1770
      End
      Begin VB.OptionButton Decl_Lat_DMS_Mode 
         BackColor       =   &H00C0C0C0&
         Caption         =   " °     '     "" "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   45
         TabIndex        =   19
         Top             =   270
         Value           =   -1  'True
         Width           =   1770
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "RA / Hour Angle Mode"
      ForeColor       =   &H00000000&
      Height          =   1455
      Left            =   10035
      TabIndex        =   31
      Top             =   2880
      Width           =   1860
      Begin VB.OptionButton HA_DD_Mode 
         BackColor       =   &H00C0C0C0&
         Caption         =   " Deg.ddddd "
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   45
         TabIndex        =   16
         Top             =   1125
         Width           =   1770
      End
      Begin VB.OptionButton HA_DMS_Mode 
         BackColor       =   &H00C0C0C0&
         Caption         =   " °     '     "" "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   45
         TabIndex        =   15
         Top             =   810
         Width           =   1770
      End
      Begin VB.OptionButton HA_DH_Mode 
         BackColor       =   &H00C0C0C0&
         Caption         =   " HH.hhhhh "
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   45
         TabIndex        =   14
         Top             =   540
         Width           =   1770
      End
      Begin VB.OptionButton HA_HMS_Mode 
         BackColor       =   &H00C0C0C0&
         Caption         =   " HH MM SS.sss "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   45
         TabIndex        =   13
         Top             =   270
         Value           =   -1  'True
         Width           =   1770
      End
   End
   Begin VB.ListBox Work 
      BackColor       =   &H00FFFFFF&
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   7275
      ItemData        =   "NeoEphemerix_2001_Interface.frx":0442
      Left            =   0
      List            =   "NeoEphemerix_2001_Interface.frx":0444
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   630
      Width           =   10005
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Object"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   600
      Left            =   5535
      TabIndex        =   29
      ToolTipText     =   " Selects Ephemeris Object for Computations "
      Top             =   0
      Width           =   1320
      Begin VB.ComboBox The_Object 
         BackColor       =   &H00FFFF00&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "NeoEphemerix_2001_Interface.frx":0446
         Left            =   90
         List            =   "NeoEphemerix_2001_Interface.frx":0448
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   225
         Width           =   1140
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Universal Time"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   600
      Left            =   2340
      TabIndex        =   28
      Top             =   0
      Width           =   1500
      Begin VB.TextBox The_Time 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   90
         MaxLength       =   8
         TabIndex        =   3
         Text            =   "00:00:00"
         Top             =   225
         Width           =   1320
      End
   End
   Begin VB.TextBox The_Object_Index 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   900
      TabIndex        =   26
      TabStop         =   0   'False
      Text            =   "Obj #"
      Top             =   8145
      Width           =   735
   End
   Begin VB.TextBox Program_Initialized 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   0
      TabIndex        =   25
      TabStop         =   0   'False
      Text            =   "Init ?"
      Top             =   8145
      Width           =   870
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Calendar Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   600
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   2310
      Begin VB.TextBox The_Year 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1170
         MaxLength       =   7
         TabIndex        =   2
         Text            =   "2001 AD"
         Top             =   225
         Width           =   1050
      End
      Begin VB.TextBox The_Month 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   540
         MaxLength       =   3
         TabIndex        =   1
         Text            =   "Jan"
         Top             =   225
         Width           =   600
      End
      Begin VB.TextBox The_Day 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   90
         MaxLength       =   2
         TabIndex        =   0
         Text            =   "1"
         Top             =   225
         Width           =   420
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "COMPUTE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   600
      Left            =   6885
      TabIndex        =   35
      ToolTipText     =   " Perform the Computations Indicated by Interface Settings "
      Top             =   0
      Width           =   1095
      Begin VB.CommandButton COMPUTE_Button 
         BackColor       =   &H00C0C0C0&
         Default         =   -1  'True
         Height          =   330
         Left            =   90
         TabIndex        =   7
         Top             =   225
         Width           =   915
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0C0C0&
      Caption         =   " ± Delta  T"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   600
      Left            =   3870
      TabIndex        =   36
      ToolTipText     =   " Delta T        (Dynamical Time  =  UT + Delta T) "
      Top             =   0
      Width           =   1635
      Begin VB.TextBox The_Delta_T 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   90
         MaxLength       =   9
         TabIndex        =   5
         Text            =   "00:00:00"
         Top             =   225
         Width           =   1455
      End
   End
   Begin VB.Label Message 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Left            =   8010
      TabIndex        =   41
      Top             =   90
      Visible         =   0   'False
      Width           =   3885
   End
   Begin VB.Label Label0 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Special Control Values"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   45
      TabIndex        =   27
      Top             =   7920
      Width           =   2085
   End
End
Attribute VB_Name = "NeoEphemerix_2001_Interface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Option Explicit

' NeoEphemerix 2001 - v1.101 Beta 2
' =========================================================================
' =========================================================================
' =========================================================================

' The SUBs in this section handle the interface.
'
' The main program built to use the interface is written
' following this section.

  Private Sub Form_Load()
' What to do when this program starts up

  LOCK_ON_TOP

  RECALL_PROGRAM_SETTINGS

  If Not INITIALIZED Then GoSub INITIALIZATION_ROUTINE

  The_Object.ListIndex = Val(The_Object_Index)
 
  Exit Sub

' This program code is executed ONLY if not initialized yet.
' This occurs the first time the program has been run so it
' can set up some initial Windows registry values first.
INITIALIZATION_ROUTINE:

' Define the objects list
  With The_Object
  .AddItem " Sun"
  .AddItem " Mercury"
  .AddItem " Venus"
  .AddItem " Earth"
  .AddItem " Mars"
  .AddItem " Jupiter"
  .AddItem " Saturn"
  .AddItem " Uranus"
  .AddItem " Neptune"
  End With

' Define universal time default
  The_Time = "00:00:00"

' Define Delta T default
  The_Delta_T = "00:00:00"

' Set initialization flag
  Program_Initialized = "YES"

' Store initial defaults setting in the Windows registry
  STORE_PROGRAM_SETTINGS
  DISPLAY_INFO
  Return

  End Sub

  Private Sub Form_Click()
' What to do when a blank area of the main form is clicked on
  Work.ListIndex = -1
  COMPUTE_Button.SetFocus
  End Sub

  Private Sub Form_Unload(Cancel As Integer)
' What to do when this form is unloaded
  Work.Clear
  STORE_PROGRAM_SETTINGS
  End Sub

  Private Sub Form_Terminate()
' What to do when this program is terminated
  Unload Me
  End Sub

' Lock this form on top of all other forms
' or restore to unlocked mode as needed

  Private Sub Keep_On_Top_CheckBox_Click()
  If Keep_On_Top_CheckBox.Value = 1 Then
     LOCK_ON_TOP
     Keep_On_Top_CheckBox.FontBold = True
  Else
     UNLOCK_ON_TOP
     Keep_On_Top_CheckBox.FontBold = False
  End If
  End Sub

' Check to see if program was initialized
  Private Function INITIALIZED() As Boolean
  If Program_Initialized = "YES" Then INITIALIZED = True _
     Else INITIALIZED = False
  End Function


  Private Sub The_Object_Click()
  The_Object_Index = The_Object.ListIndex
  Message.Visible = False
  HM_TABLE_MODE = ""
  STORE_PROGRAM_SETTINGS
  End Sub

  Private Sub Work_LostFocus()
  Work.ListIndex = -1
  COMPUTE_Button.SetFocus
  End Sub

  Private Sub COMP_HC_Mode_Click()
  Message.Visible = False
  HM_TABLE_MODE = ""
  COMP_MODE = "HC"
  The_Day.Enabled = False
  The_Object.Enabled = True
  ENABLE_DISTANCE_UNITS (True)
  ENABLE_HA_MODE (False)
  UNHIGHLIGHT_COMP_MODE_TEXT
  COMP_HC_Mode.FontBold = True
  End Sub
  Private Sub COMP_EC_Mode_Click()
  Message.Visible = False
  HM_TABLE_MODE = ""
  COMP_MODE = "EC"
  The_Day.Enabled = False
  The_Object.Enabled = True
  ENABLE_DISTANCE_UNITS (True)
  ENABLE_HA_MODE (True)
  UNHIGHLIGHT_COMP_MODE_TEXT
  COMP_EC_Mode.FontBold = True
  End Sub
  Private Sub COMP_EQ_Mode_Click()
  Message.Visible = False
  HM_TABLE_MODE = ""
  COMP_MODE = "EQ"
  The_Day.Enabled = False
  The_Object.Enabled = True
  ENABLE_DISTANCE_UNITS (True)
  ENABLE_HA_MODE (True)
  UNHIGHLIGHT_COMP_MODE_TEXT
  COMP_EQ_Mode.FontBold = True
  End Sub
  Private Sub COMP_ALL_Mode_Click()
  Message.Visible = False
  HM_TABLE_MODE = ""
  COMP_MODE = "ALL"
  UNHIGHLIGHT_COMP_MODE_TEXT
  COMP_ALL_Mode.FontBold = True
  ENABLE_DISTANCE_UNITS (True)
  The_Object.Enabled = False
  The_Day.Enabled = True
  End Sub
  Private Sub COMP_STATS_Mode_Click()
  Message.Visible = False
  HM_TABLE_MODE = ""
  COMP_MODE = "STATS"
  The_Day.Enabled = True
  The_Object.Enabled = False
  ENABLE_DISTANCE_UNITS (False)
  UNHIGHLIGHT_COMP_MODE_TEXT
  COMP_STATS_Mode.FontBold = True
  End Sub

' Unhighlight all computation mode option buttons text
  Private Sub UNHIGHLIGHT_COMP_MODE_TEXT()
  COMP_HC_Mode.FontBold = False
  COMP_EC_Mode.FontBold = False
  COMP_EQ_Mode.FontBold = False
  COMP_ALL_Mode.FontBold = False
  COMP_STATS_Mode.FontBold = False
  End Sub

' Unhighlight all hour angle option buttons text
  Private Sub UNHIGHLIGHT_HA_MODE_TEXT()
  HA_HMS_Mode.FontBold = False
  HA_DH_Mode.FontBold = False
  HA_DMS_Mode.FontBold = False
  HA_DD_Mode.FontBold = False
  End Sub

' Unhighlight all longitude angle mode option buttons text
  Private Sub UNHIGHLIGHT_LNG_MODE_TEXT()
  Lng_DMS_Mode.FontBold = False
  Lng_DD_Mode.FontBold = False
  End Sub

' Unhighlight all declination/latitude angle mode option buttons text
  Private Sub UNHIGHLIGHT_DECL_LAT_MODE_TEXT()
  Decl_Lat_DMS_Mode.FontBold = False
  Decl_Lat_DD_Mode.FontBold = False
  End Sub

' Unhighlight all distance units mode option buttons text
  Private Sub UNHIGHLIGHT_DIST_UNITS_MODE_TEXT()
  Dist_Units_AU_Mode.FontBold = False
  Dist_Units_KM_Mode.FontBold = False
  Dist_Units_MI_Mode.FontBold = False
  End Sub

  Private Sub HA_HMS_Mode_Click()
  HA_MODE = "HMS"
  UNHIGHLIGHT_HA_MODE_TEXT
  HA_HMS_Mode.FontBold = True
  End Sub
  Private Sub HA_DH_Mode_Click()
  HA_MODE = "DH"
  UNHIGHLIGHT_HA_MODE_TEXT
  HA_DH_Mode.FontBold = True
  End Sub
  Private Sub HA_DMS_Mode_Click()
  HA_MODE = "DMS"
  UNHIGHLIGHT_HA_MODE_TEXT
  HA_DMS_Mode.FontBold = True
  End Sub
  Private Sub HA_DD_Mode_Click()
  HA_MODE = "DD"
  UNHIGHLIGHT_HA_MODE_TEXT
  HA_DD_Mode.FontBold = True
  End Sub
  Private Sub Lng_DMS_Mode_Click()
  LNG_MODE = "DMS"
  UNHIGHLIGHT_LNG_MODE_TEXT
  Lng_DMS_Mode.FontBold = True
  End Sub
  Private Sub Lng_DD_Mode_Click()
  LNG_MODE = "DD"
  UNHIGHLIGHT_LNG_MODE_TEXT
  Lng_DD_Mode.FontBold = True
  End Sub
  Private Sub Decl_Lat_DMS_Mode_Click()
  DECL_LAT_MODE = "DMS"
  UNHIGHLIGHT_DECL_LAT_MODE_TEXT
  Decl_Lat_DMS_Mode.FontBold = True
  End Sub
  Private Sub Decl_Lat_DD_Mode_Click()
  DECL_LAT_MODE = "DD"
  UNHIGHLIGHT_DECL_LAT_MODE_TEXT
  Decl_Lat_DD_Mode.FontBold = True
  End Sub
  Private Sub Dist_Units_AU_Mode_Click()
  DIST_UNITS = "AU"
  UNHIGHLIGHT_DIST_UNITS_MODE_TEXT
  Dist_Units_AU_Mode.FontBold = True
  End Sub
  Private Sub Dist_Units_km_Mode_Click()
  DIST_UNITS = "KM"
  UNHIGHLIGHT_DIST_UNITS_MODE_TEXT
  Dist_Units_KM_Mode.FontBold = True
  End Sub
  Private Sub Dist_Units_mi_Mode_Click()
  DIST_UNITS = "MI"
  UNHIGHLIGHT_DIST_UNITS_MODE_TEXT
  Dist_Units_MI_Mode.FontBold = True
  End Sub

' Date and time focus highlighting routines
  Private Sub The_Day_GotFocus()
  The_Day.SelStart = 0
  The_Day.SelLength = Len(The_Day.Text)
  End Sub
  Private Sub The_Month_GotFocus()
  The_Month.SelStart = 0
  The_Month.SelLength = Len(The_Month.Text)
  End Sub
  Private Sub The_Year_GotFocus()
  The_Year.SelStart = 0
  The_Year.SelLength = Len(The_Year.Text)
  End Sub
  Private Sub The_Time_GotFocus()
  The_Time.SelStart = 0
  The_Time.SelLength = Len(The_Time.Text)
  End Sub
  Private Sub The_TZ_Adjustment_GotFocus()
  The_TZ_Adjustment.SelStart = 0
  The_TZ_Adjustment.SelLength = Len(The_TZ_Adjustment.Text)
  End Sub


  Private Sub The_Delta_T_GotFocus()
  The_Delta_T.SelStart = 0
  The_Delta_T.SelLength = Len(The_Delta_T.Text)
  End Sub

' Enable or disable distance units display
  Private Sub ENABLE_DISTANCE_UNITS(True_False As Boolean)
  If True_False = True Then
     Dist_Units_AU_Mode.Enabled = True
     Dist_Units_KM_Mode.Enabled = True
     Dist_Units_MI_Mode.Enabled = True
  Else
     Dist_Units_AU_Mode.Enabled = False
     Dist_Units_KM_Mode.Enabled = False
     Dist_Units_MI_Mode.Enabled = False
  End If
  End Sub

' Enable or disable the hour angle mode options
  Private Sub ENABLE_HA_MODE(True_False As Boolean)
  If True_False = True Then
  HA_HMS_Mode.Enabled = True
  HA_DH_Mode.Enabled = True
  HA_DMS_Mode.Enabled = True
  HA_DD_Mode.Enabled = True
  Else
  HA_HMS_Mode.Enabled = False
  HA_DH_Mode.Enabled = False
  HA_DMS_Mode.Enabled = False
  HA_DD_Mode.Enabled = False
  End If
  End Sub



' Read and process interface global settings and
' check for and report erroneous values.
'
' This routine store any error messages in global DATA_ERROR.
' and the COMPUTE button code reads this error status prior
' to performing computations to prevent crashing errors.

  Private Sub READ_INTERFACE()

  Dim Q As Variant
  Dim U As Variant
  Dim W As Variant

  Work.Clear

  DATA_ERROR = "" ' Initialize error status to null

' Read the date values
  The_Date = Trim(The_Day) & " " & Trim(The_Month) _
  & " " & Trim(The_Year)

  Q = JDE_For(The_Date, 0)

  If Error_In(Q) Then DATA_ERROR = Q: Exit Sub

' If date is OK, then reconstruct it
  Q = Calendar_Date_For(Q)
  The_Day = Trim(Val(Q)): U = InStr(Q, " ") + 1
  The_Month = Mid(Q, U, 3): U = U + 3
  The_Year = Trim(Mid(Q, U, Len(Q)))
  The_Date = The_Day & " " & The_Month & " " & The_Year

' Read the time setting - adjust if necessary
  Q = 86400 * (Day_Frac_Equiv_To(The_Time))
  The_Time = Time_Equiv_To(Q)
  If Val(The_Time) > 24 Then
  DATA_ERROR = "ERROR: Invalid local time"
  Exit Sub
  End If

' Read the time zone adjustment.  This value is added to
' the LMT to get the UT
  Q = Day_Frac_Equiv_To(The_TZ_Adjustment)
  U = 86400 * Q
  Q = Time_Equiv_To(U)
  The_TZ_Adjustment = Left(Q, Len(Q) - 3)
  If Abs(Val(The_TZ_Adjustment)) > 24 Then
     DATA_ERROR = "ERROR: " & The_TZ_Adjustment _
    & " = Invalid Time Zone Adjustment "
     Exit Sub
  End If

' Read the delta T setting - adjust if necessary.
' The dT is rounded to the nearest second.
' This value is SUBTRACTED from the TD to get the
' corresponding UT or added to the UT to get the TD.
  The_dT = Day_Frac_Equiv_To(The_Delta_T)
  Q = 86400 * The_dT
  The_Delta_T = Time_Equiv_To(Q)
  If Abs(Val(The_Delta_T)) > 24 Then
     DATA_ERROR = "ERROR: " & The_Delta_T & " = Invalid Delta T "
     Exit Sub
  Else
     The_dT = Day_Frac_Equiv_To(The_Delta_T) ' Reavaluate dT
  End If

' Now compute the complete JDE dynamical value
' accounting for the Date, LMT, TZ Adjustment, Delta T
  Q = Day_Frac_Equiv_To(The_Time) _
    + Day_Frac_Equiv_To(The_TZ_Adjustment)
  The_JDE = JDE_For(The_Date, Q) + The_dT

' Read computation mode
  If COMP_HC_Mode.Value = True Then COMP_MODE = "HC"
  If COMP_EC_Mode.Value = True Then COMP_MODE = "EC"
  If COMP_EQ_Mode.Value = True Then COMP_MODE = "EQ"
    If COMP_ALL_Mode.Value = True Then COMP_MODE = "ALL"
  If COMP_STATS_Mode.Value = True Then COMP_MODE = "STATS"

' Read hour angle display mode
  If HA_HMS_Mode.Value = True Then HA_MODE = "HMS"
  If HA_DH_Mode.Value = True Then HA_MODE = "DH"
  If HA_DMS_Mode.Value = True Then HA_MODE = "DMS"
  If HA_DD_Mode.Value = True Then HA_MODE = "DD"

' Read longitude angle display mode
  If Lng_DMS_Mode.Value = True Then LNG_MODE = "DMS"
  If Lng_DD_Mode.Value = True Then LNG_MODE = "DD"

' Read declination/latitude display mode
  If Decl_Lat_DMS_Mode.Value = True Then DECL_LAT_MODE = "DMS"
  If Decl_Lat_DD_Mode.Value = True Then DECL_LAT_MODE = "DD"

' Read distance units display mode
  If Dist_Units_AU_Mode.Value = True Then DIST_UNITS = "AU"
  If Dist_Units_KM_Mode.Value = True Then DIST_UNITS = "KM"
  If Dist_Units_MI_Mode.Value = True Then DIST_UNITS = "MI"

' Save adjusted interface settings
  STORE_PROGRAM_SETTINGS

  End Sub

' Save current contents of work display to a text file
  Private Sub Save_Work_Button_Click()

  On Error GoTo ERROR_HANDLER

  Dim File_Name As String  ' Name of file to save
  Dim i         As Integer ' Loop control index
  Dim Q As String
   
  ChDrive (App.Path)
  ChDir (App.Path)
  
  Work.Enabled = True
      
  If Work.ListCount = 0 Then
     Q = MsgBox("The computations work area is blank." & vbCrLf _
       & "There is nothing to be saved.", vbExclamation, " NeoEphemerix 2001")
         COMPUTE_Button.SetFocus
         Exit Sub
  End If

  SAVE_Dialog.FileName = "Work - " _
  & Format(Now, "Yyyy Mmm Dd HHMMSS")
  
  File_Name = ""
  SAVE_Dialog.ShowSave

  File_Name = SAVE_Dialog.FileTitle
  If File_Name = "" Then
     COMPUTE_Button.SetFocus
     Exit Sub
  End If

  Open File_Name For Output As #1
  For i = 0 To Work.ListCount - 1
      Print #1, Work.List(i)
  Next i
  Print #1, " "

  Print #1, "Generated by NeoEphemerix 2001 - v1.0 Beta"
  Close 1

  Q = MsgBox("The work area has been saved as" & vbCrLf & vbCrLf _
    & """" & File_Name & """", vbInformation, " NeoEphemerix 2001")
  COMPUTE_Button.SetFocus
  Exit Sub

' Exit if CANCEL generates an error
ERROR_HANDLER:
  COMPUTE_Button.SetFocus
  Exit Sub

  End Sub

' Special meta-command to insert a blank line in Work display
' simulating a dual carriage return and line feed in normal text
' by printing a line with a single space character.
  Private Sub BL()
  Work.AddItem " "
  End Sub

' Special output meta-command to output a string to a Work line
' simulating a PRINT command.
  Private Sub OUT(Q)
  Work.AddItem Q
  End Sub

' Perform the computations indicated by the interface settings when
' the COMPUTE button is clicked on.

  Private Sub Compute_Button_Click()

  Dim Q   As Variant
  Dim LBR As String
  Dim Day As Integer

  COMPUTE_Button.SetFocus

  Message.Visible = False
  
  READ_INTERFACE

  If DATA_ERROR <> "" Then
     OUT " "
     OUT " " & DATA_ERROR
     Beep
     Exit Sub
  End If

  NeoEphemerix_2001_Interface.MousePointer = vbHourglass

  Message.Visible = False
  HM_TABLE_MODE = ""
  STORE_PROGRAM_SETTINGS
  OUT COMP_MODE

  If COMP_MODE = "HC" Then COMPUTE_HELIOCENTRIC
  If COMP_MODE = "EC" Or COMP_MODE = "EQ" Then COMPUTE_GEOCENTRIC
  If COMP_MODE = "STATS" Then COMPUTE_DATE_STATS
  If COMP_MODE = "ALL" Then COMPUTE_ALL_OBJECTS

  NeoEphemerix_2001_Interface.MousePointer = vbDefault

  End Sub

  Private Sub INFO_Button_Click()
' What to do when the INFO button is clicked on
  DISPLAY_INFO
  COMPUTE_Button.SetFocus

  End Sub

' Display some general program info
  Private Sub DISPLAY_INFO()

  Message.Visible = False
  HM_TABLE_MODE = ""

  Work.Clear
  BL
  OUT " NeoEphemerix 2001  v1.0 Beta version - NeoProgrammics"
  OUT " Jay Tanner"
  BL
  OUT " This program was designed to act as a quick reference to where the major planets"
  OUT " are, were or will be at any given moment."
  BL
  OUT " It is not a graphical program.  The primary emphasis is on numerical accuracy so"
  OUT " that the computed results may be used for other computations beyond the current"
  OUT " scope of the program, such as some form of phenomenon that may be predicted in"
  OUT " turn by using the values computed by this program."
  BL
  OUT " It implements the full VSOP87 theory of Pierre Bretagnon of the Bureau des"
  OUT " Longitudes in Paris.  As a result, it yields computations that compare favorably"
  OUT " with those of published astronomical almanacs."
  BL
  OUT " This beta version is the first fully functional test of the computation modules"
  OUT " that went into its design.  Future development is planned depending on the"
  OUT " results of testing the program over time."
  BL
  OUT " The program computes general ephemerides of the eight major planets from Mercury"
  OUT " to Neptune and displays the results in several different modes for convenience."
  BL
  OUT " Ephemeris tables can be generated in day, hour or minute intervals and also be"
  OUT " saved to disk as text files."
  BL
  OUT " Ephemerides include:"
  OUT " Heliocentric VSOP87 spherical coordinates: Longitude, Latitude and Radius Vector"
  OUT " Apparent geocentric ecliptical coordinates (FK5)"
  OUT " Apparent geocentric equatorial (FK5) Right Ascension, Declination and Distance"
  BL
  OUT " Corrections include those for light time, aberration, nutation and reduction to"
  OUT " the FK5 system of coordinates."
  BL
  OUT " Suggestions, bug reports and comments on this program are encouraged and may be"
  OUT " directed to: Jay@NeoProgrammics.com"
  
  Work.ToolTipText = ""
  End Sub

' -------------------------------------------------------------------------
' Unhighlight any selected work area line when pointer moves outside it

  Private Sub Work_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

  If x >= 9900 Or x = 0 Then Work.ListIndex = -1: COMPUTE_Button.SetFocus
  If y >= 7200 Or y <= 15 Then Work.ListIndex = -1: COMPUTE_Button.SetFocus
  End Sub

' =========================================================================
' =========================================================================
' =========================================================================
' THE MAIN NEOEPHEMERIX 2001 PROGRAM USING THE INTERFACE FOLLOWS HERE



' Compute apparent geocentric coordinates for selected object for current
' interface month at the given time on each date.


' ------------------------------------------------------------------------
' Compute the VSOP87 heliocentric L,B,R coordinates for indicated object
' at given dynamical JDE value.
    
  Private Sub COMPUTE_HELIOCENTRIC()

  Dim Q    As Variant

  Dim At_JDE  As String
  Dim Day     As Integer
  Dim D       As String

  Dim Date_String As String

  Work.Clear

  If Trim(The_Object) = "Sun" Then
  BL
  OUT " " & "ALERT: Heliocentric computations do not apply to the Sun itself."
  Beep
  Exit Sub
  End If

  OUT "VSOP87 heliocentric ephemeris of" & The_Object _
  & " for the month of " & The_Month & " " & The_Year
  OUT "at " & The_Time & " UT  on each date.     Delta T = " & The_Delta_T _
  & "   * = Millions"
 
  BL
  OUT "Day|   H Longitude   |    H Latitude   |     H Distance"

  For Day = 1 To 31

  D = Trim(Str(Day)) & " ": If Day < 10 Then D = " " & D
  
  Date_String = D & The_Month & The_Year

  At_JDE = JDE_For(Date_String, 0)
  If Error_In(At_JDE) Then Exit For

' Compute the dynamical JDE value
  At_JDE = At_JDE + Day_Frac_Equiv_To(The_Time) _
      + Day_Frac_Equiv_To(The_TZ_Adjustment) + The_dT

' Compute the L,B,R values
  Q = Heliocentric_Position_Of(The_Object, At_JDE)

  OUT D & "|" & Q
  
  Next Day

  HM_TABLE_MODE = "H"
  Message.Visible = True
  Message.Caption = " Double-click date in the ephemeris for a 24 hour ephemeris for that date. "

  End Sub


' ------------------------------------------------------------------------
' Compute the apparent geocentric FK5 coordinates for indicated object
' at given dynamical JDE value.


  Private Sub COMPUTE_GEOCENTRIC()

' Work variables
  Dim Q    As Variant
  Dim W    As Variant

' The full dynamical Julian day number used for the computations
  Dim At_JDE  As String

' Table generator day counter and string copy
  Dim Day  As Integer
  Dim D    As String

' Working date string
  Dim Date_String As String

  Work.Clear

  If Trim(The_Object) = "Earth" Then
  BL
  OUT " " & "ALERT: Geocentric computations do not apply to the Earth itself."
  Beep
  Exit Sub
  End If

  Q = " ecliptical ": If COMP_MODE = "EQ" Then Q = " equatorial "

  OUT "Apparent geocentric" & Q & "FK5 ephemeris of" & The_Object & " for the month of "
  OUT The_Month & " " & The_Year & "  at " _
  & The_Time & " UT on each date.    Delta T = " & The_Delta_T & "   * = Millions"
  
  BL
  If COMP_MODE = "EQ" Then
  OUT "Day|       RA        |       Decl      |     Distance    |  Ang Diam  | VMag"
  Else
  OUT "Day|    Longitude    |     Latitude    |     Distance    |  Ang Diam  | VMag"
  End If

  For Day = 1 To 31

  D = Trim(Str(Day)) & " "
  If Day < 10 Then D = " " & D
  
  Date_String = D & The_Month & The_Year

  At_JDE = JDE_For(Date_String, 0)
  If Error_In(At_JDE) Then Exit For

' Compute the dynamical JDE value
  At_JDE = At_JDE + Day_Frac_Equiv_To(The_Time) _
      + Day_Frac_Equiv_To(The_TZ_Adjustment) + The_dT

' Compute the heliocentric L,B,R coordinates for Earth
  Q = Geocentric_Position_Of(The_Object, At_JDE)

  OUT D & "|" & Q

  Next Day

  HM_TABLE_MODE = "H"
  Message.Visible = True
  Message.Caption = " Double-click date line in the ephemeris for a 24 hour ephemeris for that date. "

  End Sub

' Compute general statistics for interface date/time settings

  Private Sub COMPUTE_DATE_STATS()

  Dim Q As Variant
  Dim U As Variant
  Dim V As Variant
  Dim W As Variant

  Work.Clear

  OUT "Basic astronomical satistics for: " _
  & Day_Of_Week_For(The_Date) & " - " & The_Date & "  at UT " _
  & The_Time
  BL

  Q = Day_Frac_Equiv_To(The_Time) + Day_Frac_Equiv_To(The_TZ_Adjustment)
  OUT "Dynamical date & time: " & Dynamical_Date_and_Time(The_JDE)

  OUT " "

  OUT "The delta T correction is  " & The_Delta_T & " = " _
  & Format(The_dT, "#0.#############0") & " day (TD = UT + Delta T)"
  BL

  OUT "The JDE value is " & The_JDE & "  (adjusted for delta T)"
  BL

  Q = (The_JDE - 2451545) / 36525
  OUT "T  from J2000.0 = " & Format(Q, "#0.##############0") & _
  " Julian centuries"
  OUT "T  from J2000.0 = " & Format(Q / 10, "#0.##############0") & _
  " Julian millennia"

  BL
  OUT "Sidereal time at Greenwich"
  Q = LST_For(The_Date, The_Time, 0, False)

  OUT "Mean: " & Ang_Out(Q, HA_MODE, False)
  U = LST_For(The_Date, The_Time, 0, True)
  OUT "True: " & Ang_Out(U, HA_MODE, False)
  Q = U - Q
  OUT "Diff: " & Ang_Out(Q, HA_MODE, True)
  
  BL
  OUT "Obliquity of the ecliptic"
  Q = Ecliptic_Obliquity(The_JDE, "mean")
  OUT "Mean:" & Ang_Out(Q, DECL_LAT_MODE, False)
  Q = Ecliptic_Obliquity(The_JDE, "apparent")
  OUT "True:" & Ang_Out(Q, DECL_LAT_MODE, False)
  Q = Delta_e(The_JDE)
  OUT "Diff:" & Ang_Out(Q, DECL_LAT_MODE, True)

  BL
  OUT "Nutation"
  U = Delta_Psi(The_JDE)
  
  OUT "In longitude:" & Ang_Out(U, LNG_MODE, True)
  OUT "In obliquity:" & Ang_Out(Q, DECL_LAT_MODE, True)

  End Sub


' Generate a single heliocentric ephemeris line

  Private Function Heliocentric_Position_Of(Object_Name, At_JDE)
 
' Heliocentric VSOP87 spherical coordinates

  Dim HLBR As String
  Dim L    As String
  Dim B    As String
  Dim R    As String

' Compute the L,B,R values
  HLBR = LBR_For(Object_Name, At_JDE)

' Extract coordinates from returned data vector
  L = Val_of_Coord("L", HLBR)
  B = Val_of_Coord("B", HLBR)
  R = Val_of_Coord("R", HLBR)

' Format coordinates for output
  L = Ang_Out(L, LNG_MODE, False)
  B = Ang_Out(B, DECL_LAT_MODE, True)
  R = Dist_Out(R, DIST_UNITS, False)

  Heliocentric_Position_Of = L & " |" & B & " | " & R

  End Function


' Compute single geocentric ephemeris line.  This routine also computes the
' angular width of the object and the ideal visual magnitude.

  Private Function Geocentric_Position_Of(Object_Name, At_JDE)

' Work variables
  Dim Q As Variant
  Dim W As Variant

' Heliocentric coordinates data vector
  Dim HLBR As String

' Heliocentric spherical coordinates
  Dim L    As String
  Dim B    As String
  Dim R    As String

' Heliocentric rectangular coordinates
  Dim Xe   As Double
  Dim Ye   As Double
  Dim Ze   As Double

' Heliocentric & geocentric rectangular coordinates
  Dim x    As Double
  Dim y    As Double
  Dim Z    As Double

' Distance between Earth and Sun
  Dim Re   As Double

' Distance between object and Sun
  Dim Rp   As Double

' Distance between Earth and object
  Dim Dist As Double

' Light time between earth and object
  Dim LT   As Double

' Ecliptic obliquity (apparent)
  Dim e    As Double

' Angular diameter (equatorial) of object
  Dim Ang_Diam As String

' Phase angle between Earth and Object
  Dim Phase  As Double

' Visual magnitude
  Dim VMag   As String

' Compute the heliocentric L,B,R coordinates for Earth
  HLBR = LBR_For("Earth", At_JDE)

'  Extract coordinates from returned data vector
   L = Val_of_Coord("L", HLBR) * Atn(1) / 45
   B = Val_of_Coord("B", HLBR) * Atn(1) / 45
   R = Val_of_Coord("R", HLBR)
  Re = R

' Compute heliocentric X,Y,Z values for Earth
  Xe = R * Cos(B) * Cos(L)
  Ye = R * Cos(B) * Sin(L)
  Ze = R * Sin(B)

' If object is the Sun, then the Earth coordinates are the origin
  If Trim(Object_Name) = "Sun" Then
  Xe = 0: Ye = 0: Ze = 0
  End If

' Compute the heliocentric L,B,R coordinates for object
  HLBR = LBR_For(Object_Name, At_JDE)

'  Extract coordinates from returned data vector
   L = Val_of_Coord("L", HLBR) * Atn(1) / 45
   B = Val_of_Coord("B", HLBR) * Atn(1) / 45
   R = Val_of_Coord("R", HLBR)
  Rp = R

' Compute heliocentric X,Y,Z values for object
  x = R * Cos(B) * Cos(L)
  y = R * Cos(B) * Sin(L)
  Z = R * Sin(B)

' Replace heliocentric X,Y,Z of object with geocentric X,Y,Z
  x = x - Xe
  y = y - Ye
  Z = Z - Ze

' Compute true geometric distance between earth and object
  Dist = Sqr(x * x + y * y + Z * Z)

' Compute light time correction (in days) for distance
  LT = Dist * 5.77551830441213E-03

' RECOMPUTE VALUES TO CORRECT FOR LIGHT TIME AND ABERRATION

' Recompute same heliocentric and geocentric coordinates again
' at original time minus the light time
  
' Recompute the heliocentric L,B,R coordinates for Earth
  HLBR = LBR_For("Earth", At_JDE - LT)

' Extract coordinates from returned data vector
  L = Val_of_Coord("L", HLBR) * Atn(1) / 45
  B = Val_of_Coord("B", HLBR) * Atn(1) / 45
  R = Val_of_Coord("R", HLBR)

' Recompute heliocentric X,Y,Z values for Earth
  Xe = R * Cos(B) * Cos(L)
  Ye = R * Cos(B) * Sin(L)
  Ze = R * Sin(B)

' Recompute the heliocentric L,B,R coordinates for object
  HLBR = LBR_For(Object_Name, At_JDE - LT)

' Extract coordinates from returned data vector
  L = Val_of_Coord("L", HLBR) * Atn(1) / 45
  B = Val_of_Coord("B", HLBR) * Atn(1) / 45
  R = Val_of_Coord("R", HLBR)

' Recompute heliocentric X,Y,Z values for object
  x = R * Cos(B) * Cos(L)
  y = R * Cos(B) * Sin(L)
  Z = R * Sin(B)

' Replace heliocentric X,Y,Z of object with geocentric X,Y,Z
  x = x - Xe
  y = y - Ye
  Z = Z - Ze

' Now compute the raw geocentric ecliptical coordinates in degrees
  L = Atn(y / x): If x < 0 Then L = L + 4 * Atn(1)
      If L < 0 Then L = L + 8 * Atn(1)
  L = L * 45 / Atn(1)
  B = Atn(Z / Sqr(x * x + y * y)) * 45 / Atn(1)

' Apply reductions to FK5 system coordinates
  Q = FK5_Lng_Corr(At_JDE, L, B)
  B = B + FK5_Lat_Corr(At_JDE, L)
  L = L + Q

' Compute the apparent obliquity of the ecliptic (in radians)
  e = Ecliptic_Obliquity(At_JDE, "Apparent") * Atn(1) / 45

' Apply corrections for nutation to ecliptical longitude
  L = L + Delta_Psi(At_JDE)

' Compute phase angle between Earth and object in degrees
  Q = ((Rp * Rp) + (Dist * Dist) - (Re * Re)) / (2 * Rp * Dist)
  Phase = (Atn(-Q / Sqr((-Q * Q) + 1)) + 2 * Atn(1)) * 45 / Atn(1)

' Compute apparent equatorial angular diameter of object in degrees
' and visual magnitude (except for Saturn)
  Object_Name = Trim(Object_Name)
  If Object_Name = "Sun" Then
  Ang_Diam = 959.63 / 1800 / Dist
  VMag = "- "
  End If

  If Object_Name = "Mercury" Then
     Ang_Diam = 3.36 / 1800 / Dist
     VMag = -0.42 + 5 * Log10(Rp * Dist) + 0.038 * Phase _
          - 0.000273 * Phase * Phase + 0.000002 * Phase * Phase * Phase
  End If

  If Object_Name = "Venus" Then
     Ang_Diam = 8.41 / 1800 / Dist
     VMag = -4.4 + 5 * Log10(Rp * Dist) + 0.0009 * Phase _
          + 0.000239 * Phase * Phase - 0.00000065 * Phase * Phase * Phase
  End If

  If Object_Name = "Mars" Then
     Ang_Diam = 4.68 / 1800 / Dist
     VMag = -1.52 + 5 * Log10(Rp * Dist) + 0.016 * Phase
  End If

  If Object_Name = "Jupiter" Then
     Ang_Diam = 98.44 / 1800 / Dist
     VMag = -9.4 + 5 * Log10(Rp * Dist) + 0.005 * Phase
  End If

  If Object_Name = "Saturn" Then
     Ang_Diam = 82.73 / 1800 / Dist
     VMag = "- "
  End If

  If Object_Name = "Uranus" Then
     Ang_Diam = 35.02 / 1800 / Dist
     VMag = -7.19 + 5 * Log10(Rp * Dist)
  End If

  If Object_Name = "Neptune" Then
     Ang_Diam = 33.5 / 1800 / Dist
     VMag = -6.87 + 5 * Log10(Rp * Dist)
  End If
  
' Format the visual magnitude value
  If Object_Name <> "Saturn" And Object_Name <> "Sun" Then
  VMag = Format(VMag, "#0.0")
  If VMag > 0 Then VMag = "+" & VMag
  VMag = Right(Space(4) & VMag, 4)
  End If

' Format angular diameter
  Ang_Diam = Ang_Out(Ang_Diam, "DMS", False)
  Ang_Diam = Right(Ang_Diam, 10)
  
' Determine output mode and return results accordingly
  If COMP_MODE = "EC" Then
     L = Ang_Out(L, LNG_MODE, False)
     B = Ang_Out(B, DECL_LAT_MODE, True)
     R = Dist_Out(Dist, DIST_UNITS, False)
     Geocentric_Position_Of = L & " |" & B & " | " & R & " | " _
     & Ang_Diam & " | " & VMag
  Else
     L = L * Atn(1) / 45
     B = B * Atn(1) / 45
     y = Sin(L) * Cos(e) - Tan(B) * Sin(e)
     x = Cos(L)
     Q = Atn(y / x): If x < 0 Then Q = Q + 4 * Atn(1)
     If Q < 0 Then Q = Q + 8 * Atn(1)
     Q = Q * 45 / Atn(1)
     W = ArcSin(Sin(B) * Cos(e) + Cos(B) * Sin(e) * Sin(L), "d")
     L = Q: B = W
     L = Ang_Out(L, HA_MODE, False)
     B = Ang_Out(B, DECL_LAT_MODE, True)
     R = Dist_Out(Dist, DIST_UNITS, False)
     Geocentric_Position_Of = L & " |" & B & " | " & R & " | " _
     & Ang_Diam & " | " & VMag
  End If
  
  End Function

' -------------------------------------------------------------
' Compute ephemeris for each hour of selected date of a daily
' ephemeris listing.  To select a date, simply double-click the
' line within the displayed ephemeris.
'
' If the listing is already an hourly listing, the double-clicking
' on a line will display a table for each minute of that hour.


  Private Sub Work_DblClick()

  Dim Q  As Variant
  Dim W  As Variant

  Dim D  As String
  Dim dd As String
  Dim h  As Integer
  Dim HH As String
  Dim M  As Integer
  Dim MM As String
  Dim HG As String

  Dim i  As Integer ' List index to selected line

  Dim Date_String As String
  Dim Time_String As String
  Dim Base_JDE    As String
  Dim At_JDE      As String

  If COMP_MODE = "STATS" Then Exit Sub
  If HM_TABLE_MODE = "" Then Exit Sub
  If HM_TABLE_MODE = "H" Then GoTo HH_TABLE
  If HM_TABLE_MODE = "M" Then GoTo MM_TABLE

  Exit Sub

HH_TABLE:

' Check if safe to compute hourly table
  If InStr(Work.List(0), "hour") > 0 Then Exit Sub
  
' Get index to selected line
  i = Work.ListIndex

' Get the selected line contents
  Q = Work.List(i)

' If line evaluates to zero, then ignore double-click
  D = Val(Trim(Q)): If D = 0 Then Exit Sub
  TEMP_D = D

  NeoEphemerix_2001_Interface.MousePointer = vbHourglass

' Construct the base date string
  Date_String = D & " " & The_Month & " " & The_Year

' Compute the base dynamical JDE value
  Base_JDE = JDE_For(Date_String, 0) + The_dT
  
  Work.Clear

' Determine which table header to display
  If COMP_MODE = "HC" Then
  OUT "VSOP87 heliocentric ephemeris of" & The_Object _
  & " for each hour of " & Date_String
  OUT "Delta T = " & The_Delta_T & "    * = Millions"
  BL
  OUT "HH |   H Longitude   |    H Latitude   |     H Distance"
  End If

  If COMP_MODE = "EC" Then
  OUT "Geocentric ecliptical FK5 ephemeris of" & The_Object _
  & " for each hour of " & Date_String
  OUT "Delta T = " & The_Delta_T & "     * = Millions"

  BL
  OUT "HH      Longitude    |     Latitude    |      Distance   |  Ang Diam  |" _
  & " VMag"
  End If

  If COMP_MODE = "EQ" Then
  OUT "Geocentric equatorial FK5 ephemeris of" & The_Object _
  & " for each hour of " & Date_String
  OUT "Delta T = " & The_Delta_T & "     * = Millions"

  BL
  OUT "HH         RA        |       Decl      |     Distance    |  Ang Diam  |" _
  & " VMag"
  End If

' Generate a 24 hour ephemeris table for the selected date
  For h = 0 To 24
  HH = Format(h, "0#") & " |"

  Time_String = Str(h) & ":00:00"
  At_JDE = Base_JDE + Day_Frac_Equiv_To(Time_String)
  
' Generate an ephemeris line according to computation mode
  If COMP_MODE = "HC" Then
     Q = HH & Heliocentric_Position_Of(The_Object, At_JDE)
  Else
     Q = HH & Geocentric_Position_Of(The_Object, At_JDE)
  End If
   
  OUT Q

  Next h

  BL

' Prepare for possible selection of extended minute mode
  HM_TABLE_MODE = "M"

  NeoEphemerix_2001_Interface.MousePointer = vbDefault
  Work.ToolTipText = ""

  Message.Visible = True
  Message.Caption = " Double-click hour line in the ephemeris for a 60 minute ephemeris for that hour. "

  COMPUTE_Button.SetFocus
  Exit Sub

' --------
' Generate an ephemeris for each minute of the selected hour
MM_TABLE:
  
' Get index to selected line
  i = Work.ListIndex

' Get the selected line contents
  Q = Work.List(i)

' If line is not an hour line, then ignore double-click
  If i < 4 Or i > 28 Then Exit Sub

' Get value of hour
  h = Val(Trim(Q))

' If hour=24, then ignore double-click
  If h >= 24 Then Exit Sub

  NeoEphemerix_2001_Interface.MousePointer = vbHourglass

' Construct the base date string
  Date_String = TEMP_D & " " & The_Month & " " & The_Year

' Compute the base dynamical JDE value
  Base_JDE = JDE_For(Date_String, 0) + The_dT

  Work.Clear

  HH = Format(h, "0#")

' Determine which table header to display
  If COMP_MODE = "HC" Then
  OUT "VSOP87 heliocentric ephemeris of" & The_Object & " for each minute"
  OUT "of hour " & HH & ":00 of date " & Date_String & "   Delta T = " & The_Delta_T _
  & "    * = Millions"

  BL
  OUT "MM |   H Longitude   |    H Latitude   |     H Distance"
  End If

  If COMP_MODE = "EC" Then
  OUT "Geocentric ecliptical FK5 ephemeris of" & The_Object & " for each minute"
  OUT "of hour " & HH & ":00 of date " & Date_String & "   Delta T = " & The_Delta_T _
  & "    * = Millions"

  BL
  OUT "MM |    Longitude    |     Latitude    |     Distance    |  Ang Diam  |" _
  & " VMag"
  End If

  If COMP_MODE = "EQ" Then
  OUT "Geocentric equatorial FK5 ephemeris of" & The_Object & " for each minute"
  OUT "of hour " & HH & ":00 of date " & Date_String & "   Delta T = " & The_Delta_T _
  & "    * = Millions"

  BL
  OUT "MM |       RA        |       Decl      |     Distance    |  Ang Diam  |" _
  & " VMag"
  End If

' Generate a 60 minute ephemeris table for the selected hour/date
  For M = 0 To 60
  MM = Format(M, "0#") & " |"

  Time_String = Format(h, "0#") & ":" & MM & ":00"
  At_JDE = Base_JDE + Day_Frac_Equiv_To(Time_String)
  
' Generate an ephemeris line according to computation mode
  If COMP_MODE = "HC" Then
     Q = MM & Heliocentric_Position_Of(The_Object, At_JDE)
  Else
     Q = MM & Geocentric_Position_Of(The_Object, At_JDE)
  End If
   
  OUT Q

  Next M

  BL

' Set end of extended table select mode
  HM_TABLE_MODE = ""

  NeoEphemerix_2001_Interface.MousePointer = vbDefault
  Work.ToolTipText = ""
  Message.Visible = False

  COMPUTE_Button.SetFocus

  End Sub

' ------------------------------------------------------------------------
' Compute an apparent geocentric ephemeris for all objects at current
' interface date and time settings

  Private Sub COMPUTE_ALL_OBJECTS()

  Dim Q   As Variant
  Dim W   As Variant
  
  Dim At_JDE As Double
      At_JDE = The_JDE

  Work.Clear

' Compute apparent geocentric ecliptical ephemeris in FK5 coordinates
  COMP_MODE = "EQ"
  OUT "Apparent geocentric equatorial FK5 ephemeris for " & The_Date & "  at " _
  & The_Time & " UT"
  OUT "* = millions"
  BL
  OUT "               RA       |       Decl      |     Distance    |  Ang Diam  | VMag"
  OUT "Sun    " & Geocentric_Position_Of("Sun", At_JDE)
  OUT "Mercury" & Geocentric_Position_Of("Mercury", At_JDE)
  OUT "Venus  " & Geocentric_Position_Of("Venus", At_JDE)
  OUT "Mars   " & Geocentric_Position_Of("Mars", The_JDE)
  OUT "Jupiter" & Geocentric_Position_Of("Jupiter", The_JDE)
  OUT "Saturn " & Geocentric_Position_Of("Saturn", The_JDE)
  OUT "Uranus " & Geocentric_Position_Of("Uranus", The_JDE)
  OUT "Neptune" & Geocentric_Position_Of("Neptune", The_JDE)

' Compute apparent geocentric equatorial ephemeris in FK5 coordinates
  OUT String(80, "-")
  COMP_MODE = "EQ"
  OUT "Apparent geocentric ecliptical FK5 ephemeris for " & The_Date & "  at " _
  & The_Time & " UT"
  BL
  OUT "       Longitude    |     Latitude    |"
  COMP_MODE = "EC"
  OUT "Sun" & Left(Geocentric_Position_Of("Sun", The_JDE), 36)
  Q = Left(Geocentric_Position_Of("Mercury", The_JDE), 36)
  OUT "Mer" & Q
  Q = Left(Geocentric_Position_Of("Venus", The_JDE), 36)
  OUT "Ven" & Q
  Q = Left(Geocentric_Position_Of("Mars", The_JDE), 36)
  OUT "Mar" & Q
  Q = Left(Geocentric_Position_Of("Jupiter", The_JDE), 36)
  OUT "Jup" & Q
  Q = Left(Geocentric_Position_Of("Saturn", The_JDE), 36)
  OUT "Sat" & Q
  Q = Left(Geocentric_Position_Of("Uranus", The_JDE), 36)
  OUT "Ura" & Q
  W = Left(Geocentric_Position_Of("Neptune", The_JDE), 36)
  OUT "Nep" & W

' Compute VSOP87 heliocentric ephemeris
  OUT String(80, "-")
  OUT "VSOP87 heliocentric ephemeris for " & The_Date & "  at " _
  & The_Time & " UT"
  BL
  OUT "       H Longitude  |     H Latitude  |     H Distance"
  Q = Heliocentric_Position_Of("Mercury", The_JDE)
  OUT "Mer" & Q
  Q = Heliocentric_Position_Of("Venus", The_JDE)
  OUT "Ven" & Q
  Q = Heliocentric_Position_Of("Earth", The_JDE)
    OUT "Ear" & Q
  Q = Heliocentric_Position_Of("Mars", The_JDE)
  OUT "Mar" & Q
  Q = Heliocentric_Position_Of("Jupiter", The_JDE)
  OUT "Jup" & Q
  Q = Heliocentric_Position_Of("Saturn", The_JDE)
  OUT "Sat" & Q
  Q = Heliocentric_Position_Of("Uranus", The_JDE)
  OUT "Ura" & Q
  W = Heliocentric_Position_Of("Neptune", The_JDE)
  OUT "Nep" & W
 
  COMP_MODE = "ALL"
  HM_TABLE_MODE = True
  Work.ToolTipText = ""

  End Sub

