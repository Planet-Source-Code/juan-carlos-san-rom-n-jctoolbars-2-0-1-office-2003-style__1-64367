VERSION 5.00
Object = "*\Ajc_Toolbar.vbp"
Begin VB.Form ToolbarText 
   Caption         =   "jcToolbars  v2.0.1"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9075
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   9075
   StartUpPosition =   2  'CenterScreen
   Begin jcToolbars.JCToolbar JCToolbar1 
      Height          =   705
      Index           =   4
      Left            =   30
      TabIndex        =   12
      ToolTipText     =   "Toolbar options"
      Top             =   1770
      Visible         =   0   'False
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   1244
      BackColor       =   -2147483633
      ButtonCount     =   7
      BtnCaption1     =   "Access doc"
      BtnEnabled1     =   0   'False
      BtnIcon1        =   "ToolbarText.frx":0000
      BeginProperty BtnFont1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BtnLeft1        =   2
      BtnTop1         =   2
      BtnWidth1       =   84
      BtnHeight1      =   24
      BtnCaption2     =   "jcbtn"
      BtnEnabled2     =   0   'False
      BtnType2        =   1
      BeginProperty BtnFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BtnLeft2        =   88
      BtnTop2         =   4
      BtnWidth2       =   2
      BtnHeight2      =   38
      BtnCaption3     =   "Excel doc"
      BtnEnabled3     =   0   'False
      BtnIcon3        =   "ToolbarText.frx":0352
      BtnAlignment3   =   1
      BeginProperty BtnFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BtnLeft3        =   92
      BtnTop3         =   2
      BtnWidth3       =   75
      BtnHeight3      =   24
      BtnCaption4     =   "jcbtn"
      BtnEnabled4     =   0   'False
      BtnType4        =   1
      BeginProperty BtnFont4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BtnLeft4        =   169
      BtnTop4         =   4
      BtnWidth4       =   2
      BtnHeight4      =   38
      BtnCaption5     =   "Pdf doc"
      BtnEnabled5     =   0   'False
      BtnIcon5        =   "ToolbarText.frx":06A4
      BtnAlignment5   =   2
      BeginProperty BtnFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BtnLeft5        =   173
      BtnTop5         =   2
      BtnWidth5       =   45
      BtnHeight5      =   42
      BtnCaption6     =   "jcbtn"
      BtnEnabled6     =   0   'False
      BtnIconSize6    =   32
      BtnType6        =   1
      BeginProperty BtnFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BtnLeft6        =   220
      BtnTop6         =   4
      BtnWidth6       =   2
      BtnHeight6      =   38
      BtnCaption7     =   "Word doc"
      BtnEnabled7     =   0   'False
      BtnIcon7        =   "ToolbarText.frx":09F6
      BtnAlignment7   =   3
      BeginProperty BtnFont7 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BtnLeft7        =   224
      BtnTop7         =   2
      BtnWidth7       =   55
      BtnHeight7      =   42
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4485
      Left            =   6660
      ScaleHeight     =   4485
      ScaleWidth      =   2295
      TabIndex        =   3
      ToolTipText     =   "Here is the simplest way of drawing xp style frame"
      Top             =   2400
      Width           =   2295
      Begin VB.CommandButton CmdForeColor 
         Caption         =   "Change color text of Copy button to red"
         Height          =   555
         Left            =   150
         TabIndex        =   17
         Top             =   3750
         Width           =   1965
      End
      Begin VB.CommandButton CmdDisabled 
         Caption         =   "Home Button disabled"
         Height          =   345
         Left            =   150
         TabIndex        =   16
         Top             =   2940
         Width           =   1965
      End
      Begin VB.CommandButton CmdIconSize 
         Caption         =   "Increase IconSize"
         Height          =   345
         Left            =   150
         TabIndex        =   15
         Top             =   3330
         Width           =   1965
      End
      Begin VB.CommandButton CmdMove 
         Caption         =   "Move Paste button left"
         Height          =   345
         Left            =   150
         TabIndex        =   10
         Top             =   2550
         Width           =   1965
      End
      Begin VB.CommandButton CmdDeleteBtn 
         Caption         =   "Delete Find button"
         Height          =   345
         Left            =   150
         TabIndex        =   9
         Top             =   2160
         Width           =   1965
      End
      Begin VB.CommandButton CmdChangeCaption 
         Caption         =   "Change Toolbar options button caption"
         Height          =   555
         Left            =   150
         TabIndex        =   7
         Tag             =   "This is a button caption changed at runtime"
         Top             =   1530
         Width           =   1965
      End
      Begin VB.CommandButton CmdAddBtn 
         Caption         =   "Add button"
         Height          =   345
         Left            =   150
         TabIndex        =   6
         ToolTipText     =   "Add button to the last toolbar"
         Top             =   750
         Width           =   1965
      End
      Begin VB.CommandButton CmdAddSep_Btn 
         Caption         =   "Add separator"
         Height          =   345
         Left            =   150
         TabIndex        =   5
         ToolTipText     =   "Add separator to the last toolbar"
         Top             =   1140
         Width           =   1965
      End
      Begin VB.CommandButton CmdToolbar 
         Caption         =   "Add Toolbar"
         Height          =   345
         Left            =   150
         TabIndex        =   4
         ToolTipText     =   "Add toolbar at the end"
         Top             =   360
         Width           =   1965
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   90
      ScaleHeight     =   3015
      ScaleWidth      =   6435
      TabIndex        =   1
      ToolTipText     =   "Here is the simplest way of drawing xp style frame"
      Top             =   3870
      Width           =   6435
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "aaa"
         Height          =   195
         Left            =   210
         TabIndex        =   2
         Top             =   300
         Width           =   270
      End
   End
   Begin jcToolbars.JCToolbar JCToolbar1 
      Height          =   435
      Index           =   0
      Left            =   8640
      TabIndex        =   0
      ToolTipText     =   "Toolbar options"
      Top             =   30
      Visible         =   0   'False
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   767
      BackColor       =   -2147483633
   End
   Begin VB.Timer TimMove 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8640
      Top             =   540
   End
   Begin jcToolbars.JCToolbar JCToolbar1 
      Height          =   435
      Index           =   1
      Left            =   30
      TabIndex        =   11
      ToolTipText     =   "Toolbar options"
      Top             =   60
      Visible         =   0   'False
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   767
      BackColor       =   -2147483633
      ButtonCount     =   7
      BtnCaption1     =   "Home"
      BtnEnabled1     =   0   'False
      BtnIcon1        =   "ToolbarText.frx":0D48
      BtnToolTipText1 =   "Go to home page"
      BtnKey1         =   "Home"
      BeginProperty BtnFont1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BtnLeft1        =   2
      BtnTop1         =   2
      BtnWidth1       =   56
      BtnHeight1      =   24
      BtnCaption2     =   "Contact"
      BtnEnabled2     =   0   'False
      BtnIcon2        =   "ToolbarText.frx":109A
      BtnToolTipText2 =   "Contact"
      BtnKey2         =   "Contact"
      BeginProperty BtnFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BtnLeft2        =   60
      BtnTop2         =   2
      BtnWidth2       =   65
      BtnHeight2      =   24
      BtnCaption3     =   "jcbtn"
      BtnEnabled3     =   0   'False
      BtnType3        =   1
      BeginProperty BtnFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BtnLeft3        =   127
      BtnTop3         =   4
      BtnWidth3       =   2
      BtnHeight3      =   20
      BtnCaption4     =   "Toolbar options"
      BtnEnabled4     =   0   'False
      BtnToolTipText4 =   "Example of dropdown button"
      BtnKey4         =   "tlbrOpt"
      BtnStyle4       =   2
      BeginProperty BtnFont4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BtnLeft4        =   131
      BtnTop4         =   2
      BtnWidth4       =   94
      BtnHeight4      =   21
      BtnCaption5     =   "jcbtn"
      BtnEnabled5     =   0   'False
      BtnType5        =   1
      BeginProperty BtnFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BtnLeft5        =   227
      BtnTop5         =   4
      BtnWidth5       =   2
      BtnHeight5      =   20
      BtnUseMaskColor5=   0   'False
      BtnMaskColor5   =   0
      BtnCaption6     =   "Full screen"
      BtnEnabled6     =   0   'False
      BtnIcon6        =   "ToolbarText.frx":13EC
      BtnToolTipText6 =   "Full screen"
      BtnKey6         =   "Full screen"
      BeginProperty BtnFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BtnLeft6        =   231
      BtnTop6         =   2
      BtnWidth6       =   79
      BtnHeight6      =   24
      BtnCaption7     =   "Find"
      BtnEnabled7     =   0   'False
      BtnIcon7        =   "ToolbarText.frx":173E
      BtnToolTipText7 =   "Find"
      BtnKey7         =   "Find"
      BeginProperty BtnFont7 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BtnLeft7        =   312
      BtnTop7         =   2
      BtnWidth7       =   48
      BtnHeight7      =   24
   End
   Begin jcToolbars.JCToolbar JCToolbar1 
      Height          =   435
      Index           =   2
      Left            =   90
      TabIndex        =   13
      ToolTipText     =   "Toolbar options"
      Top             =   570
      Visible         =   0   'False
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   767
      BackColor       =   -2147483633
      ButtonCount     =   11
      BtnCaption1     =   "Exit Demo"
      BtnEnabled1     =   0   'False
      BtnIcon1        =   "ToolbarText.frx":1A90
      BtnToolTipText1 =   "Exit this demo"
      BtnKey1         =   "Exit"
      BeginProperty BtnFont1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BtnLeft1        =   2
      BtnTop1         =   2
      BtnWidth1       =   76
      BtnHeight1      =   24
      BtnCaption2     =   "jcbtn"
      BtnEnabled2     =   0   'False
      BtnType2        =   1
      BeginProperty BtnFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BtnLeft2        =   80
      BtnTop2         =   4
      BtnWidth2       =   2
      BtnHeight2      =   20
      BtnEnabled3     =   0   'False
      BtnIcon3        =   "ToolbarText.frx":1DE2
      BtnKey3         =   "left"
      BtnStyle3       =   1
      BeginProperty BtnFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BtnLeft3        =   84
      BtnTop3         =   2
      BtnWidth3       =   24
      BtnHeight3      =   24
      BtnEnabled4     =   0   'False
      BtnIcon4        =   "ToolbarText.frx":2134
      BtnKey4         =   "center"
      BtnStyle4       =   1
      BeginProperty BtnFont4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BtnState4       =   2
      BtnValue4       =   -1  'True
      BtnLeft4        =   110
      BtnTop4         =   2
      BtnWidth4       =   24
      BtnHeight4      =   24
      BtnEnabled5     =   0   'False
      BtnIcon5        =   "ToolbarText.frx":2486
      BtnKey5         =   "right"
      BtnStyle5       =   1
      BeginProperty BtnFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BtnLeft5        =   136
      BtnTop5         =   2
      BtnWidth5       =   24
      BtnHeight5      =   24
      BtnCaption6     =   "jcbtn"
      BtnEnabled6     =   0   'False
      BtnType6        =   1
      BeginProperty BtnFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BtnLeft6        =   162
      BtnTop6         =   4
      BtnWidth6       =   2
      BtnHeight6      =   20
      BtnCaption7     =   "Copy"
      BtnEnabled7     =   0   'False
      BtnIcon7        =   "ToolbarText.frx":27D8
      BeginProperty BtnFont7 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BtnLeft7        =   166
      BtnTop7         =   2
      BtnWidth7       =   52
      BtnHeight7      =   24
      BtnCaption8     =   "Paste"
      BtnEnabled8     =   0   'False
      BtnIcon8        =   "ToolbarText.frx":2B2A
      BeginProperty BtnFont8 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BtnLeft8        =   220
      BtnTop8         =   2
      BtnWidth8       =   55
      BtnHeight8      =   24
      BtnCaption9     =   "Cut"
      BtnEnabled9     =   0   'False
      BtnIcon9        =   "ToolbarText.frx":2E7C
      BeginProperty BtnFont9 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BtnLeft9        =   277
      BtnTop9         =   2
      BtnWidth9       =   44
      BtnHeight9      =   24
      BtnCaption10    =   "jcbtn"
      BtnEnabled10    =   0   'False
      BtnType10       =   1
      BeginProperty BtnFont10 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BtnLeft10       =   323
      BtnTop10        =   4
      BtnWidth10      =   2
      BtnHeight10     =   20
      BtnCaption11    =   "Help"
      BtnEnabled11    =   0   'False
      BtnIcon11       =   "ToolbarText.frx":31CE
      BeginProperty BtnFont11 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BtnLeft11       =   327
      BtnTop11        =   2
      BtnWidth11      =   50
      BtnHeight11     =   24
   End
   Begin jcToolbars.JCToolbar JCToolbar1 
      Height          =   675
      Index           =   3
      Left            =   30
      TabIndex        =   14
      ToolTipText     =   "Toolbar options"
      Top             =   1080
      Visible         =   0   'False
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   1191
      BackColor       =   -2147483633
      ButtonCount     =   11
      BtnEnabled1     =   0   'False
      BtnIcon1        =   "ToolbarText.frx":3520
      BtnToolTipText1 =   "Iconsize 16"
      BeginProperty BtnFont1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BtnLeft1        =   2
      BtnTop1         =   2
      BtnWidth1       =   24
      BtnHeight1      =   24
      BtnEnabled2     =   0   'False
      BtnIcon2        =   "ToolbarText.frx":5072
      BtnIconSize2    =   24
      BtnToolTipText2 =   "Iconsize 24"
      BeginProperty BtnFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BtnLeft2        =   28
      BtnTop2         =   2
      BtnWidth2       =   32
      BtnHeight2      =   32
      BtnEnabled3     =   0   'False
      BtnIcon3        =   "ToolbarText.frx":6BC4
      BtnIconSize3    =   32
      BtnToolTipText3 =   "Iconsize 32"
      BeginProperty BtnFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BtnLeft3        =   62
      BtnTop3         =   2
      BtnWidth3       =   40
      BtnHeight3      =   40
      BtnCaption4     =   "jcbtn"
      BtnEnabled4     =   0   'False
      BtnType4        =   1
      BeginProperty BtnFont4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BtnLeft4        =   104
      BtnTop4         =   4
      BtnWidth4       =   2
      BtnHeight4      =   36
      BtnCaption5     =   "normal text"
      BtnEnabled5     =   0   'False
      BeginProperty BtnFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BtnLeft5        =   108
      BtnTop5         =   2
      BtnWidth5       =   73
      BtnHeight5      =   28
      BtnCaption6     =   "bold text"
      BtnEnabled6     =   0   'False
      BeginProperty BtnFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BtnLeft6        =   183
      BtnTop6         =   2
      BtnWidth6       =   61
      BtnHeight6      =   28
      BtnCaption7     =   "italic text"
      BtnEnabled7     =   0   'False
      BeginProperty BtnFont7 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      BtnLeft7        =   246
      BtnTop7         =   2
      BtnWidth7       =   60
      BtnHeight7      =   28
      BtnCaption8     =   "jcbtn"
      BtnEnabled8     =   0   'False
      BtnType8        =   1
      BeginProperty BtnFont8 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BtnLeft8        =   308
      BtnTop8         =   4
      BtnWidth8       =   2
      BtnHeight8      =   36
      BtnCaption9     =   "AB"
      BtnEnabled9     =   0   'False
      BtnToolTipText9 =   "fontsize 8 pt"
      BeginProperty BtnFont9 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BtnLeft9        =   312
      BtnTop9         =   2
      BtnWidth9       =   23
      BtnHeight9      =   22
      BtnCaption10    =   "AB"
      BtnEnabled10    =   0   'False
      BtnToolTipText10=   "fontsize 12 pt"
      BeginProperty BtnFont10 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BtnLeft10       =   337
      BtnTop10        =   2
      BtnWidth10      =   30
      BtnHeight10     =   26
      BtnCaption11    =   "AB"
      BtnEnabled11    =   0   'False
      BtnToolTipText11=   "fontsize 16 pt"
      BeginProperty BtnFont11 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BtnLeft11       =   369
      BtnTop11        =   2
      BtnWidth11      =   35
      BtnHeight11     =   32
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Move toolbars and resize form to see what happen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   825
      Left            =   6930
      TabIndex        =   8
      Top             =   150
      Width           =   1875
   End
   Begin VB.Menu mnufile 
      Caption         =   "file"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu mnuoption 
         Caption         =   "Show menu"
         Index           =   0
         Begin VB.Menu mnushow 
            Caption         =   "Hide Toolbar number 2"
            Index           =   0
         End
         Begin VB.Menu mnushow 
            Caption         =   "Show  toolbar number 2"
            Index           =   1
         End
      End
      Begin VB.Menu mnuoption 
         Caption         =   "Change Theme color"
         Index           =   1
         Begin VB.Menu MnuTheme 
            Caption         =   "Blue"
            Index           =   0
         End
         Begin VB.Menu MnuTheme 
            Caption         =   "Silver"
            Index           =   1
         End
         Begin VB.Menu MnuTheme 
            Caption         =   "Olive"
            Index           =   2
         End
         Begin VB.Menu MnuTheme 
            Caption         =   "Visual studio 2005"
            Index           =   3
         End
         Begin VB.Menu MnuTheme 
            Caption         =   "Norton 2004"
            Index           =   4
         End
         Begin VB.Menu MnuTheme 
            Caption         =   "Autodetect"
            Index           =   5
         End
      End
   End
End
Attribute VB_Name = "ToolbarText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NeoX As Long, IniWidth() As Integer, selection As Integer, inum As Integer
'/* Used to draw the form's rounded border
Private Declare Function RoundRect Lib "gdi32" _
    (ByVal hDC As Long, ByVal Left As Long, ByVal Top As Long, _
     ByVal Right As Long, ByVal Bottom As Long, ByVal EllipseWidth As Long, _
     ByVal EllipseHeight As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hDC As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As Any) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private lpp As POINTAPI

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Const DT_CENTER = &H1
Private Const DT_BOTTOM = &H8
Private Const DT_RIGHT = &H2
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_SINGLELINE = &H20
Private Const DT_NOCLIP = &H100
Private Const DT_LEFT = &H0

Dim themeType As ThemeConst
Dim StateBtn As jcBtnState
Dim blnFormLoaded As Boolean

Private Sub CmdDisabled_Click()
    If CmdDisabled.Caption = "Home Button disabled" Then
        StateBtn = STA_DISABLED
        CmdDisabled.Caption = "Home Button enabled"
    Else
        StateBtn = STA_NORMAL
        CmdDisabled.Caption = "Home Button disabled"
    End If
    JCToolbar1(1).ChangeBtnProperty jcState, 1, StateBtn
    
End Sub

Private Sub CmdAddBtn_Click()
    If inum > 13 Then Exit Sub
    With JCToolbar1(JCToolbar1.Count - 1)
        .AddButton Button, LoadResString(100 + inum), , LoadResPicture(inum, vbResIcon), 16, , "New button " & inum & " added", , , , , True, vbBlack
        IniWidth(JCToolbar1.Count - 1) = JCToolbar1(JCToolbar1.Count - 1).Width
    End With
    inum = inum + 1
    Form_Resize
End Sub

Private Sub CmdAddSep_Btn_Click()
    If inum > 13 Then Exit Sub
    With JCToolbar1(JCToolbar1.Count - 1)
        .AddButton Separator
        IniWidth(JCToolbar1.Count - 1) = JCToolbar1(JCToolbar1.Count - 1).Width
    End With
    Form_Resize
End Sub

Private Sub CmdForeColor_Click()
    If CmdForeColor.Caption = "Change color text of Copy button to red" Then
        JCToolbar1(2).ChangeBtnProperty jcBtnForeColor, 7, vbRed
        CmdForeColor.Caption = "Change color text of Copy button to black"
    Else
        JCToolbar1(2).ChangeBtnProperty jcBtnForeColor, 7, vbBlack
        CmdForeColor.Caption = "Change color text of Copy button to red"
    End If
End Sub

Private Sub CmdIconSize_Click()
    If CmdIconSize.Caption = "Increase IconSize" Then
        JCToolbar1(3).ChangeBtnProperty jcIconSize, 3, 48
        CmdIconSize.Caption = "Reduce IconSize"
    Else
        JCToolbar1(3).ChangeBtnProperty jcIconSize, 3, 32
        CmdIconSize.Caption = "Increase IconSize"
    End If
    AlignToolbars
End Sub

Private Sub CmdToolbar_Click()
    If CmdToolbar.Caption = "Add Toolbar" Then
        CmdToolbar.Caption = "Remove Toolbar"
        CmdToolbar.ToolTipText = "Remove last toolbar"
        AddToolbar
    Else
        CmdToolbar.Caption = "Add Toolbar"
        CmdToolbar.ToolTipText = "Add toolbar at the end"
        RemoveToolbar
    End If
End Sub

Private Sub CmdChangeCaption_Click()
    If CmdChangeCaption.Tag = "This is a button caption changed at runtime" Then
        JCToolbar1(1).ChangeBtnProperty jcCaption, 4, "This is a button caption changed at runtime"
        CmdChangeCaption.Tag = "Toolbar options"
    Else
        JCToolbar1(1).ChangeBtnProperty jcCaption, 4, "Toolbar options"
        CmdChangeCaption.Tag = "This is a button caption changed at runtime"
    End If
    IniWidth(1) = JCToolbar1(1).Width
    Form_Resize
End Sub

Private Sub CmdDeleteBtn_Click()
    If CmdDeleteBtn.Caption = "Delete Find button" Then
        JCToolbar1(1).DeleteButton (7)
        CmdDeleteBtn.Caption = "Add Find button"
    Else
       JCToolbar1(1).AddButton Button, "Find", "Find", LoadPicture(App.Path & "\ctf_iconos\find.ico"), 16, , "Find"
       CmdDeleteBtn.Caption = "Delete Find button"
    End If
    IniWidth(1) = JCToolbar1(1).Width
End Sub

Private Sub CmdMove_Click()
    If CmdMove.Caption = "Move Paste button left" Then
        CmdMove.Caption = "Move Paste button right"
        JCToolbar1(2).MoveButton 8, ToLeft
    Else
        CmdMove.Caption = "Move Paste button left"
        JCToolbar1(2).MoveButton 7, ToRight
    End If
End Sub

Private Sub CmdMoveToRight_Click()
    JCToolbar1(1).MoveButton 2, ToRight
End Sub

Private Sub Command7_Click()
    JCToolbar1(3).ChangeBtnProperty jcIconSize, 1, 16
    AlignToolbars
End Sub

Private Sub Form_Load()
    Dim i As Integer
    inum = 1
    StateBtn = STA_NORMAL
    blnFormLoaded = False

'*****************************************************
' differente examples of toolbars designed at runtime
'*****************************************************

'=====================================================================
' Example # 1: the use of dropdown button style and caption color
'=====================================================================
'    i = JCToolbar1.Count
'    Load JCToolbar1(i)
'    With JCToolbar1(i)
'        .AddButton Button, "Home", , LoadPicture(App.Path & "\ctf_iconos\home.ico"), 16, , "Go to home page"
'        .AddButton Button, "Contact", , LoadPicture(App.Path & "\ctf_iconos\contact.ico"), 16, , "Contact"
'        .AddButton Separator
'        .AddButton Button, "Toolbar options", "tlbrOpt", , , , "Example of dropdown button", vbBlue, , [Dropdown button], STA_NORMAL
'        .AddButton Separator
'        .AddButton Button, "Find", "Find", LoadPicture(App.Path & "\ctf_iconos\find.ico"), 16, , "Find"
'        .AddButton Button, "Full screen", "Full screen", LoadPicture(App.Path & "\ctf_iconos\full screen.ico"), 16, , "Full screen"
'        .Visible = True
'    End With
'---------------------------------------------------------------------

'=====================================================================
' Example # 2: the use of resource file to store pictures and strings
'=====================================================================
'    i = JCToolbar1.Count
'    Load JCToolbar1(i)
'    With JCToolbar1(i)
'        .AddButton Button, , "left", LoadResPicture(1, vbResIcon), 16, , "Left aligment", vbRed, , True
'        .AddButton Button, , "center", LoadResPicture(2, vbResIcon), 16, , "Center aligment", , , True
'        .AddButton Button, , "right", LoadResPicture(3, vbResIcon), 16, , "Right aligment", , , True
'        .AddButton Separator
'        .AddButton Button, "Copy", "copy", LoadPicture(App.Path & "\ctf_iconos\copy.ico"), 16
'        .AddButton Button, "Paste", "paste", LoadPicture(App.Path & "\ctf_iconos\paste.ico"), 16
'        .AddButton Button, "Cut", "cut", LoadPicture(App.Path & "\ctf_iconos\cut.ico"), 16
'        .AddButton Separator
'        .AddButton Button, "Help", "help", LoadPicture(App.Path & "\ctf_iconos\help-2.ico"), 16
'        .AddButton Button, "Exit", "exit", LoadPicture(App.Path & "\ctf_iconos\exit.ico"), 16, , "Exit demo"
'        .Visible = True
'    End With
'---------------------------------------------------------------------

'====================================================
' Example # 3: the use of different fonts and sizes
'====================================================
'    Dim F1 As New StdFont
'    Dim F2 As New StdFont
'    Dim F3 As New StdFont
'    Dim F4 As New StdFont
'    Dim F5 As New StdFont
'    Dim F6 As New StdFont
'
'    With F1
'        .Size = 12
'        .Bold = False
'        .Italic = False
'        .Name = "Arial Narrow"
'    End With
'
'    With F2
'        .Size = 12
'        .Bold = True
'        .Italic = False
'        .Name = "Arial Narrow"
'    End With
'
'    With F3
'        .Size = 12
'        .Bold = False
'        .Italic = True
'        .Name = "Arial Narrow"
'    End With
'
'    With F4
'        .Size = 8
'        .Name = "Arial"
'    End With
'
'    With F5
'        .Size = 12
'        .Name = "Arial"
'    End With
'
'    With F6
'        .Size = 16
'        .Name = "Arial"
'    End With
'
'    i = JCToolbar1.Count
'    Load JCToolbar1(i)
'    With JCToolbar1(i)
'        .AddButton Button, , , LoadPicture(App.Path & "\ctf_iconos\padesk.ico"), 16, , "Iconsize 16"
'        .AddButton Button, , , LoadPicture(App.Path & "\ctf_iconos\padesk.ico"), 24, , "Iconsize 24"
'        .AddButton Button, , , LoadPicture(App.Path & "\ctf_iconos\padesk.ico"), 32, , "Iconsize 32"
'        .AddButton Separator
'        .AddButton Button, "normal text", "Only text", , , , , QBColor(1), F1
'        .AddButton Button, "bold text", "Only text", , , , , QBColor(1), F2
'        .AddButton Button, "italic text", "Only text", , , , , QBColor(1), F3
'        .AddButton Separator
'        .AddButton Button, "AB", , , , , "fontsize 8 pt", , F4
'        .AddButton Button, "AB", , , , , "fontsize 12 pt", , F5
'        .AddButton Button, "AB", , , , , "fontsize 16 pt", , F6
'        .Visible = True
'    End With
'---------------------------------------------------------------------

'=============================================================
' Example # 4: the use of 4 types of icon and text alignments
'=============================================================
'    i = JCToolbar1.Count
'    Load JCToolbar1(i)
'    With JCToolbar1(i)
'        .AddButton Button, "Access doc", , LoadPicture(App.Path & "\ctf_iconos\doc access.ico"), 16, IconLeftTextRightI, "Icon left and text right"
'        .AddButton Separator
'        .AddButton Button, "Excel doc", , LoadPicture(App.Path & "\ctf_iconos\doc excel.ico"), 16, IconRightTextLeft, "Icon right and text left"
'        .AddButton Separator
'        .AddButton Button, "Pdf doc", , LoadPicture(App.Path & "\ctf_iconos\doc pdf.ico"), 16, IconTopTextBottom, "Icon top and text bottom"
'        .AddButton Separator
'        .AddButton Button, "Word doc", , LoadPicture(App.Path & "\ctf_iconos\doc word.ico"), 16, IconBottomTextTop, "Icon bottom and text top"
'        .Visible = True
'    End With
'-------------------------------------------------------------
    
    ReDim IniWidth(JCToolbar1.Count)
    
    AlignToolbars
    
    JCToolbar1(1).MenuLanguage = Spanish
    JCToolbar1(3).ShowMenuColor = False
        
    Label3.Caption = "- Buttons and separators can be added at design and runtime " & Chr(13)
    Label3.Caption = Label3.Caption & "- A property page have been added for toolbar design" & Chr(13)
    Label3.Caption = Label3.Caption & "- Autohide toolbar buttons when toolbar is sizing" & Chr(13)
    Label3.Caption = Label3.Caption & "- Toolbar autosizing taking into account icon and font size used in toolbar buttons " & Chr(13)
    Label3.Caption = Label3.Caption & "- Added function to convert color icon to grayscale icon when button is disabled" & Chr(13)
    Label3.Caption = Label3.Caption & "- Different icon sizes and fonts (type, size, bold, italic, forecolor) can be used" & Chr(13)
    Label3.Caption = Label3.Caption & "- Four type of icon and caption alignments " & Chr(13)
    Label3.Caption = Label3.Caption & "- Added dropdown button style, useful for popup menu use (see Toolbar options button)" & Chr(13)
    Label3.Caption = Label3.Caption & "- Popup theme color Menu on ThemeColorClick event that shows five theme colors" & Chr(13)
    Label3.Caption = Label3.Caption & "- Windows XP theme auto detection or selection  (blue, silver and olive) " & Chr(13)
    Label3.Caption = Label3.Caption & "- Two customs theme colors (norton 2004 and visual studio 2005) have been added" & Chr(13)
    Label3.Caption = Label3.Caption & "- You can select language for theme color Menu (spanish and english)" & Chr(13)
    Label3.Caption = Label3.Caption & "- You can determine if theme color Menu is shown or not (see the third toolbar)" & Chr(13)
    PaintFrame Picture1, "jcToolbars version 2.0.1"
    PaintFrame Picture3, "Changes at runtime"
    
    blnFormLoaded = True
End Sub

Private Sub Form_Resize()
    Dim i As Integer
    If (Me.WindowState = vbMinimized) Then
        Me.Visible = False
        bln_minimized = True
        Exit Sub
    ElseIf (Me.WindowState = vbMaximized) Then
        For i = 0 To JCToolbar1.Count - 1
            JCToolbar1(i).Left = JCToolbar1(i).Tag
            If IniWidth(i) < Me.ScaleWidth Then 'toolbar width is less than form scalewidth
                JCToolbar1(i).Width = IniWidth(i)
            Else    'toolbar width is greater than form scalewidth
                JCToolbar1(i).Width = Me.ScaleWidth
            End If
        Next i
    Else    'not minimized
        For i = 0 To JCToolbar1.Count - 1
            If (Me.ScaleWidth - IniWidth(i)) > 0 And (Me.ScaleWidth - IniWidth(i)) < JCToolbar1(i).Tag Then
                JCToolbar1(i).Left = Me.ScaleWidth - IniWidth(i)
            Else
                ResizeToolbar JCToolbar1(i), IniWidth(i)
            End If
        Next i
    End If
End Sub

Private Sub JCToolbar1_ButtonClick(Index As Integer, btnIndex As Integer, sKey As String, iLeft As Integer, iTop As Integer, iWidth As Integer, iHeight As Integer, blnVisible As Boolean)
    Select Case sKey
        Case "Full screen"
            Me.WindowState = vbMaximized
            JCToolbar1(1).ChangeBtnProperty jcCaption, 6, "Restore"
            JCToolbar1(1).ChangeBtnProperty jckey, 6, "Restore"
            JCToolbar1(1).ChangeBtnProperty jcTooltip, 6, "Restore"
            IniWidth(1) = JCToolbar1(1).Width
        Case "Restore"
            JCToolbar1(1).ChangeBtnProperty jcCaption, 6, "Full screen"
            JCToolbar1(1).ChangeBtnProperty jckey, 6, "Full screen"
            JCToolbar1(1).ChangeBtnProperty jcTooltip, 6, "Full screen"
            IniWidth(1) = JCToolbar1(1).Width
            Me.WindowState = vbNormal
        Case "Exit"
            Unload Me
        Case "left"
            JCToolbar1(2).ChangeBtnProperty jcValue, 4, False
            JCToolbar1(2).ChangeBtnProperty jcValue, 5, False
        Case "center"
            JCToolbar1(2).ChangeBtnProperty jcValue, 3, False
            JCToolbar1(2).ChangeBtnProperty jcValue, 5, False
        Case "right"
            JCToolbar1(2).ChangeBtnProperty jcValue, 3, False
            JCToolbar1(2).ChangeBtnProperty jcValue, 4, False
        Case "tlbrOpt"
            JCToolbar1(Index).ChangeBtnProperty jcState, btnIndex, STA_SELECTED
            selection = 2
            If blnVisible Then
                Me.PopupMenu mnuoption(0), , JCToolbar1(Index).Left + iLeft, JCToolbar1(Index).Top + iHeight
            Else
                Me.PopupMenu mnuoption(0), vbPopupMenuRightAlign, JCToolbar1(Index).Left + iLeft, JCToolbar1(Index).Top + JCToolbar1(Index).Height
            End If
            JCToolbar1(Index).ChangeBtnProperty jcState, btnIndex, STA_NORMAL
    End Select
End Sub

Private Sub JCToolbar1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        TimMove.Enabled = True
    End If
    selection = Index
End Sub

Private Sub JCToolbar1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    NeoX = JCToolbar1(Index).Left + 15 * X
    If NeoX < 0 Then
        NeoX = 0
    ElseIf NeoX > Me.ScaleWidth - JCToolbar1(Index).Width Then
        NeoX = Me.ScaleWidth - JCToolbar1(Index).Width
    End If
End Sub

Private Sub JCToolbar1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    TimMove.Enabled = False
End Sub

Private Sub JCToolbar1_ThemeColorClick(Index As Integer, themeIndex As Integer)
    Select Case themeIndex
        Case 0
            'MsgBox "blue"
        Case 1
            'MsgBox "silver"
        Case 2
            'MsgBox "olive"
        Case 3
            'MsgBox "visual studio"
        Case 4
            'MsgBox "norton"
        Case 5
            'MsgBox "autodetect"
    End Select
    'Unload Me
End Sub


Private Sub mnushow_Click(Index As Integer)
    Dim i As Integer
    Select Case Index
        Case 0  'hide toolbar
            JCToolbar1(selection).Visible = False
        Case 1 'show toolbar
            JCToolbar1(selection).Visible = True
    End Select
    AlignToolbars
End Sub

Private Sub MnuTheme_Click(Index As Integer)
    JCToolbar1(selection).ThemeColor = Index
End Sub

Private Sub TimMove_Timer()
    JCToolbar1(selection).Left = NeoX
    JCToolbar1(selection).Tag = JCToolbar1(selection).Left
End Sub

'resize toolbar when form is resizing
Private Sub ResizeToolbar(ToolBar As Object, IniWidth As Integer)
    If Me.ScaleWidth < (ToolBar.Left + ToolBar.Width) Then
        If ToolBar.Left > 0 Then ToolBar.Left = 0
        If (Me.ScaleWidth - ToolBar.Left) > 450 Then
            ToolBar.Width = Me.ScaleWidth - ToolBar.Left
        End If
    Else
        If (ToolBar.Left + IniWidth) < Me.ScaleWidth Then
            ToolBar.Width = IniWidth
        Else
            ToolBar.Width = Me.ScaleWidth - ToolBar.Left
        End If
    End If
End Sub

Private Sub AddToolbar()
    Dim i As Integer

    i = JCToolbar1.Count
    Load JCToolbar1(i)
    
    With JCToolbar1(i)
        .Width = 500
    End With
    ReDim Preserve IniWidth(i + 1)
    IniWidth(i) = JCToolbar1(i).Width
    JCToolbar1(i).Top = JCToolbar1(i - 1).Top + JCToolbar1(i - 1).Height
    JCToolbar1(i).Left = 0
    If i > 5 Then
        JCToolbar1(i).ThemeColor = i - 6
    Else
        JCToolbar1(i).ThemeColor = i - 1
    End If
    JCToolbar1(i).Tag = JCToolbar1(i).Left
    
    JCToolbar1(i).Visible = True
    CmdAddBtn_Click
End Sub

Private Sub RemoveToolbar()
    Dim i As Integer
    i = JCToolbar1.Count - 1
    Unload JCToolbar1(i)
    inum = 1
End Sub

Private Sub PaintFrame(Pic As PictureBox, p_Caption As String)
    Dim p_LenghOfCaption As Long, p_HeightOfCaption As Long, p_offset As Long, R As RECT
    
    Pic.ScaleMode = vbPixels
    
    p_LenghOfCaption = Pic.TextWidth(p_Caption)
    p_HeightOfCaption = Pic.TextHeight(p_Caption)
    p_offset = 2
    
    'define border color
    Pic.ForeColor = RGB(195, 195, 195)
    
    'paint rounded rectangle
    RoundRect Pic.hDC, 0&, p_offset + p_HeightOfCaption / 2, Pic.ScaleWidth - 1, Pic.ScaleHeight - 1, 8&, 8&
    
    'define caption rectangle
    SetRect R, p_offset + 12, 0, 12 + p_LenghOfCaption + 2 * p_offset, p_offset + p_HeightOfCaption
    
    'Set line color
    Pic.ForeColor = Pic.BackColor
    MoveToEx Pic.hDC, 12, p_offset + p_HeightOfCaption / 2, lpp
    LineTo Pic.hDC, 12 + p_LenghOfCaption + 2 * p_offset, p_offset + p_HeightOfCaption / 2
    
    'Set text color
    Pic.ForeColor = &HCF3603
    
    'Draw frame caption
    DrawTextEx Pic.hDC, p_Caption, Len(p_Caption), R, DT_LEFT Or DT_SINGLELINE Or DT_VCENTER, ByVal 0&

    Pic.ScaleMode = vbTwips
    Pic.Picture = Pic.Image
End Sub

Private Sub AlignToolbars()
    Dim i As Integer, iPrev As Integer
    iPrev = 0
    For i = 0 To JCToolbar1.Count - 1
        If blnFormLoaded = False Then
            IniWidth(i) = JCToolbar1(i).Width
            JCToolbar1(i).Left = 0
            JCToolbar1(i).ThemeColor = i - 1
            JCToolbar1(i).Tag = JCToolbar1(i).Left
            If i = 1 Then JCToolbar1(i).Top = 10
            If i > 1 Then JCToolbar1(i).Top = JCToolbar1(i - 1).Top + JCToolbar1(i - 1).Height
            If i > 0 Then JCToolbar1(i).Visible = True
        Else
            If JCToolbar1(i).Visible = True Then
                IniWidth(i) = JCToolbar1(i).Width
                JCToolbar1(i).Left = 0
                JCToolbar1(i).Tag = JCToolbar1(i).Left
                If i = 1 Or iPrev = 0 Then JCToolbar1(i).Top = 10
                If iPrev > 0 Then JCToolbar1(i).Top = JCToolbar1(iPrev).Top + JCToolbar1(iPrev).Height
                iPrev = i
            End If
        End If
    Next i

End Sub
