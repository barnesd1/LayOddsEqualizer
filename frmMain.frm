VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lay Odds Equalizer"
   ClientHeight    =   4965
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   6675
   FillStyle       =   6  'Cross
   ForeColor       =   &H00800080&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   6675
   Begin VB.Frame fraEW 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   2880
      TabIndex        =   35
      Top             =   520
      Visible         =   0   'False
      Width           =   3600
      Begin VB.OptionButton optPlaceLayPc 
         BackColor       =   &H00E0E0E0&
         Caption         =   "5"
         Height          =   255
         Index           =   2
         Left            =   2460
         TabIndex        =   49
         ToolTipText     =   "Right-Click to custom"
         Top             =   1920
         Width           =   530
      End
      Begin VB.OptionButton optPlaceLayPc 
         BackColor       =   &H00E0E0E0&
         Caption         =   "2"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   48
         ToolTipText     =   "Right-Click to custom"
         Top             =   1920
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton optPlaceLayPc 
         BackColor       =   &H00E0E0E0&
         Caption         =   "0"
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   47
         ToolTipText     =   "Fixed at Zero"
         Top             =   1920
         Width           =   375
      End
      Begin VB.TextBox txtPlaceLayStake 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1260
         TabIndex        =   44
         Top             =   2400
         Width           =   1270
      End
      Begin VB.TextBox txtPlaceLayOdds 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1260
         TabIndex        =   43
         Top             =   1320
         Width           =   1270
      End
      Begin VB.TextBox txtPlaceBackOdds 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2580
         TabIndex        =   40
         Top             =   720
         Width           =   795
      End
      Begin VB.ComboBox cboTerms 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtPlaceStake 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """£""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1260
         TabIndex        =   37
         Top             =   120
         Width           =   2115
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Place Lay %ge"
         Height          =   555
         Left            =   420
         TabIndex        =   46
         Top             =   1860
         Width           =   735
      End
      Begin VB.Label lblPlaceLayStake 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Place Lay Stake"
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   180
         TabIndex        =   45
         Top             =   2400
         Width           =   1005
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Place Lay Odds"
         Height          =   585
         Left            =   120
         TabIndex        =   42
         Top             =   1320
         Width           =   1005
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Calculated Place Odds"
         Height          =   390
         Left            =   1500
         TabIndex        =   41
         Top             =   720
         Width           =   1005
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Place Terms"
         Height          =   375
         Left            =   120
         TabIndex        =   38
         Top             =   720
         Width           =   465
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblPlaceStake 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Place Stake / Total Stake "
         Height          =   435
         Left            =   60
         TabIndex        =   36
         Top             =   120
         Width           =   1035
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CheckBox chkEW 
      Appearance      =   0  'Flat
      Caption         =   "E/ W"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   390
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   240
      Width           =   390
   End
   Begin VB.Frame fraBackPc 
      BorderStyle     =   0  'None
      Height          =   370
      Left            =   4660
      TabIndex        =   30
      Top             =   135
      Width           =   1860
      Begin VB.OptionButton optBackPc 
         Caption         =   "5"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   33
         ToolTipText     =   "Right-Click to custom"
         Top             =   120
         Width           =   545
      End
      Begin VB.OptionButton optBackPc 
         Caption         =   "0"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   32
         ToolTipText     =   "Fixed at Zero"
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Back %ge"
         Height          =   255
         Left            =   0
         TabIndex        =   31
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.PictureBox picDutch 
      AutoSize        =   -1  'True
      Height          =   1260
      Left            =   45
      Picture         =   "frmMain.frx":0442
      ScaleHeight     =   1200
      ScaleWidth      =   330
      TabIndex        =   29
      Top             =   1230
      Width           =   390
   End
   Begin VB.Frame fraFull 
      BackColor       =   &H00C0C0C0&
      Caption         =   "   Under Lay                                   Quick %ge  Slider                        Over Lay"
      ClipControls    =   0   'False
      ForeColor       =   &H00800080&
      Height          =   1215
      Left            =   120
      TabIndex        =   19
      Top             =   3600
      Width           =   6375
      Begin VB.CommandButton cmdPredictLose 
         Caption         =   "Predict &Lose"
         Height          =   560
         Left            =   4920
         TabIndex        =   53
         ToolTipText     =   "Works out best Lay bet"
         Top             =   1020
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.OptionButton optUO 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Both Stakes"
         Height          =   255
         Index           =   2
         Left            =   6180
         TabIndex        =   52
         Top             =   840
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.OptionButton optUO 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Place Stake"
         Height          =   255
         Index           =   1
         Left            =   6180
         TabIndex        =   51
         Top             =   540
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.OptionButton optUO 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Win Stake"
         Height          =   255
         Index           =   0
         Left            =   6180
         TabIndex        =   50
         Top             =   240
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.CommandButton cmdClearHistory 
         Caption         =   "CLEAR"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1620
         TabIndex        =   28
         Top             =   990
         Width           =   735
      End
      Begin VB.CommandButton cmdLoseZeroBack 
         Caption         =   "Lose 0 on &Back"
         Height          =   560
         Left            =   240
         TabIndex        =   24
         ToolTipText     =   "Works out best Lay bet"
         Top             =   375
         Width           =   1230
      End
      Begin VB.CommandButton cmdLoseZeroLay 
         Caption         =   "Lose 0 on &Lay"
         Height          =   560
         Left            =   4920
         TabIndex        =   23
         ToolTipText     =   "Works out best Lay bet"
         Top             =   375
         Width           =   1230
      End
      Begin VB.Label lblO 
         BackStyle       =   0  'Transparent
         Caption         =   ">5%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   2
         Left            =   4500
         TabIndex        =   55
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblU 
         BackStyle       =   0  'Transparent
         Caption         =   "<5%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   54
         Top             =   240
         Width           =   375
      End
      Begin VB.Image imgHistory 
         BorderStyle     =   1  'Fixed Single
         Height          =   390
         Index           =   3
         Left            =   4140
         MousePointer    =   8  'Size NW SE
         Picture         =   "frmMain.frx":0A90
         Stretch         =   -1  'True
         Top             =   600
         Width           =   570
      End
      Begin VB.Image imgHistory 
         BorderStyle     =   1  'Fixed Single
         Height          =   390
         Index           =   2
         Left            =   3600
         MousePointer    =   7  'Size N S
         Picture         =   "frmMain.frx":1079
         Stretch         =   -1  'True
         Top             =   600
         Width           =   540
      End
      Begin VB.Image imgHistory 
         BorderStyle     =   1  'Fixed Single
         Height          =   390
         Index           =   1
         Left            =   3060
         MousePointer    =   9  'Size W E
         Picture         =   "frmMain.frx":1505
         Stretch         =   -1  'True
         Top             =   600
         Width           =   540
      End
      Begin VB.Image imgHistory 
         BorderStyle     =   1  'Fixed Single
         Height          =   390
         Index           =   0
         Left            =   2520
         MousePointer    =   6  'Size NE SW
         Picture         =   "frmMain.frx":19C3
         Stretch         =   -1  'True
         Top             =   600
         Width           =   540
      End
      Begin VB.Label lblHistory 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Bet History"
         ForeColor       =   &H00800080&
         Height          =   435
         Left            =   1570
         TabIndex        =   27
         Top             =   600
         Width           =   850
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblO 
         BackStyle       =   0  'Transparent
         Caption         =   ">1%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   1
         Left            =   4020
         TabIndex        =   26
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblU 
         BackStyle       =   0  'Transparent
         Caption         =   "<0.5%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   2520
         TabIndex        =   25
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblO 
         BackStyle       =   0  'Transparent
         Caption         =   ">0.5%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   3360
         TabIndex        =   22
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblEqual 
         BackStyle       =   0  'Transparent
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Left            =   3120
         TabIndex        =   21
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblU 
         BackStyle       =   0  'Transparent
         Caption         =   "<1%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   20
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.TextBox txtLayDesc 
      Height          =   285
      Left            =   180
      TabIndex        =   18
      ToolTipText     =   "Lay Desc"
      Top             =   2700
      Width           =   1155
   End
   Begin VB.TextBox txtBackDesc 
      Height          =   285
      Left            =   180
      TabIndex        =   17
      ToolTipText     =   "Back Desc"
      Top             =   900
      Width           =   1155
   End
   Begin VB.OptionButton optLayPc 
      Caption         =   "5"
      Height          =   255
      Index           =   2
      Left            =   4080
      TabIndex        =   16
      ToolTipText     =   "Right-Click to custom"
      Top             =   240
      Value           =   -1  'True
      Width           =   530
   End
   Begin VB.OptionButton optLayPc 
      Caption         =   "2"
      Height          =   255
      Index           =   1
      Left            =   3480
      TabIndex        =   15
      ToolTipText     =   "Right-Click to custom"
      Top             =   240
      Width           =   615
   End
   Begin VB.OptionButton optLayPc 
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   13
      ToolTipText     =   "Fixed at Zero"
      Top             =   240
      Width           =   375
   End
   Begin VB.CheckBox chkStakeNotReturned 
      Caption         =   "Check1"
      Height          =   255
      Left            =   1950
      TabIndex        =   12
      Top             =   240
      Width           =   255
   End
   Begin VB.TextBox txtResult 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   6400
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   545
      Width           =   3495
   End
   Begin VB.CommandButton cmdNECalc 
      Caption         =   "Not Equal &Calc"
      Height          =   560
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Calculates using values you specify"
      Top             =   3000
      Width           =   1230
   End
   Begin VB.TextBox txtLayStake 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1500
      TabIndex        =   3
      Top             =   2400
      Width           =   1270
   End
   Begin VB.TextBox txtLayOdds 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1500
      TabIndex        =   2
      Top             =   1815
      Width           =   1270
   End
   Begin VB.TextBox txtBackOdds 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1500
      TabIndex        =   1
      Top             =   1200
      Width           =   1270
   End
   Begin VB.TextBox txtBackStake 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """£""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1500
      TabIndex        =   0
      Top             =   615
      Width           =   1270
   End
   Begin VB.CommandButton cmdEqualize 
      Caption         =   "&Equalize"
      Height          =   560
      Left            =   1560
      TabIndex        =   4
      ToolTipText     =   "Works out best Lay bet"
      Top             =   3000
      Width           =   1230
   End
   Begin VB.Shape shpBox 
      BorderColor     =   &H00800080&
      BorderWidth     =   10
      FillColor       =   &H00800080&
      Height          =   4860
      Left            =   60
      Top             =   60
      Width           =   6555
   End
   Begin VB.Label Label6 
      Caption         =   "Lay %ge"
      Height          =   375
      Left            =   2280
      TabIndex        =   14
      Top             =   240
      Width           =   735
   End
   Begin VB.Label lblLayStake 
      Alignment       =   1  'Right Justify
      Caption         =   "Win Lay Stake"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2460
      Width           =   1215
   End
   Begin VB.Label lblLayOdds 
      Alignment       =   1  'Right Justify
      Caption         =   "Win Lay Odds"
      Height          =   435
      Left            =   480
      TabIndex        =   9
      Top             =   1860
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblBackOdds 
      Alignment       =   1  'Right Justify
      Caption         =   "Win Back Odds"
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   1260
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Stake Not Returned"
      Height          =   195
      Left            =   450
      TabIndex        =   7
      Top             =   240
      Width           =   1425
   End
   Begin VB.Label lblBackStake 
      Alignment       =   1  'Right Justify
      Caption         =   "Win Back Stake "
      Height          =   435
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuNew 
      Caption         =   "&New"
   End
   Begin VB.Menu mnuReset 
      Caption         =   "&Reset"
   End
   Begin VB.Menu mnuComm 
      Caption         =   "&Setup"
      Begin VB.Menu mnuBFDiscount 
         Caption         =   "Betfair Commission %  Discount Explained"
      End
      Begin VB.Menu mnuBackPc 
         Caption         =   "Override Back Commission %"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLayPc 
         Caption         =   "Override Lay Commission %"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCurrSymbol 
         Caption         =   "Set Currency Symbol"
      End
      Begin VB.Menu mnuLn2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCorporateColours 
         Caption         =   "Corporate Colours"
      End
      Begin VB.Menu mnuln4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuResetClears 
         Caption         =   "Reset Clears Bet History"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuln3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMaxI 
         Caption         =   "Max Calc Iterations"
         Visible         =   0   'False
         Begin VB.Menu mnuMax 
            Caption         =   "10,000"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuMax 
            Caption         =   "25,000"
            Index           =   1
         End
         Begin VB.Menu mnuMax 
            Caption         =   "50,000"
            Index           =   2
         End
         Begin VB.Menu mnuMax 
            Caption         =   "100,000"
            Index           =   3
         End
         Begin VB.Menu mnuMax 
            Caption         =   "1,000,000"
            Index           =   4
         End
      End
   End
   Begin VB.Menu mnuCalc 
      Caption         =   "Calc&ulators"
      Begin VB.Menu mnuStandard 
         Caption         =   "Standard"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuDutch 
         Caption         =   "Dutch"
      End
      Begin VB.Menu mnuEW 
         Caption         =   "Each Way"
      End
   End
   Begin VB.Menu mnuUpgrade 
      Caption         =   "&Unlock Full Features"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About Lay Odds Equalizer"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mbLoading As Boolean
Dim mbCalcing As Boolean
Dim sCompare As String



Private Sub cboTerms_Click()
If mbLoading Then Exit Sub
Call txtBackOdds_LostFocus
End Sub



Private Sub chkEW_Click()
Const sCaption As String = "   Under Lay                                   Quick %ge  Slider                        Over Lay"
Const sCap2 As String = "                     Apply U/O-Lay to"
Dim i As Integer
For i = 0 To 2
    optUO(i).Visible = chkEW.Value
    optUO(i).Left = 6400
Next
If chkEW.Value Then
    bEWMode = True
    frmMain.Width = 10230
    txtResult.Left = 6485
    shpBox.Width = 10000
    fraFull.Width = 9810
    fraFull.Caption = sCaption + sCap2
    fraEW.Visible = True
    mnuEW.Checked = True
    mnuStandard.Checked = False
    cmdPredictLose.Top = 375
    cmdPredictLose.Left = 8000
    cmdPredictLose.Visible = True
    chkEW.BackColor = &HC0C0C0
Else
    bEWMode = False
    frmMain.Width = 6765
    txtResult.Left = 3000
    shpBox.Width = 6555
    fraFull.Width = 6375
    fraFull.Caption = sCaption
    fraEW.Visible = False
    mnuEW.Checked = False
    mnuStandard.Checked = True
    cmdPredictLose.Visible = False
    chkEW.BackColor = &H8000000F
End If
If Not mbLoading Then Call UpdateCurrency
End Sub

Private Sub chkEW_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtBackStake.SetFocus
End Sub

Private Sub chkStakeNotReturned_Click()
If chkStakeNotReturned.Value Then
    bStakeNotReturned = True
Else
    bStakeNotReturned = False
End If
End Sub



Private Sub cmdClearHistory_Click()
    iHistoryUsed = 0
    lblHistory = "Bet History"
    cmdEqualize.SetFocus
End Sub

Private Sub cmdEqualize_Click()
If BackStake = 0 Or LayOdds = 0 Or BackOdds = 0 Then
    Beep
    txtLayStake.Text = Format(0, "0.00")
    Exit Sub
End If
If bEWMode Then
    If PlaceLayOdds = 0 Then
        Beep
        txtPlaceLayOdds.SetFocus
        Exit Sub
    End If
End If
Me.MousePointer = vbHourglass
If chkStakeNotReturned.Value Then
    StakeNotReturned = BackStake
Else
    StakeNotReturned = 0
End If
LayStake = BackStake 'seed
Call calc("Equal")
If bEWMode Then Call EWcalc("Equal")
txtResult.Text = IIf(bEWMode, EWresult(), result())
sCompare = compare()
txtLayStake = Format(LayStake, "0.00")
txtPlaceLayStake.Text = Format(PlaceLayStake, "0.00")
If Not mbCalcing Then Call AddToHistory
Me.MousePointer = vbDefault
End Sub

Private Sub cmdLoseZeroBack_Click()
'If picUpgrade.Visible Then mnuAbout_Click: Exit Sub
If BackStake = 0 Or LayOdds = 0 Or BackOdds = 0 Then Beep: Exit Sub
Dim lIterations As Long

If bEWMode Then
    mbCalcing = True
    cmdEqualize.Value = True
    If BetfairBackCost + PlaceBetfairBackCost < 0 And BetfairLayCost + PlaceBetfairBackCost < 0 And BetfairLayCost + PlaceBetfairLayCost < 0 Then
        MsgBox "You have no edge on this bet", vbExclamation, "No way to edge"
        mbCalcing = False
        Exit Sub
    End If

    optUO(0).Value = True
    Do
        Call lblU_Click(0)
        lIterations = lIterations + 1
    Loop While BetfairLayCost + PlaceBetfairLayCost > 0.105 And lIterations < 10000
    mbCalcing = False
    
Else
    mbCalcing = True
    cmdEqualize.Value = True
    
    If BetfairBackCost = 0 Then Beep: mbCalcing = False: Exit Sub
    If BetfairBackCost > 0.01 Then
        Do
            LayStake = LayStake + 0.01
            txtLayStake.Text = Format(LayStake, "0.00")
            cmdNECalc.Value = True
            lIterations = lIterations + 1
        Loop While Val(Format(BetfairBackCost, "0.00")) > 0 And lIterations < 10000
    Else
        Do
            LayStake = LayStake - 0.01
            txtLayStake.Text = Format(LayStake, "0.00")
            cmdNECalc.Value = True
            lIterations = lIterations + 1
        Loop While Val(Format(BetfairBackCost, "0.00")) < 0 And lIterations < 10000
    End If
    
    mbCalcing = False
End If
Call AddToHistory
End Sub

Private Sub cmdNECalc_Click()
If BackStake = 0 Or LayOdds = 0 Or BackOdds = 0 Then Beep: Exit Sub
'If Not bEWMode And LayStake = 0 Then Beep: Exit Sub
If bEWMode And PlaceLayOdds = 0 Then Beep: txtPlaceLayOdds.SetFocus: Exit Sub
Me.MousePointer = vbHourglass
If chkStakeNotReturned.Value Then
    StakeNotReturned = BackStake
Else
    StakeNotReturned = 0
End If
Call calc("NotEqual")
If bEWMode Then Call EWcalc("NotEqual")
txtResult.Text = IIf(bEWMode, EWresult(), result())
txtLayStake.Text = Format(LayStake, "0.00")
txtPlaceLayStake.Text = Format(PlaceLayStake, "0.00")
If Not mbCalcing Then Call AddToHistory
Me.MousePointer = vbDefault
End Sub

Private Sub cmdLoseZeroLay_Click()
'If picUpgrade.Visible Then mnuAbout_Click: Exit Sub
If BackStake = 0 Or LayOdds = 0 Or BackOdds = 0 Then Beep: Exit Sub
Dim lIterations As Long

If bEWMode Then
    mbCalcing = True
    cmdEqualize.Value = True
    If BetfairBackCost + PlaceBetfairBackCost < 0 And BetfairLayCost + PlaceBetfairBackCost < 0 And BetfairLayCost + PlaceBetfairLayCost < 0 Then
        MsgBox "You have no edge on this bet", vbExclamation, "No way to edge"
        mbCalcing = False
        Exit Sub
    End If
    optUO(1).Value = True
    Do
        Call lblU_Click(0)
        lIterations = lIterations + 1
    Loop While BetfairLayCost + PlaceBetfairLayCost > 0.105 And lIterations < 10000
    mbCalcing = False
Else

    mbCalcing = True
    cmdEqualize.Value = True
    
    If BetfairLayCost = 0 Then Beep: mbCalcing = False: Exit Sub
    
    If BetfairLayCost > 0.01 Then
        Do
            LayStake = LayStake - 0.01
            txtLayStake.Text = Format(LayStake, "0.00")
            cmdNECalc.Value = True
            lIterations = lIterations + 1
        Loop While Val(Format(BetfairLayCost, "0.00")) > 0 And lIterations < 10000
    Else
        Do
            LayStake = LayStake + 0.01
            txtLayStake.Text = Format(LayStake, "0.00")
            cmdNECalc.Value = True
            lIterations = lIterations + 1
        Loop While Val(Format(BetfairLayCost, "0.00")) < 0 And lIterations < 10000
    End If
    mbCalcing = False
End If
Call AddToHistory
End Sub

Private Sub cmdPredictLose_Click()
    Dim lIterations As Long
    mbCalcing = True
    cmdEqualize.Value = True
    If BetfairBackCost + PlaceBetfairBackCost < 0 And BetfairLayCost + PlaceBetfairBackCost < 0 And BetfairLayCost + PlaceBetfairLayCost < 0 Then
        MsgBox "You have no edge on this bet", vbExclamation, "No way to edge"
        mbCalcing = False
        Exit Sub
    End If
    optUO(1).Value = True
    Do
        lIterations = lIterations + 1
        Call lblO_Click(0)
    Loop While BetfairLayCost + PlaceBetfairBackCost > 0.105 And lIterations < 10000
    mbCalcing = False
    Call AddToHistory
End Sub

Private Sub Form_Load()
'If App.Revision = 99 Then
'    mnuUpgrade.Visible = True
'    picUpgrade.Visible = True
'Else
'    mnuUpgrade.Visible = False
'    picUpgrade.Visible = False
'End If

LayPc = 2#
BackPc = 0#
PlacePc = 2#
mnuBackPc.Caption = "Back Commission " + CStr(BackPc) + "%"
mnuLayPc.Caption = "Lay Commission " + CStr(LayPc) + "%"
chkStakeNotReturned.Value = 0
cmdEqualize.Enabled = True
cmdNECalc.Enabled = True
mbLoading = True
sCompare = ""
txtResult.Left = 3000
fraEW.BorderStyle = 0
cboTerms.AddItem "1/3"
cboTerms.AddItem "1/4"
cboTerms.AddItem "1/5"
cboTerms.ListIndex = 2
End Sub









Private Sub Form_Paint()
If mbLoading Then
    ReadSettings
    UpdateCurrency
    mbLoading = False
End If
End Sub



Private Sub Form_Unload(Cancel As Integer)
SaveSettings
End Sub





Private Sub imgHistory_Click(Index As Integer)
Dim i As Integer
For i = 1 To 300
imgHistory(Index).Appearance = 0
DoEvents
DoEvents
imgHistory(Index).Appearance = 1
Next i
'If picUpgrade.Visible Then Call mnuAbout_Click
If iHistoryUsed = 0 Then Beep: Exit Sub
' move to history item
Dim iMoveTo As Integer
Select Case Index
    Case 0:
    ' move first item
    iMoveTo = 1
    Case 1:
    ' move back one
    iMoveTo = iHistoryPosition - 1
    If iMoveTo < 1 Then iMoveTo = 1
    Case 2:
    ' move forward one
    iMoveTo = iHistoryPosition + 1
    If iMoveTo > iHistoryUsed Then iMoveTo = iHistoryUsed
    Case 3:
    ' move last item
    iMoveTo = iHistoryUsed
End Select

If (iHistoryPosition <> iMoveTo) Or txtLayStake.Text = "" Then
    iHistoryPosition = iMoveTo
    BackPc = iHistBackPc(iMoveTo)
    LayPc = iHistLayPc(iMoveTo)
    PlacePc = iHistPlacePc(iMoveTo)
    SetPcOpt
    txtBackStake.Text = ConvertOdds(Val(iHistBackStake(iMoveTo)))
    BackStake = iHistBackStake(iMoveTo)
    txtBackOdds.Text = ConvertOdds(Val(iHistBackOdds(iMoveTo)))
    BackOdds = iHistBackOdds(iMoveTo)
    txtLayStake.Text = ConvertOdds(Val(iHistLayStake(iMoveTo)))
    LayStake = iHistLayStake(iMoveTo)
    txtLayOdds.Text = ConvertOdds(Val(iHistLayOdds(iMoveTo)))
    LayOdds = iHistLayOdds(iMoveTo)
    chkStakeNotReturned.Value = Abs(bHistSNR(iMoveTo))
    Call chkStakeNotReturned_Click
    cboTerms.ListIndex = iHistTerms(iMoveTo)
    txtPlaceLayStake.Text = ConvertOdds(Val(iHistPlaceLayStake(iMoveTo)))
    PlaceLayStake = iHistPlaceLayStake(iMoveTo)
    txtPlaceLayOdds.Text = ConvertOdds(Val(iHistPlaceLayOdds(iMoveTo)))
    PlaceLayOdds = iHistPlaceLayOdds(iMoveTo)
    If bEW(iMoveTo) <> bEWMode Then
        bEWMode = bEW(iMoveTo)
        If bEWMode Then
            chkEW.Value = 1
        Else
            chkEW.Value = 0
        End If
        Call chkEW_Click
    End If
    lblHistory.Caption = "Bet History " + CStr(iHistoryPosition) + " of " + CStr(iHistoryUsed)
    mbCalcing = True
    cmdNECalc.SetFocus
    cmdNECalc.Value = True
    mbCalcing = False
End If
End Sub
Public Sub SetPcOpt()
    Dim i As Integer
    For i = 0 To 2
        If Val(optLayPc(i).Caption) = LayPc Then
            optLayPc(i).Value = True
        End If
        If i < 2 Then
            If Val(optBackPc(i).Caption) = BackPc Then
                optBackPc(i).Value = True
            End If
        End If
        If Val(optPlaceLayPc(i).Caption) = PlacePc Then
            optPlaceLayPc(i).Value = True
        End If
    Next
End Sub
Public Sub AddToHistory()
If iHistoryUsed >= 100 Then
    If MsgBox("Maximum 100 history reached, click Ok to make to clear the history and make this item 1, or Cancel to make this item 100", vbOKCancel, "History full") = vbOK Then
        iHistoryUsed = 0
    Else
        iHistoryUsed = 99
    End If
End If
Dim i As Integer
Dim iTerms As Integer
iTerms = cboTerms.ListIndex
LayStake = Format(LayStake, "0.00")
PlaceLayStake = Format(PlaceLayStake, "0.00")
If iHistoryUsed > 0 Then
    i = iHistoryUsed
    If iHistBackPc(i) = BackPc And _
        iHistLayPc(i) = LayPc And _
        iHistBackStake(i) = BackStake And _
        iHistBackOdds(i) = BackOdds And _
        iHistLayStake(i) = LayStake And _
        iHistLayOdds(i) = LayOdds And _
        bHistSNR(i) = bStakeNotReturned And _
        iHistTerms(i) = iTerms And _
        iHistPlacePc(i) = PlacePc And _
        iHistPlaceLayStake(i) = PlaceLayStake And _
        iHistPlaceLayOdds(i) = PlaceLayOdds And _
        bEW(i) = bEWMode Then
        Exit Sub
    End If
End If
' add to history
iHistoryUsed = iHistoryUsed + 1
i = iHistoryUsed
iHistoryPosition = i
iHistBackPc(i) = BackPc
iHistLayPc(i) = LayPc
iHistBackStake(i) = Val(txtBackStake.Text)
iHistBackOdds(i) = Val(txtBackOdds.Text)
iHistLayStake(i) = Val(txtLayStake.Text)
iHistLayOdds(i) = Val(txtLayOdds.Text)
bHistSNR(i) = bStakeNotReturned
iHistTerms(i) = iTerms
iHistPlacePc(i) = PlacePc
iHistPlaceLayStake(i) = PlaceLayStake
iHistPlaceLayOdds(i) = PlaceLayOdds
bEW(i) = bEWMode
lblHistory.Caption = "Bet History " + CStr(iHistoryPosition) + " of " + CStr(iHistoryUsed)
End Sub



Private Sub lblEqual_Click()
'If picUpgrade.Visible Then picUpgrade_Click: Exit Sub

cmdEqualize.SetFocus
cmdEqualize.Value = True

End Sub





Private Sub lblO_Click(Index As Integer)
'If picUpgrade.Visible Then picUpgrade_Click: Exit Sub
Dim mult As Single
Select Case (Index)
Case 0: mult = 1.005
Case 1: mult = 1.01
Case 2: mult = 1.05
End Select

If optUO(0).Value Or optUO(2).Value Or Not bEWMode Then
txtLayStake.Text = ConvertOdds(LayStake * mult)
LayStake = Val(txtLayStake.Text)
End If
If bEWMode And (optUO(1).Value Or optUO(2).Value) Then
txtPlaceLayStake.Text = ConvertOdds(PlaceLayStake * mult)
PlaceLayStake = Val(txtPlaceLayStake.Text)
End If

cmdNECalc.SetFocus
cmdNECalc.Value = True
End Sub

Private Sub lblU_Click(Index As Integer)
'If picUpgrade.Visible Then picUpgrade_Click: Exit Sub
Dim mult As Single
Select Case (Index)
Case 0: mult = 0.995
Case 1: mult = 0.99
Case 2: mult = 0.95
End Select

If optUO(0).Value Or optUO(2).Value Or Not bEWMode Then
txtLayStake.Text = ConvertOdds(LayStake * mult)
LayStake = Val(txtLayStake.Text)
End If
If bEWMode And (optUO(1).Value Or optUO(2).Value) Then
txtPlaceLayStake.Text = ConvertOdds(PlaceLayStake * mult)
PlaceLayStake = Val(txtPlaceLayStake.Text)
End If

cmdNECalc.SetFocus
cmdNECalc.Value = True
End Sub


Private Sub mnuAbout_Click()
Dim sMsg As String
Dim iResult As Integer
sMsg = "Lay Odds Equalizer v" + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision)
If App.Revision = 99 Then
    sMsg = sMsg + " TRIAL"
Else
    sMsg = sMsg + " FULL"
End If
sMsg = sMsg + " Edition"
sMsg = sMsg + vbCrLf + "Written by David J. Barnes, London, UK" + vbCrLf + "Released under public license at GitHub" + vbCrLf + vbCrLf + "Note: use of this application is entirely at your own risk, gamble responsibly" + vbCrLf + vbCrLf + "Click Yes to Visit https://github.com/barnesd1/LayOddsEqualizer" + vbCrLf + "Click No to Donate £5.00 towards development of this application at http://paypal.me/LayOddsEqualizer/5"
iResult = MsgBox(sMsg, vbInformation + vbYesNoCancel, "About Lay Odds Equalizer")
If iResult = vbYes Then Call ShellExecute(Me.hwnd, "open", "https://github.com/barnesd1/LayOddsEqualizer", "", "", 4)
If iResult = vbNo Then Call ShellExecute(Me.hwnd, "open", "http://paypal.me/LayOddsEqualizer/5", "", "", 4)
End Sub

Private Sub mnuBackPc_Click()
BackPc = Val(InputBox("Back Commission Percent ?", , BackPc))
mnuBackPc.Caption = "Back Commission " + CStr(BackPc) + "%"
End Sub

Private Sub mnuBFDiscount_Click()
Dim BF As Double
Dim BFDiscount As Double
Dim BFLayCom As Double
BF = Val(InputBox("First Confirm Betfair Standard Commission Percentage (Normally 5%)?", "Betfair Discount Calculator", "5.0"))
If BF = 0 Then Beep: Exit Sub
BFDiscount = Val(InputBox("Enter Betfair Discount Percentage on your Account?", "Betfair Discount Calculator"))
If BFDiscount = 0 Then Beep: Exit Sub
BFLayCom = BF * (100 - BFDiscount) / 100
MsgBox "You will be paying " + CStr(BF) + " * (100 - " + CStr(BFDiscount) + ") / 100 so you should use" + vbCrLf + CStr(BFLayCom) + "% in your Betfair lay calculations", vbExclamation + vbOKOnly
End Sub

Private Sub mnuCorporateColours_Click()
Const Purple As Long = &H800080
Const Grey As Long = &HC0C0C0
Dim i As Integer
If mnuCorporateColours.Checked Then
    mnuCorporateColours.Checked = False
    bCorporate = False
    shpBox.BorderColor = Purple
    shpBox.FillColor = Purple
    fraFull.ForeColor = Purple
    fraFull.BackColor = Grey
    lblEqual.ForeColor = Purple
    lblHistory.BackColor = Grey
    lblHistory.ForeColor = Purple
    For i = 0 To 2
    lblU(i).ForeColor = vbBlue
    lblO(i).ForeColor = vbRed
    Next
Else
    mnuCorporateColours.Checked = True
    bCorporate = True
    shpBox.BorderColor = vbButtonFace
    shpBox.FillColor = vbButtonFace
    fraFull.ForeColor = vbBlack
    fraFull.BackColor = vbButtonFace
    lblEqual.ForeColor = vbBlack
    lblHistory.ForeColor = vbBlack
    lblHistory.BackColor = vbButtonFace
    For i = 0 To 2
    lblU(i).ForeColor = vbBlack
    lblO(i).ForeColor = vbBlack
    Next
End If
If mnuCorporateColours.Checked Then
    SaveSetting App.EXEName, "Colours", "Corporate", "YES"
Else
    SaveSetting App.EXEName, "Colours", "Corporate", "NO"
End If
End Sub

Private Sub mnuCurrSymbol_Click()
Dim sCurrSymbol As String
sCurrSymbol = InputBox("Enter Currency Symbol ?", , CurrSymbol)
If Len(sCurrSymbol) > 3 Then MsgBox "Maximum currency symbol length is 3": Exit Sub
If sCurrSymbol <> CurrSymbol Then
   CurrSymbol = sCurrSymbol
   UpdateCurrency
   
End If
End Sub
Sub UpdateCurrency()
    Dim s As String
    If bEWMode Then
        s = "Win "
        cmdLoseZeroBack.Caption = "Predict &Win"
        cmdLoseZeroLay.Caption = "Predict &Place"
        lblBackStake.Caption = s + "Back Stake " + CurrSymbol
    Else
        cmdLoseZeroBack.Caption = "Lose " + CurrSymbol + "0 on &Back"
        cmdLoseZeroLay.Caption = "Lose " + CurrSymbol + "0 on &Lay"
        lblBackStake.Caption = "        Back Stake " + CurrSymbol
    End If
    
    lblLayStake.Caption = s + "Lay Stake " + CurrSymbol
    lblBackOdds.Caption = s + "Back Odds"
    lblLayOdds.Caption = s + "Lay Odds"
    lblPlaceStake.Caption = "Place Stake / Total Stake " + CurrSymbol
    lblPlaceLayStake.Caption = "Place Lay Stake " + CurrSymbol
End Sub

Private Sub mnuDutch_Click()
    Call picDutch_Click
End Sub

Private Sub mnuEW_Click()
    If chkEW.Value Then
      chkEW.Value = False
    Else
      chkEW.Value = 1
    End If
    Call chkEW_Click
End Sub

Private Sub mnuExit_Click()
    SaveSettings
    End
End Sub

Private Sub mnuLayPc_Click()
LayPc = Val(InputBox("Lay Commission Percent ?", , LayPc))
mnuLayPc.Caption = "Lay Commission " + CStr(LayPc) + "%"
End Sub




Private Sub SaveSettings()
SaveSetting App.EXEName, "Settings", "Currency", CurrSymbol
SaveSetting App.EXEName, "Settings", "LayOdds1", optLayPc(1).Caption
SaveSetting App.EXEName, "Settings", "LayOdds2", optLayPc(2).Caption
SaveSetting App.EXEName, "Settings", "PlaceLayOdds1", optPlaceLayPc(1).Caption
SaveSetting App.EXEName, "Settings", "PlaceLayOdds2", optPlaceLayPc(2).Caption
SaveSetting App.EXEName, "Settings", "BackOdds1", optBackPc(1).Caption
SaveSetting App.EXEName, "Position", "Top", CStr(frmMain.Top)
SaveSetting App.EXEName, "Position", "Left", CStr(frmMain.Left)
SaveSetting App.EXEName, "Settings", "Iterations", CStr(lMaxIterations)
SaveSetting App.EXEName, "Settings", "DefaultLayPercent", CStr(LayPc)
SaveSetting App.EXEName, "Settings", "DefaultBackPercent", CStr(BackPc)
SaveSetting App.EXEName, "Settings", "DefaultPlaceLayPercent", CStr(PlacePc)

If bEWMode Then
    SaveSetting App.EXEName, "Settings", "Mode", "EW"
Else
    SaveSetting App.EXEName, "Settings", "Mode", "Normal"
End If
End Sub
Private Sub ReadSettings()
CurrSymbol = Left(GetSetting(App.EXEName, "Settings", "Currency", "£"), 3)

Dim sMode As String
sMode = GetSetting(App.EXEName, "Settings", "Mode", "Normal")

optLayPc(1).Caption = CStr(Val(GetSetting(App.EXEName, "Settings", "LayOdds1")))
optLayPc(2).Caption = CStr(Val(GetSetting(App.EXEName, "Settings", "LayOdds2")))
optPlaceLayPc(1).Caption = CStr(Val(GetSetting(App.EXEName, "Settings", "PlaceLayOdds1")))
optPlaceLayPc(2).Caption = CStr(Val(GetSetting(App.EXEName, "Settings", "PlaceLayOdds2")))
optBackPc(1).Caption = CStr(Val(GetSetting(App.EXEName, "Settings", "BackOdds1")))
lMaxIterations = Val(GetSetting(App.EXEName, "Settings", "Iterations"))
Select Case lMaxIterations
Case 10000:
    Call mnuMax_Click(0)
Case 25000:
    Call mnuMax_Click(1)
Case 50000:
    Call mnuMax_Click(2)
Case 100000:
    Call mnuMax_Click(3)
Case 1000000:
    Call mnuMax_Click(4)
Case Else:
    Call mnuMax_Click(4)
End Select

Dim dDefaultLayPercent As Double, dDefaultPlacePc As Double
dDefaultLayPercent = Val(GetSetting(App.EXEName, "Settings", "DefaultLayPercent"))
dDefaultPlacePc = Val(GetSetting(App.EXEName, "Settings", "DefaultPlaceLayPercent"))

Dim i As Integer
Dim dTotal As Double, dThis As Double
dTotal = 0
For i = 0 To 2
    dThis = Val(optLayPc(i).Caption)
    If dThis = dDefaultLayPercent Then optLayPc(i).Value = True
    dThis = Val(optPlaceLayPc(i).Caption)
    If dThis = dDefaultPlacePc Then optPlaceLayPc(i).Value = True
    dTotal = dTotal + dThis
Next i
If dTotal = 0 Then
    ' reset to default
    optLayPc(1).Caption = "2"
    optLayPc(2).Caption = "5"
    optBackPc(1).Caption = "2"
    optPlaceLayPc(1).Caption = "2"
    optPlaceLayPc(2).Caption = "5"
End If
Dim dDefaultBackPercent As Double
dDefaultBackPercent = Val(GetSetting(App.EXEName, "Settings", "DefaultBackPercent"))
If dDefaultBackPercent = 0 Then optBackPc(0).Value = True Else optBackPc(1).Value = True

frmMain.Top = Val(GetSetting(App.EXEName, "Position", "Top", "1000"))
frmMain.Left = Val(GetSetting(App.EXEName, "Position", "Left", "1000"))
mnuCorporateColours.Checked = False
If GetSetting(App.EXEName, "Colours", "Corporate") = "YES" Then
    mnuCorporateColours_Click
Else
    bCorporate = False
End If

Select Case sMode
Case "EW": chkEW.Value = 1: Call chkEW_Click
Case "Dutch": Call picDutch_Click
End Select
End Sub

Private Sub mnuMax_Click(Index As Integer)
Dim i As Integer
For i = 0 To 4
    If i = Index Then
        mnuMax(i).Checked = True
        Select Case i
        Case 0: lMaxIterations = 10000
        Case 1: lMaxIterations = 25000
        Case 2: lMaxIterations = 50000
        Case 3: lMaxIterations = 100000
        Case 4: lMaxIterations = 1000000
        End Select
    Else
        mnuMax(i).Checked = False
    End If
Next i
End Sub

Private Sub mnuNew_Click()
SaveSettings
DoEvents
Call ShellExecute(Me.hwnd, "open", App.EXEName + ".exe", "", "", 4)
Dim i As Integer
i = frmMain.Left
If i > 80 Then frmMain.Left = i - 80
End Sub

Private Sub mnuReset_Click()
txtBackDesc.Text = ""
txtLayDesc.Text = ""
txtBackOdds.Text = ""
BackOdds = 0
txtBackStake.Text = ""
BackStake = 0
txtLayOdds.Text = ""
LayOdds = 0
txtLayStake.Text = ""
LayStake = 0
txtPlaceLayStake.Text = ""
txtPlaceLayOdds.Text = ""
PlaceLayOdds = 0
PlaceLayStake = 0
txtPlaceStake.Text = ""
txtPlaceBackOdds.Text = ""
sCompare = ""
txtResult.Text = ""
frmMain.Caption = "Lay Odds Equalizer"
If mnuResetClears.Checked Then Call cmdClearHistory_Click
End Sub




Private Sub mnuResetClears_Click()
    mnuResetClears.Checked = Not (mnuResetClears.Checked)
        
    
End Sub

Private Sub mnuStandard_Click()
    mnuStandard.Checked = True
    If mnuEW.Checked Then
        mnuEW.Checked = False
        chkEW.Value = 0
        Call chkEW_Click
    End If
End Sub

Private Sub mnuUpgrade_Click()
'picUpgrade_Click
End Sub

Private Sub optLayPc_Click(Index As Integer)
LayPc = Val(optLayPc(Index).Caption)
mnuLayPc.Caption = "Lay Commission " + CStr(LayPc) + "%"
End Sub

Private Sub optBackPc_Click(Index As Integer)
BackPc = Val(optBackPc(Index).Caption)
mnuBackPc.Caption = "Back Commission " + CStr(BackPc) + "%"
End Sub
Private Sub optBackPc_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim d As Double
If Button = vbRightButton Then
    If Index > 0 Then
        d = Val(InputBox("Enter Custom Back Percentage", "Back %ge"))
        If d > 0 And d < 100 Then
            optBackPc(Index).Caption = CStr(d)
            optBackPc(Index).SetFocus
            BackPc = d
            optBackPc(Index).Value = True
            Call optBackPc_Click(Index)
        End If
    Else
        MsgBox "First Percentage is fixed at zero percent", vbOKOnly + vbInformation, "Custom Back Percentage"
    End If
End If
End Sub
Private Sub optLayPc_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim d As Double
If Button = vbRightButton Then
    If Index > 0 Then
        d = Val(InputBox("Enter Custom Lay Percentage", "Lay %ge"))
        If d > 0 And d < 100 Then
            optLayPc(Index).Caption = CStr(d)
            optLayPc(Index).SetFocus
            LayPc = d
            optLayPc(Index).Value = True
            Call optLayPc_Click(Index)
        End If
    Else
        MsgBox "First Percentage is fixed at zero percent", vbOKOnly + vbInformation, "Custom Lay Percentage"
    End If
End If
End Sub



Private Sub optPlaceLayPc_Click(Index As Integer)
PlacePc = Val(optPlaceLayPc(Index).Caption)
End Sub

Private Sub optPlaceLayPc_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim d As Double
If Button = vbRightButton Then
    If Index > 0 Then
        d = Val(InputBox("Enter Custom Lay Place Percentage", "Place Lay %ge"))
        If d > 0 And d < 100 Then
            optPlaceLayPc(Index).Caption = CStr(d)
            optPlaceLayPc(Index).SetFocus
            PlacePc = d
            optPlaceLayPc(Index).Value = True
            Call optPlaceLayPc_Click(Index)
        End If
    Else
        MsgBox "First Percentage is fixed at zero percent", vbOKOnly + vbInformation, "Custom Place Lay Percentage"
    End If
End If
End Sub




Private Sub picDutch_Click()
frmMain.Visible = False 'hide

If sCompare = "" Then
    frmDutch.lblCompare.Caption = "Comparison Bet: None Set"
Else
    frmDutch.lblCompare.Caption = "Compare : " & sCompare
End If
frmDutch.Top = frmMain.Top
frmDutch.Left = frmMain.Left
frmDutch.Show

End Sub

'Private Sub picUpgrade_Click()
'If picUpgrade.Visible Then Call mnuAbout_Click
'
'End Sub




Private Sub txtBackDesc_KeyUp(KeyCode As Integer, Shift As Integer)
Dim sCap As String
sCap = "Lay Odds Equalizer - Bet=" + txtBackDesc.Text
If Len(txtLayDesc.Text) Then sCap = sCap + "/" + txtLayDesc.Text
frmMain.Caption = sCap
End Sub

Private Sub txtBackOdds_GotFocus()
If Len(txtBackOdds.Text) Then
txtBackOdds.SelStart = 0
txtBackOdds.SelLength = Len(txtBackOdds.Text)
End If
End Sub

Private Sub txtBackOdds_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc("\") Then KeyAscii = Asc("/")
End Sub

Private Sub txtBackOdds_LostFocus()
txtBackOdds.Text = ConvertOdds(txtBackOdds.Text)
BackOdds = Val(txtBackOdds.Text)
If BackOdds = 0 Then txtPlaceBackOdds.Text = "0.0": Exit Sub
Dim d As Double
Select Case cboTerms.Text
Case "1/5": d = 0.2
Case "1/4": d = 0.25
Case "1/3": d = 0.333333
End Select
txtPlaceBackOdds.Text = ConvertOdds(((BackOdds - 1) * d) + 1)
PlaceBackOdds = Val(txtPlaceBackOdds.Text)
End Sub

Private Sub txtBackStake_GotFocus()
If Len(txtBackStake.Text) Then
txtBackStake.SelStart = 0
txtBackStake.SelLength = Len(txtBackStake.Text)
End If
End Sub

Private Sub txtBackStake_LostFocus()
BackStake = Val(txtBackStake.Text)
txtBackStake.Text = Format(BackStake, "0.00")
txtPlaceStake.Text = txtBackStake.Text + " / " + Format(BackStake * 2, "0.00")
End Sub



Private Sub txtBackDesc_GotFocus()
If Len(txtBackDesc.Text) Then
    txtBackDesc.SelStart = 0
    txtBackDesc.SelLength = Len(txtBackDesc.Text)
End If
End Sub



Private Sub txtPlaceBackOdds_GotFocus()
txtPlaceLayOdds.SetFocus
End Sub

Private Sub txtLayDesc_GotFocus()
If Len(txtLayDesc.Text) Then
txtLayDesc.SelStart = 0
txtLayDesc.SelLength = Len(txtLayDesc.Text)
End If
End Sub

Private Sub txtLayDesc_KeyUp(KeyCode As Integer, Shift As Integer)
Dim sCap As String
sCap = "Lay Odds Equalizer - Bet=" + txtBackDesc.Text
If Len(txtLayDesc.Text) Then sCap = sCap + "/" + txtLayDesc.Text
frmMain.Caption = sCap
End Sub

Private Sub txtLayOdds_GotFocus()
If Len(txtLayOdds.Text) Then
txtLayOdds.SelStart = 0
txtLayOdds.SelLength = Len(txtLayOdds.Text)
End If
End Sub

Private Sub txtLayOdds_LostFocus()
txtLayOdds.Text = ConvertOdds(txtLayOdds.Text)
LayOdds = Val(txtLayOdds.Text)
End Sub

Private Sub txtLayStake_GotFocus()
If Len(txtLayStake.Text) Then
txtLayStake.SelStart = 0
txtLayStake.SelLength = Len(txtLayStake.Text)
End If
End Sub

Private Sub txtLayStake_LostFocus()
LayStake = Val(txtLayStake.Text)
txtLayStake.Text = Format(LayStake, "0.00")
End Sub
Private Sub txtPlaceLayOdds_LostFocus()
txtPlaceLayOdds.Text = ConvertOdds(txtPlaceLayOdds.Text)
PlaceLayOdds = Val(txtPlaceLayOdds.Text)
End Sub
Private Sub txtPlaceLayStake_LostFocus()
PlaceLayStake = Val(txtPlaceLayStake.Text)
txtPlaceLayStake.Text = Format(PlaceLayStake, "0.00")
End Sub



Private Sub txtPlaceLayOdds_GotFocus()
If Len(txtPlaceLayOdds.Text) Then
txtPlaceLayOdds.SelStart = 0
txtPlaceLayOdds.SelLength = Len(txtPlaceLayOdds.Text)
End If
End Sub

Private Sub txtPlaceLayStake_GotFocus()
If Len(txtPlaceLayStake.Text) Then
txtPlaceLayStake.SelStart = 0
txtPlaceLayStake.SelLength = Len(txtPlaceLayStake.Text)
End If
End Sub

Private Sub txtPlaceStake_GotFocus()
txtPlaceLayOdds.SetFocus
End Sub
