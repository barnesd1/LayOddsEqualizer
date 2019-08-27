VERSION 5.00
Begin VB.Form frmDutch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lay Odds Equalizer - Dutch Calculator"
   ClientHeight    =   7665
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6450
   BeginProperty Font 
   EndProperty
   Font            =   "frmDutch.frx":0000
   Icon            =   "frmDutch.frx":0018
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   6450
   Begin VB.Frame fraCompare 
      BorderStyle     =   0  'None
      BeginProperty Font 
      EndProperty
      Font            =   "frmDutch.frx":045A
      Height          =   255
      Left            =   60
      TabIndex        =   150
      Top             =   7440
      Width           =   6375
      Begin VB.Label lblCompare 
         AutoSize        =   -1  'True
         Caption         =   "Comparison Bet: Back"
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0472
         Height          =   195
         Left            =   0
         TabIndex        =   151
         Top             =   30
         Width           =   1575
      End
   End
   Begin VB.ComboBox cmbRunners 
      BeginProperty Font 
      EndProperty
      Font            =   "frmDutch.frx":048A
      Height          =   315
      Left            =   3540
      TabIndex        =   3
      Top             =   180
      Width           =   1215
   End
   Begin VB.CommandButton cmdDistribute 
      Appearance      =   0  'Flat
      Caption         =   "DISTRIBUTE"
      BeginProperty Font 
      EndProperty
      Font            =   "frmDutch.frx":04A2
      Height          =   285
      Left            =   3730
      TabIndex        =   146
      Top             =   6960
      Width           =   1155
   End
   Begin VB.TextBox txtTotalBet 
      BeginProperty Font 
      EndProperty
      Font            =   "frmDutch.frx":04BA
      Height          =   285
      Left            =   2680
      TabIndex        =   145
      Top             =   6960
      Width           =   975
   End
   Begin VB.Frame fraDutch 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Bet                      Odds      Comm.   Stake            Action    Result           Profit"
      BeginProperty Font 
      EndProperty
      Font            =   "frmDutch.frx":04D2
      ForeColor       =   &H00800080&
      Height          =   6855
      Left            =   120
      TabIndex        =   148
      Top             =   525
      Width           =   6255
      Begin VB.ComboBox cboRound 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":04EA
         Height          =   315
         Left            =   240
         TabIndex        =   149
         Top             =   6420
         Width           =   1335
      End
      Begin VB.TextBox txtBet 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0502
         Height          =   285
         Index           =   19
         Left            =   120
         TabIndex        =   137
         Top             =   6000
         Width           =   1215
      End
      Begin VB.TextBox txtOdds 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":051A
         Height          =   285
         Index           =   19
         Left            =   1320
         TabIndex        =   138
         Top             =   6000
         Width           =   675
      End
      Begin VB.TextBox txtComm 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0532
         Height          =   285
         Index           =   19
         Left            =   1980
         TabIndex        =   139
         Top             =   6000
         Width           =   615
      End
      Begin VB.TextBox txtStake 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":054A
         Height          =   285
         Index           =   19
         Left            =   2580
         TabIndex        =   140
         Top             =   6000
         Width           =   975
      End
      Begin VB.TextBox txtResult 
         Enabled         =   0   'False
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0562
         Height          =   285
         Index           =   19
         Left            =   4200
         TabIndex        =   142
         Top             =   6000
         Width           =   975
      End
      Begin VB.TextBox txtProfit 
         Enabled         =   0   'False
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":057A
         Height          =   285
         Index           =   19
         Left            =   5160
         TabIndex        =   143
         Top             =   6000
         Width           =   975
      End
      Begin VB.CommandButton cmdBase 
         Appearance      =   0  'Flat
         Caption         =   "BASE"
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0592
         Height          =   285
         Index           =   19
         Left            =   3585
         TabIndex        =   141
         Top             =   6015
         Width           =   615
      End
      Begin VB.TextBox txtBet 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":05AC
         Height          =   285
         Index           =   18
         Left            =   120
         TabIndex        =   130
         Top             =   5700
         Width           =   1215
      End
      Begin VB.TextBox txtOdds 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":05C4
         Height          =   285
         Index           =   18
         Left            =   1320
         TabIndex        =   131
         Top             =   5700
         Width           =   675
      End
      Begin VB.TextBox txtComm 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":05DC
         Height          =   285
         Index           =   18
         Left            =   1980
         TabIndex        =   132
         Top             =   5700
         Width           =   615
      End
      Begin VB.TextBox txtStake 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":05F4
         Height          =   285
         Index           =   18
         Left            =   2580
         TabIndex        =   133
         Top             =   5700
         Width           =   975
      End
      Begin VB.TextBox txtResult 
         Enabled         =   0   'False
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":060C
         Height          =   285
         Index           =   18
         Left            =   4200
         TabIndex        =   135
         Top             =   5700
         Width           =   975
      End
      Begin VB.TextBox txtProfit 
         Enabled         =   0   'False
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0624
         Height          =   285
         Index           =   18
         Left            =   5160
         TabIndex        =   136
         Top             =   5700
         Width           =   975
      End
      Begin VB.CommandButton cmdBase 
         Appearance      =   0  'Flat
         Caption         =   "BASE"
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":063C
         Height          =   285
         Index           =   18
         Left            =   3585
         TabIndex        =   134
         Top             =   5715
         Width           =   615
      End
      Begin VB.TextBox txtBet 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0656
         Height          =   285
         Index           =   17
         Left            =   120
         TabIndex        =   123
         Top             =   5400
         Width           =   1215
      End
      Begin VB.TextBox txtOdds 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":066E
         Height          =   285
         Index           =   17
         Left            =   1320
         TabIndex        =   124
         Top             =   5400
         Width           =   675
      End
      Begin VB.TextBox txtComm 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0686
         Height          =   285
         Index           =   17
         Left            =   1980
         TabIndex        =   125
         Top             =   5400
         Width           =   615
      End
      Begin VB.TextBox txtStake 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":069E
         Height          =   285
         Index           =   17
         Left            =   2580
         TabIndex        =   126
         Top             =   5400
         Width           =   975
      End
      Begin VB.TextBox txtResult 
         Enabled         =   0   'False
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":06B6
         Height          =   285
         Index           =   17
         Left            =   4200
         TabIndex        =   128
         Top             =   5400
         Width           =   975
      End
      Begin VB.TextBox txtProfit 
         Enabled         =   0   'False
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":06CE
         Height          =   285
         Index           =   17
         Left            =   5160
         TabIndex        =   129
         Top             =   5400
         Width           =   975
      End
      Begin VB.CommandButton cmdBase 
         Appearance      =   0  'Flat
         Caption         =   "BASE"
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":06E6
         Height          =   285
         Index           =   17
         Left            =   3585
         TabIndex        =   127
         Top             =   5415
         Width           =   615
      End
      Begin VB.TextBox txtBet 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0700
         Height          =   285
         Index           =   16
         Left            =   120
         TabIndex        =   116
         Top             =   5100
         Width           =   1215
      End
      Begin VB.TextBox txtOdds 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0718
         Height          =   285
         Index           =   16
         Left            =   1320
         TabIndex        =   117
         Top             =   5100
         Width           =   675
      End
      Begin VB.TextBox txtComm 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0730
         Height          =   285
         Index           =   16
         Left            =   1980
         TabIndex        =   118
         Top             =   5100
         Width           =   615
      End
      Begin VB.TextBox txtStake 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0748
         Height          =   285
         Index           =   16
         Left            =   2580
         TabIndex        =   119
         Top             =   5100
         Width           =   975
      End
      Begin VB.TextBox txtResult 
         Enabled         =   0   'False
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0760
         Height          =   285
         Index           =   16
         Left            =   4200
         TabIndex        =   121
         Top             =   5100
         Width           =   975
      End
      Begin VB.TextBox txtProfit 
         Enabled         =   0   'False
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0778
         Height          =   285
         Index           =   16
         Left            =   5160
         TabIndex        =   122
         Top             =   5100
         Width           =   975
      End
      Begin VB.CommandButton cmdBase 
         Appearance      =   0  'Flat
         Caption         =   "BASE"
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0790
         Height          =   285
         Index           =   16
         Left            =   3585
         TabIndex        =   120
         Top             =   5115
         Width           =   615
      End
      Begin VB.TextBox txtBet 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":07AA
         Height          =   285
         Index           =   15
         Left            =   120
         TabIndex        =   109
         Top             =   4800
         Width           =   1215
      End
      Begin VB.TextBox txtOdds 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":07C2
         Height          =   285
         Index           =   15
         Left            =   1320
         TabIndex        =   110
         Top             =   4800
         Width           =   675
      End
      Begin VB.TextBox txtComm 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":07DA
         Height          =   285
         Index           =   15
         Left            =   1980
         TabIndex        =   111
         Top             =   4800
         Width           =   615
      End
      Begin VB.TextBox txtStake 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":07F2
         Height          =   285
         Index           =   15
         Left            =   2580
         TabIndex        =   112
         Top             =   4800
         Width           =   975
      End
      Begin VB.TextBox txtResult 
         Enabled         =   0   'False
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":080A
         Height          =   285
         Index           =   15
         Left            =   4200
         TabIndex        =   114
         Top             =   4800
         Width           =   975
      End
      Begin VB.TextBox txtProfit 
         Enabled         =   0   'False
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0822
         Height          =   285
         Index           =   15
         Left            =   5160
         TabIndex        =   115
         Top             =   4800
         Width           =   975
      End
      Begin VB.CommandButton cmdBase 
         Appearance      =   0  'Flat
         Caption         =   "BASE"
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":083A
         Height          =   285
         Index           =   15
         Left            =   3585
         TabIndex        =   113
         Top             =   4815
         Width           =   615
      End
      Begin VB.TextBox txtBet 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0854
         Height          =   285
         Index           =   14
         Left            =   120
         TabIndex        =   102
         Top             =   4500
         Width           =   1215
      End
      Begin VB.TextBox txtOdds 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":086C
         Height          =   285
         Index           =   14
         Left            =   1320
         TabIndex        =   103
         Top             =   4500
         Width           =   675
      End
      Begin VB.TextBox txtComm 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0884
         Height          =   285
         Index           =   14
         Left            =   1980
         TabIndex        =   104
         Top             =   4500
         Width           =   615
      End
      Begin VB.TextBox txtStake 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":089C
         Height          =   285
         Index           =   14
         Left            =   2580
         TabIndex        =   105
         Top             =   4500
         Width           =   975
      End
      Begin VB.TextBox txtResult 
         Enabled         =   0   'False
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":08B4
         Height          =   285
         Index           =   14
         Left            =   4200
         TabIndex        =   107
         Top             =   4500
         Width           =   975
      End
      Begin VB.TextBox txtProfit 
         Enabled         =   0   'False
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":08CC
         Height          =   285
         Index           =   14
         Left            =   5160
         TabIndex        =   108
         Top             =   4500
         Width           =   975
      End
      Begin VB.CommandButton cmdBase 
         Appearance      =   0  'Flat
         Caption         =   "BASE"
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":08E4
         Height          =   285
         Index           =   14
         Left            =   3585
         TabIndex        =   106
         Top             =   4515
         Width           =   615
      End
      Begin VB.TextBox txtBet 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":08FE
         Height          =   285
         Index           =   13
         Left            =   120
         TabIndex        =   95
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox txtOdds 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0916
         Height          =   285
         Index           =   13
         Left            =   1320
         TabIndex        =   96
         Top             =   4200
         Width           =   675
      End
      Begin VB.TextBox txtComm 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":092E
         Height          =   285
         Index           =   13
         Left            =   1980
         TabIndex        =   97
         Top             =   4200
         Width           =   615
      End
      Begin VB.TextBox txtStake 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0946
         Height          =   285
         Index           =   13
         Left            =   2580
         TabIndex        =   98
         Top             =   4200
         Width           =   975
      End
      Begin VB.TextBox txtResult 
         Enabled         =   0   'False
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":095E
         Height          =   285
         Index           =   13
         Left            =   4200
         TabIndex        =   100
         Top             =   4200
         Width           =   975
      End
      Begin VB.TextBox txtProfit 
         Enabled         =   0   'False
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0976
         Height          =   285
         Index           =   13
         Left            =   5160
         TabIndex        =   101
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton cmdBase 
         Appearance      =   0  'Flat
         Caption         =   "BASE"
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":098E
         Height          =   285
         Index           =   13
         Left            =   3585
         TabIndex        =   99
         Top             =   4215
         Width           =   615
      End
      Begin VB.TextBox txtBet 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":09A8
         Height          =   285
         Index           =   12
         Left            =   120
         TabIndex        =   88
         Top             =   3900
         Width           =   1215
      End
      Begin VB.TextBox txtOdds 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":09C0
         Height          =   285
         Index           =   12
         Left            =   1320
         TabIndex        =   89
         Top             =   3900
         Width           =   675
      End
      Begin VB.TextBox txtComm 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":09D8
         Height          =   285
         Index           =   12
         Left            =   1980
         TabIndex        =   90
         Top             =   3900
         Width           =   615
      End
      Begin VB.TextBox txtStake 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":09F0
         Height          =   285
         Index           =   12
         Left            =   2580
         TabIndex        =   91
         Top             =   3900
         Width           =   975
      End
      Begin VB.TextBox txtResult 
         Enabled         =   0   'False
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0A08
         Height          =   285
         Index           =   12
         Left            =   4200
         TabIndex        =   93
         Top             =   3900
         Width           =   975
      End
      Begin VB.TextBox txtProfit 
         Enabled         =   0   'False
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0A20
         Height          =   285
         Index           =   12
         Left            =   5160
         TabIndex        =   94
         Top             =   3900
         Width           =   975
      End
      Begin VB.CommandButton cmdBase 
         Appearance      =   0  'Flat
         Caption         =   "BASE"
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0A38
         Height          =   285
         Index           =   12
         Left            =   3585
         TabIndex        =   92
         Top             =   3915
         Width           =   615
      End
      Begin VB.TextBox txtBet 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0A52
         Height          =   285
         Index           =   11
         Left            =   120
         TabIndex        =   81
         Top             =   3600
         Width           =   1215
      End
      Begin VB.TextBox txtOdds 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0A6A
         Height          =   285
         Index           =   11
         Left            =   1320
         TabIndex        =   82
         Top             =   3600
         Width           =   675
      End
      Begin VB.TextBox txtComm 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0A82
         Height          =   285
         Index           =   11
         Left            =   1980
         TabIndex        =   83
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox txtStake 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0A9A
         Height          =   285
         Index           =   11
         Left            =   2580
         TabIndex        =   84
         Top             =   3600
         Width           =   975
      End
      Begin VB.TextBox txtResult 
         Enabled         =   0   'False
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0AB2
         Height          =   285
         Index           =   11
         Left            =   4200
         TabIndex        =   86
         Top             =   3600
         Width           =   975
      End
      Begin VB.TextBox txtProfit 
         Enabled         =   0   'False
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0ACA
         Height          =   285
         Index           =   11
         Left            =   5160
         TabIndex        =   87
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton cmdBase 
         Appearance      =   0  'Flat
         Caption         =   "BASE"
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0AE2
         Height          =   285
         Index           =   11
         Left            =   3585
         TabIndex        =   85
         Top             =   3615
         Width           =   615
      End
      Begin VB.TextBox txtBet 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0AFC
         Height          =   285
         Index           =   10
         Left            =   120
         TabIndex        =   74
         Top             =   3300
         Width           =   1215
      End
      Begin VB.TextBox txtOdds 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0B14
         Height          =   285
         Index           =   10
         Left            =   1320
         TabIndex        =   75
         Top             =   3300
         Width           =   675
      End
      Begin VB.TextBox txtComm 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0B2C
         Height          =   285
         Index           =   10
         Left            =   1980
         TabIndex        =   76
         Top             =   3300
         Width           =   615
      End
      Begin VB.TextBox txtStake 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0B44
         Height          =   285
         Index           =   10
         Left            =   2580
         TabIndex        =   77
         Top             =   3300
         Width           =   975
      End
      Begin VB.TextBox txtResult 
         Enabled         =   0   'False
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0B5C
         Height          =   285
         Index           =   10
         Left            =   4200
         TabIndex        =   79
         Top             =   3300
         Width           =   975
      End
      Begin VB.TextBox txtProfit 
         Enabled         =   0   'False
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0B74
         Height          =   285
         Index           =   10
         Left            =   5160
         TabIndex        =   80
         Top             =   3300
         Width           =   975
      End
      Begin VB.CommandButton cmdBase 
         Appearance      =   0  'Flat
         Caption         =   "BASE"
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0B8C
         Height          =   285
         Index           =   10
         Left            =   3585
         TabIndex        =   78
         Top             =   3315
         Width           =   615
      End
      Begin VB.TextBox txtBet 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0BA6
         Height          =   285
         Index           =   9
         Left            =   120
         TabIndex        =   67
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox txtOdds 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0BBE
         Height          =   285
         Index           =   9
         Left            =   1320
         TabIndex        =   68
         Top             =   3000
         Width           =   675
      End
      Begin VB.TextBox txtComm 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0BD6
         Height          =   285
         Index           =   9
         Left            =   1980
         TabIndex        =   69
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox txtStake 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0BEE
         Height          =   285
         Index           =   9
         Left            =   2580
         TabIndex        =   70
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox txtResult 
         Enabled         =   0   'False
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0C06
         Height          =   285
         Index           =   9
         Left            =   4200
         TabIndex        =   72
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox txtProfit 
         Enabled         =   0   'False
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0C1E
         Height          =   285
         Index           =   9
         Left            =   5160
         TabIndex        =   73
         Top             =   3000
         Width           =   975
      End
      Begin VB.CommandButton cmdBase 
         Appearance      =   0  'Flat
         Caption         =   "BASE"
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0C36
         Height          =   285
         Index           =   9
         Left            =   3585
         TabIndex        =   71
         Top             =   3015
         Width           =   615
      End
      Begin VB.TextBox txtBet 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0C50
         Height          =   285
         Index           =   8
         Left            =   120
         TabIndex        =   60
         Top             =   2700
         Width           =   1215
      End
      Begin VB.TextBox txtOdds 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0C68
         Height          =   285
         Index           =   8
         Left            =   1320
         TabIndex        =   61
         Top             =   2700
         Width           =   675
      End
      Begin VB.TextBox txtComm 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0C80
         Height          =   285
         Index           =   8
         Left            =   1980
         TabIndex        =   62
         Top             =   2700
         Width           =   615
      End
      Begin VB.TextBox txtStake 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0C98
         Height          =   285
         Index           =   8
         Left            =   2580
         TabIndex        =   63
         Top             =   2700
         Width           =   975
      End
      Begin VB.TextBox txtResult 
         Enabled         =   0   'False
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0CB0
         Height          =   285
         Index           =   8
         Left            =   4200
         TabIndex        =   65
         Top             =   2700
         Width           =   975
      End
      Begin VB.TextBox txtProfit 
         Enabled         =   0   'False
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0CC8
         Height          =   285
         Index           =   8
         Left            =   5160
         TabIndex        =   66
         Top             =   2700
         Width           =   975
      End
      Begin VB.CommandButton cmdBase 
         Appearance      =   0  'Flat
         Caption         =   "BASE"
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0CE0
         Height          =   285
         Index           =   8
         Left            =   3585
         TabIndex        =   64
         Top             =   2715
         Width           =   615
      End
      Begin VB.TextBox txtBet 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0CFA
         Height          =   285
         Index           =   7
         Left            =   120
         TabIndex        =   53
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txtOdds 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0D12
         Height          =   285
         Index           =   7
         Left            =   1320
         TabIndex        =   54
         Top             =   2400
         Width           =   675
      End
      Begin VB.TextBox txtComm 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0D2A
         Height          =   285
         Index           =   7
         Left            =   1980
         TabIndex        =   55
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox txtStake 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0D42
         Height          =   285
         Index           =   7
         Left            =   2580
         TabIndex        =   56
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox txtResult 
         Enabled         =   0   'False
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0D5A
         Height          =   285
         Index           =   7
         Left            =   4200
         TabIndex        =   58
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox txtProfit 
         Enabled         =   0   'False
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0D72
         Height          =   285
         Index           =   7
         Left            =   5160
         TabIndex        =   59
         Top             =   2400
         Width           =   975
      End
      Begin VB.CommandButton cmdBase 
         Appearance      =   0  'Flat
         Caption         =   "BASE"
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0D8A
         Height          =   285
         Index           =   7
         Left            =   3585
         TabIndex        =   57
         Top             =   2415
         Width           =   615
      End
      Begin VB.TextBox txtBet 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0DA4
         Height          =   285
         Index           =   6
         Left            =   120
         TabIndex        =   46
         Top             =   2100
         Width           =   1215
      End
      Begin VB.TextBox txtOdds 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0DBC
         Height          =   285
         Index           =   6
         Left            =   1320
         TabIndex        =   47
         Top             =   2100
         Width           =   675
      End
      Begin VB.TextBox txtComm 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0DD4
         Height          =   285
         Index           =   6
         Left            =   1980
         TabIndex        =   48
         Top             =   2100
         Width           =   615
      End
      Begin VB.TextBox txtStake 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0DEC
         Height          =   285
         Index           =   6
         Left            =   2580
         TabIndex        =   49
         Top             =   2100
         Width           =   975
      End
      Begin VB.TextBox txtResult 
         Enabled         =   0   'False
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0E04
         Height          =   285
         Index           =   6
         Left            =   4200
         TabIndex        =   51
         Top             =   2100
         Width           =   975
      End
      Begin VB.TextBox txtProfit 
         Enabled         =   0   'False
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0E1C
         Height          =   285
         Index           =   6
         Left            =   5160
         TabIndex        =   52
         Top             =   2100
         Width           =   975
      End
      Begin VB.CommandButton cmdBase 
         Appearance      =   0  'Flat
         Caption         =   "BASE"
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0E34
         Height          =   285
         Index           =   6
         Left            =   3585
         TabIndex        =   50
         Top             =   2115
         Width           =   615
      End
      Begin VB.TextBox txtBet 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0E4E
         Height          =   285
         Index           =   5
         Left            =   120
         TabIndex        =   39
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtOdds 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0E66
         Height          =   285
         Index           =   5
         Left            =   1320
         TabIndex        =   40
         Top             =   1800
         Width           =   675
      End
      Begin VB.TextBox txtComm 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0E7E
         Height          =   285
         Index           =   5
         Left            =   1980
         TabIndex        =   41
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox txtStake 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0E96
         Height          =   285
         Index           =   5
         Left            =   2580
         TabIndex        =   42
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtResult 
         Enabled         =   0   'False
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0EAE
         Height          =   285
         Index           =   5
         Left            =   4200
         TabIndex        =   44
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtProfit 
         Enabled         =   0   'False
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0EC6
         Height          =   285
         Index           =   5
         Left            =   5160
         TabIndex        =   45
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton cmdBase 
         Appearance      =   0  'Flat
         Caption         =   "BASE"
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0EDE
         Height          =   285
         Index           =   5
         Left            =   3585
         TabIndex        =   43
         Top             =   1815
         Width           =   615
      End
      Begin VB.TextBox txtBet 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0EF8
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   32
         Top             =   1500
         Width           =   1215
      End
      Begin VB.TextBox txtOdds 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0F10
         Height          =   285
         Index           =   4
         Left            =   1320
         TabIndex        =   33
         Top             =   1500
         Width           =   675
      End
      Begin VB.TextBox txtComm 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0F28
         Height          =   285
         Index           =   4
         Left            =   1980
         TabIndex        =   34
         Top             =   1500
         Width           =   615
      End
      Begin VB.TextBox txtStake 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0F40
         Height          =   285
         Index           =   4
         Left            =   2580
         TabIndex        =   35
         Top             =   1500
         Width           =   975
      End
      Begin VB.TextBox txtResult 
         Enabled         =   0   'False
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0F58
         Height          =   285
         Index           =   4
         Left            =   4200
         TabIndex        =   37
         Top             =   1500
         Width           =   975
      End
      Begin VB.TextBox txtProfit 
         Enabled         =   0   'False
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0F70
         Height          =   285
         Index           =   4
         Left            =   5160
         TabIndex        =   38
         Top             =   1500
         Width           =   975
      End
      Begin VB.CommandButton cmdBase 
         Appearance      =   0  'Flat
         Caption         =   "BASE"
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0F88
         Height          =   285
         Index           =   4
         Left            =   3585
         TabIndex        =   36
         Top             =   1515
         Width           =   615
      End
      Begin VB.TextBox txtBet 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0FA2
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   25
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtOdds 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0FBA
         Height          =   285
         Index           =   3
         Left            =   1320
         TabIndex        =   26
         Top             =   1200
         Width           =   675
      End
      Begin VB.TextBox txtComm 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0FD2
         Height          =   285
         Index           =   3
         Left            =   1980
         TabIndex        =   27
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtStake 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":0FEA
         Height          =   285
         Index           =   3
         Left            =   2580
         TabIndex        =   28
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtResult 
         Enabled         =   0   'False
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":1002
         Height          =   285
         Index           =   3
         Left            =   4200
         TabIndex        =   30
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtProfit 
         Enabled         =   0   'False
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":101A
         Height          =   285
         Index           =   3
         Left            =   5160
         TabIndex        =   31
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdBase 
         Appearance      =   0  'Flat
         Caption         =   "BASE"
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":1032
         Height          =   285
         Index           =   3
         Left            =   3585
         TabIndex        =   29
         Top             =   1215
         Width           =   615
      End
      Begin VB.TextBox txtBet 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":104C
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   900
         Width           =   1215
      End
      Begin VB.TextBox txtOdds 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":1064
         Height          =   285
         Index           =   2
         Left            =   1320
         TabIndex        =   19
         Top             =   900
         Width           =   675
      End
      Begin VB.TextBox txtComm 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":107C
         Height          =   285
         Index           =   2
         Left            =   1980
         TabIndex        =   20
         Top             =   900
         Width           =   615
      End
      Begin VB.TextBox txtStake 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":1094
         Height          =   285
         Index           =   2
         Left            =   2580
         TabIndex        =   21
         Top             =   900
         Width           =   975
      End
      Begin VB.TextBox txtResult 
         Enabled         =   0   'False
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":10AC
         Height          =   285
         Index           =   2
         Left            =   4200
         TabIndex        =   23
         Top             =   900
         Width           =   975
      End
      Begin VB.TextBox txtProfit 
         Enabled         =   0   'False
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":10C4
         Height          =   285
         Index           =   2
         Left            =   5160
         TabIndex        =   24
         Top             =   900
         Width           =   975
      End
      Begin VB.CommandButton cmdBase 
         Appearance      =   0  'Flat
         Caption         =   "BASE"
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":10DC
         Height          =   285
         Index           =   2
         Left            =   3585
         TabIndex        =   22
         Top             =   915
         Width           =   615
      End
      Begin VB.TextBox txtBet 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":10F6
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtOdds 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":110E
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   12
         Top             =   600
         Width           =   675
      End
      Begin VB.TextBox txtComm 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":1126
         Height          =   285
         Index           =   1
         Left            =   1980
         TabIndex        =   13
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtStake 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":113E
         Height          =   285
         Index           =   1
         Left            =   2580
         TabIndex        =   14
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtResult 
         Enabled         =   0   'False
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":1156
         Height          =   285
         Index           =   1
         Left            =   4200
         TabIndex        =   16
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtProfit 
         Enabled         =   0   'False
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":116E
         Height          =   285
         Index           =   1
         Left            =   5160
         TabIndex        =   17
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdBase 
         Appearance      =   0  'Flat
         Caption         =   "BASE"
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":1186
         Height          =   285
         Index           =   1
         Left            =   3585
         TabIndex        =   15
         Top             =   615
         Width           =   615
      End
      Begin VB.CommandButton cmdBase 
         Appearance      =   0  'Flat
         Caption         =   "BASE"
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":11A0
         Height          =   285
         Index           =   0
         Left            =   3585
         TabIndex        =   8
         Top             =   315
         Width           =   615
      End
      Begin VB.TextBox txtProfit 
         Enabled         =   0   'False
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":11BA
         Height          =   285
         Index           =   0
         Left            =   5160
         TabIndex        =   10
         Top             =   300
         Width           =   975
      End
      Begin VB.TextBox txtResult 
         Enabled         =   0   'False
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":11D2
         Height          =   285
         Index           =   0
         Left            =   4200
         TabIndex        =   9
         Top             =   300
         Width           =   975
      End
      Begin VB.TextBox txtStake 
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
         EndProperty
         Font            =   "frmDutch.frx":11EA
         Height          =   285
         Index           =   0
         Left            =   2580
         TabIndex        =   7
         Top             =   300
         Width           =   975
      End
      Begin VB.TextBox txtComm 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":1202
         Height          =   285
         Index           =   0
         Left            =   1980
         TabIndex        =   6
         Top             =   300
         Width           =   615
      End
      Begin VB.TextBox txtOdds 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":121A
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   5
         Top             =   300
         Width           =   675
      End
      Begin VB.TextBox txtBet 
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":1232
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Total Bet"
         BeginProperty Font 
         EndProperty
         Font            =   "frmDutch.frx":124A
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   1815
         TabIndex        =   144
         Top             =   6480
         Width           =   675
         WordWrap        =   -1  'True
      End
   End
   Begin VB.OptionButton optLegs 
      Caption         =   "Race"
      BeginProperty Font 
      EndProperty
      Font            =   "frmDutch.frx":1262
      Height          =   195
      Index           =   2
      Left            =   2640
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
   Begin VB.OptionButton optLegs 
      Caption         =   "Treble"
      BeginProperty Font 
      EndProperty
      Font            =   "frmDutch.frx":127A
      Height          =   195
      Index           =   1
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.OptionButton optLegs 
      Caption         =   "Double"
      BeginProperty Font 
      EndProperty
      Font            =   "frmDutch.frx":1292
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BeginProperty Font 
      EndProperty
      Font            =   "frmDutch.frx":12AA
      Height          =   195
      Left            =   4080
      TabIndex        =   147
      Top             =   240
      Width           =   45
   End
   Begin VB.Shape shpBox 
      BorderColor     =   &H00800080&
      BorderWidth     =   10
      FillColor       =   &H00800080&
      Height          =   7395
      Left            =   30
      Top             =   20
      Width           =   6435
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuNew 
      Caption         =   "&New"
   End
   Begin VB.Menu mnuReset 
      Caption         =   "&Reset"
   End
   Begin VB.Menu mnuGlobalComm 
      Caption         =   "&Global Commission"
   End
End
Attribute VB_Name = "frmDutch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim miLegs As Integer
Dim mbLoading As Boolean
Dim miBase As Integer
Dim CalcStake(0 To 19) As Double
Dim CalcOdds(0 To 19) As Double

Private Sub cboRound_Click()
    If mbLoading Then Exit Sub
    Dim i As Integer
    Dim dRoundTo As Double
    Select Case cboRound.listindex
    Case 0:
        cmdDistribute.value = True
        Exit Sub
    Case 1:
        dRoundTo = 0.25
    Case 2:
        dRoundTo = 0.5
    Case 3:
        dRoundTo = 1
    Case 4:
        dRoundTo = 5
    End Select
    For i = 0 To miLegs - 1
        If InStr(txtStake(i).Text, CurrSymbol) Then
            txtStake(i).Text = CurrSymbol + Format(RoundNumber(Val(Mid(txtStake(i).Text, Len(CurrSymbol) + 1)), dRoundTo), "0.00")
        Else
            txtStake(i).Text = Format(RoundNumber(Val(txtStake(i).Text), dRoundTo), "0.00")
        End If
    Next i
    Call CalcTotalBet
    For i = 0 To miLegs - 1
        If InStr(txtStake(i).Text, CurrSymbol) Then
                txtResult(i).Text = CurrSymbol + Format(Val(Mid(txtStake(i).Text, Len(CurrSymbol) + 1)) * CalcOdds(i), "0.00")
                txtProfit(i).Text = CurrSymbol + Format(Val(Mid(txtResult(i).Text, Len(CurrSymbol) + 1)) - Val(Mid(txtTotalBet.Text, Len(CurrSymbol) + 1)), "0.00")
        Else
                txtResult(i).Text = Format(Val(txtStake(i).Text) * CalcOdds(i), "0.00")
                txtProfit(i).Text = Format(Val(txtResult(i).Text) - Val(txtTotalBet.Text), "0.00")
        End If
    Next
End Sub

Function RoundNumber(ByVal OriginalNumber As Double, ByVal RoundTo As Double) As Double

' Have to do this due to problem with round(1.5) = round(2.5)
If (OriginalNumber / RoundTo) * 2 = CInt((OriginalNumber / RoundTo) * 2) Then
OriginalNumber = OriginalNumber + RoundTo / 10
End If

RoundNumber = Round(OriginalNumber / RoundTo, 0) * RoundTo

End Function
Private Sub cmbRunners_Click()
    If mbLoading Then Exit Sub
    miLegs = cmbRunners.listindex + 2
    SaveSetting app.exename, "Settings", "DefaultRunners", CStr(cmbRunners.listindex)
    Call resizeForm
End Sub
Private Function bMissingOdds() As Boolean
    Dim i As Integer
    For i = 0 To miLegs - 1
        If Val(txtOdds(i).Text) = 0 Then
            bMissingOdds = True
            Exit Function
        End If
    Next i
End Function

Private Sub cmdBase_Click(Index As Integer)
txtOdds(Index).SetFocus
If bMissingOdds() Then
    Msgbox "Missing Odds!  Please enter Odds for each leg", vbExclamation
    Exit Sub
End If
miBase = Index
Call cmdDistribute_Click
miBase = -1
End Sub

Private Sub cmdDistribute_Click()
Dim dOverround As Double
Dim dTotalBet As Double
Dim dComm As Double
Dim dBaseReturn As Double
Dim sComm As String
Dim i As Integer
txtTotalBet.SetFocus
If bMissingOdds() Then
    Msgbox "Missing Odds!  Please enter Odds for each bet", vbExclamation
    Exit Sub
End If

For i = 0 To miLegs - 1
    dOverround = dOverround + (1 / (Val(txtOdds(i).Text)))
Next

If miBase <> -1 Then
    ' total bet for base button
    sComm = Trim$(txtComm(miBase).Text)
    If sComm = "" Then
        dComm = 0
    ElseIf Right(sComm, 1) = "%" Then
        dComm = Val(left(sComm, Len(sComm) - 1))
    Else
        dComm = Val(sComm)
    End If
    CalcOdds(miBase) = (Val(txtOdds(miBase).Text) - 1) * (1 - (dComm / 100)) + 1
    If InStr(txtStake(miBase).Text, CurrSymbol) Then
            dBaseReturn = Format(Val(Mid(txtStake(miBase).Text, Len(CurrSymbol) + 1)) * CalcOdds(miBase), "0.00")
    Else
            dBaseReturn = Format(Val(txtStake(miBase).Text) * CalcOdds(miBase), "0.00")
    End If
    dTotalBet = dBaseReturn * dOverround
Else
    ' total bet for distribute button
    If InStr(txtTotalBet.Text, CurrSymbol) Then
        dTotalBet = Val(Mid(txtTotalBet.Text, Len(CurrSymbol) + 1))
    Else
        dTotalBet = Val(txtTotalBet.Text)
    End If
End If

For i = 0 To miLegs - 1
    sComm = Trim$(txtComm(i).Text)
    If sComm = "" Then
        dComm = 0
    ElseIf Right(sComm, 1) = "%" Then
        dComm = Val(left(sComm, Len(sComm) - 1))
    Else
        dComm = Val(sComm)
    End If
    CalcOdds(i) = (Val(txtOdds(i).Text) - 1) * (1 - (dComm / 100)) + 1
    If Val(txtOdds(i).Text) = 0 Then
        CalcStake(i) = 0
        txtStake(i).Text = CurrSymbol + "0.00"
    Else
        CalcStake(i) = dTotalBet * ((1 / dOverround)) / CalcOdds(i)
        txtStake(i).Text = CurrSymbol + Format(CalcStake(i), "0.00")
    End If
Next
Call CalcTotalBet
For i = 0 To miLegs - 1
    If InStr(txtStake(i).Text, CurrSymbol) Then
            txtResult(i).Text = CurrSymbol + Format(Val(Mid(txtStake(i).Text, Len(CurrSymbol) + 1)) * CalcOdds(i), "0.00")
            txtProfit(i).Text = CurrSymbol + Format(Val(Mid(txtResult(i).Text, Len(CurrSymbol) + 1)) - Val(Mid(txtTotalBet.Text, Len(CurrSymbol) + 1)), "0.00")
    Else
            txtResult(i).Text = Format(Val(txtStake(i).Text) * CalcOdds(i), "0.00")
            txtProfit(i).Text = Format(Val(txtResult(i).Text) - Val(txtTotalBet.Text), "0.00")
    End If
Next
'round totals?
If cboRound.listindex <> 0 Then Call cboRound_Click
End Sub



Private Sub Form_Load()
'Const Purple As Long = &H800080
'Const Grey As Long = &HC0C0C0
Dim i As Integer, iDefaultRunners As Integer
iDefaultRunners = CInt("0" + GetSetting(app.exename, "Settings", "DefaultRunners"))
With cmbRunners
For i = 2 To 20
    .AddItem (CStr(i) + " Runners")
Next i
End With
With cboRound
    .AddItem "No Rounding"
    .AddItem CurrSymbol + "0.25"
    .AddItem CurrSymbol + "0.5"
    .AddItem CurrSymbol + "1"
    .AddItem CurrSymbol + "5"
End With
miBase = -1

mbLoading = True
cboRound.listindex = 0

cmbRunners.listindex = iDefaultRunners
miLegs = iDefaultRunners + 2
Select Case miLegs
    Case 2: optLegs(0).value = True
    Case 3: optLegs(1).value = True
    Case Else: optLegs(2).value = True
End Select
mbLoading = False

If mbCorporate Then
    shpBox.BorderColor = vbButtonFace
    shpBox.FillColor = vbButtonFace
    fraDutch.ForeColor = vbBlack
    fraDutch.BackColor = vbButtonFace
    lblTotal.BackColor = vbButtonFace
    lblTotal.ForeColor = vbBlack
End If
Call resizeForm
End Sub





Private Sub Form_Terminate()
mnuExit_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mnuExit_Click
End Sub

Private Sub mnuExit_Click()
frmDutch.visible = False 'hide
frmMain.visible = True 'show

End Sub



Private Sub mnuGlobalComm_Click()
Dim GlobalComm As Double
Dim Comm As String
Dim i As Integer
GlobalComm = Val(InputBox("Use this to set commission to same percentage on every outcome", "Global Commission", 5))
Comm = Format(GlobalComm, "0.000") + "%"
For i = 0 To 19 'miLegs - 1
    txtComm(i).Text = Comm
Next i
End Sub

Private Sub mnuReset_Click()
    Dim i As Integer
    For i = 0 To 19
        txtBet(i).Text = ""
        txtOdds(i).Text = ""
        txtComm(i).Text = ""
        txtStake(i).Text = ""
        txtResult(i).Text = ""
        txtProfit(i).Text = ""
    Next i
    Call resizeForm
End Sub

Private Sub mnuNew_Click()
Call ShellExecute(Me.hwnd, "open", app.exename + ".exe", "", "", 4)
Dim i As Integer
i = frmMain.left
If i > 80 Then frmMain.left = i - 80
End Sub

Private Sub optLegs_Click(Index As Integer)
    If mbLoading Then Exit Sub
    If optLegs(0).value Then miLegs = 2: cmbRunners.listindex = 0
    If optLegs(1).value Then miLegs = 3: cmbRunners.listindex = 1
    If optLegs(2).value Then
        cmbRunners.SetFocus
    End If
    Call resizeForm
End Sub
        
Private Sub resizeForm()
    If miLegs < 1 Or miLegs > 20 Then Exit Sub
    Dim i As Integer
    For i = miLegs To 19
        txtBet(i).visible = False
        txtComm(i).visible = False
        txtOdds(i).visible = False
        txtProfit(i).visible = False
        txtResult(i).visible = False
        txtStake(i).visible = False
        cmdBase(i).visible = False
    Next i
    Select Case miLegs
        Case 2:
        txtBet(0).Text = "Bet 1"
        txtBet(1).Text = "Bet 2"
        Case 3:
        txtBet(0).Text = "Home"
        txtBet(1).Text = "Draw (X)"
        txtBet(2).Text = "Away"
        Case Else:
        For i = 0 To 19
            txtBet(i).Text = "Horse " + CStr(i + 1)
        Next i
    End Select
    For i = 2 To miLegs - 1
        txtBet(i).visible = True
        txtComm(i).visible = True
        txtOdds(i).visible = True
        txtProfit(i).visible = True
        txtResult(i).visible = True
        txtStake(i).visible = True
        cmdBase(i).visible = True
    Next i
    shpBox.Height = 7395 - ((20 - miLegs) * 300)
    txtTotalBet.Top = 6960 - ((20 - miLegs) * 300)
    cmdDistribute.Top = 6960 - ((20 - miLegs) * 300)
    cboRound.Top = 6420 - ((20 - miLegs) * 300)
    'cmdPlaceCalc.Top = 6440 - ((20 - miLegs) * 300)
    fraDutch.Height = 6855 - ((20 - miLegs) * 300)
    frmDutch.Height = 8250 - ((20 - miLegs) * 300) + 160
    lblTotal.Top = 6480 - ((20 - miLegs) * 300)
    fraCompare.Top = 7480 - ((20 - miLegs) * 300)
End Sub






Private Sub txtBet_GotFocus(Index As Integer)
If Len(txtBet(Index).Text) Then
    txtBet(Index).SelStart = 0
    txtBet(Index).SelLength = Len(txtBet(Index).Text)
End If
End Sub





Private Sub txtBet_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 0 Then
        frmDutch.caption = "Lay Odds Equalizer - Dutch Calculator - Bet=" + txtBet(0).Text
    End If
End Sub

Private Sub txtComm_GotFocus(Index As Integer)
If Len(txtComm(Index).Text) Then
txtComm(Index).SelStart = 0
txtComm(Index).SelLength = Len(txtComm(Index).Text)
End If
End Sub





Private Sub txtComm_LostFocus(Index As Integer)
Dim sComm As String
sComm = txtComm(Index).Text
If Right(sComm, 1) = "%" Then sComm = left(sComm, Len(sComm) - 1)
If Len(sComm) Then txtComm(Index).Text = Format(Val(sComm), "0.000") + "%"
End Sub









Private Sub txtOdds_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = Asc("\") Then KeyAscii = Asc("/")
End Sub

Private Sub txtOdds_LostFocus(Index As Integer)
txtOdds(Index).Text = ConvertOdds(txtOdds(Index).Text)
End Sub

Private Sub txtStake_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim total As Double
    For i = 0 To miLegs - 2
        If InStr(txtStake(i).Text, CurrSymbol) Then
            total = total + Val(Mid(txtStake(i).Text, Len(CurrSymbol) + 1))
        Else
            total = total + Val(txtStake(i).Text)
        End If
    Next i
    txtTotalBet.Text = CurrSymbol + Format(total, "0.00")
End Sub

Private Sub txtTotalBet_gotfocus()
If Len(txtTotalBet.Text) Then
txtTotalBet.SelStart = 0
txtTotalBet.SelLength = Len(txtTotalBet.Text)
End If
End Sub

Private Sub txtOdds_GotFocus(Index As Integer)
If Len(txtOdds(Index).Text) Then
txtOdds(Index).SelStart = 0
txtOdds(Index).SelLength = Len(txtOdds(Index).Text)
End If
End Sub





Private Sub txtStake_gotfocus(Index As Integer)
If Len(txtStake(Index).Text) Then
txtStake(Index).SelStart = 0
txtStake(Index).SelLength = Len(txtStake(Index).Text)
End If
End Sub

Private Sub txtStake_LostFocus(Index As Integer)
txtStake(Index).Text = Format(txtStake(Index).Text, CurrSymbol + "0.00")
Call CalcTotalBet

End Sub
Private Sub CalcTotalBet()
Dim i As Integer
Dim total As Double
    For i = 0 To miLegs - 1
        If InStr(txtStake(i).Text, CurrSymbol) Then
            total = total + Val(Mid(txtStake(i).Text, Len(CurrSymbol) + 1))
        Else
            total = total + Val(txtStake(i).Text)
        End If
    Next i
txtTotalBet.Text = CurrSymbol + Format(total, "0.00")
End Sub
Private Sub txtTotalBet_LostFocus()
txtTotalBet.Text = Format(txtTotalBet.Text, CurrSymbol + "0.00")
End Sub
