VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lay Odds Equalizer"
   ClientHeight    =   3780
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   6690
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   6690
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLayDesc 
      Height          =   285
      Left            =   120
      TabIndex        =   21
      ToolTipText     =   "Lay Desc"
      Top             =   2700
      Width           =   1155
   End
   Begin VB.TextBox txtBackDesc 
      Height          =   285
      Left            =   180
      TabIndex        =   20
      ToolTipText     =   "Back Desc"
      Top             =   900
      Width           =   1155
   End
   Begin VB.OptionButton optLayPc 
      Caption         =   "5"
      Height          =   255
      Index           =   5
      Left            =   6000
      TabIndex        =   19
      Top             =   240
      Width           =   495
   End
   Begin VB.OptionButton optLayPc 
      Caption         =   "4"
      Height          =   255
      Index           =   4
      Left            =   5400
      TabIndex        =   18
      Top             =   240
      Width           =   495
   End
   Begin VB.OptionButton optLayPc 
      Caption         =   "3"
      Height          =   255
      Index           =   3
      Left            =   4800
      TabIndex        =   17
      Top             =   240
      Width           =   495
   End
   Begin VB.OptionButton optLayPc 
      Caption         =   "2"
      Height          =   255
      Index           =   2
      Left            =   4200
      TabIndex        =   16
      Top             =   240
      Value           =   -1  'True
      Width           =   495
   End
   Begin VB.OptionButton optLayPc 
      Caption         =   "1"
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   15
      Top             =   240
      Width           =   495
   End
   Begin VB.OptionButton optLayPc 
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   13
      Top             =   240
      Width           =   495
   End
   Begin VB.CheckBox chkStakeNotReturned 
      Caption         =   "Check1"
      Height          =   255
      Left            =   1920
      TabIndex        =   12
      Top             =   240
      Width           =   255
   End
   Begin VB.TextBox txtResult 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3075
      Left            =   3000
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   600
      Width           =   3555
   End
   Begin VB.CommandButton cmdNECalc 
      Caption         =   "Not Equal &Calc"
      Height          =   630
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
         Name            =   "MS Sans Serif"
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
         Name            =   "MS Sans Serif"
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
         Name            =   "MS Sans Serif"
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
         Name            =   "MS Sans Serif"
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
      Height          =   630
      Left            =   1560
      TabIndex        =   4
      ToolTipText     =   "Works out best Lay bet"
      Top             =   3000
      Width           =   1230
   End
   Begin VB.Label Label6 
      Caption         =   "Lay %ge"
      Height          =   375
      Left            =   2280
      TabIndex        =   14
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Lay Stake"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2460
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Lay Odds"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1860
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Back Odds"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1260
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Stake Not Returned"
      Height          =   300
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Back Stake "
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   660
      Width           =   1095
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
   Begin VB.Menu mnuComm 
      Caption         =   "&Setup Commission"
      Begin VB.Menu mnuBackPc 
         Caption         =   "Back Commission %"
      End
      Begin VB.Menu mnuLayPc 
         Caption         =   "Lay Commission %"
      End
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



Private Sub chkStakeNotReturned_Click()
If chkStakeNotReturned.Value Then
    bStakeNotReturned = True
Else
    bStakeNotReturned = False
End If
End Sub



Private Sub cmdEqualize_Click()
If LayOdds = 0 Or BackOdds = 0 Then Beep: Exit Sub
Me.MousePointer = vbHourglass
If chkStakeNotReturned.Value Then
    StakeNotReturned = BackStake
Else
    StakeNotReturned = 0
End If
LayStake = BackStake 'seed
Call calc("Equal")
txtResult.Text = result()
txtLayStake = Format(LayStake, "0.00")
Me.MousePointer = vbDefault
End Sub

Private Sub cmdNECalc_Click()
If LayOdds = 0 Or BackOdds = 0 Or LayStake = 0 Then Beep: Exit Sub
Me.MousePointer = vbHourglass
If chkStakeNotReturned.Value Then
    StakeNotReturned = BackStake
Else
    StakeNotReturned = 0
End If
Call calc("NotEqual")
txtResult.Text = result()
txtLayStake = Format(LayStake, "0.00")
Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
LayPc = 2#
BackPc = 0#
mnuBackPc.Caption = "Back Commission " + CStr(BackPc) + "%"
mnuLayPc.Caption = "Lay Commission " + CStr(LayPc) + "%"
chkStakeNotReturned.Value = 0
cmdEqualize.Enabled = True
cmdNECalc.Enabled = True
End Sub





Private Sub mnuAbout_Click()
MsgBox "Lay Odds Equalizer v" + CStr(App.Major) + "." + CStr(App.Minor) + vbCrLf + "(c) 2014, 2015 David J. Barnes, London" + vbCrLf + "All rights reserved" + vbCrLf + vbCrLf + "Note: use of this application is entirely at your own risk, gamble responsibly" + vbCrLf + vbCrLf + "Visit http://layoddsequalizer.uk", vbInformation + vbOKOnly, "About Lay Odds Equalizer"

End Sub

Private Sub mnuBackPc_Click()
BackPc = Val(InputBox("Back Commission Percent ?", , BackPc))
mnuBackPc.Caption = "Back Commission " + CStr(BackPc) + "%"
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuLayPc_Click()
LayPc = Val(InputBox("Lay Commission Percent ?", , LayPc))
mnuLayPc.Caption = "Lay Commission " + CStr(LayPc) + "%"
End Sub






Private Sub mnuNew_Click()
txtBackDesc.Text = ""
txtLayDesc.Text = ""
txtBackOdds.Text = ""
txtBackStake.Text = ""
txtLayOdds.Text = ""
txtLayStake.Text = ""
End Sub

Private Sub optLayPc_Click(Index As Integer)
LayPc = Index
mnuLayPc.Caption = "Lay Commission " + CStr(LayPc) + "%"
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
End Sub



Private Sub txtBackDesc_GotFocus()
If Len(txtBackDesc.Text) Then
txtBackDesc.SelStart = 0
txtBackDesc.SelLength = Len(txtBackDesc.Text)
End If
End Sub

Private Sub txtLayDesc_GotFocus()
If Len(txtLayDesc.Text) Then
txtLayDesc.SelStart = 0
txtLayDesc.SelLength = Len(txtLayDesc.Text)
End If
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
