VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lay Odds Equalizer"
   ClientHeight    =   4965
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   6630
   BeginProperty Font 
   EndProperty
   Font            =   "frmMain.frx":0000
   Icon            =   "frmMain.frx":0018
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   6630
   Begin VB.PictureBox picDutch 
      AutoSize        =   -1  'True
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":045A
      Height          =   1260
      Left            =   60
      Picture         =   "frmMain.frx":0472
      ScaleHeight     =   1200
      ScaleWidth      =   330
      TabIndex        =   32
      Top             =   1260
      Width           =   390
   End
   Begin VB.Frame fraFull 
      BackColor       =   &H00C0C0C0&
      Caption         =   "   Under Lay                                   Quick %ge  Slider                        Over Lay"
      ClipControls    =   0   'False
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":0AC0
      ForeColor       =   &H00800080&
      Height          =   1215
      Left            =   120
      TabIndex        =   21
      Top             =   3600
      Width           =   6375
      Begin VB.CommandButton cmdClearHistory 
         Caption         =   "CLEAR"
         BeginProperty Font 
         EndProperty
         Font            =   "frmMain.frx":0AD8
         Height          =   210
         Left            =   1620
         TabIndex        =   31
         Top             =   960
         Width           =   735
      End
      Begin VB.PictureBox picUpgrade 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BeginProperty Font 
         EndProperty
         Font            =   "frmMain.frx":0AF2
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   0
         Picture         =   "frmMain.frx":0B0A
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   29
         ToolTipText     =   "Shareware: Unlock full version today1"
         Top             =   240
         Width           =   390
      End
      Begin VB.CommandButton cmdLoseZeroBack 
         Caption         =   "Lose 0 on &Back"
         BeginProperty Font 
         EndProperty
         Font            =   "frmMain.frx":120C
         Height          =   560
         Left            =   240
         TabIndex        =   26
         ToolTipText     =   "Works out best Lay bet"
         Top             =   375
         Width           =   1230
      End
      Begin VB.CommandButton cmdLoseZeroLay 
         Caption         =   "Lose 0 on &Lay"
         BeginProperty Font 
         EndProperty
         Font            =   "frmMain.frx":1224
         Height          =   560
         Left            =   4920
         TabIndex        =   25
         ToolTipText     =   "Works out best Lay bet"
         Top             =   375
         Width           =   1230
      End
      Begin VB.Image imgHistory 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Index           =   3
         Left            =   4140
         MousePointer    =   8  'Size NW SE
         Picture         =   "frmMain.frx":123C
         Top             =   600
         Width           =   570
      End
      Begin VB.Image imgHistory 
         BorderStyle     =   1  'Fixed Single
         Height          =   510
         Index           =   2
         Left            =   3600
         MousePointer    =   7  'Size N S
         Picture         =   "frmMain.frx":1825
         Top             =   600
         Width           =   540
      End
      Begin VB.Image imgHistory 
         BorderStyle     =   1  'Fixed Single
         Height          =   510
         Index           =   1
         Left            =   3060
         MousePointer    =   9  'Size W E
         Picture         =   "frmMain.frx":1CB1
         Top             =   600
         Width           =   540
      End
      Begin VB.Image imgHistory 
         BorderStyle     =   1  'Fixed Single
         Height          =   510
         Index           =   0
         Left            =   2520
         MousePointer    =   6  'Size NE SW
         Picture         =   "frmMain.frx":216F
         Top             =   600
         Width           =   540
      End
      Begin VB.Label lblHistory 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Bet History"
         BeginProperty Font 
         EndProperty
         Font            =   "frmMain.frx":269F
         ForeColor       =   &H00800080&
         Height          =   435
         Left            =   1570
         TabIndex        =   30
         Top             =   600
         Width           =   850
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblO 
         BackStyle       =   0  'Transparent
         Caption         =   ">1%"
         BeginProperty Font 
         EndProperty
         Font            =   "frmMain.frx":26B7
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   1
         Left            =   3960
         TabIndex        =   28
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblU 
         BackStyle       =   0  'Transparent
         Caption         =   "<0.5%"
         BeginProperty Font 
         EndProperty
         Font            =   "frmMain.frx":26C7
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   2520
         TabIndex        =   27
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblO 
         BackStyle       =   0  'Transparent
         Caption         =   ">0.5%"
         BeginProperty Font 
         EndProperty
         Font            =   "frmMain.frx":26D7
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   3360
         TabIndex        =   24
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblEqual 
         BackStyle       =   0  'Transparent
         Caption         =   "="
         BeginProperty Font 
         EndProperty
         Font            =   "frmMain.frx":26E7
         ForeColor       =   &H00800080&
         Height          =   255
         Left            =   3120
         TabIndex        =   23
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblU 
         BackStyle       =   0  'Transparent
         Caption         =   "<1%"
         BeginProperty Font 
         EndProperty
         Font            =   "frmMain.frx":26F7
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   22
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.TextBox txtLayDesc 
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":2707
      Height          =   285
      Left            =   180
      TabIndex        =   20
      ToolTipText     =   "Lay Desc"
      Top             =   2700
      Width           =   1155
   End
   Begin VB.TextBox txtBackDesc 
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":271F
      Height          =   285
      Left            =   180
      TabIndex        =   19
      ToolTipText     =   "Back Desc"
      Top             =   900
      Width           =   1155
   End
   Begin VB.OptionButton optLayPc 
      Caption         =   "5.0"
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":2737
      Height          =   255
      Index           =   4
      Left            =   5760
      TabIndex        =   18
      ToolTipText     =   "Right-Click to custom"
      Top             =   240
      Width           =   615
   End
   Begin VB.OptionButton optLayPc 
      Caption         =   "3.0"
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":274F
      Height          =   255
      Index           =   3
      Left            =   5040
      TabIndex        =   17
      ToolTipText     =   "Right-Click to custom"
      Top             =   240
      Width           =   615
   End
   Begin VB.OptionButton optLayPc 
      Caption         =   "2.0"
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":2767
      Height          =   255
      Index           =   2
      Left            =   4320
      TabIndex        =   16
      ToolTipText     =   "Right-Click to custom"
      Top             =   240
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.OptionButton optLayPc 
      Caption         =   "1.0"
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":277F
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   15
      ToolTipText     =   "Right-Click to custom"
      Top             =   240
      Width           =   615
   End
   Begin VB.OptionButton optLayPc 
      Caption         =   "0"
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":2797
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   13
      ToolTipText     =   "Fixed at Zero"
      Top             =   240
      Width           =   495
   End
   Begin VB.CheckBox chkStakeNotReturned 
      Caption         =   "Check1"
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":27AF
      Height          =   255
      Left            =   1920
      TabIndex        =   12
      Top             =   240
      Width           =   255
   End
   Begin VB.TextBox txtResult 
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":27C7
      Height          =   3015
      Left            =   3000
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   540
      Width           =   3495
   End
   Begin VB.CommandButton cmdNECalc 
      Caption         =   "Not Equal &Calc"
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":27D7
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
      EndProperty
      Font            =   "frmMain.frx":27EF
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
      EndProperty
      Font            =   "frmMain.frx":27FF
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
      EndProperty
      Font            =   "frmMain.frx":280F
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
      EndProperty
      Font            =   "frmMain.frx":281F
      Height          =   465
      Left            =   1500
      TabIndex        =   0
      Top             =   615
      Width           =   1270
   End
   Begin VB.CommandButton cmdEqualize 
      Caption         =   "&Equalize"
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":282F
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
      Height          =   4905
      Left            =   20
      Top             =   40
      Width           =   6615
   End
   Begin VB.Label Label6 
      Caption         =   "Lay %ge"
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":2847
      Height          =   375
      Left            =   2280
      TabIndex        =   14
      Top             =   240
      Width           =   735
   End
   Begin VB.Label lblLayStake 
      Alignment       =   1  'Right Justify
      Caption         =   "Lay Stake"
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":285F
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2460
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Lay Odds"
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":2877
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1860
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Back Odds"
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":288F
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1260
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Stake Not Returned"
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":28A7
      Height          =   300
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lblBackStake 
      Alignment       =   1  'Right Justify
      Caption         =   "Back Stake "
      BeginProperty Font 
      EndProperty
      Font            =   "frmMain.frx":28BF
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
      End
      Begin VB.Menu mnuLayPc 
         Caption         =   "Override Lay Commission %"
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
      Begin VB.Menu mnuLn3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMaxI 
         Caption         =   "Max Calc Iterations"
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
   Begin VB.Menu mnuUpgrade 
      Caption         =   "&Unlock Full Features"
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

Private Sub chkStakeNotReturned_Click()
If chkStakeNotReturned.value Then
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
Me.MousePointer = vbHourglass
If chkStakeNotReturned.value Then
    StakeNotReturned = BackStake
Else
    StakeNotReturned = 0
End If
LayStake = BackStake 'seed
Call calc("Equal")
txtResult.Text = result()
sCompare = compare()
txtLayStake = Format(LayStake, "0.00")
If Not mbCalcing Then Call AddToHistory
Me.MousePointer = vbDefault
End Sub

Private Sub cmdLoseZeroBack_Click()
If picUpgrade.visible Then mnuAbout_Click: Exit Sub
If BackStake = 0 Or LayOdds = 0 Or BackOdds = 0 Then Beep: Exit Sub

mbCalcing = True
cmdEqualize.value = True

If BetfairBackCost = 0 Then Beep: mbCalcing = False: Exit Sub
If BetfairBackCost > 0.01 Then
    Do
        LayStake = LayStake + 0.01
        txtLayStake.Text = Format(LayStake, "0.00")
        cmdNECalc.value = True
    Loop While Val(Format(BetfairBackCost, "0.00")) > 0
Else
    Do
        LayStake = LayStake - 0.01
        txtLayStake.Text = Format(LayStake, "0.00")
        cmdNECalc.value = True
    Loop While Val(Format(BetfairBackCost, "0.00")) < 0
End If

mbCalcing = False
Call AddToHistory
End Sub

Private Sub cmdNECalc_Click()
If BackStake = 0 Or LayOdds = 0 Or BackOdds = 0 Or LayStake = 0 Then Beep: Exit Sub
Me.MousePointer = vbHourglass
If chkStakeNotReturned.value Then
    StakeNotReturned = BackStake
Else
    StakeNotReturned = 0
End If
Call calc("NotEqual")
txtResult.Text = result()
txtLayStake.Text = Format(LayStake, "0.00")
If Not mbCalcing Then Call AddToHistory
Me.MousePointer = vbDefault
End Sub

Private Sub cmdLoseZeroLay_Click()
If picUpgrade.visible Then mnuAbout_Click: Exit Sub
If BackStake = 0 Or LayOdds = 0 Or BackOdds = 0 Then Beep: Exit Sub

mbCalcing = True
cmdEqualize.value = True

If BetfairLayCost = 0 Then Beep: mbCalcing = False: Exit Sub

If BetfairLayCost > 0.01 Then
    Do
        LayStake = LayStake - 0.01
        txtLayStake.Text = Format(LayStake, "0.00")
        cmdNECalc.value = True
    Loop While Val(Format(BetfairLayCost, "0.00")) > 0
Else
    Do
        LayStake = LayStake + 0.01
        txtLayStake.Text = Format(LayStake, "0.00")
        cmdNECalc.value = True
    Loop While Val(Format(BetfairLayCost, "0.00")) < 0
End If
mbCalcing = False
Call AddToHistory
End Sub

Private Sub Form_Load()
If app.Revision = 99 Then
    mnuUpgrade.visible = True
    picUpgrade.visible = True
Else
    mnuUpgrade.visible = False
    picUpgrade.visible = False
End If

LayPc = 2#
BackPc = 0#
mnuBackPc.caption = "Back Commission " + CStr(BackPc) + "%"
mnuLayPc.caption = "Lay Commission " + CStr(LayPc) + "%"
chkStakeNotReturned.value = 0
cmdEqualize.Enabled = True
cmdNECalc.Enabled = True
mbLoading = True
sCompare = ""
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
If picUpgrade.visible Then Call mnuAbout_Click
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
    SetLayPcOpt
    txtBackStake.Text = ConvertOdds(Val(iHistBackStake(iMoveTo)))
    BackStake = iHistBackStake(iMoveTo)
    txtBackOdds.Text = ConvertOdds(Val(iHistBackOdds(iMoveTo)))
    BackOdds = iHistBackOdds(iMoveTo)
    txtLayStake.Text = ConvertOdds(Val(iHistLayStake(iMoveTo)))
    LayStake = iHistLayStake(iMoveTo)
    txtLayOdds.Text = ConvertOdds(Val(iHistLayOdds(iMoveTo)))
    LayOdds = iHistLayOdds(iMoveTo)
    chkStakeNotReturned.value = Abs(bHistSNR(iMoveTo))
    Call chkStakeNotReturned_Click
    
    lblHistory.caption = "Bet History " + CStr(iHistoryPosition) + " of " + CStr(iHistoryUsed)
    mbCalcing = True
    cmdNECalc.SetFocus
    cmdNECalc.value = True
    mbCalcing = False
End If
End Sub
Public Sub SetLayPcOpt()
    Dim i As Integer
    For i = 0 To 4
        If Val(optLayPc(i).caption) = LayPc Then
            optLayPc(i).value = True
            Exit Sub
        End If
    Next
End Sub
Public Sub AddToHistory()
If iHistoryUsed >= 100 Then
    If Msgbox("Maximum 100 history reached, click Ok to make to clear the history and make this item 1, or Cancel to make this item 100", vbOKCancel, "History full") = vbOk Then
        iHistoryUsed = 0
    Else
        iHistoryUsed = 99
    End If
End If
Dim i As Integer
If iHistoryUsed > 0 Then
    i = iHistoryUsed
    If iHistBackPc(i) = BackPc And _
        iHistLayPc(i) = LayPc And _
        iHistBackStake(i) = BackStake And _
        iHistBackOdds(i) = BackOdds And _
        iHistLayStake(i) = LayStake And _
        iHistLayOdds(i) = LayOdds And _
        bHistSNR(i) = bStakeNotReturned Then
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
lblHistory.caption = "Bet History " + CStr(iHistoryPosition) + " of " + CStr(iHistoryUsed)
End Sub

Private Sub lblEqual_Click()
If picUpgrade.visible Then picUpgrade_Click: Exit Sub

cmdEqualize.SetFocus
cmdEqualize.value = True

End Sub

Private Sub lblO_Click(Index As Integer)
If picUpgrade.visible Then picUpgrade_Click: Exit Sub
Dim mult As Single
Select Case (Index)
Case 0: mult = 1.005
Case 1: mult = 1.01
End Select

txtLayStake.Text = ConvertOdds(LayStake * mult)
LayStake = Val(txtLayStake.Text)
cmdNECalc.SetFocus
cmdNECalc.value = True
End Sub

Private Sub lblU_Click(Index As Integer)
If picUpgrade.visible Then picUpgrade_Click: Exit Sub
Dim mult As Single
Select Case (Index)
Case 0: mult = 0.995
Case 1: mult = 0.99
End Select

txtLayStake.Text = ConvertOdds(LayStake * mult)
LayStake = Val(txtLayStake.Text)
cmdNECalc.SetFocus
cmdNECalc.value = True
End Sub


Private Sub mnuAbout_Click()
Dim sMsg As String
Dim iResult As Integer
sMsg = "Lay Odds Equalizer v" + CStr(app.Major) + "." + CStr(app.Minor) + "." + CStr(app.Revision)
If app.Revision = 99 Then
    sMsg = sMsg + " TRIAL"
Else
    sMsg = sMsg + " FULL"
End If
sMsg = sMsg + " Edition"
sMsg = sMsg + vbCrLf + "Written by David J. Barnes, London, UK" + vbCrLf + "Released under public license at GitHub" + vbCrLf + vbCrLf + "Note: use of this application is entirely at your own risk, gamble responsibly" + vbCrLf + vbCrLf + "Click Yes to Visit https://github.com/barnesd1/LayOddsEqualizer" + vbCrLf + "Click No to Donate £5.00 towards development of this application at http://paypal.me/LayOddsEqualizer/5"
iResult = Msgbox(sMsg, vbInformation + vbYesNoCancel, "About Lay Odds Equalizer")
If iResult = vbYes Then Call ShellExecute(Me.hwnd, "open", "https://github.com/barnesd1/LayOddsEqualizer", "", "", 4)
If iResult = vbNo Then Call ShellExecute(Me.hwnd, "open", "http://paypal.me/LayOddsEqualizer/5", "", "", 4)
End Sub

Private Sub mnuBackPc_Click()
BackPc = Val(InputBox("Back Commission Percent ?", , BackPc))
mnuBackPc.caption = "Back Commission " + CStr(BackPc) + "%"
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
Msgbox "You will be paying " + CStr(BF) + " * (100 - " + CStr(BFDiscount) + ") / 100 so you should use" + vbCrLf + CStr(BFLayCom) + "% in your Betfair lay calculations", vbExclamation + vbOkOnly
End Sub

Private Sub mnuCorporateColours_Click()
Const Purple As Long = &H800080
Const Grey As Long = &HC0C0C0
If mnuCorporateColours.Checked Then
    mnuCorporateColours.Checked = False
    mbCorporate = False
    shpBox.BorderColor = Purple
    shpBox.FillColor = Purple
    fraFull.ForeColor = Purple
    fraFull.BackColor = Grey
    lblEqual.ForeColor = Purple
    lblHistory.BackColor = Grey
    lblHistory.ForeColor = Purple
    lblU(1).ForeColor = vbBlue
    lblU(0).ForeColor = vbBlue
    lblO(1).ForeColor = vbRed
    lblO(0).ForeColor = vbRed
Else
    mnuCorporateColours.Checked = True
    mbCorporate = True
    shpBox.BorderColor = vbButtonFace
    shpBox.FillColor = vbButtonFace
    fraFull.ForeColor = vbBlack
    fraFull.BackColor = vbButtonFace
    lblEqual.ForeColor = vbBlack
    lblHistory.ForeColor = vbBlack
    lblHistory.BackColor = vbButtonFace
    lblU(1).ForeColor = vbBlack
    lblU(0).ForeColor = vbBlack
    lblO(1).ForeColor = vbBlack
    lblO(0).ForeColor = vbBlack
End If
If mnuCorporateColours.Checked Then
    SaveSetting app.exename, "Colours", "Corporate", "YES"
Else
    SaveSetting app.exename, "Colours", "Corporate", "NO"
End If
End Sub

Private Sub mnuCurrSymbol_Click()
Dim sCurrSymbol As String
sCurrSymbol = InputBox("Enter Currency Symbol ?", , CurrSymbol)
If Len(sCurrSymbol) > 3 Then Msgbox "Maximum currency symbol length is 3": Exit Sub
If sCurrSymbol <> CurrSymbol Then
   CurrSymbol = sCurrSymbol
   UpdateCurrency
   
End If
End Sub
Sub UpdateCurrency()
    cmdLoseZeroBack.caption = "Lose " + CurrSymbol + "0 on &Back"
    cmdLoseZeroLay.caption = "Lose " + CurrSymbol + "0 on &Lay"
    lblBackStake.caption = "Back Stake " + CurrSymbol
    lblLayStake.caption = "Lay Stake " + CurrSymbol
End Sub
Private Sub mnuExit_Click()
    SaveSettings
    End
End Sub

Private Sub mnuLayPc_Click()
LayPc = Val(InputBox("Lay Commission Percent ?", , LayPc))
mnuLayPc.caption = "Lay Commission " + CStr(LayPc) + "%"
End Sub




Private Sub SaveSettings()
SaveSetting app.exename, "Settings", "Currency", CurrSymbol
SaveSetting app.exename, "Settings", "LayOdds1", optLayPc(1).caption
SaveSetting app.exename, "Settings", "LayOdds2", optLayPc(2).caption
SaveSetting app.exename, "Settings", "LayOdds3", optLayPc(3).caption
SaveSetting app.exename, "Settings", "LayOdds4", optLayPc(4).caption
SaveSetting app.exename, "Position", "Top", CStr(frmMain.Top)
SaveSetting app.exename, "Position", "Left", CStr(frmMain.left)
SaveSetting app.exename, "Settings", "Iterations", CStr(lMaxIterations)
SaveSetting app.exename, "Settings", "DefaultPercent", CStr(LayPc)
End Sub
Private Sub ReadSettings()
CurrSymbol = left(GetSetting(app.exename, "Settings", "Currency", "£"), 3)
optLayPc(1).caption = CStr(Val(GetSetting(app.exename, "Settings", "LayOdds1")))
optLayPc(2).caption = CStr(Val(GetSetting(app.exename, "Settings", "LayOdds2")))
optLayPc(3).caption = CStr(Val(GetSetting(app.exename, "Settings", "LayOdds3")))
optLayPc(4).caption = CStr(Val(GetSetting(app.exename, "Settings", "LayOdds4")))
lMaxIterations = Val(GetSetting(app.exename, "Settings", "Iterations"))
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
    Call mnuMax_Click(3)
End Select

Dim dDefaultPercent As Double
dDefaultPercent = Val(GetSetting(app.exename, "Settings", "DefaultPercent"))
Dim i As Integer
Dim dTotal As Double, dThis As Double
dTotal = 0
For i = 0 To 4
    dThis = Val(optLayPc(i).caption)
    If dThis = dDefaultPercent Then optLayPc(i).value = True
    dTotal = dTotal + dThis
Next i
If dTotal = 0 Then
    ' reset to default
    optLayPc(1).caption = "1.15"
    optLayPc(2).caption = "2"
    optLayPc(3).caption = "4.7"
    optLayPc(4).caption = "5"
End If
frmMain.Top = Val(GetSetting(app.exename, "Position", "Top", "1000"))
frmMain.left = Val(GetSetting(app.exename, "Position", "Left", "1000"))
mnuCorporateColours.Checked = False
If GetSetting(app.exename, "Colours", "Corporate") = "YES" Then
    mnuCorporateColours_Click
Else
    mbCorporate = False
End If

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
Call ShellExecute(Me.hwnd, "open", app.exename + ".exe", "", "", 4)
Dim i As Integer
i = frmMain.left
If i > 80 Then frmMain.left = i - 80
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
sCompare = ""
End Sub


Private Sub mnuUpgrade_Click()
picUpgrade_Click
End Sub

Private Sub optLayPc_Click(Index As Integer)
LayPc = Val(optLayPc(Index).caption)
mnuLayPc.caption = "Lay Commission " + CStr(LayPc) + "%"
End Sub



Private Sub optLayPc_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim d As Double
If Button = vbRightButton Then
    If Index > 0 Then
        d = Val(InputBox("Enter Custom Lay Percentage", "Lay %ge"))
        If d > 0 And d < 100 Then
            optLayPc(Index).caption = CStr(d)
            optLayPc(Index).SetFocus
            LayPc = d
            optLayPc(Index).value = True
            Call optLayPc_Click(Index)
        End If
    Else
        Msgbox "First Percentage is fixed at zero percent", vbOkOnly + vbInformation, "Custom Lay Percentage"
    End If
End If
End Sub



Private Sub picDutch_Click()
frmMain.visible = False 'hide

If sCompare = "" Then
    frmDutch.lblCompare.caption = "Comparison Bet: None Set"
Else
    frmDutch.lblCompare.caption = "Compare : " & sCompare
End If
frmDutch.Top = frmMain.Top
frmDutch.left = frmMain.left
frmDutch.show

End Sub

Private Sub picUpgrade_Click()
If picUpgrade.visible Then Call mnuAbout_Click

End Sub


Private Sub txtBackDesc_KeyUp(KeyCode As Integer, Shift As Integer)
frmMain.caption = "Lay Odds Equalizer - Bet=" + txtBackDesc.Text
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
