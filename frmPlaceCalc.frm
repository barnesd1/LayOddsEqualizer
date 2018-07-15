VERSION 5.00
Begin VB.Form frmPlaceCalc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lay Odds Equalizer -  Place Profit Calculator"
   ClientHeight    =   3960
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   6240
   Icon            =   "frmPlaceCalc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraTotals 
      BackColor       =   &H00C0C0C0&
      Caption         =   "  Total Bet                      Total Returned           Added BOG         PROFIT  "
      ClipControls    =   0   'False
      ForeColor       =   &H00800080&
      Height          =   615
      Left            =   120
      TabIndex        =   12
      Top             =   3240
      Width           =   5955
      Begin VB.TextBox txtBOG 
         Height          =   285
         Left            =   3420
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.ComboBox cboHorse 
      Height          =   315
      Index           =   4
      Left            =   1680
      TabIndex        =   11
      Top             =   2820
      Width           =   1815
   End
   Begin VB.ComboBox cboHorse 
      Height          =   315
      Index           =   3
      Left            =   1680
      TabIndex        =   10
      Top             =   2460
      Width           =   1815
   End
   Begin VB.ComboBox cboHorse 
      Height          =   315
      Index           =   2
      Left            =   1680
      TabIndex        =   9
      Top             =   2100
      Width           =   1815
   End
   Begin VB.ComboBox cboHorse 
      Height          =   315
      Index           =   1
      Left            =   1680
      TabIndex        =   8
      Top             =   1740
      Width           =   1815
   End
   Begin VB.ComboBox cboHorse 
      Height          =   315
      Index           =   0
      Left            =   1680
      TabIndex        =   5
      Top             =   1380
      Width           =   1815
   End
   Begin VB.TextBox txtTotalLay 
      Height          =   285
      Left            =   1680
      TabIndex        =   4
      Top             =   660
      Width           =   975
   End
   Begin VB.TextBox txtTotalBet 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   300
      Width           =   975
   End
   Begin VB.ComboBox cboPlaces 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   1020
      Width           =   1215
   End
   Begin VB.Shape shpBox 
      BorderColor     =   &H00800080&
      BorderWidth     =   10
      FillColor       =   &H00800080&
      Height          =   3960
      Left            =   0
      Top             =   0
      Width           =   6240
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fifth Place"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   7
      Left            =   360
      TabIndex        =   16
      Top             =   2880
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fourth Place"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   6
      Left            =   360
      TabIndex        =   15
      Top             =   2520
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Third Place"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   5
      Left            =   360
      TabIndex        =   14
      Top             =   2160
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Second Place"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   360
      TabIndex        =   13
      Top             =   1800
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "First Place"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   360
      TabIndex        =   7
      Top             =   1440
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Places Paid"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   360
      TabIndex        =   6
      Top             =   1080
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total Lay Bet"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   660
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total Back Bet"
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   300
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuReset 
      Caption         =   "Reset"
   End
   Begin VB.Menu mnuPayout 
      Caption         =   "Global Payout"
   End
End
Attribute VB_Name = "frmPlaceCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim miPlaces As Integer
Dim mbLoading As Integer
Private Sub resizeForm()

If miPlaces < 2 Or miPlaces > 5 Then Exit Sub
    Dim i As Integer
    For i = miPlaces + 1 To 5
        lblTotal(i + 2).Visible = False
        cboHorse(i - 1).Visible = False
    Next
    For i = 2 To miPlaces
        lblTotal(i + 2).Visible = True
        cboHorse(i - 1).Visible = True
    Next
    fraTotals.Top = 3240 - ((5 - miPlaces) * 300)
    shpBox.Height = 3960 - ((5 - miPlaces) * 300)
    frmPlaceCalc.Height = 4695 - ((5 - miPlaces) * 300)
End Sub

Private Sub cboPlaces_Click()
    If mbLoading Then Exit Sub
    miPlaces = cboPlaces.ListIndex + 2
    Call resizeForm
End Sub



Private Sub Form_Load()
Dim i As Integer
With cboPlaces
For i = 2 To 5
    .AddItem CStr(i) + " Places"
Next
.ListIndex = 0
End With
miPlaces = 2
mbLoading = True
cboPlaces.ListIndex = 0
mbLoading = False
If mbCorporate Then
    fraTotals.ForeColor = vbBlack
    fraTotals.BackColor = vbButtonFace
End If
Call resizeForm
End Sub
