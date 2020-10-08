Attribute VB_Name = "Module1"
Option Explicit
'VB Constants
Public Const vbBlue As Long = 16711680
Public Const vbRed As Long = 255
Public Const vbBlack As Long = 0
Public Const vbButtonFace As Long = -2147483633
Public Const vbHourglass As Integer = 11
Public Const vbDefault As Integer = 0
Public Const vbRightButton As Integer = 2

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long

Public bCorporate As Boolean
Public LayPc As Single
Public BackPc As Single
Public PlacePc As Single
Public bStakeNotReturned As Boolean
Public bEWMode As Boolean
Public BackStake As Single
Public BackOdds As Single
Public PlaceBackOdds As Single
Public LayOdds As Single
Public LayStake As Single
Public PlaceLayStake As Single
Public PlaceLayOdds As Single
Public BackReturn As Single
Public PlaceBackReturn As Single
Public PlaceLayRisk As Single
Public LayRisk As Single
Public BackProfit As Single
Public PlaceBackProfit As Single
Public LayProfit As Single
Public PlaceLayProfit As Single
Public Difference As Single
Public BetfairBackCost As Single
Public BetfairLayCost As Single
Public PlaceBetfairBackCost As Single
Public PlaceBetfairLayCost As Single
Public ExtraPlace As Single
Public StakeNotReturned As Single
Public RetentionBack As Single
Public RetentionLay As Single
Public CurrSymbol As String
Public lMaxIterations As Long
Public sHistBackDesc(1 To 100) As String
Public sHistLayDesc(1 To 100) As String
Public iHistBackPc(1 To 100) As Single
Public iHistLayPc(1 To 100) As Single
Public iHistBackStake(1 To 100) As Single
Public iHistBackOdds(1 To 100) As Single
Public iHistLayStake(1 To 100) As Single
Public iHistLayOdds(1 To 100) As Single
Public bHistSNR(1 To 100) As Boolean
Public iHistTerms(1 To 100) As Integer
Public iHistPlacePc(1 To 100) As Single
Public iHistPlaceLayStake(1 To 100) As Single
Public iHistPlaceLayOdds(1 To 100) As Single
Public bEW(1 To 100) As Boolean
Public iHistoryPosition As Integer
Public iHistoryUsed As Integer

Function compare() As String
'comparison bet text
Dim s As String
s = "Back " + CurrSymbol + Format(BackStake, "0.00") + ", Back Odds " + Format(BackOdds, "0.00") + ", Lay Odds " + Format(LayOdds, "0.00") + ", Lay " + CurrSymbol + Format(LayStake, "0.00")
If BetfairLayCost > 0 Then
    s = s + ", Profit " + CurrSymbol + Format(BetfairLayCost, "0.00")
Else
    s = s + ", Cost " + CurrSymbol + Format(BetfairLayCost, "0.00")
End If
compare = s
End Function
Function ConvertOdds(sTyped As String) As String
Dim iSlash As Integer
Dim iLeft As Double
Dim iRight As Double
Dim dOdds As Double
iSlash = InStr(sTyped, "/")
If iSlash Then
    If iSlash = 1 Then
        Beep
    End If
    iLeft = Val(Left(sTyped, iSlash))
    iRight = Val(Mid(sTyped, iSlash + 1, Len(sTyped) - iSlash))
    If iRight = 0 Then iRight = 1
    dOdds = (iLeft / iRight) + 1
    
    ConvertOdds = Format(dOdds, "0.0#")
Else
    dOdds = Val(sTyped)
    ConvertOdds = Format(dOdds, "0.0##")
End If
End Function
Function EWresult() As String
EWresult = "Stake Not Returned = " + IIf(StakeNotReturned, "YES", "NO") + vbCrLf
EWresult = EWresult + "Win Back Commission = " + CStr(BackPc) + "%" + vbCrLf
EWresult = EWresult + "Win Lay Commission = " + CStr(LayPc) + "%" + vbCrLf
EWresult = EWresult + "Place Lay Commission = " + CStr(PlacePc) + "%" + vbCrLf
'ewresult = ewresult + "Back Return = " + CurrSymbol + Format(BackReturn, "0.00") + vbCrLf
EWresult = EWresult + "Win Lay Risk = " + CurrSymbol + Format(LayRisk, "0.00") + vbCrLf
EWresult = EWresult + "Place Lay Risk = " + CurrSymbol + Format(PlaceLayRisk, "0.00") + vbCrLf
EWresult = EWresult + "Bookie Total Required = " + CurrSymbol + Format(BackStake * 2, "0.00") + vbCrLf
EWresult = EWresult + "Exchange Total Required = " + CurrSymbol + Format(LayRisk + PlaceLayRisk, "0.00") + vbCrLf
'ewresult = ewresult + "Back Profit = " + CurrSymbol + Format(BackProfit, "0.00") + vbCrLf
'ewresult = ewresult + "Lay Profit = " + CurrSymbol + Format(LayProfit, "0.00") + vbCrLf
EWresult = EWresult + "Horse Wins = " + CurrSymbol + Format(BetfairBackCost + PlaceBetfairBackCost, "0.00") + vbCrLf
EWresult = EWresult + "Horse Places = " + CurrSymbol + Format(BetfairLayCost + PlaceBetfairBackCost, "0.00") + vbCrLf
EWresult = EWresult + "Horse Loses = " + CurrSymbol + Format(BetfairLayCost + PlaceBetfairLayCost, "0.00") + vbCrLf
EWresult = EWresult + "Extra Place = " + CurrSymbol + Format(ExtraPlace, "0.00") + vbCrLf

End Function

Function result() As String
result = "Stake Not Returned = " + IIf(StakeNotReturned, "YES", "NO") + vbCrLf
result = result + "Back Commission = " + CStr(BackPc) + "%" + vbCrLf
result = result + "Lay Commission = " + CStr(LayPc) + "%" + vbCrLf
result = result + "Back Return = " + CurrSymbol + Format(BackReturn, "0.00") + vbCrLf
result = result + "Lay Risk = " + CurrSymbol + Format(LayRisk, "0.00") + vbCrLf
result = result + "Back Profit = " + CurrSymbol + Format(BackProfit, "0.00") + vbCrLf
result = result + "Lay Profit = " + CurrSymbol + Format(LayProfit, "0.00") + vbCrLf
Difference = Abs(Val(Format(BackProfit, "0.00")) - Val(Format(LayProfit, "0.00")))
result = result + "Difference = " + CurrSymbol + Format(Difference, "0.00") + vbCrLf
If BetfairBackCost > 0 Then
    result = result + "Exchange Back Profit = " + CurrSymbol + Format(BetfairBackCost, "0.00") + vbCrLf
Else
    result = result + "Exchange Back Cost = " + CurrSymbol + Format(BetfairBackCost, "0.00") + vbCrLf
End If
If BetfairLayCost > 0 Then
    result = result + "Exchange Lay Profit = " + CurrSymbol + Format(BetfairLayCost, "0.00") + vbCrLf
Else
    result = result + "Exchange Lay Cost = " + CurrSymbol + Format(BetfairLayCost, "0.00") + vbCrLf
End If
result = result + "Retention for Back Bet = " + Format(RetentionBack, "0.0") + "%" + vbCrLf
result = result + "Retention for Lay Bet = " + Format(RetentionLay, "0.0") + "%"

End Function
Sub calc(ByVal How As String)
Dim bIterate As Boolean
Dim iTerations As Long
iTerations = 0
If How = "NotEqual" Then
    bIterate = False
Else
    bIterate = True
    If LayStake = 0 Then LayStake = BackStake 'seed
End If
Difference = 1000 'seed

BackReturn = BackOdds * BackStake - (BackPc * ((BackOdds - 1) * BackStake) / 100) - StakeNotReturned

Do
If bIterate Then
    If BackProfit < LayProfit Then
        LayStake = LayStake - 0.01
    ElseIf BackProfit > LayProfit Then
        LayStake = LayStake + 0.01
    End If
End If

LayRisk = LayStake * (LayOdds - 1)
'BackReturn = BackOdds * BackStake - (BackPc * ((BackOdds - 1) * BackStake) / 100) - StakeNotReturned
BackProfit = BackReturn - LayRisk
LayProfit = LayStake * (100 - LayPc) / 100

Difference = Abs(BackProfit - LayProfit)

If bIterate = False Then
    Exit Do
End If
iTerations = iTerations + 1
If iTerations > lMaxIterations Then Beep: Beep: Beep: Exit Do
Loop While (Difference > 0.02)

BetfairBackCost = BackProfit - BackStake
BetfairLayCost = LayProfit - BackStake

If StakeNotReturned Then
    RetentionBack = (BackProfit / BackStake) * 100
    RetentionLay = (LayProfit / BackStake) * 100
Else
    If BetfairBackCost > 0 Then
        RetentionBack = Abs(((BackProfit / BackStake) * 100) - 100)
    Else
        RetentionBack = ((BackProfit / BackStake) * 100) - 100
    End If
    If BetfairLayCost > 0 Then
        RetentionLay = Abs(100 - ((LayProfit / BackStake) * 100) - 100)
    Else
        RetentionLay = ((LayProfit / BackStake) * 100) - 100
    End If
End If

End Sub

Sub EWcalc(ByVal How As String)
'each way calculations
Dim bIterate As Boolean
Dim iTerations As Long
Dim nDifference As Single
iTerations = 0
If How = "NotEqual" Then
    bIterate = False
Else
    bIterate = True
    If PlaceLayStake = 0 Then PlaceLayStake = BackStake 'seed
End If
nDifference = 1000 'seed

PlaceBackReturn = PlaceBackOdds * BackStake - (BackPc * ((PlaceBackOdds - 1) * BackStake) / 100) - StakeNotReturned

Do
If bIterate Then
    If PlaceBackProfit < PlaceLayProfit Then
        PlaceLayStake = PlaceLayStake - 0.01
    ElseIf PlaceBackProfit > PlaceLayProfit Then
        PlaceLayStake = PlaceLayStake + 0.01
    End If
End If

PlaceLayRisk = PlaceLayStake * (PlaceLayOdds - 1)
PlaceBackProfit = PlaceBackReturn - PlaceLayRisk
PlaceLayProfit = PlaceLayStake * (100 - PlacePc) / 100

nDifference = Abs(PlaceBackProfit - PlaceLayProfit)

If bIterate = False Then
    Exit Do
End If
iTerations = iTerations + 1
If iTerations > lMaxIterations Then Beep: Beep: Beep: Exit Do
Loop While (nDifference > 0.02)

PlaceBetfairBackCost = Format(PlaceBackProfit - BackStake, "0.00")
PlaceBetfairLayCost = Format(PlaceLayProfit - BackStake, "0.00")
ExtraPlace = Format((-BackStake) + LayProfit + (PlaceBackReturn - BackStake) + PlaceLayProfit, "0.00")

End Sub

