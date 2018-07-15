Attribute VB_Name = "Module1"
Option Explicit
Public LayPc As Single
Public BackPc As Single
Public bStakeNotReturned As Boolean
Public BackStake As Single
Public BackOdds As Single
Public LayOdds As Single
Public LayStake As Single
Public BackReturn As Single
Public LayRisk As Single
Public BackProfit As Single
Public LayProfit As Single
Public Difference As Single
Public BetfairBackCost As Single
Public BetfairLayCost As Single
Public NonBetfairBackCost As Single
Public NonBetfairLayCost As Single
Public StakeNotReturned As Single
Function ConvertOdds(sTyped As String) As String
Dim iSlash As Integer
Dim iLeft As Integer
Dim iRight As Integer
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
Function result() As String
result = "Stake Not Returned = " + IIf(StakeNotReturned, "YES", "NO") + vbCrLf
result = result + "Back Commission = " + CStr(BackPc) + "%" + vbCrLf
result = result + "Lay Commission = " + CStr(LayPc) + "%" + vbCrLf
result = result + "Back Return = " + Format(BackReturn, "0.00") + vbCrLf
result = result + "Lay Risk = " + Format(LayRisk, "0.00") + vbCrLf
result = result + "Back Profit = " + Format(BackProfit, "0.00") + vbCrLf
result = result + "Lay Profit = " + Format(LayProfit, "0.00") + vbCrLf
result = result + "Difference = " + Format(Difference, "0.00") + vbCrLf
If BetfairBackCost > 0 Then
    result = result + "Exchange Back Profit = " + Format(BetfairBackCost, "0.00") + vbCrLf
Else
    result = result + "Exchange Back Cost = " + Format(BetfairBackCost, "0.00") + vbCrLf
End If
If BetfairLayCost > 0 Then
    result = result + "Exchange Lay Profit = " + Format(BetfairLayCost, "0.00") + vbCrLf
Else
    result = result + "Exchange Lay Cost = " + Format(BetfairLayCost, "0.00") + vbCrLf
End If
If NonBetfairBackCost > 0 Then
    result = result + "Dutch Profit 1 = " + Format(NonBetfairBackCost, "0.00") + vbCrLf
Else
    result = result + "Dutch Cost 1 = " + Format(NonBetfairBackCost, "0.00") + vbCrLf
End If
If NonBetfairLayCost > 0 Then
    result = result + "Dutch Profit 2 = " + Format(NonBetfairLayCost, "0.00")
Else
    result = result + "Dutch Cost 2 = " + Format(NonBetfairLayCost, "0.00")
End If
End Function
Sub calc(ByVal How As String)
Dim bIterate As Boolean
Dim iTerations As Integer
iTerations = 0
If How = "NotEqual" Then
    bIterate = False
Else
    bIterate = True
    If LayStake = 0 Then LayStake = BackStake 'seed
End If
Difference = 1000 'seed

Do
If bIterate Then

        If BackProfit < LayProfit Then
            LayStake = LayStake - 0.01
        ElseIf BackProfit > LayProfit Then
            LayStake = LayStake + 0.01
        End If
End If

LayRisk = LayStake * (LayOdds - 1)
BackReturn = BackOdds * BackStake - (BackPc * ((BackOdds - 1) * BackStake) / 100) - StakeNotReturned
BackProfit = BackReturn - LayRisk
LayProfit = LayStake * (100 - LayPc) / 100

Difference = Abs(BackProfit - LayProfit)

If bIterate = False Then
    Exit Do
End If
iTerations = iTerations + 1
If iTerations > 10000 Then Beep: Beep: Beep: Exit Do
Loop While (Difference > 0.02)
Dim OutGoing As Single
OutGoing = BackStake + Format(LayStake, "0.00")
NonBetfairBackCost = (BackStake * BackOdds) - OutGoing
BetfairBackCost = BackProfit - BackStake
NonBetfairLayCost = (LayStake * LayOdds) - OutGoing
BetfairLayCost = LayProfit - BackStake
End Sub



