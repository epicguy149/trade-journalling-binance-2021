Attribute VB_Name = "TradesGenerator"
Sub FindPositions()

Dim Coins As Object
Set Coins = CreateObject("Scripting.Dictionary")
Dim PnL As Object
Set PnL = CreateObject("Scripting.Dictionary")
Dim testValue As Double
Dim testValue2 As Variant
Dim WinLoss As Double
Dim numReplace As Double
Dim PositionClosed As Boolean
Dim WinCount As Integer
Dim LossCount As Integer
Dim StartDate As Date
Dim EndDate As Date
Dim TradeCount As Integer
Dim BiggestWin As Double
Dim BiggestLoss As Double
Dim totalWin As Double
Dim totalLoss As Double

BiggestWin = 0
BiggestLoss = 0

TradeCount = 0


Dim Rng As Range
Set Rng = ActiveSheet.Range("B2:B" & ActiveSheet.Range("B2").End(xlDown).Row)
    
ActiveSheet.Range("A2:J" & ActiveSheet.Range("A2").End(xlDown).Row).Sort key1:=Range("A2:A" & ActiveSheet.Range("A2").End(xlDown).Row), _
   order1:=xlAscending, Header:=xlNo
   
StartDate = Range("A2").Value
EndDate = Range("A" & ActiveSheet.Range("A2").End(xlDown).Row).Value

ActiveSheet.Range("F2:F" & ActiveSheet.Range("F2").End(xlDown).Row).TextToColumns Destination:=Range("F2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True

For Each cell In Rng

    If Not Coins.Exists(cell.Value) Or Coins.Count = 0 Then
        Coins.Add cell.Value, 0
    End If
    
Next cell

For Each cell In Rng
    
    numReplace = Range("D" & cell.Row).Value
    
    With ActiveSheet.Range("D" & cell.Row)
        .NumberFormat = Number
        .Value = numReplace
    End With
    
    numReplace = Range("E" & cell.Row).Value
    
    With ActiveSheet.Range("E" & cell.Row)
        .NumberFormat = Number
        .Value = numReplace
    End With

    numReplace = Range("F" & cell.Row).Value
    
    With ActiveSheet.Range("F" & cell.Row)
        .NumberFormat = Number
        .Value = numReplace
    End With
    
    numReplace = Range("G" & cell.Row).Value
    
    With ActiveSheet.Range("G" & cell.Row)
        .NumberFormat = Number
        .Value = numReplace
    End With
    
    numReplace = Range("I" & cell.Row).Value
    
    With ActiveSheet.Range("I" & cell.Row)
        .NumberFormat = Number
        .Value = numReplace
    End With
Next cell
    
    

For Each cell In Rng

PositionClosed = False
'Colour row

    If cell.Offset(0, 1).Value = "SELL" Then
        With Range("A" & cell.Row & ":J" & cell.Row).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 192
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        
        With Range("A" & cell.Row & ":J" & cell.Row).Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With
        
        'Change sell amount to negative and add to dictionary
        
        cell.Offset(0, 3).Value = "-" & cell.Offset(0, 3).Value
        testValue2 = Round(Format(cell.Offset(0, 3).Value, "#.000"), 8)
        PnL(cell.Value) = PnL(cell.Value) + cell.Offset(0, 7).Value
        Coins(cell.Value) = Round(Coins(cell.Value) + testValue2, 8)
        testValue = PnL(cell.Value)
        
    Else
'Change row colour
        With Range("A" & cell.Row & ":J" & cell.Row).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent6
            .TintAndShade = 0.399975585192419
            .PatternTintAndShade = 0
        End With
        
'Add to dictionary
        testValue2 = Round(Format(cell.Offset(0, 3).Value, "#.000"), 8)
        PnL(cell.Value) = PnL(cell.Value) + cell.Offset(0, 7).Value
        Coins(cell.Value) = Round(Coins(cell.Value) + testValue2, 8)
        testValue = PnL(cell.Value)
    End If
    
    If Coins(cell.Value) = 0 Then
        With Range("A" & cell.Row & ":J" & cell.Row).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 10498160
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        
        With Range("A" & cell.Row & ":J" & cell.Row).Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With
        
        PositionClosed = True
        
        TradeCount = TradeCount + 1
        
        WinLoss = PnL(cell.Value)
        PnL(cell.Value) = 0
        
        If WinLoss < BiggestLoss Then
        BiggestLoss = WinLoss
        End If
        
        If WinLoss > BiggestWin Then
        BiggestWin = WinLoss
        End If
        
        If WinLoss > 0.5 Then
            WinLoss = Format(WinLoss, "0.00")
            With Range("K" & cell.Row)
                .Value = "Position closed at a win of " & WinLoss
                .Font.Bold = True
                .Font.Size = 14
                .Font.Color = vbGreen
            End With
            
            WinCount = WinCount + 1
            totalWin = totalWin + WinLoss
        Else
            If WinLoss < -0.5 Then
           With Range("K" & cell.Row)
                .Value = "Position closed at a loss of " & WinLoss
                .Font.Bold = True
                .Font.Size = 14
                .Font.Color = vbRed
            End With
            LossCount = LossCount + 1
            totalLoss = totalLoss + WinLoss
            Else
            With Range("K" & cell.Row)
                .Value = "Position closed at breakeven"
                .Font.Bold = True
                .Font.Size = 14
                .Font.Color = vbYellow
            End With
                WinCount = WinCount + 1
            totalWin = totalWin + WinLossWinCount = WinCount + 1
            totalWin = totalWin + WinLoss
            End If
        End If
        
    End If
    

    
Next cell

Dim WinRate As Double

WinRate = WinCount / Application.WorksheetFunction.Sum(WinCount, LossCount)
WinRate = Application.WorksheetFunction.Product(WinRate, 100)

Range("D" & ActiveSheet.Range("A2").End(xlDown).Row).Offset(1, 0).Value = "PnL from " & StartDate & " to " & EndDate & " is " & Application.WorksheetFunction.Sum(Range("I2:I" & ActiveSheet.Range("A2").End(xlDown).Row)) & " with a win rate of " & WinRate & "% from " & Application.WorksheetFunction.Sum(LossCount, WinCount) & " trades"
Range("D" & ActiveSheet.Range("A2").End(xlDown).Row).Offset(2, 0).Value = "Biggest win was " & BiggestWin & " and biggest loss was " & BiggestLoss
Range("D" & ActiveSheet.Range("A2").End(xlDown).Row).Offset(3, 0).Value = "Average win was " & (totalWin / WinCount) & " and average loss was " & (totalLoss / LossCount)
Range("D" & ActiveSheet.Range("A2").End(xlDown).Row).Offset(2, 0).Font.Bold = True
Range("D" & ActiveSheet.Range("A2").End(xlDown).Row).Offset(2, 0).Font.Size = 14
Range("D" & ActiveSheet.Range("A2").End(xlDown).Row).Offset(1, 0).Font.Bold = True
Range("D" & ActiveSheet.Range("A2").End(xlDown).Row).Offset(1, 0).Font.Size = 14
Range("D" & ActiveSheet.Range("A2").End(xlDown).Row).Offset(1, 0).Font.Bold = True
Range("D" & ActiveSheet.Range("A2").End(xlDown).Row).Offset(1, 0).Font.Size = 14
Range("D" & ActiveSheet.Range("A2").End(xlDown).Row).Offset(3, 0).Font.Bold = True
Range("D" & ActiveSheet.Range("A2").End(xlDown).Row).Offset(3, 0).Font.Size = 14



End Sub

























