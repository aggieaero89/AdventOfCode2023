Attribute VB_Name = "Day6"
Sub Part1()
'
' Macro1 Macro
'
    Range("A1").Select
    
    'Time
    arrInput = Split(ActiveCell.Value, ":")
    strTime = Trim(arrInput(1))
    
    Do
        strA = strTime
        strTime = Replace(strTime, "  ", " ")
    Loop Until strA = strTime
    
    arrTime = Split(strTime, " ") 'milliseconds
    
    
    'Distance
    arrInput = Split(ActiveCell.Offset(1, 0).Value, ":")
    strDist = Trim(arrInput(1))
    
    Do
        strA = strDist
        strDist = Replace(strDist, "  ", " ")
    Loop Until strA = strDist
    
    arrDistance = Split(strDist, " ") 'millimeters
    
    NumWaysWin = 1
    For j = LBound(arrTime) To UBound(arrTime)
        WaysToWin = 0
        MaxTime = CInt(arrTime(j))
        BestDistance = CInt(arrDistance(j))
        For i = 0 To MaxTime
            MyDist = i * (MaxTime - i)
            If MyDist > BestDistance Then
                WaysToWin = WaysToWin + 1
            End If
        Next i
        NumWaysWin = NumWaysWin * WaysToWin
    Next j
    
    MsgBox "Number of ways to win is " & NumWaysWin
End Sub

Sub Part2()
'
' Macro1 Macro
'
    Range("A1").Select
    
    'Time
    arrInput = Split(ActiveCell.Value, ":")
    strTime = Trim(arrInput(1))
    
    Do
        strA = strTime
        strTime = Replace(strTime, " ", "")
    Loop Until strA = strTime
    
    MaxTime = CDbl(strTime) 'milliseconds
    
    
    'Distance
    arrInput = Split(ActiveCell.Offset(1, 0).Value, ":")
    strDist = Trim(arrInput(1))
    
    Do
        strA = strDist
        strDist = Replace(strDist, " ", "")
    Loop Until strA = strDist
    
    BestDistance = CDbl(strDist) 'millimeters
    
    NumWaysWin = 0
    For i = 0 To MaxTime
        MyDist = i * (MaxTime - i)
        If MyDist > BestDistance Then
            NumWaysWin = NumWaysWin + 1
        End If
    Next i

    MsgBox "Number of ways to win is " & NumWaysWin
End Sub

