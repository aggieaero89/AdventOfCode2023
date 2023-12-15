Attribute VB_Name = "Day14"
Sub Part1()
'
' Macro1 Macro
'
    ActiveSheet.UsedRange
    
    FirstRow = 1
    FirstCol = 1
    Cells(FirstRow, FirstCol).Select
    LastRow = ActiveCell.CurrentRegion.Rows.Count
    LastCol = ActiveCell.CurrentRegion.Columns.Count
    
    
    For j = FirstCol To LastCol
        CurrRow = FirstRow
        strBefore = WorksheetFunction.Concat(Range(Cells(FirstRow, j), Cells(LastRow, j)))
        boolDone = False
        Do
            strCurr = Cells(CurrRow, j).Value
            strNext = Cells(CurrRow + 1, j).Value
            If strCurr = "." And strNext = "O" Then
                Cells(CurrRow, j).Value = "O"
                Cells(CurrRow + 1, j).Value = "."
            End If
            CurrRow = CurrRow + 1
            If CurrRow = LastRow Then
                CurrRow = FirstRow
                strAfter = WorksheetFunction.Concat(Range(Cells(FirstRow, j), Cells(LastRow, j)))
                If strBefore = strAfter Then
                    boolDone = True
                Else
                    strBefore = strAfter
                End If
            End If
        Loop Until boolDone
                
    Next j
        
    MsgBox "Done"
End Sub

Sub Part2()
'
' Macro1 Macro
'
    ActiveSheet.UsedRange
    
    k = 0
    boolQuit = False
    Do
        Call SpinNorth
        Call SpinWest
        Call SpinSouth
        Call SpinEast
        k = k + 1
        Cells(k, 102) = WorksheetFunction.Concat(Range(Cells(1, 1), Cells(100, 100)))
        Application.StatusBar = "Cycle: " & k
        
        'AnswerYes = MsgBox("Cycle: " & k & ". Keep going?", vbYesNo)
        'If AnswerYes = vbYes Then boolQuit = False Else boolQuit
        If k = 142 Then boolQuit = True
    Loop Until boolQuit
    
    MsgBox "Done"
    Application.StatusBar = False
End Sub

Sub Part2_sample()
'
' Macro1 Macro
'
    ActiveSheet.UsedRange
    
    k = 0
    Do
        Call SpinNorth
        Call SpinWest
        Call SpinSouth
        Call SpinEast
        k = k + 1
        Cells(k, 12) = WorksheetFunction.Concat(Range(Cells(1, 1), Cells(10, 10)))
        
        AnswerYes = MsgBox("Cycle: " & k & ". Keep going?", vbYesNo)
        If AnswerYes = vbYes Then boolQuit = False Else boolQuit = True
    
    Loop Until boolQuit
    
    MsgBox "Done"
End Sub

Private Sub SpinNorth()
    FirstRow = 1
    FirstCol = 1
    Cells(FirstRow, FirstCol).Select
    LastRow = ActiveCell.CurrentRegion.Rows.Count
    LastCol = ActiveCell.CurrentRegion.Columns.Count
    
    
    For j = FirstCol To LastCol
        CurrRow = FirstRow
        strBefore = WorksheetFunction.Concat(Range(Cells(FirstRow, j), Cells(LastRow, j)))
        boolDone = False
        Do
            strCurr = Cells(CurrRow, j).Value
            strNext = Cells(CurrRow + 1, j).Value
            If strCurr = "." And strNext = "O" Then
                Cells(CurrRow, j).Value = "O"
                Cells(CurrRow + 1, j).Value = "."
            End If
            CurrRow = CurrRow + 1
            If CurrRow = LastRow Then
                CurrRow = FirstRow
                strAfter = WorksheetFunction.Concat(Range(Cells(FirstRow, j), Cells(LastRow, j)))
                If strBefore = strAfter Then
                    boolDone = True
                Else
                    strBefore = strAfter
                End If
            End If
        Loop Until boolDone
                
    Next j
End Sub

Private Sub SpinSouth()
    LastRow = 1
    LastCol = 1
    Cells(LastRow, LastCol).Select
    FirstRow = ActiveCell.CurrentRegion.Rows.Count
    FirstCol = ActiveCell.CurrentRegion.Columns.Count
    
    
    For j = FirstCol To LastCol Step -1
        CurrRow = FirstRow
        strBefore = WorksheetFunction.Concat(Range(Cells(LastRow, j), Cells(FirstRow, j)))
        boolDone = False
        Do
            strCurr = Cells(CurrRow, j).Value
            strNext = Cells(CurrRow - 1, j).Value
            If strCurr = "." And strNext = "O" Then
                Cells(CurrRow, j).Value = "O"
                Cells(CurrRow - 1, j).Value = "."
            End If
            CurrRow = CurrRow - 1
            If CurrRow = LastRow Then
                CurrRow = FirstRow
                strAfter = WorksheetFunction.Concat(Range(Cells(LastRow, j), Cells(FirstRow, j)))
                If strBefore = strAfter Then
                    boolDone = True
                Else
                    strBefore = strAfter
                End If
            End If
        Loop Until boolDone
                
    Next j
End Sub

Private Sub SpinEast()
    LastRow = 1
    LastCol = 1
    Cells(LastRow, LastCol).Select
    FirstRow = ActiveCell.CurrentRegion.Rows.Count
    FirstCol = ActiveCell.CurrentRegion.Columns.Count
    
    
    For i = FirstRow To LastRow Step -1
        CurrCol = FirstCol
        strBefore = WorksheetFunction.Concat(Range(Cells(i, LastCol), Cells(i, FirstCol)))
        boolDone = False
        Do
            strCurr = Cells(i, CurrCol).Value
            strNext = Cells(i, CurrCol - 1).Value
            If strCurr = "." And strNext = "O" Then
                Cells(i, CurrCol).Value = "O"
                Cells(i, CurrCol - 1).Value = "."
            End If
            CurrCol = CurrCol - 1
            If CurrCol = LastCol Then
                CurrCol = FirstCol
                strAfter = WorksheetFunction.Concat(Range(Cells(i, LastCol), Cells(i, FirstCol)))
                If strBefore = strAfter Then
                    boolDone = True
                Else
                    strBefore = strAfter
                End If
            End If
        Loop Until boolDone
                
    Next i
End Sub

Private Sub SpinWest()
    FirstRow = 1
    FirstCol = 1
    Cells(FirstRow, FirstCol).Select
    LastRow = ActiveCell.CurrentRegion.Rows.Count
    LastCol = ActiveCell.CurrentRegion.Columns.Count
    
    
    For i = FirstRow To LastRow
        CurrCol = FirstCol
        strBefore = WorksheetFunction.Concat(Range(Cells(i, FirstCol), Cells(i, LastCol)))
        boolDone = False
        Do
            strCurr = Cells(i, CurrCol).Value
            strNext = Cells(i, CurrCol + 1).Value
            If strCurr = "." And strNext = "O" Then
                Cells(i, CurrCol).Value = "O"
                Cells(i, CurrCol + 1).Value = "."
            End If
            CurrCol = CurrCol + 1
            If CurrCol = LastCol Then
                CurrCol = FirstCol
                strAfter = WorksheetFunction.Concat(Range(Cells(i, FirstCol), Cells(i, LastCol)))
                If strBefore = strAfter Then
                    boolDone = True
                Else
                    strBefore = strAfter
                End If
            End If
        Loop Until boolDone
                
    Next i
End Sub

Sub FindFirstBoardThatRepeats()
    Range("DB1").Select
    For i = 1 To 1087
        ActiveCell.Value = i
        If ActiveCell.Offset(1, -2).Value > 1 Then Exit For
    Next i
End Sub
