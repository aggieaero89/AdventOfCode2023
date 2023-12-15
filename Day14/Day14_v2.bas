Attribute VB_Name = "Day14_v2"
Sub Part1()
'
' Macro1 Macro
'
    ActiveSheet.UsedRange
    
    Call SpinNorth
        
    MsgBox "Done"
End Sub

Sub Part2()
'
' Macro1 Macro
'
    ActiveSheet.UsedRange
    
    boolManual = True 'set True for user interaction with each cycle, or False for automatic up to a certain count
    
    Range("A1").Select
    MaxRows = ActiveCell.CurrentRegion.Rows.Count
    MaxCols = ActiveCell.CurrentRegion.Columns.Count
    OutputCol = MaxCols + 2
    
    k = 0
    boolQuit = False
    Do
        Application.StatusBar = "Spinning North": Call SpinNorth
        Application.StatusBar = "Spinning West":  Call SpinWest
        Application.StatusBar = "Spinning South": Call SpinSouth
        Application.StatusBar = "Spinning East":  Call SpinEast
        k = k + 1
        Cells(k, OutputCol) = WorksheetFunction.Concat(Range(Cells(1, 1), Cells(MaxRows, MaxCols)))
        Application.StatusBar = "Cycle: " & k
        
        If boolManual Then
            AnswerYes = MsgBox("Cycle: " & k & ". Keep going?", vbYesNo)
            If AnswerYes = vbYes Then boolQuit = False Else boolQuit = True
        Else
            If k = 142 Then boolQuit = True '142 was manually chosen based off a run for 1087 cycles of Part 2
        End If
    Loop Until boolQuit
    
    MsgBox "Done"
    Application.StatusBar = False
End Sub

Private Sub SpinNorth()
    FirstRow = 1
    FirstCol = 1
    Cells(FirstRow, FirstCol).Select
    LastRow = ActiveCell.CurrentRegion.Rows.Count
    LastCol = ActiveCell.CurrentRegion.Columns.Count
    
    
    For j = FirstCol To LastCol
        CurrRow = FirstRow + 1
        
        Do
            strCurr = Cells(CurrRow, j).Value
            strNorth = Cells(CurrRow - 1, j).Value
            If strCurr = "O" And strNorth = "." Then
                TheRow = CurrRow
                boolDone = False
                
                Do
                    Cells(TheRow, j).Value = "."
                    Cells(TheRow - 1, j).Value = "O"
                    TheRow = TheRow - 1
                    If TheRow = FirstRow Then
                        boolDone = True
                    ElseIf Cells(TheRow - 1, j).Value <> "." Then
                        boolDone = True
                    End If
                Loop Until boolDone
                
            End If
            
            CurrRow = CurrRow + 1
        Loop Until CurrRow > LastRow
        
    Next j
    
End Sub

Private Sub SpinSouth()
    LastRow = 1
    LastCol = 1
    Cells(LastRow, LastCol).Select
    FirstRow = ActiveCell.CurrentRegion.Rows.Count
    FirstCol = ActiveCell.CurrentRegion.Columns.Count
    
    
    For j = FirstCol To LastCol Step -1
        CurrRow = FirstRow - 1
        
        Do
            strCurr = Cells(CurrRow, j).Value
            strSouth = Cells(CurrRow + 1, j).Value
            If strCurr = "O" And strSouth = "." Then
                TheRow = CurrRow
                boolDone = False
                
                Do
                    Cells(TheRow, j).Value = "."
                    Cells(TheRow + 1, j).Value = "O"
                    TheRow = TheRow + 1
                    If TheRow = FirstRow Then
                        boolDone = True
                    ElseIf Cells(TheRow + 1, j).Value <> "." Then
                        boolDone = True
                    End If
                Loop Until boolDone
                
            End If
            
            CurrRow = CurrRow - 1
        Loop Until CurrRow < LastRow
        
    Next j
    
End Sub

Private Sub SpinEast()
    LastRow = 1
    LastCol = 1
    Cells(LastRow, LastCol).Select
    FirstRow = ActiveCell.CurrentRegion.Rows.Count
    FirstCol = ActiveCell.CurrentRegion.Columns.Count
    
    
    For i = FirstRow To LastRow Step -1
        CurrCol = FirstCol - 1
        
        Do
            strCurr = Cells(i, CurrCol).Value
            strEast = Cells(i, CurrCol + 1).Value
            If strCurr = "O" And strEast = "." Then
                TheCol = CurrCol
                boolDone = False
                
                Do
                    Cells(i, TheCol).Value = "."
                    Cells(i, TheCol + 1).Value = "O"
                    TheCol = TheCol + 1
                    If TheCol = FirstCol Then
                        boolDone = True
                    ElseIf Cells(i, TheCol + 1).Value <> "." Then
                        boolDone = True
                    End If
                Loop Until boolDone
                
            End If
            
            CurrCol = CurrCol - 1
        Loop Until CurrCol < LastCol
        
    Next i
    
End Sub

Private Sub SpinWest()
    FirstRow = 1
    FirstCol = 1
    Cells(FirstRow, FirstCol).Select
    LastRow = ActiveCell.CurrentRegion.Rows.Count
    LastCol = ActiveCell.CurrentRegion.Columns.Count
    
    
    For i = FirstRow To LastRow
        CurrCol = FirstCol + 1
        
        Do
            strCurr = Cells(i, CurrCol).Value
            strWest = Cells(i, CurrCol - 1).Value
            If strCurr = "O" And strWest = "." Then
                TheCol = CurrCol
                boolDone = False
                
                Do
                    Cells(i, TheCol).Value = "."
                    Cells(i, TheCol - 1).Value = "O"
                    TheCol = TheCol - 1
                    If TheCol = FirstCol Then
                        boolDone = True
                    ElseIf Cells(i, TheCol - 1).Value <> "." Then
                        boolDone = True
                    End If
                Loop Until boolDone
                
            End If
            
            CurrCol = CurrCol + 1
        Loop Until CurrCol > LastCol
        
    Next i
    
End Sub

Sub FindFirstBoardThatRepeats()
    Range("DB1").Select
    For i = 1 To 1087
        ActiveCell.Value = i
        If ActiveCell.Offset(1, -2).Value > 1 Then Exit For
    Next i
End Sub
