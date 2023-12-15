Attribute VB_Name = "Day15"
Sub Part1()
'
' Macro1 Macro
'
    ActiveSheet.UsedRange
    
    Range("A1").Select
    arrHASH = Split(ActiveCell.Value, ",")
    
    SumTotal = 0
    For Each Item In arrHASH
        CurrentValue = 0
        For i = 1 To Len(Item)
            charItem = Mid(Item, i, 1)
            CurrentValue = (17 * (CurrentValue + Asc(charItem))) Mod 256
        Next i
        SumTotal = SumTotal + CurrentValue
    Next Item
    
    MsgBox "The sum of the results is " & SumTotal
End Sub

Sub Part2()
'
' Macro1 Macro
'
    ActiveSheet.UsedRange
    
    Range("A1").Select
    arrHASH = Split(ActiveCell.Value, ",")
    
    SumTotal = 0
    For Each Item In arrHASH
        CurrentValue = 0
        CurrentLens = ""
        For i = 1 To Len(Item)
            charItem = Mid(Item, i, 1)
            Select Case charItem
                Case "="
                    FocalLength = Right(Item, 1)
                    CurrCol = 3
                    NumLensesInBox = Cells(CurrentValue + 1, CurrCol - 1)
                    
                    If NumLensesInBox = 0 Then
                        Cells(CurrentValue + 1, CurrCol).Value = CurrentLens & " " & FocalLength
                        Cells(CurrentValue + 1, CurrCol - 1) = 1
                    Else
                        Do
                            arrLens = Split(Cells(CurrentValue + 1, CurrCol), " ")
                            If arrLens(0) = CurrentLens Then
                                Cells(CurrentValue + 1, CurrCol) = CurrentLens & " " & FocalLength
                                boolDone = True
                            Else
                                boolDone = False
                                CurrCol = CurrCol + 1
                                If CurrCol - 3 = NumLensesInBox Then
                                    Cells(CurrentValue + 1, CurrCol).Value = CurrentLens & " " & FocalLength
                                    Cells(CurrentValue + 1, 2) = NumLensesInBox + 1
                                    boolDone = True
                                End If
                            End If
                        Loop Until boolDone
                    End If
                    Exit For
                    
                Case "-"
                    CurrCol = 3
                    NumLensesInBox = Cells(CurrentValue + 1, CurrCol - 1)
                    
                    If NumLensesInBox > 0 Then
                        Do
                            arrLens = Split(Cells(CurrentValue + 1, CurrCol), " ")
                            If arrLens(0) = CurrentLens Then
                                Cells(CurrentValue + 1, CurrCol).Delete Shift:=xlToLeft
                                Cells(CurrentValue + 1, 2) = NumLensesInBox - 1
                                boolDone = True
                            Else
                                CurrCol = CurrCol + 1
                                If CurrCol - 3 = NumLensesInBox Then boolDone = True Else boolDone = False
                            End If
                        Loop Until boolDone
                    End If
                    
                    
                Case Else:
                    CurrentValue = (17 * (CurrentValue + Asc(charItem))) Mod 256
                    CurrentLens = CurrentLens & charItem
            End Select
        Next i
    Next Item
    
    SumTotal = 0
    For k = 1 To 256
        NumberOfLensesInBox = Cells(k, 2).Value
        If NumberOfLensesInBox > 0 Then
            For n = 1 To NumberOfLensesInBox
                arrLens = Split(Cells(k, 2 + n), " ")
                SumTotal = SumTotal + k * n * CInt(arrLens(1))
            Next n
        End If
    Next k
    
    MsgBox "Sum total is " & SumTotal
End Sub
