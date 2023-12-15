Attribute VB_Name = "Day13"
Sub Part1()
'
' Macro1 Macro
'
    ActiveSheet.UsedRange
    lngSum = 0
    
    Range("A1").Select
    
    Do While Not IsEmpty(ActiveCell.Value)
    
        FirstCol = ActiveCell.Column
        ActiveCell.End(xlToRight).Select
        LastCol = ActiveCell.Column
        
        FirstRow = ActiveCell.Row
        ActiveCell.End(xlDown).Select
        LastRow = ActiveCell.Row
        
        'CHECK ROW MIRROR
        boolFoundRow = False
        For i = FirstRow To LastRow - 1
            strCurrRow = WorksheetFunction.Concat(Range(Cells(i, FirstCol), Cells(i, LastCol)))
            strNextRow = WorksheetFunction.Concat(Range(Cells(i + 1, FirstCol), Cells(i + 1, LastCol)))
            If strCurrRow = strNextRow Then
                CurrRow = i
                PrevRow = CurrRow
                NextRow = CurrRow + 1
                If (PrevRow - 1 >= FirstRow) And (NextRow + 1 <= LastRow) Then
                    Do
                        PrevRow = PrevRow - 1
                        NextRow = NextRow + 1
                        strPrevRow = WorksheetFunction.Concat(Range(Cells(PrevRow, FirstCol), Cells(PrevRow, LastCol)))
                        strNextRow = WorksheetFunction.Concat(Range(Cells(NextRow, FirstCol), Cells(NextRow, LastCol)))
                        If strPrevRow = strNextRow Then
                            boolFoundRow = True
                        Else
                            boolFoundRow = False
                            Exit Do
                        End If
                    Loop Until (PrevRow = FirstRow) Or (NextRow = LastRow)
                Else
                    boolFoundRow = True
                End If
            End If
            If boolFoundRow Then Exit For
        Next i
        
        If boolFoundRow Then
            lngSum = lngSum + (CurrRow - FirstRow + 1) * 100
        
        Else
            'CHECK COLUMN MIRROR
            boolFoundCol = False
            For i = FirstCol To LastCol - 1
                strCurrCol = WorksheetFunction.Concat(Range(Cells(FirstRow, i), Cells(LastRow, i)))
                strNextCol = WorksheetFunction.Concat(Range(Cells(FirstRow, i + 1), Cells(LastRow, i + 1)))
                If strCurrCol = strNextCol Then
                    CurrCol = i
                    PrevCol = CurrCol
                    NextCol = CurrCol + 1
                    If (PrevCol - 1 >= FirstCol) And (NextCol + 1 <= LastCol) Then
                        Do
                            PrevCol = PrevCol - 1
                            NextCol = NextCol + 1
                            strPrevCol = WorksheetFunction.Concat(Range(Cells(FirstRow, PrevCol), Cells(LastRow, PrevCol)))
                            strNextCol = WorksheetFunction.Concat(Range(Cells(FirstRow, NextCol), Cells(LastRow, NextCol)))
                            If strPrevCol = strNextCol Then
                                boolFoundCol = True
                            Else
                                boolFoundCol = False
                                Exit Do
                            End If
                        Loop Until (PrevCol = FirstCol) Or (NextCol = LastCol)
                    Else
                        boolFoundCol = True
                    End If
                End If
                If boolFoundCol Then Exit For
            Next i
            
            If boolFoundCol Then
                lngSum = lngSum + CurrCol
            End If
        
        End If
        
        Cells(LastRow + 2, 1).Select
    Loop
    
    MsgBox "Summarizing all the notes is " & lngSum
End Sub

Function StrCompareDiff(ByVal Str1 As String, ByVal Str2 As String) As Integer
    knt_Diff = 0
    For k = 1 To Len(Str1)
        chr1 = Mid(Str1, k, 1)
        chr2 = Mid(Str2, k, 1)
        If chr1 <> chr2 Then knt_Diff = knt_Diff + 1
    Next k
    
    StrCompareDiff = knt_Diff
End Function


Sub Part2()
'
' Macro1 Macro
'
    ActiveSheet.UsedRange
    lngSum = 0
    
    Range("A1").Select
    
    Do While Not IsEmpty(ActiveCell.Value)
    
        FirstCol = ActiveCell.Column
        ActiveCell.End(xlToRight).Select
        LastCol = ActiveCell.Column
        
        FirstRow = ActiveCell.Row
        ActiveCell.End(xlDown).Select
        LastRow = ActiveCell.Row
        
        'CHECK ROW MIRROR
        boolFoundRow = False
        For i = FirstRow To LastRow - 1
            strCurrRow = WorksheetFunction.Concat(Range(Cells(i, FirstCol), Cells(i, LastCol)))
            strNextRow = WorksheetFunction.Concat(Range(Cells(i + 1, FirstCol), Cells(i + 1, LastCol)))
            numDiffs = StrCompareDiff(strCurrRow, strNextRow)
            If numDiffs <= 1 Then
                CurrRow = i
                PrevRow = CurrRow
                NextRow = CurrRow + 1
                If (PrevRow - 1 >= FirstRow) And (NextRow + 1 <= LastRow) Then
                    Do
                        PrevRow = PrevRow - 1
                        NextRow = NextRow + 1
                        strPrevRow = WorksheetFunction.Concat(Range(Cells(PrevRow, FirstCol), Cells(PrevRow, LastCol)))
                        strNextRow = WorksheetFunction.Concat(Range(Cells(NextRow, FirstCol), Cells(NextRow, LastCol)))
                        numDiffs = numDiffs + StrCompareDiff(strPrevRow, strNextRow)
                        If numDiffs > 1 Then
                            boolFoundRow = False
                            Exit Do
                        Else
                            If numDiffs = 1 Then boolFoundRow = True Else boolFoundRow = False
                        End If
                    Loop Until (PrevRow = FirstRow) Or (NextRow = LastRow)
                Else
                    If numDiffs = 1 Then boolFoundRow = True Else boolFoundRow = False
                End If
            End If
            If boolFoundRow Then Exit For
        Next i
        
        If boolFoundRow Then
            lngSum = lngSum + (CurrRow - FirstRow + 1) * 100
        
        Else
            'CHECK COLUMN MIRROR
            boolFoundCol = False
            For i = FirstCol To LastCol - 1
                strCurrCol = WorksheetFunction.Concat(Range(Cells(FirstRow, i), Cells(LastRow, i)))
                strNextCol = WorksheetFunction.Concat(Range(Cells(FirstRow, i + 1), Cells(LastRow, i + 1)))
                numDiffs = StrCompareDiff(strCurrCol, strNextCol)
                If numDiffs <= 1 Then
                    CurrCol = i
                    PrevCol = CurrCol
                    NextCol = CurrCol + 1
                    If (PrevCol - 1 >= FirstCol) And (NextCol + 1 <= LastCol) Then
                        Do
                            PrevCol = PrevCol - 1
                            NextCol = NextCol + 1
                            strPrevCol = WorksheetFunction.Concat(Range(Cells(FirstRow, PrevCol), Cells(LastRow, PrevCol)))
                            strNextCol = WorksheetFunction.Concat(Range(Cells(FirstRow, NextCol), Cells(LastRow, NextCol)))
                            numDiffs = numDiffs + StrCompareDiff(strPrevCol, strNextCol)
                            If numDiffs > 1 Then
                                boolFoundCol = False
                                Exit Do
                            Else
                                If numDiffs = 1 Then boolFoundCol = True Else boolFoundCol = False
                            End If
                        Loop Until (PrevCol = FirstCol) Or (NextCol = LastCol)
                    Else
                        If numDiffs = 1 Then boolFoundCol = True Else boolFoundCol = False
                    End If
                End If
                If boolFoundCol Then Exit For
            Next i
            
            If boolFoundCol Then
                lngSum = lngSum + CurrCol
            End If
        
        End If
        
        Cells(LastRow + 2, 1).Select
    Loop
    
    MsgBox "Summarizing all the notes with the new reflection is " & lngSum
End Sub
