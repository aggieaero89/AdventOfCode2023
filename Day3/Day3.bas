Attribute VB_Name = "Day3"
Public maxRow As Integer
Public maxCol As Integer
Public boolStar As Boolean

Public col_Star As Integer
Public row_Star As Integer

Sub Part1()
'
    maxRow = Range("A1").CurrentRegion.Rows.Count
    maxCol = Range("A1").CurrentRegion.Columns.Count

    sum_total = 0
    Range("A1").Select
    
    A = ""
    boolNumStart = False
    boolStar = False
    
    For i = 1 To maxRow
        invalid = False
        
        For j = 1 To maxCol
            strText = Cells(i, j).Value
            If IsNumeric(strText) Then
                If Not boolNumStart Then
                    boolNumStart = True
                    valid = CheckAdjacent(i, j, True, False)
                    invalid = Not valid
                Else
                    If invalid Then
                        valid = CheckAdjacent(i, j)
                        invalid = Not valid
                    End If
                End If
                A = A & strText
                
                If j = maxCol Then
                    If Not invalid Then
                        sum_total = sum_total + CInt(A)
                    End If
                    
                    boolNumStart = False
                    A = ""
                End If
                
                
            Else
                If boolNumStart Then 'end of number reached previously
                    If invalid Then
                        valid = CheckAdjacent(i, j - 1, False, True)
                        invalid = Not valid
                    End If
                    
                    If Not invalid Then
                        sum_total = sum_total + CInt(A)
                    End If
                    
                    boolNumStart = False
                    A = ""

                End If
                
            End If
                
        Next j
    Next i
    
    MsgBox "Sum of all of the part numbers in the engine schematic is " & sum_total
    
End Sub

Function CheckAdjacent(ByVal intRow As Integer, ByVal intCol As Integer, _
                       Optional boolStart As Boolean = False, Optional boolEnd As Boolean = False) As Boolean

    boolPrevTop = False
    boolPrevCol = False
    boolPrevBot = False
    
    boolNextRow = False
    boolPrevRow = False
    
    boolNextTop = False
    boolNextCol = False
    boolNextBot = False

    
    If boolStart Then
        If intRow = 1 Then
            If intCol > 1 Then
                boolPrevCol = CheckForSymbol(intRow + 0, intCol - 1)
                boolPrevBot = CheckForSymbol(intRow + 1, intCol - 1)
            End If
            boolNextRow = CheckForSymbol(intRow + 1, intCol)

        ElseIf intRow = maxRow Then
            If intCol > 1 Then
                boolPrevTop = CheckForSymbol(intRow - 1, intCol - 1)
                boolPrevCol = CheckForSymbol(intRow + 0, intCol - 1)
            End If
            boolPrevRow = CheckForSymbol(intRow - 1, intCol)
            
        Else
            If intCol > 1 Then
                boolPrevTop = CheckForSymbol(intRow - 1, intCol - 1)
                boolPrevCol = CheckForSymbol(intRow + 0, intCol - 1)
                boolPrevBot = CheckForSymbol(intRow + 1, intCol - 1)
            End If
            boolNextRow = CheckForSymbol(intRow + 1, intCol)
            boolPrevRow = CheckForSymbol(intRow - 1, intCol)
            
        End If

    ElseIf boolEnd Then
        If intRow = 1 Then
            If intCol < maxCol Then
                boolNextCol = CheckForSymbol(intRow + 0, intCol + 1)
                boolNextBot = CheckForSymbol(intRow + 1, intCol + 1)
            End If
        
        ElseIf intRow = maxRow Then
            If intCol < maxCol Then
                boolNextTop = CheckForSymbol(intRow - 1, intCol + 1)
                boolNextCol = CheckForSymbol(intRow + 0, intCol + 1)
            End If
        
        Else
            If intCol < maxCol Then
                boolNextTop = CheckForSymbol(intRow - 1, intCol + 1)
                boolNextCol = CheckForSymbol(intRow + 0, intCol + 1)
                boolNextBot = CheckForSymbol(intRow + 1, intCol + 1)
            End If
            
        End If
        
    Else
        If intRow = 1 Then
            boolNextRow = CheckForSymbol(intRow + 1, intCol)
        ElseIf intRow = maxRow Then
            boolPrevRow = CheckForSymbol(intRow - 1, intCol)
        Else
            boolPrevRow = CheckForSymbol(intRow - 1, intCol)
            boolNextRow = CheckForSymbol(intRow + 1, intCol)
        End If
        
    End If

    CheckAdjacent = boolPrevTop Or _
                    boolPrevCol Or _
                    boolPrevBot Or _
                    boolNextRow Or _
                    boolPrevRow Or _
                    boolNextTop Or _
                    boolNextCol Or _
                    boolNextBot

End Function

Function CheckForSymbol(intRow As Integer, intCol As Integer) As Boolean
'returns False if it's a number
'when boolStar = False, returns True if it's any symbol (e.g., *, $, #, +) except "."
'              = True   returns True only if it's the star (*) symbol

    strA = Cells(intRow, intCol).Value
    
    CheckForSymbol = False
    
    If Not IsNumeric(strA) Then
        If strA <> "." Then
            If boolStar Then
                If strA = "*" Then
                    CheckForSymbol = True
                    col_Star = intCol
                    row_Star = intRow
                End If
            Else
                CheckForSymbol = True
            End If
        End If
    End If

End Function

Sub Part2()
'---------------------------------------------------------------------------------
'Find all values next to stars (*) and write the value and star location to the 'stars' worksheet

    Dim shtStars As Worksheet
    Set shtStars = Sheets("stars")
    shtStars.UsedRange.Clear
    shtStars.Range("A1") = "Value"
    shtStars.Range("B1") = "Row_Star"
    shtStars.Range("C1") = "Col_Star"
    
    maxRow = Range("A1").CurrentRegion.Rows.Count
    maxCol = Range("A1").CurrentRegion.Columns.Count
    
    sum_total = 0
    Range("A1").Select
    row_sht = 2
    
    A = ""
    boolNumStart = False
    boolStar = True
    
    
    For i = 1 To maxRow
        invalid = False
        
        For j = 1 To maxCol
            strText = Cells(i, j).Value
            If IsNumeric(strText) Then
                If Not boolNumStart Then
                    boolNumStart = True
                    valid = CheckAdjacent(i, j, True, False)
                    invalid = Not valid
                Else
                    If invalid Then
                        valid = CheckAdjacent(i, j)
                        invalid = Not valid
                    End If
                End If
                A = A & strText
                
                If j = maxCol Then
                    If Not invalid Then
                        shtStars.Cells(row_sht, 1).Value = CInt(A)
                        shtStars.Cells(row_sht, 2).Value = row_Star
                        shtStars.Cells(row_sht, 3).Value = col_Star
                        row_sht = row_sht + 1
                    End If
                    
                    boolNumStart = False
                    A = ""
                End If
                
                
            Else
                If boolNumStart Then 'end of number reached previously
                    If invalid Then
                        valid = CheckAdjacent(i, j - 1, False, True)
                        invalid = Not valid
                    End If
                    
                    If Not invalid Then
                        shtStars.Cells(row_sht, 1).Value = CInt(A)
                        shtStars.Cells(row_sht, 2).Value = row_Star
                        shtStars.Cells(row_sht, 3).Value = col_Star
                        row_sht = row_sht + 1
                    End If
                    
                    boolNumStart = False
                    A = ""

                End If
                
            End If
                
        Next j
    Next i
    
'---------------------------------------------------------------------------------
'Sort "stars" worksheet based on the star row and then the star column. Compute gear ratio based on common star location.

    shtStars.Activate
    Range("A1").CurrentRegion.Sort Key1:=Range("B1"), Key2:=Range("C1"), Header:=xlYes
    maxRow = Range("A1").CurrentRegion.Rows.Count
    
    sum_total = 0
    
    boolSkipIt = False
    For i = 2 To maxRow
        If Not boolSkipIt Then
            curr_rowStar = Cells(i, 2).Value
            curr_colStar = Cells(i, 3).Value
            
            next_rowStar = Cells(i + 1, 2).Value
            next_colStar = Cells(i + 1, 3).Value
            
            If (curr_rowStar = next_rowStar) And (curr_colStar = next_colStar) Then
                curr_value = Cells(i, 1).Value
                next_value = Cells(i + 1, 1).Value
                sum_total = sum_total + curr_value * next_value
                boolSkipIt = True
            End If
        Else
            boolSkipIt = False
        End If
        
    Next i
    
    MsgBox ("Sum of all the gear ratios is " & sum_total)
    
End Sub

