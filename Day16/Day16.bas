Attribute VB_Name = "Day16"
Const intUp = 1
Const intDown = 2
Const intRight = 4
Const intLeft = 8
    
Private Type ptGrid
    boolEnergized As Boolean
    boolUp As Boolean
    boolDown As Boolean
    boolRight As Boolean
    boolLeft As Boolean
End Type

Private Type TrajPath
    strDir As String
    intRow As Integer
    intCol As Integer
    boolExists As Boolean
End Type

Function MoveIt(ByRef ThePath As TrajPath, ByVal MaxRow As Integer, ByVal MaxCol As Integer) As Boolean
    boolMove = False
    
    Select Case ThePath.strDir
        Case "Up":
            If ThePath.intRow > 1 Then
                ThePath.intRow = ThePath.intRow - 1
                boolMove = True
            End If
            
        Case "Down":
            If ThePath.intRow < MaxRow Then
                ThePath.intRow = ThePath.intRow + 1
                boolMove = True
            End If
            
        Case "Right":
            If ThePath.intCol < MaxCol Then
                ThePath.intCol = ThePath.intCol + 1
                boolMove = True
            End If
            
        Case "Left":
            If ThePath.intCol > 1 Then
                ThePath.intCol = ThePath.intCol - 1
                boolMove = True
            End If
            
    End Select
    
    MoveIt = boolMove
End Function

Function NextDir(ByRef ThePath As TrajPath) As Integer
    '1=Up, 2=Down, 4=Right, 8=Left
    
    Select Case ThePath.strDir
        Case "Up":
            Select Case Cells(ThePath.intRow, ThePath.intCol).Value
                Case "|", ".": NextDir = intUp
                Case "-":      NextDir = intRight + intLeft
                Case "/":      NextDir = intRight
                Case "\":      NextDir = intLeft
            End Select
            
        Case "Down":
            Select Case Cells(ThePath.intRow, ThePath.intCol).Value
                Case "|", ".": NextDir = intDown
                Case "-":      NextDir = intRight + intLeft
                Case "/":      NextDir = intLeft
                Case "\":      NextDir = intRight
            End Select
            
        Case "Right":
            Select Case Cells(ThePath.intRow, ThePath.intCol).Value
                Case "|":      NextDir = intUp + intDown
                Case "-", ".": NextDir = intRight
                Case "/":      NextDir = intUp
                Case "\":      NextDir = intDown
            End Select
            
        Case "Left":
            Select Case Cells(ThePath.intRow, ThePath.intCol).Value
                Case "|":      NextDir = intUp + intDown
                Case "-", ".": NextDir = intLeft
                Case "/":      NextDir = intDown
                Case "\":      NextDir = intUp
            End Select
            
    End Select
    
End Function

Function EnergizeGrid(ByRef ThePath As TrajPath, ByRef TheGrid As ptGrid) As Boolean
    boolRepeat = False
    If TheGrid.boolEnergized Then
        Select Case ThePath.strDir
            Case "Up":    If TheGrid.boolUp Then boolRepeat = True Else TheGrid.boolUp = True
            Case "Down":  If TheGrid.boolDown Then boolRepeat = True Else TheGrid.boolDown = True
            Case "Right": If TheGrid.boolRight Then boolRepeat = True Else TheGrid.boolRight = True
            Case "Left":  If TheGrid.boolLeft Then boolRepeat = True Else TheGrid.boolLeft = True
        End Select
    Else
        TheGrid.boolEnergized = True
        
        Select Case ThePath.strDir
            Case "Up":    TheGrid.boolUp = True
            Case "Down":  TheGrid.boolDown = True
            Case "Right": TheGrid.boolRight = True
            Case "Left":  TheGrid.boolLeft = True
        End Select
    End If
    
    EnergizeGrid = Not boolRepeat

End Function

Sub Part1()
'
' Macro1 Macro
'
    ActiveSheet.UsedRange
    
    Range("A1").Select
    MaxRow = ActiveCell.CurrentRegion.Rows.Count
    MaxCol = ActiveCell.CurrentRegion.Columns.Count
    
    Dim matGrid() As ptGrid
    ReDim matGrid(1 To MaxRow, 1 To MaxCol)
    
    For i = 1 To MaxRow
        For j = 1 To MaxCol
            matGrid(i, j).boolEnergized = False
            matGrid(i, j).boolUp = False
            matGrid(i, j).boolDown = False
            matGrid(i, j).boolRight = False
            matGrid(i, j).boolLeft = False
        Next j
    Next i
    
    Dim arrPath() As TrajPath
    ReDim arrPath(1 To 1)
    
    arrPath(1).intRow = 1
    arrPath(1).intCol = 1
    arrPath(1).boolExists = True
    Select Case Cells(1, 1).Value
        Case ".", "-": arrPath(1).strDir = "Right"
        Case "\", "|": arrPath(1).strDir = "Down"
        Case "/": arrPath(1).strDir = "Up"
    End Select
    
    matGrid(1, 1).boolEnergized = True
    matGrid(1, 1).boolRight = True
    
    Dim boolMoved As Boolean
    
    NumPaths = 1
    Do
        boolDone = True
        For i = 1 To NumPaths
            If arrPath(i).boolExists Then
                boolMoved = MoveIt(arrPath(i), MaxRow, MaxCol)
                If boolMoved Then
                    boolNotRepeated = EnergizeGrid(arrPath(i), matGrid(arrPath(i).intRow, arrPath(i).intCol))
                    If boolNotRepeated Then
                        boolDone = False
                        NewDir = NextDir(arrPath(i))
                        Select Case NewDir
                            Case intUp: arrPath(i).strDir = "Up"
                            Case intDown: arrPath(i).strDir = "Down"
                            Case intRight: arrPath(i).strDir = "Right"
                            Case intLeft: arrPath(i).strDir = "Left"
                            Case intUp + intDown
                                NumPaths = NumPaths + 1
                                arrPath(i).strDir = "Up"
                                
                                ReDim Preserve arrPath(1 To NumPaths)
                                arrPath(NumPaths).strDir = "Down"
                                arrPath(NumPaths).intRow = arrPath(i).intRow
                                arrPath(NumPaths).intCol = arrPath(i).intCol
                                arrPath(NumPaths).boolExists = True
                            Case intRight + intLeft
                                NumPaths = NumPaths + 1
                                arrPath(i).strDir = "Right"
                                
                                ReDim Preserve arrPath(1 To NumPaths)
                                arrPath(NumPaths).strDir = "Left"
                                arrPath(NumPaths).intRow = arrPath(i).intRow
                                arrPath(NumPaths).intCol = arrPath(i).intCol
                                arrPath(NumPaths).boolExists = True
                        End Select
                    Else
                        arrPath(i).boolExists = False
                    End If
                Else
                    arrPath(i).boolExists = False
                End If
            End If
        Next i
    Loop Until boolDone
    
    k = MaxCol + 3
    For i = 1 To MaxRow
        For j = 1 To MaxCol
            If matGrid(i, j).boolEnergized Then
                Cells(i, j + k).Value = "#"
            End If
        Next j
    Next i
    
    MsgBox "The sum of the results is " & WorksheetFunction.CountA(Cells(1, k + 1).CurrentRegion)
End Sub

Sub Part2()
'
' Macro1 Macro
'
    ActiveSheet.UsedRange
    
    Range("A1").Select
    MaxRow = ActiveCell.CurrentRegion.Rows.Count
    MaxCol = ActiveCell.CurrentRegion.Columns.Count
    
    MaxNumEnergized = -1
    'Top Row
    Application.StatusBar = "Working on Top Row"
    i = 1
    For j = 1 To MaxCol
        NumEnergized = HowManyEnergized(i, j, "Down")
        If NumEnergized > MaxNumEnergized Then
            MaxNumEnergized = NumEnergized
            i_Max = i
            j_Max = j
            dir_Max = "Down"
        End If
    Next j

    'Bottom Row
    Application.StatusBar = "Working on Bottom Row"
    i = MaxRow
    For j = 1 To MaxCol
        NumEnergized = HowManyEnergized(i, j, "Up")
        If NumEnergized > MaxNumEnergized Then
            MaxNumEnergized = NumEnergized
            i_Max = i
            j_Max = j
            dir_Max = "Up"
        End If
    Next j

    'Left Col
    Application.StatusBar = "Working on Left Col"
    j = 1
    For i = 1 To MaxRow
        NumEnergized = HowManyEnergized(i, j, "Right")
        If NumEnergized > MaxNumEnergized Then
            MaxNumEnergized = NumEnergized
            i_Max = i
            j_Max = j
            dir_Max = "Right"
        End If
    Next i

    'Right Col
    Application.StatusBar = "Working on Right Col"
    j = MaxCol
    For i = 1 To MaxRow
        NumEnergized = HowManyEnergized(i, j, "Left")
        If NumEnergized > MaxNumEnergized Then
            MaxNumEnergized = NumEnergized
            i_Max = i
            j_Max = j
            dir_Max = "Left"
        End If
    Next i
    
    Application.StatusBar = False
    NumEnergized = HowManyEnergized(i_Max, j_Max, dir_Max)
    
    MsgBox "Max Energy: " & MaxNumEnergized & " at (" & i_Max & ", " & j_Max & ")"
End Sub

Function HowManyEnergized(ByVal StartRow As Integer, ByVal StartCol As Integer, ByVal StartDir As String) As Integer
    Range("A1").Select
    MaxRow = ActiveCell.CurrentRegion.Rows.Count
    MaxCol = ActiveCell.CurrentRegion.Columns.Count
    
    Dim matGrid() As ptGrid
    ReDim matGrid(1 To MaxRow, 1 To MaxCol)
    
    For i = 1 To MaxRow
        For j = 1 To MaxCol
            matGrid(i, j).boolEnergized = False
            matGrid(i, j).boolUp = False
            matGrid(i, j).boolDown = False
            matGrid(i, j).boolRight = False
            matGrid(i, j).boolLeft = False
        Next j
    Next i
    
    Dim arrPath() As TrajPath
    ReDim arrPath(1 To 1)
    
    NumPaths = 1
    arrPath(1).intRow = StartRow
    arrPath(1).intCol = StartCol
    arrPath(1).boolExists = True
    Select Case StartDir
        Case "Up"
            Select Case Cells(StartRow, StartCol).Value
                Case ".", "|": arrPath(1).strDir = "Up"
                Case "\":      arrPath(1).strDir = "Left"
                Case "/":      arrPath(1).strDir = "Right"
                Case "-":
                    NumPaths = NumPaths + 1
                    arrPath(1).strDir = "Right"
                    
                    ReDim Preserve arrPath(1 To NumPaths)
                    arrPath(NumPaths).strDir = "Left"
                    arrPath(NumPaths).intRow = StartRow
                    arrPath(NumPaths).intCol = StartCol
                    arrPath(NumPaths).boolExists = True
            End Select
        Case "Down"
            Select Case Cells(StartRow, StartCol).Value
                Case ".", "|": arrPath(1).strDir = "Down"
                Case "\":      arrPath(1).strDir = "Right"
                Case "/":      arrPath(1).strDir = "Left"
                Case "-":
                    NumPaths = NumPaths + 1
                    arrPath(1).strDir = "Right"
                    
                    ReDim Preserve arrPath(1 To NumPaths)
                    arrPath(NumPaths).strDir = "Left"
                    arrPath(NumPaths).intRow = StartRow
                    arrPath(NumPaths).intCol = StartCol
                    arrPath(NumPaths).boolExists = True
            End Select
        Case "Right"
            Select Case Cells(StartRow, StartCol).Value
                Case ".", "-": arrPath(1).strDir = "Right"
                Case "/":      arrPath(1).strDir = "Up"
                Case "\":      arrPath(1).strDir = "Down"
                Case "|":
                    NumPaths = NumPaths + 1
                    arrPath(1).strDir = "Up"
                    
                    ReDim Preserve arrPath(1 To NumPaths)
                    arrPath(NumPaths).strDir = "Down"
                    arrPath(NumPaths).intRow = StartRow
                    arrPath(NumPaths).intCol = StartCol
                    arrPath(NumPaths).boolExists = True
            End Select
            
        Case "Left"
            Select Case Cells(StartRow, StartCol).Value
                Case ".", "-": arrPath(1).strDir = "Left"
                Case "/":      arrPath(1).strDir = "Down"
                Case "\":      arrPath(1).strDir = "Up"
                Case "|":
                    NumPaths = NumPaths + 1
                    arrPath(1).strDir = "Up"
                    
                    ReDim Preserve arrPath(1 To NumPaths)
                    arrPath(NumPaths).strDir = "Down"
                    arrPath(NumPaths).intRow = StartRow
                    arrPath(NumPaths).intCol = StartCol
                    arrPath(NumPaths).boolExists = True
            End Select
    End Select
    
    
    For i = 1 To NumPaths
        matGrid(StartRow, StartCol).boolEnergized = True
        Select Case arrPath(i).strDir
            Case "Up":    matGrid(StartRow, StartCol).boolUp = True
            Case "Down":  matGrid(StartRow, StartCol).boolDown = True
            Case "Right": matGrid(StartRow, StartCol).boolRight = True
            Case "Left":  matGrid(StartRow, StartCol).boolLeft = True
        End Select
    Next i
    
    Dim boolMoved As Boolean
    
    Do
        boolDone = True
        For i = 1 To NumPaths
            If arrPath(i).boolExists Then
                boolMoved = MoveIt(arrPath(i), MaxRow, MaxCol)
                If boolMoved Then
                    boolNotRepeated = EnergizeGrid(arrPath(i), matGrid(arrPath(i).intRow, arrPath(i).intCol))
                    If boolNotRepeated Then
                        boolDone = False
                        NewDir = NextDir(arrPath(i))
                        Select Case NewDir
                            Case intUp: arrPath(i).strDir = "Up"
                            Case intDown: arrPath(i).strDir = "Down"
                            Case intRight: arrPath(i).strDir = "Right"
                            Case intLeft: arrPath(i).strDir = "Left"
                            Case intUp + intDown
                                NumPaths = NumPaths + 1
                                arrPath(i).strDir = "Up"
                                
                                ReDim Preserve arrPath(1 To NumPaths)
                                arrPath(NumPaths).strDir = "Down"
                                arrPath(NumPaths).intRow = arrPath(i).intRow
                                arrPath(NumPaths).intCol = arrPath(i).intCol
                                arrPath(NumPaths).boolExists = True
                            Case intRight + intLeft
                                NumPaths = NumPaths + 1
                                arrPath(i).strDir = "Right"
                                
                                ReDim Preserve arrPath(1 To NumPaths)
                                arrPath(NumPaths).strDir = "Left"
                                arrPath(NumPaths).intRow = arrPath(i).intRow
                                arrPath(NumPaths).intCol = arrPath(i).intCol
                                arrPath(NumPaths).boolExists = True
                        End Select
                    Else
                        arrPath(i).boolExists = False
                    End If
                Else
                    arrPath(i).boolExists = False
                End If
            End If
        Next i
    Loop Until boolDone
    
    k = MaxCol + 3
    Range(Cells(1, k + 1), Cells(MaxRow, k + MaxCol)).ClearContents
    For i = 1 To MaxRow
        For j = 1 To MaxCol
            If matGrid(i, j).boolEnergized Then
                Cells(i, j + k).Value = "#"
            End If
        Next j
    Next i
    
    HowManyEnergized = WorksheetFunction.CountA(Range(Cells(1, k + 1), Cells(MaxRow, k + MaxCol)))
End Function
