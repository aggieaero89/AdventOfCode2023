Attribute VB_Name = "Day10"
Sub Part1()
'
' Macro1 Macro
'
    Range("A1").Select
    LastRow = ActiveCell.CurrentRegion.Rows.Count
    LastCol = ActiveCell.CurrentRegion.Columns.Count
    
    Dim rngStart As Range
    With ActiveSheet.Cells
        Set rngStart = .Find("S", LookIn:=xlValues)
        If rngStart Is Nothing Then
            MsgBox "No S for start found"
            End
        Else
            rngStart.Select
            ActiveCell.Interior.Color = 5296274
        End If
    End With
    
    If CheckNextPipe("North") Then
        NextDir = "North"
    ElseIf CheckNextPipe("East") Then
        NextDir = "East"
    ElseIf CheckNextPipe("South") Then
        NextDir = "South"
    ElseIf CheckNextPipe("West") Then
        NextDir = "West"
    Else
        MsgBox "No valid next direction"
        End
    End If
    
    Kount = 0
    Do
        Select Case NextDir
            Case "North"
                ActiveCell.Offset(-1, 0).Select
                ActiveCell.Interior.Color = 5296274
            Case "East"
                ActiveCell.Offset(0, 1).Select
                ActiveCell.Interior.Color = 5296274
            Case "South"
                ActiveCell.Offset(1, 0).Select
                ActiveCell.Interior.Color = 5296274
            Case "West"
                ActiveCell.Offset(0, -1).Select
                ActiveCell.Interior.Color = 5296274
        End Select
        Kount = Kount + 1
        
        PrevDir = NextDir
        CurrPipe = ActiveCell.Value
        If CurrPipe = "S" Then Exit Do
        
        Select Case CurrPipe
            Case "|": If PrevDir = "North" Then NextDir = "North" Else NextDir = "South"
            Case "-": If PrevDir = "East" Then NextDir = "East" Else NextDir = "West"
            Case "L": If PrevDir = "South" Then NextDir = "East" Else NextDir = "North"
            Case "J": If PrevDir = "South" Then NextDir = "West" Else NextDir = "North"
            Case "7": If PrevDir = "North" Then NextDir = "West" Else NextDir = "South"
            Case "F": If PrevDir = "North" Then NextDir = "East" Else NextDir = "South"
        End Select
           
    Loop 'Until Kount > 1000
    
    MsgBox "Number of steps from start to farthest point is " & Kount / 2
End Sub

Function CheckNextPipe(ByVal strDir As String) As Boolean
    Select Case strDir
        Case "North"
            'North
            strNext = ActiveCell.Offset(-1, 0).Value
            Select Case strNext
                Case "7", "|", "F": boolValid = True
                Case Else:          boolValid = False
            End Select
            
        Case "East"
            'East
            strNext = ActiveCell.Offset(0, 1).Value
            Select Case strNext
                Case "J", "-", "7": boolValid = True
                Case Else:          boolValid = False
            End Select
            
        Case "South"
            'South
            strNext = ActiveCell.Offset(1, 0).Value
            Select Case strNext
                Case "J", "|", "L": boolValid = True
                Case Else:          boolValid = False
            End Select
        
        Case "West"
            'West
            strNext = ActiveCell.Offset(0, -1).Value
            Select Case strNext
                Case "F", "-", "L": boolValid = True
                Case Else:          boolValid = False
            End Select
            
    End Select
    
    CheckNextPipe = boolValid
End Function
Sub Part2()
'
' Macro1 Macro
'
    Range("A1").Select
    LastRow = ActiveCell.CurrentRegion.Rows.Count
    LastCol = ActiveCell.CurrentRegion.Columns.Count
    
    OutputRow = 1
    OutputCol_X = LastCol + 2
    OutputCol_Y = LastCol + 3
    
    Dim rngStart As Range
    With ActiveSheet.Cells
        Set rngStart = .Find("S", LookIn:=xlValues)
        If rngStart Is Nothing Then
            MsgBox "No S for start found"
            End
        Else
            rngStart.Select
            ActiveCell.Interior.Color = 5296274
            Cells(OutputRow, OutputCol_X).Value = ActiveCell.Row
            Cells(OutputRow, OutputCol_Y).Value = ActiveCell.Column
        End If
    End With
    
    If CheckNextPipe("North") Then
        NextDir = "North"
    ElseIf CheckNextPipe("East") Then
        NextDir = "East"
    ElseIf CheckNextPipe("South") Then
        NextDir = "South"
    ElseIf CheckNextPipe("West") Then
        NextDir = "West"
    Else
        MsgBox "No valid next direction"
        End
    End If
    
    Kount = 0
    Do
        Select Case NextDir
            Case "North"
                ActiveCell.Offset(-1, 0).Select
                ActiveCell.Interior.Color = 5296274
            Case "East"
                ActiveCell.Offset(0, 1).Select
                ActiveCell.Interior.Color = 5296274
            Case "South"
                ActiveCell.Offset(1, 0).Select
                ActiveCell.Interior.Color = 5296274
            Case "West"
                ActiveCell.Offset(0, -1).Select
                ActiveCell.Interior.Color = 5296274
        End Select
        Kount = Kount + 1
        
        OutputRow = OutputRow + 1
        Cells(OutputRow, OutputCol_X).Value = ActiveCell.Row
        Cells(OutputRow, OutputCol_Y).Value = ActiveCell.Column
        
        PrevDir = NextDir
        CurrPipe = ActiveCell.Value
        If CurrPipe = "S" Then Exit Do
        
        Select Case CurrPipe
            Case "|": If PrevDir = "North" Then NextDir = "North" Else NextDir = "South"
            Case "-": If PrevDir = "East" Then NextDir = "East" Else NextDir = "West"
            Case "L": If PrevDir = "South" Then NextDir = "East" Else NextDir = "North"
            Case "J": If PrevDir = "South" Then NextDir = "West" Else NextDir = "North"
            Case "7": If PrevDir = "North" Then NextDir = "West" Else NextDir = "South"
            Case "F": If PrevDir = "North" Then NextDir = "East" Else NextDir = "South"
        End Select
           
    Loop 'Until Kount > 1000
    
'======================================= Count Points Inside Polygon
    Dim polygon As Range
    Set polygon = Cells(1, OutputCol_X).CurrentRegion
    
    Knt = 0
    For i = 1 To LastRow
        For j = 1 To LastCol
            If Cells(i, j).Interior.Color <> 5296274 Then
                If PtInPoly(i, j, polygon) Then
                    Knt = Knt + 1
                    Cells(i, j).Interior.Color = 255
                End If
            End If
        Next j
    Next i
    
    MsgBox "Number of points in pipes is " & Knt
End Sub

Public Function PtInPoly(ByVal Xcoord As Integer, ByVal Ycoord As Integer, ByRef polygon As Range) As Boolean
'https://www.excelfox.com/forum/showthread.php/1579-Test-Whether-A-Point-Is-In-A-Polygon-Or-Not
'
'NOTE #1: The polygon must be closed, meaning the first listed point and the last listed point must be the same. If they are not the same, the function will raise "Error #998 - Polygon Does Not Close!" if the function was called from other VB code or it will return #UnclosedPolygon! if called from the worksheet. Normally, if called from a worksheet, you would probably be using the function in a formula something like this...
'
'=IF(PtInPoly(B9,C9,E18:F37),"In Polygon","Out Polygon")
'
'In that case, the formula will return a #VALUE! error, not the #UnclosedPolygon! error, because the returned value to the IF function is not a Boolean; however, if you select the "PtInPoly(B9,C9,E18:F37)" part of the function in the Formula Bar and press F9, it will show you the returned value from the PtInPoly function as being #UnclosedPolygon!.
'
'NOTE #2: The range or array specified for the third argument must be two-dimensional. If it is not, then the function will raise "Error #999 - Array Has Wrong Number Of Coordinates!" if the function was called from other VB code or it will return #WrongNumberOfCoordinates! if called from the worksheet. Error reporting when called from the worksheet will be the same as described in NOTE #1.
    Dim x As Long, NumSidesCrossed As Long, m As Double, b As Double, Poly As Variant
    
    Poly = polygon
    For x = LBound(Poly) To UBound(Poly) - 1
        If Poly(x, 1) > Xcoord Xor Poly(x + 1, 1) > Xcoord Then
            m = (Poly(x + 1, 2) - Poly(x, 2)) / (Poly(x + 1, 1) - Poly(x, 1))
            b = (Poly(x, 2) * Poly(x + 1, 1) - Poly(x, 1) * Poly(x + 1, 2)) / (Poly(x + 1, 1) - Poly(x, 1))
            If m * Xcoord + b > Ycoord Then NumSidesCrossed = NumSidesCrossed + 1
        End If
    Next
    
    PtInPoly = CBool(NumSidesCrossed Mod 2)
End Function

'======================================================================================================================================
'Public Function PtInPoly_withErrorChecking(Xcoord As Double, Ycoord As Double, Polygon As Variant) As Variant
''https://www.excelfox.com/forum/showthread.php/1579-Test-Whether-A-Point-Is-In-A-Polygon-Or-Not
''
''NOTE #1: The polygon must be closed, meaning the first listed point and the last listed point must be the same. If they are not the same, the function will raise "Error #998 - Polygon Does Not Close!" if the function was called from other VB code or it will return #UnclosedPolygon! if called from the worksheet. Normally, if called from a worksheet, you would probably be using the function in a formula something like this...
''
''=IF(PtInPoly(B9,C9,E18:F37),"In Polygon","Out Polygon")
''
''In that case, the formula will return a #VALUE! error, not the #UnclosedPolygon! error, because the returned value to the IF function is not a Boolean; however, if you select the "PtInPoly(B9,C9,E18:F37)" part of the function in the Formula Bar and press F9, it will show you the returned value from the PtInPoly function as being #UnclosedPolygon!.
''
''NOTE #2: The range or array specified for the third argument must be two-dimensional. If it is not, then the function will raise "Error #999 - Array Has Wrong Number Of Coordinates!" if the function was called from other VB code or it will return #WrongNumberOfCoordinates! if called from the worksheet. Error reporting when called from the worksheet will be the same as described in NOTE #1.
''NOTE #3: Points extremely close to, or theoretically lying on, the polygon borders may or may not report back correctly... the vagrancies of floating point math, coupled with the limitations that the "significant digits limit" in VBA imposes, can rear its ugly head in those circumstances producing values that can calculate to either side of the polygon border being tested (remember, a mathematical line has no thickness, so it does not take too much of a difference in the significant digits to "move" a calculated point's position to one side or the other of such a line).
''NOTE #4: While I think error checking is a good thing, the setup for this function is rather straightforward and, with the possible exception of the requirement for the first and last point needing to be the same, easy enough for the programmer to maintain control over during coding. If you feel confident in your ability to meet the needs of NOTE #1 and NOTE #2 without having the code "looking over your shoulder", then the function can be simplified down to this compact code...
'
'  Dim x As Long, m As Double, b As Double, Poly As Variant
'  Dim LB1 As Long, LB2 As Long, UB1 As Long, UB2 As Long, NumSidesCrossed As Long
'  Poly = Polygon
'  If Not (Poly(LBound(Poly), 1) = Poly(UBound(Poly), 1) And _
'        Poly(LBound(Poly), 2) = Poly(UBound(Poly), 2)) Then
'    If TypeOf Application.Caller Is Range Then
'      PtInPoly = "#UnclosedPolygon!"
'    Else
'      Err.Raise 998, , "Polygon Does Not Close!"
'    End If
'    Exit Function
'  ElseIf UBound(Poly, 2) - LBound(Poly, 2) <> 1 Then
'    If TypeOf Application.Caller Is Range Then
'      PtInPoly = "#WrongNumberOfCoordinates!"
'    Else
'      Err.Raise 999, , "Array Has Wrong Number Of Coordinates!"
'    End If
'    Exit Function
'  End If
'  For x = LBound(Poly) To UBound(Poly) - 1
'    If Poly(x, 1) > Xcoord Xor Poly(x + 1, 1) > Xcoord Then
'      m = (Poly(x + 1, 2) - Poly(x, 2)) / (Poly(x + 1, 1) - Poly(x, 1))
'      b = (Poly(x, 2) * Poly(x + 1, 1) - Poly(x, 1) * Poly(x + 1, 2)) / (Poly(x + 1, 1) - Poly(x, 1))
'      If m * Xcoord + b > Ycoord Then NumSidesCrossed = NumSidesCrossed + 1
'    End If
'  Next
'  PtInPoly = CBool(NumSidesCrossed Mod 2)
'End Function

