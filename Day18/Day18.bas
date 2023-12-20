Attribute VB_Name = "Day18"
Dim CurrRow As Long
Dim CurrCol As Long

Sub MoveIt_ORIG(ByVal strDir As String, ByVal intDist As Integer, ByVal hexColor As String)
    Select Case strDir
        Case "U":
            Range(ActiveCell, ActiveCell.Offset(-intDist, 0)).Value = "#"
            Range(ActiveCell.Offset(-1, 0), ActiveCell.Offset(-intDist, 0)).Interior.Color = WorksheetFunction.Hex2Dec(hexColor)
            ActiveCell.Offset(-intDist, 0).Select
        Case "D":
            Range(ActiveCell, ActiveCell.Offset(intDist, 0)).Value = "#"
            Range(ActiveCell.Offset(1, 0), ActiveCell.Offset(intDist, 0)).Interior.Color = WorksheetFunction.Hex2Dec(hexColor)
            ActiveCell.Offset(intDist, 0).Select
        Case "R":
            Range(ActiveCell, ActiveCell.Offset(0, intDist)).Value = "#"
            Range(ActiveCell.Offset(0, 1), ActiveCell.Offset(0, intDist)).Interior.Color = WorksheetFunction.Hex2Dec(hexColor)
            ActiveCell.Offset(0, intDist).Select
        Case "L":
            Range(ActiveCell, ActiveCell.Offset(0, -intDist)).Value = "#"
            Range(ActiveCell.Offset(0, -1), ActiveCell.Offset(0, -intDist)).Interior.Color = WorksheetFunction.Hex2Dec(hexColor)
            ActiveCell.Offset(0, -intDist).Select
    End Select
End Sub

Sub MoveIt(ByVal strDir As String, ByVal intDist As Integer, ByVal hexColor As String, ByVal strDirPrev, ByVal strDirNext)
    Select Case strDir
        Case "U":
            If strDirPrev = "R" Then
                If strDirNext = "L" Then intDist = intDist - 1
            Else 'strDirPrev = "L"
                If strDirNext = "R" Then intDist = intDist + 1
            End If
            
            Range(ActiveCell, ActiveCell.Offset(-intDist, 0)).Value = "#"
            Range(ActiveCell.Offset(-1, 0), ActiveCell.Offset(-intDist, 0)).Interior.Color = WorksheetFunction.Hex2Dec(hexColor)
            ActiveCell.Offset(-intDist, 0).Select
            
        Case "D":
            If strDirPrev = "R" Then
                If strDirNext = "L" Then intDist = intDist + 1
            Else 'strDirPrev = "L"
                If strDirNext = "R" Then intDist = intDist - 1
            End If
            
            Range(ActiveCell, ActiveCell.Offset(intDist, 0)).Value = "#"
            Range(ActiveCell.Offset(1, 0), ActiveCell.Offset(intDist, 0)).Interior.Color = WorksheetFunction.Hex2Dec(hexColor)
            ActiveCell.Offset(intDist, 0).Select
            
        Case "R":
            If strDirPrev = "U" Then
                If strDirNext = "D" Then intDist = intDist + 1
            Else 'strDirPrev = "D"
                If strDirNext = "U" Then intDist = intDist - 1
            End If
                
            Range(ActiveCell, ActiveCell.Offset(0, intDist)).Value = "#"
            Range(ActiveCell.Offset(0, 1), ActiveCell.Offset(0, intDist)).Interior.Color = WorksheetFunction.Hex2Dec(hexColor)
            ActiveCell.Offset(0, intDist).Select
        Case "L":
            If strDirPrev = "D" Then
                If strDirNext = "U" Then intDist = intDist + 1
            Else 'strDirPrev = "U"
                If strDirNext = "D" Then intDist = intDist - 1
            End If
            
            Range(ActiveCell, ActiveCell.Offset(0, -intDist)).Value = "#"
            Range(ActiveCell.Offset(0, -1), ActiveCell.Offset(0, -intDist)).Interior.Color = WorksheetFunction.Hex2Dec(hexColor)
            ActiveCell.Offset(0, -intDist).Select
    End Select
End Sub

Sub MoveIt_NoPlot(ByVal strDir As String, ByVal lngDist As Long, ByVal strDirPrev, ByVal strDirNext)
    Select Case strDir
        Case "U":
            If strDirPrev = "R" Then
                If strDirNext = "L" Then lngDist = lngDist - 1
            Else 'strDirPrev = "L"
                If strDirNext = "R" Then lngDist = lngDist + 1
            End If
            
            CurrRow = CurrRow - lngDist
            
        Case "D":
            If strDirPrev = "R" Then
                If strDirNext = "L" Then lngDist = lngDist + 1
            Else 'strDirPrev = "L"
                If strDirNext = "R" Then lngDist = lngDist - 1
            End If
            
            CurrRow = CurrRow + lngDist
            
        Case "R":
            If strDirPrev = "U" Then
                If strDirNext = "D" Then lngDist = lngDist + 1
            Else 'strDirPrev = "D"
                If strDirNext = "U" Then lngDist = lngDist - 1
            End If
              
            CurrCol = CurrCol + lngDist
        
        Case "L":
            If strDirPrev = "D" Then
                If strDirNext = "U" Then lngDist = lngDist + 1
            Else 'strDirPrev = "U"
                If strDirNext = "D" Then lngDist = lngDist - 1
            End If
            
            CurrCol = CurrCol - lngDist
    End Select
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

Sub Part1()
'
' Macro1 Macro
'
    ActiveSheet.UsedRange
    
    Range("A1").Select
    i = ActiveCell.Interior.Color
    MaxRow = ActiveCell.CurrentRegion.Rows.Count
    
    OutputRow = 1
    OutputCol_X = 3
    OutputCol_Y = 4
    
    If ActiveSheet.Name = "Sample input" Then
        StartRow = 10
        StartCol = 10
    Else
        StartRow = 500
        StartCol = 500
    End If
    
    Cells(StartRow, StartCol).Select 'Sample 1 input
    Cells(OutputRow, OutputCol_X).Value = ActiveCell.Row
    Cells(OutputRow, OutputCol_Y).Value = ActiveCell.Column

    For i = 1 To MaxRow
        arrInput = Split(Cells(i, 1).Value, " ")
        strDir = arrInput(0)
        intDist = CInt(arrInput(1))
        hexColr = Mid(arrInput(2), 3, Len(arrInput(2)) - 3)
        
        Call MoveIt_ORIG(strDir, intDist, hexColr)
        
        OutputRow = OutputRow + 1
        Cells(OutputRow, OutputCol_X).Value = ActiveCell.Row
        Cells(OutputRow, OutputCol_Y).Value = ActiveCell.Column
    Next i
    
    '===========================================================
    Dim polygon As Range
    Set polygon = Cells(1, OutputCol_X).CurrentRegion
    
    For Each c In ActiveCell.CurrentRegion.Cells
        If ActiveCell.Interior.Color <> 16777215 Then 'no fill color
            If PtInPoly(c.Row, c.Column, polygon) Then
                c.Value = "#"
            End If
        End If
    Next c
    
    MsgBox "The sum of the results is " & WorksheetFunction.CountA(Cells(StartRow, StartCol).CurrentRegion)
End Sub

Sub Part1_NewPlot()
'
' Macro1 Macro
'
    ActiveSheet.UsedRange
    
    Range("A1").Select
    i = ActiveCell.Interior.Color
    MaxRow = ActiveCell.CurrentRegion.Rows.Count
    
    OutputRow = 1
    OutputCol_X = 3
    OutputCol_Y = 4
    
    If ActiveSheet.Name = "Sample input" Then
        StartRow = 10 '10 Sample 1 input
        StartCol = 10 '10 Sample 1 input
    Else
        StartRow = 500
        StartCol = 500
    End If
    
    Cells(StartRow, StartCol).Select 'Sample 1 input
    Cells(OutputRow, OutputCol_X).Value = ActiveCell.Row
    Cells(OutputRow, OutputCol_Y).Value = ActiveCell.Column

    For i = 1 To MaxRow - 1
        If i = 1 Then
            arrInputPrev = Split(Cells(MaxRow, 1).Value, " ")
            strDirPrev = arrInputPrev(0)
        Else
            strDirPrev = strDir
        End If
            
        arrInput = Split(Cells(i, 1).Value, " ")
        strDir = arrInput(0)
        intDist = CInt(arrInput(1))
        hexColr = Mid(arrInput(2), 3, Len(arrInput(2)) - 3)
        
        arrInputNext = Split(Cells(i + 1, 1).Value, " ")
        strDirNext = arrInputNext(0)
        
        Call MoveIt(strDir, intDist, hexColr, strDirPrev, strDirNext)
        
        OutputRow = OutputRow + 1
        Cells(OutputRow, OutputCol_X).Value = ActiveCell.Row
        Cells(OutputRow, OutputCol_Y).Value = ActiveCell.Column
    Next i
    
    OutputRow = OutputRow + 1
    Cells(OutputRow, OutputCol_X).Value = StartRow
    Cells(OutputRow, OutputCol_Y).Value = StartCol
    
    '===========================================================
    Dim polygon_x As Range, polygon_y As Range
    MaxRow = Cells(1, OutputCol_X).CurrentRegion.Rows.Count
    Set polygon_x = Range(Cells(1, OutputCol_X), Cells(MaxRow, OutputCol_X))
    Set polygon_y = Range(Cells(1, OutputCol_Y), Cells(MaxRow, OutputCol_Y))
    
    MsgBox "The sum of the results is " & Shoelace(polygon_x, polygon_y)
End Sub

Sub Part1_NoPlot()
'
' Macro1 Macro
'
    ActiveSheet.UsedRange
    
    Range("A1").Select
    i = ActiveCell.Interior.Color
    MaxRow = ActiveCell.CurrentRegion.Rows.Count
    
    OutputRow = 1
    OutputCol_X = 3
    OutputCol_Y = 4
    
    If ActiveSheet.Name = "Sample input" Then
        StartRow = 10 '10 Sample 1 input
        StartCol = 10 '10 Sample 1 input
    Else
        StartRow = 500
        StartCol = 500
    End If
    
    CurrRow = StartRow
    CurrCol = StartCol
    
    Cells(OutputRow, OutputCol_X).Value = CurrRow
    Cells(OutputRow, OutputCol_Y).Value = CurrCol

    For i = 1 To MaxRow - 1
        If i = 1 Then
            arrInputPrev = Split(Cells(MaxRow, 1).Value, " ")
            strDirPrev = arrInputPrev(0)
        Else
            strDirPrev = strDir
        End If
            
        arrInput = Split(Cells(i, 1).Value, " ")
        strDir = arrInput(0)
        intDist = CInt(arrInput(1))
        hexColr = Mid(arrInput(2), 3, Len(arrInput(2)) - 3)
        
        arrInputNext = Split(Cells(i + 1, 1).Value, " ")
        strDirNext = arrInputNext(0)
        
        Call MoveIt_NoPlot(strDir, intDist, strDirPrev, strDirNext)
        
        OutputRow = OutputRow + 1
        Cells(OutputRow, OutputCol_X).Value = CurrRow
        Cells(OutputRow, OutputCol_Y).Value = CurrCol
    Next i
    
    OutputRow = OutputRow + 1
    Cells(OutputRow, OutputCol_X).Value = StartRow
    Cells(OutputRow, OutputCol_Y).Value = StartCol
    
    '===========================================================
    Dim polygon_x As Range, polygon_y As Range
    MaxRow = Cells(1, OutputCol_X).CurrentRegion.Rows.Count
    Set polygon_x = Range(Cells(1, OutputCol_X), Cells(MaxRow, OutputCol_X))
    Set polygon_y = Range(Cells(1, OutputCol_Y), Cells(MaxRow, OutputCol_Y))
    
    MsgBox "The sum of the results is " & Shoelace(polygon_x, polygon_y)
End Sub

Function Find_Direction(intDir) As String
    Select Case intDir
        Case 0: Find_Direction = "R"
        Case 1: Find_Direction = "D"
        Case 2: Find_Direction = "L"
        Case 3: Find_Direction = "U"
    End Select
End Function

Sub Part2()
'
' Macro1 Macro
'
    ActiveSheet.UsedRange
    
    Range("A1").Select
    i = ActiveCell.Interior.Color
    MaxRow = ActiveCell.CurrentRegion.Rows.Count
    
    OutputRow = 1
    OutputCol_X = 3
    OutputCol_Y = 4
    
    If ActiveSheet.Name = "Sample input" Then
        StartRow = 10 '10 Sample 1 input
        StartCol = 10 '10 Sample 1 input
    Else
        StartRow = 2000000
        StartCol = 2000000
    End If
    
    CurrRow = StartRow
    CurrCol = StartCol
    
    Cells(OutputRow, OutputCol_X).Value = CurrRow
    Cells(OutputRow, OutputCol_Y).Value = CurrCol

    For i = 1 To MaxRow - 1
        If i = 1 Then
            arrInputPrev = Split(Cells(MaxRow, 1).Value, "#")
            hexPrev = Left(arrInputPrev(1), Len(arrInputPrev(1)) - 1)
            strDirPrev = Find_Direction(Int(Right(hexPrev, 1)))
        Else
            strDirPrev = strDir
        End If
            
        arrInput = Split(Cells(i, 1).Value, "#")
        hexCommand = Left(arrInput(1), Len(arrInput(1)) - 1)
        intDist = WorksheetFunction.Hex2Dec(Left(hexCommand, Len(hexCommand) - 1))
        strDir = Find_Direction(Int(Right(hexCommand, 1)))
        
        arrInputNext = Split(Cells(i + 1, 1).Value, "#")
        hexNext = Left(arrInputNext(1), Len(arrInputNext(1)) - 1)
        strDirNext = Find_Direction(Int(Right(hexNext, 1)))
        
        Call MoveIt_NoPlot(strDir, intDist, strDirPrev, strDirNext)
        
        OutputRow = OutputRow + 1
        Cells(OutputRow, OutputCol_X).Value = CurrRow
        Cells(OutputRow, OutputCol_Y).Value = CurrCol
    Next i
    
    OutputRow = OutputRow + 1
    Cells(OutputRow, OutputCol_X).Value = StartRow
    Cells(OutputRow, OutputCol_Y).Value = StartCol
    
    '===========================================================
    Dim polygon_x As Range, polygon_y As Range
    MaxRow = Cells(1, OutputCol_X).CurrentRegion.Rows.Count
    Set polygon_x = Range(Cells(1, OutputCol_X), Cells(MaxRow, OutputCol_X))
    Set polygon_y = Range(Cells(1, OutputCol_Y), Cells(MaxRow, OutputCol_Y))
    
    MsgBox "The sum of the results is " & Shoelace(polygon_x, polygon_y)
End Sub

Public Function Shoelace(Xs As Range, Ys As Range) As Double
    Dim Area As Double
    If Xs.Rows.Count = Ys.Rows.Count Then
        For i = 1 To Xs.Rows.Count - 1
            Area = Area + (Xs(i + 1) + Xs(i)) * (Ys(i + 1) - Ys(i))
        Next i
        'Use the coordinates of the first point to "close" the polygon.
        Area = Area + (Xs(1) + Xs(Xs.Rows.Count)) * (Ys(1) - Ys(Ys.Rows.Count))
    End If
    'Finally, calculate the polygon area.
    Shoelace = Abs(Area / 2)
End Function
