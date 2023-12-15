Attribute VB_Name = "Day8"
Sub Part1()
'
' Macro1 Macro
'
    Dim rngNodes As Range
    
    Range("A1").Select
    strLR = ActiveCell.Value
    maxNum = Len(strLR)
    
    LastRow = Range("C3").End(xlDown).Row
    
    
    boolSuccess = False
    Kount = 0
    CurrNode = "AAA"
    Do
        For i = 1 To maxNum
            Kount = Kount + 1
            Select Case Mid(strLR, i, 1)
                Case "L":
                    NextNode = Application.WorksheetFunction.VLookup(CurrNode, Range("C3:E" & LastRow), 2, False)
                Case "R":
                    NextNode = Application.WorksheetFunction.VLookup(CurrNode, Range("C3:E" & LastRow), 3, False)
                Case Else:
                    MsgBox "Invalid direction"
                    End
            End Select
            
            If NextNode = "ZZZ" Then
                boolSuccess = True
            Else
                CurrNode = NextNode
            End If
            
        Next i
        
    Loop Until boolSuccess
    
    MsgBox "Number of steps to reach ZZZ is " & Kount
End Sub

Sub Part2()
'
' Macro1 Macro
'
    Dim rngNodes As Range
    
    Range("A1").Select
    strLR = ActiveCell.Value
    maxNum = Len(strLR)
    
    Const intUBound = 5
    
    Range("C3").Select
    LastRow = Range("C3").End(xlDown).Row
    
    Dim arrCurrNode(0 To intUBound) As String
    For j = 0 To intUBound
        arrCurrNode(j) = ActiveCell.Offset(j, 0).Value
    Next j
    Dim arrNextNode(0 To intUBound) As String
    
    boolSuccess = False
    Kount = 0
    Do
        For i = 1 To maxNum
            Kount = Kount + 1
            Application.StatusBar = "Command = " & Kount
            Select Case Mid(strLR, i, 1)
                Case "L":
                    For j = 0 To intUBound
                        arrNextNode(j) = Application.WorksheetFunction.VLookup(arrCurrNode(j), Range("C3:E" & LastRow), 2, False)
                    Next j
                Case "R":
                    For j = 0 To intUBound
                        arrNextNode(j) = Application.WorksheetFunction.VLookup(arrCurrNode(j), Range("C3:E" & LastRow), 3, False)
                    Next j
                Case Else:
                    MsgBox "Invalid direction"
                    End
            End Select
            
            For j = 0 To intUBound
                If Right(arrNextNode(j), 1) = "Z" Then
                    boolSuccess = True
                Else
                    boolSuccess = False
                    For k = 0 To intUBound
                        arrCurrNode(k) = arrNextNode(k)
                    Next k
                    Exit For
                End If
            Next j
            If boolSuccess Then Exit For
        Next i
        
    Loop Until boolSuccess
    
    Application.StatusBar = False

    MsgBox "Number of steps where all nodes simultaneously end with Z is " & Kount
End Sub

Sub ATest()
'
' Macro1 Macro
'
    Dim rngNodes As Range
    
    Range("A1").Select
    strLR = ActiveCell.Value
    maxNum = Len(strLR)
    
    LastRow = Range("C3").End(xlDown).Row
    
    
    boolSuccess = False
    Kount = 0
    
    OutputRow = 3
    'CurrNode = "DVA": OutputCol = 8
    'CurrNode = "MPA": OutputCol = 9
    'CurrNode = "TDA": OutputCol = 10
    'CurrNode = "AAA": OutputCol = 11
    'CurrNode = "FJA": OutputCol = 12
    CurrNode = "XPA": OutputCol = 13
    
    Do
        For i = 1 To maxNum
            Kount = Kount + 1
            Select Case Mid(strLR, i, 1)
                Case "L":
                    NextNode = Application.WorksheetFunction.VLookup(CurrNode, Range("C3:E" & LastRow), 2, False)
                Case "R":
                    NextNode = Application.WorksheetFunction.VLookup(CurrNode, Range("C3:E" & LastRow), 3, False)
                Case Else:
                    MsgBox "Invalid direction"
                    End
            End Select
            
            If Right(NextNode, 1) = "Z" Then
                boolSuccess = True
            Else
                CurrNode = NextNode
            End If
            
        Next i
        
        If boolSuccess Then
            AnswerYes = MsgBox("Number of steps to reach ZZZ is " & Kount & ". Keep going?", vbYesNo)
            Cells(OutputRow, OutputCol).Value = Kount
            If AnswerYes = vbYes Then
                boolSuccess = False
                OutputRow = OutputRow + 1
                Kount = 0
            Else
                boolSuccess = True
            End If
        End If
    Loop Until boolSuccess
    
End Sub

Function FindCommon(ByVal Value1 As Double, ByVal Rate1 As Double, ByRef Factor1 As Double, ByVal Value2 As Double, ByVal Rate2 As Double, ByRef Factor2 As Double) As Double
    Dim dblValue1 As Double
    Dim dblValue2 As Double
    
    dblValue1 = Value1 + Rate1 * Factor1
    dblValue2 = Value2 + Rate2 * Factor2
    
    If dblValue1 < dblValue2 Then
        Factor1 = Factor1 + WorksheetFunction.RoundDown((dblValue2 - dblValue1) / Rate1, 0)
    ElseIf dblValue2 < dblValue1 Then
        Factor2 = Factor2 + WorksheetFunction.RoundDown((dblValue1 - dblValue2) / Rate2, 0)
    End If

    Do
        If dblValue1 < dblValue2 Then
            Factor1 = Factor1 + 1
            dblValue1 = Value1 + Rate1 * Factor1
        ElseIf dblValue2 < dblValue1 Then
            Factor2 = Factor2 + 1
            dblValue2 = Value2 + Rate2 * Factor2
        End If
        Application.StatusBar = "Val1: " & dblValue1 & " Val2: " & dblValue2
    Loop Until dblValue1 = dblValue2
    
    FindCommon = dblValue1
    Application.StatusBar = False
End Function


Sub BTest()
    Dim Afactor As Double, Bfactor As Double, Cfactor As Double, Dfactor As Double, Efactor As Double, Ffactor As Double
    
    A = 23147
    Arate = 7032
    Afactor = 0
    Aprime = A + Arate * Afactor
    
    b = 19631
    Brate = 17287
    Bfactor = 0
    Bprime = b + Brate * Bfactor

    C = 12599
    Crate = 4688
    Cfactor = 0
    Cprime = C + Crate * Cfactor
    
    D = 21389
    Drate = 293
    Dfactor = 0
    Dprime = D + Drate * Dfactor
    
    E = 17873
    Erate = 1465
    Efactor = 0
    Eprime = E + Erate * Efactor
    
    F = 20803
    Frate = 2344
    Ffactor = 0
    Fprime = F + Frate * Ffactor
    
    Do
        N1 = FindCommon(A, Arate, Afactor, b, Brate, Bfactor)
        N2 = FindCommon(b, Brate, Bfactor, C, Crate, Cfactor)
        N3 = FindCommon(C, Crate, Cfactor, D, Drate, Dfactor)
        N4 = FindCommon(D, Drate, Dfactor, E, Erate, Efactor)
        N5 = FindCommon(E, Erate, Efactor, F, Frate, Ffactor)
        
        If N1 <> N5 Then
            N1 = FindCommon(A, Arate, Afactor, F, Frate, Ffactor)
            boolDone = False
        Else
            boolDone = True
        End If
    Loop Until boolDone
    
    MsgBox "The number is: " & N1
End Sub
