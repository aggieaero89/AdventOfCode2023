Attribute VB_Name = "Day9"
Sub Part1()
'
' Macro1 Macro
'
    Dim arrLngInput() As Long
    
    Range("A1").Select
    
    SumValues = 0
    
    Do While Not IsEmpty(ActiveCell.Value)
    
        arrStrInput = Split(ActiveCell.Value, " ")
        max_i = UBound(arrStrInput)
        
        ReDim arrLngInput(LBound(arrStrInput) To UBound(arrStrInput))
        For i = LBound(arrStrInput) To UBound(arrStrInput)
            arrLngInput(i) = CLng(arrStrInput(i))
        Next i
        
        NextDiff = DiffArray(arrLngInput)
        NextSequence = arrLngInput(max_i) + NextDiff
        SumValues = SumValues + NextSequence
        
        ActiveCell.Offset(1, 0).Select
    Loop
    
    MsgBox "Sum of the extrapolated values is " & SumValues
End Sub

Function DiffArray(arr As Variant) As Long
    Dim arrDiffs() As Long
    
    max_i = UBound(arr) - 1
    
    ReDim arrDiffs(LBound(arr) To max_i)
    For i = LBound(arr) To max_i
        arrDiffs(i) = arr(i + 1) - arr(i)
    Next i
    
    boolEqual = True
    For i = LBound(arrDiffs) To max_i - 1
        If arrDiffs(i) <> arrDiffs(i + 1) Then boolEqual = False
    Next i
    
    If boolEqual And (arrDiffs(0) = 0) Then
        DiffArray = 0
    Else
        DiffArray = DiffArray(arrDiffs) + arrDiffs(max_i)
    End If

End Function

Function DiffArrayPrev(arr As Variant) As Long
    Dim arrDiffs() As Long
    
    max_i = UBound(arr) - 1
    
    ReDim arrDiffs(LBound(arr) To max_i)
    For i = LBound(arr) To max_i
        arrDiffs(i) = arr(i + 1) - arr(i)
    Next i
    
    boolEqual = True
    For i = LBound(arrDiffs) To max_i - 1
        If arrDiffs(i) <> arrDiffs(i + 1) Then boolEqual = False
    Next i
    
    If boolEqual And (arrDiffs(0) = 0) Then
        DiffArrayPrev = 0
    Else
        DiffArrayPrev = arrDiffs(0) - DiffArrayPrev(arrDiffs)
    End If

End Function

Sub Part2()
'
' Macro1 Macro
'
    Dim arrLngInput() As Long
    
    Range("A1").Select
    
    SumValues = 0
    
    Do While Not IsEmpty(ActiveCell.Value)
    
        arrStrInput = Split(ActiveCell.Value, " ")
        max_i = UBound(arrStrInput)
        
        ReDim arrLngInput(LBound(arrStrInput) To UBound(arrStrInput))
        For i = LBound(arrStrInput) To UBound(arrStrInput)
            arrLngInput(i) = CLng(arrStrInput(i))
        Next i
        
        PrevDiff = DiffArrayPrev(arrLngInput)
        PrevSequence = arrLngInput(0) - PrevDiff
        SumValues = SumValues + PrevSequence
        
        ActiveCell.Offset(1, 0).Select
    Loop
    
    MsgBox "Sum of the extrapolated values is " & SumValues
End Sub

