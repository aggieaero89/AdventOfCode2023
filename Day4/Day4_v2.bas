Attribute VB_Name = "Day4_v2"
Sub Part1()
'
' Macro1 Macro
'

'
    Range("A1").Select
    
    sumPoints = 0
    
    Do While Not IsEmpty(ActiveCell.Value)
        arrInput = Split(Replace(ActiveCell.Value, "|", ":"), ":")
    
        strWinNums = Trim(arrInput(1))
        strCrdNums = Trim(arrInput(2))
        
        strWinNums = RemoveExtraSpaces(strWinNums)
        strCrdNums = RemoveExtraSpaces(strCrdNums)
        
        sumWinNums = 0
        
        arrWinNums = Split(strWinNums, " ")
        arrCrdNums = Split(strCrdNums, " ")
        
        For Each i In arrCrdNums
            If InArray(i, arrWinNums) Then
                sumWinNums = sumWinNums + 1
            End If
        Next i
        
        If sumWinNums > 0 Then sumPoints = sumPoints + 2 ^ (sumWinNums - 1)
        
        ActiveCell.Offset(1, 0).Select
        
    Loop
    
    MsgBox ("Total number of points is " & sumPoints)
    
End Sub

Function RemoveExtraSpaces(ByVal A As String) As String
    Do
        strA = A
        A = Replace(A, "  ", " ")
    Loop Until strA = A
        
    RemoveExtraSpaces = A
End Function

Function InArray(ByVal strToBeFound As String, ByRef arr As Variant) As Boolean
    For i = LBound(arr) To UBound(arr)
        If arr(i) = strToBeFound Then
            InArray = True
            Exit Function
        End If
    Next i
    InArray = False

End Function

Sub Part2()
'-----------------------------------------------------
'Create an array of the number of winning numbers (number of copies) for each card

    intCards = Range("A1").CurrentRegion.Rows.Count
    Dim arrCopNums() As Integer
    ReDim arrCopNums(1 To intCards)
    
    Range("A1").Select
    For j = 1 To intCards
        arrInput = Split(Replace(ActiveCell.Value, "|", ":"), ":")
    
        strWinNums = Trim(arrInput(1))
        strCrdNums = Trim(arrInput(2))
        
        strWinNums = RemoveExtraSpaces(strWinNums)
        strCrdNums = RemoveExtraSpaces(strCrdNums)
        
        sumWinNums = 0
        
        arrWinNums = Split(strWinNums, " ")
        arrCrdNums = Split(strCrdNums, " ")
        
        For Each i In arrCrdNums
            If InArray(i, arrWinNums) Then
                sumWinNums = sumWinNums + 1
            End If
        Next i
        
        arrCopNums(j) = sumWinNums
        
        ActiveCell.Offset(1, 0).Select
        
    Next j 'Card num
    
'-----------------------------------------------------
'Create an array of the total number of original and copies for each card

    Dim arrTotNums() As Long
    ReDim arrTotNums(1 To intCards)
        
    For j = 1 To intCards
        arrTotNums(j) = arrTotNums(j) + 1
        num_Wins = arrCopNums(j)
        If num_Wins > 0 Then
            num_Cards = arrTotNums(j)
            For i = 1 To num_Wins
                arrTotNums(j + i) = arrTotNums(j + i) + num_Cards
            Next i
        End If
        
    Next j 'Card num
    
'-----------------------------------------------------
'Sum up the total number of original and copies for each card

    sumCrdNums = 0
    For j = 1 To intCards
        sumCrdNums = sumCrdNums + arrTotNums(j)
    Next j 'Card num
    MsgBox ("Total number of cards is " & sumCrdNums)
    
End Sub

