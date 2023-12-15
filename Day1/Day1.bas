Attribute VB_Name = "Day1"
Sub Part1()
'
' Macro1 Macro
'

'
    Sum_Total = 0
    Range("A1").Select
    
    Do While Not IsEmpty(ActiveCell.Value)
        boolFound1 = False
        strtext = ActiveCell.Value
        For i = 1 To Len(strtext)
            chr1 = Mid(strtext, i, 1)
            If IsNumeric(chr1) Then
                If Not boolFound1 Then
                    strFirst = chr1
                    boolFound1 = True
                End If
                strLast = chr1
            End If
        Next i
        
        Sum_Total = Sum_Total + CInt(strFirst & strLast)
        ActiveCell.Offset(1, 0).Select
    
    Loop
    
    MsgBox "Sum of all the calibration values is " & Sum_Total
    
End Sub

Sub Part2()
'
' Macro1 Macro
'

'
    Sum_Total = 0
    Range("A1").Select
    
    Do While Not IsEmpty(ActiveCell.Value)
        boolFound1 = False
        strtext = ActiveCell.Value
        For i = 1 To Len(strtext)
            chr1 = Mid(strtext, i, 1)
            
            Select Case chr1
                Case "o"
                    If Mid(strtext, i, 3) = "one" Then chr1 = 1
                Case "t"
                    If Mid(strtext, i, 3) = "two" Then chr1 = 2
                    If Mid(strtext, i, 5) = "three" Then chr1 = 3
                Case "f"
                    If Mid(strtext, i, 4) = "four" Then chr1 = 4
                    If Mid(strtext, i, 4) = "five" Then chr1 = 5
                Case "s"
                    If Mid(strtext, i, 3) = "six" Then chr1 = 6
                    If Mid(strtext, i, 5) = "seven" Then chr1 = 7
                Case "e"
                    If Mid(strtext, i, 5) = "eight" Then chr1 = 8
                Case "n"
                    If Mid(strtext, i, 4) = "nine" Then chr1 = 9
            End Select
            
            If IsNumeric(chr1) Then
                If Not boolFound1 Then
                    strFirst = chr1
                    boolFound1 = True
                End If
                strLast = chr1
            End If
        Next i
        
        Sum_Total = Sum_Total + CInt(strFirst & strLast)
        ActiveCell.Offset(1, 0).Select
    
    Loop
    
    MsgBox "Sum of all the calibration values is " & Sum_Total
        
End Sub

