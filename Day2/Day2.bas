Attribute VB_Name = "Day2"
Sub Part1()
'
' Macro1 Macro
'

'
    Sum_Total = 0
    Range("A1").Select
    
    max_red = 12
    max_green = 13
    max_blue = 14
    
    Do While Not IsEmpty(ActiveCell.Value)
        invalid = False
        Do While Not IsEmpty(ActiveCell.Value)
            strtext = ActiveCell.Value
            Select Case UCase(strtext)
                Case "GAME"
                    GameID = ActiveCell.Offset(0, 1).Value
                    ActiveCell.Offset(0, 2).Select
                    
                Case "BLUE"
                    intBlue = ActiveCell.Offset(0, -1).Value
                    If intBlue > max_blue Then
                        invalid = True
                        Exit Do
                    End If
                    ActiveCell.Offset(0, 1).Select
                    
                Case "RED"
                    intRed = ActiveCell.Offset(0, -1).Value
                    If intRed > max_red Then
                        invalid = True
                        Exit Do
                    End If
                    ActiveCell.Offset(0, 1).Select
                    
                Case "GREEN"
                    intGreen = ActiveCell.Offset(0, -1).Value
                    If intGreen > max_green Then
                        invalid = True
                        Exit Do
                    End If
                    ActiveCell.Offset(0, 1).Select
                
                Case Else
                    ActiveCell.Offset(0, 1).Select
                    
            End Select
        Loop
        
        If Not invalid Then Sum_Total = Sum_Total + GameID
        
        Cells(ActiveCell.Row + 1, 1).Select
    Loop
    
    MsgBox "Sum of the valid Game IDs is " & Sum_Total
    
End Sub

Sub Part2()
'
' Macro1 Macro
'

'
    Sum_Total = 0
    Range("A1").Select
    
    Do While Not IsEmpty(ActiveCell.Value)
        min_red = 0
        min_green = 0
        min_blue = 0
    
        Do While Not IsEmpty(ActiveCell.Value)
            strtext = ActiveCell.Value
            Select Case UCase(strtext)
                Case "GAME"
                    GameID = ActiveCell.Offset(0, 1).Value
                    ActiveCell.Offset(0, 2).Select
                    
                Case "BLUE"
                    intBlue = ActiveCell.Offset(0, -1).Value
                    If intBlue > min_blue Then min_blue = intBlue
                    ActiveCell.Offset(0, 1).Select
                    
                Case "RED"
                    intRed = ActiveCell.Offset(0, -1).Value
                    If intRed > min_red Then min_red = intRed
                    ActiveCell.Offset(0, 1).Select
                    
                Case "GREEN"
                    intGreen = ActiveCell.Offset(0, -1).Value
                    If intGreen > min_green Then min_green = intGreen
                    ActiveCell.Offset(0, 1).Select
                
                Case Else
                    ActiveCell.Offset(0, 1).Select
                    
            End Select
        Loop
        
        Sum_Total = Sum_Total + min_blue * min_red * min_green
        
        Cells(ActiveCell.Row + 1, 1).Select
    Loop
    
    MsgBox "Sum of the power of the sets is " & Sum_Total
    
End Sub

