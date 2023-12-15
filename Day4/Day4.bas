Attribute VB_Name = "Day4"
Sub Part2()
'
' Macro1 Macro
'

'
    col_Source = ActiveCell.Column
    row_Source = ActiveCell.Row
    
    Do While Not IsEmpty(ActiveCell.Value)
        ActiveCell.Offset(0, 1).Value = ActiveCell.Offset(0, 1).Value + 1
        num_Wins = ActiveCell.Value
        If num_Wins > 0 Then
            num_Cards = ActiveCell.Offset(0, 1).Value
            For i = 1 To num_Wins
                ActiveCell.Offset(i, 1).Value = ActiveCell.Offset(i, 1).Value + num_Cards
            Next i
        End If
        
        Cells(ActiveCell.Row + 1, col_Source).Select
    Loop
End Sub

