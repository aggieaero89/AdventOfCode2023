Attribute VB_Name = "Day11"
Sub Part1()
'
' Macro1 Macro
'
    ActiveSheet.UsedRange
    
    Range("A1").Select
    LastRow = ActiveCell.CurrentRegion.SpecialCells(xlCellTypeLastCell).Row - 3
    LastCol = ActiveCell.CurrentRegion.SpecialCells(xlCellTypeLastCell).Column - 3
    
    OutputRow = 2
    OutputCol = LastCol + 5
    
    For i = 1 To LastRow
        For j = 1 To LastCol
            If Cells(i, j).Value = "#" Then
                Cells(OutputRow, OutputCol).Value = OutputRow
                Cells(OutputRow, OutputCol + 1).Value = i
                Cells(OutputRow, OutputCol + 2).Value = j
                OutputRow = OutputRow + 1
            End If
        Next j
    Next i
    
    MsgBox "Done"
End Sub

Sub Part2()
'
' Macro1 Macro
'
    ActiveSheet.UsedRange
    
    Expansion = 1000000
    
    Range("A1").Select
    ExpansRow = ActiveCell.CurrentRegion.SpecialCells(xlCellTypeLastCell).Column
    ExpansCol = ExpansRow
    LastRow = ExpansCol - 4
    LastCol = ExpansRow - 4
    
    OutputRow = 1
    OutputCol = LastCol + 6
    
    For i = 1 To LastRow
        For j = 1 To LastCol
            If Cells(i, j).Value = "#" Then
                Cells(OutputRow, OutputCol).Value = OutputRow
                Cells(OutputRow, OutputCol + 1).Value = i
                Cells(OutputRow, OutputCol + 2).Value = j
                Cells(OutputRow, OutputCol + 3).Value = Cells(i, ExpansRow).Value * (Expansion - 1)
                Cells(OutputRow, OutputCol + 4).Value = Cells(ExpansCol, j).Value * (Expansion - 1)
                OutputRow = OutputRow + 1
            End If
        Next j
    Next i
    
    MsgBox "Done"
End Sub

