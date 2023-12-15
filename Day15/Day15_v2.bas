Attribute VB_Name = "Day15_v2"
Sub Part1()
'
' Macro1 Macro
'
    ActiveSheet.UsedRange
    
    Range("A1").Select
    arrHASH = Split(ActiveCell.Value, ",")
    
    SumTotal = 0
    For Each Item In arrHASH
        CurrentValue = ComputeHashValue(Item)
        SumTotal = SumTotal + CurrentValue
    Next Item
    
    MsgBox "The sum of the results is " & SumTotal
End Sub

Function ComputeHashValue(ByVal strItem As String) As Integer

    CurrentValue = 0
    For i = 1 To Len(strItem)
        charItem = Mid(strItem, i, 1)
        CurrentValue = (17 * (CurrentValue + Asc(charItem))) Mod 256
    Next i

    ComputeHashValue = CurrentValue
    
End Function

Sub Part2()
'
' Macro1 Macro
'
    ActiveSheet.UsedRange
    
    'Dim dictLens As New Scripting.Dictionary
    Dim arrBox(0 To 255) As New Dictionary
    
    Range("A1").Select
    arrHASH = Split(ActiveCell.Value, ",")
    
    SumTotal = 0
    For Each Item In arrHASH
        Select Case Right(Item, 1)
            Case "-"
                strLens = Left(Item, Len(Item) - 1)
                intBox = ComputeHashValue(strLens)
                If arrBox(intBox).Exists(strLens) Then arrBox(intBox).Remove strLens
                
            Case Else
                arrLens = Split(Item, "=")
                intBox = ComputeHashValue(arrLens(0))
                If arrBox(intBox).Exists(arrLens(0)) Then
                    arrBox(intBox)(arrLens(0)) = CInt(arrLens(1))
                Else
                    arrBox(intBox).Add Key:=arrLens(0), Item:=CInt(arrLens(1))
                End If
                
        End Select
        
    Next Item
    
    
    SumTotal = 0
    k = 0
    For Each Item In arrBox
        k = k + 1
        If Not Item Is Nothing Then
            If Item.Count > 0 Then
                n = 0
                For Each Key In Item.Keys
                    n = n + 1
                    SumTotal = SumTotal + k * n * Item(Key)
                Next Key
            End If
        End If
    Next Item
    
    MsgBox "The focusing power of the resulting lens configuration is " & SumTotal
End Sub
