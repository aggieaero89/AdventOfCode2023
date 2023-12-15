Attribute VB_Name = "Day5_v2"
Sub Part1()
'
' Macro1 Macro
'

'
    Range("A1").Select
    
    'Seeds
    arrInput = Split(ActiveCell.Value, ":")
    strSeeds = Trim(arrInput(1))
    arrSeeds = Split(strSeeds, " ")
    
    arrSoil = arrSeeds
    arrFertilizer = arrSeeds
    arrWater = arrSeeds
    arrLight = arrSeeds
    arrTemperature = arrSeeds
    arrHumidity = arrSeeds
    arrLocation = arrSeeds
    
    ActiveCell.Offset(1, 0).Select
    
    Call MapIt(arrSeeds, arrSoil, "Soil")
    Call MapIt(arrSoil, arrFertilizer, "Fertilizer")
    Call MapIt(arrFertilizer, arrWater, "Water")
    Call MapIt(arrWater, arrLight, "Light")
    Call MapIt(arrLight, arrTemperature, "Temperature")
    Call MapIt(arrTemperature, arrHumidity, "Humidity")
    Call MapIt(arrHumidity, arrLocation, "Location")
        
    boolFirst = True
    For Each j In arrLocation
        dblSeed = CDbl(j)
        If boolFirst Then
            minLoc = dblSeed
            boolFirst = False
        Else
            If dblSeed < minLoc Then minLoc = dblSeed
        End If
    Next j
    MsgBox "Min location is " & minLoc
End Sub

Sub MapIt(ByRef arrSource, ByRef arrDestination, ByVal strStatus As String)

    Application.StatusBar = "Processing " & strStatus & " ..."
    
    First = LBound(arrSource)
    Last = UBound(arrSource)
    
    Dim arrFound() As Boolean
    ReDim arrFound(First To Last)
    
    For i = First To Last
        arrFound(i) = False
    Next i
    
    ActiveCell.Offset(2, 0).Select
    Items = ActiveCell.CurrentRegion.Rows.Count - 1
    NotFoundKnt = Last - First + 1
    For i = 1 To Items
        If NotFoundKnt > 0 Then
            arrInput = Split(ActiveCell.Value, " ")
            lngDestination = CDbl(arrInput(0))
            lngSourceMin = CDbl(arrInput(1))
            lngSourceMax = lngSourceMin + CDbl(arrInput(2)) - 1
            knt = 0
            For Each j In arrSource
                lngSource = CDbl(j)
                If Not arrFound(knt) Then
                    If (lngSource < lngSourceMin) Or (lngSource > lngSourceMax) Then
                        arrFound(knt) = False
                    Else
                        arrFound(knt) = True
                        arrDestination(knt) = lngDestination + (lngSource - lngSourceMin)
                        NotFoundKnt = NotFoundKnt - 1
                    End If
                End If
                knt = knt + 1
            Next j
        End If
        ActiveCell.Offset(1, 0).Select
    Next i
    
    knt = 0
    For Each j In arrSource
        If Not arrFound(knt) Then
            arrDestination(knt) = CDbl(j)
        End If
        knt = knt + 1
    Next j

    Application.StatusBar = False

End Sub

Sub Part2()
'
' Macro1 Macro
'

'
    Range("A1").Select
    
    'Seeds
    Application.StatusBar = "Processing Seeds ..."
    arrInput = Split(ActiveCell.Value, ":")
    strSeeds = Trim(arrInput(1))
    arrGroup = Split(strSeeds, " ")
    firstSeedVal = CDbl(arrGroup(0))
    numSeeds = CDbl(arrGroup(1))
    upperIndex = numSeeds - 1
    
    Dim arrSeeds() As Double
    ReDim arrSeeds(0 To upperIndex)
    
    Dim arrSoil() As Double
    ReDim arrSoil(0 To upperIndex)
    
    Dim arrFertilizer() As Double
    ReDim arrFertilizer(0 To upperIndex)
    
    Dim arrWater() As Double
    ReDim arrWater(0 To upperIndex)
    
    Dim arrLight() As Double
    ReDim arrLight(0 To upperIndex)
    
    Dim arrTemperature() As Double
    ReDim arrTemperature(0 To upperIndex)
    
    Dim arrHumidity() As Double
    ReDim arrHumidity(0 To upperIndex)
    
    Dim arrLocation() As Double
    ReDim arrLocation(0 To upperIndex)
    
    Dim arrFound() As Boolean
    ReDim arrFound(0 To upperIndex)
    
    For i = 0 To upperIndex
        arrSeeds(i) = CStr(firstSeedVal + i)
    Next i
    
    ActiveCell.Offset(1, 0).Select
    
    Call MapIt(arrSeeds, arrSoil, "Soil")
    Call MapIt(arrSoil, arrFertilizer, "Fertilizer")
    Call MapIt(arrFertilizer, arrWater, "Water")
    Call MapIt(arrWater, arrLight, "Light")
    Call MapIt(arrLight, arrTemperature, "Temperature")
    Call MapIt(arrTemperature, arrHumidity, "Humidity")
    Call MapIt(arrHumidity, arrLocation, "Location")
    
    Application.StatusBar = "Finding Location min ..."
    boolFirst = True
    For Each j In arrLocation
        dblSeed = CDbl(j)
        If boolFirst Then
            minLoc = dblSeed
            boolFirst = False
        Else
            If dblSeed < minLoc Then minLoc = dblSeed
        End If
    Next j
    MsgBox "Min location is " & minLoc
    
    Application.StatusBar = False
    
End Sub



