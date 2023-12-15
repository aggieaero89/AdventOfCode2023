Attribute VB_Name = "Day5"
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
    
    arrFound = arrSeeds
    numSeeds = UBound(arrSeeds) + 1 'zero base
    
    For i = 0 To UBound(arrSeeds)
        arrFound(i) = False
    Next i

    'Soil
    ActiveCell.Offset(3, 0).Select
    Items = ActiveCell.CurrentRegion.Rows.Count - 1
    For i = 1 To Items
        arrInput = Split(ActiveCell.Value, " ")
        lngSoil = CDbl(arrInput(0))
        lngMin = CDbl(arrInput(1))
        lngMax = lngMin + CDbl(arrInput(2)) - 1
        knt = 0
        For Each j In arrSeeds
            lngSeed = CDbl(j)
            If Not arrFound(knt) Then
                If (lngSeed < lngMin) Or (lngSeed > lngMax) Then
                    arrFound(knt) = False
                Else
                    arrFound(knt) = True
                    arrSoil(knt) = lngSoil + (lngSeed - lngMin)
                End If
            End If
            knt = knt + 1
        Next j
        ActiveCell.Offset(1, 0).Select
    Next i
    
    knt = 0
    For Each j In arrSeeds
        lngSeed = CDbl(j)
        If Not arrFound(knt) Then
            arrSoil(knt) = lngSeed
        End If
        knt = knt + 1
    Next j
    
    
    For i = 0 To UBound(arrSeeds)
        arrFound(i) = False
    Next i

    'Fertilizer
    ActiveCell.Offset(2, 0).Select
    Items = ActiveCell.CurrentRegion.Rows.Count - 1
    For i = 1 To Items
        arrInput = Split(ActiveCell.Value, " ")
        lngFertilizer = CDbl(arrInput(0))
        lngMin = CDbl(arrInput(1))
        lngMax = lngMin + CDbl(arrInput(2)) - 1
        knt = 0
        For Each j In arrSoil
            lngSoil = CDbl(j)
            If Not arrFound(knt) Then
                If (lngSoil < lngMin) Or (lngSoil > lngMax) Then
                    arrFound(knt) = False
                Else
                    arrFound(knt) = True
                    arrFertilizer(knt) = lngFertilizer + (lngSoil - lngMin)
                End If
            End If
            knt = knt + 1
        Next j
        ActiveCell.Offset(1, 0).Select
    Next i
    
    knt = 0
    For Each j In arrSoil
        lngSoil = CDbl(j)
        If Not arrFound(knt) Then
            arrFertilizer(knt) = lngSoil
        End If
        knt = knt + 1
    Next j
    
    
    For i = 0 To UBound(arrSeeds)
        arrFound(i) = False
    Next i

    'Water
    ActiveCell.Offset(2, 0).Select
    Items = ActiveCell.CurrentRegion.Rows.Count - 1
    For i = 1 To Items
        arrInput = Split(ActiveCell.Value, " ")
        lngWater = CDbl(arrInput(0))
        lngMin = CDbl(arrInput(1))
        lngMax = lngMin + CDbl(arrInput(2)) - 1
        knt = 0
        For Each j In arrFertilizer
            lngFertilizer = CDbl(j)
            If Not arrFound(knt) Then
                If (lngFertilizer < lngMin) Or (lngFertilizer > lngMax) Then
                    arrFound(knt) = False
                Else
                    arrFound(knt) = True
                    arrWater(knt) = lngWater + (lngFertilizer - lngMin)
                End If
            End If
            knt = knt + 1
        Next j
        ActiveCell.Offset(1, 0).Select
    Next i
    
    knt = 0
    For Each j In arrFertilizer
        lngFertilizer = CDbl(j)
        If Not arrFound(knt) Then
            arrWater(knt) = lngFertilizer
        End If
        knt = knt + 1
    Next j
    
    
    For i = 0 To UBound(arrSeeds)
        arrFound(i) = False
    Next i

    'Light
    ActiveCell.Offset(2, 0).Select
    Items = ActiveCell.CurrentRegion.Rows.Count - 1
    For i = 1 To Items
        arrInput = Split(ActiveCell.Value, " ")
        lngLight = CDbl(arrInput(0))
        lngMin = CDbl(arrInput(1))
        lngMax = lngMin + CDbl(arrInput(2)) - 1
        knt = 0
        For Each j In arrWater
            lngWater = CDbl(j)
            If Not arrFound(knt) Then
                If (lngWater < lngMin) Or (lngWater > lngMax) Then
                    arrFound(knt) = False
                Else
                    arrFound(knt) = True
                    arrLight(knt) = lngLight + (lngWater - lngMin)
                End If
            End If
            knt = knt + 1
        Next j
        ActiveCell.Offset(1, 0).Select
    Next i
    
    knt = 0
    For Each j In arrWater
        lngWater = CDbl(j)
        If Not arrFound(knt) Then
            arrLight(knt) = lngWater
        End If
        knt = knt + 1
    Next j
    
    
    For i = 0 To UBound(arrSeeds)
        arrFound(i) = False
    Next i

    'Temperature
    ActiveCell.Offset(2, 0).Select
    Items = ActiveCell.CurrentRegion.Rows.Count - 1
    For i = 1 To Items
        arrInput = Split(ActiveCell.Value, " ")
        lngTemperature = CDbl(arrInput(0))
        lngMin = CDbl(arrInput(1))
        lngMax = lngMin + CDbl(arrInput(2)) - 1
        knt = 0
        For Each j In arrLight
            lngLight = CDbl(j)
            If Not arrFound(knt) Then
                If (lngLight < lngMin) Or (lngLight > lngMax) Then
                    arrFound(knt) = False
                Else
                    arrFound(knt) = True
                    arrTemperature(knt) = lngTemperature + (lngLight - lngMin)
                End If
            End If
            knt = knt + 1
        Next j
        ActiveCell.Offset(1, 0).Select
    Next i
    
    knt = 0
    For Each j In arrLight
        lngLight = CDbl(j)
        If Not arrFound(knt) Then
            arrTemperature(knt) = lngLight
        End If
        knt = knt + 1
    Next j
    
    
    For i = 0 To UBound(arrSeeds)
        arrFound(i) = False
    Next i

    'Humidity
    ActiveCell.Offset(2, 0).Select
    Items = ActiveCell.CurrentRegion.Rows.Count - 1
    For i = 1 To Items
        arrInput = Split(ActiveCell.Value, " ")
        lngHumidity = CDbl(arrInput(0))
        lngMin = CDbl(arrInput(1))
        lngMax = lngMin + CDbl(arrInput(2)) - 1
        knt = 0
        For Each j In arrTemperature
            lngTemperature = CDbl(j)
            If Not arrFound(knt) Then
                If (lngTemperature < lngMin) Or (lngTemperature > lngMax) Then
                    arrFound(knt) = False
                Else
                    arrFound(knt) = True
                    arrHumidity(knt) = lngHumidity + (lngTemperature - lngMin)
                End If
            End If
            knt = knt + 1
        Next j
        ActiveCell.Offset(1, 0).Select
    Next i
    
    knt = 0
    For Each j In arrTemperature
        lngTemperature = CDbl(j)
        If Not arrFound(knt) Then
            arrHumidity(knt) = lngTemperature
        End If
        knt = knt + 1
    Next j
    
    
    For i = 0 To UBound(arrSeeds)
        arrFound(i) = False
    Next i

    'Location
    ActiveCell.Offset(2, 0).Select
    Items = ActiveCell.CurrentRegion.Rows.Count - 1
    For i = 1 To Items
        arrInput = Split(ActiveCell.Value, " ")
        lngLocation = CDbl(arrInput(0))
        lngMin = CDbl(arrInput(1))
        lngMax = lngMin + CDbl(arrInput(2)) - 1
        knt = 0
        For Each j In arrHumidity
            lngHumidity = CDbl(j)
            If Not arrFound(knt) Then
                If (lngHumidity < lngMin) Or (lngHumidity > lngMax) Then
                    arrFound(knt) = False
                Else
                    arrFound(knt) = True
                    arrLocation(knt) = lngLocation + (lngHumidity - lngMin)
                End If
            End If
            knt = knt + 1
        Next j
        ActiveCell.Offset(1, 0).Select
    Next i
    
    knt = 0
    For Each j In arrHumidity
        lngHumidity = CDbl(j)
        If Not arrFound(knt) Then
            arrLocation(knt) = lngHumidity
        End If
        knt = knt + 1
    Next j
    
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
    Dim arrSeeds() As String
    ReDim arrSeeds(0 To upperIndex)
    
    Dim arrSoil() As String
    ReDim arrSoil(0 To upperIndex)
    
    Dim arrFertilizer() As String
    ReDim arrFertilizer(0 To upperIndex)
    
    Dim arrWater() As String
    ReDim arrWater(0 To upperIndex)
    
    Dim arrLight() As String
    ReDim arrLight(0 To upperIndex)
    
    Dim arrTemperature() As String
    ReDim arrTemperature(0 To upperIndex)
    
    Dim arrHumidity() As String
    ReDim arrHumidity(0 To upperIndex)
    
    Dim arrLocation() As String
    ReDim arrLocation(0 To upperIndex)
    
    Dim arrFound() As Boolean
    ReDim arrFound(0 To upperIndex)
    
    For i = 0 To upperIndex
        arrSeeds(i) = CStr(firstSeedVal + i)
    Next i
    
    'numSeeds = UBound(arrSeeds) + 1 'zero base
    
    For i = 0 To UBound(arrSeeds)
        arrFound(i) = False
    Next i

    'Soil
    Application.StatusBar = "Processing Soil ..."
    ActiveCell.Offset(3, 0).Select
    Items = ActiveCell.CurrentRegion.Rows.Count - 1
    NotFoundKnt = numSeeds
    For i = 1 To Items
        If NotFoundKnt > 0 Then
            arrInput = Split(ActiveCell.Value, " ")
            lngSoil = CDbl(arrInput(0))
            lngMin = CDbl(arrInput(1))
            lngMax = lngMin + CDbl(arrInput(2)) - 1
            knt = 0
            For Each j In arrSeeds
                lngSeed = CDbl(j)
                If Not arrFound(knt) Then
                    If (lngSeed < lngMin) Or (lngSeed > lngMax) Then
                        arrFound(knt) = False
                    Else
                        arrFound(knt) = True
                        arrSoil(knt) = lngSoil + (lngSeed - lngMin)
                        NotFoundKnt = NotFoundKnt - 1
                    End If
                End If
                knt = knt + 1
            Next j
        End If
        ActiveCell.Offset(1, 0).Select
    Next i
    
    knt = 0
    For Each j In arrSeeds
        lngSeed = CDbl(j)
        If Not arrFound(knt) Then
            arrSoil(knt) = lngSeed
        End If
        knt = knt + 1
    Next j
    
    
    For i = 0 To UBound(arrSeeds)
        arrFound(i) = False
    Next i

    'Fertilizer
    Application.StatusBar = "Processing Fertilizer ..."
    ActiveCell.Offset(2, 0).Select
    Items = ActiveCell.CurrentRegion.Rows.Count - 1
    NotFoundKnt = numSeeds
    For i = 1 To Items
        If NotFoundKnt > 0 Then
            arrInput = Split(ActiveCell.Value, " ")
            lngFertilizer = CDbl(arrInput(0))
            lngMin = CDbl(arrInput(1))
            lngMax = lngMin + CDbl(arrInput(2)) - 1
            knt = 0
            For Each j In arrSoil
                lngSoil = CDbl(j)
                If Not arrFound(knt) Then
                    If (lngSoil < lngMin) Or (lngSoil > lngMax) Then
                        arrFound(knt) = False
                    Else
                        arrFound(knt) = True
                        arrFertilizer(knt) = lngFertilizer + (lngSoil - lngMin)
                        NotFoundKnt = NotFoundKnt - 1
                    End If
                End If
                knt = knt + 1
            Next j
        End If
        ActiveCell.Offset(1, 0).Select
    Next i
    
    knt = 0
    For Each j In arrSoil
        lngSoil = CDbl(j)
        If Not arrFound(knt) Then
            arrFertilizer(knt) = lngSoil
        End If
        knt = knt + 1
    Next j
    
    
    For i = 0 To UBound(arrSeeds)
        arrFound(i) = False
    Next i

    'Water
    Application.StatusBar = "Processing Water ..."
    ActiveCell.Offset(2, 0).Select
    Items = ActiveCell.CurrentRegion.Rows.Count - 1
    NotFoundKnt = numSeeds
    For i = 1 To Items
        If NotFoundKnt > 0 Then
            arrInput = Split(ActiveCell.Value, " ")
            lngWater = CDbl(arrInput(0))
            lngMin = CDbl(arrInput(1))
            lngMax = lngMin + CDbl(arrInput(2)) - 1
            knt = 0
            For Each j In arrFertilizer
                lngFertilizer = CDbl(j)
                If Not arrFound(knt) Then
                    If (lngFertilizer < lngMin) Or (lngFertilizer > lngMax) Then
                        arrFound(knt) = False
                    Else
                        arrFound(knt) = True
                        arrWater(knt) = lngWater + (lngFertilizer - lngMin)
                        NotFoundKnt = NotFoundKnt - 1
                    End If
                End If
                knt = knt + 1
            Next j
        End If
        ActiveCell.Offset(1, 0).Select
    Next i
    
    knt = 0
    For Each j In arrFertilizer
        lngFertilizer = CDbl(j)
        If Not arrFound(knt) Then
            arrWater(knt) = lngFertilizer
        End If
        knt = knt + 1
    Next j
    
    
    For i = 0 To UBound(arrSeeds)
        arrFound(i) = False
    Next i

    'Light
    Application.StatusBar = "Processing Light ..."
    ActiveCell.Offset(2, 0).Select
    Items = ActiveCell.CurrentRegion.Rows.Count - 1
    NotFoundKnt = numSeeds
    For i = 1 To Items
        If NotFoundKnt > 0 Then
            arrInput = Split(ActiveCell.Value, " ")
            lngLight = CDbl(arrInput(0))
            lngMin = CDbl(arrInput(1))
            lngMax = lngMin + CDbl(arrInput(2)) - 1
            knt = 0
            For Each j In arrWater
                lngWater = CDbl(j)
                If Not arrFound(knt) Then
                    If (lngWater < lngMin) Or (lngWater > lngMax) Then
                        arrFound(knt) = False
                    Else
                        arrFound(knt) = True
                        arrLight(knt) = lngLight + (lngWater - lngMin)
                        NotFoundKnt = NotFoundKnt - 1
                    End If
                End If
                knt = knt + 1
            Next j
        End If
        ActiveCell.Offset(1, 0).Select
    Next i
    
    knt = 0
    For Each j In arrWater
        lngWater = CDbl(j)
        If Not arrFound(knt) Then
            arrLight(knt) = lngWater
        End If
        knt = knt + 1
    Next j
    
    
    For i = 0 To UBound(arrSeeds)
        arrFound(i) = False
    Next i

    'Temperature
    Application.StatusBar = "Processing Temperature ..."
    ActiveCell.Offset(2, 0).Select
    Items = ActiveCell.CurrentRegion.Rows.Count - 1
    NotFoundKnt = numSeeds
    For i = 1 To Items
        If NotFoundKnt > 0 Then
            arrInput = Split(ActiveCell.Value, " ")
            lngTemperature = CDbl(arrInput(0))
            lngMin = CDbl(arrInput(1))
            lngMax = lngMin + CDbl(arrInput(2)) - 1
            knt = 0
            For Each j In arrLight
                lngLight = CDbl(j)
                If Not arrFound(knt) Then
                    If (lngLight < lngMin) Or (lngLight > lngMax) Then
                        arrFound(knt) = False
                    Else
                        arrFound(knt) = True
                        arrTemperature(knt) = lngTemperature + (lngLight - lngMin)
                        NotFoundKnt = NotFoundKnt - 1
                    End If
                End If
                knt = knt + 1
            Next j
        End If
        ActiveCell.Offset(1, 0).Select
    Next i
    
    knt = 0
    For Each j In arrLight
        lngLight = CDbl(j)
        If Not arrFound(knt) Then
            arrTemperature(knt) = lngLight
        End If
        knt = knt + 1
    Next j
    
    
    For i = 0 To UBound(arrSeeds)
        arrFound(i) = False
    Next i

    'Humidity
    Application.StatusBar = "Processing Humidity ..."
    ActiveCell.Offset(2, 0).Select
    Items = ActiveCell.CurrentRegion.Rows.Count - 1
    NotFoundKnt = numSeeds
    For i = 1 To Items
        If NotFoundKnt > 0 Then
            arrInput = Split(ActiveCell.Value, " ")
            lngHumidity = CDbl(arrInput(0))
            lngMin = CDbl(arrInput(1))
            lngMax = lngMin + CDbl(arrInput(2)) - 1
            knt = 0
            For Each j In arrTemperature
                lngTemperature = CDbl(j)
                If Not arrFound(knt) Then
                    If (lngTemperature < lngMin) Or (lngTemperature > lngMax) Then
                        arrFound(knt) = False
                    Else
                        arrFound(knt) = True
                        arrHumidity(knt) = lngHumidity + (lngTemperature - lngMin)
                        NotFoundKnt = NotFoundKnt - 1
                    End If
                End If
                knt = knt + 1
            Next j
        End If
        ActiveCell.Offset(1, 0).Select
    Next i
    
    knt = 0
    For Each j In arrTemperature
        lngTemperature = CDbl(j)
        If Not arrFound(knt) Then
            arrHumidity(knt) = lngTemperature
        End If
        knt = knt + 1
    Next j
    
    
    For i = 0 To UBound(arrSeeds)
        arrFound(i) = False
    Next i

    'Location
    Application.StatusBar = "Processing Location ..."
    ActiveCell.Offset(2, 0).Select
    Items = ActiveCell.CurrentRegion.Rows.Count - 1
    NotFoundKnt = numSeeds
    For i = 1 To Items
        If NotFoundKnt > 0 Then
            arrInput = Split(ActiveCell.Value, " ")
            lngLocation = CDbl(arrInput(0))
            lngMin = CDbl(arrInput(1))
            lngMax = lngMin + CDbl(arrInput(2)) - 1
            knt = 0
            For Each j In arrHumidity
                lngHumidity = CDbl(j)
                If Not arrFound(knt) Then
                    If (lngHumidity < lngMin) Or (lngHumidity > lngMax) Then
                        arrFound(knt) = False
                    Else
                        arrFound(knt) = True
                        arrLocation(knt) = lngLocation + (lngHumidity - lngMin)
                        NotFoundKnt = NotFoundKnt - 1
                    End If
                End If
                knt = knt + 1
            Next j
        End If
        ActiveCell.Offset(1, 0).Select
    Next i
    
    knt = 0
    For Each j In arrHumidity
        lngHumidity = CDbl(j)
        If Not arrFound(knt) Then
            arrLocation(knt) = lngHumidity
        End If
        knt = knt + 1
    Next j
    
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



