' Hello future keeper of the template. Yeah this code is kinda a mess but it works. If something needs changing, I'd reccomend just scrapping it and starting over.
' Seriously, I'm not sure how it works anymore lol. At some point we probably need a dedicated quoting software, how about you propose that?
' Anyway, good luck. -T

Public Function StormPrice(Data1 As Range, Data2 As Range, Data3 As Range) As Variant
    Dim price As Variant
    Dim structureType As Variant
    Dim lookup As Variant
    Dim height As Variant
    Dim roundheight As Variant
    Dim cut As Variant
    Dim structure As Variant
    
    
    structure = Data1.Cells(1, 6).Value
    height = Data3.Cells(1, 1).Value
    cut = Data3.Cells(1, 1).Value
    
    ' price = Application.WorksheetFunction.VLookup(Right(Left(Data1.Cells(1, 6), Application.WorksheetFunction.Search("'", Data1.Cells(1, 6)) + 14), 20), StormLookups, 2, False)
    If InStr(1, structure, "Trap", vbTextCompare) > 0 Then
        price = Application.WorksheetFunction.VLookup(Data1.Cells(1, 6).Value, Range("GreaseTrapLookups"), 2, False)
    ElseIf InStr(1, structure, "24", vbTextCompare) > 0 And InStr(1, structure, "Solid", vbTextCompare) > 0 Then
        price = Application.WorksheetFunction.VLookup(Data1.Cells(1, 6).Value, Range("TFPriceLookups"), 2, False)
    ElseIf InStr(1, structure, "Waffle", vbTextCompare) > 0 Then
        If InStr(1, structure, "Structure", vbTextCompare) > 0 Then
            If Left(Data1.Cells(1, 6).Value, 2) = "27" Then
                If height <= 5 Then
                    If height < 3 Then
                        roundheight = 3
                    Else
                        roundheight = Application.WorksheetFunction.Ceiling(height, 0.5)
                    End If
                
                    price = Application.WorksheetFunction.VLookup(roundheight, Range("TSWaffleBases"), 2, False)
                Else
                    price = Application.WorksheetFunction.VLookup(5, Range("TSWaffleBases"), 2, False) + Application.WorksheetFunction.VLookup(Application.WorksheetFunction.Ceiling(height - 5, 1), Range("Risers"), 2, False)
                End If
            ElseIf Left(Data1.Cells(1, 6).Value, 2) = "24" Then
                If height < 6 Then
                    If height < 3 Then
                        roundheight = 3
                    Else
                        roundheight = Application.WorksheetFunction.Ceiling(height, 0.5)
                    End If
                
                    price = Application.WorksheetFunction.VLookup(roundheight, Range("TFWaffleBases"), 2, False)
                Else
                    price = Application.WorksheetFunction.VLookup(6, Range("TFWaffleBases"), 2, False) + (((2 * (height - 6) * 0.5 * 0.037) + ((44 / 12) * (height - 6) * 0.5 * 0.037)) + ((3 * (height - 6) * 0.5 * 0.037) + (((44 / 12) + 1) * (height - 6) * 0.5 * 0.037))) * Application.WorksheetFunction.VLookup(3, Range("WaffleRiserLookup"), 2, False)
                End If
            ElseIf Left(Data1.Cells(1, 6).Value, 1) = "3" Then
                If height < 5 Then
                    If height < 3 Then
                        roundheight = 3
                    Else
                        roundheight = Application.WorksheetFunction.Ceiling(height, 0.5)
                    End If
                
                    price = Application.WorksheetFunction.VLookup(roundheight, Range("TWaffleBases"), 2, False) + Application.WorksheetFunction.VLookup(3, Range("Lids"), 2, False)
                Else
                    price = Application.WorksheetFunction.VLookup(5, Range("TWaffleBases"), 2, False) + Application.WorksheetFunction.VLookup(3, Range("Lids"), 2, False) + ((3 * (height - 5) * 0.5 * 0.037) * 2 + (4 * (height - 5) * 0.5 * 0.037) * 2) * Application.WorksheetFunction.VLookup(3, Range("WaffleRiserLookup"), 2, False)
                End If
            Else
                If height < 4.5 Then
                    If height < 4 Then
                        roundheight = 4
                    Else
                        roundheight = Application.WorksheetFunction.Ceiling(height, 0.5)
                    End If
                
                    price = Application.WorksheetFunction.VLookup(roundheight, Range("FWaffleBases"), 2, False) + Application.WorksheetFunction.VLookup(4, Range("Lids"), 2, False)
                Else
                    price = Application.WorksheetFunction.VLookup(4.5, Range("FWaffleBases"), 2, False) + Application.WorksheetFunction.VLookup(4, Range("Lids"), 2, False) + (((4 * (height - 5)) * 0.5 * 0.037) * 2 + ((5 * (height - 5)) * 0.5 * 0.037) * 2) * Application.WorksheetFunction.VLookup(4, Range("WaffleRiserLookup"), 2, False)
                End If
            End If
        Else
            price = Application.WorksheetFunction.VLookup(Data1.Cells(1, 6).Value, Range("WaffleBases"), 2, False)
        End If
    Else
    
    If Left(Data1.Cells(1, 6).Value, 1) = "D" Then
        structure = Right(Data1.Cells(1, 6).Value, Len(Data1.Cells(1, 6)) - Application.WorksheetFunction.Search(" ", Data1.Cells(1, 6)))
        lookup = Right(Left(structure, Application.WorksheetFunction.Search("'", structure) + 14), 20)
    Else
        lookup = Right(Left(Data1.Cells(1, 6).Value, Application.WorksheetFunction.Search("'", Data1.Cells(1, 6)) + 14), 20)
    End If
    structureType = Application.WorksheetFunction.VLookup(lookup, Range("TypeLookups"), 2, False)
    If structureType = "OP" Then
       price = cut * Application.WorksheetFunction.VLookup(lookup, Range("StormLookups"), 2, False) + Application.WorksheetFunction.VLookup(lookup, Range("StormLookups"), 3, False)
    ElseIf structureType = "B" Then
    
        If height >= 15 Then
            price = "USE ROUND or THICKER WALLS"
        ElseIf Left(lookup, 2) = "24" Then
            price = (((3 * ((44 / 12) + 1) * 0.5 * 0.037)) + (((2 * (height) * 0.5 * 0.037) + ((44 / 12) * (height) * 0.5 * 0.037)) + ((3 * (height) * 0.5 * 0.037) + (((44 / 12) + 1) * (height) * 0.5 * 0.037)))) * Application.WorksheetFunction.VLookup(3, Range("WaffleRiserLookup"), 2, False)
        Else
            price = (Application.WorksheetFunction.Sum(Application.WorksheetFunction.Product(Left(lookup, 1) + 1, Right(Left(lookup, 4), 1) + 1, (height - 0.5) + 1) - Application.WorksheetFunction.Product(Left(lookup, 1), Right(Left(lookup, 4), 1), height - 0.5)) / 27) * Application.WorksheetFunction.VLookup(lookup, Range("BoxLookups"), 2, False)
        End If
        

    ElseIf structureType = "TT" Then
        If height < 4 Then
            price = Range("LETH4")
        ElseIf height >= 9 Then
            price = "USE ROUND"
        Else
            roundheight = Application.WorksheetFunction.RoundUp(height + 0.01, 0)
        
            price = Range("LETH" & roundheight).Value
        End If
        
    ElseIf structureType = "SP" Then
        price = Application.WorksheetFunction.VLookup(Data1.Cells(1, 6).Value, Range("SPLookups"), 2, False)
    ElseIf structureType = "HW" Then
        price = Application.WorksheetFunction.VLookup(Left(Data1.Cells(1, 6).Value, 4), Range("HeadwallLookups"), 2, False)
    ElseIf structureType = "DHW" Then
        price = Application.WorksheetFunction.VLookup(Left(Data1.Cells(1, 6).Value, 4), Range("DoubleHeadwallLookups"), 3, False)
    ElseIf structureType = "NP" Then
        Dim leftoverHeight As Variant
        Dim vfPrice As Variant
        
        leftoverHeight = cut - 5
        
        If leftoverHeight <= 0 Then
            vfPrice = 0
        Else
            vfPrice = Application.WorksheetFunction.VLookup(lookup, Range("NPStormLookups"), 3, False) * leftoverHeight
        End If
        
        price = Application.WorksheetFunction.VLookup(lookup, Range("NPStormLookups"), 2, False) + vfPrice

        
    End If
    End If
    
    
    StormPrice = price
End Function


Public Function SewerPrice(Data1 As Range) As Variant
    Dim price As Variant
    Dim structureType As Variant
    Dim lookup As Variant
    Dim height As Variant
    Dim roundheight As Variant
    Dim cut As Variant
    Dim structure As Variant
    
    
    structure = Data1.Cells(1, 6).Value
    height = Data1.Cells(1, 1).Value
    cut = Data1.Cells(1, 4).Value
    
    ' price = Application.WorksheetFunction.VLookup(Right(Left(Data1.Cells(1, 6), Application.WorksheetFunction.Search("'", Data1.Cells(1, 6)) + 14), 20), StormLookups, 2, False)
    If InStr(1, structure, "Trap", vbTextCompare) > 0 Then
        price = Application.WorksheetFunction.VLookup(Data1.Cells(1, 6).Value, Range("GreaseTrapLookups"), 2, False)
    Else
    lookup = Right(Left(Data1.Cells(1, 6).Value, Application.WorksheetFunction.Search("'", Data1.Cells(1, 6)) + 14), 20)
    structureType = Application.WorksheetFunction.VLookup(lookup, Range("TypeLookups"), 2, False)
    If structureType = "OP" Then
       price = cut * Application.WorksheetFunction.VLookup(lookup, Range("SewerLookups"), 2, False) + Application.WorksheetFunction.VLookup(lookup, Range("SewerLookups"), 3, False)
    ElseIf structureType = "B" Then
        price = (Application.WorksheetFunction.Sum(Application.WorksheetFunction.Product(Left(lookup, 1) + 1, Right(Left(lookup, 4), 1) + 1, (height - 0.5) + 1) - Application.WorksheetFunction.Product(Left(lookup, 1), Right(Left(lookup, 4), 1), height - 0.5)) / 27) * Application.WorksheetFunction.VLookup(lookup, Range("BoxLookups"), 2, False)
    ElseIf structureType = "TT" Then
        If height < 4 Then
            price = Range("LETH4")
        End If
        
        If height >= 9 Then
            price = "USE ROUND"
        Else
            roundheight = Application.WorksheetFunction.RoundUp(height + 0.01, 0)
        
            price = Range("LETH" & roundheight).Value
        End If
        
    ElseIf structureType = "SP" Then
        price = Application.WorksheetFunction.VLookup(Data1.Cells(1, 6).Value, Range("SPLookups"), 2, False)
    ElseIf structureType = "HW" Then
        price = Application.WorksheetFunction.VLookup(Left(Data1.Cells(1, 6).Value, 4), Range("HeadwallLookups"), 2, False)
    ElseIf structureType = "NP" Then
        Dim leftoverHeight As Variant
        Dim vfPrice As Variant
        
        leftoverHeight = cut - 6
        
        If leftoverHeight <= 0 Then
            vfPrice = 0
        Else
            vfPrice = Application.WorksheetFunction.VLookup(lookup, Range("NPSewerLookups"), 3, False) * leftoverHeight
        End If
        
        price = Application.WorksheetFunction.VLookup(lookup, Range("NPSewerLookups"), 2, False) + vfPrice

        
    End If
    End If
    
    
    SewerPrice = price
End Function

Public Function StormWeight(Data1 As Range, Data2 As Range, Data3 As Range) As Variant
    Dim height As Variant
    Dim structure As Variant
    Dim weight As Variant
    Dim structureType As Variant
    
    structure = Data1.Cells(1, 6).Value
    height = Data3.Cells(1, 1).Value
    
    structureType = Application.WorksheetFunction.VLookup(structure, Range("WeightInfoLookups"), 2, False)
    
    If structureType = "N" Then
    
        Dim baseCY As Variant
        Dim wallCYVF As Variant
        Dim lidCY As Variant
        
        baseCY = Application.WorksheetFunction.VLookup(structure, Range("WeightInfoLookups"), 3, False)
        wallCYVF = Application.WorksheetFunction.VLookup(structure, Range("WeightInfoLookups"), 5, False)
        lidCY = Application.WorksheetFunction.VLookup(structure, Range("WeightInfoLookups"), 4, False)
        
        weight = (baseCY + lidCY + wallCYVF * height) * Range("WeightPerCY").Value
    
    ElseIf structureType = "L" Then
        weight = Application.WorksheetFunction.VLookup(structure, Range("WeightInfoLookups"), 3, False)
        
    End If
    
    StormWeight = weight
End Function
