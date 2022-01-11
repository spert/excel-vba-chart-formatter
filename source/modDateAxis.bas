Option Explicit

Public Function CalculateNewScale(arrDates As Variant) As MyNewDateScale

    Dim tAxis As MyNewDateScale
    Dim iCounter As Integer
    Dim eFreq As MyFrequency
    Dim iDiff As Integer
    Dim d1 As Date
    Dim d2 As Date
    
    iCounter = UBound(arrDates)
    
    eFreq = CalculateFrequency(arrDates)
    
    If eFreq = MyFrequency.None Then
        
        Exit Function
    
    End If
    
     
    d1 = CDbl(arrDates(1, 1))
    d2 = CDbl(arrDates(UBound(arrDates), 1))
     
    If VBA.DateDiff("d", d1, d2) <= 6 Then
       tAxis.Rescaled = True
       tAxis.MinDate = d1
       tAxis.MaxDate = d2
       tAxis.MajorUnit = 1
       tAxis.BaseUnit = XlTimeUnit.xlDays
       tAxis.FormatCode = "dd.mm.yyyy"
       CalculateNewScale = tAxis
       Exit Function
    End If
       
       If VBA.DateDiff("ww", d1, d2) <= 6 Then
            tAxis.Rescaled = True
            tAxis.MinDate = d1
            tAxis.MaxDate = d2
            tAxis.MajorUnit = 6
            tAxis.MajorUnitScale = XlTimeUnit.xlDays
            tAxis.BaseUnit = XlTimeUnit.xlDays
            tAxis.FormatCode = "dd.mm.yyyy"
            CalculateNewScale = tAxis
            Exit Function
       End If
       
       If VBA.DateDiff("m", d1, d2) <= 6 Then
            tAxis.Rescaled = True
            tAxis.MinDate = d1
            tAxis.MaxDate = WorksheetFunction.EoMonth(d2, 0)
            tAxis.MajorUnit = 1
            tAxis.MajorUnitScale = XlTimeUnit.xlMonths
            tAxis.BaseUnit = XlTimeUnit.xlMonths
            tAxis.FormatCode = "mm.yyyy"
            CalculateNewScale = tAxis
            Exit Function
       End If
       
        If VBA.DateDiff("m", d1, d2) <= 12 Then
            tAxis.Rescaled = True
            tAxis.MinDate = arrDates(1, 1)
            tAxis.MaxDate = WorksheetFunction.EoMonth(d2, 0)
            tAxis.MajorUnit = 3
            tAxis.MajorUnitScale = XlTimeUnit.xlMonths
            tAxis.BaseUnit = XlTimeUnit.xlMonths
            tAxis.FormatCode = "mm.yyyy"
            CalculateNewScale = tAxis
            Exit Function
       End If
       
        If VBA.DateDiff("m", d1, d2) <= 24 Then
            tAxis.Rescaled = True
            tAxis.MinDate = arrDates(1, 1)
            tAxis.MaxDate = WorksheetFunction.EoMonth(d2, 0)
            tAxis.MajorUnit = 4
            tAxis.MajorUnitScale = XlTimeUnit.xlMonths
            tAxis.BaseUnit = XlTimeUnit.xlMonths
            tAxis.FormatCode = "mm.yyyy"
            CalculateNewScale = tAxis
            Exit Function
       End If
       
        If VBA.DateDiff("y", d1, d2) <= 6 Then
            tAxis.Rescaled = True
            tAxis.MinDate = arrDates(1, 1)
            tAxis.MaxDate = WorksheetFunction.EoMonth(d2, 12 - Month(d2))
            tAxis.MajorUnit = 1
            tAxis.MajorUnitScale = XlTimeUnit.xlYears
            tAxis.BaseUnit = XlTimeUnit.xlDays
            tAxis.FormatCode = "yyyy"
            CalculateNewScale = tAxis
            Exit Function
       End If
       
        If VBA.DateDiff("y", d1, d2) > 6 Then
            tAxis.Rescaled = True
            tAxis.MinDate = arrDates(1, 1)
            tAxis.MaxDate = WorksheetFunction.EoMonth(d2, 12 - Month(d2))
            iDiff = Year(d2) - Year(d1)
            tAxis.MajorUnit = Application.WorksheetFunction.Ceiling_Math(iDiff / 6)
            tAxis.MajorUnitScale = XlTimeUnit.xlYears
            tAxis.BaseUnit = XlTimeUnit.xlDays
            tAxis.FormatCode = "yyyy"
            CalculateNewScale = tAxis
            Exit Function
       End If
              

End Function


Public Function CalculateFrequency(arrDates As Variant) As MyFrequency

    Dim d1 As Date
    Dim d2 As Date
    Dim tFreq As MyFrequency

    If UBound(arrDates) < 2 Then
    
        CalculateFrequency = MyFrequency.None
        
        Return
        
    End If

    d1 = CDbl(arrDates(1, 1))

    d2 = CDbl(arrDates(2, 1))

    CalculateFrequency = GetFrequencyFromTowDates(d1, d2)


End Function


Public Function GetFrequencyFromTowDates(d1 As Date, d2 As Date) As MyFrequency
        
    If VBA.DateDiff("d", d1, d2) = 1 Then
    
       GetFrequencyFromTowDates = MyFrequency.Daily
       
       Exit Function
       
    End If

    If VBA.DateDiff("ww", d1, d2) = 1 Then
    
       GetFrequencyFromTowDates = MyFrequency.Weekly
    
       Exit Function
    
    End If
    
    If VBA.DateDiff("m", d1, d2) = 1 Then
    
       GetFrequencyFromTowDates = MyFrequency.Monthly
    
       Exit Function
    
    End If
    
    If VBA.DateDiff("q", d1, d2) = 1 Then
    
       GetFrequencyFromTowDates = MyFrequency.Quarterly
    
       Exit Function
    
    End If
        
    If VBA.DateDiff("yyyy", d1, d2) = 1 Then
    
       GetFrequencyFromTowDates = MyFrequency.Yearly
    
       Exit Function
    
    End If
        
    GetFrequencyFromTowDates = MyFrequency.None
    

End Function

