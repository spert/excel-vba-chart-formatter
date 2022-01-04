
Option Explicit

Public Function ConvertTo2DArray(DoubleInput As Variant) As Variant

  Dim i As Integer, j As Integer
  Dim arrResult() As Variant
  ReDim Preserve arrResult(1 To UBound(DoubleInput) - LBound(DoubleInput) + 1, 1 To 1)
  j = 1
  
  For i = LBound(DoubleInput) To UBound(DoubleInput)
  
    arrResult(j, 1) = DoubleInput(i)
    j = j + 1
    
  Next

  ConvertTo2DArray = arrResult

End Function


Public Sub FormatRange(MyRange As Range)

    Dim rF As Range
    
    For Each rF In MyRange
    
        rF.Borders.color = RGB(255, 255, 255)
        rF.ColumnWidth = 21.45
    
        With rF.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .color = RGB(217, 225, 242)
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    
    Next

End Sub

'Public Function GetSeriesColor(index As Integer, ColOrWei) As Long
'
'     Dim color As Long
'     Dim Weight As Long
'
'     Select Case index
'        Case 1
'          color = RGB(223, 158, 48)
'          Weight = 2
'        Case 2
'          color = RGB(190, 82, 3)
'          Weight = 2
'        Case 3
'          color = RGB(44, 68, 94)
'          Weight = 2
'        Case 4
'          color = RGB(46, 108, 129)
'          Weight = 2
'        Case 5
'          color = RGB(100, 178, 199)
'          Weight = 2
'        Case Else
'          color = RGB(100, 178, 199)
'          Weight = 2
'    End Select
'
'     Select Case UCase(ColOrWei)
'        Case "COLOR"
'            GetSeriesColor = color
'        Case "WEIGHT"
'            GetSeriesColor = Weight
'    End Select
'
'End Function

Public Function GetChartObject(ob As Object) As ChartObject

    Dim oSelection As Object
    Set oSelection = ob

    Do While TypeName(oSelection) <> "Application"
    
        If TypeName(oSelection) = "ChartObject" Then
           
           Set GetChartObject = oSelection
           Exit Function
           
        Else
        
            Set oSelection = oSelection.Parent
            
        End If
    
    Loop
    
    Set GetChartObject = Nothing

End Function


Function SERIESFUNC(Optional n, Optional cat, Optional Vals, Optional order, Optional BubSize) As Variant
      
   'Code source in this module is from: Excel 2010 Power Programming with VBA By John Walkenbach, page 640
    Dim i As Integer
    Dim result(1 To 5, 1 To 6) As Variant
    result(1, 1) = "Empty"
    result(2, 1) = "Empty"
    result(3, 1) = "Empty"
    result(4, 1) = "Empty"
    result(5, 1) = "Empty"
'    Result(1, 2) = ""
'    Result(2, 2) = ""
'    Result(3, 2) = ""
'    Result(4, 2) = ""
'    Result(5, 2) = ""
    
    
    If IsMissing(n) Then
    
        result(1, 1) = "String"
        result(1, 2) = ""
        
    ElseIf TypeName(n) = "Range" Then
        
        result(1, 1) = "Range_single"
        result(1, 2) = n.Areas(1).Address(, , , True)
        
    ElseIf TypeName(n) = "String" Then
    
        result(1, 1) = "String"
        result(1, 2) = n
        'Debug.Print ("string")
        
    ElseIf TypeName(n) = "Error" Then 'Error = Closed external workbook
            
        result(1, 1) = "Closed_external_workbook"
        result(1, 2) = ""
    
    End If
    
    
    If IsMissing(cat) Then
    
        result(2, 1) = "Empty"
        
    ElseIf TypeName(cat) = "Range" Then

        If cat.Areas.count = 1 Then

            result(2, 1) = "Range_single"
            result(2, 2) = cat.Areas(1).Address(, , , True)
            
        Else
        
            result(2, 1) = "Range_multiple"
            For i = 1 To cat.Areas.count
                result(2, 2) = result(2, 2) & cat.Areas(i).Address(, , , True)
               If i <> cat.Areas.count Then result(2, 2) = result(2, 2) & "&&"
            Next i
        
        End If
                
    ElseIf TypeName(cat) = "Variant()" Then
    
            Dim bString As Boolean
            bString = False
        
            result(2, 1) = "Variant()"
            
            For i = LBound(cat) To UBound(cat)
            
                result(2, 2) = result(2, 2) & cat(i)
                
                If i <> UBound(cat) Then result(2, 2) = result(2, 2) & "&&"
                
            Next i
            
    ElseIf TypeName(cat) = "String" Then
    
        result(2, 1) = "String"
        result(2, 2) = ""
        
    ElseIf TypeName(cat) = "Error" Then
        result(2, 1) = "Closed_external_workbook"
        result(2, 2) = ""
    End If
    
'    Debug.Print (TypeName(Vals))
    
    If IsMissing(Vals) Then
        result(3, 1) = "Empty"
    ElseIf TypeName(Vals) = "Range" Then
    
       If Vals.Areas.count = 1 Then
          
            result(3, 1) = "Range_single"
            result(3, 2) = Vals.Areas(1).Address(, , , True)
       
       Else
    
            result(3, 1) = "Range_multiple"
            For i = 1 To Vals.Areas.count
                result(3, 2) = result(3, 2) & Vals.Areas(i).Address(, , , True)
                If i <> Vals.Areas.count Then result(3, 2) = result(3, 2) & "&&"
            Next i
       
       End If
        
    ElseIf TypeName(Vals) = "Variant()" Then

        result(3, 1) = "Variant()"
        'Result(3, 2) = Result(3, 2) & "{"
        
        For i = LBound(Vals) To UBound(Vals)
            result(3, 2) = result(3, 2) & Vals(i)
            If i <> UBound(Vals) Then result(3, 2) = result(3, 2) & "&&"
        Next i
        
        'Result(3, 2) = Result(3, 2) & "}"
    
    ElseIf TypeName(Vals) = "String" Then
    
        result(3, 1) = "String"
        result(3, 2) = ""
        
    ElseIf TypeName(Vals) = "Error" Then
    
        result(3, 1) = "Closed_external_workbook"
        result(3, 2) = ""
        
    End If

    If IsMissing(order) Then
    
        result(4, 1) = "Empty"
        
    ElseIf TypeName(order) = "Double" Then
    
        result(4, 1) = "Double"
        result(4, 2) = order
        
    End If
    
    SERIESFUNC = result

End Function


'Public Function CalculateNewDateSpan(DateAxis As MyNewDateScale) As MyNewDateScale
'
'    'TODO: rescaling not implemented jet
'
'
'End Function

