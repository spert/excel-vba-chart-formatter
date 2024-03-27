
Option Explicit On

Private arrCat As Variant
Private tChartProps As MyChart
Private shtNewWorkSheet As Worksheet
Private iColumn As Integer
Private tMyNewDateScale As MyNewDateScale

Const cModule = "clsAxes"


Public Sub Initialize(CategoryValues As Variant, ChartProps As MyChart)

    arrCat = CategoryValues
    tChartProps = ChartProps

End Sub

Public Function IsSameAs(axis As clsAxes) As Boolean

    Const cProc = "IsSameAs"

    On Error GoTo ErrorHandler

    IsSameAs = False

    Dim arrMySource1 As Variant
    arrMySource1 = arrCat

    Dim arrMySource2 As Variant
    arrMySource2 = axis.GetCategoryArray()


    If StrComp(arrMySource1(1), "Empty", vbBinaryCompare) = 0 Or StrComp(arrMySource2(1), "Empty", vbBinaryCompare) = 0 Then

        IsSameAs = True
        Exit Function

    End If

    If StrComp(arrMySource1(1), arrMySource2(1), vbBinaryCompare) Then

        IsSameAs = False
        Exit Function

    End If

    If UBound(arrMySource1(3)) <> UBound(arrMySource2(3)) Then

        IsSameAs = False

        'Debug.Print ("Exit 3")
        Exit Function

    End If

    Dim h As Integer

    For h = 1 To UBound(arrMySource1(3))

        If arrMySource1(3)(h, 1) <> arrMySource2(3)(h, 1) Then

            IsSameAs = False

            'Debug.Print ("Exit 4")
            Exit Function

        End If

    Next

    'Debug.Print ("Exit 5")
    IsSameAs = True

    Exit Function

ErrorHandler:

    ErrorMod.ErrorMessage cProc, cModule

End Function

Public Function IsEmptyAxis() As Boolean

    IsEmptyAxis = False

    If arrCat(1) = "Empty" Then

        IsEmptyAxis = True

    End If

End Function

Property Get GetCategoryArray() As Variant

    GetCategoryArray = arrCat

End Property

Public Property Let SetColumn(Column As Integer)

        iColumn = Column
        
End Property


Property Get GetMyNewDateScale() As MyNewDateScale

    GetMyNewDateScale = tMyNewDateScale

End Property


Public Property Let SetNewWorkSheet(NewWorkSheet As Worksheet)

    Set shtNewWorkSheet = NewWorkSheet

End Property


Public Sub SetNewCategoryRange()

    Const cProc = "SetNewCategoryRange"

    On Error GoTo ErrorHandler

    Dim firstCell As Range
    Set firstCell = shtNewWorkSheet.Cells(tChartProps.SeriesDataOffset(1, 1), tChartProps.SeriesDataOffset(2, 1) + iColumn)
        
    If IsEmpty(arrCat(3)) Then
        Exit Sub
    End If

    Dim r As Range
    Set r = firstCell.Resize(UBound(arrCat(3)), 1)
            
    arrCat(5) = r.Address(, , , True)

    Exit Sub

ErrorHandler:

    ErrorMod.ErrorMessage cProc, cModule

End Sub


Public Sub PrintCategoryValues()

    Const cProc = "PrintCategoryValues"

    On Error GoTo ErrorHandler

    If IsEmpty(arrCat(5)) Then
        Exit Sub
    End If

    Dim r As Range
    Set r = Range(arrCat(5))
    
    Call FormatRange(r)

    r = arrCat(3)
    r.NumberFormat = arrCat(4)

    Exit Sub

ErrorHandler:

    ErrorMod.ErrorMessage cProc, cModule

End Sub


Public Sub ToString(count As Integer)

    Const cProc = "ToString()"

    On Error GoTo ErrorHandler

    Dim k As Integer

    Debug.Print(" ---------- Axis " & count & " -----------")
    Debug.Print(arrCat(1))
    Debug.Print(arrCat(2))
    For k = 1 To UBound(arrCat(3))
        Debug.Print(arrCat(3)(k, 1))
    Next
    Debug.Print(arrCat(4))
    Debug.Print(arrCat(5))

    Exit Sub

ErrorHandler:


End Sub

Public Sub Rescale()

    Const cProc = "RescaleDateAxis"

    tMyNewDateScale = CalculateNewScale(arrCat(3))

    Exit Sub

ErrorHandler:

    ErrorMod.ErrorMessage cProc, cModule


End Sub



