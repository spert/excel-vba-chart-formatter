
Option Explicit

Const cModule = "frmChtFormat"

Dim oChtOb As ChartObject

Property Let SetChartObject(ChtObject As ChartObject)

    Set oChtOb = ChtObject

End Property

Private Sub cmdExecute_Click()

    Const cProc = "cmdExecute_Click"
    
    ErrorMod.ResetErrorCounter
    
    On Error GoTo ErrorHandler:
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    Dim aOptions(6) As Variant
    aOptions(0) = optSmall.Value
    aOptions(1) = optSlide.Value
    aOptions(2) = optLang1.Value
    aOptions(3) = optLang2.Value
    aOptions(4) = optTitleTop.Value
    aOptions(5) = chkTitleTextBox.Value
    
    Dim Settings As New clsSettings
    Call Settings.InitiateSettings(ActiveWorkBook, ActiveChart, aOptions)
    
    Dim props() As MyChart
    props = Settings.GetFormatProperties()
    
    Dim i As Integer
    For i = 0 To UBound(props)
        
        Call Execute(props(i))
    
    Next
    
    If ErrorMod.GetErrorCount > 0 Then
        
        Dim frmError As frmMessage
        Set frmError = New frmMessage
        frmError.SetErrorMessage = ErrorMod.GetErrorMessages
        frmError.Show
        
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
    
    End If
    
    Unload Me
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = False
    
    Exit Sub
    
ErrorHandler:
    
    MsgBox "Unexpected error in exec module", vbOKOnly
    
End Sub


Private Sub Execute(Properties As MyChart)

    Const cProc = "Execute"

   On Error GoTo ErrorHandler:

    ' -------- Initiate Chart Formatting --------------------
    Dim ChtFormat As New clsChart
    Call ChtFormat.InitiateChartFormat(oChtOb.Chart, Properties)
    Call ChtFormat.CollectSeries
    Call ChtFormat.CollectAxes
    Call ChtFormat.ValidateExcelChart
        
    If ErrorMod.GetErrorCount > 0 Then
        Exit Sub
    End If
    
    ' --------- Format Worksheet Layout ------------
    Dim NewWorkSheet As New clsWorkSheet
    Call NewWorkSheet.InitiateNewWorkSheet(ActiveWorkBook, Properties)
    Call NewWorkSheet.AddNewChartToWorkbook
    Call NewWorkSheet.TurnOffGridLines
    Call NewWorkSheet.PrintHeaders
        
    ' --------- Format Worksheet Layout ------------
    ChtFormat.SetNewWorksheet = NewWorkSheet
    Call ChtFormat.AssignNewRanges
    Call ChtFormat.CopyOldChartToNewWorksheet
    Call ChtFormat.MapNewSeries
    'Call ChtFormat.MapHeadingAndSource
        
    Call ChtFormat.PrintChartTitle
    Call ChtFormat.PrintLegend
    Call ChtFormat.PrintPlotArea
    Call ChtFormat.PrintSourceTextBox
    Call ChtFormat.PrintChartAxisTitle
    
    Call ChtFormat.ApplySeriesFormat
    'Call ChtFormat.ChartTextBoxes
    Call ChtFormat.ChartAxes

    ' --------- Format Axes Layout ------------
    If chkRescale.Value = True Then
         Call ChtFormat.Rescale
    End If
    
    'Call AxesFormat.InitiateAxesFormat(oChtOb.Chart, Properties)
    
            Exit Sub
    
ErrorHandler:
    
   ErrorMod.ErrorMessage cProc, cModule


End Sub

Private Sub optTitleNone_Click()

     chkTitleTextBox.Value = False
     chkTitleTextBox.Enabled = False


End Sub

Private Sub optTitleTop_Click()

  chkTitleTextBox.Value = True
  chkTitleTextBox.Enabled = True

End Sub

Private Sub UserForm_Click()

End Sub