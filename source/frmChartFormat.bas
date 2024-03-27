
Option Explicit On

Const cModule = "frmChtFormat"

Private Settings As New clsSettings
Private ChtOld As New clsChartOld
Private ChtNew As New clsChartNew
Private ChtModify As New clsChartModify
Private ChtFormat As New clsChartFormat
Private NewWorkSheet As New clsWorkSheet
Private arrMySeries() As clsSeries

Dim oChtOb As ChartObject

Property Let SetChartObject(ChtObject As ChartObject)

    Set oChtOb = ChtObject

End Property


Private Sub cmdExecute_Click()

    Const cProc = "cmdExecute_Click"

    ErrorMod.ResetErrorCounter

    ' On Error GoTo ErrorHandler:

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    Dim aOptions(7) As Variant
    aOptions(0) = optSmall.Value
    aOptions(1) = optSlide.Value
    aOptions(2) = optLang1.Value
    aOptions(3) = optLang2.Value
    aOptions(4) = optTitleTop.Value
    aOptions(5) = chkTitleTextBox.Value
    aOptions(6) = chkSourceBox.Value

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

    If chkRecentColors.Value = True Then
        Call ApplyRecentColors(props(0))
    End If

    Application.DisplayAlerts = True
    Application.ScreenUpdating = False

    Exit Sub

ErrorHandler:

    MsgBox "Unexpected error in exec module", vbOKOnly

End Sub


Private Sub Execute(Properties As MyChart)

    Const cProc = "Execute"

    On Error GoTo ErrorHandler

    ' -------- Initiate Chart Formatting --------------------
    Call ChtOld.Initiate(oChtOb.Chart, Properties)
    Call ChtOld.ValidateExcelChart

    If ErrorMod.GetErrorCount > 0 Then
        Exit Sub
    End If

    Call ChtOld.CollectSeries
    Call ChtOld.CollectAxes

    If chkCopy.Value = True Then

        ' --------- Create Worksheet Layout ------------
        Call NewWorkSheet.Initiate(ActiveWorkBook, Properties)
        Call NewWorkSheet.AddNewChartToWorkbook
        Call NewWorkSheet.TurnOffGridLines
        Call NewWorkSheet.PrintHeaders

        ' --------- Format Worksheet Layout ------------
        Call ChtNew.Initiate(oChtOb.Chart, Properties)
        ChtNew.SetNewWorkSheet = NewWorkSheet.GetNewWorkSheet
        ChtNew.SetMyAxes = ChtOld.GetMyAxes
        ChtNew.SetMySeries = ChtOld.GetMySeries

        'order of calling subroutines matters
        Call ChtNew.AssignNewRanges
        Call ChtNew.CopyOldChartToNewWorksheet
        Call ChtNew.MapNewSeries
        Call ModifyChart(ChtNew.GetNewChart, Properties)
        Call ChtNew.PrintChartTitle
        Call ChtNew.PrintSourceTextBox
        Call ChtNew.PrintChartNumber
        Call ChtNew.PrintChartAxisTitle
        Call FormatChart(ChtNew.GetNewChart, Properties)

    Else

        Call ModifyChart(ChtOld.GetOldChart, Properties)
        Call FormatChart(ChtOld.GetOldChart, Properties)

    End If


    Exit Sub

ErrorHandler:

    ErrorMod.ErrorMessage cProc, cModule


End Sub

Private Sub ModifyChart(MyChart As Chart, Properties As MyChart)

    Call ChtModify.Initiate(MyChart, Properties)
    Call ChtModify.ModifyChartTitle
    Call ChtModify.ModifySourceBox

End Sub


Private Sub FormatChart(NewChart As Chart, Properties As MyChart)

    Const cProc = "FormatChart"

    Call ChtFormat.Initiate(NewChart, Properties)
    Call ChtFormat.FormatChartSize
    Call ChtFormat.FormatChartTitleBox
    Call ChtFormat.FormatSourceTextBox
    Call ChtFormat.FormatChartTitle
    Call ChtFormat.FormatChartAxisTitle
    Call ChtFormat.FormatChartSeries
    Call ChtFormat.FormatChartAxes
    Call ChtFormat.FormatLegend
    Call ChtFormat.FormatPlotArea

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

Private Sub chkCopy_Click()

    If chkCopy.Value = True Then

        optLang2.Enabled = True

    Else

        optLang2.Enabled = False

    End If

End Sub


Private Sub UserForm_Click()

End Sub
