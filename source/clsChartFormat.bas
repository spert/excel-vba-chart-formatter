
Option Explicit On

Const cModule = "frmChtFormat"

Private ChtFormat As New clsChartFormat
Private NewWorkSheet As New clsWorkSheet
Private arrMySeries() As clsSeries

Dim oChtOb As ChartObject

Property Let SetChartObject(ChtObject As ChartObject)

    Set oChtOb = ChtObject

End Property


Private Sub chkTitleTextBox_Click()

End Sub

Private Sub cmdExecute_Click()

    Const cProc = "cmdExecute_Click"

    ErrorMod.ResetErrorCounter

    'On Error GoTo ErrorHandler:

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    Dim aOptions(9) As Variant
    aOptions(0) = optSmall.Value
    aOptions(1) = optSlide.Value
    aOptions(2) = optLang1.Value
    aOptions(3) = optLang2.Value
    aOptions(4) = optTitleTop.Value
    aOptions(5) = chkTitleTextBox.Value
    aOptions(6) = chkSourceBox.Value
    aOptions(7) = chkLinks.Value
    aOptions(8) = chkRescale.Value

    Dim Settings As New clsSettings
    Call Settings.InitiateSettings(ActiveWorkbook, ActiveChart, aOptions)

    Dim ChtOld As New clsChartOld
    Call ChtOld.Initiate(oChtOb.Chart)
    Call ChtOld.ValidateExcelChart

    If ErrorMod.GetErrorCount > 0 Then
        Exit Sub
    End If

    Call ChtOld.CollectSeries
    Call ChtOld.CollectAxes

    If chkCopy.Value = True Then

        ' --------- Create Worksheet Layout ------------
        Call NewWorkSheet.Initiate
        Call NewWorkSheet.AddNewChartToWorkbook
        Call NewWorkSheet.TurnOffGridLines

        Dim props() As MyChart
        props = Settings.GetCopyChartProps()

        Dim j As Integer

        For j = 0 To UBound(props)

            ' --------- Format Worksheet Layout ------------
            Dim prop_c As MyChart
            prop_c = props(j)

            Dim ChtNew As New clsChartNew
            Call ChtNew.Initiate(NewWorkSheet, ChtOld, prop_c)
            Call ChtNew.CopyOldChartToNewWorksheet
            Call ChtNew.AssignNewRanges
            Call ChtNew.MapNewSeries

            Dim ChtModify As New clsChartModify
            Call ChtModify.InitiateFromNew(ChtNew, prop_c)
            Call ChtModify.ModifyChartTitle
            Call ChtModify.ModifySourceBox
            Call ChtModify.RescaleXAxis

            Call ChtNew.PrintHeaders
            Call ChtNew.PrintChartTitle
            Call ChtNew.PrintSourceTextBox
            Call ChtNew.PrintChartNumber
            Call ChtNew.PrintChartAxisTitle
            Call ChtNew.PrintAxesAndSeriesValues
            Call ChtNew.PrintSeriesNamesLinksAndScale

            Call FormatChart(ChtNew.GetNewChart, prop_c)

        Next

    Else

        Dim prop_s As MyChart
        prop_s = Settings.GetFormatChartProps()

        Dim ChtModifySingle As New clsChartModify
        Call ChtModifySingle.InitiateFromOld(ChtOld, prop_s)
        Call ChtModifySingle.RescaleXAxis
        Call ChtModifySingle.ModifyChartTitle
        Call ChtModifySingle.ModifySourceBox
        Call FormatChart(ChtOld.GetOldChart, prop_s)

    End If


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
        Dim prop_r As MyChart
        prop_r = Settings.GetFormatChartProps()
        Call ApplyRecentColors(prop_r)
    End If

    Application.DisplayAlerts = True
    Application.ScreenUpdating = False

    Exit Sub

ErrorHandler:

    MsgBox "Unexpected error in exec module", vbOKOnly

End Sub


Private Sub FormatChart(NewChart As Chart, Properties As MyChart)

    Const cProc = "FormatChart"

    Call ChtFormat.Initiate(NewChart, Properties)
    Call ChtFormat.FormatChartSize
    Call ChtFormat.FormatChartTitle
    Call ChtFormat.FormatChartTitleBox
    Call ChtFormat.FormatSourceTextBox
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
