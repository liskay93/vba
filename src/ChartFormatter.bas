Attribute VB_Name = "ChartFormatter"
Option Explicit

'==============================================================================
' ChartFormatter
'
' Select a chart, run FormatSelectedChart. Colors, fonts, size, axes, gridlines
' and legend are applied in one pass. The chart's position (Left/Top) is never
' changed.
'
' Everything tunable lives in the SETTINGS block below.
' Note: VBA clears Excel's undo stack - there is no Ctrl+Z after running this.
'==============================================================================


'============================ SETTINGS ========================================

' --- Fonts -------------------------------------------------------------------
Private Const FONT_LATIN      As String = "Malgun Gothic"  ' latin / digits
Private Const FONT_EASTASIA   As String = "Malgun Gothic"  ' korean / CJK
Private Const FONT_SIZE_BASE  As Single = 10               ' axes, legend, labels
Private Const FONT_SIZE_TITLE As Single = 14

' --- Size (position is always preserved) -------------------------------------
Private Const APPLY_SIZE      As Boolean = True   ' False = leave size untouched
Private Const CHART_WIDTH     As Single = 480     ' points
Private Const CHART_HEIGHT    As Single = 300     ' points

' --- Marks -------------------------------------------------------------------
Private Const LINE_WEIGHT     As Single = 2       ' points, line series
Private Const MARKER_SIZE     As Single = 6       ' points, ~8px
Private Const FILL_GAP_WEIGHT As Single = 1.5     ' surface-colored ring on fills
Private Const GAP_WIDTH       As Long = 80        ' bar/column gap (0-500)
Private Const OVERLAP_CLUST   As Long = -5        ' clustered series separation
Private Const SMOOTH_OFF      As Boolean = True   ' straighten smoothed lines

' --- Behaviour ---------------------------------------------------------------
Private Const LEFT_ALIGN_TITLE  As Boolean = True
Private Const PIE_DATA_LABELS   As Boolean = True  ' pie: labels instead of legend
Private Const VALUE_AXIS_FORMAT As String = ""     ' e.g. "#,##0" - "" leaves as-is

' --- Palette (validated: CVD dE 9.1, normal-vision dE 19.6 on white) ----------
' Assigned in fixed slot order and never cycled. Slot 9+ falls back to a neutral
' gray on purpose - fold extra series into "Other" or split into small multiples.
Private Const C_S1 As Long = &HD6782A&   ' #2A78D6 blue
Private Const C_S2 As Long = &H3468EB&   ' #EB6834 orange
Private Const C_S3 As Long = &H7AAF1B&   ' #1BAF7A aqua
Private Const C_S4 As Long = &H00A1ED&   ' #EDA100 yellow
Private Const C_S5 As Long = &HA47BE8&   ' #E87BA4 magenta
Private Const C_S6 As Long = &H008300&   ' #008300 green
Private Const C_S7 As Long = &HA73A4A&   ' #4A3AA7 violet
Private Const C_S8 As Long = &H4849E3&   ' #E34948 red
Private Const C_EXTRA As Long = &H84898A& ' #8A8984 neutral, slot 9+

' --- Surface / ink -----------------------------------------------------------
Private Const C_SURFACE   As Long = &HFFFFFF&   ' #FFFFFF chart surface
Private Const C_TEXT_PRI  As Long = &H0B0B0B&   ' #0B0B0B title
Private Const C_TEXT_SEC  As Long = &H4E5152&   ' #52514E axis / legend text
Private Const C_GRID      As Long = &HD2D7D8&   ' #D8D7D2 gridlines
Private Const C_AXIS      As Long = &HBDC3C4&   ' #C4C3BD axis line

'========================== END SETTINGS ======================================


Private mWarnings As String


'--- Public entry points ------------------------------------------------------

' Select a chart, then run this.
Public Sub FormatSelectedChart()
    Dim ch As Chart

    Set ch = TargetChart()
    If ch Is Nothing Then
        MsgBox "Select a chart first, then run this macro.", _
               vbExclamation, "Chart Formatter"
        Exit Sub
    End If

    mWarnings = ""
    ApplyFormat ch
    ShowWarnings
End Sub

' Convenience: same formatting applied to every chart on the active sheet.
Public Sub FormatAllChartsOnSheet()
    Dim co As ChartObject
    Dim n As Long

    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "Run this from a worksheet that holds charts.", _
               vbExclamation, "Chart Formatter"
        Exit Sub
    End If

    mWarnings = ""
    For Each co In ActiveSheet.ChartObjects
        ApplyFormat co.Chart
        n = n + 1
    Next co

    If n = 0 Then
        MsgBox "No charts found on this sheet.", vbInformation, "Chart Formatter"
    Else
        ShowWarnings
    End If
End Sub


'--- Core ---------------------------------------------------------------------

Private Sub ApplyFormat(ch As Chart)
    Dim co As ChartObject
    Dim savedLeft As Single, savedTop As Single
    Dim embedded As Boolean
    Dim prevScreen As Boolean

    prevScreen = Application.ScreenUpdating
    On Error GoTo Cleanup
    Application.ScreenUpdating = False

    embedded = (TypeName(ch.Parent) = "ChartObject")
    If embedded Then
        Set co = ch.Parent
        savedLeft = co.Left
        savedTop = co.Top
    End If

    NoteChartWarnings ch

    StyleChartArea ch
    StylePlotArea ch
    StyleTitle ch
    StyleAxes ch
    StyleSeries ch
    StyleGroups ch
    StyleLegend ch

    If embedded Then
        If APPLY_SIZE Then
            co.Width = CHART_WIDTH
            co.Height = CHART_HEIGHT
        End If
        ' Resizing keeps the anchor, but restore explicitly so the chart can
        ' never drift from where the user put it.
        co.Left = savedLeft
        co.Top = savedTop
    End If

    Application.ScreenUpdating = prevScreen
    Exit Sub

Cleanup:
    Application.ScreenUpdating = prevScreen
    MsgBox "Could not finish formatting this chart." & vbLf & vbLf & _
           Err.Number & ": " & Err.Description, vbExclamation, "Chart Formatter"
End Sub


'--- Chart-level elements -----------------------------------------------------

Private Sub StyleChartArea(ch As Chart)
    On Error Resume Next

    ch.RoundedCorners = False

    With ch.ChartArea.Format
        .Fill.Visible = msoTrue
        .Fill.Solid
        .Fill.ForeColor.RGB = C_SURFACE
        .Fill.Transparency = 0
        .Line.Visible = msoFalse
        .Shadow.Visible = msoFalse
        .Glow.Radius = 0
    End With

    ' Base font for the whole chart; title is overridden afterwards.
    With ch.ChartArea.Format.TextFrame2.TextRange.Font
        .Name = FONT_LATIN
        .NameFarEast = FONT_EASTASIA
        .Size = FONT_SIZE_BASE
        .Bold = msoFalse
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = C_TEXT_SEC
    End With
End Sub


Private Sub StylePlotArea(ch As Chart)
    On Error Resume Next
    With ch.PlotArea.Format
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        .Shadow.Visible = msoFalse
    End With
End Sub


Private Sub StyleTitle(ch As Chart)
    On Error Resume Next
    If Not ch.HasTitle Then Exit Sub

    With ch.ChartTitle.Format.TextFrame2.TextRange.Font
        .Name = FONT_LATIN
        .NameFarEast = FONT_EASTASIA
        .Size = FONT_SIZE_TITLE
        .Bold = msoTrue
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = C_TEXT_PRI
    End With

    If LEFT_ALIGN_TITLE Then
        ch.ChartTitle.Format.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
        ch.ChartTitle.Left = ch.PlotArea.InsideLeft
    End If
End Sub


Private Sub StyleAxes(ch As Chart)
    ' Category axis keeps a baseline; the value axis is carried by gridlines.
    StyleOneAxis ch, xlCategory, xlPrimary, True, False
    StyleOneAxis ch, xlValue, xlPrimary, False, True
    If ch.HasAxis(xlValue, xlSecondary) Then
        StyleOneAxis ch, xlValue, xlSecondary, False, False
    End If
End Sub


Private Sub StyleOneAxis(ch As Chart, axType As XlAxisType, axGroup As XlAxisGroup, _
                         showLine As Boolean, showGrid As Boolean)
    Dim ax As Axis

    On Error Resume Next
    Set ax = ch.Axes(axType, axGroup)
    If ax Is Nothing Then Exit Sub

    With ax
        .MajorTickMark = xlTickMarkNone
        .MinorTickMark = xlTickMarkNone
        .HasMinorGridlines = False
        .HasMajorGridlines = showGrid

        If showLine Then
            .Format.Line.Visible = msoTrue
            .Format.Line.ForeColor.RGB = C_AXIS
            .Format.Line.Weight = 0.75
        Else
            .Format.Line.Visible = msoFalse
        End If

        .TickLabels.Font.Color = C_TEXT_SEC
        .TickLabels.Font.Size = FONT_SIZE_BASE
    End With

    If showGrid Then
        With ax.MajorGridlines.Format.Line
            .Visible = msoTrue
            .ForeColor.RGB = C_GRID
            .Weight = 0.75
            .DashStyle = msoLineSolid
        End With
    End If

    If axType = xlValue And Len(VALUE_AXIS_FORMAT) > 0 Then
        ax.TickLabels.NumberFormat = VALUE_AXIS_FORMAT
    End If
End Sub


Private Sub StyleLegend(ch As Chart)
    On Error Resume Next

    If ch.SeriesCollection.Count = 0 Then Exit Sub

    ' A pie carrying its own labels does not also need a legend.
    If IsPieLike(SeriesType(ch.SeriesCollection(1))) Then
        ch.HasLegend = Not PIE_DATA_LABELS
    Else
        ' One series is named by the title; two or more always need a legend so
        ' identity never rests on color alone.
        ch.HasLegend = (ch.SeriesCollection.Count >= 2)
    End If

    If Not ch.HasLegend Then Exit Sub

    With ch.Legend
        .Position = xlLegendPositionBottom
        .Format.Fill.Visible = msoFalse
        .Format.Line.Visible = msoFalse
        With .Format.TextFrame2.TextRange.Font
            .Name = FONT_LATIN
            .NameFarEast = FONT_EASTASIA
            .Size = FONT_SIZE_BASE
            .Fill.ForeColor.RGB = C_TEXT_SEC
        End With
    End With
End Sub


'--- Series -------------------------------------------------------------------

Private Sub StyleSeries(ch As Chart)
    Dim i As Long
    Dim srs As Series
    Dim t As Long

    For i = 1 To ch.SeriesCollection.Count
        Set srs = ch.SeriesCollection(i)
        t = SeriesType(srs)

        If IsPieLike(t) Then
            StylePieSeries srs
        ElseIf IsScatterLike(t) Then
            StyleScatterSeries srs, PaletteColor(i), t
        ElseIf IsLineLike(t) Then
            StyleLineSeries srs, PaletteColor(i)
        ElseIf IsAreaLike(t) Then
            StyleAreaSeries srs, PaletteColor(i)
        Else
            StyleFillSeries srs, PaletteColor(i)
        End If
    Next i
End Sub


Private Sub StyleFillSeries(srs As Series, clr As Long)
    On Error Resume Next

    With srs.Format.Fill
        .Visible = msoTrue
        .Solid
        .ForeColor.RGB = clr
        .Transparency = 0
    End With

    ' Surface-colored ring: the gap that keeps adjacent and stacked fills apart.
    With srs.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = C_SURFACE
        .Weight = FILL_GAP_WEIGHT
    End With

    srs.Format.Shadow.Visible = msoFalse
    srs.Format.Glow.Radius = 0
    srs.Shadow = False
End Sub


Private Sub StyleAreaSeries(srs As Series, clr As Long)
    On Error Resume Next

    With srs.Format.Fill
        .Visible = msoTrue
        .Solid
        .ForeColor.RGB = clr
        .Transparency = 0.25
    End With

    With srs.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = clr
        .Weight = LINE_WEIGHT
    End With

    srs.Format.Shadow.Visible = msoFalse
    srs.Shadow = False
End Sub


Private Sub StyleLineSeries(srs As Series, clr As Long)
    On Error Resume Next

    With srs.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = clr
        .Weight = LINE_WEIGHT
        .DashStyle = msoLineSolid
    End With

    If SMOOTH_OFF Then srs.Smooth = False
    srs.Format.Shadow.Visible = msoFalse
    srs.Format.Glow.Radius = 0
    srs.Shadow = False

    ' Only size markers the chart already has - never add them.
    If srs.MarkerStyle <> xlMarkerStyleNone Then
        srs.MarkerSize = MARKER_SIZE
        srs.MarkerBackgroundColor = clr
        srs.MarkerForegroundColor = C_SURFACE
    End If
End Sub


Private Sub StyleScatterSeries(srs As Series, clr As Long, t As Long)
    On Error Resume Next

    Select Case t
        Case xlXYScatterLinesNoMarkers, xlXYScatterSmoothNoMarkers
            ' lines only
        Case Else
            If srs.MarkerStyle = xlMarkerStyleNone Then
                srs.MarkerStyle = xlMarkerStyleCircle
            End If
            srs.MarkerSize = MARKER_SIZE
            srs.MarkerBackgroundColor = clr
            srs.MarkerForegroundColor = C_SURFACE
    End Select

    Select Case t
        Case xlXYScatterLines, xlXYScatterLinesNoMarkers, _
             xlXYScatterSmooth, xlXYScatterSmoothNoMarkers
            With srs.Format.Line
                .Visible = msoTrue
                .ForeColor.RGB = clr
                .Weight = LINE_WEIGHT
            End With
        Case Else
            srs.Format.Line.Visible = msoFalse
    End Select

    srs.Format.Shadow.Visible = msoFalse
    srs.Shadow = False
End Sub


Private Sub StylePieSeries(srs As Series)
    Dim p As Long

    On Error Resume Next

    ' A pie colors by point, not by series.
    For p = 1 To srs.Points.Count
        With srs.Points(p).Format.Fill
            .Visible = msoTrue
            .Solid
            .ForeColor.RGB = PaletteColor(p)
            .Transparency = 0
        End With
        With srs.Points(p).Format.Line
            .Visible = msoTrue
            .ForeColor.RGB = C_SURFACE
            .Weight = FILL_GAP_WEIGHT
        End With
    Next p

    If PIE_DATA_LABELS Then
        srs.HasDataLabels = True
        With srs.DataLabels
            .ShowCategoryName = True
            .ShowPercentage = True
            .ShowValue = False
            .ShowSeriesName = False
            .Separator = vbLf
            .Position = xlLabelPositionOutsideEnd
            With .Format.TextFrame2.TextRange.Font
                .Name = FONT_LATIN
                .NameFarEast = FONT_EASTASIA
                .Size = FONT_SIZE_BASE
                .Fill.ForeColor.RGB = C_TEXT_SEC
            End With
        End With
    End If
End Sub


'--- Chart groups (bar width / clustering) ------------------------------------

Private Sub StyleGroups(ch As Chart)
    Dim i As Long
    Dim g As ChartGroup
    Dim t As Long

    On Error Resume Next

    For i = 1 To ch.ChartGroups.Count
        Set g = ch.ChartGroups(i)

        t = 0
        t = g.SeriesCollection(1).ChartType

        g.GapWidth = GAP_WIDTH
        If IsStacked(t) Then
            g.Overlap = 100          ' stacks must stay fully overlapped
        Else
            g.Overlap = OVERLAP_CLUST
        End If
    Next i
End Sub


'--- Helpers ------------------------------------------------------------------

Private Function TargetChart() As Chart
    On Error Resume Next
    Set TargetChart = ActiveChart
    If TargetChart Is Nothing Then
        If TypeName(Selection) = "ChartObject" Then Set TargetChart = Selection.Chart
    End If
End Function


Private Function PaletteColor(idx As Long) As Long
    Select Case idx
        Case 1: PaletteColor = C_S1
        Case 2: PaletteColor = C_S2
        Case 3: PaletteColor = C_S3
        Case 4: PaletteColor = C_S4
        Case 5: PaletteColor = C_S5
        Case 6: PaletteColor = C_S6
        Case 7: PaletteColor = C_S7
        Case 8: PaletteColor = C_S8
        Case Else: PaletteColor = C_EXTRA
    End Select
End Function


Private Function SeriesType(srs As Series) As Long
    Dim t As Long
    On Error Resume Next
    t = 0
    t = srs.ChartType
    If t = 0 Then t = xlColumnClustered   ' unreadable type - treat as a fill
    SeriesType = t
End Function


Private Function IsPieLike(t As Long) As Boolean
    Select Case t
        Case xlPie, xlPieExploded, xlPieOfPie, xlBarOfPie, _
             xl3DPie, xl3DPieExploded, xlDoughnut, xlDoughnutExploded
            IsPieLike = True
    End Select
End Function


Private Function IsLineLike(t As Long) As Boolean
    Select Case t
        Case xlLine, xlLineStacked, xlLineStacked100, _
             xlLineMarkers, xlLineMarkersStacked, xlLineMarkersStacked100, _
             xl3DLine, xlRadar, xlRadarMarkers
            IsLineLike = True
    End Select
End Function


Private Function IsScatterLike(t As Long) As Boolean
    Select Case t
        Case xlXYScatter, xlXYScatterLines, xlXYScatterLinesNoMarkers, _
             xlXYScatterSmooth, xlXYScatterSmoothNoMarkers
            IsScatterLike = True
    End Select
End Function


Private Function IsAreaLike(t As Long) As Boolean
    Select Case t
        Case xlArea, xlAreaStacked, xlAreaStacked100, _
             xl3DArea, xl3DAreaStacked, xl3DAreaStacked100, xlRadarFilled
            IsAreaLike = True
    End Select
End Function


Private Function IsStacked(t As Long) As Boolean
    Select Case t
        Case xlColumnStacked, xlColumnStacked100, _
             xlBarStacked, xlBarStacked100, _
             xl3DColumnStacked, xl3DColumnStacked100, _
             xl3DBarStacked, xl3DBarStacked100, _
             xlAreaStacked, xlAreaStacked100, _
             xl3DAreaStacked, xl3DAreaStacked100, _
             xlLineStacked, xlLineStacked100, _
             xlLineMarkersStacked, xlLineMarkersStacked100
            IsStacked = True
    End Select
End Function


Private Function Is3D(t As Long) As Boolean
    Select Case t
        Case xl3DArea, xl3DAreaStacked, xl3DAreaStacked100, _
             xl3DBarClustered, xl3DBarStacked, xl3DBarStacked100, _
             xl3DColumn, xl3DColumnClustered, xl3DColumnStacked, xl3DColumnStacked100, _
             xl3DLine, xl3DPie, xl3DPieExploded, xlSurface, xlSurfaceWireframe
            Is3D = True
    End Select
End Function


Private Sub NoteChartWarnings(ch As Chart)
    Dim n As Long

    On Error Resume Next

    n = ch.SeriesCollection.Count
    If n > 8 Then
        AddWarning "This chart has " & n & " series. Slots 9+ are painted a " & _
                   "neutral gray on purpose - the palette is not cycled. Fold " & _
                   "the extras into an 'Other' series, or split the chart."
    End If

    If ch.HasAxis(xlValue, xlSecondary) Then
        AddWarning "This chart has a secondary value axis. Two y-scales are " & _
                   "easy to misread - consider two charts, or index both " & _
                   "series to a common base."
    End If

    If Is3D(ch.ChartType) Then
        AddWarning "3-D chart types ignore part of this formatting, and the " & _
                   "perspective distorts the values. A flat chart reads better."
    End If
End Sub


Private Sub AddWarning(msg As String)
    If InStr(1, mWarnings, msg, vbTextCompare) > 0 Then Exit Sub
    If Len(mWarnings) > 0 Then mWarnings = mWarnings & vbLf & vbLf
    mWarnings = mWarnings & "- " & msg
End Sub


Private Sub ShowWarnings()
    If Len(mWarnings) = 0 Then Exit Sub
    MsgBox "Formatting applied. Worth a look:" & vbLf & vbLf & mWarnings, _
           vbInformation, "Chart Formatter"
    mWarnings = ""
End Sub
