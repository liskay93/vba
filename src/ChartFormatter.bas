Attribute VB_Name = "ChartFormatter"
Option Explicit

'==============================================================================
' ChartFormatter
'
' 차트를 선택한 뒤 FormatSelectedChart 를 실행하면 색상/폰트/크기/축/범례
' 서식이 한 번에 적용됩니다. 차트 위치(Left/Top)는 절대 바뀌지 않습니다.
'
' 폰트와 색상 체계는 기존 ESG 리포트 매크로(bar / line / Doughnut /
' DataLabel_Add)의 하우스 스타일을 따릅니다.
'
' 주의: VBA 실행 시 엑셀의 실행 취소 기록이 지워집니다. Ctrl+Z 불가.
'==============================================================================


'============================ 설정 ============================================

' --- 폰트 (하우스 스타일) ----------------------------------------------------
Private Const FONT_BODY         As String = "맑은 고딕"
Private Const FONT_LEGEND       As String = "KoPubWorld돋움체 Medium"
Private Const FONT_SIZE_BASE    As Single = 8.5   ' 축/데이터 레이블
Private Const FONT_SIZE_LEGEND  As Single = 9
Private Const FONT_SIZE_TITLE   As Single = 12

' --- 크기 (위치는 항상 보존) -------------------------------------------------
Private Const APPLY_SIZE        As Boolean = False ' True 로 바꾸면 아래 크기 적용
Private Const CHART_WIDTH       As Single = 311    ' 리포트 그리드: 198/311/510/623
Private Const CHART_HEIGHT      As Single = 215

' --- 팔레트 모드 -------------------------------------------------------------
'   "CATEGORICAL" = 서로 다른 대상(당사/경쟁사/업종 등). 색이 이름표 역할.
'   "SEQUENTIAL"  = 순서가 있는 등급/누적 구간. 한 색의 진하기로 순서를 표현.
Private Const PALETTE_MODE      As String = "CATEGORICAL"

' --- 마크 --------------------------------------------------------------------
Private Const LINE_WEIGHT       As Single = 1.5   ' 꺾은선 두께
Private Const MARKER_SIZE       As Single = 5
Private Const FILL_GAP_WEIGHT   As Single = 1.5   ' 채우기 경계(배경색 선)
Private Const GAP_WIDTH         As Long = 80      ' 막대 간격 너비
Private Const OVERLAP_CLUST     As Long = -5
Private Const SMOOTH_OFF        As Boolean = True

' --- 동작 --------------------------------------------------------------------
Private Const AXIS_WEIGHT       As Single = 0.25  ' 축 선 두께 (하우스 스타일)
Private Const SHOW_GRIDLINES    As Boolean = False ' 하우스 스타일은 그리드 제거
Private Const LEFT_ALIGN_TITLE  As Boolean = True
Private Const PIE_DATA_LABELS   As Boolean = True
Private Const VALUE_AXIS_FORMAT As String = ""    ' 예: "#,##0". 빈 값이면 건드리지 않음

'------------------------------------------------------------------------------
' 색상 - VBA 는 &HBBGGRR& (RGB 역순). 주석의 #RRGGBB 가 원래 값입니다.
'
' 범주형 6색: 흰 배경에서 6색 전부 명암비 3:1 이상, 인접쌍 색약 구분 dEE 11.8,
' 일반 시야 dEE 15.5 로 모든 기준 통과. 슬롯 순서가 안전장치이므로 순서를
' 바꾸려면 재검증이 필요합니다.
'------------------------------------------------------------------------------
Private Const C_CAT1 As Long = &HB63900&   ' #0039B6 브랜드 블루 (기존 색 유지)
Private Const C_CAT2 As Long = &H0C73E8&   ' #E8730C 주황
Private Const C_CAT3 As Long = &H7B8900&   ' #00897B 청록
Private Const C_CAT4 As Long = &H006D8A&   ' #8A6D00 황토
Private Const C_CAT5 As Long = &HA21F7B&   ' #7B1FA2 자주
Private Const C_CAT6 As Long = &H5B18C2&   ' #C2185B 적자

' 순차 4단 (진한 -> 연한). 연한 끝도 명암비 2.11:1 로 흰 배경에서 보입니다.
Private Const C_SEQ1 As Long = &H6B2400&   ' #00246B
Private Const C_SEQ2 As Long = &HB63900&   ' #0039B6
Private Const C_SEQ3 As Long = &HD97F4A&   ' #4A7FD9
Private Const C_SEQ4 As Long = &HEFB686&   ' #86B6EF

' 역할색 - 계열 슬롯이 아니라 고정 역할. 순번으로 배정하지 않습니다.
Private Const C_NEUTRAL As Long = &H7F7F7F&  ' #7F7F7F 기타/비교군 (4.00:1)
Private Const C_ACCENT  As Long = &H0CA30C&  ' #0CA30C 강조 초록 (3.35:1)

' 잉크 / 배경
Private Const C_TEXT    As Long = &H000000&  ' #000000 모든 차트 텍스트
Private Const C_AXIS    As Long = &H000000&  ' #000000 축 선
Private Const C_GRID    As Long = &HD2D7D8&  ' #D8D7D2 그리드 (사용 시)
Private Const C_SURFACE As Long = &HFFFFFF&  ' #FFFFFF 배경

'========================== 설정 끝 ===========================================


Private mWarnings As String


'--- 실행 진입점 --------------------------------------------------------------

' 차트를 선택한 뒤 이것을 실행합니다.
Public Sub FormatSelectedChart()
    Dim ch As Chart

    Set ch = TargetChart()
    If ch Is Nothing Then
        MsgBox "차트를 먼저 선택한 뒤 실행하세요.", vbExclamation, "ChartFormatter"
        Exit Sub
    End If

    mWarnings = ""
    ApplyFormat ch
    ShowWarnings
End Sub

' 현재 시트의 모든 차트에 같은 서식을 적용합니다.
Public Sub FormatAllChartsOnSheet()
    Dim co As ChartObject
    Dim n As Long

    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "차트가 있는 워크시트에서 실행하세요.", vbExclamation, "ChartFormatter"
        Exit Sub
    End If

    mWarnings = ""
    For Each co In ActiveSheet.ChartObjects
        ApplyFormat co.Chart
        n = n + 1
    Next co

    If n = 0 Then
        MsgBox "이 시트에는 차트가 없습니다.", vbInformation, "ChartFormatter"
    Else
        ShowWarnings
    End If
End Sub


'--- 본체 ---------------------------------------------------------------------

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
        ' 크기를 바꿔도 앵커는 유지되지만, 어떤 경우에도 차트가 밀려나지
        ' 않도록 원래 위치를 명시적으로 되돌려 놓습니다.
        co.Left = savedLeft
        co.Top = savedTop
    End If

    Application.ScreenUpdating = prevScreen
    Exit Sub

Cleanup:
    Application.ScreenUpdating = prevScreen
    MsgBox "서식 적용을 끝내지 못했습니다." & vbLf & vbLf & _
           Err.Number & ": " & Err.Description, vbExclamation, "ChartFormatter"
End Sub


'--- 차트 단위 요소 -----------------------------------------------------------

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

    ' 차트 전체 기본 폰트. 제목과 범례는 뒤에서 덮어씁니다.
    With ch.ChartArea.Format.TextFrame2.TextRange.Font
        .Name = FONT_BODY
        .NameFarEast = FONT_BODY
        .NameComplexScript = FONT_BODY
        .Size = FONT_SIZE_BASE
        .Bold = msoFalse
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = C_TEXT
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
        .Name = FONT_BODY
        .NameFarEast = FONT_BODY
        .Size = FONT_SIZE_TITLE
        .Bold = msoTrue
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = C_TEXT
    End With

    If LEFT_ALIGN_TITLE Then
        ch.ChartTitle.Format.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
        ch.ChartTitle.Left = ch.PlotArea.InsideLeft
    End If
End Sub


Private Sub StyleAxes(ch As Chart)
    StyleOneAxis ch, xlCategory, xlPrimary, True, False
    StyleOneAxis ch, xlValue, xlPrimary, True, SHOW_GRIDLINES
    If ch.HasAxis(xlValue, xlSecondary) Then
        StyleOneAxis ch, xlValue, xlSecondary, True, False
    End If
End Sub


Private Sub StyleOneAxis(ch As Chart, axType As XlAxisType, axGroup As XlAxisGroup, _
                         showLine As Boolean, showGrid As Boolean)
    Dim ax As Axis

    On Error Resume Next
    Set ax = ch.Axes(axType, axGroup)
    If ax Is Nothing Then Exit Sub

    With ax
        .MajorTickMark = xlTickMarkInside
        .MinorTickMark = xlTickMarkNone
        .HasMinorGridlines = False
        .HasMajorGridlines = showGrid

        If showLine Then
            .Format.Line.Visible = msoTrue
            .Format.Line.ForeColor.RGB = C_AXIS
            .Format.Line.Weight = AXIS_WEIGHT
        Else
            .Format.Line.Visible = msoFalse
        End If

        .TickLabels.Font.Color = C_TEXT
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

    If IsPieLike(SeriesType(ch.SeriesCollection(1))) Then
        ' 원형/도넛은 레이블이 이름을 말해주므로 범례를 중복으로 두지 않습니다.
        ch.HasLegend = Not PIE_DATA_LABELS
    Else
        ' 계열 1개는 제목이 이름을 말해주고, 2개 이상은 색만으로 구분하게
        ' 두지 않기 위해 범례를 항상 켭니다.
        ch.HasLegend = (ch.SeriesCollection.Count >= 2)
    End If

    If Not ch.HasLegend Then Exit Sub

    With ch.Legend
        .Position = xlLegendPositionTop
        .Format.Fill.Visible = msoFalse
        .Format.Line.Visible = msoFalse
        With .Format.TextFrame2.TextRange.Font
            .Name = FONT_LEGEND
            .NameFarEast = FONT_LEGEND
            .Size = FONT_SIZE_LEGEND
            .Fill.Visible = msoTrue
            .Fill.ForeColor.RGB = C_TEXT
        End With
    End With
End Sub


'--- 계열 --------------------------------------------------------------------

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

    ' 배경색 경계선: 붙어 있는 막대와 누적 구간을 떼어 놓는 틈.
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

    ' 이미 마커가 있는 계열만 손봅니다. 없던 마커를 새로 만들지는 않습니다.
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
            ' 선만 쓰는 형식 - 마커를 만들지 않습니다.
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

    ' 원형/도넛은 계열이 아니라 조각별로 색을 배정합니다.
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
            .ShowValue = True
            .ShowPercentage = False
            .ShowSeriesName = False
            .Separator = vbLf
            .Position = xlLabelPositionCenter
            With .Format.TextFrame2.TextRange.Font
                .Name = FONT_BODY
                .NameFarEast = FONT_BODY
                .Size = FONT_SIZE_BASE
                .Bold = msoTrue
                .Fill.Visible = msoTrue
                .Fill.ForeColor.RGB = C_SURFACE   ' 조각 위에 흰 글씨
            End With
        End With
    End If
End Sub


'--- 차트 그룹 (막대 굵기 / 묶기) --------------------------------------------

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
            g.Overlap = 100          ' 누적은 완전히 겹쳐 있어야 합니다.
        Else
            g.Overlap = OVERLAP_CLUST
        End If
    Next i
End Sub


'--- 보조 함수 ----------------------------------------------------------------

Private Function TargetChart() As Chart
    On Error Resume Next
    Set TargetChart = ActiveChart
    If TargetChart Is Nothing Then
        If TypeName(Selection) = "ChartObject" Then Set TargetChart = Selection.Chart
    End If
End Function


' 색을 순번대로 배정합니다. 슬롯을 다 쓰면 순환하지 않고 회색으로 떨어집니다.
' 순환시키면 서로 다른 계열이 같은 색이 되어 차트가 거짓말을 하게 됩니다.
Private Function PaletteColor(idx As Long) As Long
    If PALETTE_MODE = "SEQUENTIAL" Then
        Select Case idx
            Case 1: PaletteColor = C_SEQ1
            Case 2: PaletteColor = C_SEQ2
            Case 3: PaletteColor = C_SEQ3
            Case 4: PaletteColor = C_SEQ4
            Case Else: PaletteColor = C_NEUTRAL
        End Select
    Else
        Select Case idx
            Case 1: PaletteColor = C_CAT1
            Case 2: PaletteColor = C_CAT2
            Case 3: PaletteColor = C_CAT3
            Case 4: PaletteColor = C_CAT4
            Case 5: PaletteColor = C_CAT5
            Case 6: PaletteColor = C_CAT6
            Case Else: PaletteColor = C_NEUTRAL
        End Select
    End If
End Function


Private Function PaletteSlots() As Long
    If PALETTE_MODE = "SEQUENTIAL" Then
        PaletteSlots = 4
    Else
        PaletteSlots = 6
    End If
End Function


Private Function SeriesType(srs As Series) As Long
    Dim t As Long
    On Error Resume Next
    t = 0
    t = srs.ChartType
    If t = 0 Then t = xlColumnClustered   ' 종류를 읽을 수 없으면 채우기로 취급
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
    If n > PaletteSlots() Then
        AddWarning "계열이 " & n & "개입니다. 팔레트 슬롯은 " & PaletteSlots() & _
                   "개이고, 색을 돌려쓰면 서로 다른 계열이 같은 색이 되므로 " & _
                   (PaletteSlots() + 1) & "번째부터는 회색으로 칠합니다. " & _
                   "나머지를 '기타'로 묶거나 차트를 나누세요."
    End If

    ' 원형/도넛은 모든 조각이 동시에 보이므로 구분 기준이 더 엄격합니다.
    ' 이 팔레트는 동시 비교 기준으로는 4색까지만 안전합니다.
    If ch.SeriesCollection.Count > 0 Then
        If IsPieLike(SeriesType(ch.SeriesCollection(1))) Then
            If ch.SeriesCollection(1).Points.Count > 4 Then
                AddWarning "도넛/원형 조각이 " & ch.SeriesCollection(1).Points.Count & _
                           "개입니다. 모든 조각이 한눈에 같이 보이기 때문에 이 " & _
                           "팔레트는 4조각까지만 색약 구분이 보장됩니다. 작은 " & _
                           "조각은 '기타'로 합치거나 막대차트로 바꾸세요."
            End If
        End If
    End If

    If ch.HasAxis(xlValue, xlSecondary) Then
        AddWarning "보조 축(축 2개)이 있습니다. 축척이 다른 두 지표는 잘못 " & _
                   "읽기 쉽습니다. 차트를 2개로 나누거나 같은 기준으로 " & _
                   "지수화하는 편이 오해가 적습니다."
    End If

    If Is3D(ch.ChartType) Then
        AddWarning "3차원 차트는 일부 서식이 적용되지 않고, 원근 왜곡 때문에 " & _
                   "값을 잘못 읽게 됩니다. 평면 차트를 권합니다."
    End If
End Sub


Private Sub AddWarning(msg As String)
    If InStr(1, mWarnings, msg, vbTextCompare) > 0 Then Exit Sub
    If Len(mWarnings) > 0 Then mWarnings = mWarnings & vbLf & vbLf
    mWarnings = mWarnings & "- " & msg
End Sub


Private Sub ShowWarnings()
    If Len(mWarnings) = 0 Then Exit Sub
    MsgBox "서식을 적용했습니다. 확인해 볼 점:" & vbLf & vbLf & mWarnings, _
           vbInformation, "ChartFormatter"
    mWarnings = ""
End Sub
