VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GanttChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type ActivityBar
    StartDate As Double
    EndDate As Double
    Color As Double
    DashStyle As MsoLineDashStyle
    BorderColor As Double
    Fill As MsoTriState
End Type

Private Type Activity
    Text As String
    Responsible As String
    Top As Double
    Height As Double
    ActivityBars() As ActivityBar
End Type

Private Type Callout
    Text As String
    CalloutDate As Date
    DashStyle As MsoLineDashStyle
End Type

Public Destination As Object

Public MaxWidth As Integer
Public MarginLeft As Integer
Public MarginTop As Integer
Public BarHeight As Double
Public YearHeader As Boolean
Public MonthHeader As Boolean
Public WeekHeader As Boolean
Public MonthLines As Boolean

Public LabelFontSize As Double
Public OutlineColor As Long
Public LabelFontColor As Long
Public HeaderFontColor As Long
Public ResponsibleColumn As Boolean

Public StartDate As Date
Public EndDate As Date

Private Activities() As Activity
Private NumActivities As Integer
Private Callouts() As Callout
Private NumCallouts As Integer
Private ShapeNames() As String
Private NumShapes As Integer

Private Labels() As Variant
Private Responsibles() As Variant
Private LabelWidth As Double
Private LabelHeight As Double
Private ResponsibleWidth As Double
Private BoxLeft As Double
Private BoxWidth As Double
Private NumHeaders As Integer
Private HeaderLabelWidth As Double
Private pHeaderBarHeight As Double
Private pHeaderFontSize As Double

Private Sub Class_Initialize()
    Set Destination = Application.ActiveWindow.View.Slide

    NumActivities = 0
    NumCallouts = 0
    
    MaxWidth = 850
    MarginLeft = 50
    MarginTop = 280
    BarHeight = 12
    YearHeader = True
    MonthHeader = True
    WeekHeader = True
    MonthLines = True
    
    OutlineColor = RGB(32, 32, 32)
    LabelFontColor = RGB(64, 64, 64)
    HeaderFontColor = RGB(64, 64, 64)
    
    LabelWidth = 0
    BoxLeft = 0
    BoxWidth = 0
    NumHeaders = 0
    NumShapes = 0
    LabelFontSize = 11
    HeaderFontSize = 9
    ResponsibleColumn = False
End Sub

Private Sub Class_Terminate()

End Sub

Public Property Let HeaderFontSize(size As Double)
    pHeaderFontSize = size
    pHeaderBarHeight = size + 10
End Property

Public Property Get HeaderFontSize() As Double
    HeaderFontSize = pHeaderFontSize
End Property

Private Sub AddShapeToList(Name)
    NumShapes = NumShapes + 1
    ReDim Preserve ShapeNames(0 To NumShapes)
    ShapeNames(NumShapes - 1) = Name
End Sub

Private Function CreateCenteredTextBox(x, y, Width, Height, Text, Optional TextColor = 0, Optional Border As Boolean = False)
    Dim lbl As Shape
    Set lbl = Destination.Shapes.AddLabel(msoTextOrientationHorizontal, x, y, Width, Height)
    lbl.TextFrame2.TextRange.Characters.Text = Text
    lbl.TextFrame2.WordWrap = False
    lbl.TextFrame2.AutoSize = msoAutoSizeNone
    lbl.TextFrame2.TextRange.Characters.Font.size = pHeaderFontSize
    lbl.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = TextColor
    lbl.TextFrame2.TextRange.Characters.ParagraphFormat.Alignment = msoAlignCenter
    lbl.TextFrame2.VerticalAnchor = msoAnchorMiddle
    lbl.Width = Width
    lbl.Height = Height
    
    If Border Then
        lbl.line.Visible = msoTrue
        lbl.line.ForeColor.RGB = OutlineColor
    End If
    
    AddShapeToList lbl.Name
    Set CreateCenteredTextBox = lbl
End Function

Sub CreateSquare(x, y, Width, Height, Text, Optional TextColor = 0)
    Destination.Shapes.AddShape(msoShapeRectangle, x, y, Width, Height).Select
    
    With Selection.ShapeRange
        .Fill.Visible = msoFalse
        .line.Visible = msoTrue
        .line.ForeColor.RGB = OutlineColor
    End With
    
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = Text
    With Selection.ShapeRange.TextFrame2.TextRange
        .Font.Fill.Visible = msoTrue
        .Font.Fill.ForeColor.RGB = TextColor
        .Font.Fill.Transparency = 0
        .Font.Fill.Solid
        .Font.size = pHeaderFontSize
        .Characters.ParagraphFormat.FirstLineIndent = 0
        .Characters.ParagraphFormat.Alignment = msoAlignCenter
    End With
    
    Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
    
    AddShapeToList Selection.Name
End Sub

Private Function CreateTextBox(x, y, Text, Optional TextColor = 0, Optional FontSize As Variant = 10)
    Dim lbl As Shape
    Set lbl = Destination.Shapes.AddLabel(msoTextOrientationHorizontal, x, y, 0, 0)
    lbl.TextFrame2.TextRange.Characters.Text = Text
    lbl.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
    lbl.TextFrame2.WordWrap = False
    lbl.TextFrame2.TextRange.Characters.Font.size = FontSize
    lbl.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = TextColor
    lbl.TextFrame2.VerticalAnchor = msoAnchorMiddle
    
    AddShapeToList lbl.Name
    Set CreateTextBox = lbl
End Function

Private Sub CreateBar(x, y, Width, Height, Fill As MsoTriState, FillColor, DashStyle, BorderColor)
    Dim bar As Shape
    Set bar = Destination.Shapes.AddShape(msoShapeRectangle, x, y, Width, Height)
    bar.Fill.Visible = Fill
    bar.Fill.ForeColor.RGB = FillColor
    
    If DashStyle = 0 Then
        bar.line.Visible = msoFalse
    Else
        bar.line.Visible = msoTrue
        bar.line.ForeColor.RGB = BorderColor
        bar.line.DashStyle = DashStyle
    End If

    AddShapeToList bar.Name
End Sub

Private Sub CreateLine(x1, y1, x2, y2, Optional DashStyle As MsoLineDashStyle = msoLineSolid)
    Dim line As Shape
    Set line = Destination.Shapes.AddConnector(msoConnectorStraight, x1, y1, x2, y2)
    With line.line
        .ForeColor.RGB = OutlineColor
        .Weight = 0.75
        .DashStyle = DashStyle
    End With
    
    line.ZOrder msoSendToBack
    AddShapeToList line.Name
End Sub

Private Function ConvertBorderStyleToDashStyle(LineStyle) As MsoLineDashStyle
    Select Case LineStyle
        Case xlContinuous
            ConvertBorderStyleToDashStyle = msoLineSolid
        Case xlDash
            ConvertBorderStyleToDashStyle = msoLineDash
        Case xlDashDot
            ConvertBorderStyleToDashStyle = msoLineDashDot
        Case xlDashDotDot
            ConvertBorderStyleToDashStyle = msoLineDashDotDot
        Case xlDot
            ConvertBorderStyleToDashStyle = msoLineRoundDot
        Case xlDouble
            ConvertBorderStyleToDashStyle = msoLineSolid
        Case xlLineStyleNone
            ConvertBorderStyleToDashStyle = 0
        'Case xlSlantDashDot
        '    ConvertBorderStyleToDashStyle = msoLineDashLong
        Case Else
            ConvertBorderStyleToDashStyle = 0
    End Select
End Function

Public Function AddActivity(Text, StartDate, EndDate, Optional Responsible = "", Optional Fill As MsoTriState = msoTrue, Optional Color = 0, Optional BorderStyle As XlLineStyle = xlContinuous, Optional BorderColor = -1)
    Dim NewActivity As Activity
    Dim ActivityBars(0) As ActivityBar
    
    If BorderColor = -1 Then
        BorderColor = Color
    End If
    
    NewActivity.Text = Text
    NewActivity.Responsible = Responsible
    ActivityBars(0).StartDate = StartDate
    ActivityBars(0).EndDate = EndDate
    ActivityBars(0).Color = Color
    ActivityBars(0).DashStyle = ConvertBorderStyleToDashStyle(BorderStyle)
    ActivityBars(0).BorderColor = BorderColor
    ActivityBars(0).Fill = Fill
    
    NewActivity.ActivityBars = ActivityBars
    
    NumActivities = NumActivities + 1
    ReDim Preserve Activities(0 To NumActivities - 1)
    Activities(NumActivities - 1) = NewActivity
    
    AddActivity = NumActivities - 1
End Function

Public Sub AddActivityBar(ActivityIndex As Integer, StartDate, EndDate, Optional Fill As MsoTriState = msoTrue, Optional Color = 0, Optional BorderStyle As XlLineStyle = xlContinuous, Optional BorderColor = -1)
    Dim n As Integer
    Dim ActivityBars() As ActivityBar
    ActivityBars = Activities(ActivityIndex).ActivityBars
    
    n = UBound(ActivityBars) + 1
    ReDim Preserve ActivityBars(n)
    
    If BorderColor = -1 Then
        BorderColor = Color
    End If
    
    ActivityBars(n).StartDate = StartDate
    ActivityBars(n).EndDate = EndDate
    ActivityBars(n).Color = Color
    ActivityBars(n).DashStyle = ConvertBorderStyleToDashStyle(BorderStyle)
    ActivityBars(n).BorderColor = BorderColor
    ActivityBars(n).Fill = Fill
    
    Activities(ActivityIndex).ActivityBars = ActivityBars
End Sub

Public Sub AddCallout(Text, CalloutDate, Optional DashStyle As MsoLineDashStyle = msoLineLongDash)
    Dim NewCallout As Callout
    NewCallout.Text = Text
    NewCallout.CalloutDate = CalloutDate
    NewCallout.DashStyle = msoLineLongDash
    
    NumCallouts = NumCallouts + 1
    ReDim Preserve Callouts(0 To NumCallouts - 1)
    Callouts(NumCallouts - 1) = NewCallout
End Sub

Private Sub CreateLabels()
    Dim y, i, H As Integer
    Dim Shape As Variant
    Dim ResShape As Variant
    Let y = MarginTop + pHeaderBarHeight * NumHeaders + 5
    
    ReDim Labels(0 To NumActivities)
    ReDim Responsibles(0 To NumActivities)
    For i = 0 To (NumActivities - 1)
        Set Shape = CreateTextBox(MarginLeft, y, Activities(i).Text, LabelFontColor, LabelFontSize)
        H = Shape.Height
        If Shape.Width > LabelWidth Then
            LabelWidth = Shape.Width
        End If
        Set Labels(i) = Shape
        
        ' Responsibles
        If ResponsibleColumn Then
            Set ResShape = CreateTextBox(500, y, Activities(i).Responsible, LabelFontColor, LabelFontSize)
            With ResShape.TextFrame.TextRange.ParagraphFormat.Bullet
                .Visible = True
                .Type = ppBulletUnnumbered
            End With
            If ResShape.Width > ResponsibleWidth Then
                ResponsibleWidth = ResShape.Width
            End If
            Set Responsibles(i) = ResShape
            
            If ResShape.Height > H Then
                H = ResShape.Height
                Shape.Height = H
            End If
        End If
        
        Activities(i).Top = y
        Activities(i).Height = H
        y = y + H + 2
    Next i
    
    ' Responsibles
    If ResponsibleColumn Then
        For i = 0 To (NumActivities - 1)
            Responsibles(i).Left = MarginLeft + MaxWidth - ResponsibleWidth
        Next i
    End If
    
    LabelHeight = y - (MarginTop + pHeaderBarHeight * NumHeaders + 15)
End Sub

Private Sub CreateBars()
    Dim i, j As Integer
    Dim StartValBox, EndValBox, BarTop, BarLeft, BarWidth As Double
    For i = 0 To (NumActivities - 1)
        For j = 0 To UBound(Activities(i).ActivityBars)
            StartValBox = Activities(i).ActivityBars(j).StartDate
            EndValBox = Activities(i).ActivityBars(j).EndDate
            BarTop = Activities(i).Top + Activities(i).Height / 2 - BarHeight / 2
            BarLeft = BoxLeft + ((StartValBox - StartDate) / (EndDate - StartDate)) * BoxWidth
            BarWidth = ((EndValBox - StartValBox) / (EndDate - StartDate)) * BoxWidth
            CreateBar BarLeft, BarTop, BarWidth, BarHeight, Activities(i).ActivityBars(j).Fill, Activities(i).ActivityBars(j).Color, Activities(i).ActivityBars(j).DashStyle, Activities(i).ActivityBars(j).BorderColor
        Next j
    Next i
End Sub

Private Sub CreateCallouts()
    Dim txt As Shape
    Dim i As Integer
    For i = 0 To (NumCallouts - 1)
        Dim p As Double
        p = BoxLeft + ((Callouts(i).CalloutDate - StartDate) / (EndDate - StartDate)) * BoxWidth
        CreateLine p, MarginTop + pHeaderBarHeight * NumHeaders, p, MarginTop + pHeaderBarHeight * NumHeaders + 15 + LabelHeight + 7, Callouts(i).DashStyle
        Set txt = CreateTextBox(MarginLeft, MarginTop + pHeaderBarHeight * NumHeaders + 15 + LabelHeight + 7, Callouts(i).Text)
        txt.Left = p - txt.Width / 2
    Next i
End Sub

Private Sub CreateYearHeader()
    Dim TotalDays, i, j, Position, Days, EndMonth, HeaderLabelWidth As Double
    
    TotalDays = EndDate - StartDate + 1
    j = 0
    i = StartDate
    Position = BoxLeft
    Do While i < EndDate
        EndMonth = CDbl(DateSerial(Year(i), 12, 31))
        If EndMonth > EndDate Then
            EndMonth = EndDate
        End If
        
        Days = EndMonth - i + 1
        HeaderLabelWidth = Days / TotalDays * BoxWidth
        
        CreateCenteredTextBox Position, MarginTop, HeaderLabelWidth, pHeaderBarHeight, Year(i), TextColor:=HeaderFontColor
        i = i + Days
        Position = Position + HeaderLabelWidth
    Loop
End Sub

Private Sub CreateMonthHeader(Optional offset = 0)
    Dim TotalDays, i, j, Position, Days, EndMonth, HeaderLabelWidth As Double
    
    TotalDays = EndDate - StartDate + 1
    j = 0
    i = StartDate
    Position = BoxLeft
    
    CreateLine BoxLeft, MarginTop + offset, BoxLeft + BoxWidth, MarginTop + offset
    
    Do While i < EndDate
        EndMonth = CDbl(DateSerial(Year(i), Month(i) + 1, 1)) - 1
        If EndMonth > EndDate Then
            EndMonth = EndDate
        End If
        
        Days = EndMonth - i + 1
        HeaderLabelWidth = Days / TotalDays * BoxWidth
        
        CreateCenteredTextBox Position, MarginTop + offset, HeaderLabelWidth, pHeaderBarHeight, MonthName(Month(i), 3), TextColor:=HeaderFontColor, Border:=True
        
        i = i + Days
        Position = Position + HeaderLabelWidth
    Loop
End Sub

Private Sub CreateWeekHeader(Optional offset = 0)
    Dim TotalDays, i, j, Position, Days, EndMonth, HeaderLabelWidth, LineLength, NextWeek As Double
    Dim iWeekNumber, iWeekDay As Integer
    
    LineLength = LabelHeight + pHeaderBarHeight * NumHeaders + 15
    
    TotalDays = EndDate - StartDate + 1
    j = 0
    i = StartDate
    Position = BoxLeft
        
    Do While i < EndDate
        iWeekNumber = DatePart("ww", i, vbMonday, vbFirstFourDays)
        iWeekDay = Weekday(i, vbMonday)
        
        NextWeek = i + 7 - iWeekDay
        If NextWeek > EndDate Then
            NextWeek = EndDate
        End If
        
        Days = NextWeek - i + 1
        HeaderLabelWidth = Days / TotalDays * BoxWidth
        
        CreateCenteredTextBox Position, MarginTop + offset, HeaderLabelWidth, pHeaderBarHeight, iWeekNumber, TextColor:=HeaderFontColor, Border:=True
        i = i + Days
        Position = Position + HeaderLabelWidth
    Loop
End Sub

Private Sub CreateHeader()
    Dim offset As Double
    offset = 0
    If YearHeader Then
        CreateYearHeader
        offset = pHeaderBarHeight
    End If
    If MonthHeader Then
        CreateMonthHeader offset
        offset = offset + pHeaderBarHeight
    End If
    If WeekHeader Then
        CreateWeekHeader offset
    End If
End Sub

Private Sub CreateLines()
    CreateLine MarginLeft, MarginTop + pHeaderBarHeight * NumHeaders, MarginLeft + MaxWidth, MarginTop + pHeaderBarHeight * NumHeaders
    CreateLine MarginLeft, MarginTop + LabelHeight + pHeaderBarHeight * NumHeaders + 15, MarginLeft + MaxWidth, MarginTop + LabelHeight + pHeaderBarHeight * NumHeaders + 15
    
    'Create start and end line
    Dim LineLength As Double
    LineLength = LabelHeight + pHeaderBarHeight * NumHeaders + 15
    CreateLine MarginLeft + LabelWidth + 15, MarginTop, MarginLeft + LabelWidth + 15, MarginTop + LineLength
    CreateLine MarginLeft + LabelWidth + BoxWidth + 15, MarginTop, MarginLeft + BoxWidth + LabelWidth + 15, MarginTop + LineLength
    
    'Create line every month
    Dim i, EndMonth, Days, TotalDays, Position As Double
    If MonthLines Then
        TotalDays = EndDate - StartDate + 1
        Position = BoxLeft
        i = StartDate
        Do While i <= EndDate
            EndMonth = CDbl(DateSerial(Year(i), Month(i) + 1, 1)) - 1
            If EndMonth > EndDate Then
                EndMonth = EndDate
            End If
            Days = EndMonth - i + 1
            
            HeaderLabelWidth = Days / TotalDays * BoxWidth
            i = i + Days
            Position = Position + HeaderLabelWidth
            
            If i < EndDate Then
                CreateLine Position, MarginTop + pHeaderBarHeight * NumHeaders, Position, MarginTop + pHeaderBarHeight * NumHeaders + LabelHeight + 15
            End If
        Loop
    End If
End Sub

Public Sub Generate()
    On Error Resume Next
    Destination.Shapes("GanttChart").Delete
    On Error GoTo 0

    On Error Resume Next
    If YearHeader Then
        NumHeaders = NumHeaders + 1
    End If
    If MonthHeader Then
        NumHeaders = NumHeaders + 1
    End If
    If WeekHeader Then
        NumHeaders = NumHeaders + 1
    End If

    CreateLabels
    
    BoxLeft = MarginLeft + LabelWidth + 15
    BoxWidth = MaxWidth - LabelWidth - 15 - ResponsibleWidth
    
    CreateBars
    If NumCallouts > 0 Then
        CreateCallouts
    End If
    
    CreateHeader
    CreateLines
    
    Dim Group As Shape
    Set Group = Destination.Shapes.Range(ShapeNames).Group
    Group.Name = "GanttChart"
    Group.Select
    
    On Error GoTo 0
End Sub
