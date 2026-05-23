VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Better_Access_Gantt_Chart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'###############################################################################################
'# Copyright (c) 2026 Thomas Möller                                                            #
'# MIT License  => https://github.com/team-moeller/better-access-gantt-chart/blob/main/LICENSE #
'# Version 1.34.06  published: 23.05.2026                                                      #
'###############################################################################################

Option Compare Database
Option Explicit


'### Enums

Public Enum gcScriptSource
    CDN = 1
    LocalFile = 2
End Enum

Public Enum gcViewMode
    vwHour = 1
    vwQuarterDay = 2
    vwHalfDay = 3
    vwDay = 4
    vwWeek = 5
    vwMonth = 6
    vwYear = 7
End Enum

Public Enum gcLines
    lnNone = 1
    lnVertical = 2
    lnHorizontal = 3
    lnBoth = 4
End Enum

Public Enum gcScrollTo
    stToday = 1
    stStart = 2
    stEnd = 3
    stDate = 4
End Enum

Public Enum gcTheme
    thDefault = 0
    thDarkMode = 1
    thBlue = 2
End Enum



'### Members

Private WithEvents m_Control As Edge
Attribute m_Control.VB_VarHelpID = -1
'Appearance
Private m_Language As String
Private m_ViewMode As gcViewMode
Private m_ColumnWidth As Long
Private m_BarHeight As Long
Private m_ArrowCurveRadius As Integer
Private m_TaskBarCurveRadius As Integer
Private m_ContainerHight As String
Private m_UpperHeaderHight As Integer
Private m_LowerHeaderHight As Integer
Private m_Padding As Integer
Private m_Lines As gcLines
Private m_ScrollTo As gcScrollTo
Private m_ScrollToDate As Date
Private m_ShowExpectedProgress As Boolean
'Interaction
Private m_ShowTodayButton As Boolean
Private m_ViewModeSelect As Boolean
Private m_ReadonlyProgress As Boolean
Private m_ReadonlyDates As Boolean
Private m_ReadonlyChart As Boolean
Private m_AutoMoveLabel As Boolean
Private m_MoveDependencies As Boolean
'Interactivity
Private m_ReactOnClick As Boolean
Private m_ReactOnProgressChange As Boolean
Private m_ReactOnDateChange As Boolean
'Color
Private m_Theme As gcTheme
Private m_ColorTaskBarColor As String
Private m_ColorBarBorderColor As String
Private m_ColorProgressColor As String
Private m_ColorArrowColor As String
Private m_ColorHandleColor As String
Private m_ColorTextDark As String
Private m_ColorTextLight As String
Private m_ColorTextMuted As String
Private m_ColorHeaderBackground As String
Private m_ColorRowColor As String
Private m_ColorRowBorderColor As String
Private m_ColorTodayHighlightColor As String
Private m_ColorWeekendHighlightColor As String
Private m_ColorTickColor As String
'Scripts
Private m_GanttJsSource As gcScriptSource
Private m_GanttCssSource As gcScriptSource
Private m_GanttJsCDN As String
Private m_GanttCssCDN As String
Private m_GanttJsPathFilename As String
Private m_GanttCssPathFilename As String



'### Properties

Public Property Set Control(newValue As Edge)
    Set m_Control = newValue
End Property
Public Property Get Control() As Edge
    Set Control = m_Control
End Property

'Appearance
Public Property Let Language(newValue As String)
    m_Language = newValue
End Property
Public Property Get Language() As String
    Language = m_Language
End Property

Public Property Let ViewMode(newValue As gcViewMode)
    m_ViewMode = newValue
End Property
Public Property Get ViewMode() As gcViewMode
    ViewMode = m_ViewMode
End Property

Public Property Let ColumnWidth(newValue As Long)
    m_ColumnWidth = newValue
End Property
Public Property Get ColumnWidth() As Long
    ColumnWidth = m_ColumnWidth
End Property

Public Property Let BarHeight(newValue As Long)
    m_BarHeight = newValue
End Property
Public Property Get BarHeight() As Long
    BarHeight = m_BarHeight
End Property

Public Property Let ArrowCurveRadius(newValue As Integer)
    m_ArrowCurveRadius = newValue
End Property

Public Property Get ArrowCurveRadius() As Integer
    ArrowCurveRadius = m_ArrowCurveRadius
End Property

Public Property Let TaskBarCurveRadius(newValue As Integer)
    m_TaskBarCurveRadius = newValue
End Property
Public Property Get TaskBarCurveRadius() As Integer
    TaskBarCurveRadius = m_TaskBarCurveRadius
End Property

Public Property Let ContainerHight(newValue As String)
    m_ContainerHight = newValue
End Property
Public Property Get ContainerHight() As String
    ContainerHight = m_ContainerHight
End Property

Public Property Let UpperHeaderHight(newValue As Integer)
    m_UpperHeaderHight = newValue
End Property
Public Property Get UpperHeaderHight() As Integer
    UpperHeaderHight = m_UpperHeaderHight
End Property

Public Property Let LowerHeaderHight(newValue As Integer)
    m_LowerHeaderHight = newValue
End Property
Public Property Get LowerHeaderHight() As Integer
    LowerHeaderHight = m_LowerHeaderHight
End Property

Public Property Let Padding(newValue As Integer)
    m_Padding = newValue
End Property
Public Property Get Padding() As Integer
    Padding = m_Padding
End Property

Public Property Let Lines(newValue As gcLines)
    m_Lines = newValue
End Property
Public Property Get Lines() As gcLines
    Lines = m_Lines
End Property

Public Property Let ScrollTo(newValue As gcScrollTo)
    m_ScrollTo = newValue
End Property
Public Property Get ScrollTo() As gcScrollTo
    ScrollTo = m_ScrollTo
End Property

Public Property Let ScrollToDate(newValue As Date)
    m_ScrollToDate = newValue
End Property
Public Property Get ScrollToDate() As Date
    ScrollToDate = m_ScrollToDate
End Property
Public Property Let ShowExpectedProgress(newValue As Boolean)
    m_ShowExpectedProgress = newValue
End Property
Public Property Get ShowExpectedProgress() As Boolean
    ShowExpectedProgress = m_ShowExpectedProgress
End Property

'Interaction
Public Property Let ShowTodayButton(newValue As Boolean)
    m_ShowTodayButton = newValue
End Property
Public Property Get ShowTodayButton() As Boolean
    ShowTodayButton = m_ShowTodayButton
End Property

Public Property Let ViewModeSelect(ByVal newValue As Boolean)
    m_ViewModeSelect = newValue
End Property
Public Property Get ViewModeSelect() As Boolean
    ViewModeSelect = m_ViewModeSelect
End Property

Public Property Let ReadonlyProgress(newValue As Boolean)
    m_ReadonlyProgress = newValue
End Property
Public Property Get ReadonlyProgress() As Boolean
    ReadonlyProgress = m_ReadonlyProgress
End Property

Public Property Let ReadonlyDates(newValue As Boolean)
    m_ReadonlyDates = newValue
End Property

Public Property Get ReadonlyDates() As Boolean
    ReadonlyDates = m_ReadonlyDates
End Property

Public Property Let ReadonlyChart(newValue As Boolean)
    m_ReadonlyChart = newValue
End Property
Public Property Get ReadonlyChart() As Boolean
    ReadonlyChart = m_ReadonlyChart
End Property

Public Property Let AutoMoveLabel(newValue As Boolean)
    m_AutoMoveLabel = newValue
End Property
Public Property Get AutoMoveLabel() As Boolean
    AutoMoveLabel = m_AutoMoveLabel
End Property

Public Property Let MoveDependencies(newValue As Boolean)
    m_MoveDependencies = newValue
End Property
Public Property Get MoveDependencies() As Boolean
    MoveDependencies = m_MoveDependencies
End Property

'Interactivity
Public Property Let ReactOnClick(newValue As Boolean)
    m_ReactOnClick = newValue
End Property
Public Property Get ReactOnClick() As Boolean
    ReactOnClick = m_ReactOnClick
End Property

Public Property Let ReactOnProgressChange(newValue As Boolean)
    m_ReactOnProgressChange = newValue
End Property
Public Property Get ReactOnProgressChange() As Boolean
    ReactOnProgressChange = m_ReactOnProgressChange
End Property

Public Property Let ReactOnDateChange(newValue As Boolean)
    m_ReactOnDateChange = newValue
End Property
Public Property Get ReactOnDateChange() As Boolean
    ReactOnDateChange = m_ReactOnDateChange
End Property

'Color
Public Property Let Theme(ByVal newValue As gcTheme)
    m_Theme = newValue
    Select Case m_Theme
        Case 0
            'Do Nothing
        Case 1
            Me.ColorTaskBarColor = "#3a3a3a"
            Me.ColorBarBorderColor = "#3a3a3a"
            Me.ColorProgressColor = "#4a6baf"
            Me.ColorArrowColor = "#9aa5b1"
            Me.ColorHandleColor = "#555"
            Me.ColorTextDark = "#e0e0e0"
            Me.ColorTextLight = "#f5f5f5"
            Me.ColorTextMuted = "#a0a0a0"
            Me.ColorHeaderBackground = "#1c1c1c"
            Me.ColorRowColor = "#252525"
            Me.ColorRowBorderColor = "#3d3d3d"
            Me.ColorTodayHighlightColor = "#ff9d48"
            Me.ColorWeekendHighlightColor = "#2d2d2d"
            Me.ColorTickColor = "#444"
        Case 2
            Me.ColorTaskBarColor = "#d7e8ff"
            Me.ColorBarBorderColor = "#a9c9f5"
            Me.ColorProgressColor = "#4a90e2"
            Me.ColorArrowColor = "#2f6fbf"
            Me.ColorHandleColor = "#3c5f8c"
            Me.ColorTextDark = "#1c2a3a"
            Me.ColorTextLight = "#f0f6ff"
            Me.ColorTextMuted = "#6b8bb5"
            Me.ColorHeaderBackground = "#e8f1ff"
            Me.ColorRowColor = "#f5f9ff"
            Me.ColorRowBorderColor = "#c7d9f2"
            Me.ColorTodayHighlightColor = "#4a90e2"
            Me.ColorWeekendHighlightColor = "#e1ecfa"
            Me.ColorTickColor = "#b4c9e6"
        Case Else
            'Do Nothing
    End Select
End Property
Public Property Get Theme() As gcTheme
    Theme = m_Theme
End Property

Public Property Let ColorTaskBarColor(ByVal newValue As String)
    m_ColorTaskBarColor = newValue
End Property
Public Property Get ColorTaskBarColor() As String
    ColorTaskBarColor = m_ColorTaskBarColor
End Property

Public Property Let ColorBarBorderColor(ByVal newValue As String)
    m_ColorBarBorderColor = newValue
End Property
Public Property Get ColorBarBorderColor() As String
    ColorBarBorderColor = m_ColorBarBorderColor
End Property

Public Property Let ColorProgressColor(ByVal newValue As String)
    m_ColorProgressColor = newValue
End Property
Public Property Get ColorProgressColor() As String
    ColorProgressColor = m_ColorProgressColor
End Property

Public Property Let ColorArrowColor(ByVal newValue As String)
    m_ColorArrowColor = newValue
End Property
Public Property Get ColorArrowColor() As String
    ColorArrowColor = m_ColorArrowColor
End Property

Public Property Let ColorHandleColor(ByVal newValue As String)
    m_ColorHandleColor = newValue
End Property
Public Property Get ColorHandleColor() As String
    ColorHandleColor = m_ColorHandleColor
End Property

Public Property Let ColorTextDark(ByVal newValue As String)
    m_ColorTextDark = newValue
End Property
Public Property Get ColorTextDark() As String
    ColorTextDark = m_ColorTextDark
End Property

Public Property Let ColorTextLight(ByVal newValue As String)
    m_ColorTextLight = newValue
End Property
Public Property Get ColorTextLight() As String
    ColorTextLight = m_ColorTextLight
End Property

Public Property Let ColorTextMuted(ByVal newValue As String)
    m_ColorTextMuted = newValue
End Property
Public Property Get ColorTextMuted() As String
    ColorTextMuted = m_ColorTextMuted
End Property

Public Property Let ColorHeaderBackground(ByVal newValue As String)
    m_ColorHeaderBackground = newValue
End Property
Public Property Get ColorHeaderBackground() As String
    ColorHeaderBackground = m_ColorHeaderBackground
End Property

Public Property Let ColorRowColor(ByVal newValue As String)
    m_ColorRowColor = newValue
End Property
Public Property Get ColorRowColor() As String
    ColorRowColor = m_ColorRowColor
End Property

Public Property Let ColorRowBorderColor(ByVal newValue As String)
    m_ColorRowBorderColor = newValue
End Property
Public Property Get ColorRowBorderColor() As String
    ColorRowBorderColor = m_ColorRowBorderColor
End Property

Public Property Let ColorTodayHighlightColor(ByVal newValue As String)
    m_ColorTodayHighlightColor = newValue
End Property
Public Property Get ColorTodayHighlightColor() As String
    ColorTodayHighlightColor = m_ColorTodayHighlightColor
End Property

Public Property Let ColorWeekendHighlightColor(ByVal newValue As String)
    m_ColorWeekendHighlightColor = newValue
End Property
Public Property Get ColorWeekendHighlightColor() As String
    ColorWeekendHighlightColor = m_ColorWeekendHighlightColor
End Property

Public Property Let ColorTickColor(ByVal newValue As String)
    m_ColorTickColor = newValue
End Property
Public Property Get ColorTickColor() As String
    ColorTickColor = m_ColorTickColor
End Property

'Scripts
Public Property Let GanttJsSource(ByVal newValue As gcScriptSource)
    m_GanttJsSource = newValue
End Property
Public Property Get GanttJsSource() As gcScriptSource
    GanttJsSource = m_GanttJsSource
End Property

Public Property Let GanttCssSource(ByVal newValue As gcScriptSource)
    m_GanttCssSource = newValue
End Property
Public Property Get GanttCssSource() As gcScriptSource
    GanttCssSource = m_GanttCssSource
End Property

Public Property Let GanttJsCDN(ByVal newValue As String)
    m_GanttJsCDN = newValue
End Property
Public Property Get GanttJsCDN() As String
    GanttJsCDN = m_GanttJsCDN
End Property

Public Property Let GanttCssCDN(ByVal newValue As String)
    m_GanttCssCDN = newValue
End Property
Public Property Get GanttCssCDN() As String
    GanttCssCDN = m_GanttCssCDN
End Property

Public Property Let GanttJsPathFilename(ByVal newValue As String)
    m_GanttJsPathFilename = newValue
End Property
Public Property Get GanttJsPathFilename() As String
    GanttJsPathFilename = m_GanttJsPathFilename
End Property

Public Property Let GanttCssPathFilename(ByVal newValue As String)
    m_GanttCssPathFilename = newValue
End Property
Public Property Get GanttCssPathFilename() As String
    GanttCssPathFilename = m_GanttCssPathFilename
End Property



'### Instantiation

Private Sub Class_Initialize()

    'Appearance
    Me.Language = "en"
    Me.ColumnWidth = 45
    Me.BarHeight = 30
    Me.ArrowCurveRadius = 5
    Me.TaskBarCurveRadius = 3
    Me.ContainerHight = "auto"
    Me.UpperHeaderHight = 45
    Me.LowerHeaderHight = 30
    Me.Padding = 18
    Me.Lines = gcLines.lnBoth
    Me.ScrollTo = gcScrollTo.stToday
    Me.ScrollToDate = Format$(Date, "YYYY-MM-DD")
    Me.ShowExpectedProgress = False
    'Interaction
    Me.ShowTodayButton = True
    Me.ViewModeSelect = False
    Me.ReadonlyProgress = False
    Me.ReadonlyDates = False
    Me.ReadonlyChart = False
    Me.AutoMoveLabel = False
    Me.MoveDependencies = True
    'Interactivity
    Me.ReactOnClick = True
    Me.ReactOnProgressChange = True
    Me.ReactOnDateChange = True
    'Color
    Me.Theme = gcTheme.thDefault
    'Scripts
    Me.GanttJsSource = gcScriptSource.CDN
    Me.GanttJsCDN = "https://cdn.jsdelivr.net/npm/frappe-gantt/dist/frappe-gantt.umd.js"
    Me.GanttJsPathFilename = CurrentProject.Path & "\frappe-gantt.umd.js"
    Me.GanttCssSource = gcScriptSource.CDN
    Me.GanttCssCDN = "https://cdn.jsdelivr.net/npm/frappe-gantt/dist/frappe-gantt.css"
    Me.GanttCssPathFilename = CurrentProject.Path & "\frappe-gantt.css"
    
End Sub

Private Sub Class_Terminate()
    ' Do Nothing
End Sub



'### Public Methods

Public Sub ShowGanttChart()

    Dim Datei As String
    
    CreateHtmlFile
       
    Datei = "https://msaccess/" & CurrentProject.Path & "\BetterAccessGanttChart.html"
    
    Control.Navigate Datei
    
End Sub

Public Sub ResetBridgeVariable()

    'Variables
    Dim Script As String
    
    Script = "window.AccessEdgeBridge = null;"
    Me.Control.ExecuteJavascript Script

End Sub

Public Function RetrieveBridgeVariable() As String

    'Variables
    Dim Value As String
    
    Value = Me.Control.RetrieveJavascriptValue("window.AccessEdgeBridge")

    RetrieveBridgeVariable = Value

End Function



'### Private Methods

Private Sub CreateHtmlFile()

    'Variables
    Dim d As Integer
    Dim strPath As String
    Dim strFilename As String
    Dim strContent As String
    Const ForWriting As Long = 2
    
    strPath = CurrentProject.Path
    strFilename = "\BetterAccessGanttChart.html"
    
    strContent = strContent & getHead
    strContent = strContent & getScript
    strContent = strContent & getFoot
 
    d = FreeFile()
    Open strPath & strFilename For Output As d
    Print #d, strContent;
 
    Close d
  
End Sub

Private Function getHead() As String

    'Variables
    Dim strText As String
    
    strText = strText & "<!DOCTYPE html>" & vbCrLf
    strText = strText & "<html lang='en'>" & vbCrLf
    strText = strText & "<head>" & vbCrLf
    strText = strText & "    <meta charset='UTF-8'>" & vbCrLf
    strText = strText & "    <meta http-equiv='X-UA-Compatible' content='IE=Edge'/>" & vbCrLf
    strText = strText & "    <title>Better Access Gantt Chart</title>" & vbCrLf
    strText = strText & "    <style>body{margin:0;padding:0;}</style>" & vbCrLf
    strText = strText & "    <script src=" & getLocationJs & "></script>" & vbCrLf
    strText = strText & "    <link rel='stylesheet' href=" & getLocationCss & "/>" & vbCrLf
    strText = strText & getCustomColors
    strText = strText & "</head>" & vbCrLf
    strText = strText & "<body>" & vbCrLf
    
    getHead = strText
    
End Function

Private Function getScript() As String

    'Variables
    Dim strText As String
    
    strText = strText & "    <h2>Demo project</h2>" & vbCrLf
    strText = strText & "    <div id='gantt'></div>" & vbCrLf
    
    strText = strText & "    <script>" & vbCrLf
    strText = strText & "        window.AccessEdgeBridge = null;" & vbCrLf
    strText = strText & getTaskData
    strText = strText & "        const gantt = new Gantt('#gantt', tasks, {" & vbCrLf
    strText = strText & getInteractivity
    strText = strText & getAppearance
    strText = strText & getInteraction
    strText = strText & "        });" & vbCrLf
    strText = strText & "    </script>" & vbCrLf
    
    getScript = strText

End Function

Private Function getFoot() As String

    'Variables
    Dim strText As String
    
    strText = strText & "</body>" & vbCrLf
    strText = strText & "</html>"
    
    getFoot = strText
    
End Function

Private Function getCustomColors() As String

    'Variables
    Dim strText As String

    strText = strText & "    <style>" & vbCrLf
    strText = strText & "        /* Custom colors */" & vbCrLf
    strText = strText & "        .bar-red .bar {fill: #e74c3c !important;}" & vbCrLf
    strText = strText & "        .bar-green .bar {fill: #2ecc71 !important;}" & vbCrLf
    strText = strText & "        :root {" & vbCrLf
    If Len(Nz(Me.ColorTaskBarColor, vbNullString)) > 0 Then
        strText = strText & "            --g-bar-color: " & Me.ColorTaskBarColor & ";" & vbCrLf
    End If
    If Len(Nz(Me.ColorBarBorderColor, vbNullString)) > 0 Then
        strText = strText & "            --g-bar-border: " & Me.ColorBarBorderColor & ";" & vbCrLf
    End If
    If Len(Nz(Me.ColorProgressColor, vbNullString)) > 0 Then
        strText = strText & "            --g-progress-color: " & Me.ColorProgressColor & ";" & vbCrLf
    End If
    If Len(Nz(Me.ColorArrowColor, vbNullString)) > 0 Then
        strText = strText & "            --g-arrow-color: " & Me.ColorArrowColor & ";" & vbCrLf
    End If
    If Len(Nz(Me.ColorHandleColor, vbNullString)) > 0 Then
        strText = strText & "            --g-handle-color: " & Me.ColorHandleColor & ";" & vbCrLf
    End If
    If Len(Nz(Me.ColorTextDark, vbNullString)) > 0 Then
        strText = strText & "            --g-text-dark: " & Me.ColorTextDark & ";" & vbCrLf
    End If
    If Len(Nz(Me.ColorTextLight, vbNullString)) > 0 Then
        strText = strText & "            --g-text-light: " & Me.ColorTextLight & ";" & vbCrLf
    End If
    If Len(Nz(Me.ColorTextMuted, vbNullString)) > 0 Then
        strText = strText & "            --g-text-muted: " & Me.ColorTextMuted & ";" & vbCrLf
    End If
    If Len(Nz(Me.ColorHeaderBackground, vbNullString)) > 0 Then
        strText = strText & "            --g-header-background: " & Me.ColorHeaderBackground & ";" & vbCrLf
    End If
    If Len(Nz(Me.ColorRowColor, vbNullString)) > 0 Then
        strText = strText & "            --g-row-color: " & Me.ColorRowColor & ";" & vbCrLf
    End If
    If Len(Nz(Me.ColorRowBorderColor, vbNullString)) > 0 Then
        strText = strText & "            --g-row-border-color: " & Me.ColorRowBorderColor & ";" & vbCrLf
    End If
    If Len(Nz(Me.ColorTodayHighlightColor, vbNullString)) > 0 Then
        strText = strText & "            --g-today-highlight: " & Me.ColorTodayHighlightColor & ";" & vbCrLf
    End If
    If Len(Nz(Me.ColorWeekendHighlightColor, vbNullString)) > 0 Then
        strText = strText & "            --g-weekend-highlight-color: " & Me.ColorWeekendHighlightColor & ";" & vbCrLf
    End If
    If Len(Nz(Me.ColorTickColor, vbNullString)) > 0 Then
        strText = strText & "            --g-tick-color: " & Me.ColorTickColor & ";" & vbCrLf
    End If
    strText = strText & "        }" & vbCrLf
    strText = strText & "    </style>" & vbCrLf

    getCustomColors = strText

End Function

Private Function getTaskData() As String

    'Variables
    Dim db As DAO.Database
    Dim rst As DAO.Recordset
    Dim strSQL As String
    Dim strText As String
    
    Set db = CurrentDb()
    strSQL = "Select * from tbl_DemoData"
    Set rst = db.OpenRecordset(strSQL, dbOpenSnapshot)
    
    strText = strText & "        const tasks = [" & vbCrLf
    
    Do While Not rst.EOF
        strText = strText & "            {"
        strText = strText & "id: '" & rst!TID & "', "
        strText = strText & "name: '" & rst!TName & "', "
        strText = strText & "start: '" & Format$(rst!TStart, "YYYY-MM-DD") & "', "
        strText = strText & "end: '" & Format$(rst!TEnd, "YYYY-MM-DD") & "', "
        strText = strText & "progress: '" & rst!TProgress & "', "
        If Trim$(Nz(rst!TDependencies, vbNullString)) <> vbNullString Then
            strText = strText & "dependencies: '" & rst!TDependencies & "', "
        End If
        If Trim$(Nz(rst!TCustomClass, vbNullString)) <> vbNullString Then
            strText = strText & "custom_class: '" & rst!TCustomClass & "'"
        End If
        'Cut off the final comma
        If Right$(strText, 2) = ", " Then
            strText = Left$(strText, Len(strText) - 2)
        End If
        strText = strText & "}," & vbCrLf
        rst.MoveNext
    Loop
    
    strText = Left$(strText, Len(strText) - 3) & vbCrLf
    strText = strText & "        ];" & vbCrLf
    
    getTaskData = strText
    
    rst.Close
    Set rst = Nothing
    Set db = Nothing
    
End Function

Private Function getInteractivity() As String

    'Variables
    Dim strText As String
    
    If Me.ReactOnClick = True Then
        strText = strText & "            on_click: function(task) {" & vbCrLf
        strText = strText & "                if (window.AccessEdgeBridge !== null) {" & vbCrLf
        strText = strText & "                    return;" & vbCrLf
        strText = strText & "                }" & vbCrLf
        strText = strText & "                window.AccessEdgeBridge = task.id;" & vbCrLf
        strText = strText & "            }," & vbCrLf
    End If
    If Me.ReactOnProgressChange = True Then
        strText = strText & "            on_progress_change: function(task, progress) {" & vbCrLf
        strText = strText & "                window.AccessEdgeBridge = task.id + '|' + progress;" & vbCrLf
        strText = strText & "            }," & vbCrLf
    End If
    If Me.ReactOnDateChange = True Then
        strText = strText & "            on_date_change: function(task, start, end) {" & vbCrLf
        strText = strText & "                window.AccessEdgeBridge = task.id + '|' + start + '|' + end;" & vbCrLf
        strText = strText & "            }," & vbCrLf
    End If

    getInteractivity = strText

End Function

Private Function getAppearance() As String

    'Variables
    Dim strText As String

    strText = strText & "            view_mode: " & getViewMode & "," & vbCrLf
    strText = strText & "            language: '" & Me.Language & "'," & vbCrLf
    strText = strText & "            column_width: " & Me.ColumnWidth & "," & vbCrLf
    strText = strText & "            bar_height: " & Me.BarHeight & "," & vbCrLf
    strText = strText & "            arrow_curve: " & Me.ArrowCurveRadius & "," & vbCrLf
    strText = strText & "            bar_corner_radius: " & Me.TaskBarCurveRadius & "," & vbCrLf
    If Me.ContainerHight <> "auto" Then
        strText = strText & "            container_height: " & Me.ContainerHight & "," & vbCrLf
    End If
    strText = strText & "            upper_header_height: " & Me.UpperHeaderHight & "," & vbCrLf
    strText = strText & "            lower_header_height: " & Me.LowerHeaderHight & "," & vbCrLf
    strText = strText & "            padding: " & Me.Padding & "," & vbCrLf
    strText = strText & "            lines: " & getLines & "," & vbCrLf
    strText = strText & "            scroll_to: " & getScrollTo & "," & vbCrLf
    If Me.ShowExpectedProgress = True Then
        strText = strText & "            show_expected_progress: true" & "," & vbCrLf
    End If
    
    strText = Left$(strText, Len(strText) - 3) & vbCrLf
    
    getAppearance = strText

End Function

Private Function getInteraction() As String

    'Variables
    Dim strText As String
    
    If Me.ShowTodayButton = False Then
        strText = strText & "            today_button: false," & vbCrLf
    End If
    If Me.ViewModeSelect = True Then
        strText = strText & "            view_mode_select: true" & "," & vbCrLf
    End If
    If Me.ReadonlyProgress = True Then
        strText = strText & "            readonly_progress: true" & "," & vbCrLf
    End If
    If Me.ReadonlyDates = True Then
        strText = strText & "            readonly_dates: true" & "," & vbCrLf
    End If
    If Me.ReadonlyChart = True Then
        strText = strText & "            readonly: true" & "," & vbCrLf
    End If
    If Me.AutoMoveLabel = True Then
        strText = strText & "            auto_move_label: true" & "," & vbCrLf
    End If
    If Me.MoveDependencies = False Then
        strText = strText & "            move_dependencies: false" & "," & vbCrLf
    End If
    
    If Len(strText) > 0 Then
        strText = Left$(strText, Len(strText) - 3) & vbCrLf
    End If
    
    getInteraction = strText

End Function

Private Function getLocationJs() As String

    'Declarations
    Dim strText As String
    
    Select Case Me.GanttJsSource
        Case gcScriptSource.CDN
            strText = "'" & Me.GanttJsCDN & "'"
        Case gcScriptSource.LocalFile
            strText = "'" & "https://msaccess/" & Replace(Me.GanttJsPathFilename, "\", "/") & "'"
        Case Else
            'Do nothing
    End Select
            
    getLocationJs = strText

End Function

Private Function getLocationCss() As String

    'Declarations
    Dim strText As String
    
    Select Case Me.GanttCssSource
        Case gcScriptSource.CDN
            strText = "'" & Me.GanttCssCDN & "'"
        Case gcScriptSource.LocalFile
            strText = "'" & "https://msaccess/" & Replace(Me.GanttCssPathFilename, "\", "/") & "'"
        Case Else
            'Do nothing
    End Select
            
    getLocationCss = strText

End Function

Private Function getViewMode() As String

    'Declarations
    Dim strText As String

    Select Case Me.ViewMode
        Case gcViewMode.vwHour
            strText = "Hour"
        Case gcViewMode.vwQuarterDay
            strText = "Quarter Day"
        Case gcViewMode.vwHalfDay
            strText = "Half Day"
        Case gcViewMode.vwDay
            strText = "Day"
        Case gcViewMode.vwWeek
            strText = "Week"
        Case gcViewMode.vwMonth
            strText = "Month"
        Case gcViewMode.vwYear
            strText = "Year"
        Case Else
            strText = "Day"
    End Select

    strText = "'" & strText & "'"
        
    getViewMode = strText

End Function

Private Function getLines() As String

    'Declarations
    Dim strText As String

    Select Case Me.Lines
        Case gcLines.lnNone
            strText = "none"
        Case gcLines.lnVertical
            strText = "vertical"
        Case gcLines.lnHorizontal
            strText = "horizontal"
        Case gcLines.lnBoth
            strText = "both"
        Case Else
            strText = "both"
    End Select

    strText = "'" & strText & "'"
        
    getLines = strText

End Function

Private Function getScrollTo() As String

    'Declarations
    Dim strText As String

    Select Case Me.ScrollTo
        Case gcScrollTo.stToday
            strText = "today"
        Case gcScrollTo.stStart
            strText = "start"
        Case gcScrollTo.stEnd
            strText = "end"
        Case gcScrollTo.stDate
            strText = Format$(Me.ScrollToDate, "YYYY-MM-DD")
        Case Else
            strText = "today"
    End Select

    strText = "'" & strText & "'"
        
    getScrollTo = strText

End Function

