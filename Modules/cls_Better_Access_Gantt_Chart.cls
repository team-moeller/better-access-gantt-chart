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
'# Version 1.31.01  published: 18.05.2026                                                      #
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
'Theming
Private m_ThemingTaskBarColor As String
Private m_ThemingBarBorderColor As String
Private m_ThemingProgressColor As String
Private m_ThemingArrowColor As String
Private m_ThemingHandleColor As String
Private m_ThemingTextDark As String
Private m_ThemingTextLight As String
Private m_ThemingTextMuted As String
Private m_ThemingHeaderBackground As String
Private m_ThemingRowColor As String
Private m_ThemingRowBorderColor As String
Private m_ThemingTodayHighlightColor As String
Private m_ThemingWeekendHighlightColor As String
Private m_ThemingTickColor As String
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

'Theming
Public Property Let ThemingTaskBarColor(ByVal newValue As String)
    m_ThemingTaskBarColor = newValue
End Property
Public Property Get ThemingTaskBarColor() As String
    ThemingTaskBarColor = m_ThemingTaskBarColor
End Property

Public Property Let ThemingBarBorderColor(ByVal newValue As String)
    m_ThemingBarBorderColor = newValue
End Property
Public Property Get ThemingBarBorderColor() As String
    ThemingBarBorderColor = m_ThemingBarBorderColor
End Property

Public Property Let ThemingProgressColor(ByVal newValue As String)
    m_ThemingProgressColor = newValue
End Property
Public Property Get ThemingProgressColor() As String
    ThemingProgressColor = m_ThemingProgressColor
End Property

Public Property Let ThemingArrowColor(ByVal newValue As String)
    m_ThemingArrowColor = newValue
End Property
Public Property Get ThemingArrowColor() As String
    ThemingArrowColor = m_ThemingArrowColor
End Property

Public Property Let ThemingHandleColor(ByVal newValue As String)
    m_ThemingHandleColor = newValue
End Property
Public Property Get ThemingHandleColor() As String
    ThemingHandleColor = m_ThemingHandleColor
End Property

Public Property Let ThemingTextDark(ByVal newValue As String)
    m_ThemingTextDark = newValue
End Property
Public Property Get ThemingTextDark() As String
    ThemingTextDark = m_ThemingTextDark
End Property

Public Property Let ThemingTextLight(ByVal newValue As String)
    m_ThemingTextLight = newValue
End Property
Public Property Get ThemingTextLight() As String
    ThemingTextLight = m_ThemingTextLight
End Property

Public Property Let ThemingTextMuted(ByVal newValue As String)
    m_ThemingTextMuted = newValue
End Property
Public Property Get ThemingTextMuted() As String
    ThemingTextMuted = m_ThemingTextMuted
End Property

Public Property Let ThemingHeaderBackground(ByVal newValue As String)
    m_ThemingHeaderBackground = newValue
End Property
Public Property Get ThemingHeaderBackground() As String
    ThemingHeaderBackground = m_ThemingHeaderBackground
End Property

Public Property Let ThemingRowColor(ByVal newValue As String)
    m_ThemingRowColor = newValue
End Property
Public Property Get ThemingRowColor() As String
    ThemingRowColor = m_ThemingRowColor
End Property

Public Property Let ThemingRowBorderColor(ByVal newValue As String)
    m_ThemingRowBorderColor = newValue
End Property
Public Property Get ThemingRowBorderColor() As String
    ThemingRowBorderColor = m_ThemingRowBorderColor
End Property

Public Property Let ThemingTodayHighlightColor(ByVal newValue As String)
    m_ThemingTodayHighlightColor = newValue
End Property
Public Property Get ThemingTodayHighlightColor() As String
    ThemingTodayHighlightColor = m_ThemingTodayHighlightColor
End Property

Public Property Let ThemingWeekendHighlightColor(ByVal newValue As String)
    m_ThemingWeekendHighlightColor = newValue
End Property
Public Property Get ThemingWeekendHighlightColor() As String
    ThemingWeekendHighlightColor = m_ThemingWeekendHighlightColor
End Property

Public Property Let ThemingTickColor(ByVal newValue As String)
    m_ThemingTickColor = newValue
End Property
Public Property Get ThemingTickColor() As String
    ThemingTickColor = m_ThemingTickColor
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
    If Len(Nz(Me.ThemingTaskBarColor, vbNullString)) > 0 Then
        strText = strText & "            --g-bar-color: " & Me.ThemingTaskBarColor & ";" & vbCrLf
    End If
    If Len(Nz(Me.ThemingBarBorderColor, vbNullString)) > 0 Then
        strText = strText & "            --g-bar-border: " & Me.ThemingBarBorderColor & ";" & vbCrLf
    End If
    If Len(Nz(Me.ThemingProgressColor, vbNullString)) > 0 Then
        strText = strText & "            --g-progress-color: " & Me.ThemingProgressColor & ";" & vbCrLf
    End If
    If Len(Nz(Me.ThemingArrowColor, vbNullString)) > 0 Then
        strText = strText & "            --g-arrow-color: " & Me.ThemingArrowColor & ";" & vbCrLf
    End If
    If Len(Nz(Me.ThemingHandleColor, vbNullString)) > 0 Then
        strText = strText & "            --g-handle-color: " & Me.ThemingHandleColor & ";" & vbCrLf
    End If
    If Len(Nz(Me.ThemingTextDark, vbNullString)) > 0 Then
        strText = strText & "            --g-text-dark: " & Me.ThemingTextDark & ";" & vbCrLf
    End If
    If Len(Nz(Me.ThemingTextLight, vbNullString)) > 0 Then
        strText = strText & "            --g-text-light: " & Me.ThemingTextLight & ";" & vbCrLf
    End If
    If Len(Nz(Me.ThemingTextMuted, vbNullString)) > 0 Then
        strText = strText & "            --g-text-muted: " & Me.ThemingTextMuted & ";" & vbCrLf
    End If
    If Len(Nz(Me.ThemingHeaderBackground, vbNullString)) > 0 Then
        strText = strText & "            --g-header-background: " & Me.ThemingHeaderBackground & ";" & vbCrLf
    End If
    If Len(Nz(Me.ThemingRowColor, vbNullString)) > 0 Then
        strText = strText & "            --g-row-color: " & Me.ThemingRowColor & ";" & vbCrLf
    End If
    If Len(Nz(Me.ThemingRowBorderColor, vbNullString)) > 0 Then
        strText = strText & "            --g-row-border-color: " & Me.ThemingRowBorderColor & ";" & vbCrLf
    End If
    If Len(Nz(Me.ThemingTodayHighlightColor, vbNullString)) > 0 Then
        strText = strText & "            --g-today-highlight: " & Me.ThemingTodayHighlightColor & ";" & vbCrLf
    End If
    If Len(Nz(Me.ThemingWeekendHighlightColor, vbNullString)) > 0 Then
        strText = strText & "            --g-weekend-highlight-color: " & Me.ThemingWeekendHighlightColor & ";" & vbCrLf
    End If
    If Len(Nz(Me.ThemingTickColor, vbNullString)) > 0 Then
        strText = strText & "            --g-tick-color: " & Me.ThemingTickColor & ";" & vbCrLf
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

