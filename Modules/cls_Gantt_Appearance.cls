VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Gantt_Appearance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'###############################################################################################
'# Copyright (c) 2026 Thomas Möller                                                            #
'# MIT License  => https://github.com/team-moeller/better-access-gantt-chart/blob/main/LICENSE #
'# Version 2.03.03  published: 05.06.2026                                                      #
'###############################################################################################

Option Compare Database
Option Explicit


'### Members
Private m_Language As String
Private m_ViewMode As gcViewMode
Private m_ColumnWidth As Long
Private m_BarHeight As Long
Private m_ArrowCurveRadius As Integer
Private m_TaskBarCurveRadius As Integer
Private m_ContainerHeight As String
Private m_UpperHeaderHeight As Integer
Private m_LowerHeaderHeight As Integer
Private m_Padding As Integer
Private m_Lines As gcLines
Private m_ScrollTo As gcScrollTo
Private m_ScrollToDate As Date
Private m_ShowExpectedProgress As Boolean


'### Properties
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

Public Property Let ContainerHeight(newValue As String)
    m_ContainerHeight = newValue
End Property
Public Property Get ContainerHeight() As String
    ContainerHeight = m_ContainerHeight
End Property

Public Property Let UpperHeaderHeight(newValue As Integer)
    m_UpperHeaderHeight = newValue
End Property
Public Property Get UpperHeaderHeight() As Integer
    UpperHeaderHeight = m_UpperHeaderHeight
End Property

Public Property Let LowerHeaderHeight(newValue As Integer)
    m_LowerHeaderHeight = newValue
End Property
Public Property Get LowerHeaderHeight() As Integer
    LowerHeaderHeight = m_LowerHeaderHeight
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


'### Initialization
Private Sub Class_Initialize()

    Me.Language = "en"
    Me.ViewMode = gcViewMode.vwDay
    Me.ColumnWidth = 45
    Me.BarHeight = 30
    Me.ArrowCurveRadius = 5
    Me.TaskBarCurveRadius = 3
    Me.ContainerHeight = "auto"
    Me.UpperHeaderHeight = 45
    Me.LowerHeaderHeight = 30
    Me.Padding = 18
    Me.Lines = gcLines.lnBoth
    Me.ScrollTo = gcScrollTo.stToday
    Me.ScrollToDate = Date
    Me.ShowExpectedProgress = False
    
End Sub

Private Sub Class_Terminate()
    ' Do Nothing
End Sub


'### HTML/JS Output
Public Function ToScript() As String

    'Variables
    Dim s As String

    s = s & "            view_mode: " & getViewMode() & "," & vbCrLf
    s = s & "            language: '" & Me.Language & "'," & vbCrLf
    s = s & "            column_width: " & Me.ColumnWidth & "," & vbCrLf
    s = s & "            bar_height: " & Me.BarHeight & "," & vbCrLf
    s = s & "            arrow_curve: " & Me.ArrowCurveRadius & "," & vbCrLf
    s = s & "            bar_corner_radius: " & Me.TaskBarCurveRadius & "," & vbCrLf

    If Me.ContainerHeight <> "auto" Then
        s = s & "            container_height: " & Me.ContainerHeight & "," & vbCrLf
    End If

    s = s & "            upper_header_height: " & Me.UpperHeaderHeight & "," & vbCrLf
    s = s & "            lower_header_height: " & Me.LowerHeaderHeight & "," & vbCrLf
    s = s & "            padding: " & Me.Padding & "," & vbCrLf
    s = s & "            lines: " & getLines() & "," & vbCrLf
    s = s & "            scroll_to: " & getScrollTo() & "," & vbCrLf

    If Me.ShowExpectedProgress Then
        s = s & "            show_expected_progress: true," & vbCrLf
    End If

    s = Left$(s, Len(s) - 3) & vbCrLf
    ToScript = s
    
End Function

Private Function getViewMode() As String

    Select Case Me.ViewMode
        Case vwHour: getViewMode = "'Hour'"
        Case vwQuarterDay: getViewMode = "'Quarter Day'"
        Case vwHalfDay: getViewMode = "'Half Day'"
        Case vwDay: getViewMode = "'Day'"
        Case vwWeek: getViewMode = "'Week'"
        Case vwMonth: getViewMode = "'Month'"
        Case vwYear: getViewMode = "'Year'"
        Case Else: getViewMode = "'Day'"
    End Select
    
End Function

Private Function getLines() As String

    Select Case Me.Lines
        Case lnNone: getLines = "'none'"
        Case lnVertical: getLines = "'vertical'"
        Case lnHorizontal: getLines = "'horizontal'"
        Case lnBoth: getLines = "'both'"
        Case Else: getLines = "'both'"
    End Select
    
End Function

Private Function getScrollTo() As String

    Select Case Me.ScrollTo
        Case stToday: getScrollTo = "'today'"
        Case stStart: getScrollTo = "'start'"
        Case stEnd: getScrollTo = "'end'"
        Case stDate: getScrollTo = "'" & Format$(Me.ScrollToDate, "YYYY-MM-DD") & "'"
        Case Else: getScrollTo = "'today'"
    End Select
    
End Function

