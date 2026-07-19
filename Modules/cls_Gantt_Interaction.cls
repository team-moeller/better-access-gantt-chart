VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Gantt_Interaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'###############################################################################################
'# Copyright (c) 2026 Thomas Möller                                                            #
'# MIT License  => https://github.com/team-moeller/better-access-gantt-chart/blob/main/LICENSE #
'# Version 2.14.01  published: 19.07.2026                                                      #
'###############################################################################################

Option Compare Database
Option Explicit


'### Members
Private m_ShowTodayButton As Boolean
Private m_ViewModeSelect As Boolean
Private m_ReadonlyProgress As Boolean
Private m_ReadonlyDates As Boolean
Private m_ReadonlyChart As Boolean
Private m_AutoMoveLabel As Boolean
Private m_MoveDependencies As Boolean
Private m_ShowPopup As Boolean
Private m_PopupOn As gcPopupOn


'### Properties
Public Property Let ShowTodayButton(newValue As Boolean)
    m_ShowTodayButton = newValue
End Property
Public Property Get ShowTodayButton() As Boolean
    ShowTodayButton = m_ShowTodayButton
End Property

Public Property Let ViewModeSelect(newValue As Boolean)
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

Public Property Let ShowPopup(newValue As Boolean)
    m_ShowPopup = newValue
End Property
Public Property Get ShowPopup() As Boolean
    ShowPopup = m_ShowPopup
End Property

Public Property Let PopupOn(newValue As gcPopupOn)
    m_PopupOn = newValue
End Property
Public Property Get PopupOn() As gcPopupOn
    PopupOn = m_PopupOn
End Property


'### Initialization
Private Sub Class_Initialize()

    Me.ShowTodayButton = True
    Me.ViewModeSelect = False
    Me.ReadonlyProgress = False
    Me.ReadonlyDates = False
    Me.ReadonlyChart = False
    Me.AutoMoveLabel = False
    Me.MoveDependencies = True
    Me.ShowPopup = False
    Me.PopupOn = gcPopupOn.poClick
    
End Sub

Private Sub Class_Terminate()
    ' Do Nothing
End Sub


'### JS Output
Public Function ToScript() As String

    'Variables
    Dim s As String

    If Not Me.ShowTodayButton Then s = s & "            today_button: false," & vbCrLf
    If Me.ViewModeSelect Then s = s & "            view_mode_select: true," & vbCrLf
    If Me.ReadonlyProgress Then s = s & "            readonly_progress: true," & vbCrLf
    If Me.ReadonlyDates Then s = s & "            readonly_dates: true," & vbCrLf
    If Me.ReadonlyChart Then s = s & "            readonly: true," & vbCrLf
    If Me.AutoMoveLabel Then s = s & "            auto_move_label: true," & vbCrLf
    If Not Me.MoveDependencies Then s = s & "            move_dependencies: false," & vbCrLf
    If Not Me.ShowPopup Then
        s = s & "            popup: false," & vbCrLf
    Else
        If Not Me.PopupOn = gcPopupOn.poClick Then
            s = s & "            popup_on: 'hover'," & vbCrLf
        End If
    End If
    
    If Len(s) > 0 Then s = Left$(s, Len(s) - 3) & vbCrLf

    ToScript = s
    
End Function

