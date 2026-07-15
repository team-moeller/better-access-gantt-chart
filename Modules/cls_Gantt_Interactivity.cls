VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Gantt_Interactivity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'###############################################################################################
'# Copyright (c) 2026 Thomas Möller                                                            #
'# MIT License  => https://github.com/team-moeller/better-access-gantt-chart/blob/main/LICENSE #
'# Version 2.11.01  published: 15.07.2026                                                      #
'###############################################################################################

Option Compare Database
Option Explicit


'### Members
Private m_ReactOnClick As Boolean
Private m_ReactOnProgressChange As Boolean
Private m_ReactOnDateChange As Boolean


'### Properties
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


'### Initialization
Private Sub Class_Initialize()

    Me.ReactOnClick = True
    Me.ReactOnProgressChange = True
    Me.ReactOnDateChange = True
    
End Sub

Private Sub Class_Terminate()
    ' Do Nothing
End Sub


'### JS Output
Public Function ToScript() As String

    'Variables
    Dim s As String

    If Me.ReactOnClick Then
        s = s & "            on_click: function(task) {" & vbCrLf
        s = s & "                if (window.AccessEdgeBridge !== null) { return; }" & vbCrLf
        s = s & "                window.AccessEdgeBridge = JSON.stringify({" & vbCrLf
        s = s & "                event: 'click'," & vbCrLf
        s = s & "                id: task.id " & vbCrLf
        s = s & "                });" & vbCrLf
        s = s & "            }," & vbCrLf
    End If

    If Me.ReactOnProgressChange Then
        s = s & "            on_progress_change: function(task, progress) {" & vbCrLf
        s = s & "                window.AccessEdgeBridge = JSON.stringify({" & vbCrLf
        s = s & "                event: 'progress'," & vbCrLf
        s = s & "                id: task.id , " & vbCrLf
        s = s & "                progress: progress " & vbCrLf
        s = s & "                });" & vbCrLf
        s = s & "            }," & vbCrLf
    End If

    If Me.ReactOnDateChange Then
        s = s & "            on_date_change: function(task, start, end) {" & vbCrLf
        s = s & "                window.AccessEdgeBridge = JSON.stringify({" & vbCrLf
        s = s & "                event: 'datechange'," & vbCrLf
        s = s & "                id: task.id , " & vbCrLf
        s = s & "                start: start.toISOString().substring(0, 10), " & vbCrLf
        s = s & "                end: end.toISOString().substring(0, 10) " & vbCrLf
        s = s & "                });" & vbCrLf
        s = s & "            }," & vbCrLf
    End If

    ToScript = s
    
End Function

