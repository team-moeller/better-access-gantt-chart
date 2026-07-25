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
'# Version 2.18.01  published: 25.07.2026                                                      #
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
Public Function UpdateFunctionToScript() As String

    'Variables
    Dim s As String

    If 1 = 1 Then
        s = s & "        let lastScrollLeft = 0;" & vbCrLf
        s = s & "        let lastScrollTop = 0;" & vbCrLf
        s = s & "        function saveScrollPosition() {" & vbCrLf
        s = s & "            const container = document.querySelector('#gantt .gantt-container');" & vbCrLf
        s = s & "            if (!container) return;" & vbCrLf
        s = s & "            lastScrollLeft = container.scrollLeft;" & vbCrLf
        s = s & "            lastScrollTop = container.scrollTop;" & vbCrLf
        s = s & "        }" & vbCrLf
        s = s & "        function restoreScrollPosition() {" & vbCrLf
        s = s & "            const container = document.querySelector('#gantt .gantt-container');" & vbCrLf
        s = s & "            if (!container) return;" & vbCrLf
        s = s & "            container.scrollLeft = lastScrollLeft;" & vbCrLf
        s = s & "            container.scrollTop = lastScrollTop;" & vbCrLf
        s = s & "        }" & vbCrLf
        s = s & "        function updateTaskFromAccess(taskId, updatedValues) {" & vbCrLf
        s = s & "            saveScrollPosition();" & vbCrLf
        s = s & "            console.log('LiveSync (all attributes):', taskId, updatedValues);" & vbCrLf
        s = s & "            const task = tasks.find(t => t.id === taskId);" & vbCrLf
        s = s & "            if (!task) {" & vbCrLf
        s = s & "                console.error('Task not found:', taskId);" & vbCrLf
        s = s & "                return;" & vbCrLf
        s = s & "            }" & vbCrLf
        s = s & "            for (const key in updatedValues) {" & vbCrLf
        s = s & "                if (!updatedValues.hasOwnProperty(key)) continue;" & vbCrLf
        s = s & "                let value = updatedValues[key];" & vbCrLf
        s = s & "                if (key === 'start' || key === 'end') {" & vbCrLf
        s = s & "                    const d = new Date(value);" & vbCrLf
        s = s & "                    const local = new Date(d.getTime() + d.getTimezoneOffset() * 60000);" & vbCrLf
        s = s & "                    if (key === 'end') {" & vbCrLf
        s = s & "                        local.setHours(23, 59, 59, 999);"
        s = s & "                    }" & vbCrLf
        s = s & "                    value = local;" & vbCrLf
        s = s & "                }" & vbCrLf
        s = s & "                if (key === 'progress') {" & vbCrLf
        s = s & "                    value = parseFloat(value);" & vbCrLf
        s = s & "                }" & vbCrLf
        s = s & "                if (key === 'dependencies' && typeof value === 'string') {" & vbCrLf
        s = s & "                    value = value.split(',').map(v => v.trim());" & vbCrLf
        s = s & "                }" & vbCrLf
        s = s & "                task[key] = value;" & vbCrLf
        s = s & "            }" & vbCrLf
        s = s & "            gantt.refresh(tasks);" & vbCrLf
        s = s & "            requestAnimationFrame(restoreScrollPosition);" & vbCrLf
        s = s & "        }" & vbCrLf
    End If

    UpdateFunctionToScript = s
    
End Function

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
        s = s & "                const startLocal = new Date(start.getTime() - start.getTimezoneOffset() * 60000)" & vbCrLf
        s = s & "                    .toISOString().substring(0, 10);" & vbCrLf
        s = s & "                const endLocal = new Date(end.getTime() - end.getTimezoneOffset() * 60000)" & vbCrLf
        s = s & "                    .toISOString().substring(0, 10);" & vbCrLf
        s = s & "                window.AccessEdgeBridge = JSON.stringify({" & vbCrLf
        s = s & "                event: 'datechange'," & vbCrLf
        s = s & "                id: task.id ," & vbCrLf
        s = s & "                start: startLocal," & vbCrLf
        s = s & "                end: endLocal " & vbCrLf
        s = s & "                });" & vbCrLf
        s = s & "            }," & vbCrLf
    End If

    ToScript = s
    
End Function

