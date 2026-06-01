VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Gantt_Scripts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'###############################################################################################
'# Copyright (c) 2026 Thomas Möller                                                            #
'# MIT License  => https://github.com/team-moeller/better-access-gantt-chart/blob/main/LICENSE #
'# Version 2.00.09  published: 01.06.2026                                                      #
'###############################################################################################

Option Compare Database
Option Explicit


'### Members
Private m_JS As cls_Gantt_Script_JS
Private m_CSS As cls_Gantt_Script_CSS


'### Properties
Public Property Get JS() As cls_Gantt_Script_JS
    Set JS = m_JS
End Property
Public Property Set JS(newValue As cls_Gantt_Script_JS)
    Set m_JS = newValue
End Property

Public Property Get CSS() As cls_Gantt_Script_CSS
    Set CSS = m_CSS
End Property
Public Property Set CSS(newValue As cls_Gantt_Script_CSS)
    Set m_CSS = newValue
End Property


'### Initialization
Private Sub Class_Initialize()

    Set Me.JS = New cls_Gantt_Script_JS
    Set Me.CSS = New cls_Gantt_Script_CSS
    
End Sub

Private Sub Class_Terminate()
    ' Do Nothing
End Sub


'### HTML Output
Public Function ToHtml() As String

    'Variables
    Dim s As String

    s = s & "    <script src=" & Me.JS.ToHtml() & "></script>" & vbCrLf
    s = s & "    <link rel='stylesheet' href=" & Me.CSS.ToHtml() & " />" & vbCrLf

    ToHtml = s
    
End Function

