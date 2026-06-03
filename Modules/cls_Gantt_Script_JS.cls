VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Gantt_Script_JS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'###############################################################################################
'# Copyright (c) 2026 Thomas Möller                                                            #
'# MIT License  => https://github.com/team-moeller/better-access-gantt-chart/blob/main/LICENSE #
'# Version 2.01.05  published: 03.06.2026                                                      #
'###############################################################################################

Option Compare Database
Option Explicit


'### Members
Private m_Source As gcScriptSource
Private m_CDN As String
Private m_PathFilename As String


'### Properties
Public Property Let Source(newValue As gcScriptSource)
    m_Source = newValue
End Property
Public Property Get Source() As gcScriptSource
    Source = m_Source
End Property

Public Property Let CDN(newValue As String)
    m_CDN = newValue
End Property
Public Property Get CDN() As String
    CDN = m_CDN
End Property

Public Property Let PathFilename(newValue As String)
    m_PathFilename = newValue
End Property
Public Property Get PathFilename() As String
    PathFilename = m_PathFilename
End Property


'### Initialization
Private Sub Class_Initialize()

    Me.Source = gcScriptSource.CDN
    Me.CDN = "https://cdn.jsdelivr.net/npm/frappe-gantt/dist/frappe-gantt.umd.js"
    Me.PathFilename = CurrentProject.Path & "\frappe-gantt.umd.js"
    
End Sub

Private Sub Class_Terminate()
    ' Do Nothing
End Sub


'### HTML Output
Public Function ToHtml() As String

    'Variables
    Dim s As String

    Select Case Me.Source
        Case gcScriptSource.CDN
            s = "'" & Me.CDN & "'"
        Case gcScriptSource.LocalFile
            s = "'" & "https://msaccess/" & Replace(Me.PathFilename, "\", "/") & "'"
    End Select

    ToHtml = s
    
End Function
