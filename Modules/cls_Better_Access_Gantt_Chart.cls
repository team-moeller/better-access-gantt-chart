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
'# Version 2.05.01  published: 05.07.2026                                                      #
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
Private m_Appearance As cls_Gantt_Appearance
Private m_Interaction As cls_Gantt_Interaction
Private m_Interactivity As cls_Gantt_Interactivity
Private m_Color As cls_Gantt_Color
Private m_Scripts As cls_Gantt_Scripts
Private m_Data As cls_Gantt_Data


'### Properties
Public Property Set Control(newValue As Edge)
    Set m_Control = newValue
End Property
Public Property Get Control() As Edge
    Set Control = m_Control
End Property

Public Property Get Appearance() As cls_Gantt_Appearance
    Set Appearance = m_Appearance
End Property
Public Property Set Appearance(newValue As cls_Gantt_Appearance)
    Set m_Appearance = newValue
End Property

Public Property Get Interaction() As cls_Gantt_Interaction
    Set Interaction = m_Interaction
End Property
Public Property Set Interaction(newValue As cls_Gantt_Interaction)
    Set m_Interaction = newValue
End Property

Public Property Get Interactivity() As cls_Gantt_Interactivity
    Set Interactivity = m_Interactivity
End Property
Public Property Set Interactivity(newValue As cls_Gantt_Interactivity)
    Set m_Interactivity = newValue
End Property

Public Property Get Color() As cls_Gantt_Color
    Set Color = m_Color
End Property
Public Property Set Color(newValue As cls_Gantt_Color)
    Set m_Color = newValue
End Property

Public Property Get Scripts() As cls_Gantt_Scripts
    Set Scripts = m_Scripts
End Property
Public Property Set Scripts(newValue As cls_Gantt_Scripts)
    Set m_Scripts = newValue
End Property

Public Property Get Data() As cls_Gantt_Data
    Set Data = m_Data
End Property
Public Property Set Data(newValue As cls_Gantt_Data)
    Set m_Data = newValue
End Property


'### Instantiation
Private Sub Class_Initialize()

    Set Me.Appearance = New cls_Gantt_Appearance
    Set Me.Interaction = New cls_Gantt_Interaction
    Set Me.Interactivity = New cls_Gantt_Interactivity
    Set Me.Color = New cls_Gantt_Color
    Set Me.Scripts = New cls_Gantt_Scripts
    Set Me.Data = New cls_Gantt_Data

End Sub

Private Sub Class_Terminate()
    ' Do Nothing
End Sub


'### Public Methods
Public Sub ShowGanttChart()

    'Variables
    Dim Datei As String
    
    CreateHtmlFile
       
    Datei = "https://msaccess/" & CurrentProject.Path & "\BetterAccessGanttChart.html"
    
    Me.Control.Navigate Datei
    
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
    Dim s As String
    
    s = s & "<!DOCTYPE html>" & vbCrLf
    s = s & "<html lang='en'>" & vbCrLf
    s = s & "<head>" & vbCrLf
    s = s & "    <meta charset='UTF-8'>" & vbCrLf
    s = s & "    <meta http-equiv='X-UA-Compatible' content='IE=Edge'/>" & vbCrLf
    s = s & "    <title>Better Access Gantt Chart</title>" & vbCrLf
    s = s & "    <style>body{margin:0;padding:0;}</style>" & vbCrLf

    'Scripts (JS + CSS)
    s = s & Me.Scripts.ToHtml
    'Colors
    s = s & Me.Color.ToHtml

    s = s & "</head>" & vbCrLf
    s = s & "<body>" & vbCrLf
    
    getHead = s
    
End Function

Private Function getScript() As String

    'Variables
    Dim s As String
    
    s = s & "    <h2>Demo project</h2>" & vbCrLf
    s = s & "    <div id='gantt'></div>" & vbCrLf
    
    s = s & "    <script>" & vbCrLf
    s = s & "        window.AccessEdgeBridge = null;" & vbCrLf

    'Tasks
    s = s & Me.Data.ToScript

    'Gantt init
    s = s & "        const gantt = new Gantt('#gantt', tasks, {" & vbCrLf

    s = s & Me.Interactivity.ToScript
    s = s & Me.Appearance.ToScript
    s = s & Me.Interaction.ToScript

    s = s & "        });" & vbCrLf
    s = s & "    </script>" & vbCrLf
    
    getScript = s

End Function

Private Function getFoot() As String

    'Variables
    Dim s As String
    
    s = s & "</body>" & vbCrLf
    s = s & "</html>"
    
    getFoot = s
    
End Function
