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
'# Version 1.02.09  published: 30.03.2026                                                      #
'###############################################################################################

Option Compare Database
Option Explicit


'### Enums



'### Members

Private WithEvents m_Control As Edge
Attribute m_Control.VB_VarHelpID = -1
Private m_Language As String


'### Properties

Public Property Set Control(newValue As Edge)
    Set m_Control = newValue
End Property
Public Property Get Control() As Edge
    Set Control = m_Control
End Property

Public Property Let Language(newValue As String)
    m_Language = newValue
End Property
Public Property Get Language() As String
    Language = m_Language
End Property



'### Instanzing

Private Sub Class_Initialize()

    Me.Language = "en"
    
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
    strText = strText & "    <script src='https://cdn.jsdelivr.net/npm/frappe-gantt/dist/frappe-gantt.umd.js'></script>" & vbCrLf
    strText = strText & "    <link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/frappe-gantt/dist/frappe-gantt.css'/>" & vbCrLf
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
    strText = strText & "        const tasks = [" & vbCrLf
    strText = strText & getTaskData & vbCrLf
    strText = strText & "        ];" & vbCrLf
    strText = strText & "        const gantt = new Gantt('#gantt', tasks, {" & vbCrLf
    strText = strText & "            view_mode: 'Day'," & vbCrLf
    strText = strText & "            language: '" & Me.Language & "'" & vbCrLf
    strText = strText & "        });" & vbCrLf
    strText = strText & "    </script>" & vbCrLf
    
    strText = strText & "    <style>" & vbCrLf
    strText = strText & "        /* Custom colors */" & vbCrLf
    strText = strText & "        .bar-red .bar {" & vbCrLf
    strText = strText & "            fill: #e74c3c !important;" & vbCrLf
    strText = strText & "        }" & vbCrLf
    strText = strText & "        .bar-green .bar {" & vbCrLf
    strText = strText & "            fill: #2ecc71 !important;" & vbCrLf
    strText = strText & "        }" & vbCrLf
    strText = strText & "    </style>" & vbCrLf
    
    getScript = strText

End Function

Private Function getFoot() As String

    'Variables
    Dim strText As String
    
    strText = strText & "</body>" & vbCrLf
    strText = strText & "</html>"
    
    getFoot = strText
    
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
    
    strText = Left$(strText, Len(strText) - 3)
    getTaskData = strText
    
    rst.Close
    Set rst = Nothing
    Set db = Nothing
    
End Function

