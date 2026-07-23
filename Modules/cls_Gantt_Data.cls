VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Gantt_Data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'###############################################################################################
'# Copyright (c) 2026 Thomas Möller                                                            #
'# MIT License  => https://github.com/team-moeller/better-access-gantt-chart/blob/main/LICENSE #
'# Version 2.17.01  published: 23.07.2026                                                      #
'###############################################################################################

Option Compare Database
Option Explicit


'### Members
Private m_TableName As String
Private m_FieldID As String
Private m_FieldName As String
Private m_FieldStart As String
Private m_FieldEnd As String
Private m_FieldProgress As String
Private m_FieldDependencies As String
Private m_FieldCustomClass As String


'### Properties
Public Property Let TableName(newValue As String)
    m_TableName = newValue
End Property
Public Property Get TableName() As String
    TableName = m_TableName
End Property

Public Property Let FieldID(newValue As String)
    m_FieldID = newValue
End Property
Public Property Get FieldID() As String
    FieldID = m_FieldID
End Property

Public Property Let FieldName(newValue As String)
    m_FieldName = newValue
End Property
Public Property Get FieldName() As String
    FieldName = m_FieldName
End Property

Public Property Let FieldStart(newValue As String)
    m_FieldStart = newValue
End Property
Public Property Get FieldStart() As String
    FieldStart = m_FieldStart
End Property

Public Property Let FieldEnd(newValue As String)
    m_FieldEnd = newValue
End Property
Public Property Get FieldEnd() As String
    FieldEnd = m_FieldEnd
End Property

Public Property Let FieldProgress(newValue As String)
    m_FieldProgress = newValue
End Property
Public Property Get FieldProgress() As String
    FieldProgress = m_FieldProgress
End Property

Public Property Let FieldDependencies(newValue As String)
    m_FieldDependencies = newValue
End Property
Public Property Get FieldDependencies() As String
    FieldDependencies = m_FieldDependencies
End Property

Public Property Let FieldCustomClass(newValue As String)
    m_FieldCustomClass = newValue
End Property
Public Property Get FieldCustomClass() As String
    FieldCustomClass = m_FieldCustomClass
End Property


'### Initialization
Private Sub Class_Initialize()

    Me.TableName = "tbl_DemoData"
    Me.FieldID = "TID"
    Me.FieldName = "TName"
    Me.FieldStart = "TStart"
    Me.FieldEnd = "TEnd"
    Me.FieldProgress = "TProgress"
    Me.FieldDependencies = "TDependencies"
    Me.FieldCustomClass = "TCustomClass"

End Sub

Private Sub Class_Terminate()
    ' Do Nothing
End Sub


'### JS Output
Public Function ToScript() As String

    'Variables
    Dim db As DAO.Database
    Dim rst As DAO.Recordset
    Dim s As String
    Dim sql As String

    Set db = CurrentDb()
    sql = "SELECT * FROM [" & Me.TableName & "]"
    Set rst = db.OpenRecordset(sql, dbOpenSnapshot)

    s = s & "        const tasks = [" & vbCrLf

    Do While Not rst.EOF

        s = s & "            {"
        s = s & "id: '" & rst(Me.FieldID) & "', "
        s = s & "name: '" & EscapeJsString(rst(Me.FieldName)) & "', "
        s = s & "start: '" & Format$(rst(Me.FieldStart), "YYYY-MM-DD") & "', "
        s = s & "end: '" & Format$(rst(Me.FieldEnd), "YYYY-MM-DD") & "', "
        s = s & "progress: '" & rst(Me.FieldProgress) & "', "

        If Trim$(Nz(rst(Me.FieldDependencies), vbNullString)) <> vbNullString Then
            s = s & "dependencies: '" & EscapeJsString(rst(Me.FieldDependencies)) & "', "
        End If

        If Trim$(Nz(rst(Me.FieldCustomClass), vbNullString)) <> vbNullString Then
            s = s & "custom_class: '" & EscapeJsString(rst(Me.FieldCustomClass)) & "'"
        End If

        If Right$(s, 2) = ", " Then
            s = Left$(s, Len(s) - 2)
        End If

        s = s & "}," & vbCrLf

        rst.MoveNext
    Loop

    s = Left$(s, Len(s) - 3) & vbCrLf
    s = s & "        ];" & vbCrLf

    ToScript = s

    rst.Close
    Set rst = Nothing
    Set db = Nothing

End Function

Private Function EscapeJsString(s As String) As String

    s = Replace(s, "\", "\\")
    s = Replace(s, "'", "\'")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    
    EscapeJsString = s

End Function
