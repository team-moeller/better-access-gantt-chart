VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_Demo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'###############################################################################################
'# Copyright (c) 2026 Thomas Möller                                                            #
'# MIT License  => https://github.com/team-moeller/better-access-gantt-chart/blob/main/LICENSE #
'# Version 2.08.01  published: 08.07.2026                                                      #
'###############################################################################################

Option Compare Database
Option Explicit

'Members
Dim myGantt As cls_Better_Access_Gantt_Chart


Private Sub Form_Load()

    Me.lblVersion.Caption = "Version: " & DMax("V_Number", "tbl_VersionHistory")

End Sub

Public Sub cmdCreateGanttChart_Click()
    
    'Make sure frm_Configuration is open
    If IsFormOpen("frm_Configuration") = False Then
        DoCmd.OpenForm "frm_Configuration", acNormal
    End If
    
    'Create chart
    Set myGantt = New cls_Better_Access_Gantt_Chart
    
    Set myGantt.Control = Me.ctlEdgeBrowser
    'Data
    myGantt.Data.TableName = Nz(Forms!frm_configuration.txtTableName, "tbl_DemoData")
    myGantt.Data.FieldID = Nz(Forms!frm_configuration.txtFieldID, "TID")
    myGantt.Data.FieldName = Nz(Forms!frm_configuration.txtFieldName, "TName")
    myGantt.Data.FieldStart = Nz(Forms!frm_configuration.txtFieldStart, "TStart")
    myGantt.Data.FieldEnd = Nz(Forms!frm_configuration.txtFieldEnd, "TEnd")
    myGantt.Data.FieldProgress = Nz(Forms!frm_configuration.txtFieldProgress, "TProgress")
    myGantt.Data.FieldDependencies = Nz(Forms!frm_configuration.txtFieldDependencies, "TDependencies")
    myGantt.Data.FieldCustomClass = Nz(Forms!frm_configuration.txtFieldCustomClass, "TCustomClass")
    'Appearance
    myGantt.Appearance.Language = Nz(Forms!frm_configuration.cboLanguage, "en")
    myGantt.Appearance.ViewMode = Nz(Forms!frm_configuration.cboViewMode, gcViewMode.vwDay)
    myGantt.Appearance.ColumnWidth = Nz(Forms!frm_configuration.txtColumnWidth, 45)
    myGantt.Appearance.BarHeight = Nz(Forms!frm_configuration.txtBarHeight, 30)
    myGantt.Appearance.ArrowCurveRadius = Nz(Forms!frm_configuration.txtArrowCurveRadius, 5)
    myGantt.Appearance.TaskBarCurveRadius = Nz(Forms!frm_configuration.txtTaskBarCurveRadius, 3)
    myGantt.Appearance.ContainerHeight = Nz(Forms!frm_configuration.txtContainerHeight, "auto")
    myGantt.Appearance.UpperHeaderHeight = Nz(Forms!frm_configuration.txtUpperHeaderHeight, 45)
    myGantt.Appearance.LowerHeaderHeight = Nz(Forms!frm_configuration.txtLowerHeaderHeight, 30)
    myGantt.Appearance.Padding = Nz(Forms!frm_configuration.txtPadding, 18)
    myGantt.Appearance.Lines = Nz(Forms!frm_configuration.cboLines, gcLines.lnBoth)
    myGantt.Appearance.ScrollTo = Nz(Forms!frm_configuration.cboScrollTo, gcScrollTo.stToday)
    myGantt.Appearance.ScrollToDate = Nz(Forms!frm_configuration.txtScrollToDate, Date)
    myGantt.Appearance.ShowExpectedProgress = Nz(Forms!frm_configuration.chkShowExpectedProgress, False)
    'Interaction
    myGantt.Interaction.ShowTodayButton = Nz(Forms!frm_configuration.chkShowTodayButton, True)
    myGantt.Interaction.ViewModeSelect = Nz(Forms!frm_configuration.chkViewModeSelect, False)
    myGantt.Interaction.ReadonlyProgress = Nz(Forms!frm_configuration.chkReadOnlyProgress, False)
    myGantt.Interaction.ReadonlyDates = Nz(Forms!frm_configuration.chkReadOnlyDates, False)
    myGantt.Interaction.ReadonlyChart = Nz(Forms!frm_configuration.chkReadOnlyWholeChart, False)
    myGantt.Interaction.AutoMoveLabel = Nz(Forms!frm_configuration.chkAutoMoveLabel, False)
    myGantt.Interaction.MoveDependencies = Nz(Forms!frm_configuration.chkMoveDependencies, True)
    'Interactivity
    myGantt.Interactivity.ReactOnClick = Nz(Forms!frm_configuration.chkReactOnClick, True)
    myGantt.Interactivity.ReactOnProgressChange = Nz(Forms!frm_configuration.chkReactOnProgressChange, True)
    myGantt.Interactivity.ReactOnDateChange = Nz(Forms!frm_configuration.chkReactOnDateChange, True)
    'Color
    myGantt.Color.Theme = Nz(Forms!frm_configuration.fraColorTheme, gcTheme.thDefault)
    myGantt.Color.ColorTaskBarColor = Nz(Forms!frm_configuration.txtColorTaskBarColor, vbNullString)
    myGantt.Color.ColorBarBorderColor = Nz(Forms!frm_configuration.txtColorBarBorderColor, vbNullString)
    myGantt.Color.ColorProgressColor = Nz(Forms!frm_configuration.txtColorProgressColor, vbNullString)
    myGantt.Color.ColorArrowColor = Nz(Forms!frm_configuration.txtColorArrowColor, vbNullString)
    myGantt.Color.ColorHandleColor = Nz(Forms!frm_configuration.txtColorHandleColor, vbNullString)
    myGantt.Color.ColorTextDark = Nz(Forms!frm_configuration.txtColorTextDark, vbNullString)
    myGantt.Color.ColorTextLight = Nz(Forms!frm_configuration.txtColorTextLight, vbNullString)
    myGantt.Color.ColorTextMuted = Nz(Forms!frm_configuration.txtColorTextMuted, vbNullString)
    myGantt.Color.ColorHeaderBackground = Nz(Forms!frm_configuration.txtColorHeaderBackground, vbNullString)
    myGantt.Color.ColorRowColor = Nz(Forms!frm_configuration.txtColorRowColor, vbNullString)
    myGantt.Color.ColorRowBorderColor = Nz(Forms!frm_configuration.txtColorRowBorderColor, vbNullString)
    myGantt.Color.ColorTodayHighlightColor = Nz(Forms!frm_configuration.txtColorTodayHighlight, vbNullString)
    myGantt.Color.ColorWeekendHighlightColor = Nz(Forms!frm_configuration.txtColorWeekendHighlight, vbNullString)
    myGantt.Color.ColorTickColor = Nz(Forms!frm_configuration.txtColorTickColor, vbNullString)
    'Scripts
    myGantt.Scripts.JS.Source = Forms!frm_configuration.cboFrappeJsSource
    myGantt.Scripts.JS.CDN = Forms!frm_configuration.txtFrappeJsCDN
    myGantt.Scripts.JS.PathFilename = Forms!frm_configuration.txtFrappeJsPathFilename
    myGantt.Scripts.CSS.Source = Forms!frm_configuration.cboFrappeCssSource
    myGantt.Scripts.CSS.CDN = Forms!frm_configuration.txtFrappeCssCDN
    myGantt.Scripts.CSS.PathFilename = Forms!frm_configuration.txtFrappeCssPathFilename
    
    myGantt.ShowGanttChart

End Sub

Private Sub ctlEdgeBrowser_Click()

    'Variables
    Dim value As String
    Dim obj As Object
    
    value = myGantt.RetrieveBridgeVariable
    
    If value = "null" Then Exit Sub
    
    Set obj = ParseEvent(value)
    
    Select Case obj("event")
        Case "click"
            MsgBox "ID: " & obj("id") & vbCrLf, , _
                   "Event: " & obj("event")
        Case "progress"
            MsgBox "ID: " & obj("id") & vbCrLf & _
                   "Progress: " & obj("progress"), , _
                   "Event: " & obj("event")
        Case "datechange"
            MsgBox "ID: " & obj("id") & vbCrLf & _
                   "Start: " & obj("start") & vbCrLf & _
                   "End: " & obj("end"), , _
                   "Event: " & obj("event")
        Case Else
            MsgBox "Event unknown"
    End Select
    
    myGantt.ResetBridgeVariable

End Sub

Function ParseEvent(raw As String) As Object

    'Variables
    Dim dict As Object
    Dim json As String
    Dim parts() As String
    Dim p As Variant
    Dim pos As Long
    Dim key As String
    Dim val As String

    Set dict = CreateObject("Scripting.Dictionary")
    
    'Extract json
    json = Left$(raw, InStr(raw, "}"))
    json = Mid$(json, 2, Len(json) - 2)          ' remove { }
    parts = Split(json, ",")

    For Each p In parts
        pos = InStr(1, p, ":")
        If pos > 0 Then
            key = Left$(p, pos - 1)
            val = Mid$(p, pos + 1)

            key = Replace(key, """", "")
            key = Trim$(key)
            val = Trim$(Replace(val, """", ""))

            ' Numbers vs. Strings
            If IsNumeric(val) Then
                dict(key) = CLng(val)
            Else
                dict(key) = val
            End If
        End If
    Next p

    Set ParseEvent = dict
    
End Function
