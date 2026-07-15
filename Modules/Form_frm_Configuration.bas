VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_Configuration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'###############################################################################################
'# Copyright (c) 2026 Thomas Möller                                                            #
'# MIT License  => https://github.com/team-moeller/better-access-gantt-chart/blob/main/LICENSE #
'# Version 2.11.01  published: 15.07.2026                                                      #
'###############################################################################################

Option Compare Database
Option Explicit

Private Sub cmdUpdateGanttChart_Click()
    
    'Make sure frm_Demo is open
    If IsFormOpen("frm_Demo") = False Then
        DoCmd.OpenForm "frm_Demo", acNormal
    End If
    
    Forms!frm_Demo.SetFocus
    DoEvents
    Forms!frm_Demo.cmdCreateGanttChart_Click

End Sub

Private Sub cboScrollTo_AfterUpdate()

    Select Case Me.cboScrollTo
        Case 1, 2, 3
            Me.txtScrollToDate.Enabled = False
        Case 4
            Me.txtScrollToDate.Enabled = True
        Case Else
            Me.txtScrollToDate.Enabled = False
    End Select

End Sub

Private Sub fraColorTheme_AfterUpdate()

    Select Case Me.fraColorTheme
        Case 0
            Me.txtColorTaskBarColor = "#fff"
            Me.txtColorBarBorderColor = "#fff"
            Me.txtColorProgressColor = "#dbdbdb"
            Me.txtColorArrowColor = "#1f2937"
            Me.txtColorHandleColor = "#37352f"
            Me.txtColorTextDark = "#171717"
            Me.txtColorTextLight = "#fff"
            Me.txtColorTextMuted = "#7c7c7c"
            Me.txtColorHeaderBackground = "#fff"
            Me.txtColorRowColor = "#fdfdfd"
            Me.txtColorRowBorderColor = "#c7c7c7"
            Me.txtColorTodayHighlight = "#37352f"
            Me.txtColorWeekendHighlight = "#f7f7f7"
            Me.txtColorTickColor = "#e0e0e0"
        Case 1
            Me.txtColorTaskBarColor = "#3a3a3a"
            Me.txtColorBarBorderColor = "#3a3a3a"
            Me.txtColorProgressColor = "#4a6baf"
            Me.txtColorArrowColor = "#9aa5b1"
            Me.txtColorHandleColor = "#555"
            Me.txtColorTextDark = "#e0e0e0"
            Me.txtColorTextLight = "#f5f5f5"
            Me.txtColorTextMuted = "#a0a0a0"
            Me.txtColorHeaderBackground = "#1c1c1c"
            Me.txtColorRowColor = "#252525"
            Me.txtColorRowBorderColor = "#3d3d3d"
            Me.txtColorTodayHighlight = "#ff9d48"
            Me.txtColorWeekendHighlight = "#2d2d2d"
            Me.txtColorTickColor = "#444"
        Case 2
            Me.txtColorTaskBarColor = "#d7e8ff"
            Me.txtColorBarBorderColor = "#a9c9f5"
            Me.txtColorProgressColor = "#4a90e2"
            Me.txtColorArrowColor = "#2f6fbf"
            Me.txtColorHandleColor = "#3c5f8c"
            Me.txtColorTextDark = "#1c2a3a"
            Me.txtColorTextLight = "#f0f6ff"
            Me.txtColorTextMuted = "#6b8bb5"
            Me.txtColorHeaderBackground = "#e8f1ff"
            Me.txtColorRowColor = "#f5f9ff"
            Me.txtColorRowBorderColor = "#c7d9f2"
            Me.txtColorTodayHighlight = "#4a90e2"
            Me.txtColorWeekendHighlight = "#e1ecfa"
            Me.txtColorTickColor = "#b4c9e6"
        Case Else
            'Do Nothing
    End Select

End Sub

Private Sub cmdSaveFrappeJs_Click()
    
    If SaveFileToDisk("frappe-gantt.umd.js", CurrentProject.Path) = True Then
        MsgBox "frappe-gantt.umd.js successfully saved to hdd.", vbInformation, "Better Access Gantt Chart"
    Else
        MsgBox "Something went wrong!", vbExclamation, "Better Access Gantt Chart"
    End If

End Sub

Private Sub cmdSaveFrappeCss_Click()
    
    If SaveFileToDisk("frappe-gantt.css", CurrentProject.Path) = True Then
        MsgBox "frappe-gantt.css successfully saved to hdd.", vbInformation, "Better Access Gantt Chart"
    Else
        MsgBox "Something went wrong!", vbExclamation, "Better Access Gantt Chart"
    End If

End Sub

