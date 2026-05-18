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
'# Version 1.31.01  published: 18.05.2026                                                      #
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

Private Sub fraThemingTheme_AfterUpdate()

    Select Case Me.fraThemingTheme
        Case 1
            Me.txtThemingTaskBarColor = "#fff"
            Me.txtThemingBarBorderColor = "#fff"
            Me.txtThemingProgressColor = "#dbdbdb"
            Me.txtThemingArrowColor = "#1f2937"
            Me.txtThemingHandleColor = "#37352f"
            Me.txtThemingTextDark = "#171717"
            Me.txtThemingTextLight = "#fff"
            Me.txtThemingTextMuted = "#7c7c7c"
            Me.txtThemingHeaderBackground = "#fff"
            Me.txtThemingRowColor = "#fdfdfd"
            Me.txtThemingRowBorderColor = "#c7c7c7"
            Me.txtThemingTodayHighlight = "#37352f"
            Me.txtThemingWeekendHighlight = "#f7f7f7"
            Me.txtThemingTickColor = "#e0e0e0"
        Case 2
            Me.txtThemingTaskBarColor = "#3a3a3a"
            Me.txtThemingBarBorderColor = "#3a3a3a"
            Me.txtThemingProgressColor = "#4a6baf"
            Me.txtThemingArrowColor = "#9aa5b1"
            Me.txtThemingHandleColor = "#555"
            Me.txtThemingTextDark = "#e0e0e0"
            Me.txtThemingTextLight = "#f5f5f5"
            Me.txtThemingTextMuted = "#a0a0a0"
            Me.txtThemingHeaderBackground = "#1c1c1c"
            Me.txtThemingRowColor = "#252525"
            Me.txtThemingRowBorderColor = "#3d3d3d"
            Me.txtThemingTodayHighlight = "#ff9d48"
            Me.txtThemingWeekendHighlight = "#2d2d2d"
            Me.txtThemingTickColor = "#444"
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

