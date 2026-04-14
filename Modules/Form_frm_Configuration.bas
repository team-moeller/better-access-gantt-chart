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
'# Version 1.16.24  published: 14.04.2026                                                      #
'###############################################################################################

Option Compare Database
Option Explicit

Private Sub cmdUpdatePivotTable_Click()
    
    'Make sure frm_Demo is open
    If IsFormOpen("frm_Demo") = False Then
        DoCmd.OpenForm "frm_Demo", acNormal
    End If
    
    Forms!frm_Demo.SetFocus
    DoEvents
    Forms!frm_Demo.cmdCreateGanttChart_Click

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

