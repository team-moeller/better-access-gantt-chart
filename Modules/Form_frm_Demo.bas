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
'# Version 1.15.08  published: 13.04.2026                                                      #
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
    'Chart
    myGantt.Language = Nz(Forms!frm_Configuration.cboLanguage, "en")
    myGantt.ViewMode = Nz(Forms!frm_Configuration.cboViewMode, 4)
    myGantt.ColumnWidth = Nz(Forms!frm_Configuration.txtColumnWidth, 45)
    myGantt.BarHeight = Nz(Forms!frm_Configuration.txtBarHeight, 30)
    myGantt.ShowTodayButton = Nz(Forms!frm_Configuration.chkShowTodayButton, True)
    myGantt.ArrowCurveRadius = Nz(Forms!frm_Configuration.txtArrowCurveRadius, 5)
    myGantt.TaskBarCurveRadius = Nz(Forms!frm_Configuration.txtTaskBarCurveRadius, 3)
    myGantt.ContainerHight = Nz(Forms!frm_Configuration.txtContainerHight, "auto")
    myGantt.UpperHeaderHight = Nz(Forms!frm_Configuration.txtUpperHeaderHight, 45)
    myGantt.LowerHeaderHight = Nz(Forms!frm_Configuration.txtLowerHeaderHight, 30)
    
    myGantt.ShowGanttChart

End Sub

Private Sub ctlEdgeBrowser_Click()

    'Variables
    Dim Value As String
    
    Value = myGantt.RetrieveBridgeVariable
    
    If Value <> "null" Then
        MsgBox Value
    End If
    
    myGantt.ResetBridgeVariable

End Sub
