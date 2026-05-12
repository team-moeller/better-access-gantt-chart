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
'# Version 1.27.06  published: 12.05.2026                                                      #
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
    'Appearance
    myGantt.Language = Nz(Forms!frm_configuration.cboLanguage, "en")
    myGantt.ViewMode = Nz(Forms!frm_configuration.cboViewMode, gcViewMode.vwDay)
    myGantt.ColumnWidth = Nz(Forms!frm_configuration.txtColumnWidth, 45)
    myGantt.BarHeight = Nz(Forms!frm_configuration.txtBarHeight, 30)
    myGantt.ArrowCurveRadius = Nz(Forms!frm_configuration.txtArrowCurveRadius, 5)
    myGantt.TaskBarCurveRadius = Nz(Forms!frm_configuration.txtTaskBarCurveRadius, 3)
    myGantt.ContainerHight = Nz(Forms!frm_configuration.txtContainerHight, "auto")
    myGantt.UpperHeaderHight = Nz(Forms!frm_configuration.txtUpperHeaderHight, 45)
    myGantt.LowerHeaderHight = Nz(Forms!frm_configuration.txtLowerHeaderHight, 30)
    myGantt.Padding = Nz(Forms!frm_configuration.txtPadding, 18)
    myGantt.Lines = Nz(Forms!frm_configuration.cboLines, gcLines.lnBoth)
    myGantt.ScrollTo = Nz(Forms!frm_configuration.cboScrollTo, gcScrollTo.stToday)
    myGantt.ScrollToDate = Nz(Forms!frm_configuration.txtScrollToDate, Date)
    myGantt.ShowExpectedProgress = Nz(Forms!frm_configuration.chkShowExpectedProgress, False)
    'Interaction
    myGantt.ShowTodayButton = Nz(Forms!frm_configuration.chkShowTodayButton, True)
    myGantt.ViewModeSelect = Nz(Forms!frm_configuration.chkViewModeSelect, False)
    myGantt.ReadonlyProgress = Nz(Forms!frm_configuration.chkReadOnlyProgress, False)
    myGantt.ReadonlyDates = Nz(Forms!frm_configuration.chkReadOnlyDates, False)
    myGantt.ReadonlyChart = Nz(Forms!frm_configuration.chkReadOnlyWholeChart, False)
    myGantt.AutoMoveLabel = Nz(Forms!frm_configuration.chkAutoMoveLabel, False)
    myGantt.MoveDependencies = Nz(Forms!frm_configuration.chkMoveDependencies, True)
    'Interactivity
    myGantt.ReactOnClick = Nz(Forms!frm_configuration.chkReactOnClick, True)
    'Scripts
    myGantt.GanttJsSource = Forms!frm_configuration.cboFrappeJsSource
    myGantt.GanttJsCDN = Forms!frm_configuration.txtFrappeJsCDN
    myGantt.GanttJsPathFilename = Forms!frm_configuration.txtFrappeJsPathFilename
    myGantt.GanttCssSource = Forms!frm_configuration.cboFrappeCssSource
    myGantt.GanttCssCDN = Forms!frm_configuration.txtFrappeCssCDN
    myGantt.GanttCssPathFilename = Forms!frm_configuration.txtFrappeCssPathFilename
    
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
