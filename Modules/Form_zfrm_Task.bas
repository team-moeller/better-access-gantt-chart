VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_zfrm_Task"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'###############################################################################################
'# Copyright (c) 2026 Thomas Möller                                                            #
'# MIT License  => https://github.com/team-moeller/better-access-gantt-chart/blob/main/LICENSE #
'# Version 2.18.01  published: 25.07.2026                                                      #
'###############################################################################################

Option Compare Database
Option Explicit

Public UpdateTask As Boolean

Private Sub Form_BeforeUpdate(Cancel As Integer)

    If Me.UpdateTask = False Then
        Me.Undo
    End If
    
End Sub

Private Sub cmdExit_Click()

    Me.UpdateTask = False
    Me.Visible = False
    
End Sub

Private Sub cmdUpdateTaskFromAccess_Click()

    Me.UpdateTask = True
    Me.Visible = False

End Sub
