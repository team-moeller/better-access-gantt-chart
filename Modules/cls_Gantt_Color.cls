VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Gantt_Color"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'###############################################################################################
'# Copyright (c) 2026 Thomas Möller                                                            #
'# MIT License  => https://github.com/team-moeller/better-access-gantt-chart/blob/main/LICENSE #
'# Version 2.00.09  published: 01.06.2026                                                      #
'###############################################################################################

Option Compare Database
Option Explicit


'### Members
Private m_Theme As gcTheme
Private m_ColorTaskBarColor As String
Private m_ColorBarBorderColor As String
Private m_ColorProgressColor As String
Private m_ColorArrowColor As String
Private m_ColorHandleColor As String
Private m_ColorTextDark As String
Private m_ColorTextLight As String
Private m_ColorTextMuted As String
Private m_ColorHeaderBackground As String
Private m_ColorRowColor As String
Private m_ColorRowBorderColor As String
Private m_ColorTodayHighlightColor As String
Private m_ColorWeekendHighlightColor As String
Private m_ColorTickColor As String


'### Properties
Public Property Let Theme(newValue As gcTheme)
    m_Theme = newValue

    Select Case m_Theme
        Case thDefault
            'Do nothing
        Case thDarkMode
            Me.ColorTaskBarColor = "#3a3a3a"
            Me.ColorBarBorderColor = "#3a3a3a"
            Me.ColorProgressColor = "#4a6baf"
            Me.ColorArrowColor = "#9aa5b1"
            Me.ColorHandleColor = "#555"
            Me.ColorTextDark = "#e0e0e0"
            Me.ColorTextLight = "#f5f5f5"
            Me.ColorTextMuted = "#a0a0a0"
            Me.ColorHeaderBackground = "#1c1c1c"
            Me.ColorRowColor = "#252525"
            Me.ColorRowBorderColor = "#3d3d3d"
            Me.ColorTodayHighlightColor = "#ff9d48"
            Me.ColorWeekendHighlightColor = "#2d2d2d"
            Me.ColorTickColor = "#444"
        Case thBlue
            Me.ColorTaskBarColor = "#d7e8ff"
            Me.ColorBarBorderColor = "#a9c9f5"
            Me.ColorProgressColor = "#4a90e2"
            Me.ColorArrowColor = "#2f6fbf"
            Me.ColorHandleColor = "#3c5f8c"
            Me.ColorTextDark = "#1c2a3a"
            Me.ColorTextLight = "#f0f6ff"
            Me.ColorTextMuted = "#6b8bb5"
            Me.ColorHeaderBackground = "#e8f1ff"
            Me.ColorRowColor = "#f5f9ff"
            Me.ColorRowBorderColor = "#c7d9f2"
            Me.ColorTodayHighlightColor = "#4a90e2"
            Me.ColorWeekendHighlightColor = "#e1ecfa"
            Me.ColorTickColor = "#b4c9e6"
    End Select
End Property
Public Property Get Theme() As gcTheme
    Theme = m_Theme
End Property

Public Property Let ColorTaskBarColor(newValue As String)
    m_ColorTaskBarColor = newValue
End Property
Public Property Get ColorTaskBarColor() As String
    ColorTaskBarColor = m_ColorTaskBarColor
End Property

Public Property Let ColorBarBorderColor(newValue As String)
    m_ColorBarBorderColor = newValue
End Property
Public Property Get ColorBarBorderColor() As String
    ColorBarBorderColor = m_ColorBarBorderColor
End Property

Public Property Let ColorProgressColor(newValue As String)
    m_ColorProgressColor = newValue
End Property
Public Property Get ColorProgressColor() As String
    ColorProgressColor = m_ColorProgressColor
End Property

Public Property Let ColorArrowColor(newValue As String)
    m_ColorArrowColor = newValue
End Property
Public Property Get ColorArrowColor() As String
    ColorArrowColor = m_ColorArrowColor
End Property

Public Property Let ColorHandleColor(newValue As String)
    m_ColorHandleColor = newValue
End Property
Public Property Get ColorHandleColor() As String
    ColorHandleColor = m_ColorHandleColor
End Property

Public Property Let ColorTextDark(newValue As String)
    m_ColorTextDark = newValue
End Property
Public Property Get ColorTextDark() As String
    ColorTextDark = m_ColorTextDark
End Property

Public Property Let ColorTextLight(newValue As String)
    m_ColorTextLight = newValue
End Property
Public Property Get ColorTextLight() As String
    ColorTextLight = m_ColorTextLight
End Property

Public Property Let ColorTextMuted(newValue As String)
    m_ColorTextMuted = newValue
End Property
Public Property Get ColorTextMuted() As String
    ColorTextMuted = m_ColorTextMuted
End Property

Public Property Let ColorHeaderBackground(newValue As String)
    m_ColorHeaderBackground = newValue
End Property
Public Property Get ColorHeaderBackground() As String
    ColorHeaderBackground = m_ColorHeaderBackground
End Property

Public Property Let ColorRowColor(newValue As String)
    m_ColorRowColor = newValue
End Property
Public Property Get ColorRowColor() As String
    ColorRowColor = m_ColorRowColor
End Property

Public Property Let ColorRowBorderColor(newValue As String)
    m_ColorRowBorderColor = newValue
End Property
Public Property Get ColorRowBorderColor() As String
    ColorRowBorderColor = m_ColorRowBorderColor
End Property

Public Property Let ColorTodayHighlightColor(newValue As String)
    m_ColorTodayHighlightColor = newValue
End Property
Public Property Get ColorTodayHighlightColor() As String
    ColorTodayHighlightColor = m_ColorTodayHighlightColor
End Property

Public Property Let ColorWeekendHighlightColor(newValue As String)
    m_ColorWeekendHighlightColor = newValue
End Property
Public Property Get ColorWeekendHighlightColor() As String
    ColorWeekendHighlightColor = m_ColorWeekendHighlightColor
End Property

Public Property Let ColorTickColor(newValue As String)
    m_ColorTickColor = newValue
End Property
Public Property Get ColorTickColor() As String
    ColorTickColor = m_ColorTickColor
End Property


'### Initialization
Private Sub Class_Initialize()

    Me.Theme = gcTheme.thDefault
    
End Sub

Private Sub Class_Terminate()
    ' Do Nothing
End Sub


'### HTML Output
Public Function ToHtml() As String

    'Variables
    Dim s As String

    s = s & "    <style>" & vbCrLf
    s = s & "        /* Custom colors */" & vbCrLf
    s = s & "        .bar-red .bar {fill: #e74c3c !important;}" & vbCrLf
    s = s & "        .bar-green .bar {fill: #2ecc71 !important;}" & vbCrLf
    s = s & "        :root {" & vbCrLf

    If Len(Me.ColorTaskBarColor) > 0 Then s = s & "            --g-bar-color: " & Me.ColorTaskBarColor & ";" & vbCrLf
    If Len(Me.ColorBarBorderColor) > 0 Then s = s & "            --g-bar-border: " & Me.ColorBarBorderColor & ";" & vbCrLf
    If Len(Me.ColorProgressColor) > 0 Then s = s & "            --g-progress-color: " & Me.ColorProgressColor & ";" & vbCrLf
    If Len(Me.ColorArrowColor) > 0 Then s = s & "            --g-arrow-color: " & Me.ColorArrowColor & ";" & vbCrLf
    If Len(Me.ColorHandleColor) > 0 Then s = s & "            --g-handle-color: " & Me.ColorHandleColor & ";" & vbCrLf
    If Len(Me.ColorTextDark) > 0 Then s = s & "            --g-text-dark: " & Me.ColorTextDark & ";" & vbCrLf
    If Len(Me.ColorTextLight) > 0 Then s = s & "            --g-text-light: " & Me.ColorTextLight & ";" & vbCrLf
    If Len(Me.ColorTextMuted) > 0 Then s = s & "            --g-text-muted: " & Me.ColorTextMuted & ";" & vbCrLf
    If Len(Me.ColorHeaderBackground) > 0 Then s = s & "            --g-header-background: " & Me.ColorHeaderBackground & ";" & vbCrLf
    If Len(Me.ColorRowColor) > 0 Then s = s & "            --g-row-color: " & Me.ColorRowColor & ";" & vbCrLf
    If Len(Me.ColorRowBorderColor) > 0 Then s = s & "            --g-row-border-color: " & Me.ColorRowBorderColor & ";" & vbCrLf
    If Len(Me.ColorTodayHighlightColor) > 0 Then s = s & "            --g-today-highlight: " & Me.ColorTodayHighlightColor & ";" & vbCrLf
    If Len(Me.ColorWeekendHighlightColor) > 0 Then s = s & "            --g-weekend-highlight-color: " & Me.ColorWeekendHighlightColor & ";" & vbCrLf
    If Len(Me.ColorTickColor) > 0 Then s = s & "            --g-tick-color: " & Me.ColorTickColor & ";" & vbCrLf

    s = s & "        }" & vbCrLf
    s = s & "    </style>" & vbCrLf

    ToHtml = s
    
End Function

