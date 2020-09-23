VERSION 5.00
Begin VB.Form BackGround 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "BackGround"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const mc_strModuleID   As String = "BackGround."

Private Sub Form_Click()
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    If the user should click somewhere outside the puzzle        ****
'   ****    boundaries, this will put the "Focus" back on the puzzle     ****
'   ****    again.                                                       ****
'   ****                                                                 ****
'   *************************************************************************
'
    On Error GoTo ErrorHandler
    Puzzle_Main.SetFocus
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "Form_Click()"

End Sub
