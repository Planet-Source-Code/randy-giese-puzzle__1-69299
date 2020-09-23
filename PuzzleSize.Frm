VERSION 5.00
Begin VB.Form Puzzle_Size 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2205
   ClientLeft      =   15
   ClientTop       =   -75
   ClientWidth     =   8055
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   36
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PuzzleSize.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   8055
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picColumns 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   930
      Left            =   90
      ScaleHeight     =   870
      ScaleWidth      =   7800
      TabIndex        =   3
      Top             =   1155
      Width           =   7860
      Begin VB.TextBox txtPuzzle_Size 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   825
         Left            =   7005
         TabIndex        =   4
         Top             =   30
         Width           =   720
      End
      Begin VB.Label lblNumberOfColumns 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   " Number of Rows and Columns: "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   630
         Left            =   0
         TabIndex        =   5
         Top             =   120
         Width           =   7695
      End
   End
   Begin VB.PictureBox picPuzzle_Size 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   945
      Left            =   90
      ScaleHeight     =   885
      ScaleWidth      =   4380
      TabIndex        =   1
      Top             =   90
      Width           =   4440
      Begin VB.Label lblPuzzle_Size 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   120
         TabIndex        =   2
         Top             =   150
         Width           =   120
      End
   End
   Begin VB.CommandButton cmdContinue 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Continue!"
      Height          =   945
      Left            =   4650
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   90
      Width           =   3300
   End
End
Attribute VB_Name = "Puzzle_Size"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    Puzzle_Size                                                  ****
'   ****    Written by:    Randy Giese    (01/04/2003)                   ****
'   ****                                                                 ****
'   *************************************************************************
'   ****                                                                 ****
'   ****    Puzzle_Size gives the player the opportunity to specify      ****
'   ****    the number of pieces in the next Puzzle.                     ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'
'   RandyGrams Comments - Left Align the above comments.

Option Explicit

Private Const mc_strModuleID   As String = "Puzzle_Size."

Private Sub cmdContinue_Click()

    On Error GoTo ErrorHandler
    lngPerRow = txtPuzzle_Size.Text
    lngBoxes = lngPerRow * lngPerRow
    Unload Me
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "cmdContinue_Click()"

End Sub

Private Sub Form_Load()

    On Error GoTo ErrorHandler
    txtPuzzle_Size.Text = Sqr(lngBoxes) + 1
    lblPuzzle_Size.Caption = "You may change the Puzzle size in the box below." & vbCr & _
                             "The Puzzle you just finished was:  --->  " & Sqr(lngBoxes) & "*" & Sqr(lngBoxes) & "."

    lblPuzzle_Size.Move 300, 150
    picPuzzle_Size.Move 300, 300, lblPuzzle_Size.Width + 600, lblPuzzle_Size.Height + 450

    With picColumns
        .Move picPuzzle_Size.Left, picPuzzle_Size.Top + picPuzzle_Size.Height + 150, picPuzzle_Size.Width, .Height
        txtPuzzle_Size.Left = .Width - txtPuzzle_Size.Width - 120
        cmdContinue.Move .Left + .Width - cmdContinue.Width, .Top + .Height + 150
    End With  'picColumns

    Me.Move 0, 0, picColumns.Width + 720, cmdContinue.Top + cmdContinue.Height + 300
    DoEvents
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "Form_Load()"

End Sub
