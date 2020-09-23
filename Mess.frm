VERSION 5.00
Begin VB.Form Mess 
   Caption         =   "         Congratulations!"
   ClientHeight    =   1710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   114
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   280
   Begin VB.CommandButton cmdMess 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2783
      TabIndex        =   3
      Top             =   1200
      Width           =   1125
   End
   Begin VB.CommandButton cmdMess 
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1538
      TabIndex        =   2
      Top             =   1200
      Width           =   1125
   End
   Begin VB.CommandButton cmdMess 
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   293
      TabIndex        =   1
      Top             =   1200
      Width           =   1125
   End
   Begin VB.Label lblMess 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   323
      TabIndex        =   0
      Top             =   240
      Width           =   3555
   End
End
Attribute VB_Name = "Mess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    Mess                                                         ****
'   ****                                                                 ****
'   ****    Written by:    Randy Giese    (2007/09/05)                   ****
'   ****                                                                 ****
'   ****    This is my home-made dime store version of a Message Box.    ****
'   ****    It was written as a one-time shot, so I did not              ****
'   ****    incorporate any options or graphics.  My only purpose in     ****
'   ****    creating it was so I could display my message in the         ****
'   ****    Upper Left-Hand corner of the screen.                        ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'
'   RandyGrams Comments - Left Align the above comments.

Option Explicit

Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Private Sub cmdMess_Click(Index As Integer)

    lngMsgResp = Index
    Unload Me

End Sub

Private Sub Form_Load()

    SetCursorPos cmdMess(0).Left + (cmdMess(0).Width * 3 \ 4), cmdMess(0).Top + (cmdMess(0).Height \ 2) + 30
    lblMess.Caption = "You've successfully solved the Puzzle!" & vbNewLine & vbNewLine & "Would you like to try a different Puzzle size?"

End Sub
