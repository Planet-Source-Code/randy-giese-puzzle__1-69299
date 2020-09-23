VERSION 5.00
Begin VB.Form RG_About 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   7995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8640
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleMode       =   0  'User
   ScaleWidth      =   8640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicBlack 
      BackColor       =   &H00000000&
      ForeColor       =   &H00000000&
      Height          =   7760
      Left            =   120
      ScaleHeight     =   7695
      ScaleWidth      =   8340
      TabIndex        =   0
      Top             =   120
      Width           =   8400
      Begin VB.PictureBox picBlue 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   7150
         Left            =   250
         ScaleHeight     =   7155
         ScaleWidth      =   7800
         TabIndex        =   1
         Top             =   250
         Width           =   7800
         Begin VB.PictureBox picTop 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   3500
            Left            =   300
            ScaleHeight     =   3495
            ScaleWidth      =   7200
            TabIndex        =   8
            Top             =   300
            Width           =   7200
            Begin VB.PictureBox Logo 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   1800
               Left            =   4455
               Picture         =   "Rg_About.frx":0000
               ScaleHeight     =   1740
               ScaleWidth      =   2325
               TabIndex        =   9
               Top             =   1500
               Width           =   2385
            End
            Begin VB.Label lblCompanyName 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "RandyGrams - (2006)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   18
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   495
               Left            =   1650
               TabIndex        =   16
               Top             =   120
               Width           =   4170
            End
            Begin VB.Label lblPlatform 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Windows Compatible"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   360
               TabIndex        =   15
               Top             =   1500
               Width           =   2025
            End
            Begin VB.Label lblVersion 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Version  1.2.3"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   360
               TabIndex        =   14
               Top             =   1905
               Width           =   1275
            End
            Begin VB.Label lblCopyright 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Copyright:  Randy-Grams (2006)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   360
               TabIndex        =   13
               Top             =   2280
               Width           =   2970
            End
            Begin VB.Label lblProductName 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "RandyGrams"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   32.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   765
               Left            =   1740
               TabIndex        =   12
               Top             =   540
               Width           =   4020
            End
            Begin VB.Label LblWritten 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Written By:   Randy Giese"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   360
               TabIndex        =   11
               Top             =   2655
               Width           =   2430
            End
            Begin VB.Label LblEmail 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Email:  WebMaster@RandyGrams.Com"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   360
               TabIndex        =   10
               Top             =   3060
               Width           =   3690
            End
         End
         Begin VB.PictureBox picBottom 
            BackColor       =   &H00D8E9EC&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   1950
            Left            =   300
            ScaleHeight     =   1950
            ScaleWidth      =   7200
            TabIndex        =   4
            Top             =   4100
            Width           =   7200
            Begin VB.Label LblFree 
               Alignment       =   2  'Center
               BackColor       =   &H00D8E9EC&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H00000000&
               Height          =   900
               Left            =   300
               TabIndex        =   7
               Top             =   420
               Width           =   6600
            End
            Begin VB.Label LblSend 
               Alignment       =   2  'Center
               BackColor       =   &H00D8E9EC&
               BackStyle       =   0  'Transparent
               Caption         =   "Please send any Comments, Questions or Suggestions to:  WebMaster@RandyGrams.Com"
               ForeColor       =   &H00FF0000&
               Height          =   600
               Left            =   300
               TabIndex        =   6
               Top             =   1290
               Width           =   6600
            End
            Begin VB.Label LblE_Mail_Ware 
               Alignment       =   2  'Center
               BackColor       =   &H00D8E9EC&
               BackStyle       =   0  'Transparent
               Caption         =   "This program is 'E-Mail-Ware'."
               ForeColor       =   &H00FF0000&
               Height          =   300
               Left            =   600
               TabIndex        =   5
               Top             =   120
               Width           =   6000
            End
         End
         Begin VB.CommandButton CmdOK 
            BackColor       =   &H00D8E9EC&
            Caption         =   "OK"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Index           =   0
            Left            =   6300
            TabIndex        =   3
            Top             =   6350
            Width           =   1200
         End
         Begin VB.CommandButton CmdOK 
            BackColor       =   &H00D8E9EC&
            Caption         =   "OK"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Index           =   1
            Left            =   300
            TabIndex        =   2
            Top             =   6350
            Width           =   1200
         End
         Begin VB.Label LblRG_Com 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BackStyle       =   0  'Transparent
            Caption         =   "www.RandyGrams.com"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   500
            Left            =   1950
            TabIndex        =   17
            Top             =   6350
            Width           =   3900
         End
      End
   End
End
Attribute VB_Name = "RG_About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    The About screen displays information about this             ****
'   ****    RandyGrams program including the current Version number.     ****
'   ****                                                                 ****
'   ****    Please visit:  --->  www.RandyGrams.com  <---                ****
'   ****                                                                 ****
'   ****    for more free programs.                                      ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'
'   RandyGrams Comments - Left Align the above comments.

Private Const mc_strModuleID   As String = "Puzzle_Size."

Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Private Sub CmdOK_Click(Index As Integer)
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Return (or Go) to the Main program.                          ****
'   ****                                                                 ****
'   *************************************************************************
'
    On Error GoTo ErrorHandler
    Unload Me
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "CmdOK_Click(" & Index & ")"

End Sub

Private Sub Form_Activate()
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    A special  "Thank You"  to  "David Leatherman"  for          ****
'   ****    helping me get rid of an unneeded Timer.                     ****
'   ****                                                                 ****
'   ****    Thank You David Leatherman!                                  ****
'   ****                                                                 ****
'   *************************************************************************
'
    On Error GoTo ErrorHandler
'   Locate the Cursor in the center of the "OK" button.
    With CmdOK(1)
        SetCursorPos (Me.Left + PicBlack.Left + picBlue.Left + .Left + (.Width \ 2)) \ Screen.TwipsPerPixelX, (Me.Top + PicBlack.Top + picBlue.Top + .Top + (.Height \ 2)) \ Screen.TwipsPerPixelY
        .SetFocus
    End With  'CmdOK(1)

    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "Form_Activate()"

End Sub

Private Sub Form_Load()

Dim strAboutYear               As String
'
'   Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks
'   Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks
'   Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks
'   Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks
'   Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks
'   Thanks                                                                Thanks
'   Thanks    Many, many Thanks to Roger Gilchrist, who not only went     Thanks
'   Thanks    way out of his way but also well above and beyond the       Thanks
'   Thanks    call of duty, to help someone he has never met!!!           Thanks
'   Thanks                                                                Thanks
'   Thanks    I owe you a great debt of gratitude Roger!                  Thanks
'   Thanks                                                                Thanks
'   Thanks    Please check out the best program I have ever seen on       Thanks
'   Thanks    PSC:                                                        Thanks
'   Thanks                                                                Thanks
'   Thanks    "CodeFixer"                                                 Thanks
'   Thanks    By:  Roger Gilchrist.                                       Thanks
'   Thanks                                                                Thanks
'   Thanks    The latest version as of this writing can be found at:      Thanks
'   Thanks                                                                Thanks
'   Thanks    http://www.planetsourcecode.com/vb/scripts/ShowCode         Thanks
'   Thanks    .asp?txtCodeId=53297&lngWId=1                               Thanks
'   Thanks                                                                Thanks
'   Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks
'   Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks
'   Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks
'   Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks
'   Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks Thanks
'
    On Error GoTo ErrorHandler
    Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
    strAboutYear = "RandyGrams (" & Year(Now) & ")"        'Put the Year inside the ()
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Print the Current Year on the RG-Logo.                       ****
'   ****                                                                 ****
'   *************************************************************************
'
    With Logo
        .CurrentX = (.ScaleWidth - .TextWidth(strAboutYear)) / 2
        .CurrentY = .ScaleHeight - 500
        Logo.Print strAboutYear
    End With  'Logo

    lblCopyright.Caption = "Copyright: " & strAboutYear    'Add the Year to the Copyright
    lblCompanyName.Caption = strAboutYear   'Add the Year to the Company Name
    lblProductName.Caption = App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    LblFree.Caption = "It is free and you may feel free to pass it on to your friends and E-Mail aquaintances.  All I ask is that you send me a brief E-Mail to let me know what you think of:  " & App.Title & "!  Thank you!"
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "Form_Load()"

End Sub
