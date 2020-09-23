VERSION 5.00
Begin VB.Form Puzzle_Main 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   4935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5580
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
   Icon            =   "Puzzle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   329
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   372
   Begin VB.PictureBox picRestart 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   900
      Left            =   4215
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   74
      TabIndex        =   12
      Tag             =   "0"
      Top             =   75
      Visible         =   0   'False
      Width           =   1110
      Begin VB.CommandButton cmdRestart 
         BackColor       =   &H00E0E0E0&
         Caption         =   "End!"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   510
         Width           =   900
      End
      Begin VB.CommandButton cmdRestart 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Start!"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   60
         Width           =   900
      End
   End
   Begin VB.PictureBox picFileName 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   900
      Left            =   2985
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   74
      TabIndex        =   10
      Tag             =   "0"
      Top             =   75
      Visible         =   0   'False
      Width           =   1110
      Begin VB.Label lblFileName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   345
         Left            =   465
         TabIndex        =   11
         Top             =   75
         Width           =   120
      End
   End
   Begin VB.PictureBox picFileScan 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3720
      Left            =   75
      ScaleHeight     =   244
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   357
      TabIndex        =   4
      Top             =   1095
      Width           =   5415
      Begin VB.CommandButton cmdFileScan 
         Caption         =   "Start!"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4380
         TabIndex        =   8
         Top             =   120
         Width           =   900
      End
      Begin VB.DriveListBox drvFileScan 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   1800
      End
      Begin VB.FileListBox filFileScan 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2820
         Left            =   2760
         Pattern         =   "*.Bmp;*.Gif;*.Jpg"
         TabIndex        =   6
         Top             =   570
         Width           =   2520
      End
      Begin VB.DirListBox dirFileScan 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2970
         Left            =   120
         TabIndex        =   5
         Top             =   570
         Width           =   2520
      End
      Begin VB.Label lblSelect 
         AutoSize        =   -1  'True
         Caption         =   "Select a folder containing images, then Click the Start button."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1920
         TabIndex        =   9
         Top             =   135
         Width           =   7695
      End
   End
   Begin VB.PictureBox picPuzzle_Main 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   2325
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   3
      Top             =   75
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picDisplay 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   900
      Left            =   1095
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   74
      TabIndex        =   1
      Tag             =   "0"
      Top             =   75
      Visible         =   0   'False
      Width           =   1110
      Begin VB.Label lblPuzzle_Main 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   345
         Left            =   465
         TabIndex        =   2
         Top             =   75
         Width           =   120
      End
   End
   Begin VB.PictureBox picBox 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   900
      Index           =   0
      Left            =   75
      ScaleHeight     =   900
      ScaleWidth      =   900
      TabIndex        =   0
      Tag             =   "0"
      Top             =   75
      Visible         =   0   'False
      Width           =   900
   End
End
Attribute VB_Name = "Puzzle_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    Puzzle                                                       ****
'   ****                                                                 ****
'   ****    Written by:    Randy Giese    (2002/11/16)                   ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    "Puzzle" uses the Graphic files on your computer to          ****
'   ****    create a Puzzle. It randomly selects a file from the         ****
'   ****    directory of your choice and temporarily makes it your       ****
'   ****    WallPaper.  You may also select many different graphic       ****
'   ****    images from anywhere on your system to use in the puzzle.    ****
'   ****                                                                 ****
'   ****    "Puzzle" will display your picture for 4 seconds, then it    ****
'   ****    breaks the picture into boxes and mixes them up. All you     ****
'   ****    have to do is to put the pieces back where they belong.      ****
'   ****    Each time you complete a Puzzle, you will have the           ****
'   ****    opportunity to start a new one with more pieces.             ****
'   ****                                                                 ****
'   ****    You can move the Puzzle pieces by clicking on one piece,     ****
'   ****    then another piece.  When you press the "Enter" key, the     ****
'   ****    two Puzzle pieces that you clicked will swap places.         ****
'   ****    Continue swapping Puzzle pieces until the Puzzle is          ****
'   ****    solved.                                                      ****
'   ****                                                                 ****
'   ****    If you run into trouble, you may press the right mouse       ****
'   ****    button for several options, including "Hint".  When you      ****
'   ****    click "Hint", it will tell you how many pieces are still     ****
'   ****    out of order and ask if you want to see which pieces they    ****
'   ****    are.  If you choose "Yes", it will hi-lite all of the        ****
'   ****    pieces that are still out of order.                          ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    July, 2007 - Added the ReSizer Module (ReSizer.Bas) which    ****
'   ****    is used for loading images and resizing them.  It can        ****
'   ****    load all the picture formats supported by LoadPicture        ****
'   ****    plus the PNG file format.  This results in the user being    ****
'   ****    able to select almost any graphic image for the puzzle,      ****
'   ****    regardless of size or format.                                ****
'   ****                                                                 ****
'   ****    Removed the separate "Help" form and Internet Browser and    ****
'   ****    replaced them with the ShellExecute API.  This resulted      ****
'   ****    in a code reduction and now uses the users browser rather    ****
'   ****    than forcing MSIE on them.                                   ****
'   ****                                                                 ****
'   ****    I also added the GetWP Sub-Routine.  This retrieves the      ****
'   ****    users current WallPaper and saves it, then just before       ****
'   ****    the program is exited, the original WallPaper is             ****
'   ****    restored.  This works with Windows XP, but I'm afraid I      ****
'   ****    don't know if it will work on any other OS's.                ****
'   ****                                                                 ****
'   *************************************************************************
'   ****                                                                 ****
'   ****    July 13, 2007 - Added the "Solve" process to the Menu.       ****
'   ****                                                                 ****
'   *************************************************************************
'   ****                                                                 ****
'   ****    August 22, 2007 - Added the "BackGround" form.  The          ****
'   ****    "BackGround" form is full-screen size and ensures that       ****
'   ****    the puzzle pieces will not get lost amid desktop clutter.    ****
'   ****    I originally had the WallPaper set to "Stretch" mode to      ****
'   ****    fill the entire screen but if a graphic is a size that is    ****
'   ****    a different ratio than the screen, distortion occurred.      ****
'   ****    This way eliminates that distortion while still allowing     ****
'   ****    the image to be displayed at a large size.                   ****
'   ****                                                                 ****
'   ****    Note: Either the picture length or width will be screen      ****
'   ****    size depending on the pictures proportions.                  ****
'   ****                                                                 ****
'   ****    Note: Very small graphic images will still distort           ****
'   ****    somewhat.                                                    ****
'   ****                                                                 ****
'   *************************************************************************
'   ****                                                                 ****
'   ****    August 30, 2007 - Added the "lngTimes" counter in an         ****
'   ****    effort to eliminate multiple selections of a single          ****
'   ****    picture.  (This is in "Sub Begin" in "Select Case            ****
'   ****    lngNewPuzzle".                                               ****
'   ****                                                                 ****
'   *************************************************************************
'   ****                                                                 ****
'   ****    August 30, 2007 - Added the "ShowFileName" Sub-Routine.      ****
'   ****    It gives the user the option to display the picture's        ****
'   ****    FileName.                                                    ****
'   ****                                                                 ****
'   *************************************************************************
'   ****                                                                 ****
'   ****    August 31, 2007 - Changed the "ShowFileName" Menu Option     ****
'   ****    to a Check Button.  (Default is "On")                        ****
'   ****                                                                 ****
'   *************************************************************************
'   ****                                                                 ****
'   ****    September 04, 2007 - I introduced the "blnRandom"            ****
'   ****    variable as a means of starting the "Begin" Sub-Routine.     ****
'   ****    I previously had "Begin" being accessed from 2 or 3          ****
'   ****    different locations in the program.  This resulted in        ****
'   ****    some Sub-Routines not being allowed to complete and          ****
'   ****    therefore left some ugly loose ends.  This should make       ****
'   ****    somewhat cleaner running code.                               ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'
'   RandyGrams Comments - Left Align the above comments.

Option Explicit

'** Variables **
Private blnFileScan                  As Boolean
Private blnRandom                    As Boolean
Private blnShowFileName              As Boolean
Private blnSolve                     As Boolean

Private curEnd                       As Currency
Private curStart                     As Currency
Private curTimerFreq                 As Currency

Private lngArray()                   As Long
Private lngBoxLeft()                 As Long
Private lngBoxTop()                  As Long
Private lngClickCount                As Long
Private lngHigh                      As Long
Private lngNewPuzzle                 As Long
Private lngOrigMode                  As Long
Private lngPauseTime                 As Long
Private lngWide                      As Long

Private strBmp                       As String
Private strName                      As String

Private Const HKEY_CURRENT_USER      As Long = &H80000001
Private Const mc_strModuleID         As String = "Puzzle."
Private Const MF_POPUP               As Long = &H10&
Private Const MF_SEPARATOR           As Long = &H800&
Private Const MF_STRING              As Long = &H0&
Private Const QS_HOTKEY              As Long = &H80
Private Const QS_KEY                 As Long = &H1
Private Const QS_MOUSEBUTTON         As Long = &H4
Private Const QS_MOUSEMOVE           As Long = &H2
Private Const QS_PAINT               As Long = &H20
Private Const QS_POSTMESSAGE         As Long = &H8
Private Const QS_SENDMESSAGE         As Long = &H40
Private Const QS_TIMER               As Long = &H10
Private Const QS_ALLINPUT            As Long = (QS_SENDMESSAGE Or QS_PAINT Or QS_TIMER Or QS_POSTMESSAGE Or QS_MOUSEBUTTON Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)
Private Const REG_SZ                 As Long = 1
Private Const SPI_GETDESKWALLPAPER   As Long = 115
Private Const SW_MAXIMIZE            As Long = 3
Private Const TPM_LEFTALIGN          As Long = &H0&
Private Const TPM_RETURNCMD          As Long = &H100&
Private Const TPM_RIGHTBUTTON        As Long = &H2&

Private Type POINTAPI
    lngX                             As Long
    lngY                             As Long
End Type

'** Declarations **
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetQueueStatus Lib "user32" (ByVal fuFlags As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal hWnd As Long, ByVal lptpm As Any) As Long

Private Declare Sub InitCommonControls Lib "comctl32" ()

Private Sub Begin()

Dim lngTimes                         As Long
Dim lngList                          As Long
Dim lngStop                          As Long
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    The user has clicked the mouse to start the Puzzle           ****
'   ****    process.                                                     ****
'   ****                                                                 ****
'   *************************************************************************
'
    On Error GoTo ErrorHandler

    blnRandom = False
    blnSolve = False
    lngPauseTime = 4000
    lngStop = GetTickCount + lngPauseTime
    Me.Move 0, 0, Screen.Width, Screen.Height
    picRestart.Visible = False
    Randomize
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Give the player an opportunity to choose the Picture.        ****
'   ****                                                                 ****
'   *************************************************************************
'
    Select Case lngNewPuzzle
        Case 1               'Random selected...
'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    I replaced the single line of code below with everything     ****
'   ****    else in "Case 1".  This may seem wasteful; replacing 1       ****
'   ****    line of code with so many lines but this way a picture is    ****
'   ****    not reused unless it is a very small Directory.  The         ****
'   ****    single line was truly random and therefore could select      ****
'   ****    the same picture more than once.                             ****
'   ****                                                                 ****
'   ****    strName=dirFileScan.Path & "\" & filFileScan.List(filFileScan.ListCount*Rnd)
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'
'   RandyGrams Comments - Left Align the above comments.

            Do While lngTimes < 13
'           Select a random file from the Directory.
            lngList = (filFileScan.ListCount - 1) * Rnd

'           See if this file has been selected before.
            If lngArray(lngList) <> 1 Then
'               If not, then flag it and clear the counter.
                lngArray(lngList) = 1
                lngTimes = 0
                Exit Do
            Else
'               Count the number of consecutive times that
'               we try to find an unused picture.
                lngTimes = lngTimes + 1
'               If it's more than a dozen, we might as well
'               clear the counter and start over.
'               This would likely happen if there are only a
'               couple of pictures in this Directory.
                If lngTimes > 12 Then
                    lngTimes = 0
                    ReDim lngArray(0) As Long
                    DoEvents
                    ReDim lngArray(filFileScan.ListCount) As Long
                End If
            End If
            Loop

            strName = dirFileScan.Path & "\" & filFileScan.List(lngList)

        Case 2
            lngMsgResp = MsgBox("Do you want to use the same picture for your Puzzle again?     ", vbQuestion + vbYesNo, "     Select a picture!")

            If lngMsgResp = vbYes Then
                strName = filFileScan.Path & "\" & filFileScan.FileName
            Else
                Me.Move 0, 0, Screen.Width, Screen.Height

                blnFileScan = False
                picFileScan.Visible = True

                Do While blnFileScan = False
                    DoEvents
                Loop
            End If           'lngMsgResp
        Case 3
            lngNewPuzzle = 2
            strName = filFileScan.Path & "\" & filFileScan.FileName
    End Select               'lngNewPuzzle

    ClearOldBoxes
    picFileName.Visible = False
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Temporarily make this Bmp the current WallPaper.             ****
'   ****                                                                 ****
'   *************************************************************************
'
    strBmp = App.Path & "\PuzzleRG.Bmp"
'   Resize the picture to the largest usable size.
    GDI
    SavePicture Me.Picture, strBmp
    Me.Move 0, 0, lngWide * Screen.TwipsPerPixelX, lngHigh * Screen.TwipsPerPixelY

    If blnShowFileName Then
        blnShowFileName = Not blnShowFileName
        ShowFileName
    End If

'   Dim lngWPMode                        As Long
'   lngWPMode = 0    The wallpaper is centered.
'   lngWPMode = 1    The wallpaper is tiled.
'   lngWPMode = 2    The wallpaper is stretched to fit the entire screen.
'   SetWallpaper strBmp, lngWPMode
    SetWallpaper strBmp, 0

    lngPerRow = CLng(Sqr(lngBoxes))
    lngPicTo = lngBoxes - 1

    lngSH = lngHigh \ lngPerRow
    lngSW = lngWide \ lngPerRow

    ReDim lngBoxLeft(lngBoxes) As Long
    ReDim lngBoxTop(lngBoxes) As Long
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Create the boxes on the fly.  This way I don't need to       ****
'   ****    know ahead of time how many boxes to create.                 ****
'   ****                                                                 ****
'   *************************************************************************
'
    For lngLoop1 = 1 To lngPicTo
        Load picBox(lngLoop1)
        picBox(lngLoop1).Tag = lngLoop1
    Next lngLoop1
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Make the Picture Boxes visible and set their size.           ****
'   ****    Then fill them with the Picture.                             ****
'   ****                                                                 ****
'   *************************************************************************
'
    For lngLoop1 = 0 To lngPerRow - 1
        For lngLoop2 = 0 To lngPerRow - 1
            With picBox((lngLoop1 * lngPerRow) + lngLoop2)
                .Visible = True             'This makes the Picture Boxes visible
                .Move lngLoop2 * lngSW, lngLoop1 * lngSH, lngSW, lngSH    'This sets their size
                Set .Picture = hDCToPicture(GetDC(0), lngLoop2 * lngSW, lngLoop1 * lngSH, lngSW, lngSH) 'This draws the Picture
            End With
        Next lngLoop2
    Next lngLoop1
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Display the warning message to "Study the Picture!".         ****
'   ****                                                                 ****
'   *************************************************************************
'
    Me.Picture = LoadPicture("")
    Display
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Pause for 4 seconds so the user may study the Picture.       ****
'   ****                                                                 ****
'   *************************************************************************
'
    Do While GetTickCount < lngStop
        DoEvents
    Loop

    picDisplay.Visible = False

Scramble:
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Scramble the Picture.                                        ****
'   ****                                                                 ****
'   *************************************************************************
'
    For lngLoop1 = 0 To lngPicTo
        lngSwap1 = lngLoop1
Again:
        lngSwap2 = Int(lngBoxes * Rnd)

        If lngSwap1 = lngSwap2 Then
            GoTo Again
        End If

        Swap
        DoEvents
    Next lngLoop1
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Check to see if the Picture is in order. If it is, go        ****
'   ****    back and Scramble the picture again.                         ****
'   ****                                                                 ****
'   *************************************************************************
'
    CheckIt

    If lngBoxesOOO = 0 Then
        GoTo Scramble
    End If

    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "Begin()"
    Form_Unload False

End Sub

Private Sub CheckIt()
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Scan the Puzzle to see if it's in the correct order.         ****
'   ****                                                                 ****
'   *************************************************************************
'
    On Error GoTo ErrorHandler
    lngBoxesOOO = 0          'The number of Boxes Out-Of-Order

    For lngLoop2 = 0 To lngPicTo
        If picBox(lngLoop2).Tag <> picBox(lngLoop2).Index Then
            lngBoxesOOO = lngBoxesOOO + 1
        End If
    Next lngLoop2

    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "CheckIt()"

End Sub

Private Sub ClearOldBoxes()
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Clear the old boxes from the screen.                         ****
'   ****                                                                 ****
'   *************************************************************************
'
    On Error GoTo ErrorHandler

    picDisplay.Visible = False
    picBox(0).Visible = False
    picBox(0).Tag = 0

    For lngLoop1 = 1 To lngPicTo
        Unload picBox(lngLoop1)
    Next lngLoop1

    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "ClearOldBoxes()"

End Sub

Private Sub cmdFileScan_Click()

'   The user clicked the "Start" button so the file selection is random...
    On Error GoTo ErrorHandler
    ReDim lngArray(filFileScan.ListCount) As Long
    picFileScan.Visible = False
'   Randomly select a file...
    blnFileScan = True
    blnRandom = True
    RandomPics
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "cmdFileScan_Click()"

End Sub

Private Sub cmdRestart_Click(Index As Integer)

    Select Case Index
        Case 0
            cmdFileScan_Click
        Case 1
            DeleteBmp        'Delete the temporary Bmp file.
            SetWallpaper strOldPaper, lngOrigMode
            End
    End Select  'Index

End Sub

Private Sub DeleteBmp()
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Delete the temporary Bmp file.                               ****
'   ****                                                                 ****
'   *************************************************************************
'
    On Error GoTo ErrorHandler
    DeleteFile strBmp        'Delete the Temporary Bmp
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "DeleteBmp()"

End Sub

Private Sub dirFileScan_Change()

    On Error GoTo ErrorHandler

    With dirFileScan
        ChDir .Path
        filFileScan.Path = .Path
    End With  'dirFileScan

    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "dirFileScan_Change()"

End Sub

Private Sub Display()

    On Error GoTo ErrorHandler
    strMsg = "Study the Picture!" & vbCr & "  Current Puzzle size is:  --->  " & Sqr(lngBoxes) & "*" & Sqr(lngBoxes) & "  "

    With lblPuzzle_Main
        .Caption = strMsg
        .Move 6, 6
    End With  'lblPuzzle_Main

'   Make sure the Display is correctly sized and centered.
    With picDisplay
        .Move 0, 0, lblPuzzle_Main.Width + 14, lblPuzzle_Main.Height + 14
        .Move ((Me.Width \ Screen.TwipsPerPixelX) - .Width) \ 2, ((Me.Height \ Screen.TwipsPerPixelY) - .Height) \ 2
        .Visible = True
    End With  'picDisplay

    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "Display()"

End Sub

Private Sub Done()
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Issue "Congratulations!" message, clear the "Old" Puzzle,    ****
'   ****    and start a new one.                                         ****
'   ****                                                                 ****
'   *************************************************************************
'
    On Error GoTo ErrorHandler
    DeleteBmp 'Delete the temporary Bmp file.
    Mess.Show vbModal, Me
    blnRandom = blnFileScan

    If lngNewPuzzle > 1 Then
        lngNewPuzzle = 2
    End If
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Give the player an opportunity to choose the Puzzle          ****
'   ****    size!".                                                      ****
'   ****                                                                 ****
'   *************************************************************************
'
    Select Case lngMsgResp
        Case 0  'Yes
            Puzzle_Size.Show vbModal, Me
        Case 2  'Cancel
            blnRandom = False
            picRestart.Move (Me.Left + Me.Width) \ Screen.TwipsPerPixelX - picRestart.Width, 0
            picRestart.Visible = True
    End Select  'lngMsgResp

    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "Done()"

End Sub

Private Sub drvFileScan_Change()

    On Error GoTo ErrorHandler

    With dirFileScan
        ChDrive .Path
        .Path = drvFileScan.Drive
        .Refresh
    End With  'dirFileScan

    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "drvFileScan_Change()"

End Sub

Private Sub filFileScan_Click()

    On Error GoTo ErrorHandler
    lngNewPuzzle = 3
    picFileScan.Visible = False
    cmdFileScan_Click
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "filFileScan_Click()"

End Sub

Private Sub Form_Initialize()

    On Error GoTo ErrorHandler
    InitCommonControls
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "Form_Initialize()"

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    When you press the "Enter" key, the last 2 boxes that        ****
'   ****    were clicked will be swapped.  Then the Puzzle will be       ****
'   ****    checked to see if it's in the correct order.                 ****
'   ****                                                                 ****
'   *************************************************************************
'
    On Error GoTo ErrorHandler

    If KeyAscii = 13 Then
        If lngClickCount = 2 Then
            lngClickCount = 0
            Swap

'           Remove the Borders and make the selected pieces "Full-Size" again.
            picBox(lngSwap1).BorderStyle = 0
            picBox(lngSwap2).BorderStyle = 0
            picBox(lngSwap1).Move lngBoxLeft(lngSwap2), lngBoxTop(lngSwap2), lngSW, lngSH
            picBox(lngSwap2).Move lngBoxLeft(lngSwap1), lngBoxTop(lngSwap1), lngSW, lngSH

'           Check to see if the Puzzle pieces are in the correct order.
            CheckIt

            If lngBoxesOOO = 0 Then
                blnSolve = False
                Done
            End If
        End If
    Else
        If blnSolve = True Then
            lngMsgResp = MsgBox("Do you wish to Abort the Solve process?     ", vbQuestion + vbYesNo, "         Abort Solve!")

            If lngMsgResp = vbYes Then
                lngMsgResp = 0
                blnSolve = False
            End If           'lngMsgResp
        End If
    End If

    Do While blnRandom = True
        Begin
    Loop

    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "Form_KeyPress(" & KeyAscii & ")"

End Sub

Private Sub Form_Load()
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Save the name of the current WallPaper and set some          ****
'   ****    initial values.                                              ****
'   ****                                                                 ****
'   *************************************************************************
'
    On Error GoTo ErrorHandler

    BackGround.Move 0, 0, Screen.Width, Screen.Height
    BackGround.Show

    Me.Show
    Me.Move 0, 0, Screen.Width, Screen.Height

    ChDrive App.Path
    ChDir App.Path

    lngBoxes = 4
    lngNewPuzzle = 1
    blnShowFileName = True
    QueryPerformanceFrequency curTimerFreq  'Get the timer frequency

    GetWP
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "Form_Load()"

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo ErrorHandler
    PopMenu Button
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "Form_MouseDown(" & Button & ", " & Shift & ", " & X & ", " & Y & ")"

End Sub

Private Sub Form_Resize()
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Adjust the controls, etc when the user resizes the form.     ****
'   ****                                                                 ****
'   *************************************************************************
'
    On Error GoTo ErrorHandler

    With Me
        lngW = .Width
        lngH = .Height
    End With  'Me

'   Make sure all the File tools are properly sized.
    picFileScan.Move 5, 5, Me.Width - 10, Me.Height - 10
    dirFileScan.Move drvFileScan.Left, drvFileScan.Top + drvFileScan.Height + drvFileScan.Top, Me.ScaleWidth \ 3, Me.ScaleHeight - dirFileScan.Top - lblSelect.Height - 10
    filFileScan.Move dirFileScan.Left + dirFileScan.Width + drvFileScan.Left, dirFileScan.Top, Me.ScaleWidth \ 2, dirFileScan.Height
    cmdFileScan.Left = filFileScan.Left + filFileScan.Width - cmdFileScan.Width
    picFileScan.Move drvFileScan.Left, drvFileScan.Top, cmdFileScan.Left + cmdFileScan.Width + 12, dirFileScan.Top + dirFileScan.Height + 12
    picFileScan.Move (Me.ScaleWidth - picFileScan.Width) \ 2, (Me.ScaleHeight - picFileScan.Height) \ 2
    lblSelect.Left = (picFileScan.Width - lblSelect.Width) \ 2
    picRestart.Move Me.Left + Me.Width - picRestart.Width - 4, 4, (cmdRestart(0).Width * 2) + 12, cmdRestart(0).Height + 8
    cmdRestart(0).Move 4, 4
    cmdRestart(1).Move cmdRestart(0).Left + cmdRestart(0).Width + 4, cmdRestart(0).Top

    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "Form_Resize()"

End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error GoTo ErrorHandler

    DeleteBmp 'Delete the temporary Bmp file.
    SetWallpaper strOldPaper, lngOrigMode
    End

    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "Form_Unload(" & Cancel & ")"

End Sub

Private Sub GDI()

Dim Token                            As Long
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    This Sub-Routine makes a copy of the selected image,         ****
'   ****    compressing or extending the image equal to the size of      ****
'   ****    the screen, but retains the original size ratio.             ****
'   ****                                                                 ****
'   *************************************************************************
'
'   Initialize GDI+
    On Error GoTo ErrorHandler
    Token = InitGDIPlus

    With picPuzzle_Main
        .Picture = LoadPicture("")
        .Picture = LoadPicture(strName)

        If .Width > .Height / 0.75 Then
            lngWide = Screen.Width \ Screen.TwipsPerPixelX
            lngHigh = lngWide * (.Height / .Width)
        Else
            lngHigh = Screen.Height \ Screen.TwipsPerPixelY
            lngWide = lngHigh / (.Height / .Width)
        End If

'       Resize already loaded picture (vbPicTypeIcon is not supported yet)
        Me.Picture = ReSize(.Picture.Handle, .Picture.Type, lngWide, lngHigh, vbBlack, True)
    End With  'picPuzzle_Main

'   Free GDI+
    FreeGDIPlus Token

    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "GDI()"
    Form_Unload False

End Sub

Private Sub GetWP()

Dim lngKey                           As Long
Dim lngSubKey                        As Long
Dim strBufferData                    As String
'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    This Sub-Routine gets the current Wallpaper from the         ****
'   ****    Registry.                                                    ****
'   ****                                                                 ****
'   ****    This works with MS-Windows XP but hasn't been tested on      ****
'   ****    other Operating Systems.                                     ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'
    On Error GoTo ErrorHandler
    strBufferData = Space$(256)
    RegOpenKey HKEY_CURRENT_USER, "Control Panel", lngKey
    RegOpenKey lngKey, "Desktop", lngSubKey
'   Save the Current Wallpaper
    RegQueryValueEx lngSubKey, "ConvertedWallpaper", 0, REG_SZ, strBufferData, Len(strBufferData)
    strOldPaper = strBufferData
'   Save the Current Wallpaper Mode - Centered(0), Tiled(1) or Stretched(2)
    RegQueryValueEx lngSubKey, "WallpaperStyle", 0, REG_SZ, strBufferData, Len(strBufferData)
    lngOrigMode = CLng(strBufferData)

    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "GetWP()"

End Sub

Private Sub HiLiteBox(ByVal lngIndex As Long)
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Hi-Lite the selected box.                                    ****
'   ****                                                                 ****
'   *************************************************************************
'
    On Error GoTo ErrorHandler
    lngClickCount = lngClickCount + 1

'   If more than 2 boxes were clicked
    If lngClickCount > 2 Then
        lngClickCount = 2
        lngSwap3 = lngSwap2
'       Remove the Border and reset the first one back to original size
        picBox(lngSwap3).BorderStyle = 0
        picBox(lngSwap3).Move lngBoxLeft(lngSwap3), lngBoxTop(lngSwap3), lngSW, lngSH
    End If

    lngSwap2 = lngSwap1
    lngSwap1 = lngIndex

    With picBox(lngSwap1)
        lngBoxLeft(lngSwap1) = .Left
        lngBoxTop(lngSwap1) = .Top
'       Make the selected Box a little smaller
        .Move .Left + (.Width * 0.025), .Top + (.Height * 0.025), .Width * 0.95, .Height * 0.95
'       Add the Border
        .BorderStyle = 1
    End With  'picBox(lngSwap1)

    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "HiLiteBox(" & lngIndex & ")"

End Sub

Private Sub Hint()
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Hi-Lite the pieces that are Out-Of-Order.                    ****
'   ****                                                                 ****
'   *************************************************************************
'
    On Error GoTo ErrorHandler
    CheckIt
    lngMsgResp = MsgBox("There are " & lngBoxesOOO & " pieces still out of order." & vbCr & "Do you want me to show you which ones they are?", vbYesNo, "    Puzzle Hint!    By:  Randy Giese")
    lngPauseTime = 2000 + (lngBoxes * 50)

    If lngMsgResp = vbYes Then
        lngMsgResp = 0
        MsgBox "OK, get ready because I'm only going to show them for " & lngPauseTime / 1000 & " seconds." & vbCr & vbCr & _
        "Press 'OK' when you are ready to see them.", vbOKOnly, "    Puzzle Hint!    By:  Randy Giese"  '  Display message.

        For lngLoop2 = 0 To lngPicTo
            With picBox(lngLoop2)
                If .Tag <> .Index Then
                    .BorderStyle = 1
                    .Height = lngSH * 0.9
                    .Width = lngSW * 0.9
                End If
            End With         'picBox(lngLoop2)
        Next lngLoop2
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Pause for 2+ seconds so the user may study the Picture.      ****
'   ****                                                                 ****
'   *************************************************************************
'
'   RandyGrams Comments - Left Align the above comments.

        Pause

        For lngLoop2 = 0 To lngPicTo
            With picBox(lngLoop2)
                .BorderStyle = 0
                .Height = lngSH
                .Width = lngSW
            End With         'Box(lngLoop2)
        Next lngLoop2
    End If

    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "Hint()"

End Sub

Private Sub picBox_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo ErrorHandler

    If Button = vbRightButton Then
'       Use a PopupMenu for Help, etc.
        PopMenu Button
    Else
'       Hi-Lite the selected box.
        HiLiteBox (CLng(Index))
    End If

    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "picBox_MouseDown(" & Button & ", " & Shift & ", " & X & ", " & Y & ")"

End Sub

Private Sub Pause()

Dim curFreq                          As Currency

    On Error GoTo ErrorHandler
    curFreq = curTimerFreq / 1000
    curFreq = curFreq * lngPauseTime

    QueryPerformanceCounter curStart        'Get the start time
    curStart = curStart + curFreq

    Do While curEnd < curStart
        QueryPerformanceCounter curEnd      'Get the end time
        If GetQueueStatus(QS_ALLINPUT) Then 'Yield for other processing.
            DoEvents
        End If
    Loop

    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "Pause()"

End Sub

Private Sub PopMenu(ByVal lngButton As Long)

Dim lngMainMenu                      As Long
Dim lngSub_Menu                      As Long
Dim lngTPM                           As Long
Dim pntXorY                          As POINTAPI
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Use a PopupMenu for Help, etc.                               ****
'   ****                                                                 ****
'   *************************************************************************
'
    On Error GoTo ErrorHandler
    If lngButton = vbRightButton Then
'       Create the Menu addresses.   The variable contains the address.
        lngMainMenu = CreatePopupMenu()
        lngSub_Menu = CreatePopupMenu()
'       Save the current Cursor Position.
        GetCursorPos pntXorY

'       This is the Main-Menu.
        AppendMenu lngMainMenu, MF_STRING, 1, "About"
        AppendMenu lngMainMenu, MF_STRING, 2, "Help"
        AppendMenu lngMainMenu, MF_SEPARATOR, 3, ByVal 0&
        AppendMenu lngMainMenu, MF_STRING, 4, "Hint"
        AppendMenu lngMainMenu, MF_STRING, 5, "Solve"
        AppendMenu lngMainMenu, MF_SEPARATOR, 6, ByVal 0&

'       The Popup Sub Menu
        AppendMenu lngMainMenu, MF_POPUP, lngSub_Menu, "New Puzzle"
        AppendMenu lngSub_Menu, Hex(Abs(lngNewPuzzle = 1) * 8), 7, "Random Selected"
        AppendMenu lngSub_Menu, Hex(Abs(lngNewPuzzle = 2) * 8), 8, "User Selected"
        AppendMenu lngSub_Menu, MF_STRING, 9, "Different Folder"
        AppendMenu lngMainMenu, MF_STRING, 10, "Show Puzzle"
        AppendMenu lngMainMenu, Hex(Abs(blnShowFileName) * 8), 11, "Show FileName"
        AppendMenu lngMainMenu, MF_SEPARATOR, 12, ByVal 0&
        AppendMenu lngMainMenu, MF_STRING, 13, "Quit"

'       The  TrackPopupMenuEx  function displays a shortcut menu at the specified location and tracks the selection of items on the shortcut menu.
        lngTPM = TrackPopupMenuEx(lngMainMenu, TPM_LEFTALIGN Or TPM_RETURNCMD Or TPM_RIGHTBUTTON, pntXorY.lngX, pntXorY.lngY, Me.hWnd, ByVal 0&)

'       The DestroyMenu function destroys the specified menu and frees any memory that the menu occupies.
        DestroyMenu lngMainMenu

'       This is where you control the Menu buttons.
        Select Case lngTPM
        Case 1
            RG_About.Show vbModal, Me
        Case 2
            ShellExecute Me.hWnd, "open", App.Path & "\Puzzle.Html", vbNullString, vbNullString, SW_MAXIMIZE
        Case 4
            Hint
        Case 5
            blnSolve = True
            Solve
        Case 7
            blnRandom = True
            lngNewPuzzle = 1
        Case 8 To 9
            Me.Move 0, 0, Screen.Width, Screen.Height

            blnFileScan = False
            picFileScan.Visible = True

            Do While blnFileScan = False
                DoEvents
            Loop
        Case 10
            ShowPuzzle
        Case 11
            ShowFileName
        Case 13
            DeleteBmp        'Delete the temporary Bmp file.
            SetWallpaper strOldPaper, lngOrigMode
            End
        End Select           'lngTPM
    End If

    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "PopMenu(" & lngButton & ")"

End Sub

Private Sub RandomPics()

    Do While blnRandom = True
        Begin
    Loop

End Sub

Private Sub SetWallpaper(ByVal strFilename As String, ByVal Wallpaperstyle As Long)

Dim ActiveDesktop1                   As ActiveDesktop
Dim Component1                       As COMPONENTSOPT
Dim Wallpaper1                       As WALLPAPEROPT
'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    This SetWallpaper Sub is from Michael Barnathan.             ****
'   ****    His application was originally found at the following        ****
'   ****    location:                                                    ****
'   ****                                                                 ****
'   ****    http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=36900&lngWId=1
'   ****                                                                 ****
'   ****    Michael's use of the Active Desktop Object allows many       ****
'   ****    image types (besides BMP's) to be used as the WallPaper,     ****
'   ****    and it was very easy to incorporate into this                ****
'   ****    application.                                                 ****
'   ****                                                                 ****
'   ****    Good job Michael, and Thank you!                             ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'
    On Error GoTo ErrorHandler
    Set ActiveDesktop1 = New ActiveDesktop
    Component1.dwSize = Len(Component1)
    ActiveDesktop1.GetDesktopItemOptions Component1, 0&

    If Component1.fActiveDesktop = 0 Then
        Component1.fActiveDesktop = 1       'Enable Active Desktop if it must be enabled
        ActiveDesktop1.SetDesktopItemOptions Component1, 0&
    End If

    Wallpaper1.dwSize = Len(Wallpaper1)
    Wallpaper1.dwStyle = Wallpaperstyle

    ActiveDesktop1.SetWallpaperOptions Wallpaper1, 0&
    ActiveDesktop1.SetWallpaper strFilename, 0&
    ActiveDesktop1.ApplyChanges AD_APPLY_ALL

    Set ActiveDesktop1 = Nothing            'Clean up after we're done
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "SetWallpaper(" & strFilename & ", " & Wallpaperstyle & ")"

End Sub

Private Sub ShowFileName()

    On Error GoTo ErrorHandler

    With lblFileName
        .Caption = "  " & strName & "  "
        .Move 6, 6
    End With  'lblFileName

    picFileName.Move 0, 0, lblFileName.Width + 14, lblFileName.Height + 14
    picFileName.Visible = True
    blnShowFileName = Not blnShowFileName
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "ShowFileName()"

End Sub

Private Sub ShowPuzzle()

    On Error GoTo ErrorHandler
    lngPauseTime = 3000 + (lngBoxes * 50)

'   Show the completed Puzzle for "lngPauseTime" seconds
    With picPuzzle_Main
        .Picture = LoadPicture("")
        .Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
        .Picture = LoadPicture(strBmp)
        .Visible = True
    End With  'picPuzzle_Main

    Pause
    picPuzzle_Main.Visible = False

    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "ShowPuzzle()"

End Sub

Private Sub Solve()

    On Error GoTo ErrorHandler
    Do While blnSolve
        For lngLoop1 = 0 To lngPicTo - 1
            If Not blnSolve Then
                Exit For
            End If

            If picBox(lngLoop1).Index <> picBox(lngLoop1).Tag Then
                For lngLoop2 = 0 To lngPicTo
                    If Not blnSolve Then
                        Exit For
                    End If

                    If picBox(lngLoop1).Tag = picBox(lngLoop2).Index Then
                        If lngPerRow < 16 Then
'                           Hi-Lite the selected box.
                            HiLiteBox (lngLoop1)
'                           Hi-Lite the selected box.
                            HiLiteBox (lngLoop2)

                            DoEvents
                            lngPauseTime = 175 - ((lngPerRow \ 5) * 50)
                            Pause
                            Form_KeyPress (13)
                        Else
                            lngSwap1 = lngLoop1
                            lngSwap2 = lngLoop2
                            Swap
                            DoEvents
'                           Check to see if the Puzzle pieces are in the correct order.
                            CheckIt

                            If lngBoxesOOO = 0 Then
                                blnSolve = False
                                Done
                            End If

                            Exit For
                        End If
                    End If
                Next lngLoop2
            End If
        Next lngLoop1
    Loop

    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "Solve()"

End Sub

Private Sub Swap()
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Swap the 2 selected Picture Boxes.                           ****
'   ****                                                                 ****
'   *************************************************************************
'
    On Error GoTo ErrorHandler

    lngSwap = picBox(lngSwap1).Tag
    picBox(lngSwap1).Tag = picBox(lngSwap2).Tag
    picBox(lngSwap2).Tag = lngSwap

    lngSwap = picBox(lngSwap1).Left
    picBox(lngSwap1).Left = picBox(lngSwap2).Left
    picBox(lngSwap2).Left = lngSwap

    lngSwap = picBox(lngSwap1).Top
    picBox(lngSwap1).Top = picBox(lngSwap2).Top
    picBox(lngSwap2).Top = lngSwap

    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "Swap()"

End Sub
