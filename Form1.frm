VERSION 5.00
Begin VB.Form FormRuler 
   BorderStyle     =   0  'None
   Caption         =   "Ruler"
   ClientHeight    =   3570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8490
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   8490
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   600
      Top             =   2280
   End
   Begin VB.PictureBox pRuler 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   480
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   417
      TabIndex        =   0
      Top             =   600
      Width           =   6255
      Begin VB.PictureBox pTick 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   3480
         ScaleHeight     =   615
         ScaleWidth      =   735
         TabIndex        =   5
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton cmdOnTop 
         Caption         =   "On Top"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   2160
         TabIndex        =   4
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton cmdSize 
         Caption         =   "Size"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   1440
         TabIndex        =   3
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton cmdHV 
         Caption         =   "Vertical"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   720
         TabIndex        =   2
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Width           =   615
      End
   End
End
Attribute VB_Name = "FormRuler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
' MODULE:       FormRuler
' FILENAME:     j:/allpro/develop/ruler/Form1.frm
' AUTHOR:       John Wunderlin
' EMAIL:        johnw@allprosoftware.com
' WEB:          www.allprosoftware.com
' CREATED:      02-Dec-2004
' COPYRIGHT:    Copyright 2004 All-Pro Software. All Rights Reserved.
'
' DESCRIPTION:
'
' This is a simple pixel ruler that lets you measure forms for developers.
' The ruler will stay on top by default, but can be set to move up/down the zorder
' It starts at the upper-left of the screen and is set to the width of the display.
' You can flip the orientation between horizontal and vertical and change the size
' of the ruler.
'
' I wrote this program because I couldn't find a shareware/freeware pixel ruler
' on the internet other than one that had a huge (20 meg+!!) memory footprint
' and lots of runtime files.
'
' I'm releasing this as freeware/open source as a thank you to all the other
' developers who have done the same.
'
' If you make any changes to the program, I'd like to get a copy so I can inlcude
' it in my code.  Try to make the code clear, well documented and useful! :)
'
' FUTURE PLANS:
'
' - Replace buttons with right-mouse click so I can shrink the ruler, esp. the vertical
' - Add an 'x' button to close the ruler (in addition to right-mouse menu)
' - Remember last position and orientation
' - Save/load current settings in an .ini file
'
' MODIFICATION HISTORY:
' 1.0       03-Dec-2004
'           John Wunderlin
'           Initial Version
'*******************************************************************************
Option Explicit

Const BIGTICK = 15      'Size of the larger tick marks
Dim HO As Boolean       'Is this ruler in Horizontal Orientation
Const MEDTICK = 7       'Size of the medium tick marks
Dim OnTop As Boolean    'Is the ruler set to On Top
Dim RulerSize As Long   'Size of the ruler in pixels
Const SMALLTICK = 5     'Size of the smaller tick marks

Private Sub cmdClose_Click()
   'End the program
   End
End Sub

'*******************************************************************************
' cmdHV_Click (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' Set the ruler's orientation
'*******************************************************************************
Private Sub cmdHV_Click()
   If HO Then
      HO = False
      Me.pTick.Width = pRuler.Width
      Me.pTick.Height = 1
      Me.cmdHV.Caption = "Horiz."
   Else
      HO = True
      Me.pTick.Width = 1
      Me.pTick.Height = pRuler.Height
      Me.cmdHV.Caption = "Vertical"
   End If
   DrawRuler
   Me.pRuler.SetFocus
End Sub

'*******************************************************************************
' cmdOnTop_Click (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' Set the OnTop setting
'*******************************************************************************
Public Sub cmdOnTop_Click()
   If OnTop Then
      Call AlwaysOnTop(Me, False)
      Me.cmdOnTop.Caption = "On Top"
      OnTop = False
   Else
      Call AlwaysOnTop(Me, True)
      Me.cmdOnTop.Caption = "No Top"
      OnTop = True
   End If
   'needed this trap because the proc is called from formload and the pruler isn't
   ' displayed yet.  NOTE I set the focus to the ruler rather than the buttons because
   ' with the small font I'm using, the hash marks around the text distort it so
   ' much that it becomes unreadable!
   On Error Resume Next
   Me.pRuler.SetFocus
End Sub

'*******************************************************************************
' cmdSize_Click (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' Set the size of the ruler in pixels.  Note the ruler will be the same size
' vertical and horizontal.
'*******************************************************************************
Private Sub cmdSize_Click()
   Dim UserSize As Long    'Size the user specified. 0 if cancel
   
   UserSize = Val(InputBox("Enter the size of the ruler in pixels", "Set Pixel Size", Screen.Width / Screen.TwipsPerPixelX))
   If UserSize = 0 Then
      Me.pRuler.SetFocus
      Exit Sub
   End If
   
   RulerSize = UserSize
   'Make sure it's at least 400 pixels so the buttons can be displayed
   If RulerSize < 400 Then
      RulerSize = 400
   End If
   
   Call DrawRuler
   Me.pRuler.SetFocus
End Sub

'*******************************************************************************
' DrawRuler (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' Render the ruler in either horizontal or vertical
'*******************************************************************************
Sub DrawRuler()
   Dim DisplayNumber As String   'Holds the numbers shown on the ruler
   Dim LineHeight As Long        'Size of the current tick mark
   Dim RulerHeight As Long       'Thickness of the ruler
   Dim X As Long                 'current x position
   Dim Y As Long                 'current y position
   
   pRuler.Cls
   
   'Set the ruler dimensions
   If HO Then
      Me.Width = RulerSize * Screen.TwipsPerPixelX
      Me.Height = (Me.cmdClose.Height + (SMALLTICK * 2)) * Screen.TwipsPerPixelY
   Else
      Me.Height = RulerSize * Screen.TwipsPerPixelY
      Me.Width = (Me.cmdClose.Width + (SMALLTICK * 2)) * Screen.TwipsPerPixelX
   End If
   
   Me.pRuler.Width = Me.ScaleWidth
   Me.pRuler.Height = Me.ScaleHeight

   If HO Then
      RulerHeight = pRuler.ScaleHeight
      
      'draw numbers and ticks across
      For X = 0 To RulerSize Step 10
         If X Mod 100 = 0 Then

            'Show the number and bigger tick
            DisplayNumber = Trim(Str(X))
            LineHeight = BIGTICK
            pRuler.CurrentX = X
            
            'Center the text in the middle of the ruler at the current postion
            pRuler.CurrentY = SMALLTICK + (Me.cmdClose.Height / 2) - (pRuler.TextHeight(DisplayNumber) / 2)
            If X = 0 Then
               pRuler.CurrentX = 0
            Else
               pRuler.CurrentX = pRuler.CurrentX - (pRuler.TextWidth(DisplayNumber) / 2)
            End If
            pRuler.Print DisplayNumber
         ElseIf X Mod 50 = 0 Then
            LineHeight = MEDTICK
         Else
            LineHeight = SMALLTICK
         End If
         pRuler.Line (X, 0)-(X, LineHeight)
         pRuler.Line (X, RulerHeight)-(X, (RulerHeight - LineHeight))
         
      Next
      
      'position buttons between the 100's values
      Me.cmdClose.Left = 50 - cmdClose.Width / 2
      Me.cmdHV.Left = 150 - cmdHV.Width / 2
      Me.cmdSize.Left = 250 - cmdSize.Width / 2
      Me.cmdOnTop.Left = 350 - cmdOnTop.Width / 2
      
      Me.cmdClose.Top = SMALLTICK
      Me.cmdHV.Top = SMALLTICK
      Me.cmdSize.Top = SMALLTICK
      Me.cmdOnTop.Top = SMALLTICK
      
   Else
      RulerHeight = pRuler.ScaleWidth
      'draw numbers down
      For Y = 0 To RulerSize Step 10
         If Y Mod 100 = 0 Then
         
            'show the numbers and the bigger ticks
            DisplayNumber = Trim(Str(Y))
            LineHeight = BIGTICK
            pRuler.CurrentY = Y
            
            'Center the text in the middle of the ruler at the current postion
            pRuler.CurrentX = SMALLTICK + (Me.cmdClose.Width / 2) - (pRuler.TextWidth(DisplayNumber) / 2)
            If Y = 0 Then
               pRuler.CurrentY = 0
            Else
               pRuler.CurrentY = pRuler.CurrentY - (pRuler.TextHeight(DisplayNumber) / 2)
            End If
            pRuler.Print DisplayNumber
         ElseIf Y Mod 50 = 0 Then
            LineHeight = MEDTICK
         Else
            LineHeight = SMALLTICK
         End If
         pRuler.Line (0, Y)-(LineHeight, Y)
         pRuler.Line (RulerHeight, Y)-((RulerHeight - LineHeight), Y)
         
      Next
      
      'position buttons between the 100's values
      Me.cmdClose.Top = 50 - cmdClose.Height / 2
      Me.cmdHV.Top = 150 - cmdHV.Height / 2
      Me.cmdSize.Top = 250 - cmdSize.Height / 2
      Me.cmdOnTop.Top = 350 - cmdOnTop.Height / 2
      
      Me.cmdClose.Left = SMALLTICK
      Me.cmdHV.Left = SMALLTICK
      Me.cmdSize.Left = SMALLTICK
      Me.cmdOnTop.Left = SMALLTICK
      
   End If
End Sub

'*******************************************************************************
' Form_Load (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' Position ruler at 0,0 in the horizontal position and set ontop
'*******************************************************************************
Private Sub Form_Load()
   RulerSize = (Screen.Width / Screen.TwipsPerPixelX)
   HO = True
   
   'set ontop true
   OnTop = False
   Call cmdOnTop_Click
   
   Me.Width = RulerSize * Screen.TwipsPerPixelX
   Me.Height = (10 + Me.cmdClose.Height) * Screen.TwipsPerPixelY
   Me.Move 0, 0
   pTick.Width = 1
   pTick.Height = pRuler.Height
   
   Me.pTick.ZOrder (1)

   Call DrawRuler
   Call pRuler.Move(0, 0)
End Sub

'*******************************************************************************
' pRuler_MouseDown (SUB)
'
' PARAMETERS:
' (In/Out) - Button - Integer -
' (In/Out) - Shift  - Integer -
' (In/Out) - X      - Single  -
' (In/Out) - Y      - Single  -
'
' DESCRIPTION:
' This function allows dragging of the ruler without a title bar
'*******************************************************************************
Private Sub pRuler_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

'*******************************************************************************
' ptick_MouseDown (SUB)
'
' PARAMETERS:
' (In/Out) - Button - Integer -
' (In/Out) - Shift  - Integer -
' (In/Out) - X      - Single  -
' (In/Out) - Y      - Single  -
'
' DESCRIPTION:
' This function allows dragging of the ruler without a title bar.  This was
' needed in addition to the pRuler Mousedown because the user may click
' directly on the tick mark, or they may click on the ruler
'*******************************************************************************
Private Sub ptick_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

'*******************************************************************************
' Timer1_Timer (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' Refresh the position of the tick mark indicator every 10th of a second.
' Note that this takes very little cpu to perform this task.
'*******************************************************************************
Private Sub Timer1_Timer()
   Call GetCursorPos(MousePos)
   If HO Then
      Me.pTick.Left = MousePos.X - (Me.Left / Screen.TwipsPerPixelX)
      Me.pTick.Top = 0
   Else
      Me.pTick.Left = 0
      Me.pTick.Top = MousePos.Y - (Me.Top / Screen.TwipsPerPixelY)
   End If
End Sub
