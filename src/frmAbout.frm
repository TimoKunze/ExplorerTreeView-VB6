VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Info zu |1"
   ClientHeight    =   3075
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5985
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   205
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   200
      Width           =   735
   End
   Begin VB.Label lblBugtracker 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "https://bugtracker.timosoft-software.de"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1830
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   8
      ToolTipText     =   "Einen Bug melden / Report a bug"
      Top             =   1800
      Width           =   2685
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bugtracker:"
      Height          =   195
      Index           =   3
      Left            =   960
      TabIndex        =   7
      Top             =   1800
      Width           =   825
   End
   Begin VB.Label lblForum 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "https://www.timosoft-software.de/forum/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1470
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   6
      ToolTipText     =   "Das Supportforum / The support forums"
      Top             =   1560
      Width           =   2820
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Forum:"
      Height          =   195
      Index           =   2
      Left            =   960
      TabIndex        =   5
      Top             =   1560
      Width           =   480
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Getestet von: Christian Lütgens"
      Height          =   195
      Index           =   5
      Left            =   960
      TabIndex        =   11
      Top             =   2400
      Width           =   2220
   End
   Begin VB.Label lblDC 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "https://www.dateicommander.de"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3360
      TabIndex        =   12
      ToolTipText     =   "https://www.dateicommander.de"
      Top             =   2400
      Width           =   2265
   End
   Begin VB.Image imgPaypal 
      Height          =   465
      Left            =   4680
      Picture         =   "frmAbout.frx":000C
      ToolTipText     =   "Mit einer Spende ""Danke"" sagen / Say ""Thanks"" with a donation"
      Top             =   1320
      Width           =   1200
   End
   Begin VB.Label lblNikos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.zabkat.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3360
      TabIndex        =   14
      ToolTipText     =   "http://www.zabkat.com"
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Many thanks to Nikos Bozinis!"
      Height          =   195
      Index           =   6
      Left            =   960
      TabIndex        =   13
      Top             =   2760
      Width           =   2130
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "e-Mail:"
      Height          =   195
      Index           =   4
      Left            =   960
      TabIndex        =   9
      Top             =   2040
      Width           =   465
   End
   Begin VB.Label lblMail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Timo@TimoSoft-Software.de"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1470
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   10
      ToolTipText     =   "Eine e-Mail schreiben / Write an e-mail"
      Top             =   2040
      Width           =   2040
   End
   Begin VB.Label lblURL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "https://www.TimoSoft-Software.de"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   960
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   4
      ToolTipText     =   "Webseite öffnen / Open website"
      Top             =   1320
      Width           =   2400
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "|1"
      Height          =   195
      Left            =   960
      TabIndex        =   2
      Top             =   600
      Width           =   120
   End
   Begin VB.Image imgLogo 
      Height          =   480
      Left            =   240
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (C) 2002-2018, Timo ""TimoSoft"" Kunze"
      Height          =   195
      Index           =   1
      Left            =   960
      TabIndex        =   3
      Top             =   1080
      Width           =   3465
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "|1 - ActiveX-Control"
      Height          =   195
      Index           =   0
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   1350
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Implements ISubClass


  ' lokale Variablen

  Private CtlType As String


  ' lokale APIs

  Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal Operation As String, ByVal File As String, ByVal CmdLine As String, ByVal path As String, ByVal Flags As Long) As Long


' Implementation von ISubClass

Private Function ISubClass_WinMain(ByVal iListener As Long, ByVal oldWinMain As Long, ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  ISubClass_WinMain = WinMain(iListener, oldWinMain, hWnd, MSG, wParam, lParam)
End Function


' Events

Private Sub cmdOk_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  SubClass

  With App
    lblVersion = Replace$(lblVersion, "|1", "Version " & .Major & "." & .Minor & "." & .Revision & " (Build " & BuildNumber & ")") ' - RC3")
  End With
End Sub

Private Sub Form_Terminate()
  UnSubClass
End Sub

Private Sub imgPaypal_Click()
  Const SW_SHOWNORMAL = 1

  ShellExecute hWnd, "open", "https://www.paypal.com/xclick/business=TKunze71216%40gmx.de&item_name=ExplorerTreeView&no_shipping=1&tax=0&currency_code=EUR", "", "", SW_SHOWNORMAL
End Sub

Private Sub lblBugtracker_Click()
  Const SW_SHOWNORMAL = 1

  ShellExecute hWnd, "open", "https://bugtracker.timosoft-software.de/", "", "", SW_SHOWNORMAL
End Sub

Private Sub lblDC_Click()
  Const SW_SHOWNORMAL = 1

  ShellExecute hWnd, "open", lblDC.ToolTipText, "", "", SW_SHOWNORMAL
End Sub

Private Sub lblForum_Click()
  Const SW_SHOWNORMAL = 1

  ShellExecute hWnd, "open", "https://www.timosoft-software.de/forum/viewforum.php?f=37", "", "", SW_SHOWNORMAL
End Sub

Private Sub lblMail_Click()
  Const SW_SHOWNORMAL = 1

  ShellExecute hWnd, "open", "mailto:" & lblMail & "?subject=" & CtlType, "", "", SW_SHOWNORMAL
End Sub

Private Sub lblNikos_Click()
  Const SW_SHOWNORMAL = 1

  ShellExecute hWnd, "open", lblNikos.ToolTipText, "", "", SW_SHOWNORMAL
End Sub

Private Sub lblURL_Click()
  Const SW_SHOWNORMAL = 1

  ShellExecute hWnd, "open", lblURL, "", "", SW_SHOWNORMAL
End Sub


' öffentliche Methoden

Public Sub showIt(ByVal ControlType As String, ByVal ResID As Long)
  CtlType = ControlType
  Caption = Replace(Caption, "|1", CtlType)
  lblLabels(0) = Replace(lblLabels(0), "|1", CtlType)
  Set imgLogo = LoadResPicture(ResID, LoadResConstants.vbResIcon)

  Show FormShowConstants.vbModal
End Sub


' lokale Methoden

' startet das Abfangen von Nachrichten
Private Sub SubClass()
  StartListening Me, hWnd
End Sub

' beendet das Abfangen von Nachrichten für alle Fenster, für die SubClassing aktiviert wurde
Private Sub UnSubClass()
  StopListening Me, hWnd
End Sub

' wird bei eingehenden Nachrichten aufgerufen
Private Function WinMain(ByVal iListener As Long, ByVal oldWinMain As Long, ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Const WM_SETCURSOR = &H20
  Dim hCursor As Long
  Dim pt As POINT
  Dim rc As RECT

  Select Case MSG
    Case WM_SETCURSOR
      ' Möglichkeit, den Mauszeiger anzupassen
      With lblURL
        rc.Bottom = .Top + .Height
        rc.Left = .Left
        rc.Right = .Left + .Width
        rc.Top = .Top
      End With
      GetCursorPos pt
      ScreenToClient hWnd, pt

      If IsWithinRectangle(rc, pt) Then
        hCursor = LoadImage(0, MAKEINTRESOURCE(OCR_HAND), IMAGE_CURSOR, 0, 0, LR_DEFAULTCOLOR Or LR_DEFAULTSIZE Or LR_SHARED)
      Else
        With lblMail
          rc.Bottom = .Top + .Height
          rc.Left = .Left
          rc.Right = .Left + .Width
          rc.Top = .Top
        End With
        If IsWithinRectangle(rc, pt) Then
          hCursor = LoadImage(0, MAKEINTRESOURCE(OCR_HAND), IMAGE_CURSOR, 0, 0, LR_DEFAULTCOLOR Or LR_DEFAULTSIZE Or LR_SHARED)
        Else
          With lblForum
            rc.Bottom = .Top + .Height
            rc.Left = .Left
            rc.Right = .Left + .Width
            rc.Top = .Top
          End With
          If IsWithinRectangle(rc, pt) Then
            hCursor = LoadImage(0, MAKEINTRESOURCE(OCR_HAND), IMAGE_CURSOR, 0, 0, LR_DEFAULTCOLOR Or LR_DEFAULTSIZE Or LR_SHARED)
          Else
            With lblBugtracker
              rc.Bottom = .Top + .Height
              rc.Left = .Left
              rc.Right = .Left + .Width
              rc.Top = .Top
            End With
            If IsWithinRectangle(rc, pt) Then
              hCursor = LoadImage(0, MAKEINTRESOURCE(OCR_HAND), IMAGE_CURSOR, 0, 0, LR_DEFAULTCOLOR Or LR_DEFAULTSIZE Or LR_SHARED)
            Else
              With imgPaypal
                rc.Bottom = .Top + .Height
                rc.Left = .Left
                rc.Right = .Left + .Width
                rc.Top = .Top
              End With
              If IsWithinRectangle(rc, pt) Then
                hCursor = LoadImage(0, MAKEINTRESOURCE(OCR_HAND), IMAGE_CURSOR, 0, 0, LR_DEFAULTCOLOR Or LR_DEFAULTSIZE Or LR_SHARED)
              Else
                With lblDC
                  rc.Bottom = .Top + .Height
                  rc.Left = .Left
                  rc.Right = .Left + .Width
                  rc.Top = .Top
                End With
                If IsWithinRectangle(rc, pt) Then
                  hCursor = LoadImage(0, MAKEINTRESOURCE(OCR_HAND), IMAGE_CURSOR, 0, 0, LR_DEFAULTCOLOR Or LR_DEFAULTSIZE Or LR_SHARED)
                Else
                  With lblNikos
                    rc.Bottom = .Top + .Height
                    rc.Left = .Left
                    rc.Right = .Left + .Width
                    rc.Top = .Top
                  End With
                  If IsWithinRectangle(rc, pt) Then hCursor = LoadImage(0, MAKEINTRESOURCE(OCR_HAND), IMAGE_CURSOR, 0, 0, LR_DEFAULTCOLOR Or LR_DEFAULTSIZE Or LR_SHARED)
                End If
              End If
            End If
          End If
        End If
      End If

      If hCursor Then
        SetCursor hCursor

        WinMain = 1
        Exit Function
      End If
  End Select

DoDefault:
  If iListener = 1 Then
    If oldWinMain Then WinMain = CallWindowProc(oldWinMain, hWnd, MSG, wParam, lParam)
  End If
End Function
