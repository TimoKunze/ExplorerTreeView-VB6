VERSION 5.00
Object = "{CAB249C7-BAC9-4C51-9526-12F29E40C4CE}#2.0#0"; "ExTvw.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Favoriten verwalten"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   301
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   511
   StartUpPosition =   2  'Bildschirmmitte
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '2D
      ForeColor       =   &H80000008&
      Height          =   1830
      Left            =   270
      ScaleHeight     =   1800
      ScaleWidth      =   3555
      TabIndex        =   7
      Top             =   1950
      Width           =   3585
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Löschen"
      Height          =   330
      Left            =   2145
      TabIndex        =   6
      Top             =   1530
      Width           =   1710
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Verschieben..."
      Height          =   330
      Left            =   270
      TabIndex        =   5
      Top             =   1530
      Width           =   1710
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Umbenennen"
      Height          =   330
      Left            =   2145
      TabIndex        =   4
      Top             =   1095
      Width           =   1710
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Or&dner erstellen"
      Height          =   330
      Left            =   270
      TabIndex        =   3
      Top             =   1095
      Width           =   1710
   End
   Begin VB.CommandButton Command1 
      Caption         =   "S&chließen"
      Height          =   330
      Left            =   5715
      TabIndex        =   1
      Top             =   3825
      Width           =   1710
   End
   Begin ExTvw.ExplorerTreeView ExplorerTreeView1 
      Height          =   3495
      Left            =   4140
      TabIndex        =   0
      Top             =   195
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   6165
      DragExpandTime  =   1000
      DragScrollTime  =   200
      FavoritesStyle  =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FullRowSelect   =   -1  'True
      HotForeColor    =   -2147483635
      HotTracking     =   -1  'True
      IncludedItems   =   7
      Indent          =   19
      ItemHeight      =   17
      Path            =   "Favoriten"
      ReplaceHandCursor=   -1  'True
      RootPath        =   "Favoriten"
      ShowFocusRect   =   0   'False
      ShownOverlays   =   14
      ShownToolTips   =   3
      ShowRoot        =   0   'False
      SingleExpand    =   1
      TreeViewStyle   =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMain.frx":0000
      Height          =   855
      Left            =   270
      TabIndex        =   2
      Top             =   180
      Width           =   3615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()


Private Sub Command1_Click()
  Unload Me
End Sub

Private Sub Command2_Click()
  MsgBox "ToDo"
End Sub

Private Sub Command3_Click()
  MsgBox "ToDo"
End Sub

Private Sub Command4_Click()
  MsgBox "ToDo"
End Sub

Private Sub Command5_Click()
  MsgBox "ToDo"
End Sub

Private Sub ExplorerTreeView1_ItemBeforeGetInfoTip(ByVal hItem As Long, InfoTipStyle As extvw.InfoTipStyleConstants, Cancel As Boolean)
  InfoTipStyle = itsDefault
End Sub

Private Sub Form_Initialize()
  InitCommonControls
End Sub

Private Sub Form_Load()
  ExplorerTreeView1.hWndShellUIParentWindow = Me.hWnd
End Sub
