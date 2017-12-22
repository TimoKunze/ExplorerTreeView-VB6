VERSION 5.00
Object = "{CAB249C7-BAC9-4C51-9526-12F29E40C4CE}#2.6#0"; "ExTvw.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Projekt1"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   374
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   512
   StartUpPosition =   2  'Bildschirmmitte
   Begin ComctlLib.ListView lvListView 
      Height          =   4695
      Left            =   3480
      TabIndex        =   0
      Top             =   375
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   8281
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      OLEDragMode     =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Datei"
         Object.Width           =   26458
      EndProperty
   End
   Begin ExTvw.ExplorerTreeView tvTreeView 
      Height          =   4695
      Left            =   75
      TabIndex        =   1
      Top             =   375
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   8281
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      Indent          =   19
      ItemHeight      =   17
      LimitEditInput  =   0   'False
      LineColor       =   10070188
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'Kein
      FillColor       =   &H00808080&
      Height          =   4800
      Left            =   5400
      ScaleHeight     =   2090.126
      ScaleMode       =   0  'Benutzerdefiniert
      ScaleWidth      =   780
      TabIndex        =   2
      Top             =   945
      Visible         =   0   'False
      Width           =   72
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Drag files out of the list view and drop them whereever you want."
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4305
   End
   Begin VB.Image imgSplitter 
      Height          =   4785
      Left            =   3120
      MousePointer    =   9  'Größenänderung W O
      Top             =   945
      Width           =   150
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Private Const sglSplitLimit = 25

  Private mbMoving As Boolean

  Private Declare Function PathAddBackslash Lib "shlwapi.dll" Alias "PathAddBackslashA" (ByVal Path As String) As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Private Sub Form_Initialize()
  InitCommonControls
End Sub

Private Sub Form_Load()
  tvTreeView.hWndShellUIParentWindow = Me.hWnd
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  If Me.Width < 3000 Then Me.Width = 3000
  SizeControls imgSplitter.Left
End Sub

Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  With imgSplitter
    picSplitter.Move .Left, .Top, .Width \ 2, .Height - 1
  End With
  picSplitter.Visible = True
  mbMoving = True
End Sub

Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim sglPos As Single
  
  If mbMoving Then
    sglPos = ScaleX(x, vbTwips, vbPixels) + imgSplitter.Left
    If sglPos < sglSplitLimit Then
      picSplitter.Left = sglSplitLimit
    ElseIf sglPos > Me.ScaleWidth - sglSplitLimit Then
      picSplitter.Left = Me.ScaleWidth - sglSplitLimit
    Else
      picSplitter.Left = sglPos
    End If
    SizeControls picSplitter.Left
  End If
End Sub

Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  SizeControls picSplitter.Left
  picSplitter.Visible = False
  mbMoving = False
End Sub

Private Sub lvListView_OLECompleteDrag(Effect As Long)
  listFiles
End Sub

Private Sub lvListView_OLEStartDrag(Data As ComctlLib.DataObject, AllowedEffects As Long)
  Dim lvItem As ListItem
  Dim Path As String

  Path = addBackslash(tvTreeView.ItemHandleToFSPath(tvTreeView.SelectedItem))
  With lvListView
    For Each lvItem In .ListItems
      If lvItem.Selected Then
        Data.Files.Add Path & lvItem.Text
      End If
    Next
  End With
  Data.SetData , vbCFFiles
  AllowedEffects = ccOLEDropEffectCopy Or ccOLEDropEffectMove
End Sub

Private Sub tvTreeView_SelChanged(ByVal hOldItem As Long, ByVal hNewItem As Long, ByVal CausedBy As ExTvw.CausedByConstants)
  listFiles
End Sub


Private Function addBackslash(ByVal txt As String) As String
  txt = txt & String$(2, Chr$(0))
  PathAddBackslash txt
  addBackslash = Trim$(removeNull(txt))
End Function

Private Sub listFiles()
  Dim strFile As String

  lvListView.ListItems.Clear
  ' fill listview with all files in this folder
  With tvTreeView
    If .ItemIsFSFolder(.SelectedItem) Then
      strFile = Dir(addBackslash(.ItemHandleToFSPath(.SelectedItem)) & "*.*")
      While strFile <> ""
        lvListView.ListItems.Add , , strFile
        strFile = Dir
      Wend
    End If
  End With
End Sub

Private Function removeNull(ByVal txt As String) As String
  Dim pos As Integer

  pos = InStr(txt, Chr(0))
  If pos Then txt = Left(txt, pos - 1)
  removeNull = txt
End Function

Private Sub SizeControls(x As Single)
  On Error Resume Next

  ' Festlegen der Breite
  If x < 75 Then x = 75
  If x > (Me.ScaleWidth - 75) Then x = Me.ScaleWidth - 75
  tvTreeView.Width = x - tvTreeView.Left
  imgSplitter.Left = x
  lvListView.Left = x + 4
  lvListView.Width = Me.ScaleWidth - (lvListView.Left + 5)
  lvListView.Top = tvTreeView.Top
  tvTreeView.Height = Me.ScaleHeight - (tvTreeView.Top + 5)
  lvListView.Height = tvTreeView.Height
  imgSplitter.Top = tvTreeView.Top
  imgSplitter.Height = tvTreeView.Height
End Sub
