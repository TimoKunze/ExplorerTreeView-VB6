VERSION 5.00
Begin VB.Form frmBrowseForFolder 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Nach einem Ordner durchsuchen"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBrowseForFolder.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   290
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin ExTvw.ExplorerTreeView ExTvw 
      Height          =   2160
      Left            =   285
      TabIndex        =   0
      Top             =   900
      Width           =   4230
      _ExtentX        =   7461
      _ExtentY        =   3810
      DragExpandTime  =   1000
      DragScrollTime  =   200
      FadeExpandos    =   -1  'True
      FileAttributes  =   63
      FolderAttributes=   63
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Indent          =   19
      ItemHeight      =   17
      ReplaceHandCursor=   -1  'True
      ShowFocusRect   =   0   'False
      ShownOverlays   =   2
      SingleExpand    =   2
      TreeViewStyle   =   1
   End
   Begin VB.CommandButton cmdNewFolder 
      Caption         =   "Neuen &Ordner erstellen"
      Height          =   345
      Left            =   150
      TabIndex        =   5
      Top             =   3825
      Width           =   1890
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Abbrechen"
      Height          =   345
      Left            =   3390
      TabIndex        =   4
      Top             =   3825
      Width           =   1140
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   345
      Left            =   2145
      TabIndex        =   3
      Top             =   3825
      Width           =   1140
   End
   Begin VB.TextBox txtFolder 
      Height          =   285
      Left            =   1035
      TabIndex        =   2
      Top             =   3285
      Width           =   3480
   End
   Begin VB.Label lblDescr 
      AutoSize        =   -1  'True
      Caption         =   "Ordner:"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   3375
      Width           =   570
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Height          =   540
      Left            =   240
      TabIndex        =   6
      Top             =   180
      Width           =   4320
   End
End
Attribute VB_Name = "frmBrowseForFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  ' Property-Variablen

  Private propSelected As String


  Private Declare Function SetWindowTheme Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pSubAppName As Long, ByVal pSubIDList As Long) As Long


' Events

Private Sub cmdCancel_Click()
  propSelected = ""
  Hide
End Sub

Private Sub cmdNewFolder_Click()
  ExTvw.ItemCreateNewFolder ExTvw.SelectedItem
End Sub

Private Sub cmdOk_Click()
  Hide
End Sub

Private Sub ExTvw_ItemLoadedSubItems(ByVal hItem As Long)
  ExTvw.MousePointer = MousePointerConstants.vbDefault
  DoEvents
End Sub

Private Sub ExTvw_ItemLoadingSubItems(ByVal hItem As Long)
  ExTvw.MousePointer = MousePointerConstants.vbHourglass
  DoEvents
End Sub

Private Sub ExTVW_SelChanged(ByVal hOldItem As Long, ByVal hNewItem As Long, ByVal CausedBy As CausedByConstants)
  propSelected = ExTvw.ItemHandleToTreePath(hNewItem)
  txtFolder = ExTvw.ItemHandleToDisplayName(hNewItem)
  cmdNewFolder.Enabled = ExTvw.ItemSupportsNewFolder(hNewItem)
End Sub

Private Sub Form_Load()
  Dim hMod As Long

  hMod = LoadLibrary("uxtheme.dll")
  If hMod Then
    SetWindowTheme ExTvw.hWnd, StrPtr("explorer"), 0
    FreeLibrary hMod
  End If

  ExTvw.hWndShellUIParentWindow = Me.hWnd
  ExTvw.ReplaceHandCursor = Not ver_Win_Vista

  cmdNewFolder.Enabled = ExTvw.ItemSupportsNewFolder(ExTvw.SelectedItem)
End Sub

Private Sub txtFolder_GotFocus()
  selectText txtFolder
End Sub


' öffentliche Props

Public Property Get Selected() As String
  Selected = propSelected
End Property


' öffentliche Methoden

Public Sub showIt(ParamArray Args())
  With ExTvw
    .IncludedItems = Args(3)
    .DrivesOnly = Args(4)
    .FileAttributes = Args(5)
    .FolderAttributes = Args(6)
    .FileFilters = Args(7)
    .UseFileFilters = (Args(7) <> "")
    .FolderFilters = Args(8)
    .UseFolderFilters = (Args(8) <> "")
    .RootPath = Args(1)
    .Path = Args(2)
    If Args(0) Then
      lblCaption = "Wählen Sie den Item, der gerade markiert sein soll."
      .ExpandArchives = Args(9)
    Else
      lblCaption = "Wählen Sie den Item, der als Startverzeichnis dienen soll."
      .ExpandArchives = ExpandArchivesConstants.eaNone
      .IncludedItems = .IncludedItems And Not IncludedItemsConstants.iiFSFiles And Not IncludedItemsConstants.iiNonFSFiles
    End If
  End With

  Show FormShowConstants.vbModal
End Sub
