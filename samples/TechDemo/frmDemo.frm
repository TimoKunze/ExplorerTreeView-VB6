VERSION 5.00
Object = "{CAB249C7-BAC9-4C51-9526-12F29E40C4CE}#2.0#0"; "ExTvw.ocx"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Begin VB.Form frmDemo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "ExplorerTreeView TechDemo"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8625
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   360
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   575
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CheckBox chkUseFilters 
      Caption         =   "Filter:"
      Height          =   195
      Left            =   6360
      TabIndex        =   31
      Top             =   4950
      Width           =   705
   End
   Begin vbalIml6.vbalImageList ImgLst 
      Left            =   120
      Top             =   4800
      _ExtentX        =   953
      _ExtentY        =   953
      Size            =   4592
      Images          =   "frmDemo.frx":0000
      Version         =   131072
      KeyCount        =   4
      Keys            =   "ÿÿÿ"
   End
   Begin VB.CheckBox chkCustomCheckBoxes 
      Caption         =   "custom"
      Height          =   195
      Left            =   3480
      TabIndex        =   45
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CheckBox chkCustomInfoTips 
      Caption         =   "custom"
      Height          =   255
      Left            =   5040
      TabIndex        =   44
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CheckBox chkToolTips 
      Caption         =   "ToolTips"
      Height          =   255
      Left            =   6600
      TabIndex        =   22
      ToolTipText     =   "ToolTips für abgeschnittene Items"
      Top             =   120
      Width           =   1455
   End
   Begin VB.CheckBox chkInfoTips 
      Caption         =   "InfoTips"
      Height          =   255
      Left            =   4920
      TabIndex        =   21
      Top             =   1560
      Width           =   1455
   End
   Begin VB.OptionButton optSingleExpand 
      Caption         =   "None"
      Height          =   195
      Index           =   0
      Left            =   5040
      TabIndex        =   43
      Top             =   2430
      Width           =   975
   End
   Begin VB.OptionButton optSingleExpand 
      Caption         =   "old Style"
      Height          =   195
      Index           =   1
      Left            =   5040
      TabIndex        =   42
      Top             =   2760
      Width           =   975
   End
   Begin VB.OptionButton optSingleExpand 
      Caption         =   "WinXP-Style"
      Height          =   195
      Index           =   2
      Left            =   6480
      TabIndex        =   41
      Top             =   2430
      Width           =   1335
   End
   Begin VB.ListBox lstExpandArchives 
      Height          =   960
      ItemData        =   "frmDemo.frx":1210
      Left            =   4920
      List            =   "frmDemo.frx":1229
      Style           =   1  'Kontrollkästchen
      TabIndex        =   28
      Top             =   3840
      Width           =   1290
   End
   Begin VB.CheckBox chkDragDrop 
      Caption         =   "Drag'n'Drop"
      Height          =   255
      Left            =   6600
      TabIndex        =   26
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "SingleExpand"
      Height          =   855
      Left            =   4920
      TabIndex        =   39
      Top             =   2190
      Width           =   3255
   End
   Begin VB.CheckBox chkOverlays 
      Caption         =   "Overlays"
      Height          =   255
      Left            =   3480
      TabIndex        =   15
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CheckBox chkUseSelected 
      Caption         =   "use selected"
      Height          =   255
      Left            =   3480
      TabIndex        =   12
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CheckBox chkDimHidden 
      Caption         =   "dim hidden"
      Height          =   255
      Left            =   3480
      TabIndex        =   14
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CheckBox chkDimSelected 
      Caption         =   "dim selected"
      Height          =   255
      Left            =   3480
      TabIndex        =   13
      Top             =   1920
      Width           =   1335
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      ItemData        =   "frmDemo.frx":1250
      Left            =   6360
      List            =   "frmDemo.frx":125A
      Style           =   2  'Dropdown-Liste
      TabIndex        =   29
      Top             =   3150
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   6240
      TabIndex        =   36
      Top             =   3240
      Width           =   2295
      Begin VB.TextBox txtFilters 
         Height          =   285
         Left            =   840
         TabIndex        =   32
         ToolTipText     =   "Format: ""<Filter1>|<Filter2>"", Wildcards sind möglich, ein vorangestelltes ""/"" negiert den Filter"
         Top             =   1680
         Width           =   1335
      End
      Begin VB.ListBox lstAttributes 
         Height          =   1410
         ItemData        =   "frmDemo.frx":126C
         Left            =   840
         List            =   "frmDemo.frx":1282
         Style           =   1  'Kontrollkästchen
         TabIndex        =   30
         Top             =   270
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Attribute:"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   270
         Width           =   705
      End
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   4
      Top             =   3960
      Width           =   375
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   600
      TabIndex        =   3
      Top             =   3960
      Width           =   2055
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   2
      Top             =   3600
      Width           =   375
   End
   Begin VB.TextBox txtRoot 
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CheckBox chkShowRoot 
      Caption         =   "show Root"
      Height          =   255
      Left            =   4920
      TabIndex        =   20
      Top             =   1200
      Width           =   1455
   End
   Begin VB.ListBox lstIncludedItems 
      Height          =   960
      ItemData        =   "frmDemo.frx":12C0
      Left            =   3360
      List            =   "frmDemo.frx":12D0
      Style           =   1  'Kontrollkästchen
      TabIndex        =   27
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CheckBox chkCustomMenu 
      Caption         =   "custom Menu"
      Height          =   255
      Left            =   6720
      TabIndex        =   24
      Top             =   840
      Width           =   1335
   End
   Begin VB.CheckBox chkContextMenus 
      Caption         =   "ContextMenus"
      Height          =   255
      Left            =   6600
      TabIndex        =   23
      Top             =   480
      Width           =   1455
   End
   Begin VB.CheckBox chkRenaming 
      Caption         =   "allow Renaming"
      Height          =   255
      Left            =   6600
      TabIndex        =   25
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CheckBox chkHotTracking 
      Caption         =   "HotTracking"
      Height          =   255
      Left            =   4920
      TabIndex        =   18
      Top             =   480
      Width           =   1455
   End
   Begin VB.CheckBox chkFullRowSelect 
      Caption         =   "FullRowSelect"
      Height          =   255
      Left            =   4920
      TabIndex        =   17
      Top             =   120
      Width           =   1455
   End
   Begin VB.CheckBox chkCheckBoxes 
      Caption         =   "CheckBoxes"
      Height          =   255
      Left            =   3360
      TabIndex        =   16
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CheckBox chkShowSelection 
      Caption         =   "show Selection"
      Height          =   255
      Left            =   4920
      TabIndex        =   19
      Top             =   840
      Width           =   1455
   End
   Begin VB.CheckBox chkIcons 
      Caption         =   "Icons"
      Height          =   255
      Left            =   3360
      TabIndex        =   11
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CheckBox chkHasLinesAtRoot 
      Caption         =   "Lines at Root"
      Height          =   255
      Left            =   3480
      TabIndex        =   10
      Top             =   840
      Width           =   1335
   End
   Begin VB.CheckBox chkTVWStyle 
      Caption         =   "Lines"
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   9
      Top             =   480
      Width           =   1455
   End
   Begin VB.CheckBox chkTVWStyle 
      Caption         =   "Buttons"
      Height          =   255
      Index           =   0
      Left            =   3360
      TabIndex        =   8
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About..."
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton cmdReload 
      Caption         =   "&Reload"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   4320
      Width           =   1455
   End
   Begin ExTvw.ExplorerTreeView ExTvw 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   5953
      DragExpandTime  =   1000
      DragScrollTime  =   200
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
      Indent          =   19
      ItemHeight      =   17
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Archive:"
      Height          =   195
      Left            =   4920
      TabIndex        =   40
      Top             =   3600
      Width           =   600
   End
   Begin VB.Label Label5 
      Caption         =   "Diese Demo zeigt nur einen Teil der Features!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   555
      Left            =   3360
      TabIndex        =   38
      Top             =   4800
      Width           =   2775
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Path:"
      Height          =   195
      Left            =   120
      TabIndex        =   35
      Top             =   4050
      Width           =   390
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Root:"
      Height          =   195
      Left            =   120
      TabIndex        =   34
      Top             =   3690
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "IncludedItems:"
      Height          =   195
      Left            =   3360
      TabIndex        =   33
      Top             =   3600
      Width           =   1080
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long

'    hBitmap As Long     ' ab Windows 98
  End Type

  Private settingProps As Boolean

  Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
  Private Declare Function GetTickCount Lib "kernel32" () As Long
  Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal Prop As Long) As Long
  Private Declare Function ImageList_Duplicate Lib "comctl32" (ByVal hImageList As Long) As Long
  Private Declare Sub InitCommonControls Lib "comctl32" ()
  Private Declare Function InsertMenuItem Lib "user32" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal ItemID As Long, ByVal Mode As Long, Data As MENUITEMINFO) As Long
  Private Declare Function OleTranslateColor Lib "olepro32" (ByVal Clr As Long, ByVal hPal As Long, ret As Long) As Long
  Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, lParam As Any) As Long
  Private Declare Function SendMessageAsLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal Prop As Long, ByVal NewProp As Long) As Long

Private Sub chkCheckBoxes_Click()
  If Not settingProps Then
    ExTvw.CheckBoxes = chkCheckBoxes
    If Not ExTvw.CheckBoxes Then chkCustomCheckBoxes = OLE_TRISTATE.Unchecked
  End If
End Sub

Private Sub chkContextMenus_Click()
  If Not settingProps Then
    ExTvw.ShowShellContextMenus = chkContextMenus
    ExTvw.ProcessContextMenuKeys = ExTvw.ShowShellContextMenus
  End If
End Sub

Private Sub chkCustomCheckBoxes_Click()
  If Not settingProps Then
    If chkCustomCheckBoxes Then
      ExTvw.hStateImageList(False) = ImageList_Duplicate(ImgLst.hIml)
    Else
      ExTvw.hStateImageList(False) = 0
    End If
    chkCheckBoxes = Abs(ExTvw.CheckBoxes)
  End If
End Sub

Private Sub chkDimHidden_Click()
  If Not settingProps Then
    ExTvw.ShowHiddenDimmed = chkDimHidden
  End If
End Sub

Private Sub chkDimSelected_Click()
  If Not settingProps Then
    ExTvw.DimSelectedItemsIcon = chkDimSelected
  End If
End Sub

Private Sub chkDragDrop_Click()
  If Not settingProps Then
    ExTvw.AllowDragDrop = chkDragDrop
  End If
End Sub

Private Sub chkFullRowSelect_Click()
  If Not settingProps Then
    ExTvw.FullRowSelect = chkFullRowSelect
    getProps
  End If
End Sub

Private Sub chkHasLinesAtRoot_Click()
  If Not settingProps Then
    If chkHasLinesAtRoot Then
      ExTvw.LineStyle = LineStyleConstants.lsLinesAtRoot
    Else
      ExTvw.LineStyle = LineStyleConstants.lsLinesAtItem
    End If
  End If
End Sub

Private Sub chkHotTracking_Click()
  If Not settingProps Then
    ExTvw.HotTracking = chkHotTracking
  End If
End Sub

Private Sub chkIcons_Click()
  If Not settingProps Then
    ExTvw.ShowIcons = chkIcons
  End If
End Sub

Private Sub chkInfoTips_Click()
  If Not settingProps Then
    If chkInfoTips Then
      ExTvw.ShownToolTips = ExTvw.ShownToolTips Or ShownToolTipsConstants.sttInfoTips
    Else
      ExTvw.ShownToolTips = ExTvw.ShownToolTips And Not ShownToolTipsConstants.sttInfoTips
    End If
  End If
End Sub

Private Sub chkOverlays_Click()
  If Not settingProps Then
    If chkOverlays Then
      ExTvw.ShownOverlays = ShownOverlaysConstants.soAll
    Else
      ExTvw.ShownOverlays = ShownOverlaysConstants.soNone
    End If
  End If
End Sub

Private Sub chkRenaming_Click()
  If Not settingProps Then
    If chkRenaming Then
      ExTvw.Renaming = RenamingConstants.rAutomatic
    Else
      ExTvw.Renaming = RenamingConstants.rManual
    End If
  End If
End Sub

Private Sub chkShowRoot_Click()
  If Not settingProps Then
    ExTvw.ShowRoot = chkShowRoot
  End If
End Sub

Private Sub chkShowSelection_Click()
  If Not settingProps Then
    ExTvw.ShowSelection = chkShowSelection
  End If
End Sub

Private Sub chkToolTips_Click()
  If Not settingProps Then
    If chkToolTips Then
      ExTvw.ShownToolTips = ExTvw.ShownToolTips Or ShownToolTipsConstants.sttHiddenItems
    Else
      ExTvw.ShownToolTips = ExTvw.ShownToolTips And Not ShownToolTipsConstants.sttHiddenItems
    End If
  End If
End Sub

Private Sub chkTVWStyle_Click(Index As Integer)
  Dim tmp As TreeViewStyleConstants

  If Not settingProps Then
    If chkTVWStyle(0) Then tmp = tmp Or TreeViewStyleConstants.tvsButtons
    If chkTVWStyle(1) Then tmp = tmp Or TreeViewStyleConstants.tvsLines
    ExTvw.TreeViewStyle = tmp
    getProps
  End If
End Sub

Private Sub chkUseFilters_Click()
  If Not settingProps Then
    If cmbType.ListIndex = 0 Then
      ExTvw.UseFileFilters = chkUseFilters
    Else
      ExTvw.UseFolderFilters = chkUseFilters
    End If
  End If
End Sub

Private Sub chkUseSelected_Click()
  If Not settingProps Then
    ExTvw.UseSelectedIcons = chkUseSelected
  End If
End Sub

Private Sub cmbType_Click()
  Dim Backup As Boolean
  Dim tmp As AttributesConstants

  Backup = settingProps
  settingProps = True
  If cmbType.ListIndex = 0 Then
    chkUseFilters = Abs(ExTvw.UseFileFilters)
    txtFilters = ExTvw.FileFilters
    tmp = ExTvw.FileAttributes
  Else
    chkUseFilters = Abs(ExTvw.UseFolderFilters)
    txtFilters = ExTvw.FolderFilters
    tmp = ExTvw.FolderAttributes
  End If
  With lstAttributes
    .Selected(0) = Abs((tmp And AttributesConstants.attArchive) <> 0)
    .Selected(1) = Abs((tmp And AttributesConstants.attCompressed) <> 0)
    .Selected(2) = Abs((tmp And AttributesConstants.attEncrypted) <> 0)
    .Selected(3) = Abs((tmp And AttributesConstants.attHidden) <> 0)
    .Selected(4) = Abs((tmp And AttributesConstants.attReadOnly) <> 0)
    .Selected(5) = Abs((tmp And AttributesConstants.attSystem) <> 0)
    .ListIndex = 0
  End With
  settingProps = Backup
End Sub

Private Sub cmdAbout_Click()
  ExTvw.About
End Sub

Private Sub cmdBrowse_Click(Index As Integer)
  Dim frm As frmBrowseForFolder

  Set frm = New frmBrowseForFolder
  frm.showIt Index, IIf(Index, txtRoot, "Desktop"), IIf(Index, txtPath, txtRoot), ExTvw.IncludedItems, ExTvw.FileAttributes, ExTvw.FolderAttributes, IIf(ExTvw.UseFileFilters, ExTvw.FileFilters, ""), IIf(ExTvw.UseFolderFilters, ExTvw.FolderFilters, ""), ExTvw.ExpandArchives
  If frm.Selected <> "" Then
    If Index = 1 Then
      txtPath = frm.Selected
      ExTvw.Path = txtPath
    Else
      txtRoot = frm.Selected
      ExTvw.RootPath = txtRoot
    End If
    getProps
  End If
  Set frm = Nothing
End Sub

Private Sub cmdExit_Click()
  ' zum Beenden Unload Me nutzen!
  Unload Me
End Sub

Private Sub cmdReload_Click()
  ExTvw.ReloadItems
  getProps
End Sub

Private Sub ExTvw_BeforeInsertItem(ByVal pIDLToDesktop As Long, ByVal hParentItem As Long, Cancel As Boolean)
  Debug.Print "BeforeInsertItem"
End Sub

Private Sub ExTvw_BeforeShowEditContextMenu(Cancel As Boolean)
  Debug.Print "BeforeShowEditContextMenu"
End Sub

Private Sub ExTvw_Click()
  Debug.Print "Click"
End Sub

Private Sub ExTvw_DblClick()
  Debug.Print "DblClick"
End Sub

Private Sub ExTvw_DestroyLabelEditControl()
  Debug.Print "DestroyLabelEditControl"
End Sub

Private Sub ExTvw_DragDrop(Source As Control, x As Single, y As Single)
  Debug.Print "DragDrop"
End Sub

Private Sub ExTvw_DragEnter(Data As DataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim i As Integer
  Dim txt As String

  On Error Resume Next
  For i = 1 To Data.Files.Count
    txt = txt & ", " & Data.Files(i)
  Next
  Debug.Print "DragEnter: " & Mid(txt, 3)
End Sub

Private Sub ExTvw_DragLeave()
  Debug.Print "DragLeave"
End Sub

Private Sub ExTvw_DragMove(Data As DataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'  Dim i As Integer
'  Dim txt As String
'
'  On Error Resume Next
'  For i = 1 To Data.Files.Count
'    txt = txt & ", " & Data.Files(i)
'  Next
'  Debug.Print "DragMove: " & Mid(txt, 3)
End Sub

Private Sub ExTvw_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  Debug.Print "DragOver"
End Sub

Private Sub ExTvw_Drop(Data As DataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim i As Integer
  Dim txt As String

  On Error Resume Next
  For i = 1 To Data.Files.Count
    txt = txt & ", " & Data.Files(i)
  Next
  Debug.Print "Drop: " & Mid(txt, 3)
End Sub

Private Sub ExTvw_EditClick()
  Debug.Print "EditClick"
End Sub

Private Sub ExTvw_EditDblClick()
  Debug.Print "EditDblClick"
End Sub

Private Sub ExTvw_EditKeyDown(KeyCode As Integer, Shift As Integer)
  Debug.Print "EditKeyDown " & Chr(KeyCode)
End Sub

Private Sub ExTvw_EditKeyPress(KeyAscii As Integer)
  Debug.Print "EditKeyPress " & Chr(KeyAscii)
End Sub

Private Sub ExTvw_EditKeyUp(KeyCode As Integer, Shift As Integer)
  Debug.Print "EditKeyUp " & Chr(KeyCode)
End Sub

Private Sub ExTvw_EditMClick()
  Debug.Print "EditMClick"
End Sub

Private Sub ExTvw_EditMDblClick()
  Debug.Print "EditMDblClick"
End Sub

Private Sub ExTvw_EditMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  Debug.Print "EditMouseDown"
End Sub

Private Sub ExTvw_EditMouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  Debug.Print "EditMouseEnter"
End Sub

Private Sub ExTvw_EditMouseHover(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  Debug.Print "EditMouseHover"
End Sub

Private Sub ExTvw_EditMouseLeave(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  Debug.Print "EditMouseLeave"
End Sub

Private Sub ExTvw_EditMouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'  Debug.Print "EditMouseMove"
End Sub

Private Sub ExTvw_EditMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  Debug.Print "EditMouseUp"
End Sub

Private Sub ExTvw_EditRClick()
  Debug.Print "EditRClick"
End Sub

Private Sub ExTvw_EditRDblClick()
  Debug.Print "EditRDblClick"
End Sub

Private Sub ExTvw_GotFocus()
  Debug.Print "GotFocus"
End Sub

Private Sub ExTvw_InitLabelEditControl()
  Debug.Print "InitLabelEditControl"
End Sub

Private Sub ExTvw_ItemBeforeExecuteShellContextMenuItem(ByVal hItem As Long, ByVal ItemID As Long, ByVal ItemVerb As String, ByVal ItemDescription As String, Cancel As Boolean, WindowMode As ExTvw.WindowModeConstants)
  Debug.Print "ItemBeforeExecuteShellContextMenuItem " & ExTvw.ItemHandleToDisplayName(hItem) & " - Verb: " & ItemVerb & ", Description: " & ItemDescription
  If ItemID = 999 Then
    ' unser Menuitem
    MsgBox "Hallo Welt!"
  End If
End Sub

Private Sub ExTvw_ItemBeforeGetInfoTip(ByVal hItem As Long, InfoTipStyle As ExTvw.InfoTipStyleConstants, Cancel As Boolean)
  Debug.Print "ItemBeforeGetInfoTip " & ExTvw.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw_ItemBeforeInitShellContextMenu(ByVal hItem As Long, ContextMenuStyle As ExTvw.ShellContextMenuStyleConstants, Cancel As Boolean)
  Debug.Print "ItemBeforeInitShellContextMenu " & ExTvw.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw_ItemBeforeRenaming(ByVal hItem As Long, ByVal OldName As String, NewName As String, Cancel As Boolean)
  Debug.Print "ItemBeforeRenaming " & ExTvw.ItemHandleToDisplayName(hItem) & " from " & OldName & " to " & NewName
End Sub

Private Sub ExTvw_ItemBeforeShowShellContextMenu(ByVal hItem As Long, Cancel As Boolean)
  Debug.Print "ItemBeforeShowShellContextMenu " & ExTvw.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw_ItemBeforeShowToolTip(ByVal hItem As Long, ByVal ToolTipType As ExTvw.ToolTipTypeConstants, ToolTipText As String, Cancel As Boolean)
  Const GWL_EXSTYLE = -20
  Const GWL_STYLE = -16
  Const WM_USER = &H400
  Const TTM_SETTIPBKCOLOR = (WM_USER + 19)
  Const TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
  Const TTM_SETTITLE = (WM_USER + 32)
  Const TTS_BALLOON = &H40
  Const WS_BORDER = &H800000
  Const WS_EX_WINDOWEDGE = &H100&
  Dim ExStyle As Long
  Dim Style As Long

  Debug.Print "ItemBeforeShowToolTip " & ExTvw.ItemHandleToDisplayName(hItem) & ", " & ToolTipText

  If (chkCustomInfoTips = vbChecked) And (ToolTipType = ToolTipTypeConstants.tttInfoTip) Then
    ExStyle = GetWindowLong(ExTvw.hWndToolTip, GWL_EXSTYLE)
    Style = GetWindowLong(ExTvw.hWndToolTip, GWL_STYLE)
    Style = Style Or TTS_BALLOON
    Style = Style And Not WS_BORDER
    SetWindowLong ExTvw.hWndToolTip, GWL_EXSTYLE, ExStyle
    ExStyle = ExStyle And Not WS_EX_WINDOWEDGE
    SetWindowLong ExTvw.hWndToolTip, GWL_STYLE, Style

    ToolTipText = ToolTipText & vbNewLine & vbNewLine & "Hallo Welt"
    SendMessageAsLong ExTvw.hWndToolTip, TTM_SETTIPBKCOLOR, RGB(224, 224, 224), 0
    SendMessageAsLong ExTvw.hWndToolTip, TTM_SETTIPTEXTCOLOR, RGB(220, 0, 0), 0
    SendMessageAsLong ExTvw.hWndToolTip, TTM_SETTITLE, 1, StrPtr(StrConv("InfoTip:", VbStrConv.vbFromUnicode))
  Else
    ExStyle = GetWindowLong(ExTvw.hWndToolTip, GWL_EXSTYLE)
    Style = GetWindowLong(ExTvw.hWndToolTip, GWL_STYLE)
    Style = Style And Not TTS_BALLOON
    Style = Style Or WS_BORDER
    SetWindowLong ExTvw.hWndToolTip, GWL_EXSTYLE, ExStyle
    ExStyle = ExStyle Or WS_EX_WINDOWEDGE
    SetWindowLong ExTvw.hWndToolTip, GWL_STYLE, Style

    SendMessageAsLong ExTvw.hWndToolTip, TTM_SETTIPBKCOLOR, TranslateColor(SystemColorConstants.vbInfoBackground), 0
    SendMessageAsLong ExTvw.hWndToolTip, TTM_SETTIPTEXTCOLOR, TranslateColor(SystemColorConstants.vbInfoText), 0
    SendMessageAsLong ExTvw.hWndToolTip, TTM_SETTITLE, 0, StrPtr("")
  End If
End Sub

Private Sub ExTvw_ItemBeforeStartRenaming(ByVal hItem As Long, Cancel As Boolean)
  Debug.Print "ItemBeforeStartRenaming " & ExTvw.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw_ItemBeforeStateIconChange(ByVal hItem As Long, ByVal OldStateIcon As Long, NewStateIcon As Long, Cancel As Boolean)
  Debug.Print "ItemBeforeStateIconChange " & ExTvw.ItemHandleToDisplayName(hItem) & " from " & OldStateIcon & " to " & NewStateIcon
  If NewStateIcon = 3 Then NewStateIcon = 1
End Sub

Private Sub ExTvw_ItemClick(ByVal hItem As Long)
  Debug.Print "ItemClick " & ExTvw.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw_ItemCollapsed(ByVal hItem As Long)
  Debug.Print "ItemCollapsed " & ExTvw.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw_ItemCollapsing(ByVal hItem As Long, Cancel As Boolean)
  Debug.Print "ItemCollapsing " & ExTvw.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw_ItemDblClick(ByVal hItem As Long)
  Debug.Print "ItemDblClick " & ExTvw.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw_ItemDestroyShellContextMenu(ByVal hItem As Long)
  Debug.Print "ItemDestroyShellContextMenu " & ExTvw.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw_ItemDragEnter(ByVal hItem As Long, Data As DataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim i As Integer
  Dim txt As String

  On Error Resume Next
  For i = 1 To Data.Files.Count
    txt = txt & ", " & Data.Files(i)
  Next
  Debug.Print "ItemDragEnter: " & Mid(txt, 3) & " over " & ExTvw.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw_ItemDragLeave(ByVal hItem As Long)
  Debug.Print "ItemDragLeave " & ExTvw.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw_ItemDragMove(ByVal hItem As Long, Data As DataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'  Dim i As Integer
'  Dim txt As String
'
'  On Error Resume Next
'  For i = 1 To Data.Files.Count
'    txt = txt & ", " & Data.Files(i)
'  Next
'  Debug.Print "ItemDragMove: " & Mid(txt, 3) & " over " & ExTvw.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw_ItemDrop(ByVal hItem As Long, Data As DataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim i As Integer
  Dim txt As String

  On Error Resume Next
  For i = 1 To Data.Files.Count
    txt = txt & ", " & Data.Files(i)
  Next
  Debug.Print "ItemDrop: " & Mid(txt, 3) & " on " & ExTvw.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw_ItemExecutedShellContextMenuItem(ByVal hItem As Long, ByVal ItemID As Long, ByVal ItemVerb As String, ByVal ItemDescription As String)
  Debug.Print "ItemExecutedShellContextMenuItem " & ExTvw.ItemHandleToDisplayName(hItem) & " - Verb: " & ItemVerb & ", Description: " & ItemDescription
End Sub

Private Sub ExTvw_ItemExpanded(ByVal hItem As Long)
  Debug.Print "ItemExpanded " & ExTvw.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw_ItemExpanding(ByVal hItem As Long, Cancel As Boolean)
  Debug.Print "ItemExpanding " & ExTvw.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw_ItemInitShellContextMenu(ByVal hItem As Long)
  Const MFT_SEPARATOR = &H800&
  Const MFT_STRING = &H0&
  Const MIIM_ID = &H2
  Const MIIM_TYPE = &H10
  Dim Data As MENUITEMINFO

  Debug.Print "ItemInitShellContextMenu " & ExTvw.ItemHandleToDisplayName(hItem)
  If chkCustomMenu Then
    With Data
      .fMask = MIIM_ID Or MIIM_TYPE
      .fType = MFT_SEPARATOR
      .cbSize = Len(Data)
      InsertMenuItem ExTvw.hShellContextMenu, GetMenuItemCount(ExTvw.hShellContextMenu), 1, Data

      .dwTypeData = "mein MenuItem für den Item " & ExTvw.ItemHandleToDisplayName(hItem)
      .cch = Len(.dwTypeData)
      .fMask = MIIM_ID Or MIIM_TYPE
      .fType = MFT_STRING
      ' die ID für unseren Menuitem
      .wID = 999
      .cbSize = Len(Data)
      InsertMenuItem ExTvw.hShellContextMenu, GetMenuItemCount(ExTvw.hShellContextMenu), 1, Data
    End With
  End If
End Sub

Private Sub ExTvw_ItemLoadedSubItems(ByVal hItem As Long)
  Debug.Print "ItemLoadedSubItems " & ExTvw.ItemHandleToDisplayName(hItem)
  ExTvw.MousePointer = MousePointerConstants.vbDefault
  DoEvents
End Sub

Private Sub ExTvw_ItemLoadingSubItems(ByVal hItem As Long)
  Debug.Print "ItemLoadingSubItems " & ExTvw.ItemHandleToDisplayName(hItem)
  ExTvw.MousePointer = MousePointerConstants.vbHourglass
  DoEvents
End Sub

Private Sub ExTvw_ItemMClick(ByVal hItem As Long)
  Debug.Print "ItemMClick " & ExTvw.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw_ItemMDblClick(ByVal hItem As Long)
  Debug.Print "ItemMDblClick " & ExTvw.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw_ItemMouseDown(ByVal hItem As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  Debug.Print "ItemMouseDown " & ExTvw.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw_ItemMouseEnter(ByVal hItem As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  Debug.Print "ItemMouseEnter " & ExTvw.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw_ItemMouseLeave(ByVal hItem As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  Debug.Print "ItemMouseLeave " & ExTvw.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw_ItemMouseMove(ByVal hItem As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'  Debug.Print "ItemMouseMove " & ExTvw.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw_ItemMouseUp(ByVal hItem As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  Debug.Print "ItemMouseUp " & ExTvw.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw_ItemRClick(ByVal hItem As Long)
  Debug.Print "ItemRClick " & ExTvw.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw_ItemRDblClick(ByVal hItem As Long)
  Debug.Print "ItemRDblClick " & ExTvw.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw_ItemRenamed(ByVal hItem As Long, ByVal OldName As String, ByVal NewName As String)
  Debug.Print "ItemRenamed " & ExTvw.ItemHandleToDisplayName(hItem) & " from " & OldName & " to " & NewName
End Sub

Private Sub ExTvw_ItemShellContextMenuItemEnter(ByVal hItem As Long, ByVal hMenu As Long, ByVal ItemID As Long, ByVal ItemText As String, ByVal ItemVerb As String, ByVal ItemDescription As String)
  Debug.Print "ItemShellContextMenuItemEnter " & ExTvw.ItemHandleToDisplayName(hItem) & " - Text: " & ItemText & ", Verb: " & ItemVerb & ", Description: " & ItemDescription
End Sub

Private Sub ExTvw_ItemShellContextMenuItemLeave(ByVal hItem As Long, ByVal hMenu As Long, ByVal ItemID As Long, ByVal ItemText As String, ByVal ItemVerb As String, ByVal ItemDescription As String)
  Debug.Print "ItemShellContextMenuItemLeave " & ExTvw.ItemHandleToDisplayName(hItem) & " - Text: " & ItemText & ", Verb: " & ItemVerb & ", Description: " & ItemDescription
End Sub

Private Sub ExTvw_ItemStateIconChanged(ByVal hItem As Long, ByVal OldStateIcon As Long, ByVal NewStateIcon As Long)
  Const Deselected = 1
  Const Selected = 2
  Const SemiSelected = 3
  Dim foundDeselected As Boolean
  Dim foundSelected As Boolean
  Dim foundSemiSelected As Boolean
  Dim allSelected As Boolean
  Dim hParent As Long

  Debug.Print "ItemStateIconChanged " & ExTvw.ItemHandleToDisplayName(hItem) & " from " & OldStateIcon & " to " & NewStateIcon

  setStateIcons hItem, NewStateIcon
  If chkCustomCheckBoxes Then
    ' Parent-Items updaten
    With ExTvw
      hParent = .ItemGetParentItem(hItem)
      While (hParent <> -1) And (hParent <> 0)
        foundDeselected = (.ItemGetFirstSubItem(hParent, Deselected) <> -1)
        foundSelected = (.ItemGetFirstSubItem(hParent, Selected) <> -1)
        foundSemiSelected = (.ItemGetFirstSubItem(hParent, SemiSelected) <> -1)

        If foundSemiSelected Then
          .ItemStateIconIndex(hParent) = SemiSelected
        Else
          If foundSelected And foundDeselected Then
            .ItemStateIconIndex(hParent) = SemiSelected
          ElseIf foundSelected And Not foundDeselected Then
            .ItemStateIconIndex(hParent) = Selected
          ElseIf foundDeselected And Not foundSelected Then
            .ItemStateIconIndex(hParent) = Deselected
          End If
        End If

        hParent = .ItemGetParentItem(hParent)
      Wend
    End With
  End If

  With ExTvw
    Debug.Print .ItemGetSubItemCount(.ItemGetParentItem(hItem), 2) & " sub-items are checked"
  End With
End Sub

Private Sub ExTvw_KeyDown(KeyCode As Integer, Shift As Integer)
  Debug.Print "KeyDown"
End Sub

Private Sub ExTvw_KeyPress(KeyAscii As Integer)
  Debug.Print "KeyPress"
End Sub

Private Sub ExTvw_KeyUp(KeyCode As Integer, Shift As Integer)
  Debug.Print "KeyUp"
End Sub

Private Sub ExTvw_LostFocus()
  Debug.Print "LostFocus"
End Sub

Private Sub ExTvw_MClick()
  Debug.Print "MClick"
End Sub

Private Sub ExTvw_MDblClick()
  Debug.Print "MDblClick"
End Sub

Private Sub ExTvw_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Debug.Print "MouseDown"
End Sub

Private Sub ExTvw_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  Debug.Print "MouseEnter"
End Sub

Private Sub ExTvw_MouseHover(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  Debug.Print "MouseHover"
End Sub

Private Sub ExTvw_MouseLeave(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  Debug.Print "MouseLeave"
End Sub

Private Sub ExTvw_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'  Debug.Print "MouseMove"
End Sub

Private Sub ExTvw_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Debug.Print "MouseUp"
End Sub

Private Sub ExTvw_RClick()
  Debug.Print "RClick"
End Sub

Private Sub ExTvw_RDblClick()
  Debug.Print "RDblClick"
End Sub

Private Sub ExTVW_SelChanged(ByVal hOldItem As Long, ByVal hNewItem As Long, ByVal CausedBy As ExTvw.CausedByConstants)
  Debug.Print "SelChanged from " & ExTvw.ItemHandleToDisplayName(hOldItem) & " to " & ExTvw.ItemHandleToDisplayName(hNewItem)
  txtPath = ExTvw.Path
End Sub

Private Sub ExTvw_SelChanging(ByVal hOldItem As Long, ByVal hNewItem As Long, ByVal CausedBy As ExTvw.CausedByConstants, Cancel As Boolean)
  Debug.Print "SelChanging from " & ExTvw.ItemHandleToDisplayName(hOldItem) & " to " & ExTvw.ItemHandleToDisplayName(hNewItem)
End Sub

Private Sub Form_Initialize()
  ' Support für WinXP-Themes
  InitCommonControls
End Sub

Private Sub Form_Load()
  ExTvw.hWndShellUIParentWindow = Me.hWnd

  getProps
End Sub

Private Sub lstAttributes_Click()
  Dim tmp As AttributesConstants

  If Not settingProps Then
    With lstAttributes
      If .Selected(0) Then tmp = tmp Or AttributesConstants.attArchive
      If .Selected(1) Then tmp = tmp Or AttributesConstants.attCompressed
      If .Selected(2) Then tmp = tmp Or AttributesConstants.attEncrypted
      If .Selected(3) Then tmp = tmp Or AttributesConstants.attHidden
      If .Selected(4) Then tmp = tmp Or AttributesConstants.attReadOnly
      If .Selected(5) Then tmp = tmp Or AttributesConstants.attSystem
      If cmbType.ListIndex = 0 Then
        ExTvw.FileAttributes = tmp
      Else
        ExTvw.FolderAttributes = tmp
      End If
    End With
  End If
End Sub

Private Sub lstExpandArchives_Click()
  Dim tmp As ExpandArchivesConstants

  If Not settingProps Then
    With lstExpandArchives
      If .Selected(0) Then tmp = tmp Or ExpandArchivesConstants.eaACE
      If .Selected(1) Then tmp = tmp Or ExpandArchivesConstants.eaBIN
      If .Selected(2) Then tmp = tmp Or ExpandArchivesConstants.eaCAB
      If .Selected(3) Then tmp = tmp Or ExpandArchivesConstants.eaISO
      If .Selected(4) Then tmp = tmp Or ExpandArchivesConstants.eaJAR
      If .Selected(5) Then tmp = tmp Or ExpandArchivesConstants.eaRAR
      If .Selected(6) Then tmp = tmp Or ExpandArchivesConstants.eaZIP

      ExTvw.ExpandArchives = tmp
    End With
  End If
End Sub

Private Sub lstIncludedItems_ItemCheck(Item As Integer)
  Dim tmp As IncludedItemsConstants

  If Not settingProps Then
    With lstIncludedItems
      If .Selected(0) Then tmp = tmp Or IncludedItemsConstants.iiFSFolders
      If .Selected(1) Then tmp = tmp Or IncludedItemsConstants.iiFSFiles
      If .Selected(2) Then tmp = tmp Or IncludedItemsConstants.iiNonFSFolders
      If .Selected(3) Then tmp = tmp Or IncludedItemsConstants.iiNonFSFiles

      ExTvw.IncludedItems = tmp
    End With
  End If
End Sub

Private Sub optSingleExpand_Click(Index As Integer)
  If Not settingProps Then
    ExTvw.SingleExpand = Index
  End If
End Sub

Private Sub txtFilters_LostFocus()
  If cmbType.ListIndex = 0 Then
    ExTvw.FileFilters = txtFilters
  Else
    ExTvw.FolderFilters = txtFilters
  End If
End Sub

Private Sub txtPath_LostFocus()
  ExTvw.Path = txtPath
  getProps
End Sub

Private Sub txtRoot_LostFocus()
  ExTvw.RootPath = txtRoot
  getProps
End Sub


' synchronisiert die CheckBoxes mit den Eigenschaften des TreeViews
Private Sub getProps()
  settingProps = True
  With ExTvw
    txtRoot = .RootPath
    txtPath = .Path
    chkTVWStyle(0) = Abs((.TreeViewStyle And TreeViewStyleConstants.tvsButtons) <> 0)
    chkTVWStyle(1) = Abs((.TreeViewStyle And TreeViewStyleConstants.tvsLines) <> 0)
    chkHasLinesAtRoot = Abs(.LineStyle = LineStyleConstants.lsLinesAtRoot)
    chkIcons = Abs(.ShowIcons)
    chkShowSelection = Abs(.ShowSelection)
    chkCheckBoxes = Abs(.CheckBoxes)
    chkFullRowSelect = Abs(.FullRowSelect)
    chkHotTracking = Abs(.HotTracking)
    chkRenaming = Abs(.Renaming = RenamingConstants.rAutomatic)
    chkContextMenus = Abs(.ShowShellContextMenus)
    chkShowRoot = Abs(.ShowRoot)
    chkDimHidden = Abs(.ShowHiddenDimmed)
    chkDimSelected = Abs(.DimSelectedItemsIcon)
    chkUseSelected = Abs(.UseSelectedIcons)
    chkOverlays = Abs(.ShownOverlays <> ShownOverlaysConstants.soNone)
    chkDragDrop = Abs(.AllowDragDrop)
    chkInfoTips = Abs((.ShownToolTips And ShownToolTipsConstants.sttInfoTips) = ShownToolTipsConstants.sttInfoTips)
    chkToolTips = Abs((.ShownToolTips And ShownToolTipsConstants.sttHiddenItems) = ShownToolTipsConstants.sttHiddenItems)
    cmbType.ListIndex = 0
    optSingleExpand(.SingleExpand) = True
    lstIncludedItems.Selected(0) = Abs((.IncludedItems And IncludedItemsConstants.iiFSFolders) <> 0)
    lstIncludedItems.Selected(1) = Abs((.IncludedItems And IncludedItemsConstants.iiFSFiles) <> 0)
    lstIncludedItems.Selected(2) = Abs((.IncludedItems And IncludedItemsConstants.iiNonFSFolders) <> 0)
    lstIncludedItems.Selected(3) = Abs((.IncludedItems And IncludedItemsConstants.iiNonFSFiles) <> 0)
    lstExpandArchives.Selected(0) = Abs((.ExpandArchives And ExpandArchivesConstants.eaACE) <> 0)
    lstExpandArchives.Selected(1) = Abs((.ExpandArchives And ExpandArchivesConstants.eaBIN) <> 0)
    lstExpandArchives.Selected(2) = Abs((.ExpandArchives And ExpandArchivesConstants.eaCAB) <> 0)
    lstExpandArchives.Selected(3) = Abs((.ExpandArchives And ExpandArchivesConstants.eaISO) <> 0)
    lstExpandArchives.Selected(4) = Abs((.ExpandArchives And ExpandArchivesConstants.eaJAR) <> 0)
    lstExpandArchives.Selected(5) = Abs((.ExpandArchives And ExpandArchivesConstants.eaRAR) <> 0)
    lstExpandArchives.Selected(6) = Abs((.ExpandArchives And ExpandArchivesConstants.eaZIP) <> 0)
  End With
  settingProps = False
End Sub

Private Sub setStateIcons(ByVal hParentItem As Long, ByVal IconIndex As Long)
  Dim hChild As Long

  With ExTvw
    ' alle Child-Nodes (de-)selektieren
    hChild = .ItemGetFirstSubItem(hParentItem)
    While hChild <> -1
      .ItemStateIconIndex(hChild) = IconIndex
      setStateIcons hChild, IconIndex
      hChild = .ItemGetNextItem(hChild)
    Wend
  End With
End Sub

' transformiert eine OLE-Farbe in eine RGB-Farbe
Private Function TranslateColor(ByVal Clr As OLE_COLOR) As Long
  If OleTranslateColor(Clr, 0, TranslateColor) Then TranslateColor = &HFFFF
End Function
