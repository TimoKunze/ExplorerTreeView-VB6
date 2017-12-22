VERSION 5.00
Object = "{CAB249C7-BAC9-4C51-9526-12F29E40C4CE}#2.7#0"; "ExTvw.ocx"
Begin VB.Form frmTest 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manuell
   ScaleHeight     =   328
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   372
   StartUpPosition =   2  'Bildschirmmitte
   Begin ExTvw.ExplorerTreeView ExTvw1 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   7011
      DragExpandTime  =   1000
      DragScrollTime  =   200
      ExpandArchives  =   32
      FadeExpandos    =   -1  'True
      FastInfoTipsOnly=   -1  'True
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
      IncludedItems   =   15
      Indent          =   19
      ItemHeight      =   17
      ShowFocusRect   =   0   'False
      ShownToolTips   =   3
      TreeViewStyle   =   1
      UseStandardIconsForSlowItems=   -1  'True
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5040
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Command1"
      Height          =   375
      Left            =   240
      OLEDropMode     =   1  'Manuell
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Main"
      Begin VB.Menu mnuItem1 
         Caption         =   "Item 1"
      End
      Begin VB.Menu mnuItem2 
         Caption         =   "Item 2"
      End
      Begin VB.Menu mnuItem3 
         Caption         =   "Item 3"
      End
      Begin VB.Menu mnuItem 
         Caption         =   "Item 4"
      End
      Begin VB.Menu mnuItem5 
         Caption         =   "Item 5"
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Private IDropTargetHelper As IVBDropTargetHelper
  Private IID_IDragDropHelper As UUID
  Private IID_IDropTargetHelper As UUID
  Private Start As Long
  Private themeableOS As Boolean

  Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal Size As Long)
  Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
  Private Declare Function GetTickCount Lib "kernel32" () As Long
  Private Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
  Private Declare Function GlobalSize Lib "kernel32.dll" (ByVal hMem As Long) As Long
  Private Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
  Private Declare Function ImageList_Duplicate Lib "comctl32" (ByVal hIml As Long) As Long
  Private Declare Sub InitCommonControls Lib "comctl32" ()
  Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
  Private Declare Function RegisterClipboardFormat Lib "user32.dll" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
  Private Declare Sub ReleaseStgMedium Lib "ole32.dll" (ByVal lpstgmedium As Long)
  Private Declare Function SetWindowTheme Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pSubAppName As Long, ByVal pSubIDList As Long) As Long


Private Sub Command1_Click()
  ExTvw1.ReloadSubItems ExTvw1.SelectedItem
End Sub

Private Sub Command1_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim pDO As Object
  Dim pt As POINT

  Effect = vbDropEffectNone
  pt.x = CLng(x)
  pt.y = CLng(y)
  CopyMemory pDO, ByVal ObjPtr(data) + 16, 4

  If Not (IDropTargetHelper Is Nothing) Then IDropTargetHelper.Drop ByVal pDO, pt, Effect

  ZeroMemory pDO, 4
End Sub

Private Sub Command1_OLEDragOver(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
  Dim pDO As Object
  Dim pt As POINT

  Effect = vbDropEffectNone
  pt.x = CLng(x)
  pt.y = CLng(y)
  CopyMemory pDO, ByVal ObjPtr(data) + 16, 4

  Select Case State
    Case DragOverConstants.vbEnter
      If Not (IDropTargetHelper Is Nothing) Then IDropTargetHelper.DragEnter Command1.hWnd, ByVal pDO, pt, Effect
    Case DragOverConstants.vbLeave
      If Not (IDropTargetHelper Is Nothing) Then IDropTargetHelper.DragLeave
    Case DragOverConstants.vbOver
      If Not (IDropTargetHelper Is Nothing) Then IDropTargetHelper.DragOver pt, Effect
  End Select

  ZeroMemory pDO, 4
End Sub

Private Sub ExTvw1_BeforeInsertItem(ByVal pIDLToDesktop As Long, ByVal hParentItem As Long, Cancel As Boolean)
'  Debug.Print "BeforeInsertItem: " & ExTvw1.ItemHandleToDisplayName(hParentItem) & ", " & pIDLToDesktop
End Sub

Private Sub ExTvw1_BeforeShowEditContextMenu(Cancel As Boolean)
'  Debug.Print "BeforeShowEditContextMenu"
End Sub

Private Sub ExTvw1_Click()
'  Debug.Print "Click"
End Sub

Private Sub ExTvw1_CompleteDrag(ByVal Canceled As Boolean, ByVal PerformedEffect As Long)
  Debug.Print "CompleteDrag Canceled=" & Canceled & ", PerformedEffect=" & PerformedEffect
End Sub

Private Sub ExTvw1_DblClick()
'  Debug.Print "DblClick"
End Sub

Private Sub ExTvw1_DestroyLabelEditControl()
'  Debug.Print "DestroyLabelEditControl"
End Sub

Private Sub ExTvw1_DragDrop(Source As Control, x As Single, y As Single)
'  Debug.Print "DragDrop"
End Sub

Private Sub ExTvw1_DragEnter(data As DataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'  On Error Resume Next
'  Debug.Print "DragEnter: " & Data.Files(1)
End Sub

Private Sub ExTvw1_DragLeave()
'  Debug.Print "DragLeave"
End Sub

Private Sub ExTvw1_DragMove(data As DataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'  On Error Resume Next
'  Debug.Print "DragMove: " & Data.Files(1)
End Sub

Private Sub ExTvw1_DragOver(Source As Control, x As Single, y As Single, State As Integer)
'  Debug.Print "DragOver"
End Sub

Private Sub ExTvw1_Drop(data As DataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'  On Error Resume Next
'  Debug.Print "Drop: " & Data.Files(1)
End Sub

Private Sub ExTvw1_EditClick()
'  Debug.Print "EditClick"
End Sub

Private Sub ExTvw1_EditDblClick()
'  Debug.Print "EditDblClick"
End Sub

Private Sub ExTvw1_EditKeyDown(KeyCode As Integer, Shift As Integer)
'  Debug.Print "EditKeyDown " & Chr(KeyCode)
End Sub

Private Sub ExTvw1_EditKeyPress(KeyAscii As Integer)
'  Debug.Print "EditKeyPress " & Chr(KeyAscii)
End Sub

Private Sub ExTvw1_EditKeyUp(KeyCode As Integer, Shift As Integer)
'  Debug.Print "EditKeyUp " & Chr(KeyCode)
End Sub

Private Sub ExTvw1_EditMClick()
'  Debug.Print "EditMClick"
End Sub

Private Sub ExTvw1_EditMDblClick()
'  Debug.Print "EditMDblClick"
End Sub

Private Sub ExTvw1_EditMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'  Debug.Print "EditMouseDown"
End Sub

Private Sub ExTvw1_EditMouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'  Debug.Print "EditMouseEnter"
End Sub

Private Sub ExTvw1_EditMouseHover(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'  Debug.Print "EditMouseHover"
End Sub

Private Sub ExTvw1_EditMouseLeave(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'  Debug.Print "EditMouseLeave"
End Sub

Private Sub ExTvw1_EditMouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'  Debug.Print "EditMouseMove"
End Sub

Private Sub ExTvw1_EditMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'  Debug.Print "EditMouseUp"
End Sub

Private Sub ExTvw1_EditRClick()
'  Debug.Print "EditRClick"
End Sub

Private Sub ExTvw1_EditRDblClick()
'  Debug.Print "EditRDblClick"
End Sub

Private Sub ExTvw1_GotFocus()
'  Debug.Print "GotFocus"
End Sub

Private Sub ExTvw1_InitLabelEditControl()
'  Debug.Print "InitLabelEditControl"
End Sub

Private Sub ExTvw1_ItemBeforeExecuteShellContextMenuItem(ByVal hItem As Long, ByVal ItemID As Long, ByVal ItemVerb As String, ByVal ItemDescription As String, Cancel As Boolean, WindowMode As ExTvw.WindowModeConstants)
'  Debug.Print "ItemBeforeExecuteShellContextMenuItem " & ExTvw1.ItemHandleToDisplayName(hItem) & " - Verb: " & ItemVerb & ", Description: " & ItemDescription
End Sub

Private Sub ExTvw1_ItemBeforeGetInfoTip(ByVal hItem As Long, InfoTipStyle As ExTvw.InfoTipStyleConstants, Cancel As Boolean)
'  Debug.Print "ItemBeforeGetInfoTip " & ExTvw1.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw1_ItemBeforeInitShellContextMenu(ByVal hItem As Long, ContextMenuStyle As ExTvw.ShellContextMenuStyleConstants, Cancel As Boolean)
'  Debug.Print "ItemBeforeInitShellContextMenu " & ExTvw1.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw1_ItemBeforeRenaming(ByVal hItem As Long, ByVal OldName As String, NewName As String, Cancel As Boolean)
'  Debug.Print "ItemBeforeRenaming " & ExTvw1.ItemHandleToDisplayName(hItem) & " from " & OldName & " to " & NewName
End Sub

Private Sub ExTvw1_ItemBeforeShowShellContextMenu(ByVal hItem As Long, Cancel As Boolean)
'  Debug.Print "ItemBeforeShowShellContextMenu " & ExTvw1.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw1_ItemBeforeShowToolTip(ByVal hItem As Long, ByVal ToolTipType As ExTvw.ToolTipTypeConstants, ToolTipText As String, Cancel As Boolean)
'  Debug.Print "ItemBeforeShowToolTip " & ExTvw1.ItemHandleToDisplayName(hItem) & ", " & ToolTipText
End Sub

Private Sub ExTvw1_ItemBeforeStartRenaming(ByVal hItem As Long, Cancel As Boolean)
'  Debug.Print "ItemBeforeStartRenaming " & ExTvw1.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw1_ItemBeforeStateIconChange(ByVal hItem As Long, ByVal OldStateIcon As Long, NewStateIcon As Long, Cancel As Boolean)
'  Debug.Print "ItemBeforeStateIconChange " & ExTvw1.ItemHandleToDisplayName(hItem) & " from " & OldStateIcon & " to " & NewStateIcon
  If NewStateIcon = 3 Then NewStateIcon = 1
End Sub

Private Sub ExTvw1_ItemClick(ByVal hItem As Long)
'  Debug.Print "ItemClick " & ExTvw1.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw1_ItemCollapsed(ByVal hItem As Long)
'  Debug.Print "ItemCollapsed " & ExTvw1.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw1_ItemCollapsing(ByVal hItem As Long, Cancel As Boolean)
'  Debug.Print "ItemCollapsing " & ExTvw1.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw1_ItemCompleteDrag(ByVal hItem As Long, ByVal Canceled As Boolean, ByVal PerformedEffect As Long)
'  Debug.Print "ItemCompleteDrag " & ExTvw1.ItemHandleToDisplayName(hItem) & ", Canceled=" & Canceled & ", PerformedEffect=" & PerformedEffect
End Sub

Private Sub ExTvw1_ItemDblClick(ByVal hItem As Long)
'  Debug.Print "ItemDblClick " & ExTvw1.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw1_ItemDestroyShellContextMenu(ByVal hItem As Long)
'  Debug.Print "ItemDestroyShellContextMenu " & ExTvw1.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw1_ItemDragEnter(ByVal hItem As Long, data As DataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'  On Error Resume Next
'  Debug.Print "ItemDragEnter: " & Data.Files(1) & " over " & ExTvw1.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw1_ItemDragLeave(ByVal hItem As Long)
'  Debug.Print "ItemDragLeave " & ExTvw1.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw1_ItemDragMove(ByVal hItem As Long, data As DataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'  On Error Resume Next
'  Debug.Print "ItemDragMove: " & Data.Files(1) & " over " & ExTvw1.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw1_ItemDrop(ByVal hItem As Long, data As DataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'  Const DVASPECT_CONTENT = 1
'  Const S_OK = &H0
'  Dim arraySize As Long
'  Dim dataFmt As FORMATETC
'  Dim preferedEffect As Long
'  Dim IDataObj As IVBDataObject
'  Dim pBinaryData As Long
'  Dim stgData As STGMEDIUM
'
'  CopyMemory IDataObj, ByVal ObjPtr(data) + 16, 4
'  With dataFmt
'    .cfFormat = LoWord(RegisterClipboardFormat("Preferred DropEffect"))
'    .dwAspect = DVASPECT_CONTENT
'    .lindex = -1
'  End With
'  If IDataObj.GetData(dataFmt, stgData) = S_OK Then
'    arraySize = GlobalSize(stgData.data)
'    pBinaryData = GlobalLock(stgData.data)
'    MsgBox pBinaryData
'    CopyMemory preferedEffect, ByVal pBinaryData, IIf(arraySize < LenB(preferedEffect), arraySize, LenB(preferedEffect))
'    GlobalUnlock stgData.data
'    ReleaseStgMedium VarPtr(stgData)
'    MsgBox preferedEffect
'  End If
'
'  ' unerwünschte <.Release>-Aufrufe vermeiden
'  If Not (IDataObj Is Nothing) Then ZeroMemory IDataObj, 4
'
'  On Error Resume Next
'  Debug.Print "ItemDrop: " & Data.Files(1) & " on " & ExTvw1.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw1_ItemExecutedShellContextMenuItem(ByVal hItem As Long, ByVal ItemID As Long, ByVal ItemVerb As String, ByVal ItemDescription As String)
'  Debug.Print "ItemExecutedShellContextMenuItem " & ExTvw1.ItemHandleToDisplayName(hItem) & " - Verb: " & ItemVerb & ", Description: " & ItemDescription
End Sub

Private Sub ExTvw1_ItemExpanded(ByVal hItem As Long)
'  Debug.Print "ItemExpanded " & ExTvw1.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw1_ItemExpanding(ByVal hItem As Long, Cancel As Boolean)
'  Debug.Print "ItemExpanding " & ExTvw1.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw1_ItemInitShellContextMenu(ByVal hItem As Long)
'  Debug.Print "ItemInitShellContextMenu " & ExTvw1.ItemHandleToDisplayName(hItem)
'Dim a As MENUITEMINFO
'    a.dwTypeData = "Test"
'    a.cch = Len(a.dwTypeData)
'    a.fMask = &H2 Or &H10
'    a.fType = &H0
'    a.wID = 999
'    a.cbSize = Len(a)
'    InsertMenuItem ExTvw1.hShellContextMenu, 1, 1, a
End Sub

Private Sub ExTvw1_ItemLoadedSubItems(ByVal hItem As Long)
'  MsgBox "needed " & GetTickCount - Start & " ms"
'  Debug.Print "ItemLoadedSubItems " & ExTvw1.ItemHandleToDisplayName(hItem)
  ExTvw1.MousePointer = MousePointerConstants.vbDefault
  DoEvents
End Sub

Private Sub ExTvw1_ItemLoadingSubItems(ByVal hItem As Long)
'  Debug.Print "ItemLoadingSubItems " & ExTvw1.ItemHandleToDisplayName(hItem)
  ExTvw1.MousePointer = MousePointerConstants.vbHourglass
  DoEvents
  Start = GetTickCount
End Sub

Private Sub ExTvw1_ItemMClick(ByVal hItem As Long)
'  Debug.Print "ItemMClick " & ExTvw1.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw1_ItemMDblClick(ByVal hItem As Long)
'  Debug.Print "ItemMDblClick " & ExTvw1.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw1_ItemMouseDown(ByVal hItem As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'  Debug.Print "ItemMouseDown " & ExTvw1.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw1_ItemMouseEnter(ByVal hItem As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'  Debug.Print "ItemMouseEnter " & ExTvw1.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw1_ItemMouseLeave(ByVal hItem As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'  Debug.Print "ItemMouseLeave " & ExTvw1.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw1_ItemMouseMove(ByVal hItem As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'  Debug.Print "ItemMouseMove " & ExTvw1.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw1_ItemMouseUp(ByVal hItem As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'  Debug.Print "ItemMouseUp " & ExTvw1.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw1_ItemRClick(ByVal hItem As Long)
'  Debug.Print "ItemRClick " & ExTvw1.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw1_ItemRDblClick(ByVal hItem As Long)
'  Debug.Print "ItemRDblClick " & ExTvw1.ItemHandleToDisplayName(hItem)
End Sub

Private Sub ExTvw1_ItemRenamed(ByVal hItem As Long, ByVal OldName As String, ByVal NewName As String)
'  Debug.Print "ItemRenamed " & ExTvw1.ItemHandleToDisplayName(hItem) & " from " & OldName & " to " & NewName
End Sub

Private Sub ExTvw1_ItemShellContextMenuItemEnter(ByVal hItem As Long, ByVal hMenu As Long, ByVal ItemID As Long, ByVal ItemText As String, ByVal ItemVerb As String, ByVal ItemDescription As String)
'  Debug.Print "ItemShellContextMenuItemEnter " & ExTvw1.ItemHandleToDisplayName(hItem) & " - Text: " & ItemText & ", Verb: " & ItemVerb & ", Description: " & ItemDescription
End Sub

Private Sub ExTvw1_ItemShellContextMenuItemLeave(ByVal hItem As Long, ByVal hMenu As Long, ByVal ItemID As Long, ByVal ItemText As String, ByVal ItemVerb As String, ByVal ItemDescription As String)
'  Debug.Print "ItemShellContextMenuItemLeave " & ExTvw1.ItemHandleToDisplayName(hItem) & " - Text: " & ItemText & ", Verb: " & ItemVerb & ", Description: " & ItemDescription
End Sub

Private Sub ExTvw1_ItemStateIconChanged(ByVal hItem As Long, ByVal OldStateIcon As Long, ByVal NewStateIcon As Long)
  Debug.Print "ItemStateIconChanged " & ExTvw1.ItemHandleToDisplayName(hItem) & " from " & OldStateIcon & " to " & NewStateIcon
End Sub

Private Sub ExTvw1_KeyDown(KeyCode As Integer, Shift As Integer)
'  Debug.Print "KeyDown"
End Sub

Private Sub ExTvw1_KeyPress(KeyAscii As Integer)
'  Debug.Print "KeyPress"
'  Command1.Caption = Chr(KeyAscii)
End Sub

Private Sub ExTvw1_KeyUp(KeyCode As Integer, Shift As Integer)
'  Debug.Print "KeyUp"
End Sub

Private Sub ExTvw1_LostFocus()
'  Debug.Print "LostFocus"
End Sub

Private Sub ExTvw1_MClick()
'  Debug.Print "MClick"
End Sub

Private Sub ExTvw1_MDblClick()
'  Debug.Print "MDblClick"
End Sub

Private Sub ExTvw1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'  Debug.Print "MouseDown"
End Sub

Private Sub ExTvw1_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'  Debug.Print "MouseEnter"
End Sub

Private Sub ExTvw1_MouseHover(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'  Debug.Print "MouseHover"
End Sub

Private Sub ExTvw1_MouseLeave(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'  Debug.Print "MouseLeave"
End Sub

Private Sub ExTvw1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'  Debug.Print "MouseMove"
End Sub

Private Sub ExTvw1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'  Debug.Print "MouseUp"
End Sub

Private Sub ExTvw1_RClick()
'  Debug.Print "RClick"
End Sub

Private Sub ExTvw1_RDblClick()
'  Debug.Print "RDblClick"
End Sub

Private Sub ExTvw1_SelChanged(ByVal hOldItem As Long, ByVal hNewItem As Long, ByVal CausedBy As ExTvw.CausedByConstants)
'  Debug.Print ExTvw1.ItemHandleToAddressBarText(hNewItem)
'  Debug.Print "SelChanged from " & ExTvw1.ItemHandleToDisplayName(hOldItem) & " to " & ExTvw1.ItemHandleToDisplayName(hNewItem)
'Dim txt As String
'
'  Select Case ExTvw1.GetItemType(hNewItem)
'    Case itemtypeconstants.itArchiveFile
'      txt = "ArchiveFile"
'    Case itemtypeconstants.itArchiveFolder
'      txt = "ArchiveFolder"
'    Case itemtypeconstants.itFSFile
'      txt = "FSFile"
'    Case itemtypeconstants.itFSFolder
'      txt = "FSFolder"
'    Case itemtypeconstants.itNonFSFile
'      txt = "NonFSFile"
'    Case itemtypeconstants.itNonFSFolder
'      txt = "NonFSFolder"
'  End Select
'  Debug.Print ExTvw1.ItemHandleToDisplayName(hNewItem) & ": " & txt
End Sub

Private Sub ExTvw1_SelChanging(ByVal hOldItem As Long, ByVal hNewItem As Long, ByVal CausedBy As ExTvw.CausedByConstants, Cancel As Boolean)
'  Debug.Print "SelChanging from " & ExTvw1.ItemHandleToDisplayName(hOldItem) & " to " & ExTvw1.ItemHandleToDisplayName(hNewItem)
End Sub

Private Sub Form_Click()
  Debug.Print ExTvw1.PathToItemHandle("C:\Windows", True)
End Sub

Private Sub Form_Initialize()
  Dim hMod As Long

  InitCommonControls

  hMod = LoadLibrary("uxtheme.dll")
  If hMod Then
    themeableOS = True
    FreeLibrary hMod
  End If
End Sub

Private Sub Form_Load()
  CLSIDFromString strIID_IDragDropHelper, IID_IDragDropHelper
  CLSIDFromString strIID_IDropTargetHelper, IID_IDropTargetHelper
  CoCreateInstance IID_IDragDropHelper, Nothing, CLSCTXConstants.CLSCTX_INPROC_SERVER, IID_IDropTargetHelper, IDropTargetHelper

  If themeableOS Then
    ' for Windows Vista
    SetWindowTheme ExTvw1.hWnd, StrPtr("explorer"), 0
  End If
  ExTvw1.hWndShellUIParentWindow = Me.hWnd
End Sub

Private Sub Form_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim pDO As Object
  Dim pt As POINT

  Effect = vbDropEffectNone
  pt.x = CLng(x)
  pt.y = CLng(y)
  CopyMemory pDO, ByVal ObjPtr(data) + 16, 4

  If Not (IDropTargetHelper Is Nothing) Then IDropTargetHelper.Drop ByVal pDO, pt, Effect

  ZeroMemory pDO, 4
End Sub

Private Sub Form_OLEDragOver(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
  Dim pDO As Object
  Dim pt As POINT

  Effect = vbDropEffectNone
  pt.x = CLng(x)
  pt.y = CLng(y)
  CopyMemory pDO, ByVal ObjPtr(data) + 16, 4

  Select Case State
    Case DragOverConstants.vbEnter
      If Not (IDropTargetHelper Is Nothing) Then IDropTargetHelper.DragEnter Me.hWnd, ByVal pDO, pt, Effect
    Case DragOverConstants.vbLeave
      If Not (IDropTargetHelper Is Nothing) Then IDropTargetHelper.DragLeave
    Case DragOverConstants.vbOver
      If Not (IDropTargetHelper Is Nothing) Then IDropTargetHelper.DragOver pt, Effect
  End Select

  ZeroMemory pDO, 4
End Sub

Private Sub Form_Terminate()
  If Not (IDropTargetHelper Is Nothing) Then IDropTargetHelper.Release
End Sub

Private Sub Timer1_Timer()
  Debug.Print ExTvw1.Path
End Sub


Private Function LoWord(ByVal Wert As Long) As Integer
  Dim ret As Integer

  CopyMemory ret, ByVal VarPtr(Wert), LenB(ret)

  LoWord = ret
End Function
