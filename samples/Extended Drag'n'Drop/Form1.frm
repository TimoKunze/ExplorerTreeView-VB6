VERSION 5.00
Object = "{CAB249C7-BAC9-4C51-9526-12F29E40C4CE}#2.0#0"; "ExTvw.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manuell
   ScaleHeight     =   346
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   600
   StartUpPosition =   2  'Bildschirmmitte
   Begin ComctlLib.ListView Lvw 
      Height          =   4935
      Left            =   3360
      TabIndex        =   1
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   8705
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      OLEDropMode     =   1
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Datei"
         Object.Width           =   26458
      EndProperty
   End
   Begin ExTvw.ExplorerTreeView ExTvw 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   8705
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
      LineColor       =   10070188
      Path            =   "E:\"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Private IDragSourceHelper As IVBDragSourceHelper
  Private IDropTargetHelper As IVBDropTargetHelper
  Private IID_IDragDropHelper As UUID
  Private IID_IDragSourceHelper As UUID
  Private IID_IDropTargetHelper As UUID

  Private Declare Function PathAddBackslash Lib "shlwapi.dll" Alias "PathAddBackslashA" (ByVal Path As String) As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()


Private Sub ExTVw_SelChanged(ByVal hOldItem As Long, ByVal hNewItem As Long, ByVal CausedBy As ExTvw.CausedByConstants)
  Dim strFile As String

  Lvw.ListItems.Clear
  ' fill listview with all mp3s in this folder
  With ExTvw
    If .ItemIsFSFolder(hNewItem) Then
      strFile = Dir(addBackslash(.ItemHandleToFSPath(hNewItem)) & "*.mp3")
      While strFile <> ""
        Lvw.ListItems.Add , , strFile
        strFile = Dir
      Wend
    End If
  End With
End Sub

Private Sub Form_Initialize()
  InitCommonControls
End Sub

Private Sub Form_Load()
  CLSIDFromString strIID_IDragDropHelper, IID_IDragDropHelper
  CLSIDFromString strIID_IDragSourceHelper, IID_IDragSourceHelper
  CLSIDFromString strIID_IDropTargetHelper, IID_IDropTargetHelper
  CoCreateInstance IID_IDragDropHelper, Nothing, CLSCTXConstants.CLSCTX_INPROC_SERVER, IID_IDragSourceHelper, IDragSourceHelper
  CoCreateInstance IID_IDragDropHelper, Nothing, CLSCTXConstants.CLSCTX_INPROC_SERVER, IID_IDropTargetHelper, IDropTargetHelper

  ExTvw.hWndShellUIParentWindow = Me.hWnd
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim pDO As Object
  Dim pt As POINT

  Effect = vbDropEffectNone
  pt.x = CLng(x)
  pt.y = CLng(y)
  CopyMemory pDO, ByVal ObjPtr(Data) + 16, 4

  If Not (IDropTargetHelper Is Nothing) Then IDropTargetHelper.Drop ByVal pDO, pt, Effect

  ZeroMemory pDO, 4
End Sub

Private Sub Form_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
  Dim pDO As Object
  Dim pt As POINT

  Effect = vbDropEffectNone
  pt.x = CLng(x)
  pt.y = CLng(y)
  CopyMemory pDO, ByVal ObjPtr(Data) + 16, 4

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
  If Not (IDragSourceHelper Is Nothing) Then IDragSourceHelper.Release
  If Not (IDropTargetHelper Is Nothing) Then IDropTargetHelper.Release
End Sub

Private Sub Lvw_OLEDragDrop(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim pDO As IVBDataObject
  Dim pt As POINT

  Effect = vbDropEffectNone
  pt.x = CLng(x)
  pt.y = CLng(y)
  Set pDO = Data

  If Not (IDropTargetHelper Is Nothing) Then IDropTargetHelper.Drop pDO, pt, Effect
End Sub

Private Sub Lvw_OLEDragOver(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
  Dim pDO As IVBDataObject
  Dim pt As POINT

  Effect = vbDropEffectNone
  pt.x = CLng(x)
  pt.y = CLng(y)
  Set pDO = Data

  Select Case State
    Case DragOverConstants.vbEnter
      If Not (IDropTargetHelper Is Nothing) Then IDropTargetHelper.DragEnter Lvw.hWnd, pDO, pt, Effect
    Case DragOverConstants.vbLeave
      If Not (IDropTargetHelper Is Nothing) Then IDropTargetHelper.DragLeave
    Case DragOverConstants.vbOver
      If Not (IDropTargetHelper Is Nothing) Then IDropTargetHelper.DragOver pt, Effect
  End Select
End Sub

Private Sub Lvw_OLEStartDrag(Data As ComctlLib.DataObject, AllowedEffects As Long)
  Dim lvItem As ListItem
  Dim Path As String
  Dim pt As POINT

  Path = addBackslash(ExTvw.ItemHandleToFSPath(ExTvw.SelectedItem))
  With Lvw
    For Each lvItem In .ListItems
      If lvItem.Selected Then
        Data.Files.Add Path & lvItem.Text
      End If
    Next
  End With
  Data.SetData , vbCFFiles
  AllowedEffects = ccOLEDropEffectCopy Or ccOLEDropEffectMove

  GetCursorPos pt
  ScreenToClient Lvw.hWnd, pt
  If Not (IDragSourceHelper Is Nothing) Then IDragSourceHelper.InitializeFromWindow Lvw.hWnd, pt, Data
End Sub


Private Function addBackslash(ByVal txt As String) As String
  txt = txt & String(2, Chr(0))
  PathAddBackslash txt
  addBackslash = Trim(removeNull(txt))
End Function

Private Function removeNull(ByVal txt As String) As String
  Dim pos As Integer

  pos = InStr(txt, Chr(0))
  If pos Then txt = Left(txt, pos - 1)
  removeNull = txt
End Function
