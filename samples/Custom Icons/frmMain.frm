VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Object = "{CAB249C7-BAC9-4C51-9526-12F29E40C4CE}#2.0#0"; "ExTvw.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Custom icons with ExplorerTreeView"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   331
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   324
   StartUpPosition =   2  'Bildschirmmitte
   Begin vbalIml6.vbalImageList ImgLst 
      Left            =   0
      Top             =   4080
      _ExtentX        =   953
      _ExtentY        =   953
      ColourDepth     =   24
      Size            =   13776
      Images          =   "frmMain.frx":0000
      Version         =   131072
      KeyCount        =   12
      Keys            =   "ÿÿÿÿÿÿÿÿÿÿÿ"
   End
   Begin ExTvw.ExplorerTreeView ExTvw 
      Align           =   1  'Oben ausrichten
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4860
      _ExtentX        =   8573
      _ExtentY        =   8493
      ExpandArchives  =   8
      FileAttributes  =   63
      FolderAttributes=   63
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      LimitEditInput  =   0   'False
      UseCustomIcons  =   -1  'True
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Implements ISubClass


  Private Type TVITEMEX
    Mask As Long
    hItem As Long
    State As Long
    StateMask As Long
    pszText As Long
    cchTextMax As Long
    iImage As Long
    iSelectedImage As Long
    cChildren As Long
    lParam As Long
    iIntegral As Long
  End Type

  Private Type TVINSERTSTRUCT
    hParent As Long
    hInsertAfter As Long
    itemex As TVITEMEX
  End Type


  Private lastAssignedIcon As Long


  Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal Size As Long)
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()


Private Function ISubClass_GetAppIdentifier() As String
  ISubClass_GetAppIdentifier = "CustomIcons"
End Function

Private Function ISubClass_GetProcessingOrder(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As ProcessingOrderConstants
  Const TV_FIRST = &H1100
  Const TVM_INSERTITEM = (TV_FIRST + 0)
  Dim ret As ProcessingOrderConstants

  Select Case Msg
    Case TVM_INSERTITEM
      ret = ProcessingOrderConstants.poProcessBeforeOriginal
    Case Else
      ret = ProcessingOrderConstants.poProcessOriginalOnly
  End Select

  ISubClass_GetProcessingOrder = ret
End Function

Private Function ISubClass_WinMain(ByVal iListener As Long, ByVal oldWinMain As Long, ByVal oldRetVal As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Const TV_FIRST = &H1100
  Const TVIF_IMAGE = &H2
  Const TVIF_SELECTEDIMAGE = &H20
  Const TVM_INSERTITEM = (TV_FIRST + 0)
  Dim DispName As String
  Dim hItem_ArchiveFile As Long
  Dim InsertionData As TVINSERTSTRUCT
  Dim itmType As Byte
  Dim pIDLToDesktop As Long
  Dim ret As Long

  ret = oldRetVal
  Select Case Msg
    Case TVM_INSERTITEM
      ' an item is about to be inserted
      CopyMemory InsertionData, ByVal lParam, LenB(InsertionData)

      ' You probably want to set the icon dependent on the icon being inserted. This code block
      ' demonstrates how to do it.
      CopyMemory ByVal VarPtr(itmType), ByVal InsertionData.itemex.lParam, LenB(itmType)
      DispName = String$(InsertionData.itemex.cchTextMax + 1, Chr$(0))
      CopyMemory ByVal DispName, ByVal InsertionData.itemex.pszText, InsertionData.itemex.cchTextMax
      Select Case itmType
        Case 0
          ' filesystem folder
          CopyMemory pIDLToDesktop, ByVal InsertionData.itemex.lParam + 4, 4
        Case 1
          ' filesystem file
          CopyMemory pIDLToDesktop, ByVal InsertionData.itemex.lParam + 4, 4
        Case 2
          ' virtual folder (e. g. "My Computer")
          CopyMemory pIDLToDesktop, ByVal InsertionData.itemex.lParam + 4, 4
        Case 3
          ' virtual file (e. g. "Internet Explorer")
          CopyMemory pIDLToDesktop, ByVal InsertionData.itemex.lParam + 4, 4
        Case 4
          ' folder inside an archive
          CopyMemory hItem_ArchiveFile, ByVal InsertionData.itemex.lParam + 8, 4
        Case 5
          ' file inside an archive
          CopyMemory hItem_ArchiveFile, ByVal InsertionData.itemex.lParam + 8, 4
      End Select
      ' now you've:
      '   - the item's preceding item (InsertionData.hInsertAfter)
      '   - the item's parent item (InsertionData.hParent)
      '   - the item's text (DispName)
      '   - the item's state (InsertionData.itemex.State and StateMask)
      ' and (for shell items)
      '   - the item's fully qualified pIDL (pIDLToDesktop)
      ' or (for archive items)
      '   - the item representing the archive file, which contains the archive item (hItem_ArchiveFile)
      ' What you do using these details is up to you. :)

      ' customize the icons (it's just a demo, so we ignore the item details we've determined above)
      InsertionData.itemex.Mask = InsertionData.itemex.Mask Or TVIF_IMAGE Or TVIF_SELECTEDIMAGE
      InsertionData.itemex.iImage = lastAssignedIcon
      InsertionData.itemex.iSelectedImage = lastAssignedIcon
      lastAssignedIcon = (lastAssignedIcon + 1) Mod 12

      ' copy our changes back to the data that is sent to the SysTreeView32 window
      CopyMemory ByVal lParam, InsertionData, LenB(InsertionData)
  End Select

  ISubClass_WinMain = ret
End Function

Private Sub Form_Initialize()
  InitCommonControls
End Sub

Private Sub Form_Load()
  ExTvw.hWndShellUIParentWindow = Me.hWnd
  ExTvw.UseCustomIcons = True
  ExTvw.hImageList = ImgLst.hIml
  StartListening Me, ExTvw.hWnd
End Sub

Private Sub Form_Resize()
  ExTvw.Height = Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
  StopListening Me, ExTvw.hWnd
End Sub
