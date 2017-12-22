Attribute VB_Name = "basDragDrop"
Option Explicit

  ' öffentliche Enums

  Public Enum CLSCTXConsts
    CLSCTX_INPROC_SERVER = &H1
    CLSCTX_INPROC_HANDLER = &H2
    CLSCTX_LOCAL_SERVER = &H4
    CLSCTX_INPROC_SERVER16 = &H8
    CLSCTX_REMOTE_SERVER = &H10
    CLSCTX_INPROC_HANDLER16 = &H20
    CLSCTX_INPROC_SERVERX86 = &H40
    CLSCTX_INPROC_HANDLERX86 = &H80
    CLSCTX_ESERVER_HANDLER = &H100
    CLSCTX_RESERVED = &H200
    CLSCTX_NO_CODE_DOWNLOAD = &H400
    CLSCTX_NO_WX86_TRANSLATION = &H800
    CLSCTX_NO_CUSTOM_MARSHAL = &H1000
    CLSCTX_ENABLE_CODE_DOWNLOAD = &H2000
    CLSCTX_NO_FAILURE_LOG = &H4000
    CLSCTX_DISABLE_AAA = &H8000
    CLSCTX_ENABLE_AAA = &H10000
    CLSCTX_FROM_DEFAULT_CONTEXT = &H20000
  End Enum


  ' lokale Konstanten

  Private Const strIID_IDropSource = "{00000121-0000-0000-C000-000000000046}"


  ' globale Konstanten

  #If Debuging Then
    Global Const DRAGDROP_S_FIRST = &H40100
    Global Const DRAGDROP_S_LAST = &H4010F
  #End If
  Global Const DRAGDROP_S_CANCEL = &H40101
  Global Const DRAGDROP_S_DROP = &H40100
  Global Const DRAGDROP_S_USEDEFAULTCURSORS = &H40102


  ' globale Types

  Type IDSHookData
    ' Pointer auf VTable
    pVTable As Long
    ' das Original-Interface
    oldIDS As IVBDropSource
    ' Objekt, für das das Interface ersetzt wird
    'ObjCB As ExplorerComboBox
    'ObjLVW As ExplorerListView
    ObjTVW As ExplorerTreeView
    ' Typ von <Obj...>
    '    = AppID
    ObjType As Long
    ' Pointer auf diese Struktur
    pThis As Long
  End Type


  ' lokale Variablen

  Private IID_IDropSource As UUID


  ' Property-Variablen

  Private propIDSVTable(4) As Long


  ' globale APIs

  Declare Function CoCreateInstance Lib "ole32" (CLSID As UUID, ByVal IUnknwn As Any, ByVal Context As CLSCTXConsts, IID As UUID, Interface As Any) As Long
  Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef pDest As Any, ByRef pSrc As Any, ByVal sz As Long)


' öffentliche Props

Public Property Get IDSVTable() As Long
  ' VTable auf die einzelnen Prozeduren des Interfaces setzen
  propIDSVTable(0) = FuncPtr(AddressOf QueryInterface)
  propIDSVTable(1) = FuncPtr(AddressOf AddRef)
  propIDSVTable(2) = FuncPtr(AddressOf Release)
  propIDSVTable(3) = FuncPtr(AddressOf QueryContinueDrag)
  propIDSVTable(4) = FuncPtr(AddressOf GiveFeedback)

  CLSIDFromString StrPtr(strIID_IDropSource), IID_IDropSource
  ' Pointer auf VTable zurückgeben
  IDSVTable = VarPtr(propIDSVTable(0))
End Property


' Methoden des Interfaces

Private Function AddRef(This As IDSHookData) As Long
  ' Standard-Prozedur aufrufen
  AddRef = This.oldIDS.AddRef
End Function

Private Function GiveFeedback(This As IDSHookData, Effects As DROPEFFECTConstants) As Long
  Dim ret As Long

  Select Case This.ObjType
    Case 4     ' ComboBox
'      ret = This.ObjCB.IVBDropSource_GiveFeedback(Effects)
    Case 3     ' ListView
'      ret = This.ObjLVW.IVBDropSource_GiveFeedback(Effects)
    Case 2     ' TreeView
      ret = This.ObjTVW.IVBDropSource_GiveFeedback(Effects)
  End Select

  GiveFeedback = ret
End Function

Private Function QueryContinueDrag(This As IDSHookData, ByVal PressedEscape As Long, ByVal KeyState As Long) As Long
  Dim ret As Long

  Select Case This.ObjType
'    Case x     ' ComboBox
'      ret = This.ObjCB.IVBDropSource_QueryContinueDrag(PressedEscape, KeyState)
    Case 3     ' ListView
'      ret = This.ObjLVW.IVBDropSource_QueryContinueDrag(PressedEscape, KeyState)
    Case 2     ' TreeView
      ret = This.ObjTVW.IVBDropSource_QueryContinueDrag(PressedEscape, KeyState)
  End Select

  QueryContinueDrag = ret
End Function

Private Function QueryInterface(This As IDSHookData, IID As UUID, pObject As Long) As Long
  ' wenn das Interface IDropSource angefordert wird, unser eigenes Interface zurückgeben
  If isEqualGUID(IID, IID_IDropSource) Then
    pObject = This.pThis
    AddRef This
    QueryInterface = 0
  Else
    ' Standard-Prozedur aufrufen
    QueryInterface = This.oldIDS.QueryInterface(IID, pObject)
  End If
End Function

Private Function Release(This As IDSHookData) As Long
  ' Standard-Prozedur aufrufen
  Release = This.oldIDS.Release
End Function


' lokale Methoden

' gibt zurück, ob <IID1> und <IID2> gleich sind
Private Function isEqualGUID(IID1 As UUID, IID2 As UUID) As Boolean
  Dim Tmp1 As Currency
  Dim Tmp2 As Currency

  If IID1.Data1 = IID2.Data1 Then
    If IID1.Data2 = IID2.Data2 Then
      If IID1.Data3 = IID2.Data3 Then
        ' die letzten 8 Bytes mit einem Mal vergleichen
        CopyMemory ByVal VarPtr(Tmp1), ByVal VarPtr(IID1.Data4(0)), 8
        CopyMemory ByVal VarPtr(Tmp2), ByVal VarPtr(IID2.Data4(0)), 8
        isEqualGUID = (Tmp1 = Tmp2)
      End If
    End If
  End If
End Function
