Attribute VB_Name = "basFokus"
Option Explicit

  ' lokale Konstanten

  Private Const strIID_IOleInPlaceActiveObject = "{00000117-0000-0000-C000-000000000046}"


  ' globale Konstanten

  Global Const S_FALSE = &H1


  ' globale Types

  Type IPAOHookData
    ' Pointer auf VTable
    pVTable As Long
    ' das Original-Interface
    oldIPAO As IVBOleInPlaceActiveObject
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

  Private IID_IOleInPlaceActiveObject As UUID


  ' Property-Variablen

  Private propIPAOVTable(9) As Long


' öffentliche Props

Public Property Get IPAOVTable() As Long
  ' VTable auf die einzelnen Prozeduren des Interfaces setzen
  propIPAOVTable(0) = FuncPtr(AddressOf QueryInterface)
  propIPAOVTable(1) = FuncPtr(AddressOf AddRef)
  propIPAOVTable(2) = FuncPtr(AddressOf Release)
  propIPAOVTable(3) = FuncPtr(AddressOf GetWindow)
  propIPAOVTable(4) = FuncPtr(AddressOf ContextSensitiveHelp)
  propIPAOVTable(5) = FuncPtr(AddressOf TranslateAccelerator)
  propIPAOVTable(6) = FuncPtr(AddressOf OnFrameWindowActivate)
  propIPAOVTable(7) = FuncPtr(AddressOf OnDocWindowActivate)
  propIPAOVTable(8) = FuncPtr(AddressOf ResizeBorder)
  propIPAOVTable(9) = FuncPtr(AddressOf EnableModeless)

  CLSIDFromString StrPtr(strIID_IOleInPlaceActiveObject), IID_IOleInPlaceActiveObject
  ' Pointer auf VTable zurückgeben
  IPAOVTable = VarPtr(propIPAOVTable(0))
End Property


' VTable-Methoden des Interfaces

Private Function AddRef(This As IPAOHookData) As Long
  ' Standard-Prozedur aufrufen
  AddRef = This.oldIPAO.AddRef
End Function

Private Function ContextSensitiveHelp(This As IPAOHookData, ByVal EnterMode As Long) As Long
  ' Standard-Prozedur aufrufen
  ContextSensitiveHelp = This.oldIPAO.ContextSensitiveHelp(EnterMode)
End Function

Private Function EnableModeless(This As IPAOHookData, ByVal Enable As Long) As Long
  ' Standard-Prozedur aufrufen
  EnableModeless = This.oldIPAO.EnableModeless(Enable)
End Function

Private Function GetWindow(This As IPAOHookData, hWnd As Long) As Long
  ' Standard-Prozedur aufrufen
  GetWindow = This.oldIPAO.GetWindow(hWnd)
End Function

Private Function OnDocWindowActivate(This As IPAOHookData, ByVal Activate As Long) As Long
  ' Standard-Prozedur aufrufen
  OnDocWindowActivate = This.oldIPAO.OnDocWindowActivate(Activate)
End Function

Private Function OnFrameWindowActivate(This As IPAOHookData, ByVal Activate As Long) As Long
  ' Standard-Prozedur aufrufen
  OnFrameWindowActivate = This.oldIPAO.OnFrameWindowActivate(Activate)
End Function

Private Function QueryInterface(This As IPAOHookData, IID As UUID, pObject As Long) As Long
  ' wenn das Interface IOleInPlaceActiveObject angefordert wird, unser eigenes Interface
  ' zurückgeben
  If isEqualGUID(IID, IID_IOleInPlaceActiveObject) Then
    pObject = This.pThis
    AddRef This
    QueryInterface = 0
  Else
    ' Standard-Prozedur aufrufen
    QueryInterface = This.oldIPAO.QueryInterface(IID, pObject)
  End If
End Function

Private Function Release(This As IPAOHookData) As Long
  ' Standard-Prozedur aufrufen
  Release = This.oldIPAO.Release
End Function

Private Function ResizeBorder(This As IPAOHookData, rcBorder As RECT, ByVal pUIWindow As IVBOleInPlaceUIWindow, ByVal FrameWindow As Long) As Long
  ' Standard-Prozedur aufrufen
  ResizeBorder = This.oldIPAO.ResizeBorder(rcBorder, pUIWindow, FrameWindow)
End Function

Private Function TranslateAccelerator(This As IPAOHookData, MsgData As MSG) As Long
  Dim ret As Long

  ' prüfen, ob diese Tastenanschläge abgefangen werden sollen
  Select Case This.ObjType
    Case 4     ' ComboBox
'      ret = This.ObjCB.TranslateAccelerator(MsgData)
    Case 3     ' ListView
'      ret = This.ObjLVW.TranslateAccelerator(MsgData)
    Case 2     ' TreeView
      ret = This.ObjTVW.TranslateAccelerator(MsgData)
  End Select

  If ret = S_FALSE Then
    ' Standard-Prozedur aufrufen
    ret = This.oldIPAO.TranslateAccelerator(MsgData)
  End If
  TranslateAccelerator = ret
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
        CopyMemory VarPtr(Tmp1), VarPtr(IID1.Data4(0)), 8
        CopyMemory VarPtr(Tmp2), VarPtr(IID2.Data4(0)), 8
        isEqualGUID = (Tmp1 = Tmp2)
      End If
    End If
  End If
End Function
