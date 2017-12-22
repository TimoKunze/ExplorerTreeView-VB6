Attribute VB_Name = "basHooking"
Option Explicit

  ' lokale Konstanten

  ' Konstanten für SetWindowsHookEx
  #If Debuging Then
    Private Const HC_ACTION = 0
    Private Const HC_GETNEXT = 1
    Private Const HC_NOREMOVE = 3
    Private Const HC_NOREM = HC_NOREMOVE
    Private Const HC_SKIP = 2
    Private Const HC_SYSMODALOFF = 5
    Private Const HC_SYSMODALON = 4
    Private Const WH_CALLWNDPROCRET = 12
    Private Const WH_DEBUG = 9
    Private Const WH_FOREGROUNDIDLE = 11
    Private Const WH_GETMESSAGE = 3
    Private Const WH_HARDWARE = 8
    Private Const WH_JOURNALPLAYBACK = 1
    Private Const WH_JOURNALRECORD = 0
    Private Const WH_KEYBOARD = 2
    Private Const WH_KEYBOARD_LL = 13
    Private Const WH_MOUSE_LL = 14
    Private Const WH_MSGFILTER = -1
    Private Const WH_SHELL = 10
    Private Const WH_SYSMSGFILTER = 6

    Private Const WH_MAX = 14
    Private Const WH_MIN = -1
    Private Const WH_MAXHOOK = WH_MAX
    Private Const WH_MINHOOK = WH_MIN
  #End If
  Private Const WH_CALLWNDPROC = 4
  Private Const WH_CBT = 5
  Private Const WH_MOUSE = 7


  ' lokale Variablen

  Private hCallWndProcHook As Long
  Private hCBTHook As Long
  Private hMouseHook As Long
  Private pCallWndProcHookConsumers() As Long
  Private pCBTHookConsumers() As Long
  Private pMouseHookConsumers() As Long


  ' lokale APIs

  Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal HookCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
  Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal HookType As Long, ByVal pHookProc As Long, ByVal hinst As Long, ByVal ThreadID As Long) As Long
  Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
  Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByVal pDest As Long, ByVal Length As Long)


' öffentliche Methoden

Function CallWndProc(ByVal HookCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim eatIt As Boolean
  Dim i As Integer
  Dim Obj As IHook

  If HookCode >= 0 Then
    On Error GoTo Fehler
    For i = LBound(pCallWndProcHookConsumers) To UBound(pCallWndProcHookConsumers)
      If pCallWndProcHookConsumers(i) Then
        Set Obj = GetObjectFromPtr(pCallWndProcHookConsumers(i))
        Obj.CallWndProcBefore HookCode, wParam, lParam, eatIt
        If eatIt Then
          CallWndProc = 1
        Else
          CallWndProc = CallNextHookEx(hCallWndProcHook, HookCode, wParam, lParam)
          Obj.CallWndProcAfter HookCode, wParam, lParam
        End If
      End If
    Next
  End If
  Exit Function

Fehler:
  CallWndProc = CallNextHookEx(hCallWndProcHook, HookCode, wParam, lParam)
End Function

Function CBTProc(ByVal HookCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim eatIt As Boolean
  Dim i As Integer
  Dim Obj As IHook

  If HookCode >= 0 Then
    On Error GoTo Fehler
    For i = LBound(pCBTHookConsumers) To UBound(pCBTHookConsumers)
      If pCBTHookConsumers(i) Then
        Set Obj = GetObjectFromPtr(pCBTHookConsumers(i))
        Obj.CBTProcBefore HookCode, wParam, lParam, eatIt
        If eatIt Then
          CBTProc = 1
        Else
          CBTProc = CallNextHookEx(hCBTHook, HookCode, wParam, lParam)
          Obj.CBTProcAfter HookCode, wParam, lParam
        End If
      End If
    Next
  End If
  Exit Function

Fehler:
  CBTProc = CallNextHookEx(hCBTHook, HookCode, wParam, lParam)
End Function

Sub installCallWndProcHook(ByVal Obj As IHook)
  Dim exists As Boolean
  Dim i As Integer
  Dim iAvailableSlot As Integer
  Dim pObj As Long

  If hCallWndProcHook = 0 Then hCallWndProcHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf basHooking.CallWndProc, 0&, GetCurrentThreadId)
  If hCallWndProcHook Then
    iAvailableSlot = -1
    pObj = ObjPtr(Obj)
    On Error GoTo initArray
    i = LBound(pCallWndProcHookConsumers)
    For i = LBound(pCallWndProcHookConsumers) To UBound(pCallWndProcHookConsumers)
      If pCallWndProcHookConsumers(i) = pObj Then
        exists = True
      ElseIf pCallWndProcHookConsumers(i) = 0 Then
        iAvailableSlot = i
        Exit For
      End If
    Next

    If Not exists Then
      If iAvailableSlot = 0 Then
        ReDim Preserve pCallWndProcHookConsumers(UBound(pCallWndProcHookConsumers) + 1) As Long
        iAvailableSlot = UBound(pCallWndProcHookConsumers)
      End If
      pCallWndProcHookConsumers(iAvailableSlot) = pObj
    End If
  End If
  Exit Sub

initArray:
  ReDim pCallWndProcHookConsumers(0)
  Resume Next
End Sub

Sub installCBTHook(ByVal Obj As IHook)
  Dim exists As Boolean
  Dim i As Integer
  Dim iAvailableSlot As Integer
  Dim pObj As Long

  If hCBTHook = 0 Then hCBTHook = SetWindowsHookEx(WH_CBT, AddressOf basHooking.CBTProc, 0&, GetCurrentThreadId)
  If hCBTHook Then
    iAvailableSlot = -1
    pObj = ObjPtr(Obj)
    On Error GoTo initArray
    i = LBound(pCBTHookConsumers)
    For i = LBound(pCBTHookConsumers) To UBound(pCBTHookConsumers)
      If pCBTHookConsumers(i) = pObj Then
        exists = True
      ElseIf pCBTHookConsumers(i) = 0 Then
        iAvailableSlot = i
        Exit For
      End If
    Next

    If Not exists Then
      If iAvailableSlot = 0 Then
        ReDim Preserve pCBTHookConsumers(UBound(pCBTHookConsumers) + 1) As Long
        iAvailableSlot = UBound(pCBTHookConsumers)
      End If
      pCBTHookConsumers(iAvailableSlot) = pObj
    End If
  End If
  Exit Sub

initArray:
  ReDim pCBTHookConsumers(0)
  Resume Next
End Sub

Sub installMouseHook(ByVal Obj As IHook)
  Dim exists As Boolean
  Dim i As Integer
  Dim iAvailableSlot As Integer
  Dim pObj As Long

  If hMouseHook = 0 Then hMouseHook = SetWindowsHookEx(WH_MOUSE, AddressOf MouseProc, 0&, GetCurrentThreadId)
  If hMouseHook Then
    iAvailableSlot = -1
    pObj = ObjPtr(Obj)
    On Error GoTo initArray
    i = LBound(pMouseHookConsumers)
    For i = LBound(pMouseHookConsumers) To UBound(pMouseHookConsumers)
      If pMouseHookConsumers(i) = pObj Then
        exists = True
      ElseIf pMouseHookConsumers(i) = 0 Then
        iAvailableSlot = i
        Exit For
      End If
    Next

    If Not exists Then
      If iAvailableSlot = 0 Then
        ReDim Preserve pMouseHookConsumers(UBound(pMouseHookConsumers) + 1) As Long
        iAvailableSlot = UBound(pMouseHookConsumers)
      End If
      pMouseHookConsumers(iAvailableSlot) = pObj
    End If
  End If
  Exit Sub

initArray:
  ReDim pMouseHookConsumers(0)
  Resume Next
End Sub

Function MouseProc(ByVal HookCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim eatIt As Boolean
  Dim i As Integer
  Dim Obj As IHook

  If HookCode >= 0 Then
    On Error GoTo Fehler
    For i = LBound(pMouseHookConsumers) To UBound(pMouseHookConsumers)
      If pMouseHookConsumers(i) Then
        Set Obj = GetObjectFromPtr(pMouseHookConsumers(i))
        Obj.MouseProcBefore HookCode, wParam, lParam, eatIt
        If eatIt Then
          MouseProc = 1
        Else
          MouseProc = CallNextHookEx(hMouseHook, HookCode, wParam, lParam)
          Obj.MouseProcAfter HookCode, wParam, lParam
        End If
      End If
    Next
  End If
  Exit Function

Fehler:
  MouseProc = CallNextHookEx(hMouseHook, HookCode, wParam, lParam)
End Function

Sub removeCallWndProcHook(ByVal Obj As IHook)
  Dim countRefs As Integer
  Dim i As Integer
  Dim pObj As Long

  pObj = ObjPtr(Obj)
  On Error GoTo Ende
  For i = LBound(pCallWndProcHookConsumers) To UBound(pCallWndProcHookConsumers)
    If pCallWndProcHookConsumers(i) = pObj Then
      pCallWndProcHookConsumers(i) = 0
    ElseIf pCallWndProcHookConsumers(i) Then
      countRefs = countRefs + 1
    End If
  Next

  If countRefs = 0 Then
    UnhookWindowsHookEx hCallWndProcHook
    hCallWndProcHook = 0
  End If

Ende:
End Sub

Sub removeCBTHook(ByVal Obj As IHook)
  Dim countRefs As Integer
  Dim i As Integer
  Dim pObj As Long

  pObj = ObjPtr(Obj)
  On Error GoTo Ende
  For i = LBound(pCBTHookConsumers) To UBound(pCBTHookConsumers)
    If pCBTHookConsumers(i) = pObj Then
      pCBTHookConsumers(i) = 0
    ElseIf pCBTHookConsumers(i) Then
      countRefs = countRefs + 1
    End If
  Next

  If countRefs = 0 Then
    UnhookWindowsHookEx hCBTHook
    hCBTHook = 0
  End If

Ende:
End Sub

Sub removeMouseHook(ByVal Obj As IHook)
  Dim countRefs As Integer
  Dim i As Integer
  Dim pObj As Long

  pObj = ObjPtr(Obj)
  On Error GoTo Ende
  For i = LBound(pMouseHookConsumers) To UBound(pMouseHookConsumers)
    If pMouseHookConsumers(i) = pObj Then
      pMouseHookConsumers(i) = 0
    ElseIf pMouseHookConsumers(i) Then
      countRefs = countRefs + 1
    End If
  Next

  If countRefs = 0 Then
    UnhookWindowsHookEx hMouseHook
    hMouseHook = 0
  End If

Ende:
End Sub


' lokale Methoden

Private Function GetObjectFromPtr(ByVal Ptr As Long) As Object
  Dim ret As Object

  ' Pointer in Objekt konvertieren...
  CopyMemory VarPtr(ret), VarPtr(Ptr), 4
  ' ...und dieses zurückgeben
  Set GetObjectFromPtr = ret
  ' Speicher freigeben
  ZeroMemory VarPtr(ret), 4
End Function
