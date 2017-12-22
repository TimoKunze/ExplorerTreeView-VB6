Attribute VB_Name = "basSubClassing"
Option Explicit

  ' lokale Konstanten

  Private Const strControlID = "ExTvw "


  ' globale Konstanten

  ' Konstanten für SetWindowLongPtr
  #If Debuging Then
    Global Const DWL_DLGPROC = 4
    Global Const DWL_MSGRESULT = 0
    Global Const DWL_USER = 8
    Global Const DWLP_MSGRESULT = 0
    Global Const DWLP_DLGPROC = DWLP_MSGRESULT + 4 'Len(LRESULT)
    Global Const DWLP_USER = DWLP_DLGPROC + 4 'Len(DLGPROC)
    Global Const GWL_HINSTANCE = (-6)
    Global Const GWL_HWNDPARENT = (-8)
    Global Const GWL_ID = (-12)
    Global Const GWL_USERDATA = (-21)
    Global Const GWL_WNDPROC = (-4)
    Global Const GWLP_HINSTANCE = (-6)
    Global Const GWLP_HWNDPARENT = (-8)
    Global Const GWLP_ID = (-12)
    Global Const GWLP_USERDATA = (-21)
  #End If
  Global Const GWL_EXSTYLE = (-20)
  Global Const GWL_STYLE = (-16)
  Global Const GWLP_WNDPROC = (-4)


  ' lokale APIs

  Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal Prop As String) As Long
  Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal Prop As String) As Long
  Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal Prop As String, ByVal Data As Long) As Long


  ' globale APIs

  Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal Adresse As Long, ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Declare Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal Prop As Long, ByVal NewProp As Long) As Long


' öffentliche Methoden

Sub StartListening(ByVal Obj As ISubClass, ByVal hWnd As Long)
  Dim cListeners As Long
  Dim i As Long
  Dim oldWinMain As Long
  Dim tmp As Boolean

  ' <hWnd> subclassen
  cListeners = GetProp(hWnd, strControlID & "C_" & Hex(hWnd))
  If cListeners = 0 Then
    oldWinMain = SetWindowLongPtr(hWnd, GWLP_WNDPROC, AddressOf basSubClassing.WinMain)
    If oldWinMain Then
      If SetProp(hWnd, strControlID & "oldWinMain_" & Hex(hWnd), oldWinMain) = 0 Then
        SetWindowLongPtr hWnd, GWLP_WNDPROC, oldWinMain
        cListeners = 0
      End If
    Else
      cListeners = 0
    End If
  Else
    ' <hWnd> wurde bereits gesubclassed
  End If

  ' prüfen, ob <hWnd> schon von <Obj> überwacht wird
  For i = 1 To cListeners
    If GetProp(hWnd, strControlID & Hex(hWnd) & "_" & i) = ObjPtr(Obj) Then
      tmp = True
      Exit For
    End If
  Next

  ' <Obj> speichern
  If Not tmp Then
    cListeners = cListeners + 1
    SetProp hWnd, strControlID & "C_" & Hex(hWnd), cListeners
    ' <Obj> speichern
    SetProp hWnd, strControlID & Hex(hWnd) & "_" & cListeners, ObjPtr(Obj)
  End If
End Sub

Sub StopListening(ByVal Obj As ISubClass, ByVal hWnd As Long)
  Dim cListeners As Long
  Dim i As Long
  Dim oldWinMain As Long
  Dim pObj As Long
  Dim tmp As Long

  cListeners = GetProp(hWnd, strControlID & "C_" & Hex(hWnd))
  If cListeners = 1 Then
    ' wir müssen unsubclassen
    oldWinMain = GetProp(hWnd, strControlID & "oldWinMain_" & Hex(hWnd))
    SetWindowLongPtr hWnd, GWLP_WNDPROC, oldWinMain
    RemoveProp hWnd, strControlID & "oldWinMain_" & Hex(hWnd)
    RemoveProp hWnd, strControlID & Hex(hWnd) & "_1"
    RemoveProp hWnd, strControlID & "C_" & Hex(hWnd)
  Else
    For i = 1 To cListeners
      If GetProp(hWnd, strControlID & Hex(hWnd) & "_" & i) = ObjPtr(Obj) Then
        tmp = i
        Exit For
      End If
    Next

    If tmp Then
      For i = tmp + 1 To cListeners
        pObj = GetProp(hWnd, strControlID & Hex(hWnd) & "_" & i)
        SetProp hWnd, strControlID & Hex(hWnd) & "_" & (i - 1), pObj
      Next
      RemoveProp hWnd, strControlID & Hex(hWnd) & "_" & cListeners
      cListeners = cListeners - 1
      SetProp hWnd, strControlID & "C_" & Hex(hWnd), cListeners
    End If
  End If
End Sub

Function GetOldWinMain(ByVal hWnd As Long) As Long
  GetOldWinMain = GetProp(hWnd, strControlID & "oldWinMain_" & Hex(hWnd))
End Function

Function WinMain(ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim cListeners As Long
  Dim dontProcess As Boolean
  Dim eatMsg As Boolean
  Dim i As Long
  Dim oldWinMain As Long
  Dim Obj As ISubClass
  Dim pObj As Long
  Dim ret As Long

  On Error GoTo BigTrouble
  oldWinMain = GetProp(hWnd, strControlID & "oldWinMain_" & Hex(hWnd))
  If oldWinMain Then
    cListeners = GetProp(hWnd, strControlID & "C_" & Hex(hWnd))
    If cListeners Then
      For i = cListeners To 1 Step -1
        dontProcess = False
        pObj = GetProp(hWnd, strControlID & Hex(hWnd) & "_" & i)
        If pObj = 0 Then
          ' Fehler!
          ' <oldWinMain> aufrufen
          ret = CallWindowProc(oldWinMain, hWnd, MSG, wParam, lParam)
          dontProcess = True
        End If

        If Not dontProcess Then
          Set Obj = GetObjectFromPtr(pObj)

          If Not (Obj Is Nothing) Then
            ret = Obj.WinMain(i, oldWinMain, hWnd, MSG, wParam, lParam)
          Else
            ' Fehler!
            ' <oldWinMain> aufrufen
            ret = CallWindowProc(oldWinMain, hWnd, MSG, wParam, lParam)
          End If
        End If
      Next
    Else
      ret = CallWindowProc(oldWinMain, hWnd, MSG, wParam, lParam)
    End If
  End If

  WinMain = ret
BigTrouble:
End Function


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
