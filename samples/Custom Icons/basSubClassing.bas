Attribute VB_Name = "basSubClassing"
Option Explicit

  ' global constants

  ' constants used with SetWindowLongPtr()
  #If Debuging Then
    Global Const DWL_DLGPROC = 4
    Global Const DWL_MSGRESULT = 0
    Global Const DWL_USER = 8
    Global Const DWLP_MSGRESULT = 0
    Global Const DWLP_DLGPROC = DWLP_MSGRESULT + 4 'LenB(LRESULT)
    Global Const DWLP_USER = DWLP_DLGPROC + 4 'LenB(DLGPROC)
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


  ' local vars

  Private strControlID As String


  ' local API-declarations

  Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal Size As Long)
  Private Declare Function GetProp Lib "user32.dll" Alias "GetPropA" (ByVal hWnd As Long, ByVal Prop As String) As Long
  Private Declare Function RemoveProp Lib "user32.dll" Alias "RemovePropA" (ByVal hWnd As Long, ByVal Prop As String) As Long
  Private Declare Function SetProp Lib "user32.dll" Alias "SetPropA" (ByVal hWnd As Long, ByVal Prop As String, ByVal Data As Long) As Long
  Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef pDest As Any, ByVal Length As Long)


  ' global API-declarations

  Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal Adresse As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Declare Function SetWindowLongPtr Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal Prop As Long, ByVal NewProp As Long) As Long


' public methods

' adds <obj> to the list of subclassers of <hWnd>
Sub StartListening(ByVal obj As ISubClass, ByVal hWnd As Long)
  Dim cListeners As Long
  Dim i As Long
  Dim oldWinMain As Long
  Dim tmp As Boolean

  If obj Is Nothing Then Exit Sub
  If strControlID = "" Then strControlID = obj.GetAppIdentifier

  ' check whether we've already subclassed <hWnd>
  cListeners = GetProp(hWnd, strControlID & " C_" & Hex(hWnd))
  If cListeners = 0 Then
    ' no, we didn't, so do it now
    oldWinMain = SetWindowLongPtr(hWnd, GWLP_WNDPROC, AddressOf basSubClassing.WinMain)
    If oldWinMain Then
      ' save the original WndProc
      If SetProp(hWnd, strControlID & " oldWinMain_" & Hex(hWnd), oldWinMain) = 0 Then
        ' an error occured, undo all the things we did so far
        SetWindowLongPtr hWnd, GWLP_WNDPROC, oldWinMain
        cListeners = 0
      End If
    Else
      ' an error occured
      cListeners = 0
    End If
  Else
    ' yes, we've already subclassed it
  End If

  ' check whether <obj> is already a subclasser of <hWnd>
  For i = 1 To cListeners
    If GetProp(hWnd, strControlID & " " & Hex(hWnd) & "_" & i) = ObjPtr(obj) Then
      ' yes, it is
      tmp = True
      Exit For
    End If
  Next

  If Not tmp Then
    ' save the total number of subclassers
    cListeners = cListeners + 1
    SetProp hWnd, strControlID & " C_" & Hex(hWnd), cListeners
    ' save a pointer to <obj>
    SetProp hWnd, strControlID & " " & Hex(hWnd) & "_" & cListeners, ObjPtr(obj)
  End If
End Sub

' removes <obj> from the list of subclassers of <hWnd>
Sub StopListening(ByVal obj As ISubClass, ByVal hWnd As Long)
  Dim cListeners As Long
  Dim i As Long
  Dim oldWinMain As Long
  Dim pObj As Long
  Dim tmp As Long

  ' get the number of all subclassers
  cListeners = GetProp(hWnd, strControlID & " C_" & Hex(hWnd))
  If cListeners = 1 Then
    ' we'll remove the last subclasser, so restore the original WndProc
    oldWinMain = GetProp(hWnd, strControlID & " oldWinMain_" & Hex(hWnd))
    SetWindowLongPtr hWnd, GWLP_WNDPROC, oldWinMain
    ' we've to clean up some things
    RemoveProp hWnd, strControlID & " oldWinMain_" & Hex(hWnd)
    RemoveProp hWnd, strControlID & " " & Hex(hWnd) & "_1"
    RemoveProp hWnd, strControlID & " C_" & Hex(hWnd)
  Else
    ' search the list of subclassers for <obj>
    For i = 1 To cListeners
      If GetProp(hWnd, strControlID & " " & Hex(hWnd) & "_" & i) = ObjPtr(obj) Then
        tmp = i
        Exit For
      End If
    Next

    If tmp Then
      For i = tmp + 1 To cListeners
        ' since we'll remove an item from the list, we'll need to update the following ones
        pObj = GetProp(hWnd, strControlID & " " & Hex(hWnd) & "_" & i)
        SetProp hWnd, strControlID & " " & Hex(hWnd) & "_" & (i - 1), pObj
      Next
      ' remove the item from the list
      RemoveProp hWnd, strControlID & " " & Hex(hWnd) & "_" & cListeners
      ' save the total number of subclassers
      cListeners = cListeners - 1
      SetProp hWnd, strControlID & " C_" & Hex(hWnd), cListeners
    End If
  End If
End Sub

' returns the original WndProc
Function getOldWinMain(ByVal hWnd As Long) As Long
  getOldWinMain = GetProp(hWnd, strControlID & " oldWinMain_" & Hex(hWnd))
End Function

' this is *our* WinMain for <hWnd>
' we'll need to spread the messages to all subclassers
Function WinMain(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim calledOriginal As Boolean
  Dim cListeners As Long
  Dim i As Long
  Dim oldWinMain As Long
  Dim obj As ISubClass
  Dim pObj As Long
  Dim ProcessingOrder As ProcessingOrderConstants
  Dim ret As Long

  ' check whether we really subclassed <hWnd>
  oldWinMain = GetProp(hWnd, strControlID & " oldWinMain_" & Hex(hWnd))
  If oldWinMain Then
    ' we did, now provide the message to every single subclasser
    cListeners = GetProp(hWnd, strControlID & " C_" & Hex(hWnd))
    ' the last subclasser should be called at first
    For i = cListeners To 1 Step -1
      pObj = GetProp(hWnd, strControlID & " " & Hex(hWnd) & "_" & i)
      Set obj = getObjectFromPtr(pObj)

      If Not (obj Is Nothing) Then
        ' ask the subclasser when this message should be provided to him
        ProcessingOrder = obj.GetProcessingOrder(hWnd, Msg, wParam, lParam)
        If i = cListeners Then
          Select Case ProcessingOrder
            Case ProcessingOrderConstants.poProcessAfterOriginal, ProcessingOrderConstants.poProcessOriginalOnly
              If Not calledOriginal Then
                ret = CallWindowProc(oldWinMain, hWnd, Msg, wParam, lParam)
                calledOriginal = True
              End If
          End Select
        End If

        ' let the subclasser process this message (if it wants it)
        If ProcessingOrder <> ProcessingOrderConstants.poProcessOriginalOnly Then ret = obj.WinMain(i, oldWinMain, ret, hWnd, Msg, wParam, lParam)
        If Not calledOriginal Then
          If ProcessingOrder = ProcessingOrderConstants.poProcessInsteadOfOriginal Then calledOriginal = True
        End If

        If i = 1 Then
          Select Case ProcessingOrder
            Case ProcessingOrderConstants.poProcessBeforeOriginal
              If Not calledOriginal Then
                ret = CallWindowProc(oldWinMain, hWnd, Msg, wParam, lParam)
                calledOriginal = True
              End If
          End Select
        End If
      End If
    Next

    ' call the original WndProc, if we didn't it yet
    If Not calledOriginal Then ret = CallWindowProc(oldWinMain, hWnd, Msg, wParam, lParam)
  End If

  WinMain = ret
End Function


' local methods

' returns the object <Ptr> points to
Private Function getObjectFromPtr(ByVal Ptr As Long) As Object
  Dim ret As Object

  ' get the object <Ptr> points to
  CopyMemory ret, Ptr, LenB(Ptr)
  Set getObjectFromPtr = ret
  ' free memory
  ZeroMemory ret, 4
End Function
