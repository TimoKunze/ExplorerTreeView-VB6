VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommonDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

  Implements IHook


  ' öffentliche Enums

  Public Enum CDOpenFlagsConsts
    OFN_ALLOWMULTISELECT = &H200
    OFN_CREATEPROMPT = &H2000
    OFN_DONTADDTORECENT = &H2000000     ' ab Windows 2000
    OFN_ENABLEHOOK = &H20
    OFN_ENABLEINCLUDENOTIFY = &H400000
    OFN_ENABLESIZING = &H800000
    OFN_ENABLETEMPLATE = &H40
    OFN_ENABLETEMPLATEHANDLE = &H80
    OFN_EXPLORER = &H80000
    OFN_EXTENSIONDIFFERENT = &H400
    OFN_FILEMUSTEXIST = &H1000
    OFN_FORCESHOWHIDDEN = &H10000000     ' ab Windows 2000
    OFN_HIDEREADONLY = &H4
    OFN_LONGNAMES = &H200000
    OFN_NOCHANGEDIR = &H8
    OFN_NODEREFERENCELINKS = &H100000
    OFN_NOLONGNAMES = &H40000
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOREADONLYRETURN = &H8000
    OFN_NOTESTFILECREATE = &H10000
    OFN_NOVALIDATE = &H100
    OFN_OVERWRITEPROMPT = &H2
    OFN_PATHMUSTEXIST = &H800
    OFN_READONLY = &H1
    OFN_SHAREAWARE = &H4000
    OFN_SHAREFALLTHROUGH = 2
    OFN_SHARENOWARN = 1
    OFN_SHAREWARN = 0
    OFN_SHOWHELP = &H10
  End Enum

  Public Enum ExCDOpenFlagsConsts
    OFN_EX_NOPLACESBAR = &H1     ' ab Windows 2000
  End Enum


  ' lokale Konstanten

  Private Const MAX_PATH = 260

  ' Konstanten für SetWindowPos
  #If Debuging Then
    Private Const HWND_BOTTOM = 1
    Private Const HWND_NOTOPMOST = -2
    Private Const HWND_TOP = 0
    Private Const HWND_TOPMOST = -1
    Private Const SWP_ASYNCWINDOWPOS = &H4000
    Private Const SWP_DEFERERASE = &H2000
    Private Const SWP_FRAMECHANGED = &H20
    Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED
    Private Const SWP_HIDEWINDOW = &H80
    Private Const SWP_NOACTIVATE = &H10
    Private Const SWP_NOCOPYBITS = &H100
    Private Const SWP_NOMOVE = &H2
    Private Const SWP_NOOWNERZORDER = &H200
    Private Const SWP_NOREDRAW = &H8
    Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
    Private Const SWP_NOSENDCHANGING = &H400
    Private Const SWP_SHOWWINDOW = &H40
  #End If
  Private Const SWP_NOSIZE = &H1
  Private Const SWP_NOZORDER = &H4


  ' lokale Types
  Private Type CREATESTRUCT
    lpCreateParams As Long
    hInstance As Long
    hMenu As Long
    hWndParent As Long
    cy As Long
    cx As Long
    y As Long
    x As Long
    Style As Long
    lpszName As Long
    lpszClass As Long
    ExStyle As Long
  End Type

  Private Type CWPSTRUCT
    lParam As Long
    wParam As Long
    message As Long
    hWnd As Long
  End Type

  Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
    ' ab Windows 2000
    pvReserved As Long
    dwReserved As Long
    FlagsEx As Long
  End Type


  ' lokale Variablen

  Private colFilters As Collection
  Private HookStarted_CallWndProc As Boolean


  ' Property-Variablen

  Private propActiveFilter As Integer
  Private propCaption As String
  Private propDefaultExtension As String
  Private propDefaultFile As String
  Private propExOpenFlags As ExCDOpenFlagsConsts
  Private prophWndParent As Long
  Private propOpenFlags As CDOpenFlagsConsts
  Private propStartDirection As String


  ' lokale APIs

  Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal Size As Long)
  Private Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal buffer As String, ByVal bufferSize As Long) As Long
  Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (Data As OPENFILENAME) As Long
  Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (Data As OPENFILENAME) As Long
  Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal Flags As Long) As Long


' Implementation von IHook

Private Function IHook_CallWndProcAfter(ByVal HookCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  '
End Function

Private Function IHook_CallWndProcBefore(ByVal HookCode As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef eatIt As Boolean) As Long
  Const WM_CREATE = &H1
  Dim ClsName As String
  Dim CreationData As CREATESTRUCT
  Dim Data As CWPSTRUCT
  Dim x As Long
  Dim y As Long

  CopyMemory VarPtr(Data), lParam, LenB(Data)
  With Data
    Select Case .message
      Case WM_CREATE
        ClsName = Space(MAX_PATH)
        GetClassName .hWnd, ClsName, Len(ClsName)
        ClsName = removeNull(ClsName)
        If ClsName = "#32770" Then
          CopyMemory VarPtr(CreationData), .lParam, LenB(CreationData)
          With CreationData
            x = (Screen.Width / Screen.TwipsPerPixelX) / 2 - .cx / 2
            y = (Screen.Height / Screen.TwipsPerPixelY) / 2 - .cy / 2
          End With
          SetWindowPos .hWnd, 0, x, y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER
        End If
    End Select
  End With
End Function

Private Function IHook_CBTProcAfter(ByVal HookCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  '
End Function

Private Function IHook_CBTProcBefore(ByVal HookCode As Long, ByVal wParam As Long, ByVal lParam As Long, eatIt As Boolean) As Long
  '
End Function

Private Function IHook_MouseProcAfter(ByVal HookCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  '
End Function

Private Function IHook_MouseProcBefore(ByVal HookCode As Long, ByVal wParam As Long, ByVal lParam As Long, eatIt As Boolean) As Long
  '
End Function


' Events

Private Sub Class_Initialize()
  Set colFilters = New Collection
End Sub

Private Sub Class_Terminate()
  If HookStarted_CallWndProc Then
    removeCallWndProcHook Me
    HookStarted_CallWndProc = False
  End If

  Set colFilters = Nothing
End Sub


' öffentliche Props

Public Property Get ActiveFilter() As Integer
  ActiveFilter = propActiveFilter
End Property

Public Property Let ActiveFilter(ByVal New_ActiveFilter As Integer)
  propActiveFilter = New_ActiveFilter
End Property

Public Property Get Caption() As String
  Caption = propCaption
End Property

Public Property Let Caption(ByVal New_Caption As String)
  propCaption = New_Caption
End Property

Public Property Get DefaultExtension() As String
  DefaultExtension = propDefaultExtension
End Property

Public Property Let DefaultExtension(ByVal New_DefaultExtension As String)
  propDefaultExtension = New_DefaultExtension
End Property

Public Property Get DefaultFile() As String
  DefaultFile = propDefaultFile
End Property

Public Property Let DefaultFile(ByVal New_DefaultFile As String)
  propDefaultFile = New_DefaultFile
End Property

Public Property Get ExOpenFlags() As ExCDOpenFlagsConsts
  ExOpenFlags = propExOpenFlags
End Property

Public Property Let ExOpenFlags(ByVal New_ExOpenFlags As ExCDOpenFlagsConsts)
  propExOpenFlags = New_ExOpenFlags
End Property

Public Property Get FilterCount() As Integer
  FilterCount = colFilters.Count / 2
End Property

Public Property Get FilterDescription(ByVal Index As Integer) As String
  If Not isValidNumber(Index, 1, FilterCount / 2) Then
    Err.Raise 9
    Exit Property
  End If

  FilterDescription = colFilters(Index * 2 - 1)
End Property

Public Property Get FilterMask(ByVal Index As Integer) As String
  If Not isValidNumber(Index, 1, FilterCount / 2) Then
    Err.Raise 9
    Exit Property
  End If

  FilterMask = colFilters(Index * 2)
End Property

Public Property Get hWndParent() As Long
  hWndParent = prophWndParent
End Property

Public Property Let hWndParent(ByVal New_hWndParent As Long)
  prophWndParent = New_hWndParent
End Property

Public Property Get OpenFlags() As CDOpenFlagsConsts
  OpenFlags = propOpenFlags
End Property

Public Property Let OpenFlags(ByVal New_OpenFlags As CDOpenFlagsConsts)
  propOpenFlags = New_OpenFlags
End Property

Public Property Get StartDirection() As String
  StartDirection = propStartDirection
End Property

Public Property Let StartDirection(ByVal New_StartDirection As String)
  propStartDirection = New_StartDirection
End Property


' lokale Methoden

' gibt alle Filter in einem String zurück
' Format: <Beschreibung> & Chr(0) & <Maske> & Chr(0)...
' den Abschluss bilden 2 Nullen
Private Function getFilters() As String
  Dim Description As String
  Dim i As Integer
  Dim Mask As String
  Dim ret As String

  For i = 1 To colFilters.Count Step 2
    Description = colFilters(i)
    Mask = colFilters(i + 1)

    ret = ret & Description & Chr(0) & Mask & Chr(0)
  Next
  If ret <> "" Then ret = ret & Chr(0)

  getFilters = ret
End Function

' prüft, ob <txt> eine Zahl zwischen <Start> und <Ende> ist
Private Function isValidNumber(ByVal txt As String, ByVal Start As Long, ByVal Ende As Long) As Boolean
  Dim i As Integer
  Dim ret As Boolean
  Dim tmp As String

  If txt = "" Then Exit Function

  ret = True

  tmp = txt
  If Left(tmp, 1) = "-" Then tmp = Mid(tmp, 2)

  ' jedes Zeichen einzeln prüfen, ob es eine Ziffer ist
  For i = 1 To Len(tmp)
    If InStr("0123456789", Mid(tmp, i, 1)) = 0 Then ret = False
  Next
  If ret Then
    ' Grenzen prüfen
    If (CLng(txt) < Start) Or (CLng(txt) > Ende) Then ret = False
  End If

  isValidNumber = ret
End Function

' entfernt 2xChr(0) am Ende von <txt>
Private Function removeDoubleNull(ByVal txt As String) As String
  Dim pos As Integer

  pos = InStr(txt, String(2, Chr(0)))
  If pos Then txt = Left(txt, pos - 1)
  removeDoubleNull = txt
End Function

' entfernt Chr(0) am Ende von <txt>
Private Function removeNull(ByVal txt As String) As String
  Dim pos As Integer

  pos = InStr(txt, Chr(0))
  If pos Then txt = Left(txt, pos - 1)
  removeNull = txt
End Function


' öffentliche Methoden

' fügt einen Filter mit den angegebenen Eigenschaften hinzu
Public Sub AddFilter(ByVal Description As String, ByVal Mask As String, Optional ByVal After As Integer = -1)
  If After > 0 Then
    colFilters.Add Description, , , After * 2
    colFilters.Add Mask, , , After * 2 + 1
  Else
    colFilters.Add Description
    colFilters.Add Mask
  End If
End Sub

' löscht alle Filter
Public Sub RemoveAllFilters()
  Dim i As Integer

  For i = 1 To colFilters.Count Step 2
    colFilters.remove 1
    colFilters.remove 2
  Next
End Sub

' löscht den Filter <Index>
Public Sub RemoveFilter(ByVal Index As Integer)
  ' Maske entfernen
  colFilters.remove Index * 2
  ' Beschreibung entfernen
  colFilters.remove Index * 2 - 1
End Sub

' zeigt den "Datei öffnen"-Dialog an
Public Function showOpen(File As String) As Boolean
  Dim Data As OPENFILENAME

  If Not HookStarted_CallWndProc Then
    ' Hook starten
    installCallWndProcHook Me
    HookStarted_CallWndProc = True
  End If

  ' Eigenschaften übertragen
  With Data
    .Flags = propOpenFlags
    .FlagsEx = propExOpenFlags
    .hInstance = App.hInstance
    .hwndOwner = prophWndParent
    .lpstrDefExt = propDefaultExtension
    .lpstrFile = propDefaultFile & String(MAX_PATH - Len(propDefaultFile), Chr(0))
    .lpstrFileTitle = Space(MAX_PATH)
    .lpstrFilter = getFilters
    .lpstrInitialDir = propStartDirection
    .lpstrTitle = propCaption
    .nFilterIndex = propActiveFilter
    .nMaxFile = MAX_PATH
    .nMaxFileTitle = MAX_PATH
    .lStructSize = LenB(Data) - IIf(ver_Win_2k, 0, 12)
  End With

  ' anzeigen...
  showOpen = GetOpenFileName(Data)

  ' gewählte Datei zurückgeben
  File = removeDoubleNull(Data.lpstrFile)
End Function

' zeigt den "Datei speichern unter"-Dialog an
Public Function showSave(File As String) As Boolean
  Dim Data As OPENFILENAME

  If Not HookStarted_CallWndProc Then
    ' Hook starten
    installCallWndProcHook Me
    HookStarted_CallWndProc = True
  End If

  ' Eigenschaften übertragen
  With Data
    .Flags = propOpenFlags
    .FlagsEx = propExOpenFlags
    .hInstance = App.hInstance
    .hwndOwner = prophWndParent
    .lpstrDefExt = propDefaultExtension
    .lpstrFile = propDefaultFile & String(MAX_PATH - Len(propDefaultFile), Chr(0))
    .lpstrFileTitle = Space(MAX_PATH)
    .lpstrFilter = getFilters
    .lpstrInitialDir = propStartDirection
    .lpstrTitle = propCaption
    .nFilterIndex = propActiveFilter
    .nMaxFile = MAX_PATH
    .nMaxFileTitle = MAX_PATH
    .lStructSize = LenB(Data) - IIf(ver_Win_2k, 0, 12)
  End With

  ' anzeigen...
  showSave = GetSaveFileName(Data)

  ' gewählte Datei zurückgeben
  File = removeDoubleNull(Data.lpstrFile)
End Function
