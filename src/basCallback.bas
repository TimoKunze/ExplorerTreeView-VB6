Attribute VB_Name = "basCallback"
Option Explicit

' öffentliche Methoden

' CallBack für das Sortieren von Archive-Items
' <lParam1> und <lParam2> sind Pointer auf "ITEMINFO_ARCHIVE"-Strukturen
' <lParam> ist ein Pointer auf das Objekt, welches den Sortiervorgang gestartet hat
Public Function Callback_Sort_Archive(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal lParam As Long) As Long
  Dim Obj As ISort

  Set Obj = GetObjectFromPtr(lParam)
  If Not (Obj Is Nothing) Then Callback_Sort_Archive = Obj.CBSortArchiveItems(lParam1, lParam2)
End Function

' CallBack für das Sortieren von "normalen" Items
' <lParam1> und <lParam2> sind Pointer auf "ITEMINFO_COMMON"-Strukturen
' <lParam> ist ein Pointer auf das Objekt, welches den Sortiervorgang gestartet hat
Public Function Callback_Sort_Common(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal lParam As Long) As Long
  Dim Obj As ISort

  Set Obj = GetObjectFromPtr(lParam)
  If Not (Obj Is Nothing) Then Callback_Sort_Common = Obj.CBSortCommonItems(lParam1, lParam2)
End Function

' CallBack für das Sortieren von Kategorien
' <lParam1> und <lParam2> sind die IDs
' <lParam> ist ein Pointer auf das Objekt, welches den Sortiervorgang gestartet hat
Public Function Callback_Sort_Groups(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal lParam As Long) As Long
  Dim Obj As ISort

  Set Obj = GetObjectFromPtr(lParam)
  If Not (Obj Is Nothing) Then Callback_Sort_Groups = Obj.CBSortGroups(lParam1, lParam2)
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
