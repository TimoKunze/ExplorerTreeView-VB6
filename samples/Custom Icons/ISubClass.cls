VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISubClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

  ' public enums

  Public Enum ProcessingOrderConstants
    poProcessAfterOriginal
    poProcessBeforeOriginal
    poProcessInsteadOfOriginal
    poProcessOriginalOnly
  End Enum


' public methods

Public Function GetAppIdentifier() As String
  '
End Function

Public Function GetProcessingOrder(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As ProcessingOrderConstants
  '
End Function

Public Function WinMain(ByVal iListener As Long, ByVal oldWinMain As Long, ByVal oldRetVal As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  '
End Function
