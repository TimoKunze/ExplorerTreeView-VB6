VERSION 5.00
Object = "{CAB249C7-BAC9-4C51-9526-12F29E40C4CE}#2.0#0"; "ExTvw.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "ExplorerTreeView-TechDemo"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   538
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   722
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CheckBox chkGhosted 
      Caption         =   "Ghosted"
      Height          =   255
      Left            =   5040
      TabIndex        =   3
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdAddFTP 
      Caption         =   "Add FTP path..."
      Height          =   495
      Left            =   7800
      TabIndex        =   14
      Top             =   240
      Width           =   2775
   End
   Begin VB.TextBox txtRoot 
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   6000
      Width           =   9615
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   255
      Index           =   0
      Left            =   10320
      TabIndex        =   8
      Top             =   6000
      Width           =   375
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   6360
      Width           =   9615
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   255
      Index           =   1
      Left            =   10320
      TabIndex        =   6
      Top             =   6360
      Width           =   375
   End
   Begin VB.Frame fraItemProps 
      Caption         =   "Selected item's properties"
      Height          =   5055
      Left            =   4920
      TabIndex        =   2
      Top             =   840
      Width           =   5775
      Begin VB.Label lblProps 
         Height          =   2295
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   5535
      End
   End
   Begin VB.CommandButton cmdPropPages 
      Caption         =   "Show property pages..."
      Height          =   495
      Left            =   4920
      TabIndex        =   1
      Top             =   240
      Width           =   2775
   End
   Begin ExTvw.ExplorerTreeView ExTvw 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   9975
      BackColor       =   16777215
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
      ForeColor       =   0
      Indent          =   19
      ItemHeight      =   17
      LineColor       =   10070188
      ShowFocusRect   =   0   'False
      ShownToolTips   =   3
   End
   Begin VB.Label lblAddressBarText 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   7680
      Width           =   45
   End
   Begin VB.Label lblTreePath 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   7320
      Width           =   45
   End
   Begin VB.Label lblFSPath 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   6960
      Width           =   45
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   8
      X2              =   712
      Y1              =   456
      Y2              =   456
   End
   Begin VB.Label lblDescr 
      AutoSize        =   -1  'True
      Caption         =   "Path:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   6450
      Width           =   390
   End
   Begin VB.Label lblDescr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Root:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   6090
      Width           =   390
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal pString As Long, CLSID As UUID) As Long
  Private Declare Function CoTaskMemFree Lib "ole32.dll" (ByVal cb As Long) As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
  Private Declare Function OleCreatePropertyFrame Lib "oleaut32.dll" (ByVal hWndOwner As Long, ByVal x As Long, ByVal y As Long, ByVal lpszCaption As Long, ByVal cObjects As Long, ByVal ppUnk As Long, ByVal cPages As Long, ByVal pPageClsID As Long, ByVal lcid As Long, ByVal dwReserved As Long, ByVal pvReserved As Long) As Long


Private Sub chkGhosted_Click()
  ExTvw.ItemGhosted(ExTvw.SelectedItem) = (chkGhosted.Value = vbChecked)
End Sub

Private Sub cmdAddFTP_Click()
  Dim ftp As String

  ftp = InputBox("Which FTP path shall be inserted?", "Add FTP Path", "ftp://myftp.com")
  ExTvw.AddFTPPath ftp
End Sub

Private Sub cmdBrowse_Click(Index As Integer)
  Dim frm As frmBrowseForFolder

  Set frm = New frmBrowseForFolder
  frm.showIt Index, IIf(Index, txtRoot.Text, "Desktop"), IIf(Index, txtPath.Text, txtRoot.Text), ExTvw.IncludedItems, ExTvw.FileAttributes, ExTvw.FolderAttributes, IIf(ExTvw.UseFileFilters, ExTvw.FileFilters, ""), IIf(ExTvw.UseFolderFilters, ExTvw.FolderFilters, ""), ExTvw.ExpandArchives
  If frm.Selected <> "" Then
    If Index = 1 Then
      txtPath.Text = frm.Selected
      ExTvw.Path = txtPath.Text
    Else
      txtRoot.Text = frm.Selected
      ExTvw.RootPath = txtRoot.Text
    End If
  End If
  Set frm = Nothing
End Sub

Private Sub cmdPropPages_Click()
  ShowProperties ExTvw, ExTvw.Name
End Sub

Private Sub ExTvw_ItemLoadedSubItems(ByVal hItem As Long)
  ExTvw.MousePointer = vbDefault
  MousePointer = vbDefault
End Sub

Private Sub ExTvw_ItemLoadingSubItems(ByVal hItem As Long)
  ExTvw.MousePointer = vbHourglass
  MousePointer = vbHourglass
End Sub

Private Sub ExTVW_SelChanged(ByVal hOldItem As Long, ByVal hNewItem As Long, ByVal CausedBy As ExTvw.CausedByConstants)
  Dim Target As String
  Dim txt As String

  chkGhosted.Value = IIf(ExTvw.ItemGhosted(hNewItem), vbChecked, vbUnchecked)
  txt = "IconIndex (selected/unselected): " & ExTvw.ItemIconIndex(hNewItem, True) & "/" & ExTvw.ItemIconIndex(hNewItem, False) & vbNewLine
  txt = txt & "StateIconIndex: " & ExTvw.ItemStateIconIndex(hNewItem) & vbNewLine
  txt = txt & "Level: " & ExTvw.ItemGetLevel(hNewItem) & vbNewLine
  Select Case ExTvw.ItemGetType(hNewItem)
    Case itArchiveFile
      txt = txt & "Type: file, part of an archive" & vbNewLine
    Case itArchiveFolder
      txt = txt & "Type: folder, part of an archive" & vbNewLine
    Case itFSFile
      txt = txt & "Type: file, part of the filesystem" & vbNewLine
    Case itFSFolder
      txt = txt & "Type: folder, part of the filesystem" & vbNewLine
    Case itNonFSFile
      txt = txt & "Type: ""file"", not part of the filesystem" & vbNewLine
    Case itNonFSFolder
      txt = txt & "Type: ""folder"", not part of the filesystem" & vbNewLine
  End Select
  txt = txt & "Supports ""new folder"" verb: " & IIf(ExTvw.ItemSupportsNewFolder(hNewItem), "yes", "no")
  If ExTvw.ItemIsLink(hNewItem, Target) Then
    txt = txt & vbNewLine & vbNewLine & "Item links to: " & Target
  End If
  lblProps.Caption = txt
  txtRoot.Text = ExTvw.RootPath
  txtPath.Text = ExTvw.Path
  lblFSPath.Caption = "Filesystem path: " & ExTvw.ItemHandleToFSPath(hNewItem)
  lblTreePath.Caption = "Tree path: " & ExTvw.ItemHandleToTreePath(hNewItem)
  lblAddressBarText.Caption = "Address Bar text: " & ExTvw.ItemHandleToAddressBarText(hNewItem)
End Sub

Private Sub Form_Initialize()
  ' support WinXP themes
  InitCommonControls
End Sub

Private Sub Form_Load()
  ExTvw.hWndShellUIParentWindow = Me.hWnd
End Sub


Private Sub ShowProperties(ByVal Ctrl As Object, Optional ByVal CtrlName As String, Optional ByVal hWndParent As Long = 0)
  Const strIID_ISpecifyPropertyPages = "{B196B28B-BAB4-101A-B69C-00AA00341D07}"
  Dim IID_ISpecifyPropertyPages As UUID
  Dim ISPP As IVBSpecifyPropertyPages
  Dim IUnk As IVBUnknown
  Dim propertyPages As CAUUID

  If hWndParent = 0 Then hWndParent = Me.hWnd
  CLSIDFromString StrPtr(strIID_ISpecifyPropertyPages), IID_ISpecifyPropertyPages

  Set IUnk = Ctrl.object
  IUnk.QueryInterface IID_ISpecifyPropertyPages, ISPP
  If Not (ISPP Is Nothing) Then
    ' get the CLSIDs of the pages
    ISPP.GetPages propertyPages
    ' show the pages
    OleCreatePropertyFrame hWndParent, 0, 0, StrPtr(CtrlName), 1, VarPtr(IUnk), propertyPages.cElems, propertyPages.pElems, 0, 0, 0
    ' release the pages array
    CoTaskMemFree propertyPages.pElems
  End If
End Sub
