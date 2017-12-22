VERSION 5.00
Begin VB.UserControl ExplorerTreeView 
   Alignable       =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   4410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3630
   EditAtDesignTime=   -1  'True
   HasDC           =   0   'False
   Picture         =   "ExTreeView.ctx":0000
   PropertyPages   =   "ExTreeView.ctx":058A
   ScaleHeight     =   294
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   242
   ToolboxBitmap   =   "ExTreeView.ctx":060E
End
Attribute VB_Name = "ExplorerTreeView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

  Implements IHook
  Implements ISort
  Implements ISubClass


  ' öffentliche Enums

  Public Enum AppearanceConstants
    a2D
    a3D
    a3DLight
  End Enum

  Public Enum AttributesConstants
    attReadOnly = 1
    attHidden = 2
    attArchive = 4
    attSystem = 8
    attEncrypted = 16
    attCompressed = 32
  End Enum

  Public Enum BkImageStyleConstants
    bisTiled
    #If Debuging Then
      bisBottomLeft
      bisBottomRight
      bisCentered
      bisTopLeft
      bisTopRight
      bisStretched
    #End If
  End Enum

  Public Enum BorderStyleConstants
    bsNone
    bsFixedSingle
  End Enum

  Public Enum CausedByConstants
    cbUnknown
    cbMouse
    cbKeyboard
    cbInternal
  End Enum

  Public Enum ErrorTypeConstants
    etNone
    etFileNotFound
    etCanNotOpenFile
    etCanNotReadFile
    etInvalidFormat
    etTooNewFileFormatToRead
    etNoSettingsFound
    etUnknown
    etCanNotWriteFile
  End Enum

  Public Enum ExpandArchivesConstants
    eaNone = 0
    eaACE = 1
    eaCAB = 2
    eaJAR = 4
    eaRAR = 8
    eaZIP = 16
    eaISO = 32
    eaBIN = 64
    #If NewArchiveSupport Then
      eaTAR = 128
    #End If
  End Enum

  Public Enum FireBeforeInsertItemConstants
    fbiiForFSFolders = 1
    fbiiForFSFiles = 2
    fbiiForNonFSFolders = 4
    fbiiForNonFSFiles = 8
  End Enum

  Public Enum IncludedItemsConstants
    iiFSFolders = 1
    iiFSFiles = 2
    iiNonFSFolders = 4
    iiNonFSFiles = 8
  End Enum

  Public Enum InfoTipStyleConstants
    itsDefault = QITipFlags.QITIPF_DEFAULT
    itsNoLinkTargets = QITipFlags.QITIPF_LINKNOTARGET
    itsUseLinkTargetPath = QITipFlags.QITIPF_LINKUSETARGET
    itsUseLinkNameAndTarget = QITipFlags.QITIPF_USENAME
    itsShowSlowTip = QITipFlags.QITIPF_USESLOWTIP
  End Enum

  Public Enum ItemTypeConstants
    itFSFolder          ' Ordner
    itFSFile            ' Datei
    itNonFSFolder       ' Objekt-"Ordner"
    itNonFSFile         ' Objekt-"Datei"
    itArchiveFolder     ' Ordner in einem Archiv
    itArchiveFile       ' Datei in einem Archiv
  End Enum

  Public Enum LineStyleConstants
    lsLinesAtRoot
    lsLinesAtItem
  End Enum

  Public Enum pIDLTypeConstants
    ptToDesktop
    ptToParent
    ptToRoot
  End Enum

  Public Enum RenamingConstants
    rAutomatic
    rManual
  End Enum

  Public Enum ScrollBarsConstants
    sbNone
    sbNormal
'    sbFlat
  End Enum

  Public Enum ShellContextMenuStyleConstants
    scmsDefaultOnly = CMFConstants.CMF_DEFAULTONLY
    scmsExplorerTree = CMFConstants.CMF_EXPLORE
    scmsExtendedVerbs = CMFConstants.CMF_EXTENDEDVERBS
    scmsFindHack = CMFConstants.CMF_ITEMMENU
    scmsNoDefault = CMFConstants.CMF_NODEFAULT
    scmsNormal = CMFConstants.CMF_NORMAL
    scmsNoVerbs = CMFConstants.CMF_NOVERBS
    scmsVerbsOnly = CMFConstants.CMF_VERBSONLY
  End Enum

  Public Enum ShownOverlaysConstants
    soNone = 0
    soLink = 1
    soSharedItem = 2
    soSlowFile = 4
    soOthers = 8

    soAll = soLink Or soSharedItem Or soSlowFile Or soOthers
  End Enum

  Public Enum ShownToolTipsConstants
    sttNone = 0
    sttHiddenItems = 1
    sttInfoTips = 2
  End Enum

  Public Enum SingleExpandConstants
    seNone = 0
    seNormal = 1
    seWinXPStyle = 2
  End Enum

  Public Enum SortingCriteriaConstants
    scName = 0
    scSize = 1
    scType = 2
    scLastModified = 3
  End Enum

  Public Enum SortOrderConstants
    sorAscending
    sorDescending
  End Enum

  Public Enum ToolTipTypeConstants
    tttInfoTip = 0
    tttToolTip = 1
  End Enum

  Public Enum TreeViewStyleConstants
    tvsButtons = 1
    tvsLines = 2
  End Enum


  ' lokale Enums

  Private Enum AllowNullConstants
    anAlways = 0
    anNever = 1
    anIfEmpty = 2
    anIfNoShowRoot = 4
  End Enum

  Private Enum ExWindowStyleConsts
    WS_EX_ACCEPTFILES = &H10
    WS_EX_APPWINDOW = &H40000
    WS_EX_CLIENTEDGE = &H200
    WS_EX_COMPOSITED = &H2000000
    WS_EX_CONTEXTHELP = &H400
    WS_EX_CONTROLPARENT = &H10000
    WS_EX_DLGMODALFRAME = &H1
    WS_EX_LAYERED = &H80000
    WS_EX_LAYOUTRTL = &H400000
    WS_EX_LEFT = &H0
    WS_EX_LEFTSCROLLBAR = &H4000
    WS_EX_LTRREADING = &H0
    WS_EX_MDICHILD = &H40
    WS_EX_NOACTIVATE = &H8000000
    WS_EX_NOINHERITLAYOUT = &H100000
    WS_EX_NOPARENTNOTIFY = &H4
    WS_EX_RIGHT = &H1000
    WS_EX_RIGHTSCROLLBAR = &H0
    WS_EX_RTLREADING = &H2000
    WS_EX_STATICEDGE = &H20000
    WS_EX_TOOLWINDOW = &H80
    WS_EX_TOPMOST = &H8
    WS_EX_TRANSPARENT = &H20
    WS_EX_WINDOWEDGE = &H100
    WS_EX_OVERLAPPEDWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE)
    WS_EX_PALETTEWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)
  End Enum

  Private Enum TreeViewStyleConsts
    TVS_HASBUTTONS = &H1
    TVS_HASLINES = &H2
    TVS_LINESATROOT = &H4
    TVS_EDITLABELS = &H8
    TVS_DISABLEDRAGDROP = &H10
    TVS_SHOWSELALWAYS = &H20
    TVS_RTLREADING = &H40
    TVS_NOTOOLTIPS = &H80
    TVS_CHECKBOXES = &H100
    TVS_TRACKSELECT = &H200
    TVS_SINGLEEXPAND = &H400
    TVS_INFOTIP = &H800
    TVS_FULLROWSELECT = &H1000
    TVS_NOSCROLL = &H2000
    TVS_NONEVENHEIGHT = &H4000
    TVS_NOHSCROLL = &H8000
  End Enum

  Private Enum TreeViewExStyleConsts
    TVS_EX_MULTISELECT = &H2
    TVS_EX_DOUBLEBUFFER = &H4
    TVS_EX_NOINDENTSTATE = &H8
    TVS_EX_RICHTOOLTIP = &H10
    TVS_EX_AUTOHSCROLL = &H20
    TVS_EX_FADEINOUTEXPANDOS = &H40
    TVS_EX_PARTIALCHECKBOXES = &H80
    TVS_EX_EXCLUSIONCHECKBOXES = &H100
    TVS_EX_DIMMEDCHECKBOXES = &H200
    TVS_EX_DRAWIMAGEASYNC = &H400
  End Enum

  Private Enum WindowStyleConsts
    WS_BORDER = &H800000
    WS_CAPTION = &HC00000
    WS_CHILD = &H40000000
    WS_CLIPCHILDREN = &H2000000
    WS_CLIPSIBLINGS = &H4000000
    WS_DISABLED = &H8000000
    WS_DLGFRAME = &H400000
    WS_GROUP = &H20000
    WS_HSCROLL = &H100000
    WS_MAXIMIZE = &H1000000
    WS_MAXIMIZEBOX = &H10000
    WS_MINIMIZE = &H20000000
    WS_MINIMIZEBOX = &H20000
    WS_OVERLAPPED = &H0
    WS_POPUP = &H80000000
    WS_SYSMENU = &H80000
    WS_TABSTOP = &H10000
    WS_THICKFRAME = &H40000
    WS_VISIBLE = &H10000000
    WS_VSCROLL = &H200000
    WS_CHILDWINDOW = WS_CHILD
    WS_ICONIC = WS_MINIMIZE
    WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
    WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
    WS_SIZEBOX = WS_THICKFRAME
    WS_TILED = WS_OVERLAPPED
    WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
  End Enum


  ' lokale Konstanten

  Private Const ArchiveSignature = "«|ARCHIV|»"
  Private Const CLR_DEFAULT = &HFF000000
  Private Const CLR_HILIGHT = CLR_DEFAULT
  Private Const CLR_NONE = &HFFFFFFFF
  Private Const FileCreatedBy_SettingsFile = "extvw"
  Private Const FileTypeSignature_SettingsFile = "exctlsf"
  Private Const FileTypeVersion_SettingsFile = "2.00"
  Private Const NeededVersion_SettingsFile_Read_Build_ExCmb = 0
  Private Const NeededVersion_SettingsFile_Read_Major_ExCmb = 0
  Private Const NeededVersion_SettingsFile_Read_Minor_ExCmb = 0
  Private Const NeededVersion_SettingsFile_Read_Revision_ExCmb = 0
  Private Const NeededVersion_SettingsFile_Read_Build_ExLvw = 88
  Private Const NeededVersion_SettingsFile_Read_Major_ExLvw = 0
  Private Const NeededVersion_SettingsFile_Read_Minor_ExLvw = 0
  Private Const NeededVersion_SettingsFile_Read_Revision_ExLvw = 88
  Private Const NeededVersion_SettingsFile_Read_Build_ExTvw = 437
  Private Const NeededVersion_SettingsFile_Read_Major_ExTvw = 1
  Private Const NeededVersion_SettingsFile_Read_Minor_ExTvw = 2
  Private Const NeededVersion_SettingsFile_Read_Revision_ExTvw = 2
  Private Const NeededVersion_SettingsFile_Write_Build_ExCmb = 0
  Private Const NeededVersion_SettingsFile_Write_Major_ExCmb = 0
  Private Const NeededVersion_SettingsFile_Write_Minor_ExCmb = 0
  Private Const NeededVersion_SettingsFile_Write_Revision_ExCmb = 0
  Private Const NeededVersion_SettingsFile_Write_Build_ExLvw = 88
  Private Const NeededVersion_SettingsFile_Write_Major_ExLvw = 0
  Private Const NeededVersion_SettingsFile_Write_Minor_ExLvw = 0
  Private Const NeededVersion_SettingsFile_Write_Revision_ExLvw = 88
  Private Const NeededVersion_SettingsFile_Write_Build_ExTvw = 437
  Private Const NeededVersion_SettingsFile_Write_Major_ExTvw = 1
  Private Const NeededVersion_SettingsFile_Write_Minor_ExTvw = 2
  Private Const NeededVersion_SettingsFile_Write_Revision_ExTvw = 2
  Private Const NM_FIRST = 0
  #If Debuging Then
    Private Const I_IMAGENONE = -2
  #End If
  Private Const I_CHILDRENCALLBACK = -1
  Private Const I_IMAGECALLBACK = -1
  Private Const TimerID_SubClassEdit = 1
  Private Const TimerID_SubClassParent = 2
  Private Const TimerID_MouseEvents = 10
  Private Const TimerID_DragScroll = 20
  Private Const TimerID_StateIconChange = 21
  Private Const TimerID_ResetIconCacheRebuildFlag = 22
  Private Const TimerID_DragExpand = 30
  Private Const TV_FIRST = &H1100
  Private Const TVN_FIRST = -400
  Private Const WM_USER = &H400     ' wird in basShell nochmal deklariert, liegt irgendwie an Binary Compatibility
  Private Const EXTVM_MYMOUSEHOVER = (WM_USER + 10)
  Private Const EXTVM_MYMOUSELEAVE = (WM_USER + 11)
  Private Const EXTVM_SETAUTOUPDATE = (WM_USER + 20)     ' wParam = 0/1, Rückgabewert = alte Einstellung
  Private Const EXTVM_SETROOTPATH = (WM_USER + 21)     ' wParam = buffer size, lParam = pointer to (ANSI) string buffer

  ' Klassenname für TreeView
  Private Const WC_TREEVIEW = "SysTreeView32"

  ' Fehlerkonstanten
  Private Const errCreateFileError = 100
  Private Const errNoSettingsFound = 103
  Private Const errReadFileError = 102
  Private Const errSettingsAlreadyExists = 106
  Private Const errTooNewFileFormat_Read = 104
  Private Const errTooNewFileFormat_Write = 105
  Private Const errWriteFileError = 101

  ' Ressourcen
  Private Const resIcon = 1002

  ' Konstanten für CustomDraw
  Private Const CDDS_PREPAINT = &H1
  Private Const CDDS_POSTPAINT = &H2
  Private Const CDDS_PREERASE = &H3
  Private Const CDDS_POSTERASE = &H4
  Private Const CDDS_MAPPART = &H5
  Private Const CDDS_ITEM = &H10000
  Private Const CDDS_SUBITEM = &H20000
  Private Const CDDS_ITEMPOSTERASE = (CDDS_ITEM Or CDDS_POSTERASE)
  Private Const CDDS_ITEMPOSTPAINT = (CDDS_ITEM Or CDDS_POSTPAINT)
  Private Const CDDS_ITEMPREERASE = (CDDS_ITEM Or CDDS_PREERASE)
  Private Const CDDS_ITEMPREPAINT = (CDDS_ITEM Or CDDS_PREPAINT)
  Private Const CDIS_SELECTED = &H1
  Private Const CDIS_GRAYED = &H2
  Private Const CDIS_DISABLED = &H4
  Private Const CDIS_CHECKED = &H8
  Private Const CDIS_FOCUS = &H10
  Private Const CDIS_DEFAULT = &H20
  Private Const CDIS_HOT = &H40
  Private Const CDIS_MARKED = &H80
  Private Const CDIS_INDETERMINATE = &H100
  Private Const CDIS_SHOWKEYBOARDCUES = &H200
  Private Const CDIS_NEARHOT = &H400
  Private Const CDIS_OTHERSIDEHOT = &H800
  Private Const CDIS_DROPHILITED = &H1000
  Private Const CDIS_MYFOCUS = &H100000     ' ACHTUNG! selbst hinzugefügt
  Private Const CDRF_DODEFAULT = &H0
  Private Const CDRF_NEWFONT = &H2
  Private Const CDRF_SKIPDEFAULT = &H4
  Private Const CDRF_DOERASE = &H8
  Private Const CDRF_NOTIFYPOSTPAINT = &H10
  Private Const CDRF_NOTIFYITEMDRAW = &H20
  Private Const CDRF_NOTIFYSUBITEMDRAW = &H20
  Private Const CDRF_NOTIFYPOSTERASE = &H40
  Private Const CDRF_SKIPPOSTPAINT = &H100
  Private Const TVCDRF_NOIMAGES = &H10000

  ' Konstanten für WM_HSCROLL und WM_VSCROLL
  #If Debuging Then
    Private Const SB_BOTTOM = 7
    Private Const SB_ENDSCROLL = 8
    Private Const SB_LEFT = 6
    Private Const SB_PAGEDOWN = 3
    Private Const SB_PAGELEFT = 2
    Private Const SB_PAGERIGHT = 3
    Private Const SB_PAGEUP = 2
    Private Const SB_RIGHT = 7
    Private Const SB_THUMBPOSITION = 4
    Private Const SB_THUMBTRACK = 5
    Private Const SB_TOP = 6
  #End If
  Private Const SB_LINEDOWN = 1
  Private Const SB_LINELEFT = 0
  Private Const SB_LINERIGHT = 1
  Private Const SB_LINEUP = 0

  ' Konstanten für TVM_EXPAND
  #If Debuging Then
    Private Const TVE_ACTIONMASK = &H7     ' MS internal
    Private Const TVE_EXPANDPARTIAL = &H4000
    Private Const TVE_TOGGLE = &H3
  #End If
  Private Const TVE_COLLAPSE = &H1
  Private Const TVE_COLLAPSERESET = &H8000
  Private Const TVE_EXPAND = &H2

  ' Konstanten für TVM_GETNEXTITEM
  Private Const TVGN_ROOT = &H0
  Private Const TVGN_NEXT = &H1
  Private Const TVGN_PREVIOUS = &H2
  Private Const TVGN_PARENT = &H3
  Private Const TVGN_CHILD = &H4
  Private Const TVGN_FIRSTVISIBLE = &H5
  Private Const TVGN_NEXTVISIBLE = &H6
  Private Const TVGN_PREVIOUSVISIBLE = &H7
  Private Const TVGN_DROPHILITE = &H8
  Private Const TVGN_CARET = &H9
  Private Const TVGN_LASTVISIBLE = &HA
  Private Const TVGN_NEXTSELECTED = &HB

  ' Konstanten für TVM_HITTEST
  Private Const TVHT_NOWHERE = &H1
  Private Const TVHT_ONITEMICON = &H2
  Private Const TVHT_ONITEMLABEL = &H4
  Private Const TVHT_ONITEMINDENT = &H8
  Private Const TVHT_ONITEMBUTTON = &H10
  Private Const TVHT_ONITEMRIGHT = &H20
  Private Const TVHT_ONITEMSTATEICON = &H40
  Private Const TVHT_ABOVE = &H100
  Private Const TVHT_BELOW = &H200
  Private Const TVHT_TORIGHT = &H400
  Private Const TVHT_TOLEFT = &H800
  Private Const TVHT_ONITEM = (TVHT_ONITEMICON Or TVHT_ONITEMLABEL Or TVHT_ONITEMSTATEICON)
  Private Const TVHT_OnItemLine = (TVHT_ONITEM Or TVHT_ONITEMBUTTON Or TVHT_ONITEMINDENT Or TVHT_ONITEMRIGHT)

  Public Enum HitTestConstants
    htAboveClientArea = TVHT_ABOVE
    htBelowClientArea = TVHT_BELOW
    htBelowLastItem = TVHT_NOWHERE
    htLeftOfClientArea = TVHT_TOLEFT
    htOnItem = TVHT_ONITEM
    htOnItemButton = TVHT_ONITEMBUTTON
    htOnItemIcon = TVHT_ONITEMICON
    htOnItemIndent = TVHT_ONITEMINDENT
    htOnItemLabel = TVHT_ONITEMLABEL
    htOnItemLine = TVHT_OnItemLine
    htOnItemRight = TVHT_ONITEMRIGHT
    htOnItemStateIcon = TVHT_ONITEMSTATEICON
    htRightOfClientArea = TVHT_TORIGHT
  End Enum

  ' Konstanten für TVM_GETITEMPARTRECT
  Private Const TVGIPR_BUTTON = &H1

  ' Konstanten für TVM_INSERTITEM und TVN_GETDISPINFO
  #If Debuging Then
    Private Const TVIF_ALL = &H7F     ' MS internal
    Private Const TVIF_RESERVED = &HF000     ' MS internal
  #End If
  Private Const TVIF_DI_SETITEM = &H1000
  Private Const TVI_FIRST = &HFFFF0001
  Private Const TVI_LAST = &HFFFF0002
  Private Const TVI_ROOT = &HFFFF0000
  Private Const TVI_SORT = &HFFFF0003
  Private Const TVIF_TEXT = &H1
  Private Const TVIF_IMAGE = &H2
  Private Const TVIF_PARAM = &H4
  Private Const TVIF_STATE = &H8
  Private Const TVIF_HANDLE = &H10
  Private Const TVIF_SELECTEDIMAGE = &H20
  Private Const TVIF_CHILDREN = &H40
  Private Const TVIF_INTEGRAL = &H80
  Private Const TVIF_STATEEX = &H100
  Private Const TVIF_EXPANDEDIMAGE = &H200

  ' Konstanten für TVM_SELECTITEM(?)
  #If Debuging Then
    Private Const TVSI_NOSINGLEEXPAND = &H8000
  #End If

  ' Konstanten für TVM_SETIMAGELIST
  Private Const TVSIL_NORMAL = 0
  Private Const TVSIL_STATE = 2

  ' Konstanten für TVN_ASYNCDRAW
  Private Const ADRF_DRAWSYNC = 0
  Private Const ADRF_DRAWNOTHING = 1
  Private Const ADRF_DRAWIMAGE = 2
  Private Const ILDRF_IMAGELOWQUALITY = &H1
  Private Const ILDRF_OVERLAYLOWQUALITY = &H10
  Private Const E_PENDING = &H8000000A

  ' Konstanten für TVN_SINGLEEXPAND
  Private Const TVNRET_DEFAULT = 0
  Private Const TVNRET_SKIPNEW = 2
  Private Const TVNRET_SKIPOLD = 1

  ' TreeView-Item states
  #If Debuging Then
    Private Const TVIS_ALL = &HFF7F     ' MS internal
    Private Const TVIS_USERMASK = &HF000
  #End If
  Private Const TVIS_SELECTED = &H2
  Private Const TVIS_CUT = &H4
  Private Const TVIS_DROPHILITED = &H8
  Private Const TVIS_BOLD = &H10
  Private Const TVIS_EXPANDED = &H20
  Private Const TVIS_EXPANDEDONCE = &H40
  Private Const TVIS_EXPANDPARTIAL = &H80
  Private Const TVIS_OVERLAYMASK = &HF00
  Private Const TVIS_STATEIMAGEMASK = &HF000

  ' TreeView-Item extended states
  Private Const TVIS_EX_FLAT = &H1
  Private Const TVIS_EX_DISABLED = &H2
  Private Const TVIS_EX_ALL = &H2

  ' CommonControls-Messages
  #If Debuging Then
    Private Const CCM_FIRST = &H2000
    Private Const CCM_DPISCALE = (CCM_FIRST + &HC)
    Private Const CCM_GETCOLORSCHEME = (CCM_FIRST + 3)
    Private Const CCM_GETDROPTARGET = (CCM_FIRST + 4)
    Private Const CCM_GETUNICODEFORMAT = (CCM_FIRST + 6)
    Private Const CCM_GETVERSION = (CCM_FIRST + &H8)
    Private Const CCM_SETBKCOLOR = (CCM_FIRST + 1)
    Private Const CCM_SETCOLORSCHEME = (CCM_FIRST + 2)
    Private Const CCM_SETNOTIFYWINDOW = (CCM_FIRST + &H9)
    Private Const CCM_SETUNICODEFORMAT = (CCM_FIRST + 5)
    Private Const CCM_SETVERSION = (CCM_FIRST + &H7)
    Private Const CCM_SETWINDOWTHEME = (CCM_FIRST + &HB)
  #End If

  ' TreeView-Messages
  #If Debuging Then
    Private Const TVM_CREATEDRAGIMAGE = (TV_FIRST + 18)
    Private Const TVM_FINDITEM = (TV_FIRST + 3)    ' MS internal, not implemented
    Private Const TVM_GETINSERTMARKCOLOR = (TV_FIRST + 38)
    Private Const TVM_GETISEARCHSTRINGW = (TV_FIRST + 64)
    Private Const TVM_GETISEARCHSTRING = (TV_FIRST + 23)
    Private Const TVM_GETUNICODEFORMAT = (CCM_FIRST + 6)
    Private Const TVM_MAPACCIDTOHTREEITEM = (TV_FIRST + 42)
    Private Const TVM_MAPHTREEITEMTOACCID = (TV_FIRST + 43)
    Private Const TVM_SETINSERTMARK = (TV_FIRST + 26)
    Private Const TVM_SETINSERTMARKCOLOR = (TV_FIRST + 37)
    Private Const TVM_SETTOOLTIPS = (TV_FIRST + 24)
    Private Const TVM_SETUNICODEFORMAT = (CCM_FIRST + 5)
    Private Const TVM_SORTCHILDREN = (TV_FIRST + 19)
  #End If
  Private Const TVM_DELETEITEM = (TV_FIRST + 1)
  Private Const TVM_EDITLABELW = (TV_FIRST + 65)
  Private Const TVM_EDITLABEL = (TV_FIRST + 14)
  Private Const TVM_ENDEDITLABELNOW = (TV_FIRST + 22)
  Private Const TVM_ENSUREVISIBLE = (TV_FIRST + 20)
  Private Const TVM_EXPAND = (TV_FIRST + 2)
  Private Const TVM_GETBKCOLOR = (TV_FIRST + 31)
  Private Const TVM_GETBORDER = (TV_FIRST + 36)
  Private Const TVM_GETCOUNT = (TV_FIRST + 5)
  Private Const TVM_GETEDITCONTROL = (TV_FIRST + 15)
  Private Const TVM_GETEXTENDEDSTYLE = (TV_FIRST + 45)
  Private Const TVM_GETIMAGELIST = (TV_FIRST + 8)
  Private Const TVM_GETINDENT = (TV_FIRST + 6)
  Private Const TVM_GETITEMW = (TV_FIRST + 62)
  Private Const TVM_GETITEM = (TV_FIRST + 12)
  Private Const TVM_GETITEMHEIGHT = (TV_FIRST + 28)
  Private Const TVM_GETITEMPARTRECT = (TV_FIRST + 72)
  Private Const TVM_GETITEMRECT = (TV_FIRST + 4)
  Private Const TVM_GETITEMSTATE = (TV_FIRST + 39)
  Private Const TVM_GETLINECOLOR = (TV_FIRST + 41)
  Private Const TVM_GETNEXTITEM = (TV_FIRST + 10)
  Private Const TVM_GETSCROLLTIME = (TV_FIRST + 34)
  Private Const TVM_GETSELECTEDCOUNT = (TV_FIRST + 70)
  Private Const TVM_GETTEXTCOLOR = (TV_FIRST + 32)
  Private Const TVM_GETTOOLTIPS = (TV_FIRST + 25)
  Private Const TVM_GETVISIBLECOUNT = (TV_FIRST + 16)
  Private Const TVM_HITTEST = (TV_FIRST + 17)
  Private Const TVM_INSERTITEMW = (TV_FIRST + 50)
  Private Const TVM_INSERTITEM = (TV_FIRST + 0)
  Private Const TVM_SELECTITEM = (TV_FIRST + 11)
  Private Const TVM_SETAUTOSCROLLINFO = (TV_FIRST + 59)
  Private Const TVM_SETBKCOLOR = (TV_FIRST + 29)
  Private Const TVM_SETBORDER = (TV_FIRST + 35)
  Private Const TVM_SETEXTENDEDSTYLE = (TV_FIRST + 44)
  Private Const TVM_SETIMAGELIST = (TV_FIRST + 9)
  Private Const TVM_SETINDENT = (TV_FIRST + 7)
  Private Const TVM_SETITEMW = (TV_FIRST + 63)
  Private Const TVM_SETITEM = (TV_FIRST + 13)
  Private Const TVM_SETITEMHEIGHT = (TV_FIRST + 27)
  Private Const TVM_SETLINECOLOR = (TV_FIRST + 40)
  Private Const TVM_SETSCROLLTIME = (TV_FIRST + 33)
  Private Const TVM_SETTEXTCOLOR = (TV_FIRST + 30)
  Private Const TVM_SHOWINFOTIP = (TV_FIRST + 71)
  Private Const TVM_SORTCHILDRENCB = (TV_FIRST + 21)

  ' CommonControls-Notifications
  #If Debuging Then
    Private Const NM_BTNCLK = (NM_FIRST - 11)     ' MS internal
    Private Const NM_CHAR = (NM_FIRST - 18)
    Private Const NM_DBLCLK = (NM_FIRST - 3)
    Private Const NM_ENDWAIT = (NM_FIRST - 10)     ' MS internal
    Private Const NM_HOVER = (NM_FIRST - 13)
    Private Const NM_KEYDOWN = (NM_FIRST - 15)
    Private Const NM_KILLFOCUS = (NM_FIRST - 8)
    Private Const NM_LDOWN = (NM_FIRST - 20)
    Private Const NM_NCHITTEST = (NM_FIRST - 14)
    Private Const NM_OUTOFMEMORY = (NM_FIRST - 1)
    Private Const NM_RDBLCLK = (NM_FIRST - 6)
    Private Const NM_RDOWN = (NM_FIRST - 21)
    Private Const NM_RELEASEDCAPTURE = (NM_FIRST - 16)
    Private Const NM_RETURN = (NM_FIRST - 4)
    Private Const NM_SETCURSOR = (NM_FIRST - 17)
    Private Const NM_SETFOCUS = (NM_FIRST - 7)
    Private Const NM_STARTWAIT = (NM_FIRST - 9)     ' MS internal
    Private Const NM_THEMECHANGED = (NM_FIRST - 22)
    Private Const NM_TOOLTIPSCREATED = (NM_FIRST - 19)
  #End If
  Private Const NM_CLICK = (NM_FIRST - 2)
  Private Const NM_CUSTOMDRAW = (NM_FIRST - 12)
  Private Const NM_CUSTOMTEXT = (NM_FIRST - 24)
  Private Const NM_FONTCHANGED = (NM_FIRST - 23)
  Private Const NM_RCLICK = (NM_FIRST - 5)
  Private Const NM_TVSTATEIMAGECHANGING = (NM_FIRST - 24)

  ' TreeView-Notifications
  #If Debuging Then
    Private Const TVN_GETDISPINFOW = (TVN_FIRST - 52)
    Private Const TVN_GETINFOTIPW = (TVN_FIRST - 14)
    Private Const TVN_KEYDOWN = (TVN_FIRST - 12)
    Private Const TVN_SETDISPINFOW = (TVN_FIRST - 53)
    Private Const TVN_SETDISPINFO = (TVN_FIRST - 4)
  #End If
  Private Const TVN_ASYNCDRAW = (TVN_FIRST - 20)
  Private Const TVN_BEGINDRAGW = (TVN_FIRST - 56)
  Private Const TVN_BEGINDRAG = (TVN_FIRST - 7)
  Private Const TVN_BEGINLABELEDITW = (TVN_FIRST - 59)
  Private Const TVN_BEGINLABELEDIT = (TVN_FIRST - 10)
  Private Const TVN_BEGINRDRAGW = (TVN_FIRST - 57)
  Private Const TVN_BEGINRDRAG = (TVN_FIRST - 8)
  Private Const TVN_DELETEITEMW = (TVN_FIRST - 58)
  Private Const TVN_DELETEITEM = (TVN_FIRST - 9)
  Private Const TVN_ENDLABELEDITW = (TVN_FIRST - 60)
  Private Const TVN_ENDLABELEDIT = (TVN_FIRST - 11)
  Private Const TVN_GETDISPINFO = (TVN_FIRST - 3)
  Private Const TVN_GETINFOTIP = (TVN_FIRST - 13)
  Private Const TVN_ITEMCHANGEDW = (TVN_FIRST - 19)
  Private Const TVN_ITEMCHANGED = (TVN_FIRST - 18)
  Private Const TVN_ITEMCHANGINGW = (TVN_FIRST - 17)
  Private Const TVN_ITEMCHANGING = (TVN_FIRST - 16)
  Private Const TVN_ITEMEXPANDEDW = (TVN_FIRST - 55)
  Private Const TVN_ITEMEXPANDED = (TVN_FIRST - 6)
  Private Const TVN_ITEMEXPANDINGW = (TVN_FIRST - 54)
  Private Const TVN_ITEMEXPANDING = (TVN_FIRST - 5)
  Private Const TVN_SELCHANGEDW = (TVN_FIRST - 51)
  Private Const TVN_SELCHANGED = (TVN_FIRST - 2)
  Private Const TVN_SELCHANGINGW = (TVN_FIRST - 50)
  Private Const TVN_SELCHANGING = (TVN_FIRST - 1)
  Private Const TVN_SINGLEEXPAND = (TVN_FIRST - 15)

  ' Konstanten für BitBlt
  #If Debuging Then
    Private Const CAPTUREBLT = &H40000000
    Private Const DSTINVERT = &H550009
    Private Const MERGECOPY = &HC000CA
    Private Const MERGEPAINT = &HBB0226
    Private Const NOMIRRORBITMAP = &H80000000
    Private Const NOTSRCCOPY = &H330008
    Private Const NOTSRCERASE = &H1100A6
    Private Const PATCOPY = &HF00021
    Private Const PATINVERT = &H5A0049
    Private Const PATPAINT = &HFB0A09
    Private Const SRCERASE = &H440328
    Private Const WHITENESS = &HFF0062
  #End If
  Private Const BLACKNESS = &H42&
  Private Const SRCAND = &H8800C6
  Private Const SRCCOPY = &HCC0020
  Private Const SRCINVERT = &H660046
  Private Const SRCPAINT = &HEE0086

  ' Konstanten für CreateFile
  #If Debuging Then
    Private Const CREATE_NEW = 1
    Private Const FILE_FLAG_BACKUP_SEMANTICS = &H2000000
    Private Const FILE_FLAG_DELETE_ON_CLOSE = &H4000000
    Private Const FILE_FLAG_NO_BUFFERING = &H20000000
    Private Const FILE_FLAG_OPEN_NO_RECALL = &H100000
    Private Const FILE_FLAG_OPEN_REPARSE_POINT = &H200000
    Private Const FILE_FLAG_OVERLAPPED = &H40000000
    Private Const FILE_FLAG_POSIX_SEMANTICS = &H1000000
    Private Const FILE_FLAG_RANDOM_ACCESS = &H10000000
    Private Const FILE_FLAG_SEQUENTIAL_SCAN = &H8000000
    Private Const FILE_FLAG_WRITE_THROUGH = &H80000000
    Private Const FILE_SHARE_DELETE = &H4
    Private Const FILE_SHARE_WRITE = &H2
    Private Const OPEN_ALWAYS = 4
    Private Const TRUNCATE_EXISTING = 5
  #End If
  Private Const CREATE_ALWAYS = 2
  Private Const FILE_SHARE_READ = &H1
  Private Const GENERIC_READ = &H80000000
  Private Const GENERIC_WRITE = &H40000000
  Private Const INVALID_HANDLE_VALUE = -1
  Private Const OPEN_EXISTING = 3

  ' Konstanten für FormatMessage
  #If Debuging Then
    Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
    Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
    Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
    Private Const FORMAT_MESSAGE_FROM_STRING = &H400
    Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
  #End If
  Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
  Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200

  ' Konstanten für GetAncestor
  #If Debuging Then
    Private Const GA_PARENT = 1
    Private Const GA_ROOTOWNER = 3
  #End If
  Private Const GA_ROOT = 2

  ' Konstanten für GetMenuDefaultItem
  #If Debuging Then
    Private Const GMDI_USEDISABLED = &H1
  #End If
  Private Const GMDI_GOINTOPOPUPS = &H2

  ' Konstanten für ImageList_Draw(Ex)
  Private Const ILD_NORMAL = &H0
  Private Const ILD_TRANSPARENT = &H1
  Private Const ILD_BLEND25 = &H2
  Private Const ILD_FOCUS = ILD_BLEND25
  Private Const ILD_BLEND50 = &H4
  Private Const ILD_BLEND = ILD_BLEND50
  Private Const ILD_SELECTED = ILD_BLEND50
  Private Const ILD_BLEND75 = &H8     ' MS internal, not implemented
  Private Const ILD_BLENDMASK = &HE     ' MS internal
  Private Const ILD_MASK = &H10
  Private Const ILD_IMAGE = &H20
  Private Const ILD_ROP = &H40
  Private Const ILD_OVERLAYMASK = &HF00
  Private Const ILD_PRESERVEALPHA = &H1000
  Private Const ILD_SCALE = &H2000
  Private Const ILD_DPISCALE = &H4000
  Private Const ILD_ASYNC = &H8000
  Private Const ILS_NORMAL = &H0
  Private Const ILS_GLOW = &H1
  Private Const ILS_SHADOW = &H2
  Private Const ILS_SATURATE = &H4
  Private Const ILS_ALPHA = &H8

  ' Konstanten für SendMessage
  #If Debuging Then
    Private Const HWND_BROADCAST = &HFFFF&
  #End If

  ' Konstanten für SetFilePointer
  #If Debuging Then
    Private Const FILE_END = 2
  #End If
  Private Const FILE_BEGIN = 0
  Private Const FILE_CURRENT = 1

  ' Konstanten für SetWindowPos
  #If Debuging Then
    Private Const HWND_BOTTOM = 1
    Private Const HWND_NOTOPMOST = -2
    Private Const HWND_TOP = 0
    Private Const SWP_ASYNCWINDOWPOS = &H4000
    Private Const SWP_DEFERERASE = &H2000
    Private Const SWP_HIDEWINDOW = &H80
    Private Const SWP_NOCOPYBITS = &H100
    Private Const SWP_NOOWNERZORDER = &H200
    Private Const SWP_NOREDRAW = &H8
    Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
    Private Const SWP_NOSENDCHANGING = &H400
    Private Const SWP_SHOWWINDOW = &H40
  #End If
  Private Const HWND_TOPMOST = -1
  Private Const SWP_FRAMECHANGED = &H20
  Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED
  Private Const SWP_NOACTIVATE = &H10
  Private Const SWP_NOMOVE = &H2
  Private Const SWP_NOSIZE = &H1
  Private Const SWP_NOZORDER = &H4

  ' Konstanten für ShowWindow
  #If Debuging Then
    Private Const SW_FORCEMINIMIZE = 11
    Private Const SW_MAX = 11
    Private Const SW_MAXIMIZE = 3
    Private Const SW_MINIMIZE = 6
    Private Const SW_NORMAL = 1
    Private Const SW_RESTORE = 9
    Private Const SW_SHOW = 5
    Private Const SW_SHOWNA = 8
  #End If
  Private Const SW_HIDE = 0
  Private Const SW_SHOWDEFAULT = 10
  Private Const SW_SHOWMAXIMIZED = 3
  Private Const SW_SHOWMINIMIZED = 2
  Private Const SW_SHOWMINNOACTIVE = 7
  Private Const SW_SHOWNOACTIVATE = 4
  Private Const SW_SHOWNORMAL = 1

  Public Enum WindowModeConstants
    wmShowDefault = SW_SHOWDEFAULT
    wmShowHidden = SW_HIDE
    wmShowMaximized = SW_SHOWMAXIMIZED
    wmShowMinimized = SW_SHOWMINIMIZED
    wmShowMinimizedNotActive = SW_SHOWMINNOACTIVE
    wmShowNormal = SW_SHOWNORMAL
    wmShowNotActive = SW_SHOWNOACTIVATE
  End Enum


  ' lokale Types

  #If Debuging Then
    Private Type NMCHAR
      hdr As NMHDR
      ch As Long
      dwItemPrev As Long
      dwItemNext As Long
    End Type

    Private Type NMKEY
      hdr As NMHDR
      nVKey As Long
      uFlags As Long
    End Type

    Private Type NMMOUSE
      hdr As NMHDR
      dwItemSpec As Long
      dwItemData As Long
      pt As POINT
      dwHitInfo As POINT
    End Type

    Private Type NMOBJECTNOTIFY
      hdr As NMHDR
      IItem As Long
      pObject As Long
      hResult As Long
      dwFlags As Long
    End Type

    Private Type NMTOOLTIPSCREATED
      hdr As NMHDR
      hwndToolTips As Long
    End Type

    Private Type NMTVITEMCHANGE
      hdr As NMHDR
      uChanged As Long
      hItem As Long
      uStateNew As Long
      uStateOld As Long
      lParam As Long
    End Type

    Private Type NMTVKEYDOWN
      hdr As NMHDR
      wVKey As Integer
      Flags As Long
    End Type
  #End If

  Private Type IMAGELISTDRAWPARAMS
    cbSize As Long
    himl As Long
    i As Long
    hDCDst As Long
    x As Long
    y As Long
    cx As Long
    cy As Long
    xBitmap As Long
    yBitmap As Long
    rgbBk As Long
    rgbFg As Long
    fStyle As Long
    dwRop As Long
    fState As Long
    iFrame As Long
    crEffect As Long
  End Type

  Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
  End Type

  ' für Items in einem Archiv
  Private Type ITEMINFO_ARCHIVE
    itemType As Byte
    attributes As Integer
    hItem As Long
    hItem_ArchiveFile As Long
  End Type

  ' für "normale" Items
  Private Type ITEMINFO_COMMON
    itemType As Byte
    pIDLToDesktop As Long
  End Type

  ' für benutzerdefinierte Items (not used)
  Private Type ITEMINFO_CUSTOM
    itemType As Byte
    itemData As Long
  End Type

  Private Type ITEMMAPENTRY_COMMON
    hItem As Long
    pIDLToDesktop As Long
  End Type

  Private Type NMCUSTOMDRAW
    hdr As NMHDR
    dwDrawStage As Long
    hDC As Long
    rc As RECT
    dwItemSpec As Long
    uItemState As Long
    lItemlParam As Long
  End Type

  Private Type TVITEM
    Mask As Long
    hItem As Long
    state As Long
    StateMask As Long
    pszText As Long
    cchTextMax As Long
    iImage As Long
    iSelectedImage As Long
    cChildren As Long
    lParam As Long
  End Type

  Private Type TVITEMEX
    Mask As Long
    hItem As Long
    state As Long
    StateMask As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    iSelectedImage As Long
    cChildren As Long
    lParam As Long
    iIntegral As Long
    uStateEx As Long
    hWnd As Long
    iExpandedImage As Long
  End Type

  Private Type NMTREEVIEW
    hdr As NMHDR
    Action As Long
    itemOld As TVITEM
    itemNew As TVITEM
    ptDrag As POINT
  End Type

  Private Type NMTVASYNCDRAW
    hdr As NMHDR
    pimldp As Long     ' pointer to IMAGELISTDRAWPARAMS
    hr As Long
    hItem As Long
    lParam As Long
    dwRetFlags As Long
    iRetImageIndex As Long
  End Type

  Private Type NMTVCUSTOMDRAW
    nmcd As NMCUSTOMDRAW
    clrText As Long
    clrTextBk As Long
    iLevel As Long
  End Type

  Private Type NMTVDISPINFO
    hdr As NMHDR
    Item As TVITEM
  End Type

  Private Type NMTVDISPINFOEX
    hdr As NMHDR
    Item As TVITEMEX
  End Type

  Private Type NMTVGETINFOTIP
    hdr As NMHDR
    pszText As Long
    cchTextMax As Long
    hItem As Long
    lParam As Long
  End Type

  Private Type NMTVSTATEIMAGECHANGING
    hdr As NMHDR
    hti As Long
    iOldStateImageIndex As Long
    iNewStateImageIndex As Long
  End Type

  Private Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
  End Type

  Private Type TVGETITEMPARTRECTINFO
    hti As Long
    prc As Long     ' pointer to RECT
    partID As Long
  End Type

  Private Type TVHITTESTINFO
    pt As POINT
    Flags As Long
    hItem As Long
  End Type

  Private Type TVINSERTSTRUCT
    hParent As Long
    hInsertAfter As Long
    itemex As TVITEMEX
  End Type

  Private Type TVSORTCB
    hParent As Long
    lpfnCompare As Long
    lParam As Long
  End Type


  ' öffentliche Types

  Public Type ITEMRECT
    rcLeft As Long
    rcTop As Long
    rcRight As Long
    rcBottom As Long
  End Type


  ' lokale Variablen

  #If BkImageSupport Then
    Private bkImageBrush As Long
    Private bkImageTexture As Long
  #End If
  Private blockAutoUpdate As Long
  Private bPendingShellIconCacheRebuild As Boolean
  Private blockMouseLeave As Boolean
  Private blockMouseLeave_Edit As Boolean
  Private cachedIShFolderForSorting As IVBShellFolder
  Private cachedISHIcon As IVBShellIcon
  Private cachedISHIcon_hItem As Long     ' hItem that cachedISHIcon is for
  Private canceledToolTip As Boolean
  Private CBTHookStarted As Boolean
  Private checkForMouseLeave As Boolean
  Private checkForMouseLeave_Edit As Boolean
  Private createdFolder As Boolean
  Private curEffects As DROPEFFECTConstants
  Private curKeyState As Long
  Private curMouseButtons As Long
  Private curMouseButtons_Edit As Long
  Private curPt As POINT
  Private DataObj As DataObject
  #If Debuging Then
    Private debugger As clsDebugger
  #End If
  Private DI_GETDRAGIMAGE As Long
  Private disableCreateControl As Boolean
  Private disableLoadFileSystem As Boolean
  Private dontChangeRecyclerChildrenProp As Boolean
  Private dontExpand As Boolean
  Private dontSavePath As Boolean
  Private dragButton As Long
  Private dragCursorHotSpot_Copy As POINT
  Private dragCursorHotSpot_Link As POINT
  Private dragCursorHotSpot_Move As POINT
  Private dragCursorHotSpot_No As POINT
  Private dragCursorHotSpot_Null As POINT
  Private dragCursorIndex_Copy As Long
  Private dragCursorIndex_Link As Long
  Private dragCursorIndex_Move As Long
  Private dragCursorIndex_No As Long
  Private dragCursorIndex_Null As Long
  'Private hCollapsingItem As Long
  Private hCreatedItem As Long
  #If BkImageSupport Then
    Private hDC_BkImageTexture As Long
  #End If
  Private hDraggedItem As Long
  Private hImlDragCursors As Long
  Private hItemUnderMouse As Long
  Private hItemWithContMenu As Long
  #If BkImageSupport Then
    Private hPrevBMP_BkImageTexture As Long
  #End If
  Private hSHNotify As Long
  Private hWnd_EditContextMenu As Long
  Private hWnd_ParentForm As Long
  Private hWnd_ScrollBarContextMenu As Long
  Private hWnd_ShellContextMenu As Long
  Private hWnd_ToolTip As Long
  Private IContextMenu As IVBContextMenu
  Private IContextMenu2 As IVBContextMenu2
  Private IContextMenu3 As IVBContextMenu3
  Private IDrpTrgt As IVBDropTarget
  Private IDropTargetHelper As IVBDropTargetHelper
  Private IDSData As IDSHookData
  Private IPAOData As IPAOHookData
  Private IRoot As IVBShellFolder
  Private isSettingFocus As Boolean
  Private mustCustomDraw As Boolean
  Private needRecreate As Boolean
  Private pIDL_Desktop As Long
  Private ptLastMousePos As POINT
  Private ptLastMousePos_Edit As POINT
  Private raisedMouseDown As Long
  Private raisedMouseDown_Edit As Long
  Private rcClient As RECT
  Private rcWnd As RECT
  Private rcWnd_Edit As RECT
  Private releasedMouseButton As Boolean
  Private releasedMouseButton_Edit As Boolean
  Private RootItemInfo As ITEMINFO_COMMON
  Private SelChangedDelay_hNewItem As Long
  Private SelChangedDelay_hOldItem As Long
  Private SelChangeInternal As Boolean
  Private suppressSelChangeEvents As Boolean
  Private TimeOfPossibleMouseLeave As Long
  Private TimeOfPossibleMouseLeave_Edit As Long
  Private TimeOfFirstMouseMove As Long
  Private TimeOfFirstMouseMove_Edit As Long
  Private TimeOfLastMouseMove As Long
  Private TimeOfLastMouseMove_Edit As Long
'  Private TimeOfLastTotalUpdate As Long
  Private TimeOfLastWM_CONTEXTMENU_Edit As Long
  Private TimeOfLastWM_CONTEXTMENU_Shell As Long
  Private TimeOfRealSelChanged As Long
  Private useLargeIcons As Boolean
  Private useSHDoDragDrop As Boolean
  Private XOffset As Long


  ' interne Eigenschaften

  Private propDragging As Boolean
  Private propDragging_Edit As Boolean
  Private prophDropTargetItem As Long
  Private propHot As Boolean
  Private propHot_Edit As Boolean
  Private propMouseOverControl As Boolean
  Private propMouseOverControl_Edit As Boolean
  Private propNoAutoUpdate As Boolean


  ' Property-Variablen

  Private propAllowDragDrop As Boolean
  Private propAlwaysUseSmallIcons As Boolean
  Private propAppearance As AppearanceConstants
  Private propArchiveItemBackColor As OLE_COLOR
  Private propArchiveItemForeColor As OLE_COLOR
  Private propArchivesAfterFolders As Boolean
  Private propAutoHScroll As Boolean
  Private propBackColor As OLE_COLOR
  #If BkImageSupport Then
    Private propBkImage As StdPicture
    Private propBkImageFixed As Boolean
    Private propBkImageStyle As BkImageStyleConstants
  #End If
  Private propBorderStyle As BorderStyleConstants
  Private propCheckBoxes As Boolean
  Private propCheckSlowItemsForSubItems As Boolean
  Private propCompressedBackColor As OLE_COLOR
  Private propCompressedForeColor As OLE_COLOR
  Private propDimSelectedItemsIcon As Boolean
  Private propDontCheckFolderExpandibility As Boolean
  Private propDragExpandTime As Long
  Private propDragScrollTime As Long
  Private propDragScrollZoneBottom As Long
  Private propDragScrollZoneLeft As Long
  Private propDragScrollZoneRight As Long
  Private propDragScrollZoneTop As Long
  Private propDragShowAlphaBlendedItems As Boolean
  Private propDrivesOnly As Boolean
  Private propEnabled As Boolean
  Private propEncryptedBackColor As OLE_COLOR
  Private propEncryptedForeColor As OLE_COLOR
  Private propExpandArchives As ExpandArchivesConstants
  Private propFadeExpandos As Boolean
  Private propFastInfoTipsOnly As Boolean
  Private propFavoritesStyle As Boolean
  Private propFileAttributes As AttributesConstants
  Private propFileFilters As String
  Private propFireBeforeInsertItem As FireBeforeInsertItemConstants
  Private propFolderAttributes As AttributesConstants
  Private propFolderFilters As String
  Private WithEvents propFont As StdFont
Attribute propFont.VB_VarHelpID = -1
  Private propForeColor As OLE_COLOR
  Private propFullRowSelect As Boolean
  Private prophDC As Long
  Private prophEditetItem As Long
  Private prophImageList As Long
  Private prophShellContextMenu As Long
  Private propHotBackColor As OLE_COLOR
  Private propHotForeColor As OLE_COLOR
  Private propHotTracking As Boolean
  Private propHoverTime As Long
  Private prophLargeStateImageList As Long
  Private prophSmallStateImageList As Long
  Private prophWnd As Long
  Private prophWndShellUIParentWindow As Long
  Private propIncludedItems As IncludedItemsConstants
  Private propIndent As Long
  Private propItemHeight As Long
  Private propLimitEditInput As Boolean
  Private propLineColor As OLE_COLOR
  Private propLineStyle As LineStyleConstants
  Private propMargin As Long
  Private propMaxScrollTime As Long
  Private propMouseIcon As StdPicture
  Private propMousePointer As MousePointerConstants
  Private propPath As String
  Private propProcessContextMenuKeys As Boolean
  Private propRenaming As RenamingConstants
  Private propReplaceHandCursor As Boolean
  Private propRootPath As String
  Private propRTLReading As Boolean
  Private propScrollBars As ScrollBarsConstants
  Private propSelChangedDelayTime As Long
  Private propSelectedItemBackColor As OLE_COLOR
  Private propSelectedItemForeColor As OLE_COLOR
  Private propShowFocusRect As Boolean
  Private propShowHiddenDimmed As Boolean
  Private propShowIcons As Boolean
  Private propShownOverlays As ShownOverlaysConstants
  Private propShownToolTips As ShownToolTipsConstants
  Private propShowRoot As Boolean
  Private propShowSelection As Boolean
  Private propShowShellContextMenus As Boolean
  Private propSingleExpand As SingleExpandConstants
  Private propSortingCriteria As Long
  Private propSortOrder As SortOrderConstants
  Private propTreeViewStyle As TreeViewStyleConstants
  Private propUseCustomIcons As Boolean
  Private propUseFileFilters As Boolean
  Private propUseFolderFilters As Boolean
  Private propUseSelectedIcons As Boolean
  Private propUseStandardIcons As Boolean
  Private propUseStandardIconsForSlowItems As Boolean


  ' lokale APIs

  Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDC_Dst As Long, ByVal xDst As Long, ByVal yDst As Long, ByVal cxDst As Long, ByVal cyDst As Long, ByVal hDC_Src As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal Mode As Long) As Long
  Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
  Private Declare Function CreateBitmapAsLong Lib "gdi32.dll" Alias "CreateBitmap" (ByVal cx As Long, ByVal cy As Long, ByVal planes As Long, ByVal bitCount As Long, ByVal pBits As Long) As Long
  Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hDC As Long, ByVal Width As Long, ByVal Height As Long) As Long
  Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
  Private Declare Function CreateFileAsLong Lib "kernel32.dll" Alias "CreateFileA" (ByVal File As String, ByVal AccessMode As Long, ByVal ShareMode As Long, ByVal Security As Long, ByVal CreationFlags As Long, ByVal FileAttributes As Long, ByVal hTemplateFile As Long) As Long
  Private Declare Function CreatePatternBrush Lib "gdi32.dll" (ByVal hBitmap As Long) As Long
  Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal Clr As Long) As Long
  Private Declare Function CreateWindowExAsLong Lib "user32.dll" Alias "CreateWindowExA" (ByVal StyleEx As Long, ByVal Klasse As String, ByVal Caption As String, ByVal Style As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal hWnd_Parent As Long, ByVal hMenu As Long, ByVal hInstanz As Long, ByVal lParam As Long) As Long
  Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
  Private Declare Function DestroyWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
  Private Declare Function DoDragDrop Lib "ole32.dll" (ByVal IDataObj As IVBDataObject, ByVal IDrpSrc As IVBDropSource, ByVal AllowedEffects As Long, EffectUsed As Long) As Long
'  Private Declare Function DrawFocusRect Lib "user32.dll" (ByVal hDC As Long, rc As RECT) As Long
  Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
  Private Declare Function EnableWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal TrueFalse As Long) As Long
  Private Declare Function FillRect Lib "user32.dll" (ByVal hDC As Long, rc As RECT, ByVal hBrush As Long) As Long
  Private Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageA" (ByVal Flags As Long, Source As Any, ByVal MessageId As Long, ByVal LanguageId As Long, ByVal buffer As String, ByVal bufferSize As Long, Arguments As Long) As Long
  Private Declare Function GetAncestor Lib "user32.dll" (ByVal hWnd As Long, ByVal Flags As Long) As Long
  Private Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal buffer As String, ByVal bufferSize As Long) As Long
  Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
  Private Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hDC As Long, ByVal nIndex As Long) As Long
  Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, FileSizeHigh As Long) As Long
  Private Declare Function GetFocus Lib "user32" () As Long
  Private Declare Function GetMenuDefaultItem Lib "user32.dll" (ByVal hMenu As Long, ByVal Mode As Long, ByVal Flags As Long) As Long
  Private Declare Function GetObjectAPI Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal bufferSize As Long, Data As Any) As Long
  Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
  Private Declare Function GetProcessHeap Lib "kernel32.dll" () As Long
  Private Declare Function GetScrollInfo Lib "user32.dll" (ByVal hWnd As Long, ByVal SBType As Long, SBInfo As SCROLLINFO) As Long
  Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
  Private Declare Function GetSysColorBrush Lib "user32.dll" (ByVal nIndex As Long) As Long
  Private Declare Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal Prop As Long) As Long
  Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, rc As RECT) As Long
  Private Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
  Private Declare Function GlobalSize Lib "kernel32.dll" (ByVal hMem As Long) As Long
  Private Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
  Private Declare Function HeapAlloc Lib "kernel32.dll" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
  Private Declare Function HeapFree Lib "kernel32.dll" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal lpMem As Long) As Long
  Private Declare Function ImageList_Create Lib "comctl32.dll" (ByVal cx As Long, ByVal cy As Long, ByVal Flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
  Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal himl As Long) As Long
  Private Declare Function ImageList_DrawEx Lib "comctl32.dll" (ByVal himl As Long, ByVal Index As Long, ByVal hDC_Dest As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal BackClr As Long, ByVal ForeClr As Long, ByVal Style As Long) As Long
  Private Declare Function ImageList_GetBkColor Lib "comctl32.dll" (ByVal himl As Long) As Long
  Private Declare Function ImageList_GetDragImage Lib "comctl32.dll" (ByVal ppt As Long, ByVal pptHotSpot As Long) As Long
  Private Declare Function ImageList_GetIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal Flags As Long) As Long
  Private Declare Function ImageList_GetIconSize Lib "comctl32.dll" (ByVal himl As Long, ByVal cx As Long, ByVal cy As Long) As Long
  Private Declare Function ImageList_GetImageCount Lib "comctl32.dll" (ByVal himl As Long) As Long
  Private Declare Function ImageList_SetBkColor Lib "comctl32.dll" (ByVal himl As Long, ByVal BkClr As Long) As Long
  Private Declare Function ImageList_SetDragCursorImage Lib "comctl32.dll" (ByVal himlDrag As Long, ByVal iDrag As Long, ByVal dxHotspot As Long, ByVal dyHotspot As Long) As Long
  Private Declare Function InflateRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
  Private Declare Function InvalidateRectAsLong Lib "user32.dll" Alias "InvalidateRect" (ByVal hWnd As Long, ByVal rc As Long, ByVal eraseBackground As Long) As Long
  Private Declare Function IsPathShared Lib "ntshrui.dll" Alias "IsPathSharedA" (ByVal lpcszPath As String, ByVal bRefresh As Long) As Long
  Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
  Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal IDEvent As Long) As Long
  Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpFileName As String) As Long
  Private Declare Function lstrcpyAsLong1 Lib "kernel32.dll" Alias "lstrcpyA" (ByVal pDest As Long, ByVal Src As String) As Long
  Private Declare Function OemToChar Lib "user32" Alias "OemToCharA" (ByVal Src As String, ByVal Dst As String) As Long
  Private Declare Function OffsetRect Lib "user32" (rc As RECT, ByVal dx As Long, ByVal dy As Long) As Long
  Private Declare Function ReadFileAsLong Lib "kernel32" Alias "ReadFile" (ByVal hFile As Long, buffer As Any, ByVal BytesToRead As Long, BytesRead As Long, ByVal OverlappedData As Long) As Long
  Private Declare Function RegisterClipboardFormat Lib "user32.dll" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
  Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
  Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long
  Private Declare Sub ReleaseStgMedium Lib "ole32.dll" (ByVal lpstgmedium As Long)
  Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, lParam As Any) As Long
  Private Declare Function SetBkColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
  Private Declare Function SetBkMode Lib "gdi32.dll" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
  Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal DistToMoveLow As Long, DistToMoveHigh As Long, ByVal Flags As Long) As Long
  Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
  Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
  Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
  Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal IDEvent As Long, ByVal MilliSeconds As Long, ByVal pCallBack As Long) As Long
  Private Declare Function SetViewportOrgEx Lib "gdi32.dll" (ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, ByRef lpPoint As POINT) As Long
  Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal Flags As Long) As Long
  Private Declare Function SHChangeNotifyRegister Lib "shell32" Alias "#2" (ByVal hWnd As Long, ByVal fSources As SHCNRFConstants, ByVal fEvents As SHCNEConstants, ByVal wMsg As Long, ByVal cEntries As Long, pfsne As SHCHANGENOTIFYENTRY) As Long
  Private Declare Function SHChangeNotifyUnRegister Lib "shell32" Alias "#4" (ByVal hNotify As Long) As Long
  Private Declare Function SHDoDragDrop Lib "shell32.dll" (ByVal hWnd As Long, ByVal IDataObj As IVBDataObject, ByVal IDrpSrc As IVBDropSource, ByVal AllowedEffects As Long, EffectUsed As Long) As Long
  Private Declare Function SHGetDesktopFolder Lib "shell32.dll" (IFolder As IVBShellFolder) As Long
  Private Declare Function SHHandleUpdateImage Lib "shell32" (ByVal pIDL2 As Long) As Long
  Private Declare Function SHLimitInputEdit Lib "shell32" (ByVal hWnd As Long, ByVal ISHFolder As IVBShellFolder) As Long
  Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal Flags As Long) As Long
  Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
  Private Declare Function UpdateWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
  Private Declare Function WindowFromPoint Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
  Private Declare Function WriteFileAsLong Lib "kernel32" Alias "WriteFile" (ByVal hFile As Long, buffer As Any, ByVal BytesToWrite As Long, BytesWritten As Long, ByVal OverlappedData As Long) As Long


  ' öffentliche Events

  Public Event BeforeInsertItem(ByVal pIDLToDesktop As Long, ByVal hParentItem As Long, Cancel As Boolean)
Attribute BeforeInsertItem.VB_Description = "Wird ausgelöst, bevor ein Item (kein Archiv-Item) eingefügt wird, dessen Typ zu den mit FireBeforeInsertItem festgelegten gehört."
  Public Event BeforeShowEditContextMenu(Cancel As Boolean)
Attribute BeforeShowEditContextMenu.VB_Description = "Wird ausgelöst, bevor das Kontextmenü des Textfeldes zum Umbenennen von Items angezeigt wird. Mit <Cancel> kann der Vorgang abgebrochen werden."
  Public Event Click()
Attribute Click.VB_Description = "Wird ausgelöst, wenn der Benutzer mit der linken Maustaste in das Control klickt."
Attribute Click.VB_UserMemId = -600
  Public Event CompleteDrag(ByVal Canceled As Boolean, ByVal performedEffect As Long)
  Public Event DblClick()
Attribute DblClick.VB_Description = "Wird ausgelöst, wenn der Benutzer mit der linken Maustaste in das Control doppelklickt."
Attribute DblClick.VB_UserMemId = -601
  Public Event DestroyLabelEditControl()
Attribute DestroyLabelEditControl.VB_Description = "Wird ausgelöst, bevor das Textfeld zum Umbenennen von Items zerstört wird."
  Public Event DragEnter(Data As DataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Attribute DragEnter.VB_Description = "Wird ausgelöst, wenn <Data> in das Control gezogen wird."
  Public Event DragLeave()
Attribute DragLeave.VB_Description = "Wird ausgelöst, wenn bei einem Drag'n'Drop-Vorgang das Control verlassen wird."
  Public Event DragMove(Data As DataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Attribute DragMove.VB_Description = "Wird ausgelöst, wenn <Data> über das Control gezogen wird."
  Public Event Drop(Data As DataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Attribute Drop.VB_Description = "Wird ausgelöst, wenn <Data> über dem Control abgelegt wird."
  Public Event EditClick()
Attribute EditClick.VB_Description = "Wird ausgelöst, wenn der Benutzer mit der linken Maustaste in das Textfeld zum Umbenennen von Items klickt."
  Public Event EditDblClick()
Attribute EditDblClick.VB_Description = "Wird ausgelöst, wenn der Benutzer mit der linken Maustaste in das Textfeld zum Umbenennen von Items doppelklickt."
  Public Event EditKeyDown(KeyCode As Integer, Shift As Integer)
Attribute EditKeyDown.VB_Description = "Wird ausgelöst, wenn der Benutzer eine Taste drückt während das Textfeld zum Umbenennen von Items den Fokus besitzt."
  Public Event EditKeyPress(KeyAscii As Integer)
Attribute EditKeyPress.VB_Description = "Wird ausgelöst, wenn der Benutzer eine ANSI-Taste drückt und losläßt während das Textfeld zum Umbenennen von Items den Fokus besitzt."
  Public Event EditKeyUp(KeyCode As Integer, Shift As Integer)
Attribute EditKeyUp.VB_Description = "Wird ausgelöst, wenn der Benutzer eine Taste losläßt während das Textfeld zum Umbenennen von Items den Fokus besitzt."
  Public Event EditMClick()
Attribute EditMClick.VB_Description = "Wird ausgelöst, wenn der Benutzer mit der mittleren Maustaste in das Textfeld zum Umbenennen von Items klickt."
  Public Event EditMDblClick()
Attribute EditMDblClick.VB_Description = "Wird ausgelöst, wenn der Benutzer mit der mittleren Maustaste in das Textfeld zum Umbenennen von Items doppelklickt."
  Public Event EditMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Attribute EditMouseDown.VB_Description = "Wird ausgelöst, wenn der Benutzer über dem Textfeld zum Umbenennen von Items eine Maustaste drückt."
  Public Event EditMouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Attribute EditMouseEnter.VB_Description = "Wird ausgelöst, wenn der Benutzer die Maus in das Textfeld zum Umbenennen von Items hineinbewegt."
  Public Event EditMouseHover(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Attribute EditMouseHover.VB_Description = "Wird ausgelöst, wenn sich der Mauszeiger für <HoverTime> Millisekunden über dem Textfeld zum Umbenennen von Items befindet."
  Public Event EditMouseLeave(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Attribute EditMouseLeave.VB_Description = "Wird ausgelöst, wenn der Benutzer die Maus aus dem Textfeld zum Umbenennen von Items herausbewegt."
  Public Event EditMouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Attribute EditMouseMove.VB_Description = "Wird ausgelöst, wenn der Benutzer über dem Textfeld zum Umbenennen von Items die Maus bewegt."
  Public Event EditMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Attribute EditMouseUp.VB_Description = "Wird ausgelöst, wenn der Benutzer über dem Textfeld zum Umbenennen von Items eine Maustaste losläßt."
  Public Event EditRClick()
Attribute EditRClick.VB_Description = "Wird ausgelöst, wenn der Benutzer mit der rechten Maustaste in das Textfeld zum Umbenennen von Items klickt."
  Public Event EditRDblClick()
Attribute EditRDblClick.VB_Description = "Wird ausgelöst, wenn der Benutzer mit der rechten Maustaste in das Textfeld zum Umbenennen von Items doppelklickt."
  Public Event InitLabelEditControl()
Attribute InitLabelEditControl.VB_Description = "Wird ausgelöst, nachdem das Textfeld zum Umbenennen von Items erzeugt wurde."
  Public Event ItemBeforeExecuteShellContextMenuItem(ByVal hItem As Long, ByVal ItemID As Long, ByVal ItemVerb As String, ByVal ItemDescription As String, Cancel As Boolean, WindowMode As WindowModeConstants)
Attribute ItemBeforeExecuteShellContextMenuItem.VB_Description = "Wird ausgelöst, bevor der Befehl <ItemVerb> für den Item <hItem> ausgeführt wird. Mit <Cancel> kann der Vorgang abgebrochen werden."
  Public Event ItemBeforeGetInfoTip(ByVal hItem As Long, InfoTipStyle As InfoTipStyleConstants, Cancel As Boolean)
Attribute ItemBeforeGetInfoTip.VB_Description = "Wird ausgelöst, bevor der ToolTipText des Items <hItem> ermittelt wird. Mit <Cancel> kann der Vorgang abgebrochen werden."
  Public Event ItemBeforeInitShellContextMenu(ByVal hItem As Long, ContextMenuStyle As ShellContextMenuStyleConstants, Cancel As Boolean)
Attribute ItemBeforeInitShellContextMenu.VB_Description = "Wird ausgelöst, bevor das Shell-Kontextmenü des Items <hItem> erzeugt wird. Mit <Cancel> kann der Vorgang abgebrochen werden."
  Public Event ItemBeforeStateIconChange(ByVal hItem As Long, ByVal OldStateIcon As Long, NewStateIcon As Long, Cancel As Boolean)
Attribute ItemBeforeStateIconChange.VB_Description = "Wird ausgelöst, bevor das StateIcon des Items <hItem> geändert wird. Mit <Cancel> kann der Vorgang abgebrochen werden."
  Public Event ItemBeforeRenaming(ByVal hItem As Long, ByVal OldName As String, NewName As String, Cancel As Boolean)
Attribute ItemBeforeRenaming.VB_Description = "Wird ausgelöst, bevor der Item <hItem> umbenannt wird. Mit <Cancel> kann der Vorgang abgebrochen werden."
  Public Event ItemBeforeShowShellContextMenu(ByVal hItem As Long, Cancel As Boolean)
Attribute ItemBeforeShowShellContextMenu.VB_Description = "Wird ausgelöst, bevor das Shell-Kontextmenü des Items <hItem> angezeigt wird. Mit <Cancel> kann der Vorgang abgebrochen werden."
  Public Event ItemBeforeShowToolTip(ByVal hItem As Long, ByVal ToolTipType As ToolTipTypeConstants, ToolTipText As String, Cancel As Boolean)
Attribute ItemBeforeShowToolTip.VB_Description = "Wird ausgelöst, bevor der ToolTip des Items <hItem> angezeigt wird. <ToolTipType> gibt den Typ des anzuzeigenden ToolTips an. Mit <Cancel> kann der Vorgang für InfoTips abgebrochen werden. ToolTips für abgeschnittene Items können nicht gecancelt werden."
  Public Event ItemBeforeStartRenaming(ByVal hItem As Long, Cancel As Boolean)
Attribute ItemBeforeStartRenaming.VB_Description = "Wird ausgelöst, bevor das Textfeld zum Umbenennen von Items für den Item <hItem> erzeugt wird. Mit <Cancel> kann der Vorgang abgebrochen werden."
  Public Event ItemClick(ByVal hItem As Long)
Attribute ItemClick.VB_Description = "Wird ausgelöst, wenn der Benutzer mit der linken Maustaste auf den Item <hItem> klickt."
  Public Event ItemCollapsed(ByVal hItem As Long)
Attribute ItemCollapsed.VB_Description = "Wird ausgelöst, nachdem der Item <hItem> reduziert wurde."
  Public Event ItemCollapsing(ByVal hItem As Long, Cancel As Boolean)
Attribute ItemCollapsing.VB_Description = "Wird ausgelöst, bevor der Item <hItem> reduziert wird. Mit <Cancel> kann der Vorgang abgebrochen werden."
  Public Event ItemCompleteDrag(ByVal hItem As Long, ByVal Canceled As Boolean, ByVal performedEffect As Long)
Attribute ItemCompleteDrag.VB_Description = "Wird nach einer Drag'n'Drop-Aktion ausgelöst. ACHTUNG: 'hItem' und 'PerformedEffect' können ungültig sein!"
  Public Event ItemDblClick(ByVal hItem As Long)
Attribute ItemDblClick.VB_Description = "Wird ausgelöst, wenn der Benutzer mit der linken Maustaste auf den Item <hItem> doppelklickt."
  Public Event ItemDestroyShellContextMenu(ByVal hItem As Long)
Attribute ItemDestroyShellContextMenu.VB_Description = "Wird ausgelöst, nachdem das Shell-Kontextmenü des Items <hItem> zerstört wurde."
  Public Event ItemDragEnter(ByVal hItem As Long, Data As DataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Attribute ItemDragEnter.VB_Description = "Wird ausgelöst, wenn <Data> in den Item <hItem> gezogen wird."
  Public Event ItemDragLeave(ByVal hItem As Long)
Attribute ItemDragLeave.VB_Description = "Wird ausgelöst, wenn bei einem Drag'n'Drop-Vorgang der Item <hItem> verlassen wird."
  Public Event ItemDragMove(ByVal hItem As Long, Data As DataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Attribute ItemDragMove.VB_Description = "Wird ausgelöst, wenn <Data> über den Item <hItem> gezogen wird."
  Public Event ItemDrop(ByVal hItem As Long, Data As DataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Attribute ItemDrop.VB_Description = "Wird ausgelöst, wenn <Data> über dem Item <hItem> abgelegt wird."
  Public Event ItemExecutedShellContextMenuItem(ByVal hItem As Long, ByVal ItemID As Long, ByVal ItemVerb As String, ByVal ItemDescription As String)
Attribute ItemExecutedShellContextMenuItem.VB_Description = "Wird ausgelöst, nachdem der Befehl <ItemVerb> für den Item <hItem> ausgeführt wurde."
  Public Event ItemExpanded(ByVal hItem As Long)
Attribute ItemExpanded.VB_Description = "Wird ausgelöst, nachdem der Item <hItem> erweitert wurde."
  Public Event ItemExpanding(ByVal hItem As Long, Cancel As Boolean)
Attribute ItemExpanding.VB_Description = "Wird ausgelöst, bevor der Item <hItem> erweitert wird. Mit <Cancel> kann der Vorgang abgebrochen werden."
  Public Event ItemInitShellContextMenu(ByVal hItem As Long)
Attribute ItemInitShellContextMenu.VB_Description = "Wird ausgelöst, nachdem das Shell-Kontextmenü des Items <hItem> erzeugt wurde."
  Public Event ItemLoadedSubItems(ByVal hItem As Long)
Attribute ItemLoadedSubItems.VB_Description = "Wird ausgelöst, nachdem die SubItems des Items <hItem> geladen wurden."
  Public Event ItemLoadingSubItems(ByVal hItem As Long)
Attribute ItemLoadingSubItems.VB_Description = "Wird ausgelöst, bevor die SubItems des Items <hItem> geladen werden."
  Public Event ItemMClick(ByVal hItem As Long)
Attribute ItemMClick.VB_Description = "Wird ausgelöst, wenn der Benutzer mit der mittleren Maustaste auf den Item <hItem> klickt."
  Public Event ItemMDblClick(ByVal hItem As Long)
Attribute ItemMDblClick.VB_Description = "Wird ausgelöst, wenn der Benutzer mit der mittleren Maustaste auf den Item <hItem> doppelklickt."
  Public Event ItemMouseDown(ByVal hItem As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Attribute ItemMouseDown.VB_Description = "Wird ausgelöst, wenn der Benutzer über dem Item <hItem> eine Maustaste drückt."
  Public Event ItemMouseEnter(ByVal hItem As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Attribute ItemMouseEnter.VB_Description = "Wird ausgelöst, wenn der Benutzer die Maus in den Item <hItem> hineinbewegt."
  Public Event ItemMouseLeave(ByVal hItem As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Attribute ItemMouseLeave.VB_Description = "Wird ausgelöst, wenn der Benutzer die Maus aus dem Item <hItem> herausbewegt."
  Public Event ItemMouseMove(ByVal hItem As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Attribute ItemMouseMove.VB_Description = "Wird ausgelöst, wenn der Benutzer über dem Item <hItem> die Maus bewegt."
  Public Event ItemMouseUp(ByVal hItem As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Attribute ItemMouseUp.VB_Description = "Wird ausgelöst, wenn der Benutzer über dem Item <hItem> eine Maustaste losläßt."
  Public Event ItemRClick(ByVal hItem As Long)
Attribute ItemRClick.VB_Description = "Wird ausgelöst, wenn der Benutzer mit der rechten Maustaste auf den Item <hItem> klickt."
  Public Event ItemRDblClick(ByVal hItem As Long)
Attribute ItemRDblClick.VB_Description = "Wird ausgelöst, wenn der Benutzer mit der rechten Maustaste auf den Item <hItem> doppelklickt."
  Public Event ItemRenamed(ByVal hItem As Long, ByVal OldName As String, ByVal NewName As String)
Attribute ItemRenamed.VB_Description = "Wird ausgelöst, nachdem der Item <hItem> umbenannt wurde."
  Public Event ItemShellContextMenuItemEnter(ByVal hItem As Long, ByVal hMenu As Long, ByVal ItemID As Long, ByVal ItemText As String, ByVal ItemVerb As String, ByVal ItemDescription As String)
Attribute ItemShellContextMenuItemEnter.VB_Description = "Wird ausgelöst, wenn die Maus in den Menü-Item <ItemID> des Menüs <hMenu> hineinbewegt wird."
  Public Event ItemShellContextMenuItemLeave(ByVal hItem As Long, ByVal hMenu As Long, ByVal ItemID As Long, ByVal ItemText As String, ByVal ItemVerb As String, ByVal ItemDescription As String)
Attribute ItemShellContextMenuItemLeave.VB_Description = "Wird ausgelöst, wenn die Maus aus dem Menü-Item <ItemID> des Menüs <hMenu> herausbewegt wird."
  Public Event ItemStateIconChanged(ByVal hItem As Long, ByVal OldStateIcon As Long, ByVal NewStateIcon As Long)
Attribute ItemStateIconChanged.VB_Description = "Wird ausgelöst, nachdem das StateIcon des Items <hItem> geändert wurde."
'  Public Event ItemUpdated(ByVal hItem As Long)
'  Public Event ItemUpdating(ByVal hItem As Long)
'  Public Event ItemUpdatedSubItems(ByVal hItem As Long, ByRef DeletedItems() As Long, ByRef AddedItems() As Long, ByRef ChangedItems() As Long)
'  Public Event ItemUpdatingSubItems(ByVal hItem As Long, DontDeleteOld As Boolean, DontAddNew As Boolean, DontChangeExisting As Boolean)
  Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Wird ausgelöst, wenn der Benutzer eine Taste drückt während das Control den Fokus besitzt."
Attribute KeyDown.VB_UserMemId = -602
  Public Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Wird ausgelöst, wenn der Benutzer eine ANSI-Taste drückt und losläßt während das Control den Fokus besitzt."
Attribute KeyPress.VB_UserMemId = -603
  Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Wird ausgelöst, wenn der Benutzer eine Taste losläßt während das Control den Fokus besitzt."
Attribute KeyUp.VB_UserMemId = -604
  Public Event MClick()
Attribute MClick.VB_Description = "Wird ausgelöst, wenn der Benutzer mit der mittleren Maustaste in das Control klickt."
  Public Event MDblClick()
Attribute MDblClick.VB_Description = "Wird ausgelöst, wenn der Benutzer mit der mittleren Maustaste in das Control doppelklickt."
  Public Event MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Attribute MouseDown.VB_Description = "Wird ausgelöst, wenn der Benutzer über dem Control eine Maustaste drückt."
Attribute MouseDown.VB_UserMemId = -605
  Public Event MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Attribute MouseEnter.VB_Description = "Wird ausgelöst, wenn der Benutzer die Maus in das Control hineinbewegt."
  Public Event MouseHover(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Attribute MouseHover.VB_Description = "Wird ausgelöst, wenn sich der Mauszeiger für <HoverTime> Millisekunden über dem Control befindet."
  Public Event MouseLeave(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Attribute MouseLeave.VB_Description = "Wird ausgelöst, wenn der Benutzer die Maus aus dem Control herausbewegt."
  Public Event MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Attribute MouseMove.VB_Description = "Wird ausgelöst, wenn der Benutzer über dem Control die Maus bewegt."
Attribute MouseMove.VB_UserMemId = -606
  Public Event MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Attribute MouseUp.VB_Description = "Wird ausgelöst, wenn der Benutzer über dem Control eine Maustaste losläßt."
Attribute MouseUp.VB_UserMemId = -607
  Public Event RClick()
Attribute RClick.VB_Description = "Wird ausgelöst, wenn der Benutzer mit der rechten Maustaste in das Control klickt."
  Public Event RDblClick()
Attribute RDblClick.VB_Description = "Wird ausgelöst, wenn der Benutzer mit der rechten Maustaste in das Control doppelklickt."
  Public Event SelChanged(ByVal hOldItem As Long, ByVal hNewItem As Long, ByVal CausedBy As CausedByConstants)
Attribute SelChanged.VB_Description = "Wird ausgelöst, wenn ein anderer Item markiert wurde."
Attribute SelChanged.VB_MemberFlags = "200"
  Public Event SelChanging(ByVal hOldItem As Long, ByVal hNewItem As Long, ByVal CausedBy As CausedByConstants, Cancel As Boolean)
Attribute SelChanging.VB_Description = "Wird ausgelöst, bevor ein anderer Item markiert wird. Mit <Cancel> kann der Vorgang abgebrochen werden."


' Implementation von IVBDropSource

Friend Function IVBDropSource_GiveFeedback(Effects As DROPEFFECTConstants) As Long
  IVBDropSource_GiveFeedback = DRAGDROP_S_USEDEFAULTCURSORS

'  Dim iCursor As Long
'  Dim pt As POINT
'  Dim ptHotSpot As POINT
'
'  If hImlDragCursors Then
'    If effects = 0 Then
'      iCursor = dragCursorIndex_No
'      pt = dragCursorHotSpot_No
'    ElseIf effects = DROPEFFECTConstants.DROPEFFECT_COPY Then
'      iCursor = dragCursorIndex_Copy
'      pt = dragCursorHotSpot_Copy
'    ElseIf effects = DROPEFFECTConstants.DROPEFFECT_LINK Then
'      iCursor = dragCursorIndex_Link
'      pt = dragCursorHotSpot_Link
'    ElseIf effects = DROPEFFECTConstants.DROPEFFECT_MOVE Then
'      iCursor = dragCursorIndex_Move
'      pt = dragCursorHotSpot_Move
'    Else
'      iCursor = dragCursorIndex_Move
'      pt = dragCursorHotSpot_Move
'    End If

'    DAD_ShowCursor False
'    ImageList_GetDragImage 0, VarPtr(ptHotSpot)
'    ptHotSpot.x = ptHotSpot.x - pt.x
'    ptHotSpot.y = ptHotSpot.y - pt.y
'    If ptHotSpot.x < 0 Then ptHotSpot.x = 0
'    If ptHotSpot.y < 0 Then ptHotSpot.y = 0
'    ImageList_SetDragCursorImage hImlDragCursors, iCursor, ptHotSpot.x, ptHotSpot.y
'    IVBDropSource_GiveFeedback = S_OK
'  Else
'    IVBDropSource_GiveFeedback = DRAGDROP_S_USEDEFAULTCURSORS
'  End If
End Function

Friend Function IVBDropSource_QueryContinueDrag(ByVal PressedEscape As Long, ByVal KeyState As Long) As Long
  Const IMAGE_CURSOR = 2
  Const LR_DEFAULTCOLOR = &H0
  Const LR_SHARED = &H8000
  Const OCR_NORMAL = 32512
  Const SM_CXCURSOR = 13
  Const SM_CYCURSOR = 14
  Dim cx As Long
  Dim cy As Long
  Dim hCursor As Long
  Dim ptHotSpot As POINT
  Dim resetCursor As Boolean

  If PressedEscape Then
    IVBDropSource_QueryContinueDrag = DRAGDROP_S_CANCEL
    resetCursor = True
  ElseIf KeyState And dragButton Then
    IVBDropSource_QueryContinueDrag = S_OK
  Else
    IVBDropSource_QueryContinueDrag = DRAGDROP_S_DROP
    resetCursor = True
  End If

'  If resetCursor Then
'    hCursor = LoadImage(0, MAKEINTRESOURCE(OCR_NORMAL), IMAGE_CURSOR, cx, cy, LR_DEFAULTCOLOR Or LR_SHARED)
'    SetCursor hCursor
'    If hImlDragCursors Then
'      DAD_ShowCursor True
'      ImageList_GetDragImage 0, VarPtr(ptHotSpot)
'      ptHotSpot.x = ptHotSpot.x - dragCursorHotSpot_Null.x
'      ptHotSpot.y = ptHotSpot.y - dragCursorHotSpot_Null.y
'      If ptHotSpot.x < 0 Then ptHotSpot.x = 0
'      If ptHotSpot.y < 0 Then ptHotSpot.y = 0
'      ImageList_SetDragCursorImage hImlDragCursors, dragCursorIndex_Null, ptHotSpot.x, ptHotSpot.y
'    End If
'  End If
End Function


' Implementation von IVBDropTarget

Private Sub IVBDropTarget_DragEnter(ByVal Data As DataObject, ByVal KeyState As Long, ByVal x As Long, ByVal y As Long, Effects As ExTvwInterfaces.DROPEFFECTConstants)
  Dim Button As MouseButtonConstants
  Dim hItem As Long
  Dim IDataObj As IVBDataObject
  Dim retEffect As DROPEFFECTConstants
  Dim Shift As ShiftConstants

  curEffects = Effects
  curKeyState = KeyState
  With curPt
    .x = x
    .y = y
  End With
  Set DataObj = Data

  hItem = ItemHitTest(curPt, TVHT_OnItemLine)
  GetShiftAndMouseButtonBits KeyState, Shift, Button
  RaiseDragEnter Data, curEffects, Button, Shift, curPt.x, curPt.y, hItem

  hDropTargetItem = hItem

  retEffect = curEffects
  CopyMemory VarPtr(IDataObj), ObjPtr(Data) + 16, 4
  If Not (IDataObj Is Nothing) Then
    IDataObj.AddRef
  End If
  With curPt
    .x = x
    .y = y
  End With
  ClientToScreen prophWnd, curPt
  If Not (IDropTargetHelper Is Nothing) Then
    If ver_Win_Vista Then
      IDropTargetHelper.DragEnter prophWndShellUIParentWindow, IDataObj, curPt, retEffect
    Else
      IDropTargetHelper.DragEnter prophWnd, IDataObj, curPt, retEffect
    End If
  End If
  Set IDataObj = Nothing

  Effects = retEffect
End Sub

Private Sub IVBDropTarget_DragLeave()
  Dim Button As MouseButtonConstants
  Dim hItem As Long
  Dim Shift As ShiftConstants

  KillTimer prophWnd, TimerID_DragExpand
  KillTimer prophWnd, TimerID_DragScroll

  hItem = prophDropTargetItem
  hDropTargetItem = 0

  GetShiftAndMouseButtonBits 0, Shift, Button, True
  RaiseDragLeave hItem

  curEffects = 0
  curKeyState = 0
  With curPt
    .x = 0
    .y = 0
  End With
  Set IDrpTrgt = Nothing

  If Not (IDropTargetHelper Is Nothing) Then
    IDropTargetHelper.DragLeave
  End If
End Sub

Private Sub IVBDropTarget_DragOver(ByVal KeyState As Long, ByVal x As Long, ByVal y As Long, Effects As ExTvwInterfaces.DROPEFFECTConstants)
  Dim Button As MouseButtonConstants
  Dim hItem As Long
  Dim retEffect As DROPEFFECTConstants
  Dim Shift As ShiftConstants

  curEffects = Effects
  retEffect = curEffects
  curKeyState = KeyState
  With curPt
    .x = x
    .y = y
  End With

  hItem = ItemHitTest(curPt, TVHT_OnItemLine)

  If ItemExists(hItem, AllowNullConstants.anNever) Then
    If hItem = prophDropTargetItem Then
      ' ein DragOver simulieren
      If Not (IDrpTrgt Is Nothing) Then
        On Error Resume Next
        IDrpTrgt.DragOver KeyState, x, y, curEffects
      Else
        curEffects = DROPEFFECTConstants.DROPEFFECT_NONE
      End If
    Else
      hDropTargetItem = hItem
    End If
  Else
    hDropTargetItem = hItem
  End If

  GetShiftAndMouseButtonBits KeyState, Shift, Button
  RaiseDragMove DataObj, curEffects, Button, Shift, curPt.x, curPt.y, hItem

  retEffect = curEffects
  With curPt
    .x = x
    .y = y
  End With
  ClientToScreen prophWnd, curPt
  If Not (IDropTargetHelper Is Nothing) Then
    IDropTargetHelper.DragOver curPt, retEffect
  End If

Ende:
  Effects = retEffect
End Sub

Private Sub IVBDropTarget_Drop(ByVal Data As DataObject, ByVal KeyState As Long, ByVal x As Long, ByVal y As Long, Effects As ExTvwInterfaces.DROPEFFECTConstants)
'  Dim arraySize As Long
  Dim Button As MouseButtonConstants
'  Dim dataFmt As FORMATETC
'  Dim EffectUsed As Long
  Dim hItem As Long
  Dim IDataObj As IVBDataObject
'  Dim pBinaryData As Long
  Dim retEffect As DROPEFFECTConstants
  Dim Shift As ShiftConstants
'  Dim stgData As STGMEDIUM

  KillTimer prophWnd, TimerID_DragExpand
  KillTimer prophWnd, TimerID_DragScroll

  curEffects = Effects
  curKeyState = KeyState
  With curPt
    .x = x
    .y = y
  End With
  Set DataObj = Data

  hItem = ItemHitTest(curPt, TVHT_OnItemLine)
  GetShiftAndMouseButtonBits KeyState, Shift, Button
  RaiseDrop Data, curEffects, Button, Shift, curPt.x, curPt.y, hItem

  retEffect = curEffects
  CopyMemory VarPtr(IDataObj), ObjPtr(Data) + 16, 4
  If Not (IDataObj Is Nothing) Then
    IDataObj.AddRef
  End If
  With curPt
    .x = x
    .y = y
  End With
  ClientToScreen prophWnd, curPt
  If Not (IDropTargetHelper Is Nothing) Then
    IDropTargetHelper.Drop IDataObj, curPt, retEffect
  End If

  If ItemExists(hItem, AllowNullConstants.anNever) Then
    ' ein Drop simulieren
    curEffects = retEffect
    hDropTargetItem = hItem
    retEffect = curEffects
    If Not (IDrpTrgt Is Nothing) Then
      On Error Resume Next
      IDrpTrgt.Drop IDataObj, curKeyState, curPt.x, curPt.y, retEffect
    End If
  End If
'
'  With dataFmt
'    .cfFormat = LoWord(RegisterClipboardFormat("Performed DropEffect"))
'    .dwAspect = DVASPECTConstants.DVASPECT_CONTENT
'    .lindex = -1
'  End With
'  If IDataObj.GetData(dataFmt, stgData) = S_OK Then
'    arraySize = GlobalSize(stgData.Data)
'    pBinaryData = GlobalLock(stgData.Data)
'    CopyMemory VarPtr(EffectUsed), pBinaryData, IIf(arraySize < LenB(EffectUsed), arraySize, LenB(EffectUsed))
'    GlobalUnlock stgData.Data
'    ReleaseStgMedium VarPtr(stgData)
'  End If

  hDropTargetItem = 0
  curEffects = 0
  curKeyState = 0
  With curPt
    .x = 0
    .y = 0
  End With
  Set IDrpTrgt = Nothing
  Set IDataObj = Nothing

  Effects = retEffect
End Sub


' Implementation von IHook

Private Function IHook_CallWndProcAfter(ByVal HookCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  '
End Function

Private Function IHook_CallWndProcBefore(ByVal HookCode As Long, ByVal wParam As Long, ByVal lParam As Long, eatIt As Boolean) As Long
  '
End Function

Private Function IHook_CBTProcAfter(ByVal HookCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Const HCBT_CREATEWND = 3
  Const HCBT_DESTROYWND = 4
  Dim ClsName As String
  Dim tmp As Long

  Select Case HookCode
    Case HCBT_CREATEWND
      ClsName = Space(MAX_PATH)
      GetClassName wParam, ClsName, Len(ClsName)
      ClsName = removeNull(ClsName)
      Select Case LCase(ClsName)
        Case "edit"
          ' eine TextBox
          SetTimer prophWnd, TimerID_SubClassEdit, 100, 0
        Case "#32768"
          ' ein Menü
          tmp = GetTickCount
          If tmp - TimeOfLastWM_CONTEXTMENU_Edit <= 500 Then
            hWnd_EditContextMenu = wParam
            blockMouseLeave = True
            blockMouseLeave_Edit = True
          ElseIf tmp - TimeOfLastWM_CONTEXTMENU_Shell <= 500 Then
            hWnd_ShellContextMenu = wParam
            blockMouseLeave = True
          End If
      End Select
    Case HCBT_DESTROYWND
      Select Case wParam
        Case hWnd_EditContextMenu
          hWnd_EditContextMenu = 0
          blockMouseLeave = False
          blockMouseLeave_Edit = False
        Case hWnd_ShellContextMenu
          hWnd_ShellContextMenu = 0
          blockMouseLeave = False
      End Select
  End Select
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


' Implementation von ISort

' wird beim Sortieren von Archiv-Items aufgerufen
Private Function ISort_CBSortArchiveItems(ByVal Data1 As Long, ByVal Data2 As Long, Optional ByVal ZeroIfSameType As Boolean = False) As Long
  Dim DispName1 As String
  Dim DispName2 As String
  Dim ret As Long
  Dim TVWItemInfo1 As ITEMINFO_ARCHIVE
  Dim TVWItemInfo2 As ITEMINFO_ARCHIVE

  CopyMemory VarPtr(TVWItemInfo1), Data1, LenB(TVWItemInfo1)
  CopyMemory VarPtr(TVWItemInfo2), Data2, LenB(TVWItemInfo2)

  If TVWItemInfo1.itemType = TVWItemInfo2.itemType Then
    If Not ZeroIfSameType Then
      ' selber Typ -> alphabetisch vorgehen
      DispName1 = GetItemText(TVWItemInfo1.hItem)
      DispName2 = GetItemText(TVWItemInfo2.hItem)
      ret = StrComp(DispName1, DispName2, VbCompareMethod.vbTextCompare)
    End If
  Else
    ' verschiedene Typen -> Ordner zuerst
    If TVWItemInfo1.itemType = ItemTypeConstants.itArchiveFolder Then
      ' Item1 soll zuerst stehen
      ret = -1
    Else
      ' Item2 soll zuerst stehen
      ret = 1
    End If
  End If

  If propSortOrder = SortOrderConstants.sorDescending Then ret = -ret

  ISort_CBSortArchiveItems = ret
End Function

' wird beim Sortieren von "normalen" Items aufgerufen
Private Function ISort_CBSortCommonItems(ByVal Data1 As Long, ByVal Data2 As Long) As Long
'  Dim flags As SHCIDSConstants
  Dim IParent1 As IVBShellFolder
'  Dim IParent2 As IVBShellFolder
  Dim pIDLToParent1 As Long
  Dim pIDLToParent2 As Long
  Dim ret As Long
  Dim tmp As Boolean
  Dim TVWItemInfo1 As ITEMINFO_COMMON
  Dim TVWItemInfo2 As ITEMINFO_COMMON

  ' Daten der beiden Items ermitteln
  CopyMemory VarPtr(TVWItemInfo1), Data1, LenB(TVWItemInfo1)
  CopyMemory VarPtr(TVWItemInfo2), Data2, LenB(TVWItemInfo2)
'  #If Debuging Then
'    SplitFullyQualifiedPIDL debugger, TVWItemInfo1.pIDLToDesktop, IParent1, pIDLToParent1
'    SplitFullyQualifiedPIDL debugger, TVWItemInfo2.pIDLToDesktop, IParent2, pIDLToParent2
'  #Else
'    SplitFullyQualifiedPIDL TVWItemInfo1.pIDLToDesktop, IParent1, pIDLToParent1
'    SplitFullyQualifiedPIDL TVWItemInfo2.pIDLToDesktop, IParent2, pIDLToParent2
'  #End If
'
'  ret = -2
'  If propArchivesAfterFolders Then
'    #If Debuging Then
'      tmp = IsArchiveToExpand(debugger, IParent1, pIDLToParent1, propExpandArchives)
'      If tmp Xor IsArchiveToExpand(debugger, IParent2, pIDLToParent2, propExpandArchives) Then
'    #Else
'      tmp = IsArchiveToExpand(IParent1, pIDLToParent1, propExpandArchives)
'      If tmp Xor IsArchiveToExpand(IParent2, pIDLToParent2, propExpandArchives) Then
'    #End If
'      If tmp Then
'        If TVWItemInfo2.itemType = ItemTypeConstants.itFSFile Then ret = -1
'      Else
'        If TVWItemInfo1.itemType = ItemTypeConstants.itFSFile Then ret = 1
'      End If
'    End If
'  End If
'
'  If ret = -2 Then
'    flags = propSortingCriteria
'    ' <IParent> sollte für beide pIDLs gleich sein
'    ret = IParent1.CompareIDs(flags, pIDLToParent1, pIDLToParent2)
'  End If
'  Set IParent1 = Nothing
'  Set IParent2 = Nothing
'
'  If propSortOrder = SortOrderConstants.sorDescending Then ret = -ret
'  ISort_CBSortCommonItems = ret

  #If Debuging Then
    SplitFullyQualifiedPIDL debugger, TVWItemInfo1.pIDLToDesktop, IParent1, pIDLToParent1
  #Else
    SplitFullyQualifiedPIDL TVWItemInfo1.pIDLToDesktop, IParent1, pIDLToParent1
  #End If
  If Not (IParent1 Is Nothing) Then
    'pIDLToParent1 = ILFindLastID(TVWItemInfo1.pIDLToDesktop)
    pIDLToParent2 = ILFindLastID(TVWItemInfo2.pIDLToDesktop)

    ret = -2
    If propArchivesAfterFolders Then
      #If Debuging Then
        tmp = IsArchiveToExpand(debugger, IParent1, pIDLToParent1, propExpandArchives)
        If tmp Xor IsArchiveToExpand(debugger, IParent1, pIDLToParent2, propExpandArchives) Then
      #Else
        tmp = IsArchiveToExpand(IParent1, pIDLToParent1, propExpandArchives)
        If tmp Xor IsArchiveToExpand(IParent1, pIDLToParent2, propExpandArchives) Then
      #End If
        If tmp Then
          If TVWItemInfo2.itemType = ItemTypeConstants.itFSFile Then
            ret = -1
          End If
        Else
          If TVWItemInfo1.itemType = ItemTypeConstants.itFSFile Then
            ret = 1
          End If
        End If
      End If
    End If

    If ret = -2 Then
      ' <IParent> sollte für beide pIDLs gleich sein
      ret = IParent1.CompareIDs(propSortingCriteria, pIDLToParent1, pIDLToParent2)
    End If
    Set IParent1 = Nothing
  End If

  If propSortOrder = SortOrderConstants.sorDescending Then ret = -ret
  ISort_CBSortCommonItems = ret
End Function

Private Function ISort_CBSortGroups(ByVal Data1 As Long, ByVal Data2 As Long) As Long
  '
End Function


' Implementation von ISubClass

Private Function ISubClass_WinMain(ByVal iListener As Long, ByVal oldWinMain As Long, ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Select Case hWnd
    Case UserControl.hWnd
      ISubClass_WinMain = WinMainUC(iListener, oldWinMain, hWnd, MSG, wParam, lParam)
    Case prophWnd
      ISubClass_WinMain = WinMainTree(iListener, oldWinMain, hWnd, MSG, wParam, lParam)
    Case hWnd_ParentForm
      ISubClass_WinMain = WinMainParent(iListener, oldWinMain, hWnd, MSG, wParam, lParam)
    Case hWndEdit
      ISubClass_WinMain = WinMainEdit(iListener, oldWinMain, hWnd, MSG, wParam, lParam)
  End Select
End Function


' Events

Private Sub propFont_FontChanged(ByVal PropertyName As String)
  ChangeFont
  PropertyChanged PropertyName
End Sub

Private Sub UserControl_Initialize()
  Const IDB_LINK_MERGE = 142
  Const IDB_PLUS_MERGE = 143
  Const IDC_NULL = 1005
  Const ILC_COLORDDB = &HFE
  Const ILC_MASK = &H1
  Const IMAGE_BITMAP = 0
  Const IMAGE_CURSOR = 2
  Const LR_DEFAULTCOLOR = &H0
  Const LR_SHARED = &H8000
  Const OCR_NO = 32648
  Const OCR_NORMAL = 32512
  Const RASTERCAPS = 38
  Const RC_PALETTE = &H100
  Const SM_CXCURSOR = 13
  Const SM_CYCURSOR = 14
  Dim cursorObj As IPictureDisp
  Dim cx As Long
  Dim cy As Long
  Dim Flags As Long
  Dim hCursor As Long
  Dim hDC As Long
  Dim hMod As Long
  Dim ptHotSpot As POINT
  Dim Tmp1 As IVBUnknown     ' IDropSource ist standardmäßig nicht implementiert
  Dim Tmp2 As IVBOleInPlaceActiveObject

  #If Debuging Then
    Set debugger = New clsDebugger
  #End If

  GetVersions
  GetUUIDs

  SHGetDesktopFolder IDesktop
  pIDL_Desktop = CSIDLTopIDL(CSIDLConstants.CSIDL_DESKTOP)

  disableCreateControl = True
  disableLoadFileSystem = True
  dontExpand = False

  InitCommonCtls CommonControlsConsts.ICC_TREEVIEW_CLASSES

  hMod = LoadLibrary("shell32.dll")
  If hMod Then
    If GetProcAddress(hMod, "SHDoDragDrop") Then
      useSHDoDragDrop = True
    End If
    FreeLibrary hMod
  End If

  If Not useSHDoDragDrop Then
    ' neues DropSource-Interface initiieren
    With IDSData
      Set Tmp1 = Me
      CopyMemory VarPtr(.oldIDS), VarPtr(Tmp1), 4
      .ObjType = AppID
      CopyMemory VarPtr(.ObjTVW), VarPtr(Me), 4
      .pVTable = basDragDrop.IDSVTable
      .pThis = VarPtr(IDSData)
    End With
  End If

  ' neues Fokus-Interface initiieren
  With IPAOData
    Set Tmp2 = Me
    CopyMemory VarPtr(.oldIPAO), VarPtr(Tmp2), 4
    .ObjType = AppID
    CopyMemory VarPtr(.ObjTVW), VarPtr(Me), 4
    .pVTable = basFokus.IPAOVTable
    .pThis = VarPtr(IPAOData)
  End With

  DI_GETDRAGIMAGE = RegisterWindowMessage("ShellGetDragImage")

'  Flags = ILC_MASK
'  hDC = GetDC(0)
'  If (GetDeviceCaps(hDC, RASTERCAPS) And RC_PALETTE) = 0 Then
'    Flags = ILC_COLORDDB
'  End If
'  ReleaseDC 0, hDC
'
'  cx = GetSystemMetrics(SM_CXCURSOR)
'  cy = GetSystemMetrics(SM_CYCURSOR)
'  hImlDragCursors = ImageList_Create(cx, cy, Flags, 1, 0)
'  If hImlDragCursors Then
'    hMod = LoadLibrary("shell32.dll")
'    If hMod Then
'      hCursor = LoadImage(hMod, MAKEINTRESOURCE(IDC_NULL), IMAGE_CURSOR, cx, cy, LR_DEFAULTCOLOR Or LR_SHARED)
'      If hCursor = 0 Then
'        Debug.Print "IDC_NULL could not be loaded!"
'        Set cursorObj = LoadResPicture(IDC_NULL, LoadResConstants.vbResCursor)
'        hCursor = cursorObj.Handle
'      End If
'      dragCursorIndex_Null = AddCursorToImageList(hImlDragCursors, hCursor, -1, dragCursorHotSpot_Null)
'
'      hCursor = LoadImage(0, MAKEINTRESOURCE(OCR_NO), IMAGE_CURSOR, cx, cy, LR_DEFAULTCOLOR Or LR_SHARED)
'      dragCursorIndex_No = AddCursorToImageList(hImlDragCursors, hCursor, -1, dragCursorHotSpot_No)
'
'      hCursor = LoadImage(0, MAKEINTRESOURCE(OCR_NORMAL), IMAGE_CURSOR, cx, cy, LR_DEFAULTCOLOR Or LR_SHARED)
'      dragCursorIndex_Move = AddCursorToImageList(hImlDragCursors, hCursor, -1, dragCursorHotSpot_Move)
'
'      hCursor = LoadImage(0, MAKEINTRESOURCE(OCR_NORMAL), IMAGE_CURSOR, cx, cy, LR_DEFAULTCOLOR Or LR_SHARED)
'      dragCursorIndex_Copy = AddCursorToImageList(hImlDragCursors, hCursor, MAKEINTRESOURCE(IDB_PLUS_MERGE), dragCursorHotSpot_Copy)
'
'      hCursor = LoadImage(0, MAKEINTRESOURCE(OCR_NORMAL), IMAGE_CURSOR, cx, cy, LR_DEFAULTCOLOR Or LR_SHARED)
'      dragCursorIndex_Link = AddCursorToImageList(hImlDragCursors, hCursor, MAKEINTRESOURCE(IDB_LINK_MERGE), dragCursorHotSpot_Link)
'
'      FreeLibrary hMod
'    End If
'  End If
  #If Debuging Then
    TestBed
  #End If
End Sub

Private Sub UserControl_InitProperties()
  disableCreateControl = True
  disableLoadFileSystem = True

  AllowDragDrop = True
  AlwaysUseSmallIcons = False
  Appearance = AppearanceConstants.a3D
  ArchiveItemBackColor = -1
  ArchiveItemForeColor = -1
  ArchivesAfterFolders = True
  AutoHScroll = True
  BackColor = SystemColorConstants.vbWindowBackground
  #If BkImageSupport Then
    Set BkImage = Nothing
    BkImageFixed = True
    BkImageStyle = BkImageStyleConstants.bisTiled
  #End If
  BorderStyle = BorderStyleConstants.bsNone
  CheckBoxes = False
  CheckSlowItemsForSubItems = False
  CompressedBackColor = -1
  CompressedForeColor = -1
  DimSelectedItemsIcon = False
  DontCheckFolderExpandibility = False
  DragExpandTime = 2000
  DragScrollTime = 500
  DragScrollZoneBottom = 16
  DragScrollZoneLeft = 16
  DragScrollZoneRight = 16
  DragScrollZoneTop = 16
  DragShowAlphaBlendedItems = True
  DrivesOnly = False
  Enabled = True
  EncryptedBackColor = -1
  EncryptedForeColor = -1
  ExpandArchives = 0 'ExpandArchivesConstants.eaACE Or ExpandArchivesConstants.eaBIN Or ExpandArchivesConstants.eaCAB Or ExpandArchivesConstants.eaISO Or ExpandArchivesConstants.eaRAR Or ExpandArchivesConstants.eaZIP
  FadeExpandos = False
  FastInfoTipsOnly = False
  FavoritesStyle = False
  FileAttributes = AttributesConstants.attArchive Or AttributesConstants.attCompressed Or AttributesConstants.attEncrypted Or AttributesConstants.attReadOnly
  FileFilters = ""
  FireBeforeInsertItem = 0
  FolderAttributes = AttributesConstants.attArchive Or AttributesConstants.attCompressed Or AttributesConstants.attEncrypted Or AttributesConstants.attReadOnly
  FolderFilters = ""
  Set Font = Ambient.Font
  ForeColor = SystemColorConstants.vbWindowText
  FullRowSelect = False
  HotBackColor = -1
  HotForeColor = -1
  HotTracking = False
  HoverTime = 1
  IncludedItems = IncludedItemsConstants.iiFSFolders Or IncludedItemsConstants.iiNonFSFolders
  Indent = 16
  ItemHeight = 17
  LineColor = SystemColorConstants.vbButtonShadow
  LineStyle = LineStyleConstants.lsLinesAtItem
  Margin = 3
  MaxScrollTime = 100
  Set MouseIcon = Nothing
  MousePointer = MousePointerConstants.vbDefault
  path = "Desktop"
  ProcessContextMenuKeys = True
  Renaming = RenamingConstants.rAutomatic
  ReplaceHandCursor = False
  RootPath = "Desktop"
  RTLReading = Ambient.RightToLeft
  ScrollBars = ScrollBarsConstants.sbNormal
  SelChangedDelayTime = 0
  SelectedItemBackColor = -1
  SelectedItemForeColor = -1
  ShowFocusRect = True
  ShowHiddenDimmed = True
  ShowIcons = True
  ShownOverlays = ShownOverlaysConstants.soAll
  ShownToolTips = ShownToolTipsConstants.sttHiddenItems
  ShowRoot = True
  ShowSelection = True
  ShowShellContextMenus = True
  SingleExpand = SingleExpandConstants.seNone
  SortingCriteria = SortingCriteriaConstants.scName
  SortOrder = SortOrderConstants.sorAscending
  TreeViewStyle = TreeViewStyleConstants.tvsButtons Or TreeViewStyleConstants.tvsLines
  UseCustomIcons = False
  UseFileFilters = False
  UseFolderFilters = False
  UseSelectedIcons = True
  UseStandardIcons = False
  UseStandardIconsForSlowItems = False
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
  Const MK_ALT = &H20
  Const MK_CONTROL = &H8
  Const MK_LBUTTON = &H1
  Const MK_MBUTTON = &H10
  Const MK_RBUTTON = &H2
  Const MK_SHIFT = &H4
  Dim KeyState As Long

  If Button And MouseButtonConstants.vbLeftButton Then KeyState = KeyState Or MK_LBUTTON
  If Button And MouseButtonConstants.vbMiddleButton Then KeyState = KeyState Or MK_MBUTTON
  If Button And MouseButtonConstants.vbRightButton Then KeyState = KeyState Or MK_RBUTTON
  If Shift And ShiftConstants.vbAltMask Then KeyState = KeyState Or MK_ALT
  If Shift And ShiftConstants.vbCtrlMask Then KeyState = KeyState Or MK_CONTROL
  If Shift And ShiftConstants.vbShiftMask Then KeyState = KeyState Or MK_SHIFT
  IVBDropTarget_Drop Data, KeyState, x, y, Effect
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, state As Integer)
  Const MK_ALT = &H20
  Const MK_CONTROL = &H8
  Const MK_LBUTTON = &H1
  Const MK_MBUTTON = &H10
  Const MK_RBUTTON = &H2
  Const MK_SHIFT = &H4
  Dim KeyState As Long

  If Button And MouseButtonConstants.vbLeftButton Then KeyState = KeyState Or MK_LBUTTON
  If Button And MouseButtonConstants.vbMiddleButton Then KeyState = KeyState Or MK_MBUTTON
  If Button And MouseButtonConstants.vbRightButton Then KeyState = KeyState Or MK_RBUTTON
  If Shift And ShiftConstants.vbAltMask Then KeyState = KeyState Or MK_ALT
  If Shift And ShiftConstants.vbCtrlMask Then KeyState = KeyState Or MK_CONTROL
  If Shift And ShiftConstants.vbShiftMask Then KeyState = KeyState Or MK_SHIFT
  Select Case state
    Case DragOverConstants.vbEnter
      IVBDropTarget_DragEnter Data, KeyState, x, y, Effect
    Case DragOverConstants.vbLeave
      IVBDropTarget_DragLeave
    Case DragOverConstants.vbOver
      IVBDropTarget_DragOver KeyState, x, y, Effect
  End Select
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  #If Debuging Then
    debugger.Prefix = Ambient.DisplayName
    debugger.AddLogEntry "ExplorerTreeView Version " & Me.Version, LogEntryTypeConstants.letOther
    debugger.AddLogEntry "Logfile: " & debugger.LogFileName, LogEntryTypeConstants.letOther
    debugger.AddLogEntry "Important Flags:", LogEntryTypeConstants.letOther
    debugger.AddLogEntry "   ver_Comctl32_580: " & ver_Comctl32_580, LogEntryTypeConstants.letOther
    debugger.AddLogEntry "   ver_Comctl32_600: " & ver_Comctl32_600, LogEntryTypeConstants.letOther
    debugger.AddLogEntry "   ver_Comctl32_610: " & ver_Comctl32_610, LogEntryTypeConstants.letOther
    debugger.AddLogEntry "   ver_Shell32_50: " & ver_Shell32_50, LogEntryTypeConstants.letOther
    debugger.AddLogEntry "   ver_Shlwapi_50: " & ver_Shlwapi_50, LogEntryTypeConstants.letOther
    debugger.AddLogEntry "   ver_Win_98OrNT4SP3: " & ver_Win_98OrNT4SP3, LogEntryTypeConstants.letOther
    debugger.AddLogEntry "   ver_Win_98OrNT4SP4: " & ver_Win_98OrNT4SP4, LogEntryTypeConstants.letOther
    debugger.AddLogEntry "   ver_Win_2k: " & ver_Win_2k, LogEntryTypeConstants.letOther
    debugger.AddLogEntry "   ver_Win_MEOrXP: " & ver_Win_MEOrXP, LogEntryTypeConstants.letOther
    debugger.AddLogEntry "   ver_Win_NTBased: " & ver_Win_NTBased, LogEntryTypeConstants.letOther
    debugger.AddLogEntry "   ver_Win_XP: " & ver_Win_XP, LogEntryTypeConstants.letOther
    debugger.AddLogEntry "   ver_Win_Vista: " & ver_Win_Vista, LogEntryTypeConstants.letOther
    debugger.AddLogEntry "   ver_Win_Seven: " & ver_Win_Seven & vbNewLine, LogEntryTypeConstants.letOther
  #End If

  disableCreateControl = True
  disableLoadFileSystem = True

  With PropBag
    AllowDragDrop = .ReadProperty("AllowDragDrop", True)
    AlwaysUseSmallIcons = .ReadProperty("AlwaysUseSmallIcons", False)
    Appearance = .ReadProperty("Appearance", AppearanceConstants.a3D)
    ArchiveItemBackColor = .ReadProperty("ArchiveItemBackColor", -1)
    ArchiveItemForeColor = .ReadProperty("ArchiveItemForeColor", -1)
    ArchivesAfterFolders = .ReadProperty("ArchivesAfterFolders", True)
    AutoHScroll = .ReadProperty("AutoHScroll", True)
    BackColor = .ReadProperty("BackColor", SystemColorConstants.vbWindowBackground)
    #If BkImageSupport Then
      Set BkImage = .ReadProperty("BkImage", Nothing)
      BkImageFixed = .ReadProperty("BkImageFixed", True)
      BkImageStyle = .ReadProperty("BkImageStyle", BkImageStyleConstants.bisTiled)
    #End If
    BorderStyle = .ReadProperty("BorderStyle", BorderStyleConstants.bsNone)
    CheckBoxes = .ReadProperty("CheckBoxes", False)
    CheckSlowItemsForSubItems = .ReadProperty("CheckSlowItemsForSubItems", False)
    CompressedBackColor = .ReadProperty("CompressedBackColor", -1)
    CompressedForeColor = .ReadProperty("CompressedForeColor", -1)
    DimSelectedItemsIcon = .ReadProperty("DimSelectedItemsIcon", False)
    DontCheckFolderExpandibility = .ReadProperty("DontCheckFolderExpandibility", False)
    DragExpandTime = .ReadProperty("DragExpandTime", 2000)
    DragScrollTime = .ReadProperty("DragScrollTime", 500)
    DragScrollZoneBottom = .ReadProperty("DragScrollZoneBottom", 16)
    DragScrollZoneLeft = .ReadProperty("DragScrollZoneLeft", 16)
    DragScrollZoneRight = .ReadProperty("DragScrollZoneRight", 16)
    DragScrollZoneTop = .ReadProperty("DragScrollZoneTop", 16)
    DragShowAlphaBlendedItems = .ReadProperty("DragShowAlphaBlendedItems", True)
    DrivesOnly = .ReadProperty("DrivesOnly", False)
    Enabled = .ReadProperty("Enabled", True)
    EncryptedBackColor = .ReadProperty("EncryptedBackColor", -1)
    EncryptedForeColor = .ReadProperty("EncryptedForeColor", -1)
    ExpandArchives = .ReadProperty("ExpandArchives", 0) 'ExpandArchivesConstants.eaACE Or ExpandArchivesConstants.eaBIN Or ExpandArchivesConstants.eaCAB Or ExpandArchivesConstants.eaISO Or ExpandArchivesConstants.eaRAR Or ExpandArchivesConstants.eaZIP)
    FadeExpandos = .ReadProperty("FadeExpandos", False)
    FastInfoTipsOnly = .ReadProperty("FastInfoTipsOnly", False)
    FavoritesStyle = .ReadProperty("FavoritesStyle", False)
    FileAttributes = .ReadProperty("FileAttributes", AttributesConstants.attArchive Or AttributesConstants.attCompressed Or AttributesConstants.attEncrypted Or AttributesConstants.attReadOnly)
    FileFilters = .ReadProperty("FileFilters", "")
    FireBeforeInsertItem = .ReadProperty("FireBeforeInsertItem", 0)
    FolderAttributes = .ReadProperty("FolderAttributes", AttributesConstants.attArchive Or AttributesConstants.attCompressed Or AttributesConstants.attEncrypted Or AttributesConstants.attReadOnly)
    FolderFilters = .ReadProperty("FolderFilters", "")
    Set Font = .ReadProperty("Font", Ambient.Font)
    ForeColor = .ReadProperty("ForeColor", SystemColorConstants.vbWindowText)
    FullRowSelect = .ReadProperty("FullRowSelect", False)
    HotBackColor = .ReadProperty("HotBackColor", -1)
    HotForeColor = .ReadProperty("HotForeColor", -1)
    HotTracking = .ReadProperty("HotTracking", False)
    HoverTime = .ReadProperty("HoverTime", 1)
    IncludedItems = .ReadProperty("IncludedItems", IncludedItemsConstants.iiFSFolders Or IncludedItemsConstants.iiNonFSFolders)
    Indent = .ReadProperty("Indent", 16)
    ItemHeight = .ReadProperty("ItemHeight", 16)
    LimitEditInput = .ReadProperty("LimitEditInput", True)
    LineColor = .ReadProperty("LineColor", SystemColorConstants.vbButtonShadow)
    LineStyle = .ReadProperty("LineStyle", LineStyleConstants.lsLinesAtItem)
    Margin = .ReadProperty("Margin", 3)
    MaxScrollTime = .ReadProperty("MaxScrollTime", 100)
    Set MouseIcon = .ReadProperty("MouseIcon", Nothing)
    MousePointer = .ReadProperty("MousePointer", MousePointerConstants.vbDefault)
    path = .ReadProperty("Path", "Desktop")
    ProcessContextMenuKeys = .ReadProperty("ProcessContextMenuKeys", True)
    Renaming = .ReadProperty("Renaming", RenamingConstants.rAutomatic)
    ReplaceHandCursor = .ReadProperty("ReplaceHandCursor", False)
    RootPath = .ReadProperty("RootPath", "Desktop")
    If propPath = "" Then path = RootPath
    RTLReading = .ReadProperty("RTLReading", Ambient.RightToLeft)
    ScrollBars = Abs(.ReadProperty("ScrollBars", ScrollBarsConstants.sbNormal))
    SelChangedDelayTime = .ReadProperty("SelChangedDelayTime", 0)
    SelectedItemBackColor = .ReadProperty("SelectedItemBackColor", -1)
    SelectedItemForeColor = .ReadProperty("SelectedItemForeColor", -1)
    ShowFocusRect = .ReadProperty("ShowFocusRect", True)
    ShowHiddenDimmed = .ReadProperty("ShowHiddenDimmed", True)
    ShowIcons = .ReadProperty("ShowIcons", True)
    ShownOverlays = .ReadProperty("ShownOverlays", ShownOverlaysConstants.soAll)
    ShownToolTips = .ReadProperty("ShownToolTips", ShownToolTipsConstants.sttHiddenItems)
    ShowRoot = .ReadProperty("ShowRoot", True)
    ShowSelection = .ReadProperty("ShowSelection", True)
    ShowShellContextMenus = .ReadProperty("ShowShellContextMenus", True)
    SingleExpand = .ReadProperty("SingleExpand", SingleExpandConstants.seNone)
    SortingCriteria = .ReadProperty("SortingCriteria", SortingCriteriaConstants.scName)
    SortOrder = .ReadProperty("SortOrder", SortOrderConstants.sorAscending)
    TreeViewStyle = .ReadProperty("TreeViewStyle", TreeViewStyleConstants.tvsButtons Or TreeViewStyleConstants.tvsLines)
    UseCustomIcons = .ReadProperty("UseCustomIcons", False)
    UseFileFilters = .ReadProperty("UseFileFilters", False)
    UseFolderFilters = .ReadProperty("UseFolderFilters", False)
    UseSelectedIcons = .ReadProperty("UseSelectedIcons", True)
    UseStandardIcons = .ReadProperty("UseStandardIcons", False)
    UseStandardIconsForSlowItems = .ReadProperty("UseStandardIconsForSlowItems", False)
  End With
End Sub

Private Sub UserControl_Resize()
  Dim sclMode As Long

  ' work-around a bug that VB has with non-default DPI settings
  ' http://groups.google.de/group/microsoft.public.de.vb/browse_thread/thread/61e4b1fd7d747314/d7315395726410fe
  On Error Resume Next
  sclMode = vbTwips
  sclMode = Extender.Container.ScaleMode
  SetWindowPos UserControl.hWnd, 0, 0, 0, ScaleX(Extender.Width, sclMode, vbPixels), ScaleY(Extender.Height, sclMode, vbPixels), SWP_NOMOVE Or SWP_NOZORDER
  If Err Then Err.Clear
  On Error GoTo 0

  If prophWnd Then
    ' das Control neu resizen
    SetWindowPos prophWnd, 0, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, SWP_NOMOVE Or SWP_NOZORDER
  Else
    ' TreeView erzeugen
    disableCreateControl = False
    CreateControl
  End If
End Sub

Private Sub UserControl_Terminate()
  ' TreeView zerstören
  DestroyControl

  If Not useSHDoDragDrop Then
    With IDSData
      ZeroMemory VarPtr(.oldIDS), 4
      ZeroMemory VarPtr(.ObjTVW), 4
    End With
  End If

  With IPAOData
    ZeroMemory VarPtr(.oldIPAO), 4
    ZeroMemory VarPtr(.ObjTVW), 4
  End With
  #If Debuging Then
    FreeItemIDList debugger, "UserControl_Terminate", pIDL_Desktop
  #Else
    FreeItemIDList pIDL_Desktop
  #End If

  #If BkImageSupport Then
    If bkImageBrush Then DeleteObject bkImageBrush
    If bkImageTexture Then DeleteObject bkImageTexture
  #End If

  #If Debuging Then
    Set debugger = Nothing
  #End If

'  If hImlDragCursors Then
'    ImageList_Destroy hImlDragCursors
'    ' just to be sure
'    DAD_ShowCursor True
'  End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  With PropBag
    .WriteProperty "AllowDragDrop", propAllowDragDrop, True
    .WriteProperty "AlwaysUseSmallIcons", propAlwaysUseSmallIcons, False
    .WriteProperty "Appearance", propAppearance, AppearanceConstants.a3D
    .WriteProperty "ArchiveItemBackColor", propArchiveItemBackColor, -1
    .WriteProperty "ArchiveItemForeColor", propArchiveItemForeColor, -1
    .WriteProperty "ArchivesAfterFolders", propArchivesAfterFolders, True
    .WriteProperty "AutoHScroll", propAutoHScroll, True
    .WriteProperty "BackColor", propBackColor, SystemColorConstants.vbWindowBackground
    #If BkImageSupport Then
      .WriteProperty "BkImage", propBkImage, Nothing
      .WriteProperty "BkImageFixed", propBkImageFixed, True
      .WriteProperty "BkImageStyle", propBkImageStyle, BkImageStyleConstants.bisTiled
    #End If
    .WriteProperty "BorderStyle", propBorderStyle, BorderStyleConstants.bsNone
    .WriteProperty "CheckBoxes", propCheckBoxes, False
    .WriteProperty "CheckSlowItemsForSubItems", propCheckSlowItemsForSubItems, False
    .WriteProperty "CompressedBackColor", propCompressedBackColor, -1
    .WriteProperty "CompressedForeColor", propCompressedForeColor, -1
    .WriteProperty "DimSelectedItemsIcon", propDimSelectedItemsIcon, False
    .WriteProperty "DontCheckFolderExpandibility", propDontCheckFolderExpandibility, False
    .WriteProperty "DragExpandTime", propDragExpandTime, 2000
    .WriteProperty "DragScrollTime", propDragScrollTime, 500
    .WriteProperty "DragScrollZoneBottom", propDragScrollZoneBottom, 16
    .WriteProperty "DragScrollZoneLeft", propDragScrollZoneLeft, 16
    .WriteProperty "DragScrollZoneRight", propDragScrollZoneRight, 16
    .WriteProperty "DragScrollZoneTop", propDragScrollZoneTop, 16
    .WriteProperty "DragShowAlphaBlendedItems", propDragShowAlphaBlendedItems, True
    .WriteProperty "DrivesOnly", propDrivesOnly, False
    .WriteProperty "Enabled", propEnabled, True
    .WriteProperty "EncryptedBackColor", propEncryptedBackColor, -1
    .WriteProperty "EncryptedForeColor", propEncryptedForeColor, -1
    .WriteProperty "ExpandArchives", propExpandArchives, 0 'ExpandArchivesConstants.eaACE Or ExpandArchivesConstants.eaBIN Or ExpandArchivesConstants.eaCAB Or ExpandArchivesConstants.eaISO Or ExpandArchivesConstants.eaRAR Or ExpandArchivesConstants.eaZIP
    .WriteProperty "FadeExpandos", propFadeExpandos, False
    .WriteProperty "FastInfoTipsOnly", propFastInfoTipsOnly, False
    .WriteProperty "FavoritesStyle", propFavoritesStyle, False
    .WriteProperty "FileAttributes", propFileAttributes, AttributesConstants.attArchive Or AttributesConstants.attCompressed Or AttributesConstants.attEncrypted Or AttributesConstants.attReadOnly
    .WriteProperty "FileFilters", propFileFilters, ""
    .WriteProperty "FireBeforeInsertItem", propFireBeforeInsertItem, 0
    .WriteProperty "FolderAttributes", propFolderAttributes, AttributesConstants.attArchive Or AttributesConstants.attCompressed Or AttributesConstants.attEncrypted Or AttributesConstants.attReadOnly
    .WriteProperty "FolderFilters", propFolderFilters, ""
    .WriteProperty "Font", propFont, Ambient.Font
    .WriteProperty "ForeColor", propForeColor, SystemColorConstants.vbWindowText
    .WriteProperty "FullRowSelect", propFullRowSelect, False
    .WriteProperty "HotBackColor", propHotBackColor, -1
    .WriteProperty "HotForeColor", propHotForeColor, -1
    .WriteProperty "HotTracking", propHotTracking, False
    .WriteProperty "HoverTime", propHoverTime, 1
    .WriteProperty "IncludedItems", propIncludedItems, IncludedItemsConstants.iiFSFolders Or IncludedItemsConstants.iiNonFSFolders
    .WriteProperty "Indent", propIndent, 16
    .WriteProperty "ItemHeight", propItemHeight, 16
    .WriteProperty "LimitEditInput", propLimitEditInput, True
    .WriteProperty "LineColor", propLineColor, SystemColorConstants.vbButtonShadow
    .WriteProperty "LineStyle", propLineStyle, LineStyleConstants.lsLinesAtItem
    .WriteProperty "Margin", propMargin, 3
    .WriteProperty "MaxScrollTime", propMaxScrollTime, 100
    .WriteProperty "MouseIcon", propMouseIcon, Nothing
    .WriteProperty "MousePointer", propMousePointer, MousePointerConstants.vbDefault
    .WriteProperty "Path", propPath, "Desktop"
    .WriteProperty "ProcessContextMenuKeys", propProcessContextMenuKeys, True
    .WriteProperty "Renaming", propRenaming, RenamingConstants.rAutomatic
    .WriteProperty "ReplaceHandCursor", propReplaceHandCursor, False
    .WriteProperty "RootPath", propRootPath, "Desktop"
    .WriteProperty "RTLReading", propRTLReading, Ambient.RightToLeft
    .WriteProperty "ScrollBars", propScrollBars, ScrollBarsConstants.sbNormal
    .WriteProperty "SelChangedDelayTime", propSelChangedDelayTime, 0
    .WriteProperty "SelectedItemBackColor", propSelectedItemBackColor, -1
    .WriteProperty "SelectedItemForeColor", propSelectedItemForeColor, -1
    .WriteProperty "ShowFocusRect", propShowFocusRect, True
    .WriteProperty "ShowHiddenDimmed", propShowHiddenDimmed, True
    .WriteProperty "ShowIcons", propShowIcons, True
    .WriteProperty "ShownOverlays", propShownOverlays, ShownOverlaysConstants.soAll
    .WriteProperty "ShownToolTips", propShownToolTips, ShownToolTipsConstants.sttHiddenItems
    .WriteProperty "ShowRoot", propShowRoot, True
    .WriteProperty "ShowSelection", propShowSelection, True
    .WriteProperty "ShowShellContextMenus", propShowShellContextMenus, True
    .WriteProperty "SingleExpand", propSingleExpand, SingleExpandConstants.seNone
    .WriteProperty "SortingCriteria", propSortingCriteria, SortingCriteriaConstants.scName
    .WriteProperty "SortOrder", propSortOrder, SortOrderConstants.sorAscending
    .WriteProperty "TreeViewStyle", propTreeViewStyle, TreeViewStyleConstants.tvsButtons Or TreeViewStyleConstants.tvsLines
    .WriteProperty "UseCustomIcons", propUseCustomIcons, False
    .WriteProperty "UseFileFilters", propUseFileFilters, False
    .WriteProperty "UseFolderFilters", propUseFolderFilters, False
    .WriteProperty "UseSelectedIcons", propUseSelectedIcons, True
    .WriteProperty "UseStandardIcons", propUseStandardIcons, False
    .WriteProperty "UseStandardIconsForSlowItems", propUseStandardIconsForSlowItems, False
  End With
End Sub


' private Props

Private Property Get Dragging() As Boolean
  Dragging = propDragging
End Property

Private Property Let Dragging(ByVal New_Dragging As Boolean)
  If New_Dragging = propDragging Then Exit Property

  propDragging = New_Dragging
  Refresh     ' ToDo: braucht man das?
End Property

Private Property Get Dragging_Edit() As Boolean
  Dragging_Edit = propDragging_Edit
End Property

Private Property Let Dragging_Edit(ByVal New_Dragging_Edit As Boolean)
  If New_Dragging_Edit = propDragging_Edit Then Exit Property

  propDragging_Edit = New_Dragging_Edit
End Property

Private Property Get hDropTargetItem() As Long
  hDropTargetItem = prophDropTargetItem
End Property

Private Property Let hDropTargetItem(ByVal New_hDropTargetItem As Long)
  Dim IDataObj As IVBDataObject
  Dim IParent As IVBShellFolder
  Dim pIDLToParent As Long
  Dim TVWItemInfo As ITEMINFO_COMMON

  If New_hDropTargetItem = prophDropTargetItem Then
    If Not ItemExists(prophDropTargetItem, AllowNullConstants.anNever) Then
      curEffects = DROPEFFECTConstants.DROPEFFECT_NONE
    End If
    Exit Property
  End If

  If Not (IDrpTrgt Is Nothing) Then
    ' ein DragLeave simulieren
    On Error Resume Next
    IDrpTrgt.DragLeave
    Set IDrpTrgt = Nothing
  End If

  prophDropTargetItem = New_hDropTargetItem

  If ItemExists(prophDropTargetItem, AllowNullConstants.anNever) Then
    SetTimer prophWnd, TimerID_DragExpand, propDragExpandTime, 0

    ' ein DragEnter simulieren
    If GetItemInfo_Common(prophDropTargetItem, TVWItemInfo) Then
      #If Debuging Then
        SplitFullyQualifiedPIDL debugger, TVWItemInfo.pIDLToDesktop, IParent, pIDLToParent
      #Else
        SplitFullyQualifiedPIDL TVWItemInfo.pIDLToDesktop, IParent, pIDLToParent
      #End If
      Set IDrpTrgt = Nothing
      IParent.GetUIObjectOf prophWndShellUIParentWindow, 1, pIDLToParent, IID_IDropTarget, 0, IDrpTrgt
      If Not (IDrpTrgt Is Nothing) Then
        PostMessage prophWnd, TVM_SELECTITEM, TVGN_DROPHILITE, prophDropTargetItem
        CopyMemory VarPtr(IDataObj), ObjPtr(DataObj) + 16, 4
        If Not (IDataObj Is Nothing) Then
          IDataObj.AddRef
        End If
        On Error Resume Next
        IDrpTrgt.DragEnter IDataObj, curKeyState, curPt.x, curPt.y, curEffects
        Set IDataObj = Nothing
      Else
        curEffects = DROPEFFECTConstants.DROPEFFECT_NONE
      End If
      Set IParent = Nothing
    Else
      curEffects = DROPEFFECTConstants.DROPEFFECT_NONE
    End If
  Else
    PostMessage prophWnd, TVM_SELECTITEM, TVGN_DROPHILITE, 0
    KillTimer prophWnd, TimerID_DragExpand
    curEffects = DROPEFFECTConstants.DROPEFFECT_NONE
    InvalidateRectAsLong prophWnd, 0, 1
  End If
End Property

Private Property Get Hot() As Boolean
  Hot = propHot
End Property

Private Property Let Hot(ByVal New_Hot As Boolean)
  If New_Hot = propHot Then Exit Property

  propHot = New_Hot
End Property

Private Property Get Hot_Edit() As Boolean
  Hot_Edit = propHot_Edit
End Property

Private Property Let Hot_Edit(ByVal New_Hot_Edit As Boolean)
  If New_Hot_Edit = propHot_Edit Then Exit Property

  propHot_Edit = New_Hot_Edit
End Property

Private Property Get inIDE() As Boolean
  On Error Resume Next
  inIDE = Not Ambient.UserMode
End Property

Private Property Get MouseOverControl() As Boolean
  MouseOverControl = propMouseOverControl
End Property

Private Property Let MouseOverControl(ByVal New_MouseOverControl As Boolean)
  If New_MouseOverControl = propMouseOverControl Then Exit Property

  propMouseOverControl = New_MouseOverControl
End Property

Private Property Get MouseOverControl_Edit() As Boolean
  MouseOverControl_Edit = propMouseOverControl_Edit
End Property

Private Property Let MouseOverControl_Edit(ByVal New_MouseOverControl_Edit As Boolean)
  If New_MouseOverControl_Edit = propMouseOverControl_Edit Then Exit Property

  propMouseOverControl_Edit = New_MouseOverControl_Edit
End Property


' öffentliche Props

Public Property Get AllowDragDrop() As Boolean
Attribute AllowDragDrop.VB_Description = "Gibt an, ob Drag'n'Drop-Vorgänge möglich sind."
Attribute AllowDragDrop.VB_ProcData.VB_Invoke_Property = ";Drag'n'Drop"
  AllowDragDrop = propAllowDragDrop
End Property

Public Property Let AllowDragDrop(ByVal New_AllowDragDrop As Boolean)
  Dim Style As Long

  If New_AllowDragDrop = propAllowDragDrop Then Exit Property

  propAllowDragDrop = New_AllowDragDrop

  If prophWnd Then
    If propAllowDragDrop Then
      If propDragShowAlphaBlendedItems Then
        CoCreateInstance IID_IDragDropHelper, Nothing, CLSCTXConsts.CLSCTX_INPROC_SERVER, IID_IDropTargetHelper, IDropTargetHelper
      End If
    Else
      Set IDropTargetHelper = Nothing
    End If
    UserControl.OLEDropMode = IIf(propAllowDragDrop, OLEDropConstants.vbOLEDropManual, OLEDropConstants.vbOLEDropNone)

    Style = GetWindowLongPtr(prophWnd, GWL_STYLE)
    Style = IIf(propAllowDragDrop, Style And Not TreeViewStyleConsts.TVS_DISABLEDRAGDROP, Style Or TreeViewStyleConsts.TVS_DISABLEDRAGDROP)
    SetWindowLongPtr prophWnd, GWL_STYLE, Style
  End If

  PropertyChanged "AllowDragDrop"
End Property

Public Property Get AlwaysUseSmallIcons() As Boolean
Attribute AlwaysUseSmallIcons.VB_Description = "Gibt an, ob auch dann die kleinen Icons benutzt werden sollen, wenn <ItemHeight> große Icons zulassen würde."
Attribute AlwaysUseSmallIcons.VB_ProcData.VB_Invoke_Property = ";Icons"
  AlwaysUseSmallIcons = propAlwaysUseSmallIcons
End Property

Public Property Let AlwaysUseSmallIcons(ByVal New_AlwaysUseSmallIcons As Boolean)
  Const WM_SETREDRAW = &HB
  Dim New_UseLargeIcons As Boolean
  Dim Size_LargeIcons As Long
  Dim Size_SmallIcons As Long

  If New_AlwaysUseSmallIcons = propAlwaysUseSmallIcons Then Exit Property

  propAlwaysUseSmallIcons = New_AlwaysUseSmallIcons

  If Not inIDE Then
    If prophImageList Then
      ' ggf. ImageList updaten
      If propAlwaysUseSmallIcons Then
        New_UseLargeIcons = False
      Else
        Size_LargeIcons = GetSystemMetrics(SM_CXICON)
        Size_SmallIcons = GetSystemMetrics(SM_CXSMICON)
        New_UseLargeIcons = (propItemHeight >= Size_LargeIcons * 0.75)
      End If

      If useLargeIcons <> New_UseLargeIcons Then
        useLargeIcons = New_UseLargeIcons
        SendMessageAsLong prophWnd, WM_SETREDRAW, 0, 0

        If Not propUseCustomIcons Then
          ' ImageList "befreien"
          SendMessageAsLong prophWnd, TVM_SETIMAGELIST, TVSIL_NORMAL, 0
          prophImageList = GetSysImageList(useLargeIcons)
          SendMessageAsLong prophWnd, TVM_SETIMAGELIST, TVSIL_NORMAL, prophImageList
        End If

        SendMessageAsLong prophWnd, WM_SETREDRAW, 1, 0
      End If
    End If
  End If

  PropertyChanged "AlwaysUseSmallIcons"
End Property

Public Property Get Appearance() As AppearanceConstants
Attribute Appearance.VB_Description = "Gibt an, ob das Control 3D-Effekte hat."
Attribute Appearance.VB_ProcData.VB_Invoke_Property = ";Darstellung"
Attribute Appearance.VB_UserMemId = -520
  Dim ExStyle As Long

  If prophWnd Then
    ' aktuellen Wert ermitteln
    ExStyle = GetWindowLongPtr(prophWnd, GWL_EXSTYLE)
    If ExStyle And ExWindowStyleConsts.WS_EX_CLIENTEDGE Then
      propAppearance = AppearanceConstants.a3D
    ElseIf ExStyle And ExWindowStyleConsts.WS_EX_STATICEDGE Then
      propAppearance = AppearanceConstants.a3DLight
    Else
      propAppearance = AppearanceConstants.a2D
    End If
  End If

  Appearance = propAppearance
End Property

Public Property Let Appearance(ByVal New_Appearance As AppearanceConstants)
  Dim ExStyle As Long

  If Not isValidNumber(New_Appearance, AppearanceConstants.a2D, AppearanceConstants.a3DLight) Then
    Err.Raise 380
    Exit Property
  End If
  If New_Appearance = propAppearance Then Exit Property

  propAppearance = New_Appearance

  If prophWnd Then
    ' Control anpassen
    ExStyle = GetWindowLongPtr(prophWnd, GWL_EXSTYLE)
    ExStyle = ExStyle And Not ExWindowStyleConsts.WS_EX_CLIENTEDGE And Not ExWindowStyleConsts.WS_EX_STATICEDGE
    If propAppearance = AppearanceConstants.a3D Then
      ExStyle = ExStyle Or ExWindowStyleConsts.WS_EX_CLIENTEDGE
    ElseIf propAppearance = AppearanceConstants.a3DLight Then
      ExStyle = ExStyle Or ExWindowStyleConsts.WS_EX_STATICEDGE
    End If
    SetWindowLongPtr prophWnd, GWL_EXSTYLE, ExStyle
    SetWindowPos prophWnd, 0, 0, 0, 0, 0, SWP_DRAWFRAME Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER
  End If

  PropertyChanged "Appearance"
End Property

Public Property Get ArchiveItemBackColor() As OLE_COLOR
Attribute ArchiveItemBackColor.VB_Description = "Gibt die Hintergrundfarbe für Items an, die sich in einem Archiv befinden. Bei -1 wird die Hintergrundfarbe des  Controls genutzt."
Attribute ArchiveItemBackColor.VB_ProcData.VB_Invoke_Property = ";Farben"
  ArchiveItemBackColor = propArchiveItemBackColor
End Property

Public Property Let ArchiveItemBackColor(ByVal New_ArchiveItemBackColor As OLE_COLOR)
  If New_ArchiveItemBackColor = propArchiveItemBackColor Then Exit Property

  propArchiveItemBackColor = New_ArchiveItemBackColor

  If Not inIDE Then
    mustCustomDraw = IsCustomDrawNecessary
    Refresh
  End If

  PropertyChanged "ArchiveItemBackColor"
End Property

Public Property Get ArchiveItemForeColor() As OLE_COLOR
Attribute ArchiveItemForeColor.VB_Description = "Gibt die Schriftfarbe für Items an, die sich in einem Archiv befinden. Bei -1 wird die Schriftfarbe des  Controls genutzt."
Attribute ArchiveItemForeColor.VB_ProcData.VB_Invoke_Property = ";Farben"
  ArchiveItemForeColor = propArchiveItemForeColor
End Property

Public Property Let ArchiveItemForeColor(ByVal New_ArchiveItemForeColor As OLE_COLOR)
  If New_ArchiveItemForeColor = propArchiveItemForeColor Then Exit Property

  propArchiveItemForeColor = New_ArchiveItemForeColor

  If Not inIDE Then
    mustCustomDraw = IsCustomDrawNecessary
    Refresh
  End If

  PropertyChanged "ArchiveItemForeColor"
End Property

Public Property Get ArchivesAfterFolders() As Boolean
Attribute ArchivesAfterFolders.VB_Description = "Gibt an, ob Archive, welche nach <ExpandArchives> als Ordner behandelt werden, direkt nach den Ordnern einsortiert werden."
Attribute ArchivesAfterFolders.VB_ProcData.VB_Invoke_Property = ";Sortierung"
  ArchivesAfterFolders = propArchivesAfterFolders
End Property

Public Property Let ArchivesAfterFolders(ByVal New_ArchivesAfterFolders As Boolean)
  If New_ArchivesAfterFolders = propArchivesAfterFolders Then Exit Property

  propArchivesAfterFolders = New_ArchivesAfterFolders

  If Not inIDE Then
    If prophWnd Then
      SortAllSubItems IIf(propShowRoot, GetRootOrFirstItem, 0)
    End If
  End If

  PropertyChanged "ArchivesAfterFolders"
End Property

Public Property Get AutoHScroll() As Boolean
Attribute AutoHScroll.VB_Description = "Gibt an, ob das Control für eine optimale Darstellung des markierten Items automatisch horizontal gescrollt wird. Benötigt Version 6.10 oder neuer der comctl32.dll."
  Dim Style As Long

  If ver_Comctl32_610 Then
    If prophWnd Then
      ' aktuellen Wert ermitteln
      Style = SendMessageAsLong(prophWnd, TVM_GETEXTENDEDSTYLE, 0, 0)
      propAutoHScroll = ((Style And TreeViewExStyleConsts.TVS_EX_AUTOHSCROLL) = TreeViewExStyleConsts.TVS_EX_AUTOHSCROLL)
    End If
  End If

  AutoHScroll = propAutoHScroll
End Property

Public Property Let AutoHScroll(ByVal New_AutoHScroll As Boolean)
  If New_AutoHScroll = propAutoHScroll Then Exit Property

  propAutoHScroll = New_AutoHScroll

  If ver_Comctl32_610 Then
    If prophWnd Then
      ' Control aktualisieren
      If propAutoHScroll Then
        SendMessageAsLong prophWnd, TVM_SETEXTENDEDSTYLE, TreeViewExStyleConsts.TVS_EX_AUTOHSCROLL, TreeViewExStyleConsts.TVS_EX_AUTOHSCROLL
      Else
        SendMessageAsLong prophWnd, TVM_SETEXTENDEDSTYLE, TreeViewExStyleConsts.TVS_EX_AUTOHSCROLL, 0
      End If
    End If
  End If

  PropertyChanged "AutoHScroll"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Gibt die Hintergrundfarbe des Controls an."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Farben"
Attribute BackColor.VB_UserMemId = -501
  Dim RGBClr As Long

  If prophWnd Then
    #If BkImageSupport Then
      If propBkImage Is Nothing Then
    #End If
      ' aktuellen Wert ermitteln
      RGBClr = SendMessageAsLong(prophWnd, TVM_GETBKCOLOR, 0, 0)
      If RGBClr <> TranslateColor(propBackColor) Then
        propBackColor = RGBClr
      End If
    #If BkImageSupport Then
      End If
    #End If
  End If

  BackColor = propBackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
  Dim RGBClr As Long

  propBackColor = New_BackColor

  If prophWnd Then
    mustCustomDraw = IsCustomDrawNecessary
    ' Hintergrundfarbe setzen
    #If BkImageSupport Then
      RGBClr = IIf(propBkImage Is Nothing, TranslateColor(propBackColor), -1)
    #Else
      RGBClr = TranslateColor(propBackColor)
    #End If
    SendMessageAsLong prophWnd, TVM_SETBKCOLOR, 0, RGBClr
  End If

  PropertyChanged "BackColor"
End Property

Public Property Get BkImage() As StdPicture
Attribute BkImage.VB_Description = "Not supported anymore. Left for backward compatibility."
Attribute BkImage.VB_ProcData.VB_Invoke_Property = ";Hintergrund"
Attribute BkImage.VB_MemberFlags = "40"
  #If BkImageSupport Then
    Set BkImage = propBkImage
  #End If
End Property

Public Property Set BkImage(ByVal New_BkImage As StdPicture)
  #If BkImageSupport Then
    On Error Resume Next
    If New_BkImage.Handle = 0 Then Set New_BkImage = Nothing

    Set propBkImage = New_BkImage

    If Not inIDE Then
      mustCustomDraw = IsCustomDrawNecessary
      If propBkImage Is Nothing Then
        SendMessageAsLong prophWnd, TVM_SETBKCOLOR, 0, TranslateColor(propBackColor)
      Else
        SendMessageAsLong prophWnd, TVM_SETBKCOLOR, 0, -1
      End If
      InvalidateRectAsLong prophWnd, 0, 1
    End If

    PropertyChanged "BkImage"
  #End If
End Property

Public Property Get BkImageFixed() As Boolean
Attribute BkImageFixed.VB_Description = "Not supported anymore. Left for backward compatibility."
Attribute BkImageFixed.VB_ProcData.VB_Invoke_Property = ";Hintergrund"
Attribute BkImageFixed.VB_MemberFlags = "40"
  #If BkImageSupport Then
    BkImageFixed = propBkImageFixed
  #End If
End Property

Public Property Let BkImageFixed(ByVal New_BkImageFixed As Boolean)
  #If BkImageSupport Then
    If New_BkImageFixed = propBkImageFixed Then Exit Property

    propBkImageFixed = New_BkImageFixed

    If Not inIDE Then
      If Not (propBkImage Is Nothing) Then
        InvalidateRectAsLong prophWnd, 0, 1
      End If
    End If

    PropertyChanged "BkImageFixed"
  #End If
End Property

Public Property Get BkImageStyle() As BkImageStyleConstants
Attribute BkImageStyle.VB_Description = "Not supported anymore. Left for backward compatibility."
Attribute BkImageStyle.VB_ProcData.VB_Invoke_Property = ";Hintergrund"
Attribute BkImageStyle.VB_MemberFlags = "40"
  #If BkImageSupport Then
    BkImageStyle = propBkImageStyle
  #End If
End Property

Public Property Let BkImageStyle(ByVal New_BkImageStyle As BkImageStyleConstants)
  #If BkImageSupport Then
    #If Debuging Then
      If Not isValidNumber(New_BkImageStyle, BkImageStyleConstants.bisTiled, BkImageStyleConstants.bisStretched) Then
    #Else
      If Not isValidNumber(New_BkImageStyle, BkImageStyleConstants.bisTiled, BkImageStyleConstants.bisTiled) Then
    #End If
      Err.Raise 380
      Exit Property
    End If
    If New_BkImageStyle = propBkImageStyle Then Exit Property

    propBkImageStyle = New_BkImageStyle

    If Not inIDE Then
      If Not (propBkImage Is Nothing) Then
        InvalidateRectAsLong prophWnd, 0, 1
      End If
    End If

    PropertyChanged "BkImageStyle"
  #End If
End Property

Public Property Get BorderStyle() As BorderStyleConstants
Attribute BorderStyle.VB_Description = "Gibt an, wie der Rahmen des Controls angezeigt werden soll."
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Darstellung"
Attribute BorderStyle.VB_UserMemId = -504
  Dim Style As Long

  If prophWnd Then
    ' aktuellen Wert ermitteln
    Style = GetWindowLongPtr(prophWnd, GWL_STYLE)
    propBorderStyle = IIf(Style And WindowStyleConsts.WS_BORDER, BorderStyleConstants.bsFixedSingle, BorderStyleConstants.bsNone)
  End If

  BorderStyle = propBorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleConstants)
  Dim Style As Long

  If Not isValidNumber(New_BorderStyle, BorderStyleConstants.bsNone, BorderStyleConstants.bsFixedSingle) Then
    Err.Raise 380
    Exit Property
  End If
  If New_BorderStyle = propBorderStyle Then Exit Property

  propBorderStyle = New_BorderStyle

  If prophWnd Then
    ' Control anpassen
    Style = GetWindowLongPtr(prophWnd, GWL_STYLE)
    Style = IIf(propBorderStyle = BorderStyleConstants.bsFixedSingle, Style Or WindowStyleConsts.WS_BORDER, Style And Not WindowStyleConsts.WS_BORDER)
    SetWindowLongPtr prophWnd, GWL_STYLE, Style
    SetWindowPos prophWnd, 0, 0, 0, 0, 0, SWP_DRAWFRAME Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER
  End If

  PropertyChanged "BorderStyle"
End Property

Public Property Get CheckBoxes() As Boolean
Attribute CheckBoxes.VB_Description = "Gibt an, ob vor jedem Item ein Kontrollkästchen angezeigt wird."
Attribute CheckBoxes.VB_ProcData.VB_Invoke_Property = ";Darstellung"
  Dim Style As Long

  If prophWnd Then
    ' aktuellen Wert ermitteln
    Style = GetWindowLongPtr(prophWnd, GWL_STYLE)
    propCheckBoxes = ((Style And TreeViewStyleConsts.TVS_CHECKBOXES) = TreeViewStyleConsts.TVS_CHECKBOXES)
  End If

  CheckBoxes = propCheckBoxes
End Property

Public Property Let CheckBoxes(ByVal New_CheckBoxes As Boolean)
  If New_CheckBoxes = propCheckBoxes Then Exit Property

  propCheckBoxes = New_CheckBoxes

  ' das Control neu erzeugen...
  If prophWnd Then CreateControl

  PropertyChanged "CheckBoxes"
End Property

Public Property Get CheckSlowItemsForSubItems() As Boolean
Attribute CheckSlowItemsForSubItems.VB_Description = "Gibt an, ob langsame Items (Wechseldatenträger, Netzlaufwerke etc.) auf SubItems geprüft werden."
Attribute CheckSlowItemsForSubItems.VB_ProcData.VB_Invoke_Property = ";Verhalten"
  CheckSlowItemsForSubItems = propCheckSlowItemsForSubItems
End Property

Public Property Let CheckSlowItemsForSubItems(ByVal New_CheckSlowItemsForSubItems As Boolean)
  propCheckSlowItemsForSubItems = New_CheckSlowItemsForSubItems

  PropertyChanged "CheckSlowItemsForSubItems"
End Property

Public Property Get CompressedBackColor() As OLE_COLOR
Attribute CompressedBackColor.VB_Description = "Gibt die Hintergrundfarbe für komprimierte Items an. Bei -1 wird die Hintergrundfarbe des  Controls genutzt."
Attribute CompressedBackColor.VB_ProcData.VB_Invoke_Property = ";Farben"
  CompressedBackColor = propCompressedBackColor
End Property

Public Property Let CompressedBackColor(ByVal New_CompressedBackColor As OLE_COLOR)
  If New_CompressedBackColor = propCompressedBackColor Then Exit Property

  propCompressedBackColor = New_CompressedBackColor

  If Not inIDE Then
    mustCustomDraw = IsCustomDrawNecessary
    Refresh
  End If

  PropertyChanged "CompressedBackColor"
End Property

Public Property Get CompressedForeColor() As OLE_COLOR
Attribute CompressedForeColor.VB_Description = "Gibt die Schriftfarbe für komprimierte Items an. Bei -1 wird die Standardfarbe des  Systems genutzt."
Attribute CompressedForeColor.VB_ProcData.VB_Invoke_Property = ";Farben"
  CompressedForeColor = propCompressedForeColor
End Property

Public Property Let CompressedForeColor(ByVal New_CompressedForeColor As OLE_COLOR)
  If New_CompressedForeColor = propCompressedForeColor Then Exit Property

  propCompressedForeColor = New_CompressedForeColor

  If Not inIDE Then
    mustCustomDraw = IsCustomDrawNecessary
    Refresh
  End If

  PropertyChanged "CompressedForeColor"
End Property

Public Property Get DimSelectedItemsIcon() As Boolean
Attribute DimSelectedItemsIcon.VB_Description = "Gibt an, ob das Icon des markierten Items geblendet werden soll."
Attribute DimSelectedItemsIcon.VB_ProcData.VB_Invoke_Property = ";Icons"
  DimSelectedItemsIcon = propDimSelectedItemsIcon
End Property

Public Property Let DimSelectedItemsIcon(ByVal New_DimSelectedItemsIcon As Boolean)
  If New_DimSelectedItemsIcon = propDimSelectedItemsIcon Then Exit Property

  propDimSelectedItemsIcon = New_DimSelectedItemsIcon

  If Not inIDE Then
    mustCustomDraw = IsCustomDrawNecessary
    Refresh
  End If

  PropertyChanged "DimSelectedItemsIcon"
End Property

Public Property Get DontCheckFolderExpandibility() As Boolean
Attribute DontCheckFolderExpandibility.VB_Description = "Gibt an, ob Ordner auf Sub-Items geprüft werden. Das Prüfen großer Ordner auf Sub-Items kann sehr lange dauern."
Attribute DontCheckFolderExpandibility.VB_ProcData.VB_Invoke_Property = ";Verhalten"
  DontCheckFolderExpandibility = propDontCheckFolderExpandibility
End Property

Public Property Let DontCheckFolderExpandibility(ByVal New_DontCheckFolderExpandibility As Boolean)
  propDontCheckFolderExpandibility = New_DontCheckFolderExpandibility
  PropertyChanged "DontCheckFolderExpandibility"
End Property

Public Property Get DragExpandTime() As Long
Attribute DragExpandTime.VB_Description = "Gibt an, wie lange (in Millisekunden) sich der Mauszeiger während eines Drag'n'Drop-Vorgangs über einem Item befinden muß, damit dieser erweitert wird."
Attribute DragExpandTime.VB_ProcData.VB_Invoke_Property = ";Drag'n'Drop"
  DragExpandTime = propDragExpandTime
End Property

Public Property Let DragExpandTime(ByVal New_DragExpandTime As Long)
  If Not isValidNumber(New_DragExpandTime, 1, 60000) Then
    Err.Raise 380
    Exit Property
  End If
  propDragExpandTime = New_DragExpandTime

  PropertyChanged "DragExpandTime"
End Property

Public Property Get DragScrollTime() As Long
Attribute DragScrollTime.VB_Description = "Gibt an, in welchen Abständen (in Millisekunden) das Control während eines Drag'n'Drop-Vorgangs um je 1 Einheit gescrollt wird (solte sich der Mauszeiger in der ScrollZone befinden)."
Attribute DragScrollTime.VB_ProcData.VB_Invoke_Property = ";Drag'n'Drop"
  DragScrollTime = propDragScrollTime
End Property

Public Property Let DragScrollTime(ByVal New_DragScrollTime As Long)
  If Not isValidNumber(New_DragScrollTime, 1, 60000) Then
    Err.Raise 380
    Exit Property
  End If
  propDragScrollTime = New_DragScrollTime

  PropertyChanged "DragScrollTime"
End Property

Public Property Get DragScrollZoneBottom() As Long
Attribute DragScrollZoneBottom.VB_Description = "Gibt die Breite des Bereichs am unteren Rand an, in dem sich der Mauszeiger während eines Drag'n'Drop-Vorgangs befinden muß, um das Control zu scrollen."
Attribute DragScrollZoneBottom.VB_ProcData.VB_Invoke_Property = ";Drag'n'Drop"
  DragScrollZoneBottom = propDragScrollZoneBottom
End Property

Public Property Let DragScrollZoneBottom(New_DragScrollZoneBottom As Long)
  propDragScrollZoneBottom = New_DragScrollZoneBottom

  PropertyChanged "DragScrollZoneBottom"
End Property

Public Property Get DragScrollZoneLeft() As Long
Attribute DragScrollZoneLeft.VB_Description = "Gibt die Breite des Bereichs am linken Rand an, in dem sich der Mauszeiger während eines Drag'n'Drop-Vorgangs befinden muß, um das Control zu scrollen."
Attribute DragScrollZoneLeft.VB_ProcData.VB_Invoke_Property = ";Drag'n'Drop"
  DragScrollZoneLeft = propDragScrollZoneLeft
End Property

Public Property Let DragScrollZoneLeft(New_DragScrollZoneLeft As Long)
  propDragScrollZoneLeft = New_DragScrollZoneLeft

  PropertyChanged "DragScrollZoneLeft"
End Property

Public Property Get DragScrollZoneRight() As Long
Attribute DragScrollZoneRight.VB_Description = "Gibt die Breite des Bereichs am rechten Rand an, in dem sich der Mauszeiger während eines Drag'n'Drop-Vorgangs befinden muß, um das Control zu scrollen."
Attribute DragScrollZoneRight.VB_ProcData.VB_Invoke_Property = ";Drag'n'Drop"
  DragScrollZoneRight = propDragScrollZoneRight
End Property

Public Property Let DragScrollZoneRight(New_DragScrollZoneRight As Long)
  propDragScrollZoneRight = New_DragScrollZoneRight

  PropertyChanged "DragScrollZoneRight"
End Property

Public Property Get DragScrollZoneTop() As Long
Attribute DragScrollZoneTop.VB_Description = "Gibt die Breite des Bereichs am oberen Rand an, in dem sich der Mauszeiger während eines Drag'n'Drop-Vorgangs befinden muß, um das Control zu scrollen."
Attribute DragScrollZoneTop.VB_ProcData.VB_Invoke_Property = ";Drag'n'Drop"
  DragScrollZoneTop = propDragScrollZoneTop
End Property

Public Property Let DragScrollZoneTop(New_DragScrollZoneTop As Long)
  propDragScrollZoneTop = New_DragScrollZoneTop

  PropertyChanged "DragScrollZoneTop"
End Property

Public Property Get DragShowAlphaBlendedItems() As Boolean
Attribute DragShowAlphaBlendedItems.VB_Description = "Gibt an, ob bei einem Drag'n'Drop-Vorgang der gedraggte Item am Mauszeiger angehängt wird."
Attribute DragShowAlphaBlendedItems.VB_ProcData.VB_Invoke_Property = ";Drag'n'Drop"
  DragShowAlphaBlendedItems = propDragShowAlphaBlendedItems
End Property

Public Property Let DragShowAlphaBlendedItems(New_DragShowAlphaBlendedItems As Boolean)
  If New_DragShowAlphaBlendedItems = propDragShowAlphaBlendedItems Then Exit Property

  propDragShowAlphaBlendedItems = New_DragShowAlphaBlendedItems

  If propAllowDragDrop Then
    If propDragShowAlphaBlendedItems Then
      CoCreateInstance IID_IDragDropHelper, Nothing, CLSCTXConsts.CLSCTX_INPROC_SERVER, IID_IDropTargetHelper, IDropTargetHelper
    Else
      Set IDropTargetHelper = Nothing
    End If
  Else
    Set IDropTargetHelper = Nothing
  End If

  PropertyChanged "DragShowAlphaBlendedItems"
End Property

Public Property Get DrivesOnly() As Boolean
Attribute DrivesOnly.VB_Description = "Gibt an, ob beim Arbeitsplatz nur die Laufwerke angezeigt werden."
Attribute DrivesOnly.VB_ProcData.VB_Invoke_Property = ";Filter"
  DrivesOnly = propDrivesOnly
End Property

Public Property Let DrivesOnly(ByVal New_DrivesOnly As Boolean)
  If New_DrivesOnly = propDrivesOnly Then Exit Property

  propDrivesOnly = New_DrivesOnly

  If Not inIDE Then
    If prophWnd Then
      ReloadItems
    End If
  End If

  PropertyChanged "DrivesOnly"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Gibt an, ob das Control auf vom Benutzer erzeugte Ereignisse reagieren kann."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Verhalten"
Attribute Enabled.VB_UserMemId = -514
  Enabled = propEnabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  propEnabled = New_Enabled

  UserControl.Enabled = propEnabled
  EnableWindow UserControl.hWnd, propEnabled
  If prophWnd Then EnableWindow prophWnd, propEnabled

  PropertyChanged "Enabled"
End Property

Public Property Get EncryptedBackColor() As OLE_COLOR
Attribute EncryptedBackColor.VB_Description = "Gibt die Hintergrundfarbe für verschlüsselte Items an. Bei -1 wird die Hintergrundfarbe des  Controls genutzt."
Attribute EncryptedBackColor.VB_ProcData.VB_Invoke_Property = ";Farben"
  EncryptedBackColor = propEncryptedBackColor
End Property

Public Property Let EncryptedBackColor(ByVal New_EncryptedBackColor As OLE_COLOR)
  If New_EncryptedBackColor = propEncryptedBackColor Then Exit Property

  propEncryptedBackColor = New_EncryptedBackColor

  If Not inIDE Then
    mustCustomDraw = IsCustomDrawNecessary
    Refresh
  End If

  PropertyChanged "EncryptedBackColor"
End Property

Public Property Get EncryptedForeColor() As OLE_COLOR
Attribute EncryptedForeColor.VB_Description = "Gibt die Schriftfarbe für verschlüsselte Items an. Bei -1 wird die Standardfarbe des  Systems genutzt."
Attribute EncryptedForeColor.VB_ProcData.VB_Invoke_Property = ";Farben"
  EncryptedForeColor = propEncryptedForeColor
End Property

Public Property Let EncryptedForeColor(ByVal New_EncryptedForeColor As OLE_COLOR)
  If New_EncryptedForeColor = propEncryptedForeColor Then Exit Property

  propEncryptedForeColor = New_EncryptedForeColor

  If Not inIDE Then
    mustCustomDraw = IsCustomDrawNecessary
    Refresh
  End If

  PropertyChanged "EncryptedForeColor"
End Property

Public Property Get ExpandArchives() As ExpandArchivesConstants
Attribute ExpandArchives.VB_Description = "Gibt an, welche Archive als Ordner behandelt werden sollen."
Attribute ExpandArchives.VB_ProcData.VB_Invoke_Property = ";Filter"
Attribute ExpandArchives.VB_MemberFlags = "400"
  ExpandArchives = propExpandArchives
End Property

Public Property Let ExpandArchives(ByVal New_ExpandArchives As ExpandArchivesConstants)
  Dim hItem As Long

  #If NewArchiveSupport Then
    If Not isValidNumber(New_ExpandArchives, ExpandArchivesConstants.eaNone, ExpandArchivesConstants.eaACE Or ExpandArchivesConstants.eaBIN Or ExpandArchivesConstants.eaCAB Or ExpandArchivesConstants.eaISO Or ExpandArchivesConstants.eaJAR Or ExpandArchivesConstants.eaRAR Or ExpandArchivesConstants.eaTAR Or ExpandArchivesConstants.eaZIP) Then
  #Else
    If Not isValidNumber(New_ExpandArchives, ExpandArchivesConstants.eaNone, ExpandArchivesConstants.eaACE Or ExpandArchivesConstants.eaBIN Or ExpandArchivesConstants.eaCAB Or ExpandArchivesConstants.eaISO Or ExpandArchivesConstants.eaJAR Or ExpandArchivesConstants.eaRAR Or ExpandArchivesConstants.eaZIP) Then
  #End If
    Err.Raise 380
    Exit Property
  End If
  If New_ExpandArchives = propExpandArchives Then Exit Property

  propExpandArchives = New_ExpandArchives

  If Not inIDE Then
    If prophWnd Then
      mustCustomDraw = IsCustomDrawNecessary
      ' alle Archive prüfen
      If propShowRoot Then hItem = GetRootOrFirstItem
      UpdateAllArchives hItem
    End If
  End If

  PropertyChanged "ExpandArchives"
End Property

Public Property Get FadeExpandos() As Boolean
Attribute FadeExpandos.VB_Description = "Gibt an, ob die Item-Buttons ausgeblendet werden, wenn sich die Maus nicht über dem Control befindet. Benötigt Version 6.10 oder neuer der comctl32.dll."
Attribute FadeExpandos.VB_ProcData.VB_Invoke_Property = ";Verhalten"
  Dim Style As Long

  If ver_Comctl32_610 Then
    If prophWnd Then
      ' aktuellen Wert ermitteln
      Style = SendMessageAsLong(prophWnd, TVM_GETEXTENDEDSTYLE, 0, 0)
      propFadeExpandos = ((Style And TreeViewExStyleConsts.TVS_EX_FADEINOUTEXPANDOS) = TreeViewExStyleConsts.TVS_EX_FADEINOUTEXPANDOS)
    End If
  End If

  FadeExpandos = propFadeExpandos
End Property

Public Property Let FadeExpandos(ByVal New_FadeExpandos As Boolean)
  If New_FadeExpandos = propFadeExpandos Then Exit Property

  propFadeExpandos = New_FadeExpandos

  If ver_Comctl32_610 Then
    If prophWnd Then
      ' Control aktualisieren
      If propFadeExpandos Then
        SendMessageAsLong prophWnd, TVM_SETEXTENDEDSTYLE, TreeViewExStyleConsts.TVS_EX_FADEINOUTEXPANDOS, TreeViewExStyleConsts.TVS_EX_FADEINOUTEXPANDOS
      Else
        SendMessageAsLong prophWnd, TVM_SETEXTENDEDSTYLE, TreeViewExStyleConsts.TVS_EX_FADEINOUTEXPANDOS, 0
      End If
    End If
  End If

  PropertyChanged "FadeExpandos"
End Property

Public Property Get FastInfoTipsOnly() As Boolean
Attribute FastInfoTipsOnly.VB_Description = "Gibt an oder legt fest, ob nur InfoTips angezeigt werden, deren Generierung sich nicht negativ auf die Performance auswirkt."
Attribute FastInfoTipsOnly.VB_ProcData.VB_Invoke_Property = ";Verhalten"
  FastInfoTipsOnly = propFastInfoTipsOnly
End Property

Public Property Let FastInfoTipsOnly(ByVal New_FastInfoTipsOnly As Boolean)
  propFastInfoTipsOnly = New_FastInfoTipsOnly
  PropertyChanged "FastInfoTipsOnly"
End Property

Public Property Get FavoritesStyle() As Boolean
Attribute FavoritesStyle.VB_Description = "Gibt an, ob um die Gruppe des markierten Items ein Rehteck gezeichnet wird (wie bei den Windows-Explorerleisten)."
Attribute FavoritesStyle.VB_ProcData.VB_Invoke_Property = ";Verhalten"
  FavoritesStyle = propFavoritesStyle
End Property

Public Property Let FavoritesStyle(ByVal New_FavoritesStyle As Boolean)
  If New_FavoritesStyle = propFavoritesStyle Then Exit Property

  propFavoritesStyle = New_FavoritesStyle
  If propFavoritesStyle Then SingleExpand = SingleExpandConstants.seNormal

  If Not inIDE Then
    mustCustomDraw = IsCustomDrawNecessary
    Refresh
  End If

  PropertyChanged "FavoritesStyle"
End Property

Public Property Get FileAttributes() As AttributesConstants
Attribute FileAttributes.VB_Description = "Gibt an, welche Dateien angezeigt werden."
Attribute FileAttributes.VB_ProcData.VB_Invoke_Property = ";Filter"
Attribute FileAttributes.VB_MemberFlags = "400"
  FileAttributes = propFileAttributes
End Property

Public Property Let FileAttributes(ByVal New_FileAttributes As AttributesConstants)
  If Not isValidNumber(New_FileAttributes, 0, AttributesConstants.attReadOnly Or AttributesConstants.attHidden Or AttributesConstants.attArchive Or AttributesConstants.attSystem Or AttributesConstants.attEncrypted Or AttributesConstants.attCompressed) Then
    Err.Raise 380
    Exit Property
  End If
  If New_FileAttributes = propFileAttributes Then Exit Property

  propFileAttributes = New_FileAttributes

  If Not inIDE Then
    If prophWnd Then
      ReloadItems
    End If
  End If

  PropertyChanged "FileAttributes"
End Property

Public Property Get FileFilters() As String
Attribute FileFilters.VB_Description = "Gibt an, welche Dateien angezeigt werden. Die Filter müssen durch ""|"" getrennt werden. Ein vorangestelltes ""/"" negiert den Filter. Wildcards sind möglich."
Attribute FileFilters.VB_ProcData.VB_Invoke_Property = ";Filter"
  FileFilters = propFileFilters
End Property

Public Property Let FileFilters(ByVal New_FileFilters As String)
  If New_FileFilters = propFileFilters Then Exit Property

  propFileFilters = New_FileFilters

  If Not inIDE Then
    If prophWnd Then
      ReloadItems
    End If
  End If

  PropertyChanged "FileFilters"
End Property

Public Property Get FireBeforeInsertItem() As FireBeforeInsertItemConstants
Attribute FireBeforeInsertItem.VB_Description = "Gibt an, für welche Art von Items das Event BeforeInsertItem ausgelöst werden soll."
Attribute FireBeforeInsertItem.VB_ProcData.VB_Invoke_Property = ";Verhalten"
Attribute FireBeforeInsertItem.VB_MemberFlags = "400"
  FireBeforeInsertItem = propFireBeforeInsertItem
End Property

Public Property Let FireBeforeInsertItem(ByVal New_FireBeforeInsertItem As FireBeforeInsertItemConstants)
  propFireBeforeInsertItem = New_FireBeforeInsertItem
  PropertyChanged "FireBeforeInsertItem"
End Property

Public Property Get FirstVisibleItem() As Long
Attribute FirstVisibleItem.VB_Description = "Gibt den 1. sichtbaren Item im Control an."
Attribute FirstVisibleItem.VB_ProcData.VB_Invoke_Property = ";Darstellung"
Attribute FirstVisibleItem.VB_MemberFlags = "400"
  Dim ret As Long

  ret = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_FIRSTVISIBLE, 0)
  If ret = 0 Then ret = -1

  FirstVisibleItem = ret
End Property

Public Property Let FirstVisibleItem(ByVal New_FirstVisibleItem As Long)
  If Not ItemExists(New_FirstVisibleItem, AllowNullConstants.anIfNoShowRoot) Then Exit Property

  SendMessageAsLong prophWnd, TVM_SELECTITEM, TVGN_FIRSTVISIBLE, New_FirstVisibleItem
End Property

Public Property Get FolderAttributes() As AttributesConstants
Attribute FolderAttributes.VB_Description = "Gibt an, welche Ordner angezeigt werden."
Attribute FolderAttributes.VB_ProcData.VB_Invoke_Property = ";Filter"
Attribute FolderAttributes.VB_MemberFlags = "400"
  FolderAttributes = propFolderAttributes
End Property

Public Property Let FolderAttributes(ByVal New_FolderAttributes As AttributesConstants)
  If Not isValidNumber(New_FolderAttributes, 0, AttributesConstants.attReadOnly Or AttributesConstants.attHidden Or AttributesConstants.attArchive Or AttributesConstants.attSystem Or AttributesConstants.attEncrypted Or AttributesConstants.attCompressed) Then
    Err.Raise 380
    Exit Property
  End If
  If New_FolderAttributes = propFolderAttributes Then Exit Property

  propFolderAttributes = New_FolderAttributes

  If Not inIDE Then
    If prophWnd Then
      ReloadItems
    End If
  End If

  PropertyChanged "FolderAttributes"
End Property

Public Property Get FolderFilters() As String
Attribute FolderFilters.VB_Description = "Gibt an, welche Ordner angezeigt werden. Die Filter müssen durch ""|"" getrennt werden. Ein vorangestelltes ""/"" negiert den Filter. Wildcards sind möglich."
Attribute FolderFilters.VB_ProcData.VB_Invoke_Property = ";Filter"
  FolderFilters = propFolderFilters
End Property

Public Property Let FolderFilters(ByVal New_FolderFilters As String)
  If New_FolderFilters = propFolderFilters Then Exit Property

  propFolderFilters = New_FolderFilters

  If Not inIDE Then
    If prophWnd Then
      ReloadItems
    End If
  End If

  PropertyChanged "FolderFilters"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Gibt die Schriftart für das Control an."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Schriftart"
Attribute Font.VB_UserMemId = -512
  Set Font = propFont
End Property

Public Property Set Font(ByVal New_Font As Font)
  Set propFont = New_Font

  ChangeFont

  PropertyChanged "Font"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Gibt die Textfarbe des Controls an."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Farben"
Attribute ForeColor.VB_UserMemId = -513
  Dim RGBClr As Long

  If prophWnd Then
    ' aktuellen Wert ermitteln
    RGBClr = SendMessageAsLong(prophWnd, TVM_GETTEXTCOLOR, 0, 0)
    If RGBClr <> TranslateColor(propForeColor) Then propForeColor = RGBClr
  End If

  ForeColor = propForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
  Dim RGBClr As Long

  If New_ForeColor = propForeColor Then Exit Property

  propForeColor = New_ForeColor

  If prophWnd Then
    ' Textfarbe setzen
    RGBClr = TranslateColor(propForeColor)
    SendMessageAsLong prophWnd, TVM_SETTEXTCOLOR, 0, RGBClr
  End If

  PropertyChanged "ForeColor"
End Property

Public Property Get FullRowSelect() As Boolean
Attribute FullRowSelect.VB_Description = "Gibt an, ob die ganze Zeile eines Items markiert wird."
Attribute FullRowSelect.VB_ProcData.VB_Invoke_Property = ";Verhalten"
  Dim Style As Long

  If prophWnd Then
    ' aktuellen Wert ermitteln
    Style = GetWindowLongPtr(prophWnd, GWL_STYLE)
    propFullRowSelect = ((Style And TreeViewStyleConsts.TVS_FULLROWSELECT) = TreeViewStyleConsts.TVS_FULLROWSELECT)
  End If

  FullRowSelect = propFullRowSelect
End Property

Public Property Let FullRowSelect(ByVal New_FullRowSelect As Boolean)
  Dim Style As Long

  If New_FullRowSelect = propFullRowSelect Then Exit Property

  If New_FullRowSelect Then
    ' Lines deaktivieren
    TreeViewStyle = propTreeViewStyle And Not TreeViewStyleConstants.tvsLines
  End If

  propFullRowSelect = New_FullRowSelect
  mustCustomDraw = IsCustomDrawNecessary

  If prophWnd Then
    ' Control aktualisieren
    Style = GetWindowLongPtr(prophWnd, GWL_STYLE)
    Style = IIf(propFullRowSelect, Style Or TreeViewStyleConsts.TVS_FULLROWSELECT, Style And Not TreeViewStyleConsts.TVS_FULLROWSELECT)
    SetWindowLongPtr prophWnd, GWL_STYLE, Style
    If Not inIDE Then
      InvalidateRectAsLong prophWnd, 0, 1
    End If
  End If

  PropertyChanged "FullRowSelect"
End Property

' hDC des TreeViews
Public Property Get hDC() As Long
Attribute hDC.VB_Description = "Gibt das Handle des DeviceContexts des TreeView-Fensters zurück."
Attribute hDC.VB_ProcData.VB_Invoke_Property = ";Daten"
Attribute hDC.VB_MemberFlags = "400"
  hDC = prophDC
End Property

' Handle des Items, der gerade umbenannt wird
Public Property Get hEditetItem() As Long
Attribute hEditetItem.VB_Description = "Gibt den Item zurück, der gerade umbenannt wird."
Attribute hEditetItem.VB_ProcData.VB_Invoke_Property = ";Daten"
Attribute hEditetItem.VB_MemberFlags = "40"
  hEditetItem = prophEditetItem
End Property

' Handle der ImageList
Public Property Get hImageList() As Long
Attribute hImageList.VB_Description = "Gibt das Handle der ImageList des Controls zurück. Da die SystemImageList genutzt wird, sollte sie nicht freigegeben werden."
Attribute hImageList.VB_ProcData.VB_Invoke_Property = ";Daten"
Attribute hImageList.VB_MemberFlags = "400"
  hImageList = prophImageList
End Property

Public Property Let hImageList(ByVal New_hImageList As Long)
  prophImageList = New_hImageList
  If prophWnd Then
    SendMessageAsLong prophWnd, TVM_SETIMAGELIST, TVSIL_NORMAL, prophImageList
  End If
End Property

Public Property Get HotBackColor() As OLE_COLOR
Attribute HotBackColor.VB_Description = "Bestimmt die Hintergrundfarbe, die für denjenigen Item verwendet wird, über dem sich der Mauszeiger befindet. Bei -1 wird die Hintergrundfarbe des  Systems genutzt."
Attribute HotBackColor.VB_ProcData.VB_Invoke_Property = ";Farben"
  HotBackColor = propHotBackColor
End Property

Public Property Let HotBackColor(ByVal New_HotBackColor As OLE_COLOR)
  If New_HotBackColor = propHotBackColor Then Exit Property

  propHotBackColor = New_HotBackColor

  If Not inIDE Then
    mustCustomDraw = IsCustomDrawNecessary
    Refresh
  End If

  PropertyChanged "HotBackColor"
End Property

Public Property Get HotForeColor() As OLE_COLOR
Attribute HotForeColor.VB_Description = "Bestimmt die Schriftfarbe, die für denjenigen Item verwendet wird, über dem sich der Mauszeiger befindet. Bei -1 wird die Standardfarbe des  Systems genutzt."
Attribute HotForeColor.VB_ProcData.VB_Invoke_Property = ";Farben"
  HotForeColor = propHotForeColor
End Property

Public Property Let HotForeColor(ByVal New_HotForeColor As OLE_COLOR)
  If New_HotForeColor = propHotForeColor Then Exit Property

  propHotForeColor = New_HotForeColor

  If Not inIDE Then
    mustCustomDraw = IsCustomDrawNecessary
    Refresh
  End If

  PropertyChanged "HotForeColor"
End Property

Public Property Get HotTracking() As Boolean
Attribute HotTracking.VB_Description = "Gibt an, ob der Item unter dem Mauszeiger automatisch markiert wird."
Attribute HotTracking.VB_ProcData.VB_Invoke_Property = ";Verhalten"
  Dim Style As Long

  If prophWnd Then
    ' aktuellen Wert ermitteln
    Style = GetWindowLongPtr(prophWnd, GWL_STYLE)
    propHotTracking = ((Style And TreeViewStyleConsts.TVS_TRACKSELECT) = TreeViewStyleConsts.TVS_TRACKSELECT)
  End If

  HotTracking = propHotTracking
End Property

Public Property Let HotTracking(ByVal New_HotTracking As Boolean)
  Dim Style As Long

  propHotTracking = New_HotTracking
  mustCustomDraw = IsCustomDrawNecessary

  If prophWnd Then
    ' Control aktualisieren
    Style = GetWindowLongPtr(prophWnd, GWL_STYLE)
    Style = IIf(propHotTracking, Style Or TreeViewStyleConsts.TVS_TRACKSELECT, Style And Not TreeViewStyleConsts.TVS_TRACKSELECT)
    SetWindowLongPtr prophWnd, GWL_STYLE, Style
    SetWindowPos prophWnd, 0, 0, 0, 0, 0, SWP_DRAWFRAME Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER
  End If

  PropertyChanged "HotTracking"
End Property

Public Property Get HoverTime() As Long
Attribute HoverTime.VB_Description = "Bestimmt die Zeit, die der Mauszeiger über dem Control sein muß, damit das MouseHover-Event ausgelöst wird."
Attribute HoverTime.VB_ProcData.VB_Invoke_Property = ";Verhalten"
  HoverTime = propHoverTime
End Property

Public Property Let HoverTime(ByVal New_HoverTime As Long)
  If (New_HoverTime < 1) And (New_HoverTime <> -1) Then
    Err.Raise 380
    Exit Property
  End If
  propHoverTime = New_HoverTime

  PropertyChanged "HoverTime"
End Property

' Handle des Shell-Kontextmenüs
Public Property Get hShellContextMenu() As Long
Attribute hShellContextMenu.VB_Description = "Gibt das Handle des Shell-KontextMenüs zurück."
Attribute hShellContextMenu.VB_ProcData.VB_Invoke_Property = ";Daten"
Attribute hShellContextMenu.VB_MemberFlags = "400"
  hShellContextMenu = prophShellContextMenu
End Property

' Handles der State-ImageLists
Public Property Get hStateImageList(ByVal Large As Boolean) As Long
Attribute hStateImageList.VB_Description = "Gibt das Handle der StateImageList des Controls zurück. Diese enthält die Icons für die CheckBoxes."
Attribute hStateImageList.VB_ProcData.VB_Invoke_Property = ";Daten"
Attribute hStateImageList.VB_MemberFlags = "400"
  If prophWnd Then
    ' aktuellen Wert ermitteln
    If Not inIDE Then
      If Large Then
        prophLargeStateImageList = SendMessageAsLong(prophWnd, TVM_GETIMAGELIST, TVSIL_STATE, 0)
      Else
        prophSmallStateImageList = SendMessageAsLong(prophWnd, TVM_GETIMAGELIST, TVSIL_STATE, 0)
      End If
    End If
  End If

  hStateImageList = IIf(Large, prophLargeStateImageList, prophSmallStateImageList)
End Property

' Handles der State-ImageLists
Public Property Let hStateImageList(ByVal Large As Boolean, ByVal New_hStateImageList As Long)
  If Large Then
    If New_hStateImageList = prophLargeStateImageList Then Exit Property
  Else
    If New_hStateImageList = prophSmallStateImageList Then Exit Property
  End If

  If prophWnd Then
    If Large = useLargeIcons Then
      If New_hStateImageList Then
        CheckBoxes = True
        SendMessageAsLong prophWnd, TVM_SETIMAGELIST, TVSIL_STATE, New_hStateImageList
      Else
        SendMessageAsLong prophWnd, TVM_SETIMAGELIST, TVSIL_STATE, 0
      End If
    End If
  End If

  If Large Then
    If prophLargeStateImageList Then ImageList_Destroy prophLargeStateImageList
    prophLargeStateImageList = New_hStateImageList
  Else
    If prophSmallStateImageList Then ImageList_Destroy prophSmallStateImageList
    prophSmallStateImageList = New_hStateImageList
  End If
End Property

' hWnd des TreeViews
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Gibt das Handle des TreeView-Fensters zurück."
Attribute hWnd.VB_ProcData.VB_Invoke_Property = ";Daten"
Attribute hWnd.VB_UserMemId = -515
Attribute hWnd.VB_MemberFlags = "400"
  hWnd = prophWnd
End Property

' Handle der TextBox für Renaming
Public Property Get hWndEdit() As Long
Attribute hWndEdit.VB_Description = "Gibt das Handle der TextBox für Renaming zurück."
Attribute hWndEdit.VB_ProcData.VB_Invoke_Property = ";Daten"
Attribute hWndEdit.VB_MemberFlags = "400"
  If prophWnd Then hWndEdit = SendMessageAsLong(prophWnd, TVM_GETEDITCONTROL, 0, 0)
End Property

' Handle des Parent-Fensters für UI der Shell-Interfaces
Public Property Get hWndShellUIParentWindow() As Long
Attribute hWndShellUIParentWindow.VB_Description = "Gibt das Handle des Fensters zurück, welches als Elternfenster für Meldungen der Shell-API genutzt wird."
Attribute hWndShellUIParentWindow.VB_ProcData.VB_Invoke_Property = ";Daten"
Attribute hWndShellUIParentWindow.VB_MemberFlags = "400"
  hWndShellUIParentWindow = prophWndShellUIParentWindow
End Property

Public Property Let hWndShellUIParentWindow(ByVal New_hWndShellUIParentWindow As Long)
  prophWndShellUIParentWindow = New_hWndShellUIParentWindow
End Property

' Handle des ToolTip-Controls
Public Property Get hWndToolTip() As Long
Attribute hWndToolTip.VB_Description = "Gibt das Handle des ToolTip-Controls zurück."
Attribute hWndToolTip.VB_ProcData.VB_Invoke_Property = ";Daten"
Attribute hWndToolTip.VB_MemberFlags = "400"
  If prophWnd Then hWndToolTip = SendMessageAsLong(prophWnd, TVM_GETTOOLTIPS, 0, 0)
End Property

Public Property Get IncludedItems() As IncludedItemsConstants
Attribute IncludedItems.VB_Description = "Gibt an, welche Typen von Shell-Items angezeigt werden."
Attribute IncludedItems.VB_ProcData.VB_Invoke_Property = ";Filter"
Attribute IncludedItems.VB_MemberFlags = "400"
  IncludedItems = propIncludedItems
End Property

Public Property Let IncludedItems(ByVal New_IncludedItems As IncludedItemsConstants)
  If Not isValidNumber(New_IncludedItems, 0, IncludedItemsConstants.iiFSFiles Or IncludedItemsConstants.iiFSFolders Or IncludedItemsConstants.iiNonFSFiles Or IncludedItemsConstants.iiNonFSFolders) Then
    Err.Raise 380
    Exit Property
  End If
  If New_IncludedItems = propIncludedItems Then Exit Property

  propIncludedItems = New_IncludedItems

  If Not inIDE Then
    If prophWnd Then
      ReloadItems
    End If
  End If

  PropertyChanged "IncludedItems"
End Property

Public Property Get Indent() As Long
Attribute Indent.VB_Description = "Gibt die Größe des Einzugs der Items an."
Attribute Indent.VB_ProcData.VB_Invoke_Property = ";Darstellung"
  If prophWnd Then
    ' aktuellen Wert ermitteln
    propIndent = SendMessageAsLong(prophWnd, TVM_GETINDENT, 0, 0)
  End If

  Indent = propIndent
End Property

Public Property Let Indent(ByVal New_Indent As Long)
  If Not isValidNumber(New_Indent, 1, 1000) Then
    Err.Raise 380
    Exit Property
  End If
  propIndent = New_Indent

  If prophWnd Then
    ' das Control anpassen
    SendMessageAsLong prophWnd, TVM_SETINDENT, propIndent, 0
  End If

  PropertyChanged "Indent"
End Property

Public Property Get ItemGhosted(ByVal hItem As Long) As Boolean
Attribute ItemGhosted.VB_Description = "Gibt an, ob der Item <hItem> abgeblendet ist."
Attribute ItemGhosted.VB_ProcData.VB_Invoke_Property = ";Darstellung"
Attribute ItemGhosted.VB_MemberFlags = "400"
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Property

  ItemGhosted = (GetItemState_Fast(hItem, TVIS_CUT) And TVIS_CUT)
End Property

Public Property Let ItemGhosted(ByVal hItem As Long, ByVal New_ItemGhosted As Boolean)
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Property

  SetItemState hItem, IIf(New_ItemGhosted, TVIS_CUT, 0), TVIS_CUT
End Property

Public Property Get ItemHeight() As Long
Attribute ItemHeight.VB_Description = "Gibt die Höhe eines Items an."
Attribute ItemHeight.VB_ProcData.VB_Invoke_Property = ";Darstellung"
  If prophWnd Then
    ' aktuellen Wert ermitteln
    propItemHeight = SendMessageAsLong(prophWnd, TVM_GETITEMHEIGHT, 0, 0)
  End If

  ItemHeight = propItemHeight
End Property

Public Property Let ItemHeight(ByVal New_ItemHeight As Long)
  Dim Size_LargeIcons As Long
  Dim Size_SmallIcons As Long

  If Not isValidNumber(New_ItemHeight, 1, 1000) Then
    Err.Raise 380
    Exit Property
  End If
  propItemHeight = New_ItemHeight

  If prophWnd Then
    ' das Control anpassen
    If propItemHeight = 0 Then propItemHeight = 16
    SendMessageAsLong prophWnd, TVM_SETITEMHEIGHT, propItemHeight, 0

    ' ImageLists anpassen
    If propAlwaysUseSmallIcons Then
      useLargeIcons = False
    Else
      Size_LargeIcons = GetSystemMetrics(SM_CXICON)
      Size_SmallIcons = GetSystemMetrics(SM_CXSMICON)
      useLargeIcons = (propItemHeight >= Size_LargeIcons * 0.75)
    End If

    If (prophImageList <> 0) And Not propUseCustomIcons Then
      ' ImageList neu setzen
      prophImageList = GetSysImageList(useLargeIcons)
      SendMessageAsLong prophWnd, TVM_SETIMAGELIST, TVSIL_NORMAL, prophImageList
    End If
    If propCheckBoxes Then SendMessageAsLong prophWnd, TVM_SETIMAGELIST, TVSIL_STATE, IIf(useLargeIcons, prophLargeStateImageList, prophSmallStateImageList)
  End If

  PropertyChanged "ItemHeight"
End Property

Public Property Get ItemIconIndex(ByVal hItem As Long, ByVal Selected As Boolean) As Long
Attribute ItemIconIndex.VB_Description = "Gibt den Index des Icons des Items <hItem> an."
Attribute ItemIconIndex.VB_MemberFlags = "400"
  ItemIconIndex = GetItemIconIndex(hItem, Selected)
End Property

Public Property Let ItemIconIndex(ByVal hItem As Long, ByVal Selected As Boolean, ByVal New_IconIndex As Long)
  SetItemIcon hItem, New_IconIndex, Selected
End Property

Public Property Get ItemOverlayIndex(ByVal hItem As Long) As Long
  ItemOverlayIndex = GetItemOverlayIndex(hItem)
End Property

Public Property Let ItemOverlayIndex(ByVal hItem As Long, ByVal New_OverlayIndex As Long)
  SetItemOverlay hItem, New_OverlayIndex
End Property

Public Property Get ItemStateIconIndex(ByVal hItem As Long) As Long
Attribute ItemStateIconIndex.VB_Description = "Gibt den Index des StateIcons des Items <hItem> an."
Attribute ItemStateIconIndex.VB_ProcData.VB_Invoke_Property = ";Icons"
Attribute ItemStateIconIndex.VB_MemberFlags = "400"
  Dim ret As Long

  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Property

  ret = GetItemState_Fast(hItem, TVIS_STATEIMAGEMASK) And TVIS_STATEIMAGEMASK
  ret = STATEIMAGEMASKTOINDEX(ret)

  ItemStateIconIndex = ret
End Property

Public Property Let ItemStateIconIndex(ByVal hItem As Long, ByVal New_ItemsStateIconIndex As Long)
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Property

  SetItemState hItem, INDEXTOSTATEIMAGEMASK(New_ItemsStateIconIndex), TVIS_STATEIMAGEMASK
End Property

Public Property Get LimitEditInput() As Boolean
Attribute LimitEditInput.VB_Description = "Gibt an, ob beim Umbenennen von Items der neue Dateiname schon während der Eingabe auf ungültige Zeichen überprüft wird."
Attribute LimitEditInput.VB_ProcData.VB_Invoke_Property = ";Verhalten"
  LimitEditInput = propLimitEditInput
End Property

Public Property Let LimitEditInput(ByVal New_LimitEditInput As Boolean)
  propLimitEditInput = New_LimitEditInput

  PropertyChanged "LimitEditInput"
End Property

Public Property Get LineColor() As OLE_COLOR
Attribute LineColor.VB_Description = "Gibt die Linienfarbe des Controls an."
Attribute LineColor.VB_ProcData.VB_Invoke_Property = ";Farben"
  Dim RGBClr As Long

  If prophWnd Then
    ' aktuellen Wert ermitteln
    RGBClr = SendMessageAsLong(prophWnd, TVM_GETLINECOLOR, 0, 0)
    If RGBClr <> TranslateColor(propLineColor) Then propLineColor = RGBClr
  End If

  LineColor = propLineColor
End Property

Public Property Let LineColor(ByVal New_LineColor As OLE_COLOR)
  Dim RGBClr As Long

  propLineColor = New_LineColor

  If prophWnd Then
    ' Linienfarbe setzen
    RGBClr = TranslateColor(propLineColor)
    SendMessageAsLong prophWnd, TVM_SETLINECOLOR, 0, RGBClr
  End If

  PropertyChanged "LineColor"
End Property

Public Property Get LineStyle() As LineStyleConstants
Attribute LineStyle.VB_Description = "Gibt an wie die Linien zwischen den Items gezeichnet werden."
Attribute LineStyle.VB_ProcData.VB_Invoke_Property = ";Darstellung"
  Dim Style As Long

  If prophWnd Then
    ' aktuellen Wert ermitteln
    Style = GetWindowLongPtr(prophWnd, GWL_STYLE)
    propLineStyle = IIf(Style And TreeViewStyleConsts.TVS_LINESATROOT, LineStyleConstants.lsLinesAtRoot, LineStyleConstants.lsLinesAtItem)
  End If

  LineStyle = propLineStyle
End Property

Public Property Let LineStyle(ByVal New_LineStyle As LineStyleConstants)
  Dim Style As Long

  If Not isValidNumber(New_LineStyle, LineStyleConstants.lsLinesAtRoot, LineStyleConstants.lsLinesAtItem) Then
    Err.Raise 380
    Exit Property
  End If
  propLineStyle = New_LineStyle

  If prophWnd Then
    ' Control aktualisieren
    Style = GetWindowLongPtr(prophWnd, GWL_STYLE)
    Style = IIf(propLineStyle = LineStyleConstants.lsLinesAtRoot, Style Or TreeViewStyleConsts.TVS_LINESATROOT, Style And Not TreeViewStyleConsts.TVS_LINESATROOT)
    SetWindowLongPtr prophWnd, GWL_STYLE, Style
  End If

  PropertyChanged "LineStyle"
End Property

Public Property Get Margin() As Long
Attribute Margin.VB_Description = "Gibt die Breite des linken und rechten Randes der Items an. Standard ist 0 für Windows NT 4 und 3 für neuere Windows-Versionen."
Attribute Margin.VB_ProcData.VB_Invoke_Property = ";Darstellung"
  If prophWnd Then
    ' aktuellen Wert ermitteln
    propMargin = SendMessageAsLong(prophWnd, TVM_GETBORDER, 0, 0)
  End If

  Margin = propMargin
End Property

Public Property Let Margin(ByVal New_Margin As Long)
  If New_Margin = propMargin Then Exit Property

  propMargin = New_Margin

  If prophWnd Then
    SendMessageAsLong prophWnd, TVM_SETBORDER, &H1, propMargin
    InvalidateRectAsLong prophWnd, 0, 1
  End If

  PropertyChanged "Margin"
End Property

Public Property Get MaxScrollTime() As Long
Attribute MaxScrollTime.VB_Description = "Gibt an wie lange ein Scroll-Vorgang maximal dauern darf (in Millisekunden)."
Attribute MaxScrollTime.VB_ProcData.VB_Invoke_Property = ";Verhalten"
  If prophWnd Then
    ' aktuellen Wert ermitteln
    propMaxScrollTime = SendMessageAsLong(prophWnd, TVM_GETSCROLLTIME, 0, 0)
  End If

  MaxScrollTime = propMaxScrollTime
End Property

Public Property Let MaxScrollTime(ByVal New_MaxScrollTime As Long)
  If Not isValidNumber(New_MaxScrollTime, 100, 5000) Then
    Err.Raise 380
    Exit Property
  End If
  propMaxScrollTime = New_MaxScrollTime

  If prophWnd Then SendMessageAsLong prophWnd, TVM_SETSCROLLTIME, propMaxScrollTime, 0

  PropertyChanged "MaxScrollTime"
End Property

Public Property Get MouseIcon() As StdPicture
Attribute MouseIcon.VB_Description = "Legt einen benutzerdefinierten Mauszeiger fest."
Attribute MouseIcon.VB_ProcData.VB_Invoke_Property = ";Darstellung"
  Set MouseIcon = propMouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As StdPicture)
  Set propMouseIcon = New_MouseIcon

  Set UserControl.MouseIcon = propMouseIcon

  PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Gibt den Mauszeiger für das Control an."
Attribute MousePointer.VB_ProcData.VB_Invoke_Property = ";Darstellung"
  MousePointer = propMousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
  propMousePointer = New_MousePointer

  UserControl.MousePointer = propMousePointer

  PropertyChanged "MousePointer"
End Property

Public Property Get path() As String
Attribute path.VB_Description = "Gibt den markierten Item an."
Attribute path.VB_ProcData.VB_Invoke_Property = ";Verhalten"
Attribute path.VB_MemberFlags = "600"
  path = Replace(propPath, ArchiveSignature, ":")
End Property

Public Property Let path(ByVal New_Path As String)
  Dim ArchivePart As String
  Dim i As Integer
  Dim pos As Integer
  Dim Sep As String
  Dim txt As String

  If New_Path = propPath Then Exit Property

  #If Debuging Then
    debugger.AddLogEntry "Setting Path property to " & New_Path, LogEntryTypeConstants.letInfo
  #End If

  Sep = IIf(IsFTPURL(New_Path), "/", "\")
  ' <New_Path> konvertieren
  New_Path = RemoveBackslash(New_Path, False, Sep)
  ' prüfen, ob Archiv-Pfade angehängt sind
  ' -> evt. kennzeichnen
  If InStr(New_Path, ":") Then
    ' sicher stellen, daß es nicht nur eine Laufwerksangabe ist
    If Not IncludesCLSID(New_Path) Then
      If Mid$(New_Path, InStrRev(New_Path, ":") + 1, 1) <> Sep Then
        ' sehr wahrscheinlich ein Archiv-Item
        pos = InStrRev(New_Path, ":")
        New_Path = Left$(New_Path, pos - 1) & ArchiveSignature & Mid$(New_Path, pos + 1)
      End If
    End If
  End If

  ' prüfen, ob Archiv-Pfade angehängt sind
  ' -> abtrennen
  If InStr(New_Path, ArchiveSignature) Then
    pos = InStr(New_Path, ArchiveSignature)
    ArchivePart = Mid$(New_Path, pos)
    New_Path = Left$(New_Path, pos - 1)
  End If

  Select Case True
    Case IncludesCLSID(New_Path)
      ' ein Objekt - so lassen
    Case IsDirectory(New_Path)
      ' ein FS-Folder, der sogar existiert - so lassen
    Case FileExists(New_Path)
      ' eine FS-File, die sogar existiert - so lassen
    Case IsFTPURL(New_Path)
      ' eine FTP-URL - so lassen
    #If Debuging Then
      Case IsCSIDL(debugger, New_Path)
        ' eine "echte" CSIDL - so lassen
      Case IsCSIDL(debugger, New_Path, True)
        ' ein Eintrag aus dem Kontextmenü - so lassen
    #Else
      Case IsCSIDL(New_Path)
        ' eine "echte" CSIDL - so lassen
      Case IsCSIDL(New_Path, True)
        ' ein Eintrag aus dem Kontextmenü - so lassen
    #End If
    Case Else
      ' evt. ein Pfad aus DisplayNames
      ' -> nur gültig für den Desktop
      '    -> müssen mit "Desktop" beginnen
      #If Debuging Then
        txt = pIDLToDisplayName_Light(debugger, pIDL_Desktop)
      #Else
        txt = pIDLToDisplayName_Light(pIDL_Desktop)
      #End If
      If LCase$(Left$(New_Path, Len(txt))) = LCase$(txt) Then
        ' Pfad aus DisplayNames - so lassen
      Else
        ' möglicherweise ein Pfad, an den DisplayNames angehängt sind
        ' es kann auch ein ungültiger Pfad sein - trotzdem akzeptieren
        txt = New_Path
        For i = 1 To CountSegments(New_Path)
          txt = RemoveLastFolders(New_Path, i)
          If IsDirectory(txt) Then
            ' Pfad mit DisplayNames oder ungültiger Pfad
            ' -> so lassen
            Exit For
          End If
        Next
        If txt = "" Then
          ' ungültiger Wert
          Exit Property
        End If
      End If
  End Select

  ' Archiv-Pfad wieder anhängen
  New_Path = New_Path & ArchivePart

  If New_Path <> "" Then propPath = New_Path
  If propPath = "" Then propPath = propRootPath

  If prophWnd Then SelectPath propPath

  PropertyChanged "Path"
End Property

Public Property Get ProcessContextMenuKeys() As Boolean
Attribute ProcessContextMenuKeys.VB_Description = "Gibt an, ob das Control das KontextMenü des markierten Items anzeigt, wenn die ""Kontextmenü""-Taste oder die Tasten [SHIFT]-[F10] gedrückt werden."
Attribute ProcessContextMenuKeys.VB_ProcData.VB_Invoke_Property = ";Verhalten"
  ProcessContextMenuKeys = propProcessContextMenuKeys
End Property

Public Property Let ProcessContextMenuKeys(ByVal New_ProcessContextMenuKeys As Boolean)
  propProcessContextMenuKeys = New_ProcessContextMenuKeys

  PropertyChanged "ProcessContextMenuKeys"
End Property

Public Property Get Renaming() As RenamingConstants
Attribute Renaming.VB_Description = "Gibt an, ob Items vom Benutzer umbenannt werden können."
Attribute Renaming.VB_ProcData.VB_Invoke_Property = ";Verhalten"
  Dim Style As Long

  If prophWnd Then
    ' aktuellen Wert ermitteln
    Style = GetWindowLongPtr(prophWnd, GWL_STYLE)
    propRenaming = IIf(Style And TreeViewStyleConsts.TVS_EDITLABELS, RenamingConstants.rAutomatic, RenamingConstants.rManual)
  End If

  Renaming = propRenaming
End Property

Public Property Let Renaming(ByVal New_Renaming As RenamingConstants)
  Dim Style As Long

  If Not isValidNumber(New_Renaming, RenamingConstants.rAutomatic, RenamingConstants.rManual) Then
    Err.Raise 380
    Exit Property
  End If
  propRenaming = New_Renaming

  If prophWnd Then
    ' Control aktualisieren
    Style = GetWindowLongPtr(prophWnd, GWL_STYLE)
    Style = IIf(propRenaming = RenamingConstants.rAutomatic, Style Or TreeViewStyleConsts.TVS_EDITLABELS, Style And Not TreeViewStyleConsts.TVS_EDITLABELS)
    SetWindowLongPtr prophWnd, GWL_STYLE, Style
  End If

  PropertyChanged "Renaming"
End Property

Public Property Get ReplaceHandCursor() As Boolean
Attribute ReplaceHandCursor.VB_Description = "Gibt an, ob im HotTracking-Modus der Mauszeiger durch die Eigenschaften <MouseIcon> und <MousePointer> festgelegt wird."
Attribute ReplaceHandCursor.VB_ProcData.VB_Invoke_Property = ";Verhalten"
  ReplaceHandCursor = propReplaceHandCursor
End Property

Public Property Let ReplaceHandCursor(ByVal New_ReplaceHandCursor As Boolean)
  If New_ReplaceHandCursor = propReplaceHandCursor Then Exit Property

  propReplaceHandCursor = New_ReplaceHandCursor

  PropertyChanged "ReplaceHandCursor"
End Property

Public Property Get RootPath() As String
Attribute RootPath.VB_Description = "Gibt den Startitem an."
Attribute RootPath.VB_ProcData.VB_Invoke_Property = ";Verhalten"
Attribute RootPath.VB_MemberFlags = "400"
  RootPath = propRootPath
End Property

Public Property Let RootPath(ByVal New_RootPath As String)
  Const HEAP_ZERO_MEMORY = &H8
  Dim hHeap As Long
  Dim i As Integer
  Dim pBuffer As Long
  Dim pos As Integer
  Dim txt As String

  If New_RootPath = propRootPath Then Exit Property

  #If Debuging Then
    debugger.AddLogEntry "Setting RootPath property", LogEntryTypeConstants.letInfo
    debugger.AddLogEntry "   " & New_RootPath, LogEntryTypeConstants.letOther
  #End If

  Select Case True
    Case IncludesCLSID(New_RootPath)
      ' ein Objekt - so lassen
    Case IsDirectory(New_RootPath)
      ' ein FS-Folder, der sogar existiert - so lassen
    Case IsURL(New_RootPath)
      ' eine URL
      ' ToDo: Support für URLs
      Exit Property
    #If Debuging Then
      Case IsCSIDL(debugger, New_RootPath)
        ' eine "echte" CSIDL - so lassen
      Case IsCSIDL(debugger, New_RootPath, True)
        ' ein Eintrag aus dem Kontextmenü - so lassen
    #Else
      Case IsCSIDL(New_RootPath)
        ' eine "echte" CSIDL - so lassen
      Case IsCSIDL(New_RootPath, True)
        ' ein Eintrag aus dem Kontextmenü - so lassen
    #End If
    Case Else
      ' evt. ein Pfad aus DisplayNames
      ' -> nur gültig für den Desktop
      '    -> müssen mit "Desktop" beginnen
      #If Debuging Then
        txt = pIDLToDisplayName_Light(debugger, pIDL_Desktop)
      #Else
        txt = pIDLToDisplayName_Light(pIDL_Desktop)
      #End If
      If LCase$(Left$(New_RootPath, Len(txt))) = LCase$(txt) Then
        ' Pfad aus DisplayNames - so lassen
      Else
        ' möglicherweise ein Pfad, an den DisplayNames angehängt sind
        ' es kann auch ein ungültiger Pfad sein - trotzdem akzeptieren
        txt = New_RootPath
        For i = 1 To CountSegments(New_RootPath)
          txt = RemoveLastFolders(New_RootPath, i)
          If IsDirectory(txt) Then
            ' Pfad mit DisplayNames oder ungültiger Pfad
            ' -> so lassen
            Exit For
          End If
        Next
        If txt = "" Then
          ' ungültiger Wert
          Exit Property
        End If
      End If
  End Select

  If New_RootPath = "" Then
    New_RootPath = "Desktop"
  End If

  #If Debuging Then
    debugger.AddLogEntry "   " & New_RootPath, LogEntryTypeConstants.letOther
  #End If

  If inIDE Then
    propRootPath = New_RootPath
  ElseIf prophWnd Then
    hHeap = GetProcessHeap
    If hHeap Then
      pBuffer = HeapAlloc(hHeap, HEAP_ZERO_MEMORY, Len(New_RootPath) + 1)
    End If
    If pBuffer Then
      #If Debuging Then
        debugger.AddLogEntry "   Posting EXTVM_SETROOTPATH", LogEntryTypeConstants.letOther
      #End If
      lstrcpyAsLong1 pBuffer, New_RootPath
      PostMessage prophWnd, EXTVM_SETROOTPATH, lstrlenAsLong(pBuffer), pBuffer
      #If Debuging Then
        debugger.AddLogEntry "      ...done", LogEntryTypeConstants.letOther
      #End If
    Else
      propRootPath = New_RootPath
      #If Debuging Then
        debugger.AddLogEntry "   Will reload items now!", LogEntryTypeConstants.letOther
      #End If
      ReloadItems
    End If
  Else
    propRootPath = New_RootPath
  End If

  PropertyChanged "RootPath"
End Property

Public Property Get RTLReading() As Boolean
Attribute RTLReading.VB_Description = "Gibt an, ob auf einem bidirektionalen System der Text von rechts nach links ausgegeben wird."
Attribute RTLReading.VB_ProcData.VB_Invoke_Property = ";Verhalten"
  Dim Style As Long

  If prophWnd Then
    ' aktuellen Wert ermitteln
    Style = GetWindowLongPtr(prophWnd, GWL_STYLE)
    propRTLReading = ((Style And TreeViewStyleConsts.TVS_RTLREADING) = TreeViewStyleConsts.TVS_RTLREADING)
  End If

  RTLReading = propRTLReading
End Property

Public Property Let RTLReading(ByVal New_RTLReading As Boolean)
  Dim Style As Long

  propRTLReading = New_RTLReading

  ' das Control neu erzeugen...
  If prophWnd Then CreateControl
'  If prophWnd Then
'    ' Control aktualisieren
'    Style = GetWindowLongPtr(prophWnd, GWL_STYLE)
'    Style = IIf(propRTLReading, Style Or TreeViewStyleConsts.TVS_RTLREADING, Style And Not TreeViewStyleConsts.TVS_RTLREADING)
'    SetWindowLongPtr prophWnd, GWL_STYLE, Style
'    Style = GetWindowLongPtr(prophWnd, GWL_EXSTYLE)
'    Style = IIf(propRTLReading, Style Or ExWindowStyleConsts.WS_EX_LAYOUTRTL, Style And Not ExWindowStyleConsts.WS_EX_LAYOUTRTL)
'    SetWindowLongPtr prophWnd, GWL_EXSTYLE, Style
'    Refresh
'  End If

  PropertyChanged "RTLReading"
End Property

Public Property Get ScrollBars() As ScrollBarsConstants
Attribute ScrollBars.VB_Description = "Gibt an, ob bei Bedarf Scrollbalken angezeigt werden und wie diese aussehen."
Attribute ScrollBars.VB_ProcData.VB_Invoke_Property = ";Verhalten"
  Dim Style As Long

  If prophWnd Then
    ' aktuellen Wert ermitteln
    Style = GetWindowLongPtr(prophWnd, GWL_STYLE)
    propScrollBars = IIf(Style And TreeViewStyleConsts.TVS_NOSCROLL, ScrollBarsConstants.sbNone, ScrollBarsConstants.sbNormal)
  End If

  ScrollBars = propScrollBars
End Property

Public Property Let ScrollBars(ByVal New_ScrollBars As ScrollBarsConstants)
  Dim Style As Long

  New_ScrollBars = Abs(New_ScrollBars)
  If Not isValidNumber(New_ScrollBars, ScrollBarsConstants.sbNone, ScrollBarsConstants.sbNormal) Then
    Err.Raise 380
    Exit Property
  End If
  If New_ScrollBars = propScrollBars Then Exit Property

  propScrollBars = New_ScrollBars

  If prophWnd Then
    ' Control aktualisieren
    Style = GetWindowLongPtr(prophWnd, GWL_STYLE)
    Style = IIf(propScrollBars = ScrollBarsConstants.sbNone, Style Or TreeViewStyleConsts.TVS_NOSCROLL, Style And Not TreeViewStyleConsts.TVS_NOSCROLL)
    SetWindowLongPtr prophWnd, GWL_STYLE, Style
    SetWindowPos prophWnd, 0, 0, 0, 0, 0, SWP_DRAWFRAME Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER
    UserControl_Resize
  End If

  PropertyChanged "ScrollBars"
End Property
'
'Public Property Get Selected(ByVal hItem As Long) As Boolean
'  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Property
'
'  Selected = (getItemState_Fast(hItem, TVIS_SELECTED) And TVIS_SELECTED)
'End Property
'
'Public Property Let Selected(ByVal hItem As Long, ByVal New_Selected As Boolean)
'  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Property
'
'  setItemState hItem, IIf(New_Selected, TVIS_SELECTED, 0), TVIS_SELECTED
'End Property

Public Property Get SelChangedDelayTime() As Long
Attribute SelChangedDelayTime.VB_Description = "Gibt an, wie lange gewartet wird, bevor ein SelChanged-Event, welches durch Tastatureingaben entstand, ausgelöst wird (in Millisekunden)."
Attribute SelChangedDelayTime.VB_ProcData.VB_Invoke_Property = ";Verhalten"
  SelChangedDelayTime = propSelChangedDelayTime
End Property

Public Property Let SelChangedDelayTime(ByVal New_SelChangedDelayTime As Long)
  If Not isValidNumber(New_SelChangedDelayTime, 0, 10000) Then
    Err.Raise 380
    Exit Property
  End If
  propSelChangedDelayTime = New_SelChangedDelayTime

  PropertyChanged "SelChangedDelayTime"
End Property

Public Property Get SelectedItem() As Long
  Dim ret As Long

  ret = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_CARET, 0)

  If ret = 0 Then ret = -1
  SelectedItem = ret
End Property

Public Property Let SelectedItem(ByVal New_SelectedItem As Long)
Attribute SelectedItem.VB_Description = "Bestimmt den aktuellen Item."
Attribute SelectedItem.VB_ProcData.VB_Invoke_PropertyPut = ";Daten"
Attribute SelectedItem.VB_MemberFlags = "400"
  If Not ItemExists(New_SelectedItem, AllowNullConstants.anNever) Then Exit Property

  dontExpand = True
  SendMessageAsLong prophWnd, TVM_SELECTITEM, TVGN_CARET, New_SelectedItem
  dontExpand = False
End Property

Public Property Get SelectedItemBackColor() As OLE_COLOR
Attribute SelectedItemBackColor.VB_Description = "Gibt die Hintergrundfarbe des markierten Items an. Bei -1 wird die Standardfarbe des  Systems genutzt."
Attribute SelectedItemBackColor.VB_ProcData.VB_Invoke_Property = ";Farben"
  SelectedItemBackColor = propSelectedItemBackColor
End Property

Public Property Let SelectedItemBackColor(ByVal New_SelectedItemBackColor As OLE_COLOR)
  If New_SelectedItemBackColor = propSelectedItemBackColor Then Exit Property

  propSelectedItemBackColor = New_SelectedItemBackColor

  If Not inIDE Then
    mustCustomDraw = IsCustomDrawNecessary
    Refresh
  End If

  PropertyChanged "SelectedItemBackColor"
End Property

Public Property Get SelectedItemForeColor() As OLE_COLOR
Attribute SelectedItemForeColor.VB_Description = "Gibt die Schriftfarbe des markierten Items an. Bei -1 wird die Standardfarbe des  Systems genutzt."
Attribute SelectedItemForeColor.VB_ProcData.VB_Invoke_Property = ";Farben"
  SelectedItemForeColor = propSelectedItemForeColor
End Property

Public Property Let SelectedItemForeColor(ByVal New_SelectedItemForeColor As OLE_COLOR)
  If New_SelectedItemForeColor = propSelectedItemForeColor Then Exit Property

  propSelectedItemForeColor = New_SelectedItemForeColor

  If Not inIDE Then
    mustCustomDraw = IsCustomDrawNecessary
    Refresh
  End If

  PropertyChanged "SelectedItemForeColor"
End Property

Public Property Get ShowFocusRect() As Boolean
Attribute ShowFocusRect.VB_Description = "Not supported anymore. Left for backward compatibility."
Attribute ShowFocusRect.VB_ProcData.VB_Invoke_Property = ";Darstellung"
Attribute ShowFocusRect.VB_MemberFlags = "440"
  ShowFocusRect = propShowFocusRect
End Property

Public Property Let ShowFocusRect(ByVal New_ShowFocusRect As Boolean)
  If New_ShowFocusRect = propShowFocusRect Then Exit Property

  propShowFocusRect = New_ShowFocusRect
'
'  If Not inIDE Then
'    mustCustomDraw = isCustomDrawNecessary
'    Refresh
'  End If

  PropertyChanged "ShowFocusRect"
End Property

Public Property Get ShowHiddenDimmed() As Boolean
Attribute ShowHiddenDimmed.VB_Description = "Gibt an, ob die Icons versteckter Items abgeblendet werden."
Attribute ShowHiddenDimmed.VB_ProcData.VB_Invoke_Property = ";Icons"
  ShowHiddenDimmed = propShowHiddenDimmed
End Property

Public Property Let ShowHiddenDimmed(ByVal New_ShowHiddenDimmed As Boolean)
  If New_ShowHiddenDimmed = propShowHiddenDimmed Then Exit Property

  propShowHiddenDimmed = New_ShowHiddenDimmed

  If Not inIDE Then
    mustCustomDraw = IsCustomDrawNecessary
    Refresh
  End If

  PropertyChanged "ShowHiddenDimmed"
End Property

Public Property Get ShowIcons() As Boolean
Attribute ShowIcons.VB_Description = "Gibt an, ob Icons angezeigt werden."
Attribute ShowIcons.VB_ProcData.VB_Invoke_Property = ";Icons"
  If prophWnd Then
    ' aktuellen Wert ermitteln
    If Not inIDE And Not disableLoadFileSystem Then propShowIcons = SendMessageAsLong(prophWnd, TVM_GETIMAGELIST, TVSIL_NORMAL, 0)
  End If

  ShowIcons = propShowIcons
End Property

Public Property Let ShowIcons(ByVal New_ShowIcons As Boolean)
  Const WM_SETREDRAW = &HB
  Dim hItem As Long

  If New_ShowIcons = propShowIcons Then Exit Property

  propShowIcons = New_ShowIcons
  mustCustomDraw = IsCustomDrawNecessary

  If Not inIDE Then
    If prophWnd Then
      SendMessageAsLong prophWnd, WM_SETREDRAW, 0, 0
      If propShowIcons Then
        If Not propUseCustomIcons Then prophImageList = GetSysImageList(useLargeIcons)
        SendMessageAsLong prophWnd, TVM_SETIMAGELIST, TVSIL_NORMAL, prophImageList

        If Not propUseCustomIcons Then
          UpdateDefaultIconIndices
          InvalidateIcons 0, -1, True
        End If
      Else
        If propShowRoot Then
          hItem = GetRootOrFirstItem
          SetItemIcons hItem, 0, 0
          SetItemOverlay hItem, 0
        End If
        RemoveAllIcons hItem, True

        ' ImageList "befreien"
        SendMessageAsLong prophWnd, TVM_SETIMAGELIST, TVSIL_NORMAL, 0
        prophImageList = 0
      End If
      SendMessageAsLong prophWnd, WM_SETREDRAW, 1, 0
    End If
  End If

  PropertyChanged "ShowIcons"
End Property

Public Property Get ShownOverlays() As ShownOverlaysConstants
Attribute ShownOverlays.VB_Description = "Gibt die Typen von Overlay-Icons an, die bei Bedarf angezeigt werden."
Attribute ShownOverlays.VB_ProcData.VB_Invoke_Property = ";Icons"
Attribute ShownOverlays.VB_MemberFlags = "400"
  ShownOverlays = propShownOverlays
End Property

Public Property Let ShownOverlays(ByVal New_ShownOverlays As ShownOverlaysConstants)
  If Not isValidNumber(New_ShownOverlays, ShownOverlaysConstants.soNone, ShownOverlaysConstants.soAll) Then
    Err.Raise 380
    Exit Property
  End If
  If New_ShownOverlays = propShownOverlays Then Exit Property

  propShownOverlays = New_ShownOverlays

  If Not inIDE Then
    If (prophImageList <> 0) And Not propUseCustomIcons Then
      ' alle Overlays prüfen
      UpdateDefaultIconIndices
      InvalidateOverlays 0
    End If
  End If

  PropertyChanged "ShownOverlays"
End Property

Public Property Get ShownToolTips() As ShownToolTipsConstants
Attribute ShownToolTips.VB_Description = "Bestimmt, welche Arten von ToolTips angezeigt werden."
Attribute ShownToolTips.VB_ProcData.VB_Invoke_Property = ";Verhalten"
Attribute ShownToolTips.VB_MemberFlags = "400"
  Dim Style As Long

  If prophWnd Then
    ' aktuellen Wert ermitteln
    Style = GetWindowLongPtr(prophWnd, GWL_STYLE)
    If Style And TreeViewStyleConsts.TVS_NOTOOLTIPS Then propShownToolTips = ShownToolTipsConstants.sttNone
  End If

  ShownToolTips = propShownToolTips
End Property

Public Property Let ShownToolTips(ByVal New_ShownToolTips As ShownToolTipsConstants)
  Dim Style As Long

  If Not isValidNumber(New_ShownToolTips, ShownToolTipsConstants.sttNone, ShownToolTipsConstants.sttHiddenItems Or ShownToolTipsConstants.sttInfoTips) Then
    Err.Raise 380
    Exit Property
  End If
  If New_ShownToolTips = propShownToolTips Then Exit Property

  propShownToolTips = New_ShownToolTips

  If prophWnd Then
    ' Control aktualisieren
    Style = GetWindowLongPtr(prophWnd, GWL_STYLE)
    Style = IIf(propShownToolTips = ShownToolTipsConstants.sttNone, Style Or TreeViewStyleConsts.TVS_NOTOOLTIPS And Not TreeViewStyleConsts.TVS_INFOTIP, Style Or TreeViewStyleConsts.TVS_INFOTIP And Not TreeViewStyleConsts.TVS_NOTOOLTIPS)
    SetWindowLongPtr prophWnd, GWL_STYLE, Style
  End If

  PropertyChanged "ShownToolTips"
End Property

Public Property Get ShowRoot() As Boolean
Attribute ShowRoot.VB_Description = "Gibt an, ob der Startitem angezeigt wird."
Attribute ShowRoot.VB_ProcData.VB_Invoke_Property = ";Verhalten"
  ShowRoot = propShowRoot
End Property

Public Property Let ShowRoot(ByVal New_ShowRoot As Boolean)
  If New_ShowRoot = propShowRoot Then Exit Property

  propShowRoot = New_ShowRoot

  If Not inIDE Then
    If prophWnd Then
      ReloadItems
    End If
  End If

  PropertyChanged "ShowRoot"
End Property

Public Property Get ShowSelection() As Boolean
Attribute ShowSelection.VB_Description = "Gibt an, ob die Markierung auch angezeigt werden soll, wenn das Control den Fokus nicht hat."
Attribute ShowSelection.VB_ProcData.VB_Invoke_Property = ";Darstellung"
  Dim Style As Long

  If prophWnd Then
    ' aktuellen Wert ermitteln
    Style = GetWindowLongPtr(prophWnd, GWL_STYLE)
    propShowSelection = ((Style And TreeViewStyleConsts.TVS_SHOWSELALWAYS) = TreeViewStyleConsts.TVS_SHOWSELALWAYS)
  End If

  ShowSelection = propShowSelection
End Property

Public Property Let ShowSelection(ByVal New_ShowSelection As Boolean)
  Dim Style As Long

  If New_ShowSelection = propShowSelection Then Exit Property
  propShowSelection = New_ShowSelection

  If prophWnd Then
    ' Control aktualisieren
    Style = GetWindowLongPtr(prophWnd, GWL_STYLE)
    Style = IIf(propShowSelection, Style Or TreeViewStyleConsts.TVS_SHOWSELALWAYS, Style And Not TreeViewStyleConsts.TVS_SHOWSELALWAYS)
    SetWindowLongPtr prophWnd, GWL_STYLE, Style
    InvalidateRectAsLong prophWnd, 0, 1
  End If

  PropertyChanged "ShowSelection"
End Property

Public Property Get ShowShellContextMenus() As Boolean
Attribute ShowShellContextMenus.VB_Description = "Gibt an, ob bei einem Rechtsklick auf einen Item dessen Shell-Kontextmenü angezeigt wird."
Attribute ShowShellContextMenus.VB_ProcData.VB_Invoke_Property = ";Verhalten"
  ShowShellContextMenus = propShowShellContextMenus
End Property

Public Property Let ShowShellContextMenus(ByVal New_ShowShellContextMenus As Boolean)
  propShowShellContextMenus = New_ShowShellContextMenus

  PropertyChanged "ShowShellContextMenus"
End Property

Public Property Get SingleExpand() As SingleExpandConstants
Attribute SingleExpand.VB_Description = "Gibt an, ob immer nur 1 Item erweitert sein kann."
Attribute SingleExpand.VB_ProcData.VB_Invoke_Property = ";Verhalten"
  Dim Style As Long

  If prophWnd Then
    ' aktuellen Wert ermitteln
    Style = GetWindowLongPtr(prophWnd, GWL_STYLE)
    If (Style And TreeViewStyleConsts.TVS_SINGLEEXPAND) = 0 Then propSingleExpand = SingleExpandConstants.seNone
  End If

  SingleExpand = propSingleExpand
End Property

Public Property Let SingleExpand(ByVal New_SingleExpand As SingleExpandConstants)
  Dim Style As Long

  If Not isValidNumber(New_SingleExpand, SingleExpandConstants.seNone, SingleExpandConstants.seWinXPStyle) Then
    Err.Raise 380
    Exit Property
  End If
  If New_SingleExpand = propSingleExpand Then Exit Property

  propSingleExpand = New_SingleExpand
  If propSingleExpand <> SingleExpandConstants.seNormal Then
    ' FavoritesStyle deaktivieren
    FavoritesStyle = False
  End If

  If prophWnd Then
    ' Control aktualisieren
    Style = GetWindowLongPtr(prophWnd, GWL_STYLE)
    Style = IIf(propSingleExpand <> SingleExpandConstants.seNone, Style Or TreeViewStyleConsts.TVS_SINGLEEXPAND, Style And Not TreeViewStyleConsts.TVS_SINGLEEXPAND)
    SetWindowLongPtr prophWnd, GWL_STYLE, Style
  End If

  PropertyChanged "SingleExpand"
End Property

Public Property Get SortingCriteria() As Long
Attribute SortingCriteria.VB_Description = "Gibt die ID der Eigenschaft an, nach der die Items sortiert werden."
Attribute SortingCriteria.VB_ProcData.VB_Invoke_Property = ";Sortierung"
  SortingCriteria = propSortingCriteria
End Property

Public Property Let SortingCriteria(ByVal New_SortingCriteria As Long)
  If New_SortingCriteria < 0 Then
    Err.Raise 380
    Exit Property
  End If
  If New_SortingCriteria = propSortingCriteria Then Exit Property

  propSortingCriteria = New_SortingCriteria

  If Not inIDE Then
    If prophWnd Then
      SortAllSubItems IIf(propShowRoot, GetRootOrFirstItem, 0)
    End If
  End If

  PropertyChanged "SortingCriteria"
End Property

Public Property Get SortOrder() As SortOrderConstants
Attribute SortOrder.VB_Description = "Gibt an, ob die Items in aufsteigender oder absteigender Reihenfolge sortiert werden."
Attribute SortOrder.VB_ProcData.VB_Invoke_Property = ";Sortierung"
  SortOrder = propSortOrder
End Property

Public Property Let SortOrder(ByVal New_SortOrder As SortOrderConstants)
  If Not isValidNumber(New_SortOrder, SortOrderConstants.sorAscending, SortOrderConstants.sorDescending) Then
    Err.Raise 380
    Exit Property
  End If
  If New_SortOrder = propSortOrder Then Exit Property

  propSortOrder = New_SortOrder

  If Not inIDE Then
    If prophWnd Then
      SortAllSubItems IIf(propShowRoot, GetRootOrFirstItem, 0)
    End If
  End If

  PropertyChanged "SortOrder"
End Property

Public Property Get TreeViewStyle() As TreeViewStyleConstants
Attribute TreeViewStyle.VB_Description = "Gibt an, ob Buttons und Verbindungslinien angezeigt werden."
Attribute TreeViewStyle.VB_ProcData.VB_Invoke_Property = ";Darstellung"
Attribute TreeViewStyle.VB_MemberFlags = "400"
  Dim Style As Long

  If prophWnd Then
    ' aktuellen Wert ermitteln
    Style = GetWindowLongPtr(prophWnd, GWL_STYLE)
    propTreeViewStyle = 0
    If Style And TreeViewStyleConsts.TVS_HASBUTTONS Then propTreeViewStyle = propTreeViewStyle Or TreeViewStyleConstants.tvsButtons
    If Style And TreeViewStyleConsts.TVS_HASLINES Then propTreeViewStyle = propTreeViewStyle Or TreeViewStyleConstants.tvsLines
  End If

  TreeViewStyle = propTreeViewStyle
End Property

Public Property Let TreeViewStyle(ByVal New_TreeViewStyle As TreeViewStyleConstants)
  Dim Style As Long

  If New_TreeViewStyle = propTreeViewStyle Then Exit Property

  If New_TreeViewStyle And TreeViewStyleConstants.tvsLines Then
    ' FullRowSelect deaktivieren
    FullRowSelect = False
  End If

  If Not isValidNumber(New_TreeViewStyle, 0, TreeViewStyleConstants.tvsButtons Or TreeViewStyleConstants.tvsLines) Then
    Err.Raise 380
    Exit Property
  End If
  propTreeViewStyle = New_TreeViewStyle

  If prophWnd Then
    ' Control aktualisieren
    Style = GetWindowLongPtr(prophWnd, GWL_STYLE)
    Style = IIf(propTreeViewStyle And TreeViewStyleConstants.tvsButtons, Style Or TreeViewStyleConsts.TVS_HASBUTTONS, Style And Not TreeViewStyleConsts.TVS_HASBUTTONS)
    Style = IIf(propTreeViewStyle And TreeViewStyleConstants.tvsLines, Style Or TreeViewStyleConsts.TVS_HASLINES, Style And Not TreeViewStyleConsts.TVS_HASLINES)
    SetWindowLongPtr prophWnd, GWL_STYLE, Style
  End If

  PropertyChanged "TreeViewStyle"
End Property

Public Property Get UseCustomIcons() As Boolean
Attribute UseCustomIcons.VB_Description = "Gibt an, ob eine benutzerdefinierte ImageList statt der System-ImageList verwendet wird."
Attribute UseCustomIcons.VB_ProcData.VB_Invoke_Property = ";Darstellung"
  UseCustomIcons = propUseCustomIcons
End Property

Public Property Let UseCustomIcons(ByVal New_UseCustomIcons As Boolean)
  Const WM_SETREDRAW = &HB

  If New_UseCustomIcons = propUseCustomIcons Then Exit Property

  propUseCustomIcons = New_UseCustomIcons
  If propShowIcons And (prophWnd <> 0) And Not inIDE Then
    If propUseCustomIcons Then
      SendMessageAsLong prophWnd, TVM_SETIMAGELIST, TVSIL_NORMAL, 0
      prophImageList = 0
    Else
      SendMessageAsLong prophWnd, WM_SETREDRAW, 0, 0

      prophImageList = GetSysImageList(useLargeIcons)
      SendMessageAsLong prophWnd, TVM_SETIMAGELIST, TVSIL_NORMAL, prophImageList
      If propCheckBoxes Then
        SendMessageAsLong prophWnd, TVM_SETIMAGELIST, TVSIL_STATE, IIf(useLargeIcons, prophLargeStateImageList, prophSmallStateImageList)
      End If

      UpdateDefaultIconIndices
      InvalidateIcons 0, -1, True

      SendMessageAsLong prophWnd, WM_SETREDRAW, 1, 0
    End If
  End If

  PropertyChanged "UseCustomIcons"
End Property

Public Property Get UseFileFilters() As Boolean
Attribute UseFileFilters.VB_Description = "Gibt an, ob die Eigenschaft <FileFilters> berücksichtigt wird."
Attribute UseFileFilters.VB_ProcData.VB_Invoke_Property = ";Filter"
  UseFileFilters = propUseFileFilters
End Property

Public Property Let UseFileFilters(ByVal New_UseFileFilters As Boolean)
  If New_UseFileFilters = propUseFileFilters Then Exit Property

  propUseFileFilters = New_UseFileFilters

  If Not inIDE Then
    If prophWnd Then
      ReloadItems
    End If
  End If

  PropertyChanged "UseFileFilters"
End Property

Public Property Get UseFolderFilters() As Boolean
Attribute UseFolderFilters.VB_Description = "Gibt an, ob die Eigenschaft <FolderFilters> berücksichtigt wird."
Attribute UseFolderFilters.VB_ProcData.VB_Invoke_Property = ";Filter"
  UseFolderFilters = propUseFolderFilters
End Property

Public Property Let UseFolderFilters(ByVal New_UseFolderFilters As Boolean)
  If New_UseFolderFilters = propUseFolderFilters Then Exit Property

  propUseFolderFilters = New_UseFolderFilters

  If Not inIDE Then
    If prophWnd Then
      ReloadItems
    End If
  End If

  PropertyChanged "UseFolderFilters"
End Property

Public Property Get UseSelectedIcons() As Boolean
Attribute UseSelectedIcons.VB_Description = "Gibt an, ob ein spezielles Icon genutzt wird, wenn ein Item markiert ist."
Attribute UseSelectedIcons.VB_ProcData.VB_Invoke_Property = ";Icons"
  UseSelectedIcons = propUseSelectedIcons
End Property

Public Property Let UseSelectedIcons(ByVal New_UseSelectedIcons As Boolean)
  If New_UseSelectedIcons = propUseSelectedIcons Then Exit Property

  propUseSelectedIcons = New_UseSelectedIcons
  If Not inIDE Then
    If (prophImageList <> 0) And Not propUseCustomIcons Then
      ' alle Icons prüfen
      UpdateDefaultIconIndices
      InvalidateIcons 0, -1
    End If
  End If

  PropertyChanged "UseSelectedIcons"
End Property

Public Property Get UseStandardIcons() As Boolean
Attribute UseStandardIcons.VB_Description = "Gibt an, ob generische Icons genutzt werden statt für jeden Item das exakte Icon zu ermitteln."
  UseStandardIcons = propUseStandardIcons
End Property

Public Property Let UseStandardIcons(ByVal New_UseStandardIcons As Boolean)
  If New_UseStandardIcons = propUseStandardIcons Then Exit Property

  propUseStandardIcons = New_UseStandardIcons
  If Not inIDE Then
    If (prophImageList <> 0) And Not propUseCustomIcons Then
      ' alle Icons prüfen
      UpdateDefaultIconIndices
      InvalidateIcons 0, -1
    End If
  End If

  PropertyChanged "UseStandardIcons"
End Property

Public Property Get UseStandardIconsForSlowItems() As Boolean
  UseStandardIconsForSlowItems = propUseStandardIconsForSlowItems
End Property

Public Property Let UseStandardIconsForSlowItems(ByVal New_UseStandardIconsForSlowItems As Boolean)
  If New_UseStandardIconsForSlowItems = propUseStandardIconsForSlowItems Then Exit Property

  propUseStandardIconsForSlowItems = New_UseStandardIconsForSlowItems
  If Not inIDE Then
    If (prophImageList <> 0) And Not propUseCustomIcons Then
      ' alle Icons prüfen
      UpdateDefaultIconIndices
      InvalidateIcons 0, -1
    End If
  End If

  PropertyChanged "UseStandardIconsForSlowItems"
End Property

Public Property Get Version() As String
Attribute Version.VB_Description = "Gibt die Version des Controls zurück."
Attribute Version.VB_ProcData.VB_Invoke_Property = ";Daten"
  With App
    Version = .Major & "." & .Minor & "." & .Revision & " (Build " & BuildNumber & ")" ' - RC3"
  End With
End Property

Public Property Let Version(ByVal New_Version As String)
  Err.Raise IIf(inIDE, 387, 382)
End Property


' geheime Props - Infos zum Control

Public Property Get AppID() As Long
Attribute AppID.VB_ProcData.VB_Invoke_Property = ";Daten"
Attribute AppID.VB_MemberFlags = "40"
  AppID = 2
End Property

Public Property Get AppName() As String
Attribute AppName.VB_ProcData.VB_Invoke_Property = ";Daten"
Attribute AppName.VB_MemberFlags = "40"
  AppName = "ExplorerTreeView"
End Property

Public Property Get AppShortName() As String
Attribute AppShortName.VB_ProcData.VB_Invoke_Property = ";Daten"
Attribute AppShortName.VB_MemberFlags = "40"
  AppShortName = "ExTvw"
End Property

Public Property Get Build() As Integer
Attribute Build.VB_ProcData.VB_Invoke_Property = ";Daten"
Attribute Build.VB_MemberFlags = "440"
  Build = BuildNumber
End Property

Public Property Get Company() As String
Attribute Company.VB_ProcData.VB_Invoke_Property = ";Daten"
Attribute Company.VB_MemberFlags = "40"
  Company = App.CompanyName
End Property

Public Property Get DesignedWith() As String
Attribute DesignedWith.VB_ProcData.VB_Invoke_Property = ";Daten"
Attribute DesignedWith.VB_MemberFlags = "40"
  DesignedWith = "Microsoft Visual Basic 6.0 Professional Edition (SP5)"
End Property

' Zeit von 1. Zeile bis 1. Release
Public Property Get DesignTime() As String
Attribute DesignTime.VB_ProcData.VB_Invoke_Property = ";Daten"
Attribute DesignTime.VB_MemberFlags = "40"
  DesignTime = "??.?.2000 - 15.07.2002"
End Property

Public Property Get IsRelease() As Boolean
Attribute IsRelease.VB_ProcData.VB_Invoke_Property = ";Daten"
Attribute IsRelease.VB_MemberFlags = "40"
  #If Debuging Then
    IsRelease = False
  #Else
    IsRelease = True
  #End If
End Property

Public Property Get Programmer() As String
Attribute Programmer.VB_ProcData.VB_Invoke_Property = ";Daten"
Attribute Programmer.VB_MemberFlags = "40"
  Programmer = "Timo ""TimoSoft"" Kunze"
End Property

Public Property Get TestedBy() As String
Attribute TestedBy.VB_ProcData.VB_Invoke_Property = ";Daten"
Attribute TestedBy.VB_MemberFlags = "40"
  TestedBy = "Timo ""TimoSoft"" Kunze|Christian Lütgens"
End Property


' öffentliche Methoden

Public Sub About()
Attribute About.VB_Description = "Zeigt Informationen über das Control an."
Attribute About.VB_UserMemId = -552
  Dim frm As frmAbout

  Set frm = New frmAbout
  frm.showIt UserControl.Name, resIcon
  Set frm = Nothing
End Sub

Public Function AddFTPPath(ByVal FTPPath As String) As Long
Attribute AddFTPPath.VB_Description = "Fügt die FTP-URI <FTPPath> als Sub-Item des Internet Explorers ein und gibt den Handle des letzten eingefügten Items zurück."
  Dim DispName As String
  Dim hParentItem As Long
  Dim hSubItem As Long
  Dim IItem As IVBShellFolder
  Dim IShIconOverlay As IVBShellIconOverlay
  Dim itemAttr As SFGAOConstants
  Dim overlayIndex As Long
  Dim pIDLSubItem_ToDesktop As Long
  Dim pIDLSubItem_ToParent As Long
  Dim ret As Long
  Dim Segment As String
  Dim ServerPart As String
  Dim TVWItemInfo As ITEMINFO_COMMON
  Dim TVWItemInfo_Parent As ITEMINFO_COMMON
  Dim useStandardIcon As Boolean

  ret = -1
  If Not IsFTPURL(FTPPath) Then
    AddFTPPath = -1
    Exit Function
  End If

  ServerPart = GetFirstFolders(FTPPath, 1, "/")
  FTPPath = Mid$(FTPPath, Len(ServerPart & "/") + 1)

  hParentItem = GetInternetExplorerItemHandle(True)
  If Not GetItemInfo_Common(hParentItem, TVWItemInfo_Parent) Then
    AddFTPPath = -1
    Exit Function
  End If
  #If Debuging Then
    debugger.AddLogEntry "AddFTPPath: Calling GetISHFolderInterfaceFQ()", LogEntryTypeConstants.letInfo
    debugger.AddLogEntry "   0x" & Hex(TVWItemInfo_Parent.pIDLToDesktop), LogEntryTypeConstants.letOther
    Set IItem = GetISHFolderInterfaceFQ(debugger, TVWItemInfo_Parent.pIDLToDesktop)
  #Else
    Set IItem = GetISHFolderInterfaceFQ(TVWItemInfo_Parent.pIDLToDesktop)
  #End If

  If Not (IItem Is Nothing) Then
    pIDLSubItem_ToDesktop = PathTopIDL(ServerPart)
    ' testen, ob der FTP-Pfad schon existiert
    ret = InternalpIDLToItemHandle(pIDLSubItem_ToDesktop, hParentItem)
    If ret <> -1 Then
      #If Debuging Then
        FreeItemIDList debugger, "AddFTPPath", pIDLSubItem_ToDesktop
      #Else
        FreeItemIDList pIDLSubItem_ToDesktop
      #End If
    Else
      pIDLSubItem_ToParent = ILFindLastID(pIDLSubItem_ToDesktop)
      ' Item anzeigen
      TVWItemInfo.pIDLToDesktop = pIDLSubItem_ToDesktop

      ' Typ des Items ermitteln
      #If Debuging Then
        itemAttr = GetAttributes(debugger, IItem, pIDLSubItem_ToParent, SFGAOConstants.SFGAO_FILESYSTEM Or SFGAOConstants.SFGAO_FOLDER Or SFGAOConstants.SFGAO_SHARE Or SFGAOConstants.SFGAO_LINK Or SFGAOConstants.SFGAO_GHOSTED Or SFGAOConstants.SFGAO_NEWCONTENT)
      #Else
        itemAttr = GetAttributes(IItem, pIDLSubItem_ToParent, SFGAOConstants.SFGAO_FILESYSTEM Or SFGAOConstants.SFGAO_FOLDER Or SFGAOConstants.SFGAO_SHARE Or SFGAOConstants.SFGAO_LINK Or SFGAOConstants.SFGAO_GHOSTED Or SFGAOConstants.SFGAO_NEWCONTENT)
      #End If
      If itemAttr And SFGAOConstants.SFGAO_FILESYSTEM Then
        If itemAttr And SFGAOConstants.SFGAO_FOLDER Then
          ' ist es wirklich ein Ordner?
          #If Debuging Then
            If FileExists_pIDL(debugger, IItem, pIDLSubItem_ToParent) Then
          #Else
            If FileExists_pIDL(IItem, pIDLSubItem_ToParent) Then
          #End If
            TVWItemInfo.itemType = CByte(ItemTypeConstants.itFSFile)
          Else
            TVWItemInfo.itemType = CByte(ItemTypeConstants.itFSFolder)
          End If
        Else
          TVWItemInfo.itemType = CByte(ItemTypeConstants.itFSFile)
        End If
      Else
        If itemAttr And SFGAOConstants.SFGAO_FOLDER Then
          TVWItemInfo.itemType = CByte(ItemTypeConstants.itNonFSFolder)
        Else
          TVWItemInfo.itemType = CByte(ItemTypeConstants.itNonFSFile)
        End If
      End If

      If (prophImageList <> 0) And Not propUseCustomIcons Then
        ' Overlays laden
        overlayIndex = 0
        If propShownOverlays <> ShownOverlaysConstants.soNone Then
          On Error Resume Next
          IItem.QueryInterface IID_IShellIconOverlay, IShIconOverlay
          If IShIconOverlay Is Nothing Then
            If itemAttr And SFGAOConstants.SFGAO_SHARE Then overlayIndex = 1
            If itemAttr And SFGAOConstants.SFGAO_LINK Then overlayIndex = 2
          Else
            If IShIconOverlay.GetOverlayIndex(pIDLSubItem_ToParent, overlayIndex) = S_OK Then
              ' jetzt haben wir den Overlay-Index für den Item, aber soll dieser Overlay
              ' überhaupt angezeigt werden?
              If ShouldShowOverlay(propShownOverlays, overlayIndex) Then
                ' ja!
              Else
                ' kein Overlay
                overlayIndex = 0
              End If
            Else
              ' kein Overlay
              overlayIndex = 0
            End If
            Set IShIconOverlay = Nothing
          End If
        End If
      End If

      #If Debuging Then
        DispName = pIDLToDisplayName(debugger, IItem, pIDLSubItem_ToParent, SHGDNConstants.SHGDN_INFOLDER)
        If (prophImageList <> 0) And Not propUseCustomIcons Then
          #If NoCallbacks Then
            useStandardIcon = propUseStandardIcons
            If propUseStandardIconsForSlowItems And Not useStandardIcon Then
              useStandardIcon = IsSlowItem(debugger, IItem, pIDLSubItem_ToParent, TVWItemInfo.pIDLToDesktop, True)
            End If
          #End If
        End If
      #Else
        DispName = pIDLToDisplayName(IItem, pIDLSubItem_ToParent, SHGDNConstants.SHGDN_INFOLDER)
        If (prophImageList <> 0) And Not propUseCustomIcons Then
          #If NoCallbacks Then
            useStandardIcon = propUseStandardIcons
            If propUseStandardIconsForSlowItems And Not useStandardIcon Then
              useStandardIcon = IsSlowItem(IItem, pIDLSubItem_ToParent, TVWItemInfo.pIDLToDesktop, True)
            End If
          #End If
        End If
      #End If

      ' Item hinzufügen
      ret = InsertTVWItem_Common(DispName, hParentItem, TVWItemInfo, overlayIndex, itemAttr And SFGAOConstants.SFGAO_GHOSTED, itemAttr And SFGAOConstants.SFGAO_NEWCONTENT, useStandardIcon)
    End If

    SortItems hParentItem
    ItemExpand hParentItem
  End If
  Set IItem = Nothing

  ' es folgt Teil 2 - das expandieren bis auch der Sub-Pfad komplett eingefügt ist
  hSubItem = ret
  If FTPPath <> "" Then
    ItemExpand hSubItem
    Do While FTPPath <> ""
      Segment = GetFirstFolders(FTPPath, 1, "/")
      hSubItem = ItemGetFirstSubItem(hSubItem)
      DispName = ""
      Do While hSubItem <> -1
        DispName = ItemHandleToDisplayName(hSubItem)
        If LCase$(DispName) = LCase$(Segment) Then Exit Do
        hSubItem = ItemGetNextItem(hSubItem)
      Loop
      If hSubItem = -1 Then
        Exit Do
      Else
        ItemExpand hSubItem
        ret = hSubItem
      End If

      FTPPath = RemoveFirstFolders(FTPPath, 1, "/")
    Loop
  End If

  AddFTPPath = ret
End Function

Public Function EndRenaming(ByVal Cancel As Boolean) As Boolean
Attribute EndRenaming.VB_Description = "Beendet das Umbenennen eines Items. Mit <Cancel> kann bestimmt werden, ob die Umbenennung vollzogen oder abgebrochen werden soll."
  EndRenaming = SendMessageAsLong(prophWnd, TVM_ENDEDITLABELNOW, Abs(Cancel), 0)
End Function

Public Function FileContainsValidSettings(ByVal File As String, Optional ByRef ErrorType As ErrorTypeConstants, Optional ByRef AdditionalData As Variant = "") As Boolean
Attribute FileContainsValidSettings.VB_Description = "Gibt zurück, ob <File> gültige Daten für einen Aufruf von ""LoadSettingsFromFile()"" enthält. Für <ErrorType> = etUnknown enthält <AdditionalData> den Fehlercode; für etTooNewFileFormatRead die benötigte Version von ExplorerTreeView."
  Dim BytesRead As Long
  Dim canRead As Boolean
  Dim hFile As Long
  Dim intBuffer As Integer
  Dim isVersion10 As Boolean
  Dim isVersion20 As Boolean
  Dim lngBuffer As Long
  Dim pos As Integer
  Dim ret As Boolean
  Dim ret2 As Long
  Dim Start As Long
  Dim strBuffer As String
  Dim Version_Build As Integer
  Dim Version_Major As Byte
  Dim Version_Minor As Byte
  Dim Version_Revision As Integer

  On Error GoTo Fehler

  AdditionalData = Empty
  ErrorType = ErrorTypeConstants.etNone

  If FileExists(File) Then
    ' Datei öffnen
    hFile = CreateFileAsLong(File, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If hFile = INVALID_HANDLE_VALUE Then GoTo CreateFileError

    ' Signatur einlesen...
    strBuffer = Space$(7)
    ret2 = ReadFileAsLong(hFile, ByVal strBuffer, Len(strBuffer), BytesRead, 0)
    If ret2 = 0 Then GoTo ReadFileError
    strBuffer = Left$(strBuffer, lstrlenA(strBuffer))
    '...und prüfen
    If strBuffer = FileTypeSignature_SettingsFile Then
      ' Formatversion einlesen...
      strBuffer = Space$(4)
      ret2 = ReadFileAsLong(hFile, ByVal strBuffer, Len(strBuffer), BytesRead, 0)
      If ret2 = 0 Then GoTo ReadFileError
      strBuffer = Left$(strBuffer, lstrlenA(strBuffer))
      '...und prüfen
      pos = InStr(strBuffer, ".")
      If pos = 0 Then pos = Len(strBuffer)
      isVersion20 = (CInt(Left$(strBuffer, pos - 1)) >= 2)
      isVersion10 = (CInt(Left$(strBuffer, pos - 1)) = 1)

      If isVersion10 Or isVersion20 Then
        ' Signatur und Version des Controls, welches die Datei erzeugt hat, überspringen
        If isVersion20 Then
          SetFilePointer hFile, 11, 0, FILE_CURRENT
        Else
          SetFilePointer hFile, 10, 0, FILE_CURRENT
        End If
        If isVersion20 Then
          ' Signatur und Version des Controls, welches die Datei zuletzt geändert hat, überspringen
          SetFilePointer hFile, 11, 0, FILE_CURRENT

          ' Version dieses Controls einlesen, die benötigt wird um dieses Format zu lesen...
          ret2 = ReadFileAsLong(hFile, ByVal VarPtr(lngBuffer), LenB(lngBuffer), BytesRead, 0)
          If ret2 = 0 Then GoTo ReadFileError
          intBuffer = HiWord(lngBuffer)
          Version_Major = HiByte(intBuffer)
          Version_Minor = LoByte(intBuffer)
          Version_Revision = LoWord(lngBuffer)
          ret2 = ReadFileAsLong(hFile, ByVal VarPtr(Version_Build), LenB(Version_Build), BytesRead, 0)
          If ret2 = 0 Then GoTo ReadFileError
          AdditionalData = Version_Major & "." & Version_Minor & "." & Version_Revision & " (Build " & Version_Build & ")"
          ' ...und prüfen
          If App.Major > Version_Major Then
            canRead = True
          ElseIf App.Major = Version_Major Then
            If App.Minor > Version_Minor Then
              canRead = True
            ElseIf App.Minor = Version_Minor Then
              If App.Revision > Version_Revision Then
                canRead = True
              ElseIf App.Revision = Version_Revision Then
                canRead = (BuildNumber >= Version_Build)
              End If
            End If
          End If

          If canRead Then
            ' 86 Bytes überspringen (12 Bytes Versionsinfos für ExplorerListView und ExplorerComboBox +
            ' 12 Bytes reserviert + 18 Bytes Versionsinfos für das Schreiben + 12 Bytes reserviert + 32
            ' Bytes reserviert)
            SetFilePointer hFile, 86, 0, FILE_CURRENT
          End If
        Else
          canRead = True
          ' 32 reservierte Bytes überspringen
          SetFilePointer hFile, 32, 0, FILE_CURRENT
        End If

        If canRead Then
          ' Adresse des Startbytes der Eigenschaften überspringen
          SetFilePointer hFile, 4, 0, FILE_CURRENT

          ' Länge der Eigenschaften lesen
          ret2 = ReadFileAsLong(hFile, ByVal VarPtr(lngBuffer), LenB(lngBuffer), BytesRead, 0)
          If ret2 = 0 Then GoTo ReadFileError

          If lngBuffer > 0 Then
            ret = True
            ErrorType = ErrorTypeConstants.etNone
          Else
            ErrorType = ErrorTypeConstants.etNoSettingsFound
          End If
        Else
          ErrorType = ErrorTypeConstants.etTooNewFileFormatToRead
        End If
      Else
        ErrorType = ErrorTypeConstants.etInvalidFormat
      End If
    Else
      ErrorType = ErrorTypeConstants.etInvalidFormat
    End If

    ' Datei schließen
    CloseHandle hFile
  Else
    ErrorType = ErrorTypeConstants.etFileNotFound
  End If

  FileContainsValidSettings = ret
  Exit Function

Fehler:
  ' unbekannter Fehler
  CloseHandle hFile
  ErrorType = ErrorTypeConstants.etUnknown
  AdditionalData = Err.Number
  Exit Function

CreateFileError:
  ' CreateFile ist fehlgeschlagen
  CloseHandle hFile
  ErrorType = ErrorTypeConstants.etCanNotOpenFile
  Exit Function

ReadFileError:
  ' ReadFile ist fehlgeschlagen
  CloseHandle hFile
  ErrorType = ErrorTypeConstants.etCanNotReadFile
End Function

Public Sub FreepIDLMem(ByRef pIDL As Long)
Attribute FreepIDLMem.VB_Description = "Gibt den Speicher frei, der von <pIDL> belegt wird."
  #If Debuging Then
    FreeItemIDList debugger, "FreepIDLMem", pIDL
  #Else
    FreeItemIDList pIDL
  #End If
End Sub

Public Function GetFirstItem() As Long
Attribute GetFirstItem.VB_Description = "Gibt den 1. Item auf der 1. Ebene zurück."
  GetFirstItem = GetRootOrFirstItem
End Function

Public Function GetItemCount(Optional ByVal StateIconMask As Long = -1, Optional ByVal MustMatch As Boolean = True) As Long
Attribute GetItemCount.VB_Description = "Gibt die Anzahl aller Items im Control zurück."
  Dim hItem As Long
  Dim ret As Long

  If StateIconMask = -1 Then
    ret = SendMessageAsLong(prophWnd, TVM_GETCOUNT, 0, 0)
  Else
    If propShowRoot Then
      hItem = GetRootOrFirstItem
      If ItemStateIconIndex(hItem) = StateIconMask Then
        If MustMatch Then ret = ret + 1
      Else
        If Not MustMatch Then ret = ret + 1
      End If
    End If
    ret = ret + CountAllSubItems(hItem, StateIconMask, MustMatch)
  End If

  GetItemCount = ret
End Function

Public Function GetLastItem() As Long
Attribute GetLastItem.VB_Description = "Gibt den letzten (ggf. durch scrollen) sichtbaren Item im Control zurück."
  Dim ret As Long

  ret = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_LASTVISIBLE, 0)
  If ret = 0 Then ret = -1

  GetLastItem = ret
End Function

Public Function GetVisibleItemCount() As Long
Attribute GetVisibleItemCount.VB_Description = "Gibt die maximale Anzahl der (vollständig) sichtbaren Items zurück."
  GetVisibleItemCount = SendMessageAsLong(prophWnd, TVM_GETVISIBLECOUNT, 0, 0)
End Function

Public Function HitTest(ByVal x As Long, ByVal y As Long, ByRef TestFlags As HitTestConstants) As Long
Attribute HitTest.VB_Description = "Gibt den Item unter dem Punkt [<x>;<y>] zurück. <TestFlags> enthält nähere Infos zu diesem Punkt."
  Dim pt As POINT
  Dim ret As Long

  With pt
    .x = x
    .y = y
  End With
  ret = ItemHitTest(pt, TestFlags)

  If ret = 0 Then ret = -1
  HitTest = ret
End Function

' gibt zurück, ob derzeit etwas mit Drag'n'Drop läuft
Public Function IsInDragDropOperation() As Boolean
Attribute IsInDragDropOperation.VB_Description = "Gibt an, ob das Control in einen Drag'n'Drop-Vorgang verwickelt ist."
  IsInDragDropOperation = ((hDraggedItem <> 0) Or (prophDropTargetItem <> 0))
End Function

' reduziert <hItem>
Public Sub ItemCollapse(ByVal hItem As Long, Optional ByVal DeleteSubItems As Boolean = False)
Attribute ItemCollapse.VB_Description = "Reduziert den Item <hItem>."
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub
  SendMessageAsLong prophWnd, TVM_EXPAND, TVE_COLLAPSE Or IIf(DeleteSubItems, TVE_COLLAPSERESET, 0), hItem
End Sub

Public Function ItemCreateNewFolder(ByVal hParentItem As Long) As Long
Attribute ItemCreateNewFolder.VB_Description = "Erstellt im Item <hParentItem> einen neuen Ordner."
  Dim Flags As CMFConstants
  Dim IContMenu As IVBContextMenu
  Dim IItem As IVBShellFolder
  Dim InvokeData As CMINVOKECOMMANDINFOEX
  Dim ret As Long
  Dim TVWItemInfo As ITEMINFO_COMMON

  If Not ItemExists(hParentItem, AllowNullConstants.anNever) Then Exit Function

  ' Daten über <hParentItem> sammeln
  If Not GetItemInfo_Common(hParentItem, TVWItemInfo, True) Then Exit Function

  #If Debuging Then
    debugger.AddLogEntry "ItemCreateNewFolder: Calling GetISHFolderInterfaceFQ()", LogEntryTypeConstants.letInfo
    debugger.AddLogEntry "   0x" & Hex(TVWItemInfo.pIDLToDesktop), LogEntryTypeConstants.letOther
    Set IItem = GetISHFolderInterfaceFQ(debugger, TVWItemInfo.pIDLToDesktop)
  #Else
    Set IItem = GetISHFolderInterfaceFQ(TVWItemInfo.pIDLToDesktop)
  #End If
  If Not (IItem Is Nothing) Then
    ' Interface "IContextMenu" ermitteln
    IItem.CreateViewObject prophWndShellUIParentWindow, IID_IContextMenu, IContMenu
    If Not (IContMenu Is Nothing) Then
      ' Menü initiieren
      Flags = CMFConstants.CMF_NORMAL Or CMFConstants.CMF_EXPLORE
      If propRenaming = RenamingConstants.rAutomatic Then Flags = Flags Or CMFConstants.CMF_CANRENAME
      IContMenu.QueryContextMenu 0, 0, 1, &H7FFF, Flags

      With InvokeData
        .StructSize = LenB(InvokeData)
        .hWnd = prophWnd
        .Verb = StrPtr(StrConv("NewFolder", VbStrConv.vbFromUnicode))
        .Showflags = SW_SHOWNORMAL
      End With
      createdFolder = True
      hCreatedItem = -1
      If IContMenu.InvokeCommand(InvokeData) = NOERROR Then
        ret = hCreatedItem
        If propRenaming = RenamingConstants.rAutomatic Then
          SelChangeInternal = True
          Me.SelectedItem = hCreatedItem
          SelChangeInternal = False
          Me.ItemEnsureVisible hCreatedItem
          Me.ItemStartRenaming hCreatedItem
        End If
      Else
        ret = -1
      End If
      hCreatedItem = -1
      createdFolder = False

      Set IContMenu = Nothing
    End If
  End If
  Set IItem = Nothing

  If ret = 0 Then ret = -1
  ItemCreateNewFolder = ret
End Function

Public Function ItemDoesExist(ByVal hItem As Long) As Boolean
Attribute ItemDoesExist.VB_Description = "Gibt zurück, ob der Item <hItem> existiert."
  ItemDoesExist = ItemExists(hItem, AllowNullConstants.anNever, , False)
End Function

Public Sub ItemEnsureVisible(ByVal hItem As Long)
Attribute ItemEnsureVisible.VB_Description = "Stellt sicher, dass der Item <hItem> sichtbar ist. Dazu werden ggf. Items erweitert und das Control gescrollt."
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub

  SendMessageAsLong prophWnd, TVM_ENSUREVISIBLE, 0, hItem
End Sub

' führt für <hItem> den Standard-Befehl aus
' gibt bei Erfolg True zurück
Public Function ItemExecuteDefaultCommand(ByVal hItem As Long, Optional ByVal WindowMode As WindowModeConstants = WindowModeConstants.wmShowNormal) As Boolean
Attribute ItemExecuteDefaultCommand.VB_Description = "Führt den Standardbefehl für den Item <hItem> aus. Der Standardbefehl ist im Kontextmenü meist hervorgehoben."
  Dim CmdID As Long
  Dim Flags As CMFConstants
  Dim hMenu As Long
  Dim IContMenu As IVBContextMenu
  Dim InvokeData As CMINVOKECOMMANDINFOEX
  Dim IParent As IVBShellFolder
  Dim pIDLToParent As Long
  Dim TVWItemInfo As ITEMINFO_COMMON

  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Function

  ' Daten über <hItem> sammeln
  If Not GetItemInfo_Common(hItem, TVWItemInfo, True) Then Exit Function

  #If Debuging Then
    SplitFullyQualifiedPIDL debugger, TVWItemInfo.pIDLToDesktop, IParent, pIDLToParent
  #Else
    SplitFullyQualifiedPIDL TVWItemInfo.pIDLToDesktop, IParent, pIDLToParent
  #End If
  ' Interface "IContextMenu" ermitteln
  IParent.GetUIObjectOf prophWndShellUIParentWindow, 1, pIDLToParent, IID_IContextMenu, 0, IContMenu
  If Not (IContMenu Is Nothing) Then
    ' Menü erzeugen und füllen
    Flags = CMFConstants.CMF_DEFAULTONLY
    If propRenaming = RenamingConstants.rAutomatic Then
      Flags = Flags Or CMFConstants.CMF_CANRENAME
    End If
    hMenu = CreatePopupMenu
    IContMenu.QueryContextMenu hMenu, 0, 1, &H7FFF, Flags

    CmdID = GetMenuDefaultItem(hMenu, 0, GMDI_GOINTOPOPUPS)
    If CmdID <> -1 Then
      With InvokeData
        .StructSize = LenB(InvokeData)
        .hWnd = prophWnd
        .Verb = MakeDWord(CmdID - 1, 0)
        .Showflags = WindowMode
      End With
      ItemExecuteDefaultCommand = (IContMenu.InvokeCommand(InvokeData) = NOERROR)
    End If
  End If

  Set IParent = Nothing
  DestroyMenu hMenu
  Set IContMenu = Nothing
End Function

' erweitert <hItem>
Public Sub ItemExpand(ByVal hItem As Long)
Attribute ItemExpand.VB_Description = "Erweitert den Item <hItem>."
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub

  SendMessageAsLong prophWnd, TVM_EXPAND, TVE_EXPAND, hItem
End Sub

' gibt den 1. SubItem von <hParentItem> zurück
Public Function ItemGetFirstSubItem(ByVal hParentItem As Long, Optional ByVal StateIconMask As Long = -1, Optional ByVal MustMatch As Boolean = True) As Long
Attribute ItemGetFirstSubItem.VB_Description = "Gibt den ersten Subitem des Items <hItem> zurück."
  Dim ret As Long

  If Not ItemExists(hParentItem, AllowNullConstants.anIfNoShowRoot) Then
    ItemGetFirstSubItem = -1
    Exit Function
  End If

  If hParentItem Then
    ret = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_CHILD, hParentItem)
  Else
    ret = GetRootOrFirstItem
  End If

  If StateIconMask <> -1 Then
    If ItemStateIconIndex(ret) = StateIconMask Then
      If Not MustMatch Then
        ret = ItemGetNextItem(ret, StateIconMask, MustMatch)
      End If
    Else
      If MustMatch Then
        ret = ItemGetNextItem(ret, StateIconMask, MustMatch)
      End If
    End If
  End If

  If ret = 0 Then ret = -1
  ItemGetFirstSubItem = ret
End Function

Public Function ItemGetInfoTip(ByVal hItem As Long, ByVal InfoTipStyle As InfoTipStyleConstants) As String
Attribute ItemGetInfoTip.VB_Description = "Gibt den InfoTip des Items <hItem> zurück."
  Dim Cancel As Boolean
  Dim ret As String
  Dim TVWItemInfo As ITEMINFO_COMMON

  If Not GetItemInfo_Common(hItem, TVWItemInfo) Then Exit Function

  RaiseItemBeforeGetInfoTip hItem, InfoTipStyle, Cancel
  If Not Cancel Then
    #If Debuging Then
      ret = GetItemInfo(debugger, prophWndShellUIParentWindow, TVWItemInfo.pIDLToDesktop, InfoTipStyle)
    #Else
      ret = GetItemInfo(prophWndShellUIParentWindow, TVWItemInfo.pIDLToDesktop, InfoTipStyle)
    #End If

    ' der Desktop hat als InfoTip manchmal angeblich "Desktop"
    If ret = GetItemText(hItem) Then ret = ""
  End If

  ItemGetInfoTip = ret
End Function

' 0-basiert
Public Function ItemGetLevel(ByVal hItem As Long) As Long
Attribute ItemGetLevel.VB_Description = "Gibt die Ebene zurück, auf der sich der Item <hItem> befindet. Die oberste Ebene ist Ebene 0."
  Dim ret As Long

  If Not ItemExists(hItem, AllowNullConstants.anIfNoShowRoot) Then Exit Function

  hItem = ItemGetParentItem(hItem)
  While Not IsWithinRange(hItem, -1, 0)
    ret = ret + 1
    hItem = ItemGetParentItem(hItem)
  Wend

  ItemGetLevel = ret
End Function

' gibt den Nachbar-Item von <hItem> zurück
Public Function ItemGetNextItem(ByVal hItem As Long, Optional ByVal StateIconMask As Long = -1, Optional ByVal MustMatch As Boolean = True) As Long
Attribute ItemGetNextItem.VB_Description = "Gibt den nächsten Item hinter dem Item <hItem> zurück, der sich auf der selben Ebene befindet wie <hItem>."
  Dim ret As Long

  If ItemExists(hItem, AllowNullConstants.anNever) Then
    ret = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_NEXT, hItem)
    If StateIconMask <> -1 Then
      Do While ret
        If ItemStateIconIndex(ret) = StateIconMask Then
          If MustMatch Then
            Exit Do
          Else
            ret = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_NEXT, ret)
          End If
        Else
          If Not MustMatch Then
            Exit Do
          Else
            ret = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_NEXT, ret)
          End If
        End If
      Loop
    End If
  End If

  If ret = 0 Then ret = -1
  ItemGetNextItem = ret
End Function

Public Function ItemGetNextVisibleItem(ByVal hItem As Long, Optional ByVal StateIconMask As Long = -1, Optional ByVal MustMatch As Boolean = True) As Long
Attribute ItemGetNextVisibleItem.VB_Description = "Gibt den nächsten Item hinter dem Item <hItem> zurück. Er muß sich nicht auf der selben Ebene befinden, aber (ggf. durch scrollen) sichtbar sein."
  Dim ret As Long

  If ItemExists(hItem, AllowNullConstants.anNever) Then
    ret = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_NEXTVISIBLE, hItem)
    If StateIconMask <> -1 Then
      Do While ret
        If ItemStateIconIndex(ret) = StateIconMask Then
          If MustMatch Then
            Exit Do
          Else
            ret = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_NEXTVISIBLE, ret)
          End If
        Else
          If Not MustMatch Then
            Exit Do
          Else
            ret = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_NEXTVISIBLE, ret)
          End If
        End If
      Loop
    End If
  End If

  If ret = 0 Then ret = -1
  ItemGetNextVisibleItem = ret
End Function

' gibt den Parent-Item von <hItem> zurück
Public Function ItemGetParentItem(ByVal hItem As Long, Optional ByVal StateIconMask As Long = -1, Optional ByVal MustMatch As Boolean = True) As Long
Attribute ItemGetParentItem.VB_Description = "Gibt den Parent-Item des Items <hItem> zurück."
  Dim ret As Long

  If Not ItemExists(hItem, AllowNullConstants.anNever) Then
    ret = -1
    GoTo Ende
  End If

  ret = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_PARENT, hItem)
  If StateIconMask <> -1 Then
    Do While ret
      If ItemStateIconIndex(ret) = StateIconMask Then
        If MustMatch Then
          Exit Do
        Else
          ret = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_PARENT, ret)
        End If
      Else
        If Not MustMatch Then
          Exit Do
        Else
          ret = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_PARENT, ret)
        End If
      End If
    Loop
  End If

  If ret = 0 Then
    If propShowRoot Then ret = -1
  End If

Ende:
  ItemGetParentItem = ret
End Function

' gibt die pIDL vom Typ <pIDLType> von <hItem> zurück
Public Function ItemGetpIDL(ByVal hItem As Long, ByVal pIDLType As pIDLTypeConstants) As Long
Attribute ItemGetpIDL.VB_Description = "Gibt die ItemIDList des Items <hItem> zurück. Die ItemIDList muss freigegeben werden, wenn Sie nicht mehr benötigt wird."
  Dim pIDL As Long
  Dim TVWItemInfo As ITEMINFO_COMMON

  If inIDE Then Exit Function

  pIDL = GetFullyQualifiedPIDL(hItem)
  If pIDL Then
    Select Case pIDLType
      Case pIDLTypeConstants.ptToDesktop
        ItemGetpIDL = ILClone(pIDL)
      Case pIDLTypeConstants.ptToParent
        #If Debuging Then
          debugger.AddLogEntry "ItemGetpIDL: Calling CopyLastItemIDs()", LogEntryTypeConstants.letInfo
          debugger.AddLogEntry "   0x" & Hex(hItem), LogEntryTypeConstants.letOther
          debugger.AddLogEntry "   0x" & Hex(pIDL), LogEntryTypeConstants.letOther
          ItemGetpIDL = CopyLastItemIDs(debugger, pIDL, 1)
        #Else
          ItemGetpIDL = CopyLastItemIDs(pIDL, 1)
        #End If
      Case pIDLTypeConstants.ptToRoot
        pIDL = ILFindChild(RootItemInfo.pIDLToDesktop, pIDL)
        If pIDL Then ItemGetpIDL = ILClone(pIDL)
    End Select
  End If
End Function

' gibt den Nachbar-Item von <hItem> zurück
Public Function ItemGetPreviousItem(ByVal hItem As Long, Optional ByVal StateIconMask As Long = -1, Optional ByVal MustMatch As Boolean = True) As Long
Attribute ItemGetPreviousItem.VB_Description = "Gibt den nächsten Item vor dem Item <hItem> zurück, der sich auf der selben Ebene befindet wie <hItem>."
  Dim ret As Long

  If ItemExists(hItem, AllowNullConstants.anNever) Then
    ret = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_PREVIOUS, hItem)
    If StateIconMask <> -1 Then
      Do While ret
        If ItemStateIconIndex(ret) = StateIconMask Then
          If MustMatch Then
            Exit Do
          Else
            ret = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_PREVIOUS, ret)
          End If
        Else
          If Not MustMatch Then
            Exit Do
          Else
            ret = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_PREVIOUS, ret)
          End If
        End If
      Loop
    End If
  End If

  If ret = 0 Then ret = -1
  ItemGetPreviousItem = ret
End Function

Public Function ItemGetPreviousVisibleItem(ByVal hItem As Long, Optional ByVal StateIconMask As Long = -1, Optional ByVal MustMatch As Boolean = True) As Long
Attribute ItemGetPreviousVisibleItem.VB_Description = "Gibt den nächsten Item vor dem Item <hItem> zurück. Er muss sich nicht auf der selben Ebene befinden, aber (ggf. durch scrollen) sichtbar sein."
  Dim ret As Long

  If ItemExists(hItem, AllowNullConstants.anNever) Then
    ret = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_PREVIOUSVISIBLE, hItem)
    If StateIconMask <> -1 Then
      Do While ret
        If ItemStateIconIndex(ret) = StateIconMask Then
          If MustMatch Then
            Exit Do
          Else
            ret = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_PREVIOUSVISIBLE, ret)
          End If
        Else
          If Not MustMatch Then
            Exit Do
          Else
            ret = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_PREVIOUSVISIBLE, ret)
          End If
        End If
      Loop
    End If
  End If

  If ret = 0 Then ret = -1
  ItemGetPreviousVisibleItem = ret
End Function

' gibt das ItemRect von <hItem> zurück
Public Sub ItemGetRectangle(ByVal hItem As Long, rcItem As ITEMRECT, Optional ByVal entireRow As Boolean = False)
Attribute ItemGetRectangle.VB_Description = "Gibt das Rechteck zurück, welches den Item <hItem> umgibt."
  Dim rc As RECT

  GetItemRect hItem, rc, entireRow
  LSet rcItem = rc
End Sub

Public Function ItemGetSubItemCount(ByVal hParentItem As Long, Optional ByVal StateIconMask As Long = -1, Optional ByVal MustMatch As Boolean = True) As Long
Attribute ItemGetSubItemCount.VB_Description = "Gibt die Anzahl der direkten SubItems des Items <hParentItem> zurück."
  Dim hItem As Long
  Dim ret As Long

  If ItemExists(hParentItem, AllowNullConstants.anNever) Then
    hItem = Me.ItemGetFirstSubItem(hParentItem, StateIconMask, MustMatch)
    While hItem <> -1
      ret = ret + 1

      hItem = Me.ItemGetNextItem(hItem, StateIconMask, MustMatch)
    Wend
  End If

  ItemGetSubItemCount = ret
End Function

' gibt den Typ des Items <hItem> zurück
Public Function ItemGetType(ByVal hItem As Long) As ItemTypeConstants
Attribute ItemGetType.VB_Description = "Gibt den Typ des Items <hItem> zurück."
  Dim lParam As Long
  Dim ret As Byte

  If inIDE Then
    ItemGetType = -1
    Exit Function
  End If
  If Not ItemExists(hItem, AllowNullConstants.anIfNoShowRoot) Then
    ItemGetType = -1
    Exit Function
  End If

  If hItem Then
    lParam = GetItemlParam(hItem)
    CopyMemory VarPtr(ret), lParam, LenB(ret)
  Else
    ret = RootItemInfo.itemType
  End If

  ItemGetType = CLng(ret)
End Function

' gibt den AddressBar-DisplayName von <hItem> zurück
Public Function ItemHandleToAddressBarText(ByVal hItem As Long) As String
  Dim pIDL As Long

  If Not ItemExists(hItem, AllowNullConstants.anIfNoShowRoot) Then Exit Function

  pIDL = GetFullyQualifiedPIDL(hItem)
  If pIDL Then
    #If Debuging Then
      ItemHandleToAddressBarText = pIDLToDisplayName(debugger, IDesktop, pIDL, SHGDNConstants.SHGDN_FORPARSING Or SHGDNConstants.SHGDN_FORADDRESSBAR)
    #Else
      ItemHandleToAddressBarText = pIDLToDisplayName(IDesktop, pIDL, SHGDNConstants.SHGDN_FORPARSING Or SHGDNConstants.SHGDN_FORADDRESSBAR)
    #End If
  Else
    ItemHandleToAddressBarText = GetItemText(hItem)
  End If
End Function

' gibt den DisplayName von <hItem> zurück
Public Function ItemHandleToDisplayName(ByVal hItem As Long) As String
Attribute ItemHandleToDisplayName.VB_Description = "Gibt den DisplayName des Items <hItem> zurück."
  ItemHandleToDisplayName = GetItemText(hItem)
End Function

' gibt den Pfad von <hItem> zurück
Public Function ItemHandleToFSPath(ByVal hItem As Long) As String
Attribute ItemHandleToFSPath.VB_Description = "Gibt den Pfad des Items <hItem> zurück. Dabei wird für Items des Dateisystems immer ein korrekter Dateisystempfad zurückgegeben."
  If Not ItemExists(hItem, AllowNullConstants.anIfNoShowRoot) Then Exit Function
  #If Debuging Then
    ItemHandleToFSPath = pIDLToPath_Light(debugger, GetFullyQualifiedPIDL(hItem))
  #Else
    ItemHandleToFSPath = pIDLToPath_Light(GetFullyQualifiedPIDL(hItem))
  #End If
End Function

' gibt den Pfad von <hItem> zurück
Public Function ItemHandleToTreePath(ByVal hItem As Long) As String
Attribute ItemHandleToTreePath.VB_Description = "Gibt den Pfad des Items <hItem> zurück."
  ItemHandleToTreePath = Replace$(ItemHandleToPath(hItem), ArchiveSignature, ":")
End Function

Public Function ItemHasChildren(ByVal hItem As Long) As Boolean
Attribute ItemHasChildren.VB_Description = "Gibt zurück, ob der Item <hItem> SubItems hat. Die SubItems müssen nicht geladen sein."
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Function
  ItemHasChildren = GetItemChildrenProperty(hItem)
End Function

' gibt zurück, ob <hItem> eine Datei in einem Archiv ist
Public Function ItemIsArchiveFile(ByVal hItem As Long) As Boolean
Attribute ItemIsArchiveFile.VB_Description = "Gibt zurück, ob der Item <hItem> eine Datei in einem Archiv ist, welches nach <ExpandArchives> als Ordner behandelt wird."
  ItemIsArchiveFile = (ItemGetType(hItem) = ItemTypeConstants.itArchiveFile)
End Function

' gibt zurück, ob <hItem> ein Ordner in einem Archiv ist
Public Function ItemIsArchiveFolder(ByVal hItem As Long) As Boolean
Attribute ItemIsArchiveFolder.VB_Description = "Gibt zurück, ob der Item <hItem> ein Ordner in einem Archiv ist, welches nach <ExpandArchives> als Ordner behandelt wird."
  ItemIsArchiveFolder = (ItemGetType(hItem) = ItemTypeConstants.itArchiveFolder)
End Function

' gibt zurück, ob <hItem> ein Archiv ist
' der Einfachkeit wegen wird nur geprüft, ob <hItem> eine FS-Datei und erweiterbar ist
Public Function ItemIsArchiveToExpand(ByVal hItem As Long) As Boolean
Attribute ItemIsArchiveToExpand.VB_Description = "Gibt zurück, ob der Item <hItem> ein Archiv ist, welches nach <ExpandArchives> als Ordner behandelt wird."
  If ItemIsFSFile(hItem) Then
    ItemIsArchiveToExpand = GetItemChildrenProperty(hItem)
  End If
End Function

' gibt zurück, ob <hItem> erweitert ist
Public Function ItemIsExpanded(ByVal hItem As Long) As Boolean
Attribute ItemIsExpanded.VB_Description = "Gibt zurück, ob der Item <hItem> erweitert ist."
  Dim ret As Boolean

  If Not ItemExists(hItem, AllowNullConstants.anIfNoShowRoot) Then Exit Function

  If hItem Then
    ret = (GetItemState(hItem, TVIS_EXPANDED) And TVIS_EXPANDED)
  Else
    ret = True
  End If
  ItemIsExpanded = ret
End Function

' gibt zurück, ob <hItem> eine Datei ist
Public Function ItemIsFile(ByVal hItem As Long) As Boolean
Attribute ItemIsFile.VB_Description = "Gibt zurück, ob der Item <hItem> eine Datei ist."
  Dim tmp As ItemTypeConstants

  tmp = ItemGetType(hItem)
  ItemIsFile = ((tmp = ItemTypeConstants.itArchiveFile) Or (tmp = ItemTypeConstants.itFSFile) Or (tmp = ItemTypeConstants.itNonFSFile))
End Function

' gibt zurück, ob <hItem> ein Ordner ist
Public Function ItemIsFolder(ByVal hItem As Long) As Boolean
Attribute ItemIsFolder.VB_Description = "Gibt zurück, ob der Item <hItem> ein Ordner ist."
  Dim tmp As ItemTypeConstants

  tmp = ItemGetType(hItem)
  ItemIsFolder = ((tmp = ItemTypeConstants.itArchiveFolder) Or (tmp = ItemTypeConstants.itFSFolder) Or (tmp = ItemTypeConstants.itNonFSFolder))
End Function

' gibt zurück, ob <hItem> eine FS-Datei ist
Public Function ItemIsFSFile(ByVal hItem As Long) As Boolean
Attribute ItemIsFSFile.VB_Description = "Gibt zurück, ob der Item <hItem> eine Datei ist, die zum Dateisystem gehört."
  ItemIsFSFile = (ItemGetType(hItem) = ItemTypeConstants.itFSFile)
End Function

' gibt zurück, ob <hItem> ein FS-Ordner ist
Public Function ItemIsFSFolder(ByVal hItem As Long) As Boolean
Attribute ItemIsFSFolder.VB_Description = "Gibt zurück, ob der Item <hItem> ein Ordner ist, der zum Dateisystem gehört."
  ItemIsFSFolder = (ItemGetType(hItem) = ItemTypeConstants.itFSFolder)
End Function

' gibt zurück, ob <hItem> ein Link ist
Public Function ItemIsLink(ByVal hItem As Long, Optional ByRef target As String = "") As Boolean
Attribute ItemIsLink.VB_Description = "Gibt zurück, ob <hItem> eine Verknüpfung ist. <Target> wird ggf. das Ziel der Verknüpfung zugewiesen."
  Dim IParent As IVBShellFolder
  Dim pIDLToParent As Long
  Dim ret As Boolean
  Dim TVWItemInfo As ITEMINFO_COMMON

  If GetItemInfo_Common(hItem, TVWItemInfo) Then
    #If Debuging Then
      SplitFullyQualifiedPIDL debugger, TVWItemInfo.pIDLToDesktop, IParent, pIDLToParent
      ret = HasAttribute(IParent, pIDLToParent, SFGAOConstants.SFGAO_LINK)
      If ret Then
        target = GetLinkTarget(debugger, IParent, pIDLToParent)
      End If
    #Else
      SplitFullyQualifiedPIDL TVWItemInfo.pIDLToDesktop, IParent, pIDLToParent
      ret = HasAttribute(IParent, pIDLToParent, SFGAOConstants.SFGAO_LINK)
      If ret Then
        target = GetLinkTarget(IParent, pIDLToParent)
      End If
    #End If
    Set IParent = Nothing
  End If

  ItemIsLink = ret
End Function

' gibt zurück, ob <hItem> eine NonFS-Datei ist
Public Function ItemIsNonFSFile(ByVal hItem As Long) As Boolean
Attribute ItemIsNonFSFile.VB_Description = "Gibt zurück, ob der Item <hItem> eine Datei ist, die nicht zum Dateisystem gehört und sich nicht in einem Archiv befindet, welches nach <ExpandArchives> als Ordner behandelt wird."
  ItemIsNonFSFile = (ItemGetType(hItem) = ItemTypeConstants.itNonFSFile)
End Function

' gibt zurück, ob <hItem> ein NonFS-Ordner ist
Public Function ItemIsNonFSFolder(ByVal hItem As Long) As Boolean
Attribute ItemIsNonFSFolder.VB_Description = "Gibt zurück, ob der Item <hItem> ein Ordner ist, der nicht zum Dateisystem gehört und sich nicht in einem Archiv befindet, welches nach <ExpandArchives> als Ordner behandelt wird."
  ItemIsNonFSFolder = (ItemGetType(hItem) = ItemTypeConstants.itNonFSFolder)
End Function

' gibt zurück, ob <hItem> in einem Archiv ist
Public Function ItemIsPartOfArchive(ByVal hItem As Long) As Boolean
Attribute ItemIsPartOfArchive.VB_Description = "Gibt zurück, ob der Item <hItem> in einem Archiv ist, welches nach <ExpandArchives> als Ordner behandelt wird."
  Dim tmp As ItemTypeConstants

  tmp = ItemGetType(hItem)
  ItemIsPartOfArchive = ((tmp = ItemTypeConstants.itArchiveFile) Or (tmp = ItemTypeConstants.itArchiveFolder))
End Function

' gibt zurück, ob <hItem> zum Dateisystem gehört
Public Function ItemIsPartOfFileSystem(ByVal hItem As Long) As Boolean
Attribute ItemIsPartOfFileSystem.VB_Description = "Gibt zurück, ob der Item <hItem> Teil des Dateisystems ist."
  Dim tmp As ItemTypeConstants

  tmp = ItemGetType(hItem)
  ItemIsPartOfFileSystem = ((tmp = ItemTypeConstants.itFSFile) Or (tmp = ItemTypeConstants.itFSFolder))
End Function

Public Sub ItemLoadSubItems(ByVal hParentItem As Long)
Attribute ItemLoadSubItems.VB_Description = "Lädt die direkten SubItems des Items <hParentItem>."
  If Not ItemExists(hParentItem, AllowNullConstants.anIfNoShowRoot) Then Exit Sub

  If Not ItemHasSubItems(hParentItem) Then InsertSubItems hParentItem
  If Not ItemHasSubItems(hParentItem) Then
    dontChangeRecyclerChildrenProp = True
    SetItemChildrenProperty hParentItem, False
  End If
End Sub

' zeigt das Shell-Kontextmenü für <hItem> an der Position [<x>;<y>] an
Public Sub ItemShowShellContextMenu(ByVal hItem As Long, ByVal x As Long, ByVal y As Long)
Attribute ItemShowShellContextMenu.VB_Description = "Zeigt das Shell-Kontextmenü des Items <hItem> an der Stelle [<x>;<y>] an."
  Const VK_SHIFT = &H10
  Dim Cancel As Boolean
  Dim CmdID As Long
  Dim Flags As ShellContextMenuStyleConstants
  Dim InvokeData As CMINVOKECOMMANDINFOEX
  Dim IParent As IVBShellFolder
  Dim ItemDescr As String
  Dim ItemVerb As String
  Dim pIDL As Long
  Dim pIDLToParent As Long
  Dim Style As WindowModeConstants

  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub

  ' Daten über <hItem> sammeln
  pIDL = GetFullyQualifiedPIDL(hItem, True)
  If pIDL = 0 Then Exit Sub
  #If Debuging Then
    SplitFullyQualifiedPIDL debugger, pIDL, IParent, pIDLToParent
  #Else
    SplitFullyQualifiedPIDL pIDL, IParent, pIDLToParent
  #End If

  hItemWithContMenu = hItem
  If (pIDLToParent <> 0) And Not (IParent Is Nothing) Then
    ' Interface "IContextMenu" ermitteln
    IParent.GetUIObjectOf prophWndShellUIParentWindow, 1, pIDLToParent, IID_IContextMenu, 0, IContextMenu

    If Not (IContextMenu Is Nothing) Then
      ' Interface "IContextMenu2" ermitteln
      IContextMenu.QueryInterface IID_IContextMenu2, IContextMenu2

      If Not (IContextMenu2 Is Nothing) Then
        ' Interface "IContextMenu3" ermitteln
        IContextMenu2.QueryInterface IID_IContextMenu3, IContextMenu3
      End If
    End If

    ' Menü erzeugen und füllen
    If IsKeyHoldDown(VK_SHIFT) Then
      Flags = ShellContextMenuStyleConstants.scmsExtendedVerbs
    Else
      Flags = ShellContextMenuStyleConstants.scmsNormal
    End If

    RaiseItemBeforeInitShellContextMenu hItem, Flags, Cancel
    If Not Cancel Then
      If propRenaming = RenamingConstants.rAutomatic Then Flags = Flags Or CMFConstants.CMF_CANRENAME
      prophShellContextMenu = CreatePopupMenu
      If Not (IContextMenu Is Nothing) Then IContextMenu.QueryContextMenu prophShellContextMenu, 0, 1, &H7FFF, Flags

      ' WORKAROUND: bei Zip-Archiven werden teilweise unter WinMe/XP 2 Separators hintereinander
      '             eingefügt -> diese entfernen
      RemoveDoubleSeparators prophShellContextMenu

      RaiseItemInitShellContextMenu hItem

      ' Menü anzeigen
      RaiseItemBeforeShowShellContextMenu hItem, Cancel
      If Not Cancel Then
        SetItemState SelectedItem, 0, TVIS_SELECTED
        SetItemState hItem, TVIS_SELECTED, TVIS_SELECTED
        CmdID = TrackPopupMenuExAsLong(prophShellContextMenu, TPM_LEFTALIGN Or TPM_LEFTBUTTON Or TPM_RIGHTBUTTON Or TPM_RETURNCMD, x, y, prophWnd, 0)
        SetItemState hItem, 0, TVIS_SELECTED
        SetItemState SelectedItem, TVIS_SELECTED, TVIS_SELECTED
        If CmdID Then
          ItemDescr = String$(MAX_PATH, Chr$(0))
          ItemVerb = String$(MAX_PATH, Chr$(0))
          If Not (IContextMenu Is Nothing) Then
            IContextMenu.GetCommandString CmdID - 1, GCSConstants.GCS_HELPTEXT Or GCSConstants.GCS_UNICODE, 0, StrPtr(ItemDescr), Len(ItemDescr)
            IContextMenu.GetCommandString CmdID - 1, GCSConstants.GCS_VERB Or GCSConstants.GCS_UNICODE, 0, StrPtr(ItemVerb), Len(ItemVerb)
          End If
          ItemDescr = Left$(ItemDescr, lstrlenA(ItemDescr))
          ItemVerb = Left$(ItemVerb, lstrlenA(ItemVerb))

          Style = WindowModeConstants.wmShowDefault
          RaiseItemBeforeExecuteShellContextMenuItem hItem, CmdID, ItemVerb, ItemDescr, Cancel, Style

          If Not Cancel Then
            If LCase$(ItemVerb) = "rename" Then
              ' ein Item soll umbenannt werden
              SendMessageAsLong prophWnd, TVM_EDITLABEL, 0, hItem
              RaiseItemExecutedShellContextMenuItem hItem, CmdID, ItemVerb, ItemDescr
            Else
              ' sonstige MenuItems
              With InvokeData
                .StructSize = LenB(InvokeData)
                .hWnd = prophWnd
                .Verb = MakeDWord(CmdID - 1, 0)
                .Showflags = Style
                If Not (IContextMenu Is Nothing) Then
                  If IContextMenu.InvokeCommand(InvokeData) = NOERROR Then RaiseItemExecutedShellContextMenuItem hItem, CmdID, ItemVerb, ItemDescr
                End If
              End With
            End If
          End If
        End If
      Else
        DestroyMenu prophShellContextMenu
        prophShellContextMenu = 0
        RaiseItemDestroyShellContextMenu hItem
      End If
    End If

    Set IContextMenu = Nothing
    Set IContextMenu2 = Nothing
    Set IContextMenu3 = Nothing
  End If
  Set IParent = Nothing

  hItemWithContMenu = 0
  DestroyMenu prophShellContextMenu
  prophShellContextMenu = 0
End Sub

Public Function ItemStartRenaming(ByVal hItem As Long) As Boolean
Attribute ItemStartRenaming.VB_Description = "Startet das Umbenennen des Items <hItem>. Sollte sich <hItem> nicht umbenennen lassen, wird False zurück gegeben."
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Function

  SetFocus
  ItemStartRenaming = (SendMessageAsLong(prophWnd, TVM_EDITLABEL, 0, hItem) <> 0)
End Function

' TODO: Win-Explorer seems to look for SFGAO_FILESYSTEM | SFGAO_FILESYSANCESTOR | SFGAO_FOLDER
Public Function ItemSupportsNewFolder(ByVal hItem As Long) As Boolean
Attribute ItemSupportsNewFolder.VB_Description = "Gibt zurück, ob der Item <hItem> den Befehl ""Neuer Ordner"" unterstützt."
  Dim IParent As IVBShellFolder
  Dim pIDL As Long
  Dim pIDLToParent As Long
  Dim TVWItemInfo_Common As ITEMINFO_COMMON

  If Not ItemExists(hItem, AllowNullConstants.anIfNoShowRoot) Then Exit Function

  If ItemIsFSFolder(hItem) Then
    pIDL = GetFullyQualifiedPIDL(hItem, True)
    If pIDL = 0 Then Exit Function
    #If Debuging Then
      SplitFullyQualifiedPIDL debugger, pIDL, IParent, pIDLToParent
    #Else
      SplitFullyQualifiedPIDL pIDL, IParent, pIDLToParent
    #End If
    ItemSupportsNewFolder = HasAttribute(IParent, pIDLToParent, SFGAOConstants.SFGAO_FILESYSANCESTOR)
    Set IParent = Nothing
  End If
End Function
'
'Public Function ItemSupportsVerb(ByVal hItem As Long, ByVal Verb As String) As Boolean
'  Dim i As Integer
'  Dim ret As Boolean
'  Dim SupportedVerbs() As String
'
'  If Not ItemExists(hItem, AllowNullConstants.anIfNoShowRoot) Then Exit Function
'  If Trim(Verb) = "" Then Exit Function
'
'  If GetItemSupportedVerbs(hItem, SupportedVerbs) Then
'    For i = LBound(SupportedVerbs) To UBound(SupportedVerbs)
'      If LCase(Verb) = LCase(SupportedVerbs(i)) Then
'        ' gefunden
'        ret = True
'        Exit For
'      End If
'    Next
'  End If
'
'  ItemSupportsVerb = ret
'End Function

Public Function ItemTextIsCut(ByVal hItem As Long) As Boolean
Attribute ItemTextIsCut.VB_Description = "Gibt zurück, ob der Text des Items <hItem> nur teilweise sichtbar ist."
  Dim Offset As Long
  Dim rcItem As RECT
  Dim ret As Boolean
  Dim ScrollBarWidth As Long
  Dim Style As Long

  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Function

  GetItemRect hItem, rcItem
  ' damit der Item verdeckt ist, muß <rcItem.Left> kleiner sein als <rcClient.Left> oder
  ' <rcItem.Right> muß größer sein als <rcClient.Right>
  ' prüfen, ob das Control einen vertikalen ScrollBar hat
  Style = GetWindowLongPtr(prophWnd, GWL_STYLE)
  If Style And WindowStyleConsts.WS_VSCROLL Then
    ' Breite des ScrollBars ermitteln
    ScrollBarWidth = GetSystemMetrics(SM_CXVSCROLL)
  End If
  With rcItem
'    ' 3 Pixel abziehen (2 für FokusRect-Grenze, 1 zwischen Schrift und FokusRect)
'    If (.Top > rcClient.Top) And (.Bottom < rcClient.Bottom) Then ret = ((.Left < rcClient.Left) Or (.Right - 3 > rcClient.Right - ScrollBarWidth))
    If (.Top > rcClient.Top) And (.Bottom < rcClient.Bottom) Then ret = ((.Left < rcClient.Left) Or (.Right - propMargin > rcClient.Right - ScrollBarWidth))
  End With

  ItemTextIsCut = ret
End Function

' gibt zurück, ob <hItem> schon einmal erweitert war
Public Function ItemWasExpandedOnce(ByVal hItem As Long) As Boolean
Attribute ItemWasExpandedOnce.VB_Description = "Gibt zurück, ob der Item <hItem> bereits erweitert wurde."
  If Not ItemExists(hItem, AllowNullConstants.anIfNoShowRoot) Then Exit Function

  If hItem Then
    ItemWasExpandedOnce = (GetItemState(hItem, TVIS_EXPANDEDONCE) And TVIS_EXPANDEDONCE)
  Else
    ItemWasExpandedOnce = True
  End If
End Function

Public Function LoadSettingsFromFile(ByVal File As String) As Boolean
Attribute LoadSettingsFromFile.VB_Description = "Übernimmt alle Einstellungen aus der Datei <File>."
  Dim AdditionalData As Variant
  Dim arrProps() As Byte
  Dim BytesRead As Long
  Dim Flags As ErrorTypeConstants
  Dim hFile As Long
  Dim isVersion10 As Boolean
  Dim isVersion20 As Boolean
  Dim lngBuffer As Long
  Dim MSG As String
  Dim pos As Integer
  Dim Props As PropertyBag
  Dim ret As Long
  Dim Start As Long
  Dim strBuffer As String

  On Error GoTo Fehler

  ' prüfen, ob die Datei Daten für ExplorerTreeView enthält
  If Not FileContainsValidSettings(File, Flags, AdditionalData) Then
    Select Case Flags
      Case ErrorTypeConstants.etCanNotOpenFile
        GoTo CreateFileError
      Case ErrorTypeConstants.etCanNotReadFile
        GoTo ReadFileError
      Case ErrorTypeConstants.etFileNotFound
        GoTo FileNotFound
      Case ErrorTypeConstants.etInvalidFormat
        GoTo InvalidFormat
      Case ErrorTypeConstants.etNoSettingsFound
        GoTo NoSettingsFound
      Case ErrorTypeConstants.etTooNewFileFormatToRead
        GoTo TooNewFileFormat_Read
      Case ErrorTypeConstants.etUnknown
        GoTo Fehler
    End Select
  End If

  ' Datei öffnen
  hFile = CreateFileAsLong(File, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
  If hFile = INVALID_HANDLE_VALUE Then GoTo CreateFileError

  ' die Signatur wurde bereits geprüft...
  SetFilePointer hFile, 7, 0, FILE_BEGIN

  ' Formatversion einlesen...
  strBuffer = String$(4, Chr$(0))
  ret = ReadFileAsLong(hFile, ByVal strBuffer, Len(strBuffer), BytesRead, 0)
  If ret = 0 Then GoTo ReadFileError
  strBuffer = Left$(strBuffer, lstrlenA(strBuffer))
  '...und prüfen
  pos = InStr(strBuffer, ".")
  If pos = 0 Then pos = Len(strBuffer)
  isVersion20 = (CInt(Left(strBuffer, pos - 1)) >= 2)
  isVersion10 = (CInt(Left(strBuffer, pos - 1)) = 1)

  ' gleich zur Adresse der Daten springen
  If isVersion10 Then
    SetFilePointer hFile, 42, 0, FILE_CURRENT
  ElseIf isVersion20 Then
    SetFilePointer hFile, 114, 0, FILE_CURRENT
  End If

  ' Startbyte der Eigenschaften einlesen
  ret = ReadFileAsLong(hFile, ByVal VarPtr(Start), LenB(Start), BytesRead, 0)
  If ret = 0 Then GoTo ReadFileError
  If isVersion10 Then Start = Start + 4     ' Bug von Version 1.0

  ' Länge der Eigenschaften einlesen
  ret = ReadFileAsLong(hFile, ByVal VarPtr(lngBuffer), LenB(lngBuffer), BytesRead, 0)
  If ret = 0 Then GoTo ReadFileError

  ' Eigenschaften einlesen...
  SetFilePointer hFile, Start, 0, FILE_BEGIN
  ReDim arrProps(lngBuffer - 1)
  ret = ReadFileAsLong(hFile, ByVal VarPtr(arrProps(LBound(arrProps))), ArraySize_AsByte(arrProps), BytesRead, 0)
  If ret = 0 Then GoTo ReadFileError

  ' ...und übertragen
  Set Props = New PropertyBag
  Props.Contents = arrProps
  Erase arrProps
  UserControl_ReadProperties Props
  disableCreateControl = False
  ' ggf. Control neu erzeugen
  If needRecreate Then CreateControl

  ' Datei schließen
  CloseHandle hFile

  LoadSettingsFromFile = True
  Exit Function

CreateFileError:
  ' CreateFile ist fehlgeschlagen
  On Error GoTo 0
  CloseHandle hFile
  MSG = LoadResString(resCreateFileError)
  MSG = Replace(MSG, "|1", File)
  MSG = Replace(MSG, "|2", Err.LastDllError)
  strBuffer = Space(500)
  FormatMessage FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, ByVal 0&, Err.LastDllError, &H400, strBuffer, Len(strBuffer), 0
  MSG = Replace(MSG, "|3", removeNull(strBuffer))
  MSG = Replace(MSG, "\n", vbNewLine)
  Err.Raise vbObjectError + errCreateFileError, App.EXEName & ".LoadSettingsFromFile", MSG
  Exit Function

Fehler:
  ' unbekannter Fehler
  CloseHandle hFile
  Err.Raise IIf(Err.Number, Err.Number, CLng(AdditionalData)), App.EXEName & ".LoadSettingsFromFile"
  Exit Function

FileNotFound:
  ' Datei nicht gefunden
  On Error GoTo 0
  CloseHandle hFile
  Err.Raise 53, App.EXEName & ".LoadSettingsFromFile"
  Exit Function

InvalidFormat:
  ' ungültiges Dateiformat
  On Error GoTo 0
  CloseHandle hFile
  Err.Raise 321, App.EXEName & ".LoadSettingsFromFile"
  Exit Function

NoSettingsFound:
  ' es wurden keine Daten gefunden
  On Error GoTo 0
  CloseHandle hFile
  MSG = LoadResString(resNoSettingsFound)
  MSG = Replace(MSG, "|1", File)
  MSG = Replace(MSG, "\n", vbNewLine)
  Err.Raise vbObjectError + errNoSettingsFound, App.EXEName & ".LoadSettingsFromFile", MSG
  Exit Function

ReadFileError:
  ' ReadFile ist fehlgeschlagen
  On Error GoTo 0
  CloseHandle hFile
  MSG = LoadResString(resReadFileError)
  MSG = Replace(MSG, "|1", File)
  MSG = Replace(MSG, "|2", Err.LastDllError)
  strBuffer = Space(500)
  FormatMessage FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, ByVal 0&, Err.LastDllError, 1024, strBuffer, Len(strBuffer), 0
  MSG = Replace(MSG, "|3", removeNull(strBuffer))
  MSG = Replace(MSG, "\n", vbNewLine)
  Err.Raise vbObjectError + errReadFileError, App.EXEName & ".LoadSettingsFromFile", MSG
  Exit Function

TooNewFileFormat_Read:
  ' das Dateiformat ist zu neu zum Lesen
  On Error GoTo 0
  CloseHandle hFile
  MSG = LoadResString(resTooNewFileFormat_Read)
  MSG = Replace(MSG, "|1", File)
  MSG = Replace(MSG, "|2", AppName)
  MSG = Replace(MSG, "|3", CStr(AdditionalData))
  MSG = Replace(MSG, "|4", Version)
  MSG = Replace(MSG, "\n", vbNewLine)
  Err.Raise vbObjectError + errTooNewFileFormat_Read, App.EXEName & ".LoadSettingsFromFile", MSG
End Function

'Public Function PathToItemHandle(ByVal Pathname As String, Optional ByVal insertIfNecessary As Boolean = False, Optional ByVal forceInsertion As Boolean = False) As Long
Public Function PathToItemHandle(ByVal Pathname As String, Optional ByVal insertIfNecessary As Boolean = False) As Long
Attribute PathToItemHandle.VB_Description = "Sucht zu <Pathname> den passenden Item."
  Dim ret As Long

  If LCase$(Pathname) = LCase$(propPath) Then
    ret = Me.SelectedItem
  Else
    If GetCommonPrefix(LCase$(Pathname), LCase$(propPath)) = LCase$(RemoveBackslash(propPath, False)) Then
      ret = SelectedItem
    Else
      ret = -2
    End If
    ret = InternalpIDLToItemHandle(UserPathTopIDL(Pathname, True), ret, insertIfNecessary, True)
  End If

  PathToItemHandle = ret
End Function

Public Function pIDLToItemHandle(ByVal pIDLToDesktop As Long, Optional ByVal insertIfNecessary As Boolean = False) As Long
Attribute pIDLToItemHandle.VB_Description = "Sucht zu <pIDLToDesktop> den passenden Item."
  pIDLToItemHandle = InternalpIDLToItemHandle(pIDLToDesktop, , insertIfNecessary)
End Function

' zeichnet das Control neu
Public Sub Refresh()
Attribute Refresh.VB_Description = "Erzwingt ein vollständiges Neuzeichnen des Controls."
Attribute Refresh.VB_UserMemId = -550
  If prophWnd Then
    InvalidateRectAsLong prophWnd, 0, 1
    UpdateWindow prophWnd
  End If
End Sub

' lädt alle Icons neu
' initialisiert vorher die SysImageList neu
Public Sub ReloadIcons()
Attribute ReloadIcons.VB_Description = "Lädt alle Icons neu."
  Const WM_SETREDRAW = &HB
  Dim Size_LargeIcons As Long
  Dim Size_SmallIcons As Long

  If propAlwaysUseSmallIcons Then
    useLargeIcons = False
  Else
    Size_LargeIcons = GetSystemMetrics(SM_CXICON)
    Size_SmallIcons = GetSystemMetrics(SM_CXSMICON)
    useLargeIcons = (propItemHeight >= Size_LargeIcons * 0.75)
  End If

  If (prophImageList <> 0) And Not propUseCustomIcons Then
    SendMessageAsLong prophWnd, WM_SETREDRAW, 0, 0

    SendMessageAsLong prophWnd, TVM_SETIMAGELIST, TVSIL_NORMAL, 0
    prophImageList = 0

    ' SysImageList initialisieren
    ' -> mit Standard-Icons und den Overlays
    If ver_Win_NTBased Then FileIconInit 0

    prophImageList = GetSysImageList(useLargeIcons)
    SendMessageAsLong prophWnd, TVM_SETIMAGELIST, TVSIL_NORMAL, prophImageList
    If propCheckBoxes Then
      SendMessageAsLong prophWnd, TVM_SETIMAGELIST, TVSIL_STATE, IIf(useLargeIcons, prophLargeStateImageList, prophSmallStateImageList)
    End If

    If Not propUseCustomIcons Then
      UpdateDefaultIconIndices
      InvalidateIcons 0, -1, True
    End If

    SendMessageAsLong prophWnd, WM_SETREDRAW, 1, 0
    Refresh
  End If
End Sub

' lädt alle Items neu
Public Sub ReloadItems()
Attribute ReloadItems.VB_Description = "Lädt alle Items neu."
  Dim buffer As String
  Dim expanded As Boolean
  Dim hItem As Long

  hItem = Me.SelectedItem
  buffer = Me.ItemHandleToTreePath(hItem)
  expanded = Me.ItemIsExpanded(hItem)
  suppressSelChangeEvents = True
  dontSavePath = True
  InitTreeView
  dontSavePath = False
  SelectPath propPath, False
  suppressSelChangeEvents = False

  hItem = Me.SelectedItem
  If LCase$(Me.ItemHandleToTreePath(hItem)) = LCase$(buffer) Then
    If expanded Then Me.ItemExpand hItem
  Else
    SelChangeInternal = True
    RaiseSelChanged 0, hItem, cbInternal, False
  End If
End Sub

' lädt alle Sub-Items von <hParentItem> neu
Public Sub ReloadSubItems(ByVal hParentItem As Long)
Attribute ReloadSubItems.VB_Description = "Lädt alle SubItems von <hParentItem> neu."
  Dim buffer As String
  Dim expanded As Boolean
  Dim hItem As Long

  If Not ItemExists(hParentItem, AllowNullConstants.anIfNoShowRoot) Then Exit Sub

  hItem = Me.SelectedItem
  buffer = Me.ItemHandleToTreePath(hItem)
  expanded = Me.ItemIsExpanded(hItem)
  suppressSelChangeEvents = True
  dontSavePath = True
  RemoveSubItems hParentItem
  ItemLoadSubItems hParentItem
  dontSavePath = False
  SelectPath propPath, False
  suppressSelChangeEvents = False

  hItem = Me.SelectedItem
  If LCase$(Me.ItemHandleToTreePath(hItem)) = LCase$(buffer) Then
    If expanded Then Me.ItemExpand hItem
  Else
    SelChangeInternal = True
    RaiseSelChanged 0, hItem, cbInternal, False
  End If
End Sub

Public Function SaveSettingsToFile(ByVal File As String, Optional ByVal OverwriteSettings As Boolean = True, Optional ByRef OverwriteNewerFormat As Boolean = False) As Boolean
Attribute SaveSettingsToFile.VB_Description = "Speichert alle Einstellungen in der Datei <File>."
  Dim AdditionalData As Variant
  Dim appendDataToFile As Boolean
  Dim Flags As ErrorTypeConstants
  Dim MSG As String
  Dim neededVersionToWrite As Long
  Dim neededVersionToWrite_Build As Integer
  Dim ret2 As Integer
  Dim strBuffer As String
  Dim Version_Major As Byte
  Dim Version_Minor As Byte
  Dim Version_Revision As Integer
  Dim writeNewFile As Boolean

  On Error GoTo Fehler

  writeNewFile = Not FileExists(File)
  If Not writeNewFile Then
    If FileContainsValidSettings(File, Flags, AdditionalData) Then
      If OverwriteSettings Then
        ' wenn <File> im Format 1.0 vorliegt, Datei gleich komplett neu anlegen, sonst Datei aktualisieren
        ret2 = GetSettingsFileFormatVersion(File, neededVersionToWrite, neededVersionToWrite_Build)
        If (HiByte(ret2) = 1) And (LoByte(ret2) = 0) Then
          ' Version 1.0
          writeNewFile = True
        Else
          ret2 = HiWord(neededVersionToWrite)
          Version_Major = HiByte(ret2)
          Version_Minor = LoByte(ret2)
          Version_Revision = LoWord(neededVersionToWrite)
          If App.Major > Version_Major Then
            appendDataToFile = True
          ElseIf App.Major = Version_Major Then
            If App.Minor > Version_Minor Then
              appendDataToFile = True
            ElseIf App.Minor = Version_Minor Then
              If App.Revision > Version_Revision Then
                appendDataToFile = True
              ElseIf App.Revision = Version_Revision Then
                appendDataToFile = (BuildNumber >= neededVersionToWrite_Build)
              End If
            End If
          End If

          If Not appendDataToFile Then GoTo TooNewFileFormat_Write
        End If
      Else
        ' <File> enthält bereits gültige Daten, wir sollen aber nichts überschreiben
        GoTo SettingsAlreadyExists
      End If
    Else
      Select Case Flags
        Case ErrorTypeConstants.etCanNotOpenFile
          GoTo CreateFileError
        Case ErrorTypeConstants.etCanNotReadFile
          GoTo ReadFileError
        Case ErrorTypeConstants.etCanNotWriteFile
          GoTo WriteFileError
        Case ErrorTypeConstants.etInvalidFormat
          GoTo InvalidFormat
        Case ErrorTypeConstants.etNoSettingsFound
          ' <File> enthält keine Daten
          ' -> aktualisieren
          appendDataToFile = True
        Case ErrorTypeConstants.etTooNewFileFormatToRead
          ' Dateiformat von <File> ist zu neu zum Lesen
          If OverwriteNewerFormat Then
            ' wir sollen trotzdem überschreiben
            ' -> Datei gleich komplett neu anlegen
            writeNewFile = True
          Else
            GoTo TooNewFileFormat_Read
          End If
        Case ErrorTypeConstants.etUnknown
          GoTo Fehler
      End Select
    End If
  End If

  If writeNewFile Then
    Flags = WriteSettingsFile(File, AdditionalData)
  ElseIf appendDataToFile Then
    Flags = UpdateSettingsFile(File, AdditionalData)
  End If
  Select Case Flags
    Case ErrorTypeConstants.etCanNotOpenFile
      GoTo CreateFileError
    Case ErrorTypeConstants.etCanNotReadFile
      GoTo ReadFileError
    Case ErrorTypeConstants.etCanNotWriteFile
      GoTo WriteFileError
    Case ErrorTypeConstants.etUnknown
      GoTo Fehler
  End Select

  SaveSettingsToFile = True
  Exit Function

CreateFileError:
  ' CreateFile ist fehlgeschlagen
  On Error GoTo 0
  MSG = LoadResString(resCreateFileError)
  MSG = Replace(MSG, "|1", File)
  MSG = Replace(MSG, "|2", Err.LastDllError)
  strBuffer = Space(500)
  FormatMessage FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, ByVal 0&, Err.LastDllError, &H400, strBuffer, Len(strBuffer), 0
  MSG = Replace(MSG, "|3", removeNull(strBuffer))
  MSG = Replace(MSG, "\n", vbNewLine)
  Err.Raise vbObjectError + errCreateFileError, App.EXEName & ".SaveSettingsToFile", MSG
  Exit Function

Fehler:
  ' unbekannter Fehler
  Err.Raise IIf(Err.Number, Err.Number, CLng(AdditionalData)), App.EXEName & ".SaveSettingsToFile"
  Exit Function

InvalidFormat:
  ' ungültiges Dateiformat
  On Error GoTo 0
  Err.Raise 321, App.EXEName & ".SaveSettingsToFile"
  Exit Function

ReadFileError:
  ' ReadFile ist fehlgeschlagen
  On Error GoTo 0
  MSG = LoadResString(resReadFileError)
  MSG = Replace(MSG, "|1", File)
  MSG = Replace(MSG, "|2", Err.LastDllError)
  strBuffer = Space(500)
  FormatMessage FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, ByVal 0&, Err.LastDllError, 1024, strBuffer, Len(strBuffer), 0
  MSG = Replace(MSG, "|3", removeNull(strBuffer))
  MSG = Replace(MSG, "\n", vbNewLine)
  Err.Raise vbObjectError + errReadFileError, App.EXEName & ".SaveSettingsToFile", MSG
  Exit Function

SettingsAlreadyExists:
  ' Die Datei enthält bereits gültige Daten
  On Error GoTo 0
  MSG = LoadResString(resSettingsAlreadyExists)
  MSG = Replace(MSG, "|1", File)
  MSG = Replace(MSG, "|2", AppName)
  MSG = Replace(MSG, "\n", vbNewLine)
  Err.Raise vbObjectError + errSettingsAlreadyExists, App.EXEName & ".SaveSettingsToFile", MSG
  Exit Function

TooNewFileFormat_Read:
  ' das Dateiformat ist zu neu zum Lesen
  On Error GoTo 0
  MSG = LoadResString(resTooNewFileFormat_Read)
  MSG = Replace(MSG, "|1", File)
  MSG = Replace(MSG, "|2", AppName)
  MSG = Replace(MSG, "|3", CStr(AdditionalData))
  MSG = Replace(MSG, "|4", Version)
  MSG = Replace(MSG, "\n", vbNewLine)
  Err.Raise vbObjectError + errTooNewFileFormat_Read, App.EXEName & ".SaveSettingsToFile", MSG
  Exit Function

TooNewFileFormat_Write:
  ' das Dateiformat ist zu neu zum Schreiben
  On Error GoTo 0
  MSG = LoadResString(resTooNewFileFormat_Write)
  MSG = Replace(MSG, "|1", File)
  MSG = Replace(MSG, "|2", AppName)
  MSG = Replace(MSG, "|3", Version_Major & "." & Version_Minor & "." & Version_Revision & " (Build " & neededVersionToWrite_Build & ")")
  MSG = Replace(MSG, "|4", Version)
  MSG = Replace(MSG, "\n", vbNewLine)
  Err.Raise vbObjectError + errTooNewFileFormat_Write, App.EXEName & ".SaveSettingsToFile", MSG
  Exit Function

WriteFileError:
  ' WriteFile ist fehlgeschlagen
  On Error GoTo 0
  MSG = LoadResString(resWriteFileError)
  MSG = Replace(MSG, "|1", File)
  MSG = Replace(MSG, "|2", Err.LastDllError)
  strBuffer = Space(500)
  FormatMessage FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, ByVal 0&, Err.LastDllError, 1024, strBuffer, Len(strBuffer), 0
  MSG = Replace(MSG, "|3", removeNull(strBuffer))
  MSG = Replace(MSG, "\n", vbNewLine)
  Err.Raise vbObjectError + errWriteFileError, App.EXEName & ".SaveSettingsToFile", MSG
End Function

Public Sub SetFocus()
Attribute SetFocus.VB_Description = "Setzt den Fokus auf das Control."
  UserControl.SetFocus
End Sub


' lokale Methoden

' fügt <pIDLToParent> ggf. als SubItem von <hParentItem> ein
Private Function AddSubItem(ByVal hParentItem As Long, pIDLToParent As Long, Optional ByVal freepIDL As Boolean = False, Optional ByVal forceInsertion As Boolean = False, Optional ByVal pIDLIsOkay As Boolean = False, Optional ByVal skipExistanceCheck As Boolean = True) As Long
  Dim DispName As String
  Dim IParent As IVBShellFolder
  Dim IShIconOverlay As IVBShellIconOverlay
  Dim itemAttr As SFGAOConstants
  Dim overlayIndex As Long
  Dim pIDLParent_ToDesktop As Long
  Dim pIDLToDesktop As Long
  Dim pIDLToParent2 As Long
  Dim ret As Long
  Dim shouldShowIt As Boolean
  Dim TVWItemInfo As ITEMINFO_COMMON
  Dim TVWItemInfo_Parent As ITEMINFO_COMMON
  Dim useStandardIcon As Boolean

  If pIDLToParent = 0 Then
    AddSubItem = -1
    Exit Function
  End If
  If Not GetItemInfo_Common(hParentItem, TVWItemInfo_Parent, skipExistanceCheck) Then
    ret = -1
    GoTo FreeMem
  End If

  pIDLParent_ToDesktop = GetFullyQualifiedPIDL(hParentItem, skipExistanceCheck)
  #If Debuging Then
    debugger.AddLogEntry "AddSubItem: Calling GetISHFolderInterfaceFQ()", LogEntryTypeConstants.letInfo
    debugger.AddLogEntry "   0x" & Hex(pIDLParent_ToDesktop), LogEntryTypeConstants.letOther
    Set IParent = GetISHFolderInterfaceFQ(debugger, pIDLParent_ToDesktop)
  #Else
    Set IParent = GetISHFolderInterfaceFQ(pIDLParent_ToDesktop)
  #End If
  If pIDLIsOkay Then
    pIDLToParent2 = ILClone(pIDLToParent)
  Else
    ' da <pIDLToParent> in der Regel von einem AutoUpdate-Event stammt, enthält sie oft nicht
    ' alle Daten
    ' -> eine vollständige pIDL ermitteln
    #If Debuging Then
      pIDLToParent2 = FilterSubItems_pIDL(debugger, IParent, pIDLParent_ToDesktop, pIDLToParent, True, False)
    #Else
      pIDLToParent2 = FilterSubItems_pIDL(IParent, pIDLParent_ToDesktop, pIDLToParent, True, False)
    #End If
  End If
  If pIDLToParent2 Then
    pIDLToDesktop = ILAppendID(ILClone(pIDLParent_ToDesktop), pIDLToParent2, 1)
    If forceInsertion Then
      shouldShowIt = True
    Else
      #If Debuging Then
        shouldShowIt = ShouldShowItem(debugger, Me, IParent, pIDLToParent2, pIDLToDesktop, hParentItem)
      #Else
        shouldShowIt = ShouldShowItem(Me, IParent, pIDLToParent2, pIDLToDesktop, hParentItem)
      #End If
    End If
  End If

  If shouldShowIt Then
    ' Item anzeigen
    TVWItemInfo.pIDLToDesktop = pIDLToDesktop

    ' Typ des Items ermitteln
    #If Debuging Then
      itemAttr = GetAttributes(debugger, IParent, pIDLToParent2, SFGAOConstants.SFGAO_FILESYSTEM Or SFGAOConstants.SFGAO_FOLDER Or SFGAOConstants.SFGAO_SHARE Or SFGAOConstants.SFGAO_LINK Or SFGAOConstants.SFGAO_GHOSTED Or SFGAOConstants.SFGAO_NEWCONTENT)
    #Else
      itemAttr = GetAttributes(IParent, pIDLToParent2, SFGAOConstants.SFGAO_FILESYSTEM Or SFGAOConstants.SFGAO_FOLDER Or SFGAOConstants.SFGAO_SHARE Or SFGAOConstants.SFGAO_LINK Or SFGAOConstants.SFGAO_GHOSTED Or SFGAOConstants.SFGAO_NEWCONTENT)
    #End If
    If itemAttr And SFGAOConstants.SFGAO_FILESYSTEM Then
      If itemAttr And SFGAOConstants.SFGAO_FOLDER Then
        ' ist es wirklich ein Ordner?
        #If Debuging Then
          If FileExists_pIDL(debugger, IParent, pIDLToParent2) Then
        #Else
          If FileExists_pIDL(IParent, pIDLToParent2) Then
        #End If
          TVWItemInfo.itemType = CByte(ItemTypeConstants.itFSFile)
        Else
          TVWItemInfo.itemType = CByte(ItemTypeConstants.itFSFolder)
        End If
      Else
        TVWItemInfo.itemType = CByte(ItemTypeConstants.itFSFile)
      End If
    Else
      If itemAttr And SFGAOConstants.SFGAO_FOLDER Then
        TVWItemInfo.itemType = CByte(ItemTypeConstants.itNonFSFolder)
      Else
        TVWItemInfo.itemType = CByte(ItemTypeConstants.itNonFSFile)
      End If
    End If

    If (prophImageList <> 0) And Not propUseCustomIcons Then
      ' Overlays laden
      overlayIndex = 0
      If propShownOverlays <> ShownOverlaysConstants.soNone Then
        On Error Resume Next
        IParent.QueryInterface IID_IShellIconOverlay, IShIconOverlay
        If IShIconOverlay Is Nothing Then
          If itemAttr And SFGAOConstants.SFGAO_SHARE Then overlayIndex = 1
          If itemAttr And SFGAOConstants.SFGAO_LINK Then overlayIndex = 2
        Else
          If IShIconOverlay.GetOverlayIndex(pIDLToParent2, overlayIndex) = S_OK Then
            ' jetzt haben wir den Overlay-Index für den Item, aber soll dieser Overlay überhaupt
            ' angezeigt werden?
            If ShouldShowOverlay(propShownOverlays, overlayIndex) Then
              ' ja!
            Else
              ' kein Overlay
              overlayIndex = 0
            End If
          Else
            ' kein Overlay
            overlayIndex = 0
          End If
          Set IShIconOverlay = Nothing
        End If
      End If
    End If

    #If Debuging Then
      DispName = pIDLToDisplayName(debugger, IParent, pIDLToParent2, SHGDNConstants.SHGDN_INFOLDER)
      If (prophImageList <> 0) And Not propUseCustomIcons Then
        #If NoCallbacks Then
          useStandardIcon = propUseStandardIcons
          If propUseStandardIconsForSlowItems And Not useStandardIcon Then
            useStandardIcon = IsSlowItem(debugger, IParent, pIDLToParent2, TVWItemInfo.pIDLToDesktop, True)
          End If
        #End If
      End If
    #Else
      DispName = pIDLToDisplayName(IParent, pIDLToParent2, SHGDNConstants.SHGDN_INFOLDER)
      If (prophImageList <> 0) And Not propUseCustomIcons Then
        #If NoCallbacks Then
          useStandardIcon = propUseStandardIcons
          If propUseStandardIconsForSlowItems And Not useStandardIcon Then
            useStandardIcon = IsSlowItem(IParent, pIDLToParent2, TVWItemInfo.pIDLToDesktop, True)
          End If
        #End If
      End If
    #End If

    ' Item hinzufügen
    ret = InsertTVWItem_Common(DispName, hParentItem, TVWItemInfo, overlayIndex, itemAttr And SFGAOConstants.SFGAO_GHOSTED, itemAttr And SFGAOConstants.SFGAO_NEWCONTENT, useStandardIcon)
  Else
    #If Debuging Then
      FreeItemIDList debugger, "AddSubItem #1", pIDLToDesktop
    #Else
      FreeItemIDList pIDLToDesktop
    #End If
  End If
  #If Debuging Then
    FreeItemIDList debugger, "AddSubItem #2", pIDLToParent2
  #Else
    FreeItemIDList pIDLToParent2
  #End If
  Set IParent = Nothing

FreeMem:
  If freepIDL Then
    #If Debuging Then
      FreeItemIDList debugger, "AddSubItem #3", pIDLToParent
    #Else
      FreeItemIDList pIDLToParent
    #End If
  End If
  If ret = 0 Then ret = -1
  AddSubItem = ret
End Function

Private Function GetFullyQualifiedPIDL(ByVal hItem As Long, Optional ByVal skipExistanceCheck As Boolean = False) As Long
  Dim TVWItemInfo As ITEMINFO_COMMON

  If Not GetItemInfo_Common(hItem, TVWItemInfo, skipExistanceCheck) Then Exit Function
  GetFullyQualifiedPIDL = TVWItemInfo.pIDLToDesktop
End Function

' überträgt die Schriftart auf das TreeView
Private Sub ChangeFont()
  Const WM_GETFONT = &H31
  Const WM_SETFONT = &H30
  Dim hFont As Long

  ' einfach die Schriftart des UserControls ändern, diese dann ermitteln (als Handle) und auf
  ' das Control übertragen
  On Error Resume Next
  Set UserControl.Font = propFont
  hFont = SendMessageAsLong(UserControl.hWnd, WM_GETFONT, 0, 0)
  If prophWnd Then SendMessageAsLong prophWnd, WM_SETFONT, hFont, 1
End Sub

' löscht alle Einträge aus dem TreeView
Private Sub ClearTVWItems()
  #If Debuging Then
    debugger.AddLogEntry "Entered ClearTVWItems()", LogEntryTypeConstants.letInfo
  #End If

  If prophWnd Then
    dontSavePath = True
    SendMessageAsLong prophWnd, TVM_DELETEITEM, 0, TVI_ROOT
    dontSavePath = False

    FreeItemMem 0
    cachedISHIcon_hItem = 0
  End If

  #If Debuging Then
    debugger.AddLogEntry "Leaving ClearTVWItems()", LogEntryTypeConstants.letInfo
  #End If
End Sub

Private Function CountAllSubItems(ByVal hParentItem As Long, Optional ByVal StateIconMask As Long = -1, Optional ByVal MustMatch As Boolean = True) As Long
  Dim hItem As Long
  Dim ret As Long

  If Not ItemExists(hParentItem, AllowNullConstants.anIfNoShowRoot) Then Exit Function

  hItem = ItemGetFirstSubItem(hParentItem)
  While ItemExists(hItem, AllowNullConstants.anNever)
    If StateIconMask = -1 Then
      ret = ret + 1
    Else
      If ItemStateIconIndex(hItem) = StateIconMask Then
        If MustMatch Then ret = ret + 1
      Else
        If Not MustMatch Then ret = ret + 1
      End If
    End If
    If GetItemChildrenProperty(hItem) Then ret = ret + CountAllSubItems(hItem, StateIconMask, MustMatch)

    hItem = ItemGetNextItem(hItem)
  Wend

  CountAllSubItems = ret
End Function

' erzeugt das TreeView
Private Sub CreateControl()
  Const WM_SIZE = &H5
  Dim ExStyle As ExWindowStyleConsts
  Dim Style As WindowStyleConsts
  Dim TVWExStyle As TreeViewExStyleConsts
  Dim TVWStyle As TreeViewStyleConsts

  If disableCreateControl Then
    needRecreate = True
    Exit Sub
  Else
    needRecreate = False
  End If

  ' zur Sicherheit...
  DestroyControl

  ' Styles setzen
  ExStyle = ExWindowStyleConsts.WS_EX_LEFT Or ExWindowStyleConsts.WS_EX_LTRREADING Or ExWindowStyleConsts.WS_EX_RIGHTSCROLLBAR
  If propAppearance = AppearanceConstants.a3D Then
    ExStyle = ExStyle Or ExWindowStyleConsts.WS_EX_CLIENTEDGE
  ElseIf propAppearance = AppearanceConstants.a3DLight Then
    ExStyle = ExStyle Or ExWindowStyleConsts.WS_EX_STATICEDGE
  End If
  If propRTLReading Then
    ExStyle = ExStyle Or ExWindowStyleConsts.WS_EX_LAYOUTRTL
  End If

  Style = WindowStyleConsts.WS_CHILDWINDOW Or WindowStyleConsts.WS_CLIPCHILDREN Or WindowStyleConsts.WS_CLIPSIBLINGS Or WindowStyleConsts.WS_TABSTOP Or WindowStyleConsts.WS_VISIBLE
  If propBorderStyle = BorderStyleConstants.bsFixedSingle Then Style = Style Or WindowStyleConsts.WS_BORDER

  TVWStyle = TreeViewStyleConsts.TVS_NONEVENHEIGHT
  If Not propAllowDragDrop Then TVWStyle = TVWStyle Or TreeViewStyleConsts.TVS_DISABLEDRAGDROP
  If propCheckBoxes Then TVWStyle = TVWStyle Or TreeViewStyleConsts.TVS_CHECKBOXES
  If propFullRowSelect Then TVWStyle = TVWStyle Or TreeViewStyleConsts.TVS_FULLROWSELECT
  If propHotTracking Then TVWStyle = TVWStyle Or TreeViewStyleConsts.TVS_TRACKSELECT
  If propLineStyle = LineStyleConstants.lsLinesAtRoot Then TVWStyle = TVWStyle Or TreeViewStyleConsts.TVS_LINESATROOT
  If propRenaming = RenamingConstants.rAutomatic Then TVWStyle = TVWStyle Or TreeViewStyleConsts.TVS_EDITLABELS
  If propRTLReading Then TVWStyle = TVWStyle Or TreeViewStyleConsts.TVS_RTLREADING
  If propScrollBars = ScrollBarsConstants.sbNone Then TVWStyle = TVWStyle Or TreeViewStyleConsts.TVS_NOSCROLL
  If propShownToolTips = ShownToolTipsConstants.sttNone Then TVWStyle = TVWStyle Or TreeViewStyleConsts.TVS_NOTOOLTIPS Else TVWStyle = TVWStyle Or TreeViewStyleConsts.TVS_INFOTIP
  If propShowSelection Then TVWStyle = TVWStyle Or TreeViewStyleConsts.TVS_SHOWSELALWAYS
  If propSingleExpand <> SingleExpandConstants.seNone Then TVWStyle = TVWStyle Or TreeViewStyleConsts.TVS_SINGLEEXPAND
  If propTreeViewStyle And TreeViewStyleConstants.tvsButtons Then TVWStyle = TVWStyle Or TreeViewStyleConsts.TVS_HASBUTTONS
  If propTreeViewStyle And TreeViewStyleConstants.tvsLines Then TVWStyle = TVWStyle Or TreeViewStyleConsts.TVS_HASLINES

  ' Fenster erzeugen...
  prophWnd = CreateWindowExAsLong(ExStyle, WC_TREEVIEW, Ambient.DisplayName, Style Or TVWStyle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, 0, App.hInstance, 0)
  If prophWnd = 0 Then GoTo Fehler
  prophDC = GetDC(prophWnd)

  ' für C++-Kompatibilität
  ' ToDo: Reicht das, oder ist HWND_TOPMOST wirklich nötig?
  SetWindowPos hWndToolTip, prophWnd, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

  #If UseSubClassing Then
    If Not inIDE Then SubClass
  #End If

  If ver_Comctl32_610 Then
    ' this eliminates flicker
    TVWExStyle = TreeViewExStyleConsts.TVS_EX_DOUBLEBUFFER 'Or TVS_EX_DRAWIMAGEASYNC
    If propAutoHScroll Then TVWExStyle = TVWExStyle Or TreeViewExStyleConsts.TVS_EX_AUTOHSCROLL
    If propFadeExpandos Then TVWExStyle = TVWExStyle Or TreeViewExStyleConsts.TVS_EX_FADEINOUTEXPANDOS
    SendMessageAsLong prophWnd, TVM_SETEXTENDEDSTYLE, 0, TVWExStyle
  End If

  ' ...weitere Props setzen...
  EnableWindow prophWnd, propEnabled
  SendMessageAsLong prophWnd, TVM_SETBORDER, 1, propMargin
  SendMessageAsLong prophWnd, TVM_SETINDENT, propIndent, 0
  If propItemHeight = 0 Then propItemHeight = 16
  SendMessageAsLong prophWnd, TVM_SETITEMHEIGHT, propItemHeight, 0
  SendMessageAsLong prophWnd, TVM_SETSCROLLTIME, propMaxScrollTime, 0

  ' ...Farben setzen...
  SetColors

  ' Schriftart setzen
  ChangeFont

  ' RECTs initialisieren
  WinMainTree 1, 0, prophWnd, WM_SIZE, 0, 0

  If propAllowDragDrop Then
    If propDragShowAlphaBlendedItems Then
      CoCreateInstance IID_IDragDropHelper, Nothing, CLSCTXConsts.CLSCTX_INPROC_SERVER, IID_IDropTargetHelper, IDropTargetHelper
    End If
    UserControl.OLEDropMode = OLEDropConstants.vbOLEDropManual
  End If

  If inIDE Then
    LoadIDEItems
  Else
    #If UseSubClassing Then
      ' das Dateisystem wird versetzt eingelesen (damit das Control erstmal angezeigt wird)
    #Else
      disableLoadFileSystem = False
      InitTreeView
      SelectPath propPath
    #End If
  End If
  Exit Sub

Fehler:
End Sub

' zerstört das TreeView
Private Sub DestroyControl()
  UnSubClass

  Set IDropTargetHelper = Nothing

  If prophWnd Then
    KillTimer prophWnd, TimerID_SubClassParent

    ReleaseDC prophWnd, prophDC
    prophDC = 0

    ClearTVWItems

    ShowWindow prophWnd, SW_HIDE
    SetParent prophWnd, 0
    DestroyWindow prophWnd
    prophWnd = 0
    If prophLargeStateImageList Then ImageList_Destroy prophLargeStateImageList
    prophLargeStateImageList = 0
    If prophSmallStateImageList Then ImageList_Destroy prophSmallStateImageList
    prophSmallStateImageList = 0
  End If
End Sub

' gibt den Speicher frei, der für die "ItemInfo"-Struktur von <hItem> reserviert wurde
Private Sub FreeItemMem(ByVal hItem As Long, Optional ByVal freepIDLs As Boolean = True, Optional ByVal freelParam As Boolean = True)
  Dim isRootItem As Boolean
  Dim pTVWItemInfo As Long
  Dim TVWItemInfo As ITEMINFO_COMMON

  'If inIDE Then Exit Sub
  'If Not ItemExists(hItem, AllowNullConstants.anIfNoShowRoot) Then Exit Sub

  If ItemIsPartOfArchive(hItem) Then
    ' ein Archiv-Item
    If freelParam Then
      pTVWItemInfo = GetItemlParam(hItem)
      CoTaskMemFree pTVWItemInfo
    End If
  Else
    ' ein "normaler" Item
    If hItem = 0 Then
      ' der RootItem
      If freepIDLs Then
        If RootItemInfo.pIDLToDesktop Then
          #If Debuging Then
            FreeItemIDList debugger, "FreeItemMem #1", RootItemInfo.pIDLToDesktop
          #Else
            FreeItemIDList RootItemInfo.pIDLToDesktop
          #End If
        End If
      End If
    Else
      pTVWItemInfo = GetItemlParam(hItem)
      CopyMemory VarPtr(TVWItemInfo), pTVWItemInfo, LenB(TVWItemInfo)

      If freepIDLs Then
        isRootItem = (TVWItemInfo.pIDLToDesktop = RootItemInfo.pIDLToDesktop)
        #If Debuging Then
          FreeItemIDList debugger, "FreeItemMem #2", TVWItemInfo.pIDLToDesktop
        #Else
          FreeItemIDList TVWItemInfo.pIDLToDesktop
        #End If
        If isRootItem Then RootItemInfo.pIDLToDesktop = 0
      End If

      If freelParam Then CoTaskMemFree pTVWItemInfo
    End If
  End If
End Sub

#If BkImageSupport Then
  Private Function GetBackgroundBitmap(ByRef hCompatibleDC As Long) As Long
    Const SB_HORZ = 0
    Const SIF_PAGE = &H2
    Const SIF_POS = &H4
    Const SIF_RANGE = &H1
    Const SIF_TRACKPOS = &H10
    Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)
    Dim anchorPoint As POINT
    Dim bitmapSize As Size
    Dim bmp As BITMAP
    Dim clientRectangle As RECT
    Dim cx As Long
    Dim cy As Long
    Dim hBMP_Target As Long
    Dim hBrush_Background As Long
    Dim hDC_Bitmap As Long
    Dim hDC_Target As Long
    Dim horizontalScrollInfo As SCROLLINFO
    Dim hPreviousBitmap1 As Long
    Dim hPreviousBitmap2 As Long
    Dim rc As RECT
    Dim Style As WindowStyleConsts
    Dim xDestination As Long
    Dim xSource As Long
    Dim yDestination As Long
    Dim ySource As Long

    GetClientRect prophWnd, clientRectangle

    GetObjectAPI propBkImage.Handle, LenB(bmp), bmp
    bitmapSize.cx = bmp.bmWidth
    bitmapSize.cy = bmp.bmHeight

    hDC_Target = CreateCompatibleDC(hCompatibleDC)
    hBMP_Target = CreateCompatibleBitmap(hCompatibleDC, clientRectangle.Right - clientRectangle.Left, clientRectangle.Bottom - clientRectangle.Top)
    hPreviousBitmap1 = SelectObject(hDC_Target, hBMP_Target)

    ' calculate the anchor point
    If Not propBkImageFixed Then
      ' TODO: Find out whether this style is reliable on all versions of Windows.
      Style = GetWindowLongPtr(prophWnd, GWL_STYLE)
      If Style And WindowStyleConsts.WS_HSCROLL Then
        horizontalScrollInfo.cbSize = LenB(horizontalScrollInfo)
        horizontalScrollInfo.fMask = SIF_ALL
        GetScrollInfo prophWnd, SB_HORZ, horizontalScrollInfo
      Else
        horizontalScrollInfo.nMax = clientRectangle.Right
      End If

      anchorPoint.x = clientRectangle.Left - horizontalScrollInfo.nPos
      rc.Left = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_ROOT, 0)
      If SendMessage(prophWnd, TVM_GETITEMRECT, 1, rc) Then
        anchorPoint.y = rc.Top
      End If
    Else
      anchorPoint.x = clientRectangle.Left
      anchorPoint.y = clientRectangle.Top
    End If

    ' draw

    hDC_Bitmap = CreateCompatibleDC(hDC_Target)
    hPreviousBitmap2 = SelectObject(hDC_Bitmap, propBkImage.Handle)

    hBrush_Background = CreateSolidBrush(propBackColor)
    FillRect hDC_Target, clientRectangle, hBrush_Background

    yDestination = anchorPoint.y
    Do
      ySource = ySource Mod bitmapSize.cy
      If bitmapSize.cy - ySource < clientRectangle.Bottom - yDestination Then
        cy = bitmapSize.cy - ySource
      Else
        cy = clientRectangle.Bottom - yDestination
      End If

      xDestination = anchorPoint.x
      xSource = 0
      Do
        xSource = xSource Mod bitmapSize.cx
        If bitmapSize.cx - xSource < clientRectangle.Right - xDestination Then
          cx = bitmapSize.cx - xSource
        Else
          cx = clientRectangle.Right - xDestination
        End If

        BitBlt hDC_Target, xDestination, yDestination, cx, cy, hDC_Bitmap, xSource, ySource, SRCCOPY

        xSource = xSource + cx
        xDestination = xDestination + cx
      Loop While xDestination < clientRectangle.Right

      ySource = ySource + cy
      yDestination = yDestination + cy
    Loop While yDestination < clientRectangle.Bottom

    DeleteObject hBrush_Background
    SelectObject hDC_Bitmap, hPreviousBitmap2
    DeleteDC hDC_Bitmap
    SelectObject hDC_Target, hPreviousBitmap1
    DeleteDC hDC_Target

    GetBackgroundBitmap = hBMP_Target
  End Function
#End If

Private Function GetInternetExplorerItemHandle(ByVal insertIfNecessary As Boolean) As Long
  Dim DispName As String
  Dim hFirstSubItem As Long
  Dim hParentItem As Long
  Dim IItem As IVBShellFolder
  Dim IShIconOverlay As IVBShellIconOverlay
  Dim itemAttr As SFGAOConstants
  Dim overlayIndex As Long
  Dim pIDLIE_ToDesktop As Long
  Dim pIDLToParent As Long
  Dim ret As Long
  Dim TVWItemInfo As ITEMINFO_COMMON
  Dim useStandardIcon As Boolean

  ret = -1
  IDesktop.ParseDisplayName 0, 0, "::{871C5380-42A0-1069-A2EA-08002B30309D}", Len("::{871C5380-42A0-1069-A2EA-08002B30309D}"), pIDLIE_ToDesktop, 0
  If pIDLIE_ToDesktop Then
    If ILIsEqual(RootItemInfo.pIDLToDesktop, pIDL_Desktop) Then
      ' den IE suchen (zu Gunsten der Performance nur die direkten SubItems des Desktop durchsuchen)
      If propShowRoot Then
        hFirstSubItem = ItemGetFirstSubItem(GetRootOrFirstItem)
      Else
        hFirstSubItem = GetRootOrFirstItem
      End If
      #If Debuging Then
        debugger.AddLogEntry "GetInternetExplorerItemHandle: Calling InternalpIDLToItemHandle()", letInfo
        debugger.AddLogEntry "   0x" & Hex(pIDLIE_ToDesktop), LogEntryTypeConstants.letOther
      #End If
      ret = InternalpIDLToItemHandle(pIDLIE_ToDesktop, hFirstSubItem, , , False)
      If ret = -1 Then
        If insertIfNecessary Then
          ' IE einfügen
          If propShowRoot Then
            hParentItem = GetRootOrFirstItem
          Else
            hParentItem = 0
          End If
          TVWItemInfo.pIDLToDesktop = pIDLIE_ToDesktop
          pIDLToParent = ILFindLastID(pIDLIE_ToDesktop)
          #If Debuging Then
            Set IItem = GetParentInterface(debugger, pIDLIE_ToDesktop)

            ' Typ des Items ermitteln
            itemAttr = GetAttributes(debugger, IItem, pIDLToParent, SFGAOConstants.SFGAO_FILESYSTEM Or SFGAOConstants.SFGAO_FOLDER Or SFGAOConstants.SFGAO_SHARE Or SFGAOConstants.SFGAO_LINK Or SFGAOConstants.SFGAO_GHOSTED Or SFGAOConstants.SFGAO_NEWCONTENT)
          #Else
            Set IItem = GetParentInterface(pIDLIE_ToDesktop)

            ' Typ des Items ermitteln
            itemAttr = GetAttributes(IItem, pIDLToParent, SFGAOConstants.SFGAO_FILESYSTEM Or SFGAOConstants.SFGAO_FOLDER Or SFGAOConstants.SFGAO_SHARE Or SFGAOConstants.SFGAO_LINK Or SFGAOConstants.SFGAO_GHOSTED Or SFGAOConstants.SFGAO_NEWCONTENT)
          #End If
          If itemAttr And SFGAOConstants.SFGAO_FILESYSTEM Then
            If itemAttr And SFGAOConstants.SFGAO_FOLDER Then
              ' ist es wirklich ein Ordner?
              #If Debuging Then
                If FileExists_pIDL(debugger, IItem, pIDLToParent) Then
              #Else
                If FileExists_pIDL(IItem, pIDLToParent) Then
              #End If
                TVWItemInfo.itemType = CByte(ItemTypeConstants.itFSFile)
              Else
                TVWItemInfo.itemType = CByte(ItemTypeConstants.itFSFolder)
              End If
            Else
              TVWItemInfo.itemType = CByte(ItemTypeConstants.itFSFile)
            End If
          Else
            If itemAttr And SFGAOConstants.SFGAO_FOLDER Then
              TVWItemInfo.itemType = CByte(ItemTypeConstants.itNonFSFolder)
            Else
              TVWItemInfo.itemType = CByte(ItemTypeConstants.itNonFSFile)
            End If
          End If
          ' Overlays laden
          overlayIndex = 0
          If propShownOverlays <> ShownOverlaysConstants.soNone Then
            On Error Resume Next
            IItem.QueryInterface IID_IShellIconOverlay, IShIconOverlay
            If IShIconOverlay Is Nothing Then
              If itemAttr And SFGAOConstants.SFGAO_SHARE Then overlayIndex = 1
              If itemAttr And SFGAOConstants.SFGAO_LINK Then overlayIndex = 2
            Else
              If IShIconOverlay.GetOverlayIndex(pIDLToParent, overlayIndex) = S_OK Then
                ' jetzt haben wir den Overlay-Index für den Item, aber soll dieser Overlay
                ' überhaupt angezeigt werden?
                If ShouldShowOverlay(propShownOverlays, overlayIndex) Then
                  ' ja!
                Else
                  ' kein Overlay
                  overlayIndex = 0
                End If
              Else
                ' kein Overlay
                overlayIndex = 0
              End If
              Set IShIconOverlay = Nothing
            End If
          End If

          #If Debuging Then
            DispName = pIDLToDisplayName(debugger, IItem, pIDLToParent, SHGDNConstants.SHGDN_INFOLDER)
            If (prophImageList <> 0) And Not propUseCustomIcons Then
              #If NoCallbacks Then
                useStandardIcon = propUseStandardIcons
                If propUseStandardIconsForSlowItems And Not useStandardIcon Then
                  useStandardIcon = IsSlowItem(debugger, IItem, pIDLToParent, TVWItemInfo.pIDLToDesktop, True)
                End If
              #End If
            End If
          #Else
            DispName = pIDLToDisplayName(IItem, pIDLToParent, SHGDNConstants.SHGDN_INFOLDER)
            If (prophImageList <> 0) And Not propUseCustomIcons Then
              #If NoCallbacks Then
                useStandardIcon = propUseStandardIcons
                If propUseStandardIconsForSlowItems And Not useStandardIcon Then
                  useStandardIcon = IsSlowItem(IItem, pIDLToParent, TVWItemInfo.pIDLToDesktop, True)
                End If
              #End If
            End If
          #End If

          ' Item hinzufügen
          ret = InsertTVWItem_Common(DispName, hParentItem, TVWItemInfo, overlayIndex, itemAttr And SFGAOConstants.SFGAO_GHOSTED, itemAttr And SFGAOConstants.SFGAO_NEWCONTENT, useStandardIcon)
          SortItems hParentItem
          Set IItem = Nothing
        End If
      Else
        If ILIsEqual(RootItemInfo.pIDLToDesktop, pIDLIE_ToDesktop) Then
          ' der Root-Item ist der IE
          If propShowRoot Then
            ret = GetRootOrFirstItem
          Else
            ret = 0
          End If
        End If
      End If

      If TVWItemInfo.pIDLToDesktop <> pIDLIE_ToDesktop Then
        #If Debuging Then
          FreeItemIDList debugger, "GetInternetExplorerItemHandle", pIDLIE_ToDesktop
        #Else
          FreeItemIDList pIDLIE_ToDesktop
        #End If
      End If
    End If
  End If

  GetInternetExplorerItemHandle = ret
  #If Debuging Then
    debugger.AddLogEntry "Leaving GetInternetExplorerItemHandle (ret=0x" & Hex(ret) & ")", letInfo
  #End If
End Function

Private Function GetItemChildrenProperty(ByVal hItem As Long) As Boolean
  Dim itemData As TVITEMEX

  If Not ItemExists(hItem, AllowNullConstants.anIfNoShowRoot) Then Exit Function

  If hItem Then
    With itemData
      .hItem = hItem
      .Mask = TVIF_HANDLE Or TVIF_CHILDREN
      SendMessage prophWnd, TVM_GETITEM, 0, itemData

      GetItemChildrenProperty = .cChildren
    End With
  Else
    GetItemChildrenProperty = (GetRootOrFirstItem <> -1)
  End If
End Function

' gibt den IconIndex für <hItem> zurück
Private Function GetItemIconIndex(ByVal hItem As Long, ByVal Selected As Boolean) As Long
  Dim itemData As TVITEMEX

  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Function

  With itemData
    .hItem = hItem
    .Mask = TVIF_HANDLE Or TVIF_IMAGE Or TVIF_SELECTEDIMAGE
    SendMessage prophWnd, TVM_GETITEM, 0, itemData
    GetItemIconIndex = IIf(Selected, .iSelectedImage, .iImage)
  End With
End Function

' gibt die ITEMINFO_ARCHIVE-Struktur für <hItem> zurück
' gibt False zurück, wenn <hItem> ein "normaler" Item ist
Private Function GetItemInfo_Archive(ByVal hItem As Long, TVWItemInfo As ITEMINFO_ARCHIVE) As Boolean
  Dim pItemInfo As Long

  If inIDE Then Exit Function
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Function

  If ItemIsPartOfArchive(hItem) Then
    pItemInfo = GetItemlParam(hItem)
    If pItemInfo Then
      GetItemInfo_Archive = True
      CopyMemory VarPtr(TVWItemInfo), pItemInfo, LenB(TVWItemInfo)
    End If
  End If
End Function

' gibt die ITEMINFO_COMMON-Struktur für <hItem> zurück
' gibt False zurück, wenn <hItem> ein Archiv-Item ist
Private Function GetItemInfo_Common(ByVal hItem As Long, TVWItemInfo As ITEMINFO_COMMON, Optional ByVal skipExistanceCheck As Boolean = False) As Boolean
  Dim itmType As Byte
  Dim pItemInfo As Long

  If inIDE Then Exit Function
  If Not skipExistanceCheck Then
    If Not ItemExists(hItem, AllowNullConstants.anIfNoShowRoot) Then Exit Function
  End If

  If hItem Then
    pItemInfo = GetItemlParam(hItem)
    If pItemInfo Then
      CopyMemory VarPtr(itmType), pItemInfo, LenB(itmType)
    End If
  Else
    itmType = RootItemInfo.itemType
  End If
  If (itmType = ItemTypeConstants.itArchiveFile) Or (itmType = ItemTypeConstants.itArchiveFolder) Then
    GetItemInfo_Common = False
  Else
    GetItemInfo_Common = True
    If hItem Then
      If pItemInfo Then
        CopyMemory VarPtr(TVWItemInfo), pItemInfo, LenB(TVWItemInfo)
      End If
    Else
      ' der Root-Item
      If Not propShowRoot Then
        TVWItemInfo = RootItemInfo
      End If
    End If
  End If
End Function

' gibt den <.lParam>-Wert von <hItem> zurück
Private Function GetItemlParam(ByVal hItem As Long) As Long
  Dim itemData As TVITEMEX

  If (hItem <> -1) And (hItem <> 0) Then
    With itemData
      .hItem = hItem
      .Mask = TVIF_HANDLE Or TVIF_PARAM Or TVIF_INTEGRAL
      SendMessage prophWnd, TVM_GETITEM, 0, itemData
      If .iIntegral Then GetItemlParam = .lParam
    End With
  End If
End Function

' gibt den Overlay von <hItem> zurück
Private Function GetItemOverlayIndex(ByVal hItem As Long) As Long
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Function
  GetItemOverlayIndex = OVERLAYMASKTOINDEX((GetItemState_Fast(hItem, TVIS_OVERLAYMASK) And TVIS_OVERLAYMASK))
End Function

' gibt das ItemRect von <hItem> zurück
Private Sub GetItemRect(ByVal hItem As Long, rcItem As RECT, Optional ByVal entireRow As Boolean = False)
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub

  rcItem.Left = hItem
  SendMessage prophWnd, TVM_GETITEMRECT, Abs(Not entireRow), rcItem
End Sub

' gibt den Status des Items <hItem> zurück
Private Function GetItemState(hItem As Long, StateMask As Long) As Long
  Dim itemData As TVITEMEX

  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Function

  If ver_Comctl32_580 Then
    GetItemState = SendMessageAsLong(prophWnd, TVM_GETITEMSTATE, hItem, StateMask)
  Else
    With itemData
      .hItem = hItem
      .StateMask = StateMask
      .Mask = TVIF_HANDLE Or TVIF_STATE
      SendMessage prophWnd, TVM_GETITEM, 0, itemData
      GetItemState = .state
    End With
  End If
End Function

' gibt den Status des Items <hItem> zurück
Private Function GetItemState_Fast(hItem As Long, StateMask As Long) As Long
  Dim itemData As TVITEMEX

  If ver_Comctl32_580 Then
    GetItemState_Fast = SendMessageAsLong(prophWnd, TVM_GETITEMSTATE, hItem, StateMask)
  Else
    With itemData
      .hItem = hItem
      .StateMask = StateMask
      .Mask = TVIF_HANDLE Or TVIF_STATE
      SendMessage prophWnd, TVM_GETITEM, 0, itemData
      GetItemState_Fast = .state
    End With
  End If
End Function

' gibt die Beschriftung des Items <hItem> zurück
Private Function GetItemText(ByVal hItem As Long) As String
  Dim IParent As IVBShellFolder
  Dim itemData As TVITEMEX
  Dim pIDLToParent As Long

  If Not ItemExists(hItem, AllowNullConstants.anIfNoShowRoot) Then Exit Function

  If hItem Then
    With itemData
      .hItem = hItem
      .pszText = String$(MAX_PATH, Chr$(0))
      .cchTextMax = Len(.pszText)
      .Mask = TVIF_HANDLE Or TVIF_TEXT
      SendMessage prophWnd, TVM_GETITEM, 0, itemData
      GetItemText = Left$(.pszText, lstrlenA(.pszText))
    End With
  Else
    ' der Root-Item
    #If Debuging Then
      SplitFullyQualifiedPIDL debugger, RootItemInfo.pIDLToDesktop, IParent, pIDLToParent
    #Else
      SplitFullyQualifiedPIDL RootItemInfo.pIDLToDesktop, IParent, pIDLToParent
    #End If
    If (pIDLToParent <> 0) And Not (IParent Is Nothing) Then
      #If Debuging Then
        GetItemText = pIDLToDisplayName(debugger, IParent, pIDLToParent, SHGDNConstants.SHGDN_INFOLDER)
      #Else
        GetItemText = pIDLToDisplayName(IParent, pIDLToParent, SHGDNConstants.SHGDN_INFOLDER)
      #End If
    End If
    Set IParent = Nothing
  End If
End Function

' gibt den Root-Item oder den 1. Item zurück
Private Function GetRootOrFirstItem() As Long
  Dim ret As Long

  ret = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_ROOT, 0)
  If ret = 0 Then ret = -1
  GetRootOrFirstItem = ret
End Function

Private Function GetSettingsFileFormatVersion(ByVal File As String, Optional ByRef neededVersionToWrite As Long = 0, Optional ByRef neededVersionToWrite_Build As Integer = 0) As Integer
  Dim bytBuffer As Byte
  Dim BytesRead As Long
  Dim Flags As ErrorTypeConstants
  Dim hFile As Long
  Dim isVersion20 As Boolean
  Dim pos As Integer
  Dim ret As Integer
  Dim ret2 As Boolean
  Dim ret3 As Long
  Dim strBuffer As String

  On Error GoTo Fehler

  ret2 = FileContainsValidSettings(File, Flags)
  If Not ret2 Then ret2 = (Flags = ErrorTypeConstants.etNoSettingsFound)

  If ret2 Then
    ' Datei öffnen
    hFile = CreateFileAsLong(File, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If hFile <> INVALID_HANDLE_VALUE Then
      ' Formatversion einlesen...
      SetFilePointer hFile, 7, 0, FILE_CURRENT
      strBuffer = String$(4, Chr$(0))
      ret3 = ReadFileAsLong(hFile, ByVal strBuffer, Len(strBuffer), BytesRead, 0)
      If ret3 = 0 Then GoTo Fehler
      strBuffer = Left$(strBuffer, lstrlenA(strBuffer))
      '...und prüfen
      pos = InStr(strBuffer, ".")
      If pos Then
        ' bei Version 1.0 waren für die Version nur 3 Bytes vorgesehen
        If CByte(Left$(strBuffer, pos - 1)) = 1 Then strBuffer = Left$(strBuffer, Len(strBuffer) - 1)
        bytBuffer = CByte(Mid$(strBuffer, pos + 1))
        ret = MakeWord(bytBuffer, CByte(Left$(strBuffer, pos - 1)))
        isVersion20 = (CByte(Left$(strBuffer, pos - 1)) >= 2)
      End If

      If isVersion20 Then
        ' zur Versionsangabe für das Schreiben springen
        SetFilePointer hFile, 52, 0, FILE_CURRENT
        ' Version dieses Controls einlesen, die benötigt wird um dieses Format zu schreiben...
        ret3 = ReadFileAsLong(hFile, ByVal VarPtr(neededVersionToWrite), LenB(neededVersionToWrite), BytesRead, 0)
        If ret3 = 0 Then GoTo Fehler
        ret3 = ReadFileAsLong(hFile, ByVal VarPtr(neededVersionToWrite_Build), LenB(neededVersionToWrite_Build), BytesRead, 0)
        If ret3 = 0 Then GoTo Fehler
      End If

      ' Datei schließen
      CloseHandle hFile
    End If
  End If

  GetSettingsFileFormatVersion = ret
  Exit Function

Fehler:
  CloseHandle hFile
End Function

Private Function GetToolTipText(ByVal hItem As Long, ByVal InfoTipStyle As InfoTipStyleConstants, ByRef ToolTipType As ToolTipTypeConstants) As String
  Dim ret As String
  Dim txt As String

  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Function

  ToolTipType = ToolTipTypeConstants.tttToolTip
  If propShownToolTips And ShownToolTipsConstants.sttHiddenItems Then
    ' ist der Item versteckt?
    If ItemTextIsCut(hItem) Then ret = GetItemText(hItem)
  End If

  If propShownToolTips And ShownToolTipsConstants.sttInfoTips Then
    ' InfoTip laden
    txt = ItemGetInfoTip(hItem, InfoTipStyle)
    If (txt <> "") And (ret <> "") Then ret = ret & vbNewLine
    ret = ret & txt
    ToolTipType = ToolTipTypeConstants.tttInfoTip
  End If

  GetToolTipText = ret
End Function

' behandelt das AutoUpdate-Event für das Hinzufügen eines Laufwerks
' <SHNotData.dwItem1> ist die pIDL (wenn das Laufwerk über den Gerätemanager hinzugefügt wurde, ist es
' irgendetwas anderes)
' <SHNotData.dwItem2> ist 0
' gibt True zurück, wenn <Refresh> aufgerufen werden muß
Private Function HandleAutoUpdate_AddedDrive(IParent As IVBShellFolder, pIDLToParent As Long) As Boolean
  Dim hParentItem As Long
  Dim pIDLToDesktop As Long

  pIDLToDesktop = CSIDLTopIDL(CSIDLConstants.CSIDL_DRIVES)
  #If Debuging Then
    debugger.AddLogEntry "HandleAutoUpdate_AddedDrive: Calling InternalpIDLToItemHandle()", letInfo
    debugger.AddLogEntry "   0x" & Hex(pIDLToDesktop), LogEntryTypeConstants.letOther
  #End If
  hParentItem = InternalpIDLToItemHandle(pIDLToDesktop, , , True)

  If hParentItem = -1 Then
    ' der ParentItem existiert nicht
    If GetRootOrFirstItem = -1 Then
      ' das Control ist leer
      ' -> möglicherweise ist das hinzugefügte Laufwerk der RootItem
      ' -> in diesem Fall das Control füllen
      pIDLToDesktop = UserPathTopIDL(propRootPath, True)
      If pIDLToDesktop Then
        InitTreeView
        #If Debuging Then
          FreeItemIDList debugger, "HandleAutoUpdate_AddedDrive", pIDLToDesktop
        #Else
          FreeItemIDList pIDLToDesktop
        #End If
      End If
    Else
      ' keine weiteren Aktionen nötig
    End If
  Else
    ' der ParentItem  existiert bereits
    ' -> wenn seine SubItems schon geladen sind, diese ggf. ergänzen
    ' -> hat der ParentItem keine SubItems, den Item ggf. ergänzen
    If GetItemChildrenProperty(hParentItem) And Not ItemHasSubItems(hParentItem) Then
      ' der Item hat SubItems, diese wurden aber noch nicht geladen
      ' -> keine weiteren Aktionen nötig
    Else
      ' Item ggf. hinzufügen
      InsertNewAndRemoveOldSubItems hParentItem, True, False
    End If
  End If

  HandleAutoUpdate_AddedDrive = False
End Function

' behandelt das AutoUpdate-Event für das Ändern einer Dateityp-Verknüpfung
' <pIDL> ist die pIDL der Dateiendung
'    -> DispName: "*.xxx"
'    -> Pfad    : "C:\*.xxx"
' gibt True zurück, wenn <Refresh> aufgerufen werden muß
' ToDo: Scheint nicht zu funktionieren (Win-Explorer aktualisiert sich auch nicht)
Private Function HandleAutoUpdate_ChangedFileTypeLink(pIDL As Long) As Boolean
  HandleAutoUpdate_ChangedFileTypeLink = HandleAutoUpdate_ChangedIcons(pIDL)
End Function

Private Function HandleAutoUpdate_ChangedIcon(ByVal oldIconIndex As Long, ByVal pExtraData As Long) As Boolean
  If (prophImageList = 0) Or propUseCustomIcons Then Exit Function

  If pExtraData Then
    If ver_Win_XP Then
      oldIconIndex = SHHandleUpdateImage(pExtraData)
      If oldIconIndex = -1 Then
        ' failure
        Exit Function
      End If
    Else
      ' SHHandleUpdateImage crashes on Win2k, so go safe and update all icons
      oldIconIndex = -1
    End If
  End If

  UpdateDefaultIconIndices
  InvalidateIcons 0, oldIconIndex, oldIconIndex = -1

  HandleAutoUpdate_ChangedIcon = True
End Function

' behandelt das AutoUpdate-Event für das Ändern eines Icons für einen Dateitypen oder einen NonFS-Item
' <pDetails> ist 0, die pIDL des Desktops oder eine Struktur der Form:
'   Type UIData
'     Size As Integer
'     OldIconIndex As Long
'     ? As Long
'   End Type
' <SHNotData.dwItem2> ist eine Struktur? der Form:
'   Type UIData
'     ? As Long?
'     ? As Integer?
'     OldIconIndex As Long
'   End Type
' gibt True zurück, wenn <Refresh> aufgerufen werden muß
Private Function HandleAutoUpdate_ChangedIcons(pDetails As Long) As Boolean
  If (prophImageList = 0) Or propUseCustomIcons Then Exit Function

  RebuildShellIconCache
  InvalidateIcons 0, -1, True
  HandleAutoUpdate_ChangedIcons = False
End Function

' behandelt das AutoUpdate-Event für das Ändern einer Drucker-Freigabe
' <SHNotData.dwItem1> ist die pIDL
' gibt True zurück, wenn <Refresh> aufgerufen werden muß
Private Function handleAutoUpdate_ChangedPrinterSharing(pIDLToDesktop As Long) As Boolean
  Dim hChangedItem As Long

  #If Debuging Then
    debugger.AddLogEntry "handleAutoUpdate_ChangedPrinterSharing: Calling InternalpIDLToItemHandle()", letInfo
    debugger.AddLogEntry "   0x" & Hex(pIDLToDesktop), LogEntryTypeConstants.letOther
  #End If
  hChangedItem = InternalpIDLToItemHandle(pIDLToDesktop)
  If Not IsWithinRange(hChangedItem, -1, 0) Then
    ReloadTVWItemProps hChangedItem
  End If
  handleAutoUpdate_ChangedPrinterSharing = False
End Function

' behandelt das AutoUpdate-Event für das Ändern einer Freigabe
' <SHNotData.dwItem1> ist die pIDL
' gibt True zurück, wenn <Refresh> aufgerufen werden muß
Private Function HandleAutoUpdate_ChangedSharing(pIDLToDesktop As Long) As Boolean
  Dim hChangedItem As Long
  Dim path As String

  #If Debuging Then
    debugger.AddLogEntry "HandleAutoUpdate_ChangedSharing: Calling InternalpIDLToItemHandle()", letInfo
    debugger.AddLogEntry "   0x" & Hex(pIDLToDesktop), LogEntryTypeConstants.letOther
  #End If
  hChangedItem = InternalpIDLToItemHandle(pIDLToDesktop)
  If Not IsWithinRange(hChangedItem, -1, 0) Then
    #If Debuging Then
      path = pIDLToPath_Light(debugger, pIDLToDesktop)
    #Else
      path = pIDLToPath_Light(pIDLToDesktop)
    #End If
    If IsDirectory(path) Or FileExists(path) Then
      ' flush the share cache
      On Error Resume Next
      IsPathShared path, 1
    End If

    ' now reload the overlay
    ReloadTVWItemProps hChangedItem
  End If

  HandleAutoUpdate_ChangedSharing = False
End Function

' behandelt das AutoUpdate-Event für das Erstellen einer Datei
' <SHNotData.dwItem1> ist die pIDL
' <SHNotData.dwItem2> ist 0
' gibt True zurück, wenn <Refresh> aufgerufen werden muß
Private Function HandleAutoUpdate_FileCreation(IParent As IVBShellFolder, pIDLToParent As Long) As Boolean
  If propIncludedItems And IncludedItemsConstants.iiFSFiles Then
    HandleAutoUpdate_FileCreation = HandleAutoUpdate_FolderCreation(IParent, pIDLToParent)
  End If
End Function

' behandelt das AutoUpdate-Event für das Erstellen einer Datei
' <SHNotData.dwItem1> ist die pIDL
' <SHNotData.dwItem2> ist 0
' gibt True zurück, wenn <Refresh> aufgerufen werden muß
Private Function HandleAutoUpdate_FileCreationFQ(pIDLToDesktop As Long) As Boolean
  If propIncludedItems And IncludedItemsConstants.iiFSFiles Then
    HandleAutoUpdate_FileCreationFQ = HandleAutoUpdate_FolderCreationFQ(pIDLToDesktop)
  End If
End Function

' behandelt das AutoUpdate-Event für das Löschen einer Datei
' <SHNotData.dwItem1> ist die pIDL
' sollte die Datei in den Papierkorb verschoben worden sein, ist <SHNotData.dwItem2> die neue
' pIDL
' gibt True zurück, wenn <Refresh> aufgerufen werden muß
Private Function HandleAutoUpdate_FileDeletion(pIDLToDesktop As Long) As Boolean
  If propIncludedItems And IncludedItemsConstants.iiFSFiles Then
    HandleAutoUpdate_FileDeletion = HandleAutoUpdate_FolderDeletion(pIDLToDesktop)
  End If
End Function

' behandelt das AutoUpdate-Event für das Verschieben einer Datei
' <SHNotData.dwItem1> ist die alte pIDL,
' <SHNotData.dwItem2> ist die neue pIDL
' gibt True zurück, wenn <Refresh> aufgerufen werden muß
Private Function HandleAutoUpdate_FileMoving(SHNotData As SHNOTIFY, IParent_New As IVBShellFolder, pIDLToParent_New As Long) As Boolean
  If propIncludedItems And IncludedItemsConstants.iiFSFiles Then
    HandleAutoUpdate_FileMoving = HandleAutoUpdate_FolderMoving(SHNotData, IParent_New, pIDLToParent_New)
  End If
End Function

' behandelt das AutoUpdate-Event für das Umbenennen einer Datei oder eines Archivs (unter WinMe/WinXP)
' <SHNotData.dwItem1> ist die alte pIDL,
' <SHNotData.dwItem2> ist die neue pIDL
' gibt True zurück, wenn <Refresh> aufgerufen werden muß
Private Function HandleAutoUpdate_FileRenaming(SHNotData As SHNOTIFY, pIDLParent_ToDesktop1 As Long, pIDLParent_ToDesktop2 As Long) As Boolean
  If propIncludedItems And IncludedItemsConstants.iiFSFiles Then
    HandleAutoUpdate_FileRenaming = handleAutoUpdate_FolderRenaming(SHNotData, pIDLParent_ToDesktop1, pIDLParent_ToDesktop2)
  End If
End Function

' behandelt das AutoUpdate-Event für das Erstellen eines Ordners
' <SHNotData.dwItem1> ist die pIDL
' <SHNotData.dwItem2> ist 0
' gibt True zurück, wenn <Refresh> aufgerufen werden muß
Private Function HandleAutoUpdate_FolderCreation(IParent As IVBShellFolder, pIDLToParent As Long) As Boolean
  Dim hItem As Long
  Dim hParentItem As Long
  Dim hRecycler As Long
  Dim IPF As IVBPersistFolder
  Dim IPF2 As IVBPersistFolder2
  Dim pIDLParent_ToDesktop As Long
  Dim pIDLRecycler As Long
  Dim pIDLToDesktop As Long

  IParent.QueryInterface IID_IPersistFolder, IPF
  If Not (IPF Is Nothing) Then
    IPF.QueryInterface IID_IPersistFolder2, IPF2
  Else
    IParent.QueryInterface IID_IPersistFolder2, IPF2
  End If
  If Not (IPF2 Is Nothing) Then
    IPF2.GetCurFolder pIDLParent_ToDesktop
  End If
  Set IPF = Nothing
  Set IPF2 = Nothing
  #If Debuging Then
    debugger.AddLogEntry "HandleAutoUpdate_FolderCreation: Calling InternalpIDLToItemHandle() #1", letInfo
    debugger.AddLogEntry "   0x" & Hex(pIDLParent_ToDesktop), LogEntryTypeConstants.letOther
  #End If
  hParentItem = InternalpIDLToItemHandle(pIDLParent_ToDesktop)

  Select Case hParentItem
    Case -1
      ' der ParentItem existiert nicht
      If GetRootOrFirstItem = -1 Then
        ' das Control ist leer
        ' -> möglicherweise ist der erstellte Ordner der RootItem
        ' -> in diesem Fall das Control füllen
        pIDLToDesktop = UserPathTopIDL(propRootPath, True)
        If pIDLToDesktop Then
          InitTreeView
          #If Debuging Then
            FreeItemIDList debugger, "HandleAutoUpdate_FolderCreation #1", pIDLToDesktop
          #Else
            FreeItemIDList pIDLToDesktop
          #End If
        End If
      Else
        ' keine weiteren Aktionen nötig
      End If
    Case Else
      ' der ParentItem  existiert bereits
      ' -> wenn seine SubItems schon geladen sind, diese ggf. ergänzen
      ' -> hat der ParentItem keine SubItems, den Item ggf. ergänzen
      If (GetItemChildrenProperty(hParentItem) = True) And (ItemHasSubItems(hParentItem) = False) Then
        ' der Item hat SubItems, diese wurden aber noch nicht geladen
        ' -> wenn wir den Item per "CreateNewFolder" selbst erzeugt haben, den Item
        '    unbedingt einfügen
        If createdFolder And (pIDLToParent <> 0) Then
          pIDLToDesktop = ILAppendID(ILClone(pIDLParent_ToDesktop), pIDLToParent, 1)
          #If Debuging Then
            debugger.AddLogEntry "HandleAutoUpdate_FolderCreation: Calling InternalpIDLToItemHandle() #2", letInfo
            debugger.AddLogEntry "   0x" & Hex(pIDLToDesktop), LogEntryTypeConstants.letOther
          #End If
          hCreatedItem = InternalpIDLToItemHandle(pIDLToDesktop, hParentItem, True)
          createdFolder = False
          #If Debuging Then
            FreeItemIDList debugger, "HandleAutoUpdate_FolderCreation #2", pIDLToDesktop
          #Else
            FreeItemIDList pIDLToDesktop
          #End If
        End If
      Else
        ' Item ggf. hinzufügen
        If pIDLToParent Then
          ' prüfen, ob der Item nicht schon existiert
          pIDLToDesktop = ILAppendID(ILClone(pIDLParent_ToDesktop), pIDLToParent, 1)
          #If Debuging Then
            debugger.AddLogEntry "HandleAutoUpdate_FolderCreation: Calling InternalpIDLToItemHandle() #3", letInfo
            debugger.AddLogEntry "   0x" & Hex(pIDLToDesktop), LogEntryTypeConstants.letOther
          #End If
          hCreatedItem = InternalpIDLToItemHandle(pIDLToDesktop, hParentItem, True)
          If hCreatedItem <> -1 Then createdFolder = False
        End If
      End If
  End Select

  If pIDLToParent Then
    pIDLToDesktop = ILAppendID(ILClone(pIDLParent_ToDesktop), pIDLToParent, 1)
    #If Debuging Then
      If IsPartOfRecycler(debugger, pIDLToDesktop) Then
    #Else
      If IsPartOfRecycler(pIDLToDesktop) Then
    #End If
      'HandleAutoUpdate_UpdatedRecyclers
      ' Papierkorb auf dem Desktop aktualisieren
      pIDLRecycler = CSIDLTopIDL(CSIDLConstants.CSIDL_BITBUCKET)
      #If Debuging Then
        debugger.AddLogEntry "HandleAutoUpdate_FolderCreation: Calling InternalpIDLToItemHandle() #4", letInfo
        debugger.AddLogEntry "   0x" & Hex(pIDLRecycler), LogEntryTypeConstants.letOther
      #End If
      hRecycler = InternalpIDLToItemHandle(pIDLRecycler, , , True)
      If hRecycler <> -1 Then
        InsertNewAndRemoveOldSubItems hRecycler
      End If
    End If
  End If
  #If Debuging Then
    FreeItemIDList debugger, "HandleAutoUpdate_FolderCreation #3", pIDLToDesktop

    FreeItemIDList debugger, "HandleAutoUpdate_FolderCreation #4", pIDLParent_ToDesktop
  #Else
    FreeItemIDList pIDLToDesktop

    FreeItemIDList pIDLParent_ToDesktop
  #End If
  HandleAutoUpdate_FolderCreation = False
End Function

' behandelt das AutoUpdate-Event für das Erstellen eines Ordners
' <SHNotData.dwItem1> ist die pIDL
' <SHNotData.dwItem2> ist 0
' gibt True zurück, wenn <Refresh> aufgerufen werden muß
Private Function HandleAutoUpdate_FolderCreationFQ(pIDLToDesktop As Long) As Boolean
  Dim hItem As Long
  Dim hParentItem As Long
  Dim hRecycler As Long
  Dim pIDL As Long
  Dim pIDLParent_ToDesktop As Long
  Dim pIDLRecycler As Long

  pIDLParent_ToDesktop = ILClone(pIDLToDesktop)
  MakeParentItemIDList pIDLParent_ToDesktop
  #If Debuging Then
    debugger.AddLogEntry "HandleAutoUpdate_FolderCreationFQ: Calling InternalpIDLToItemHandle() #1", letInfo
    debugger.AddLogEntry "   0x" & Hex(pIDLParent_ToDesktop), LogEntryTypeConstants.letOther
  #End If
  hParentItem = InternalpIDLToItemHandle(pIDLParent_ToDesktop)

  Select Case hParentItem
    Case -1
      ' der ParentItem existiert nicht
      If GetRootOrFirstItem = -1 Then
        ' das Control ist leer
        ' -> möglicherweise ist der erstellte Ordner der RootItem
        ' -> in diesem Fall das Control füllen
        pIDL = UserPathTopIDL(propRootPath, True)
        If pIDL Then
          InitTreeView
          #If Debuging Then
            FreeItemIDList debugger, "HandleAutoUpdate_FolderCreationFQ #1", pIDL
          #Else
            FreeItemIDList pIDL
          #End If
        End If
      Else
        ' keine weiteren Aktionen nötig
      End If
    Case Else
      ' der ParentItem  existiert bereits
      ' -> wenn seine SubItems schon geladen sind, diese ggf. ergänzen
      ' -> hat der ParentItem keine SubItems, den Item ggf. ergänzen
      If (GetItemChildrenProperty(hParentItem) = True) And (ItemHasSubItems(hParentItem) = False) Then
        ' der Item hat SubItems, diese wurden aber noch nicht geladen
        ' -> wenn wir den Item per "CreateNewFolder" selbst erzeugt haben, den Item
        '    unbedingt einfügen
        If createdFolder Then
          #If Debuging Then
            debugger.AddLogEntry "HandleAutoUpdate_FolderCreationFQ: Calling InternalpIDLToItemHandle() #2", letInfo
            debugger.AddLogEntry "   0x" & Hex(pIDLToDesktop), LogEntryTypeConstants.letOther
          #End If
          hCreatedItem = InternalpIDLToItemHandle(pIDLToDesktop, hParentItem, True)
          createdFolder = False
          #If Debuging Then
            FreeItemIDList debugger, "HandleAutoUpdate_FolderCreationFQ #2", pIDL
          #Else
            FreeItemIDList pIDL
          #End If
        End If
      Else
        ' Item ggf. hinzufügen
        ' prüfen, ob der Item nicht schon existiert
        #If Debuging Then
          debugger.AddLogEntry "HandleAutoUpdate_FolderCreationFQ: Calling InternalpIDLToItemHandle() #3", letInfo
          debugger.AddLogEntry "   0x" & Hex(pIDLToDesktop), LogEntryTypeConstants.letOther
        #End If
        hCreatedItem = InternalpIDLToItemHandle(pIDLToDesktop, hParentItem, True)
        If hCreatedItem <> -1 Then createdFolder = False
      End If
  End Select

  #If Debuging Then
    If IsPartOfRecycler(debugger, pIDLToDesktop) Then
  #Else
    If IsPartOfRecycler(pIDLToDesktop) Then
  #End If
    'HandleAutoUpdate_UpdatedRecyclers
    ' Papierkorb auf dem Desktop aktualisieren
    pIDLRecycler = CSIDLTopIDL(CSIDLConstants.CSIDL_BITBUCKET)
    #If Debuging Then
      debugger.AddLogEntry "HandleAutoUpdate_FolderCreationFQ: Calling InternalpIDLToItemHandle() #4", letInfo
      debugger.AddLogEntry "   0x" & Hex(pIDLRecycler), LogEntryTypeConstants.letOther
    #End If
    hRecycler = InternalpIDLToItemHandle(pIDLRecycler, , , True)
    If hRecycler <> -1 Then
      InsertNewAndRemoveOldSubItems hRecycler
    End If
  End If

  #If Debuging Then
    FreeItemIDList debugger, "HandleAutoUpdate_FolderCreationFQ #3", pIDLParent_ToDesktop
  #Else
    FreeItemIDList pIDLParent_ToDesktop
  #End If
  HandleAutoUpdate_FolderCreationFQ = False
End Function

' behandelt das AutoUpdate-Event für das Löschen eines Ordners
' <SHNotData.dwItem1> ist die pIDL
' sollte der Ordner in den Papierkorb verschoben worden sein, ist <SHNotData.dwItem2> die neue pIDL
' gibt True zurück, wenn <Refresh> aufgerufen werden muß
Private Function HandleAutoUpdate_FolderDeletion(pIDLToDesktop As Long) As Boolean
  Dim hDeletedItem As Long
  Dim hRecycler As Long
  Dim pIDLRecycler As Long

  #If Debuging Then
    debugger.AddLogEntry "HandleAutoUpdate_FolderDeletion: Calling InternalpIDLToItemHandle() #1", letInfo
    debugger.AddLogEntry "   0x" & Hex(pIDLToDesktop), LogEntryTypeConstants.letOther
  #End If
  hDeletedItem = InternalpIDLToItemHandle(pIDLToDesktop)
  If hDeletedItem = -1 Then
    ' der Item existiert nicht
    ' -> wenn der RootItem ein SubItem des Ordners war, wurde er mit gelöscht
    If ILIsParent(pIDLToDesktop, RootItemInfo.pIDLToDesktop, 0) Then
      ' alle Items entfernen
      ClearTVWItems
    Else
      ' keine weiteren Aktionen nötig
    End If
  Else
    RemoveTVWItem hDeletedItem
  End If

  #If Debuging Then
    If IsPartOfRecycler(debugger, pIDLToDesktop) Then
  #Else
    If IsPartOfRecycler(pIDLToDesktop) Then
  #End If
    'HandleAutoUpdate_UpdatedRecyclers
    ' Papierkorb auf dem Desktop aktualisieren
    pIDLRecycler = CSIDLTopIDL(CSIDLConstants.CSIDL_BITBUCKET)
    #If Debuging Then
      debugger.AddLogEntry "HandleAutoUpdate_FolderDeletion: Calling InternalpIDLToItemHandle() #2", letInfo
      debugger.AddLogEntry "   0x" & Hex(pIDLRecycler), LogEntryTypeConstants.letOther
    #End If
    hRecycler = InternalpIDLToItemHandle(pIDLRecycler, , , True)
    If hRecycler <> -1 Then
      InsertNewAndRemoveOldSubItems hRecycler
    End If
  End If

  HandleAutoUpdate_FolderDeletion = False
End Function

' behandelt das AutoUpdate-Event für das Verschieben eines Ordners
' <SHNotData.dwItem1> ist die alte pIDL,
' <SHNotData.dwItem2> ist die neue pIDL
' gibt True zurück, wenn <Refresh> aufgerufen werden muß
Private Function HandleAutoUpdate_FolderMoving(SHNotData As SHNOTIFY, IParent_New As IVBShellFolder, pIDLToParent_New As Long) As Boolean
  HandleAutoUpdate_FolderMoving = HandleAutoUpdate_FolderDeletion(SHNotData.dwItem1) Or HandleAutoUpdate_FolderCreation(IParent_New, pIDLToParent_New)
End Function

' behandelt das AutoUpdate-Event für das Umbenennen eines Ordners
' <SHNotData.dwItem1> ist die alte pIDL,
' <SHNotData.dwItem2> ist die neue pIDL
' gibt True zurück, wenn <Refresh> aufgerufen werden muß
Private Function handleAutoUpdate_FolderRenaming(SHNotData As SHNOTIFY, pIDLParent_ToDesktop1 As Long, pIDLParent_ToDesktop2 As Long) As Boolean
  Dim hChangedItem As Long
  Dim hParentItem As Long
  Dim hSubItem As Long
  Dim IParent2 As IVBShellFolder
  Dim pIDLToDesktop2 As Long
  Dim pIDLToParent1 As Long
  Dim pIDLToParent2 As Long
  Dim ret As Boolean

  With SHNotData
    If .dwItem1 Then pIDLToParent1 = ILFindLastID(.dwItem1)

    If .dwItem2 Then
      #If Debuging Then
        SplitFullyQualifiedPIDL debugger, .dwItem2, IParent2, pIDLToParent2
      #Else
        SplitFullyQualifiedPIDL .dwItem2, IParent2, pIDLToParent2
      #End If
    End If
    If (pIDLToParent2 <> 0) And Not (IParent2 Is Nothing) Then
      #If Debuging Then
        pIDLToParent2 = SimplePIDLToRealPIDL(debugger, IParent2, pIDLToParent2)
        If pIDLToParent2 Then
          pIDLToDesktop2 = ILAppendID(ILClone(pIDLParent_ToDesktop2), pIDLToParent2, 1)
          FreeItemIDList debugger, "handleAutoUpdate_FolderRenaming #1", pIDLToParent2
        End If
      #Else
        pIDLToParent2 = SimplePIDLToRealPIDL(IParent2, pIDLToParent2)
        If pIDLToParent2 Then
          pIDLToDesktop2 = ILAppendID(ILClone(pIDLParent_ToDesktop2), pIDLToParent2, 1)
          FreeItemIDList pIDLToParent2
        End If
      #End If
      If pIDLToDesktop2 Then
        pIDLToParent2 = ILFindLastID(pIDLToDesktop2)
      End If
    End If

    If ILIsParent(SHNotData.dwItem1, RootItemInfo.pIDLToDesktop, 0) Then
      hChangedItem = IIf(propShowRoot, GetRootOrFirstItem, 0)
    Else
      #If Debuging Then
        debugger.AddLogEntry "handleAutoUpdate_FolderRenaming: Calling InternalpIDLToItemHandle()", letInfo
        debugger.AddLogEntry "   0x" & Hex(.dwItem1), LogEntryTypeConstants.letOther
      #End If
      hChangedItem = InternalpIDLToItemHandle(.dwItem1)
    End If

    If hChangedItem = -1 Then
      ' Okay - wir müssen nichts updaten, aber den Item vielleicht einfügen, weil er jetzt
      ' vielleicht angezeigt werden soll
      If pIDLToParent2 Then
        ret = HandleAutoUpdate_FolderCreation(IParent2, pIDLToParent2)
      End If
    ElseIf ILIsEqual(.dwItem1, .dwItem2) Then
      ' die pIDLs sind gleich
      ' -> wahrscheinlich ein umbenanntes Laufwerk
      ' -> Eigenschaften trotzdem neu laden
      hParentItem = ItemGetParentItem(hChangedItem)
      ReloadTVWItemProps hChangedItem
      SortItems hParentItem
    Else
      ' <hChangedItem> updaten
      hParentItem = ItemGetParentItem(hChangedItem)
      If pIDLToDesktop2 Then
        ReloadTVWItemProps hChangedItem, pIDLToDesktop2
        SortItems hParentItem
      End If
    End If

    #If Debuging Then
      FreeItemIDList debugger, "handleAutoUpdate_FolderRenaming #2", pIDLToDesktop2
    #Else
      FreeItemIDList pIDLToDesktop2
    #End If
  End With
  Set IParent2 = Nothing

  handleAutoUpdate_FolderRenaming = ret
End Function

' behandelt das AutoUpdate-Event für das Einlegen eines Datenträgers
' <SHNotData.dwItem1> ist die pIDL
' <SHNotData.dwItem2> ist 0
' gibt True zurück, wenn <Refresh> aufgerufen werden muß
Private Function HandleAutoUpdate_InsertedMedia(pIDLToDesktop As Long) As Boolean
  Dim hChangedItem As Long
  Dim pIDL As Long

  #If Debuging Then
    debugger.AddLogEntry "HandleAutoUpdate_InsertedMedia: Calling InternalpIDLToItemHandle()", letInfo
    debugger.AddLogEntry "   0x" & Hex(pIDLToDesktop), LogEntryTypeConstants.letOther
  #End If
  hChangedItem = InternalpIDLToItemHandle(pIDLToDesktop)
  Select Case hChangedItem
    Case -1
      ' der Item existiert nicht
      If GetRootOrFirstItem = -1 Then
        ' das Control ist leer
        ' -> möglicherweise ist der RootItem auf dem Datenträger
        ' -> in diesem Fall das Control füllen
        pIDL = UserPathTopIDL(propRootPath, True)
        If pIDL Then
          InitTreeView
          #If Debuging Then
            FreeItemIDList debugger, "HandleAutoUpdate_InsertedMedia", pIDL
          #Else
            FreeItemIDList pIDL
          #End If
        End If
      Else
        ' keine weiteren Aktionen nötig
      End If
    Case 0
      ' der Item ist der RootItem, wird aber nicht angezeigt
      ' -> Control füllen
      InitTreeView
    Case Else
      ' der Item existiert
      ' -> Eigenschaften neu einlesen
      ReloadTVWItemProps hChangedItem
  End Select

  HandleAutoUpdate_InsertedMedia = False
End Function

' behandelt das AutoUpdate-Event für das Entfernen eines Laufwerks
' <SHNotData.dwItem1> ist die pIDL (wenn das Laufwerk über den Gerätemanager entfernt wurde,
' ist es irgendetwas anderes)
' <SHNotData.dwItem2> ist 0
' gibt True zurück, wenn <Refresh> aufgerufen werden muß
Private Function HandleAutoUpdate_RemovedDrive(IParent As IVBShellFolder, pIDLToParent As Long) As Boolean
  Dim hParentItem As Long
  Dim IMyComputer As IVBShellFolder
  Dim pIDL As Long
  Dim pIDLMyComputer_ToDesktop As Long

  pIDLMyComputer_ToDesktop = CSIDLTopIDL(CSIDLConstants.CSIDL_DRIVES)
  #If Debuging Then
    debugger.AddLogEntry "HandleAutoUpdate_RemovedDrive: Calling InternalpIDLToItemHandle()", letInfo
    debugger.AddLogEntry "   0x" & Hex(pIDLMyComputer_ToDesktop), LogEntryTypeConstants.letOther
  #End If
  hParentItem = InternalpIDLToItemHandle(pIDLMyComputer_ToDesktop)

  If hParentItem = -1 Then
    ' der Arbeitsplatz existiert nicht
    ' -> wenn der RootItem ein SubItem des Laufwerks war, wurde er mit entfernt
    #If Debuging Then
      If pIDLIncludesMyComputer(debugger, RootItemInfo.pIDLToDesktop) Then
        ' prüfen, ob der RootItem noch existiert
        Set IMyComputer = GetISHFolderInterface(debugger, IDesktop, pIDLMyComputer_ToDesktop)
        pIDL = CopyFirstItemIDs(debugger, RootItemInfo.pIDLToDesktop, 2)

        If Not ShouldShowDrive(debugger, IMyComputer, ILFindLastID(pIDL)) Then
          ClearTVWItems
        End If
        FreeItemIDList debugger, "HandleAutoUpdate_RemovedDrive #1", pIDL
        Set IMyComputer = Nothing
      Else
        ' keine weiteren Aktionen nötig
      End If
    #Else
      If pIDLIncludesMyComputer(RootItemInfo.pIDLToDesktop) Then
        ' prüfen, ob der RootItem noch existiert
        Set IMyComputer = GetISHFolderInterface(IDesktop, pIDLMyComputer_ToDesktop)
        pIDL = CopyFirstItemIDs(RootItemInfo.pIDLToDesktop, 2)

        If Not ShouldShowDrive(IMyComputer, ILFindLastID(pIDL)) Then
          ClearTVWItems
        End If
        FreeItemIDList pIDL
        Set IMyComputer = Nothing
      Else
        ' keine weiteren Aktionen nötig
      End If
    #End If
  Else
    ' der Arbeitsplatz  existiert
    ' -> wenn seine SubItems schon geladen sind, diese ggf. ergänzen/alte SubItems entfernen
    If (GetItemChildrenProperty(hParentItem) = True) And (ItemHasSubItems(hParentItem) = False) Then
      ' der Arbeitsplatz hat SubItems, diese wurden aber noch nicht geladen
      ' -> keine weiteren Aktionen nötig
    Else
      ' neue SubItems einfügen/alte entfernen
      InsertNewAndRemoveOldSubItems hParentItem, False, True, False
    End If
  End If

  #If Debuging Then
    FreeItemIDList debugger, "HandleAutoUpdate_RemovedDrive #2", pIDLMyComputer_ToDesktop
  #Else
    FreeItemIDList pIDLMyComputer_ToDesktop
  #End If
  HandleAutoUpdate_RemovedDrive = False
End Function

' behandelt das AutoUpdate-Event für das Entnehmen eines Datenträgers
' <SHNotData.dwItem1> ist die pIDL
' <SHNotData.dwItem2> ist 0
' gibt True zurück, wenn <Refresh> aufgerufen werden muß
Private Function HandleAutoUpdate_RemovedMedia(pIDLToDesktop As Long) As Boolean
  Dim hChangedItem As Long

  #If Debuging Then
    debugger.AddLogEntry "HandleAutoUpdate_RemovedMedia: Calling InternalpIDLToItemHandle()", letInfo
    debugger.AddLogEntry "   0x" & Hex(pIDLToDesktop), LogEntryTypeConstants.letOther
  #End If
  hChangedItem = InternalpIDLToItemHandle(pIDLToDesktop)
  If hChangedItem = -1 Then
    ' der Item existiert nicht
    ' -> wenn der RootItem ein SubItem des Ordners war, wurde er mit gelöscht
    If ILIsParent(pIDLToDesktop, RootItemInfo.pIDLToDesktop, 0) Then
      ' alle Items entfernen
      ClearTVWItems
    Else
      ' keine weiteren Aktionen nötig
    End If
  Else
    ' der Item existiert
    ' -> Eigenschaften neu einlesen
    ' WORKAROUND: wenn RemoveSubItems erst innerhalb von ReloadTVWItemProps aufgerufen
    '             wird, scheint TVE_COLLAPSERESET keine Wirkung zu haben
    RemoveSubItems hChangedItem
    ReloadTVWItemProps hChangedItem
  End If

  HandleAutoUpdate_RemovedMedia = False
End Function

' behandelt das AutoUpdate-Event für das Umbenennen eines NonFS-Items
' <SHNotData.dwItem1> ist die neue (und alte) pIDL
' gibt True zurück, wenn <Refresh> aufgerufen werden muß
Private Function handleAutoUpdate_RenamedNonFSItem(pIDLToDesktop As Long) As Boolean
  Dim hChangedItem As Long

  #If Debuging Then
    debugger.AddLogEntry "handleAutoUpdate_RenamedNonFSItem: Calling InternalpIDLToItemHandle()", letInfo
    debugger.AddLogEntry "   0x" & Hex(pIDLToDesktop), LogEntryTypeConstants.letOther
  #End If
  hChangedItem = InternalpIDLToItemHandle(pIDLToDesktop)
  If hChangedItem <> -1 Then ReloadTVWItemProps hChangedItem
  handleAutoUpdate_RenamedNonFSItem = False
End Function

' behandelt das AutoUpdate-Event für das Ein-/Ausblenden eines NonFS-Items bzw. für die
' geänderte Anordnung von NonFS-Items
' <SHNotData.dwItem1> ist ??
' gibt True zurück, wenn <Refresh> aufgerufen werden muß
Private Function HandleAutoUpdate_ShowOrHideNonFSItem() As Boolean
  ReloadItems
  HandleAutoUpdate_ShowOrHideNonFSItem = False
End Function

' behandelt das AutoUpdate-Event für das allgemeine Updaten eines Ordners und seines Inhalts
' <SHNotData.dwItem1> ist die neue pIDL
' gibt True zurück, wenn <Refresh> aufgerufen werden muß
Private Function HandleAutoUpdate_UpdatedFolder(pIDLToDesktop As Long) As Boolean
  Const WM_SETREDRAW = &HB
  Dim hChangedItem As Long
  Dim hParentItem As Long
  Dim IParent As IVBShellFolder
  Dim isNeroScout As Boolean
  Dim pIDLToParent As Long
  Dim ret As Boolean

  #If Debuging Then
    isNeroScout = (pIDLToPath_Light(debugger, pIDLToDesktop) Like "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}\::{3D6BE802-FC0D-4595-A304-E611F97089DC}\*")
  #Else
    isNeroScout = (pIDLToPath_Light(pIDLToDesktop) Like "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}\::{3D6BE802-FC0D-4595-A304-E611F97089DC}\*")
  #End If

  #If Debuging Then
    debugger.AddLogEntry "HandleAutoUpdate_UpdatedFolder: Calling InternalpIDLToItemHandle()", letInfo
    debugger.AddLogEntry "   0x" & Hex(pIDLToDesktop), LogEntryTypeConstants.letOther
  #End If
  hChangedItem = InternalpIDLToItemHandle(pIDLToDesktop)
  If hChangedItem = -1 Then
    ' Okay - wir müssen nichts updaten, aber den Item vielleicht einfügen, weil er jetzt vielleicht
    ' angezeigt werden soll
    #If Debuging Then
      If ValidateItemFQ(debugger, pIDLToDesktop) Then
        SplitFullyQualifiedPIDL debugger, pIDLToDesktop, IParent, pIDLToParent
    #Else
      If ValidateItemFQ(pIDLToDesktop) Then
        SplitFullyQualifiedPIDL pIDLToDesktop, IParent, pIDLToParent
    #End If
      ret = HandleAutoUpdate_FolderCreation(IParent, pIDLToParent)
      Set IParent = Nothing
    Else
      ' Item wurde wohl entfernt
    End If
  #If Debuging Then
    ElseIf Not ValidateItemFQ(debugger, pIDLToDesktop) Then
      debugger.AddLogEntry "   The updated item is invalid - probably has been removed.", LogEntryTypeConstants.letOther
  #Else
    ElseIf Not ValidateItemFQ(pIDLToDesktop) Then
  #End If
    ' Item wurde wohl entfernt
    RemoveTVWItem hChangedItem
  Else
    ' Should we still display the item?
    If ILIsEqual(pIDLToDesktop, pIDL_Desktop) Then
      InsertNewAndRemoveOldSubItems hChangedItem
    Else
      InsertNewAndRemoveOldSubItems hChangedItem, Not isNeroScout, , False
      hParentItem = ItemGetParentItem(hChangedItem)
      ReloadTVWItemPIDL hChangedItem
      If hParentItem <> -1 Then SortItems hParentItem
    End If
  End If
  #If Debuging Then
    debugger.AddLogEntry "Leaving HandleAutoUpdate_UpdatedFolder...", LogEntryTypeConstants.letInfo
  #End If

  HandleAutoUpdate_UpdatedFolder = ret
End Function

' behandelt das AutoUpdate-Event für das allgemeine Updaten eines FS-Items
' <SHNotData.dwItem1> ist die neue pIDL
' gibt True zurück, wenn <Refresh> aufgerufen werden muß
Private Function handleAutoUpdate_UpdatedFSItem(pIDLToDesktop As Long) As Boolean
  handleAutoUpdate_UpdatedFSItem = HandleAutoUpdate_UpdatedFolder(pIDLToDesktop)
End Function

' behandelt das AutoUpdate-Event für das allgemeine Updaten eines Papierkorbes
' <SHNotData.dwItem1> ist die pIDL des betroffenen Papierkorbs
' gibt True zurück, wenn <Refresh> aufgerufen werden muß
Private Function HandleAutoUpdate_UpdatedRecycler(pIDLRecyclerToDesktop As Long, ISHFRecycler As IVBShellFolder, Optional ByVal reEnumerateSubItems As Boolean = False) As Boolean
'  Dim arrRecyclers() As String
  Dim hRecycler As Long
'  Dim i As Long
  Dim pIDLRecycler As Long
  Dim TVWItemInfo As ITEMINFO_COMMON
  Static lastCall As Long

  If (GetTickCount - lastCall) <= 700 Then Exit Function
  lastCall = GetTickCount

  ' alle Papierkörbe suchen und nicht mehr vorhandene Sub-Items löschen
'  #If Debuging Then
'    arrRecyclers = Split(GetAllRecycleBins(debugger), "|")
'  #Else
'    arrRecyclers = Split(GetAllRecycleBins, "|")
'  #End If
'  If Not IsEmpty(arrRecyclers) Then
'    For i = LBound(arrRecyclers) To UBound(arrRecyclers)
      #If Debuging Then
        debugger.AddLogEntry "HandleAutoUpdate_UpdatedRecycler: Calling InternalpIDLToItemHandle() #1", letInfo
        debugger.AddLogEntry "   0x" & Hex(pIDLRecyclerToDesktop), LogEntryTypeConstants.letOther
      #End If
      hRecycler = InternalpIDLToItemHandle(pIDLRecyclerToDesktop)

      If hRecycler <> -1 Then
        InvalidateIcons hRecycler, -1, True, False
        If reEnumerateSubItems Then InsertNewAndRemoveOldSubItems hRecycler
        If Not dontChangeRecyclerChildrenProp Then
          #If Debuging Then
            SetItemChildrenProperty hRecycler, HasSubItems(debugger, pIDLRecyclerToDesktop, Me)
          #Else
            SetItemChildrenProperty hRecycler, HasSubItems(pIDLRecyclerToDesktop, Me)
          #End If
        End If
      End If
'    Next i
'  End If

  ' Papierkorb auf dem Desktop aktualisieren
  pIDLRecycler = CSIDLTopIDL(CSIDLConstants.CSIDL_BITBUCKET)
  #If Debuging Then
    debugger.AddLogEntry "HandleAutoUpdate_UpdatedRecycler: Calling InternalpIDLToItemHandle() #2", letInfo
    debugger.AddLogEntry "   0x" & Hex(pIDLRecycler), LogEntryTypeConstants.letOther
  #End If
  hRecycler = InternalpIDLToItemHandle(pIDLRecycler, , , True)
  If hRecycler <> -1 Then
    GetItemInfo_Common hRecycler, TVWItemInfo

    InvalidateIcons hRecycler, -1, True, False
    If reEnumerateSubItems Then InsertNewAndRemoveOldSubItems hRecycler
    If Not dontChangeRecyclerChildrenProp Then
      #If Debuging Then
        SetItemChildrenProperty hRecycler, HasSubItems(debugger, TVWItemInfo.pIDLToDesktop, Me)
      #Else
        SetItemChildrenProperty hRecycler, HasSubItems(TVWItemInfo.pIDLToDesktop, Me)
      #End If
    End If
  End If

  HandleAutoUpdate_UpdatedRecycler = False
End Function

' behandelt das AutoUpdate-Event für das allgemeine Updaten eines Papierkorbes
' <SHNotData.dwItem1> ist die pIDL des betroffenen Papierkorbs
' gibt True zurück, wenn <Refresh> aufgerufen werden muß
Private Function HandleAutoUpdate_UpdatedRecyclers(Optional ByVal reEnumerateSubItems As Boolean = False) As Boolean
  Dim arrRecyclers() As String
  Dim hRecycler As Long
  Dim i As Long
  Dim pIDLRecycler As Long
  Dim TVWItemInfo As ITEMINFO_COMMON
'  Static lastCall As Long

'  If (GetTickCount - lastCall) <= 700 Then Exit Function
'  lastCall = GetTickCount

  ' alle Papierkörbe suchen und nicht mehr vorhandene Sub-Items löschen
  #If Debuging Then
    arrRecyclers = Split(GetAllRecycleBins(debugger), "|")
  #Else
    arrRecyclers = Split(GetAllRecycleBins, "|")
  #End If
  If Not IsEmpty(arrRecyclers) Then
    For i = LBound(arrRecyclers) To UBound(arrRecyclers)
      pIDLRecycler = PathTopIDL(arrRecyclers(i))
      #If Debuging Then
        debugger.AddLogEntry "HandleAutoUpdate_UpdatedRecyclers: Calling InternalpIDLToItemHandle() #1", letInfo
        debugger.AddLogEntry "   0x" & Hex(pIDLRecycler), LogEntryTypeConstants.letOther
      #End If
      hRecycler = InternalpIDLToItemHandle(pIDLRecycler, , , True)

      If hRecycler <> -1 Then
        InvalidateIcons hRecycler, -1, True, False
        If reEnumerateSubItems Then InsertNewAndRemoveOldSubItems hRecycler
        If Not dontChangeRecyclerChildrenProp Then
          #If Debuging Then
            SetItemChildrenProperty hRecycler, HasSubItems(debugger, pIDLRecycler, Me)
          #Else
            SetItemChildrenProperty hRecycler, HasSubItems(pIDLRecycler, Me)
          #End If
        End If
      End If
    Next i
  End If

  ' Papierkorb auf dem Desktop aktualisieren
  pIDLRecycler = CSIDLTopIDL(CSIDLConstants.CSIDL_BITBUCKET)
  #If Debuging Then
    debugger.AddLogEntry "HandleAutoUpdate_UpdatedRecyclers: Calling InternalpIDLToItemHandle() #2", letInfo
    debugger.AddLogEntry "   0x" & Hex(pIDLRecycler), LogEntryTypeConstants.letOther
  #End If
  hRecycler = InternalpIDLToItemHandle(pIDLRecycler, , , True)
  If hRecycler <> -1 Then
    GetItemInfo_Common hRecycler, TVWItemInfo

    InvalidateIcons hRecycler, -1, True, False
    If reEnumerateSubItems Then InsertNewAndRemoveOldSubItems hRecycler
    If Not dontChangeRecyclerChildrenProp Then
      #If Debuging Then
        SetItemChildrenProperty hRecycler, HasSubItems(debugger, TVWItemInfo.pIDLToDesktop, Me)
      #Else
        SetItemChildrenProperty hRecycler, HasSubItems(TVWItemInfo.pIDLToDesktop, Me)
      #End If
    End If
  End If

  HandleAutoUpdate_UpdatedRecyclers = False
End Function

' behandelt die AutoUpdate-Events
Private Sub HandleSHNotify(ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long)
  Const WM_SETREDRAW = &HB
  Dim IParent As IVBShellFolder
  Dim isRecycler As Boolean
  Dim needRefresh As Boolean
  Dim oldIconIndex As Long
  Dim pIDL As Long
  Dim pIDL2 As Long
  Dim pIDLParent_ToDesktop1 As Long
  Dim pIDLParent_ToDesktop2 As Long
  Dim pIDLToParent As Long
  Dim SHNotData As SHNOTIFY

  If blockAutoUpdate Then Exit Sub

  If lParam And SHCNEConstants.SHCNE_INTERRUPT Then
    lParam = lParam And &HFFFFFFF
    #If Debuging Then
      debugger.AddLogEntry "   AutoUpdate: Interrupt event 0x" & Hex(lParam), LogEntryTypeConstants.letInfo
    #End If
  End If

  #If Debuging Then
    If lParam And SHCNEConstants.SHCNE_EXTENDED_EVENT Then
      CopyMemory VarPtr(SHNotData), wParam, LenB(SHNotData)
      debugger.AddLogEntry "   AutoUpdate: Unhandled extended event 0x" & Hex(SHNotData.dwItem1), LogEntryTypeConstants.letWarning
      Exit Sub
    End If
  #End If

  Select Case lParam
    Case SHCNEConstants.SHCNE_ASSOCCHANGED
      ' eine Dateiverknüpfung hat sich geändert, das Icon für einen Dateitypen wurde geändert
      ' oder ein NonFS-Item wurde ein-/ausgeblendet oder die Anordnung von NonFS-Items wurde
      ' geändert oder das Icon eines NonFS-Items wurde geändert
      ' Dateiverknüpfung:
      '   <SHNotData.dwItem1> ist die pIDL der Dateiendung
      '      -> DispName: "*.xxx"
      '      -> Pfad    : "C:\*.xxx"
      ' Icon:
      '   <SHNotData.dwItem1> ist die pIDL des Desktops oder 0
      ' NonFS-Item:
      '   <SHNotData.dwItem1> ist ??
      CopyMemory VarPtr(SHNotData), wParam, LenB(SHNotData)
      If SHNotData.dwItem1 = 0 Then
        ' geändertes Icon
        #If Debuging Then
          debugger.AddLogEntry "   AutoUpdate: The icon of a file type or NonFS item has changed.", LogEntryTypeConstants.letInfo
        #End If
        needRefresh = HandleAutoUpdate_ChangedIcons(0)
      ElseIf ILIsEqual(SHNotData.dwItem1, pIDL_Desktop) Then
        ' geändertes Icon
        #If Debuging Then
          debugger.AddLogEntry "   AutoUpdate: The icon of a file type or NonFS item has changed.", LogEntryTypeConstants.letInfo
        #End If
        needRefresh = HandleAutoUpdate_ChangedIcons(SHNotData.dwItem1)
      #If Debuging Then
        ElseIf pIDLToDisplayName_Light(debugger, SHNotData.dwItem1) <> "" Then
      #Else
        ElseIf pIDLToDisplayName_Light(SHNotData.dwItem1) <> "" Then
      #End If
        ' geänderte Dateityp-Verknüpfung
        #If Debuging Then
          debugger.AddLogEntry "   AutoUpdate: A file type link has changed.", LogEntryTypeConstants.letInfo
          debugger.AddLogEntry "      " & pIDLToPath_Light(debugger, SHNotData.dwItem1), LogEntryTypeConstants.letOther
        #End If
        needRefresh = HandleAutoUpdate_ChangedFileTypeLink(SHNotData.dwItem1)
      Else
        ' ein-/ausgeblendeter NonFS-Item oder geänderte Anordnung von NonFS-Items
        #If Debuging Then
          debugger.AddLogEntry "   AutoUpdate: The visibility of a NonFS item or the order of NonFS items has changed.", LogEntryTypeConstants.letInfo
        #End If
        needRefresh = HandleAutoUpdate_ShowOrHideNonFSItem()
      End If

    Case SHCNEConstants.SHCNE_ATTRIBUTES
      ' der Status eines Druckers hat sich geändert
      ' <SHNotData.dwItem1> ist die pIDL
      ' keine weiteren Aktionen nötig
      #If Debuging Then
        debugger.AddLogEntry "   AutoUpdate: A printer's state has changed.", LogEntryTypeConstants.letInfo
        CopyMemory VarPtr(SHNotData), wParam, LenB(SHNotData)
        debugger.AddLogEntry "      " & pIDLToDisplayName_Light(debugger, SHNotData.dwItem1), LogEntryTypeConstants.letOther
      #End If

    Case SHCNEConstants.SHCNE_CREATE
      ' eine Datei wurde erstellt
      CopyMemory VarPtr(SHNotData), wParam, LenB(SHNotData)
      If SHNotData.dwItem1 <> 0 Then
        #If Debuging Then
          debugger.AddLogEntry "   AutoUpdate: A file has been created.", LogEntryTypeConstants.letInfo
          debugger.AddLogEntry "      " & pIDLToPath_Light(debugger, SHNotData.dwItem1), LogEntryTypeConstants.letOther
        #End If

        #If Debuging Then
          SplitFullyQualifiedPIDL debugger, SHNotData.dwItem1, IParent, pIDLToParent
        #Else
          SplitFullyQualifiedPIDL SHNotData.dwItem1, IParent, pIDLToParent
        #End If
        ' beim Drucken wird SHCNE_CREATE mit einer ungültigen pIDL geschickt!
        If pIDLToParent <> 0 And Not (IParent Is Nothing) Then
          #If Debuging Then
            pIDLToParent = SimplePIDLToRealPIDL(debugger, IParent, pIDLToParent)
          #Else
            pIDLToParent = SimplePIDLToRealPIDL(IParent, pIDLToParent)
          #End If
          needRefresh = HandleAutoUpdate_FileCreation(IParent, pIDLToParent)
          #If Debuging Then
            FreeItemIDList debugger, "HandleSHNotify #1", pIDLToParent
          #Else
            FreeItemIDList pIDLToParent
          #End If
          Set IParent = Nothing
        End If
      End If

    Case SHCNEConstants.SHCNE_DELETE
      ' eine Datei wurde gelöscht
      CopyMemory VarPtr(SHNotData), wParam, LenB(SHNotData)
      #If Debuging Then
        debugger.AddLogEntry "   AutoUpdate: A file has been deleted.", LogEntryTypeConstants.letInfo
        debugger.AddLogEntry "      " & pIDLToPath_Light(debugger, SHNotData.dwItem1), LogEntryTypeConstants.letOther
      #End If
      needRefresh = HandleAutoUpdate_FileDeletion(SHNotData.dwItem1)
      If SHNotData.dwItem2 Then
        #If Debuging Then
          SplitFullyQualifiedPIDL debugger, SHNotData.dwItem2, IParent, pIDLToParent
        #Else
          SplitFullyQualifiedPIDL SHNotData.dwItem2, IParent, pIDLToParent
        #End If
        pIDL = GetpIDL(IParent)
        If pIDL = 0 Then
          pIDL = ILClone(SHNotData.dwItem2)
          MakeParentItemIDList pIDL
        End If
        needRefresh = needRefresh Or HandleAutoUpdate_UpdatedRecycler(pIDL, IParent, True)
        If pIDLToParent <> SHNotData.dwItem2 Then
          #If Debuging Then
            FreeItemIDList debugger, "HandleSHNotify #2", pIDLToParent
          #Else
            FreeItemIDList pIDLToParent
          #End If
        End If
        #If Debuging Then
          FreeItemIDList debugger, "HandleSHNotify #3", pIDL
        #Else
          FreeItemIDList pIDL
        #End If
        Set IParent = Nothing
      End If

    Case SHCNEConstants.SHCNE_DRIVEADD
      ' ein Laufwerk wurde hinzugefügt
      ' <SHNotData.dwItem1> ist die pIDL des neuen Items (wurde das Laufwerk über den Gerätemanager
      ' entfernt, ist es etwas Anderes)
      CopyMemory VarPtr(SHNotData), wParam, LenB(SHNotData)
      #If Debuging Then
        debugger.AddLogEntry "   AutoUpdate: A drive has been added.", LogEntryTypeConstants.letInfo
        debugger.AddLogEntry "      " & pIDLToPath_Light(debugger, SHNotData.dwItem1), LogEntryTypeConstants.letOther
      #End If
      #If Debuging Then
        SplitFullyQualifiedPIDL debugger, SHNotData.dwItem1, IParent, pIDLToParent
        If pIDLToParent <> 0 And Not (IParent Is Nothing) Then
          pIDLToParent = SimplePIDLToRealPIDL(debugger, IParent, pIDLToParent)
      #Else
        SplitFullyQualifiedPIDL SHNotData.dwItem1, IParent, pIDLToParent
        If pIDLToParent <> 0 And Not (IParent Is Nothing) Then
          pIDLToParent = SimplePIDLToRealPIDL(IParent, pIDLToParent)
      #End If
        needRefresh = HandleAutoUpdate_AddedDrive(IParent, pIDLToParent)
        #If Debuging Then
          FreeItemIDList debugger, "HandleSHNotify #4", pIDLToParent
        #Else
          FreeItemIDList pIDLToParent
        #End If
        Set IParent = Nothing
      End If

    Case SHCNEConstants.SHCNE_DRIVEADDGUI
      ' ein Laufwerk wurde hinzugefügt oder ein Datenträger wurde eingelegt
      '--------------------------------------------------------------------------------------
      ' möglicherweise wird dieses Event ausgelöst, um zu sagen: "Ein Laufwerk wurde
      ' hinzugefügt/ein Datenträger eingelegt und dies sollte der User auch sehen." (->
      ' DisplayName und Icon ändern sich)
      '--------------------------------------------------------------------------------------
      ' <SHNotData.dwItem1> ist die pIDL des neuen Items
      ' keine weiteren Aktionen nötig
      #If Debuging Then
        debugger.AddLogEntry "   AutoUpdate: A drive has been added or a media has been inserted.", LogEntryTypeConstants.letInfo
        CopyMemory VarPtr(SHNotData), wParam, LenB(SHNotData)
        debugger.AddLogEntry "      " & pIDLToPath_Light(debugger, SHNotData.dwItem1), LogEntryTypeConstants.letOther
      #End If

    Case SHCNEConstants.SHCNE_DRIVEREMOVED
      ' ein Laufwerk wurde entfernt
      ' <SHNotData.dwItem1> ist die pIDL (wurde das Laufwerk über den Gerätemanager entfernt,
      ' ist es etwas Anderes)
      CopyMemory VarPtr(SHNotData), wParam, LenB(SHNotData)
      #If Debuging Then
        debugger.AddLogEntry "   AutoUpdate: A drive has been removed.", LogEntryTypeConstants.letInfo
        debugger.AddLogEntry "      " & pIDLToPath_Light(debugger, SHNotData.dwItem1), LogEntryTypeConstants.letOther
      #End If
      #If Debuging Then
        SplitFullyQualifiedPIDL debugger, SHNotData.dwItem1, IParent, pIDLToParent
        If pIDLToParent <> 0 And Not (IParent Is Nothing) Then
          pIDLToParent = SimplePIDLToRealPIDL(debugger, IParent, pIDLToParent)
      #Else
        SplitFullyQualifiedPIDL SHNotData.dwItem1, IParent, pIDLToParent
        If pIDLToParent <> 0 And Not (IParent Is Nothing) Then
          pIDLToParent = SimplePIDLToRealPIDL(IParent, pIDLToParent)
      #End If
        needRefresh = HandleAutoUpdate_RemovedDrive(IParent, pIDLToParent)
        #If Debuging Then
          FreeItemIDList debugger, "HandleSHNotify #5", pIDLToParent
        #Else
          FreeItemIDList pIDLToParent
        #End If
        Set IParent = Nothing
      End If

    Case SHCNEConstants.SHCNE_FREESPACE
      ' auf mindestens einem Laufwerk hat sich die Menge des freien Speicherplatzes verändert
      ' <SHNotData.dwItem1> ist eine Struktur der Form:
      '   Type UIData
      '     Size As Integer
      '     BitFieldChangedDrives As Long
      '     Dummy As Long    ' immer 0
      '   End Type
      ' keine weiteren Aktionen nötig
      #If Debuging Then
        CopyMemory VarPtr(SHNotData), wParam, LenB(SHNotData)
        debugger.AddLogEntry "   AutoUpdate: The free space on one or more drives has changed.", LogEntryTypeConstants.letInfo
      #End If

    Case SHCNEConstants.SHCNE_MEDIAINSERTED
      ' ein Medium (z.B. CD/DVD) wurde eingelegt
      ' <SHNotData.dwItem1> ist die pIDL
      CopyMemory VarPtr(SHNotData), wParam, LenB(SHNotData)
      #If Debuging Then
        debugger.AddLogEntry "   AutoUpdate: A media has been inserted.", LogEntryTypeConstants.letInfo
        debugger.AddLogEntry "      " & pIDLToPath_Light(debugger, SHNotData.dwItem1), LogEntryTypeConstants.letOther
      #End If
      needRefresh = HandleAutoUpdate_InsertedMedia(SHNotData.dwItem1)

    Case SHCNEConstants.SHCNE_MEDIAREMOVED
      ' ein Medium (z.B. CD/DVD) wurde entnommen
      ' <SHNotData.dwItem1> ist die pIDL
      CopyMemory VarPtr(SHNotData), wParam, LenB(SHNotData)
      #If Debuging Then
        debugger.AddLogEntry "   AutoUpdate: A media has been removed.", LogEntryTypeConstants.letInfo
        debugger.AddLogEntry "      " & pIDLToPath_Light(debugger, SHNotData.dwItem1), LogEntryTypeConstants.letOther
      #End If
      needRefresh = HandleAutoUpdate_RemovedMedia(SHNotData.dwItem1)

    Case SHCNEConstants.SHCNE_MKDIR
      ' ein Ordner wurde erstellt
      CopyMemory VarPtr(SHNotData), wParam, LenB(SHNotData)
      #If Debuging Then
        debugger.AddLogEntry "   AutoUpdate: A folder has been created.", LogEntryTypeConstants.letInfo
        debugger.AddLogEntry "      " & pIDLToPath_Light(debugger, SHNotData.dwItem1), LogEntryTypeConstants.letOther
      #End If
      #If Debuging Then
        SplitFullyQualifiedPIDL debugger, SHNotData.dwItem1, IParent, pIDLToParent
        If pIDLToParent <> 0 And Not (IParent Is Nothing) Then
          pIDLToParent = SimplePIDLToRealPIDL(debugger, IParent, pIDLToParent)
      #Else
        SplitFullyQualifiedPIDL SHNotData.dwItem1, IParent, pIDLToParent
        If pIDLToParent <> 0 And Not (IParent Is Nothing) Then
          pIDLToParent = SimplePIDLToRealPIDL(IParent, pIDLToParent)
      #End If
        needRefresh = HandleAutoUpdate_FolderCreation(IParent, pIDLToParent)
        #If Debuging Then
          FreeItemIDList debugger, "HandleSHNotify #6", pIDLToParent
        #Else
          FreeItemIDList pIDLToParent
        #End If
        Set IParent = Nothing
      End If

    Case SHCNEConstants.SHCNE_NETSHARE
      ' die Freigabe eines (FS-?)Items wurde geändert
      ' <SHNotData.dwItem1> ist die pIDL des Items
      CopyMemory VarPtr(SHNotData), wParam, LenB(SHNotData)
      #If Debuging Then
        debugger.AddLogEntry "   AutoUpdate: A share has been changed.", LogEntryTypeConstants.letInfo
        debugger.AddLogEntry "      " & pIDLToPath_Light(debugger, SHNotData.dwItem1), LogEntryTypeConstants.letOther
      #End If
      needRefresh = HandleAutoUpdate_ChangedSharing(SHNotData.dwItem1)

    Case SHCNEConstants.SHCNE_NETUNSHARE
      ' die Freigabe eines Items wurde per Computerverwaltung aufgehoben
      ' <SHNotData.dwItem1> ist die pIDL des Items
      CopyMemory VarPtr(SHNotData), wParam, LenB(SHNotData)
      #If Debuging Then
        debugger.AddLogEntry "   AutoUpdate: A share has been removed.", LogEntryTypeConstants.letInfo
        debugger.AddLogEntry "      " & pIDLToPath_Light(debugger, SHNotData.dwItem1), LogEntryTypeConstants.letOther
      #End If
      needRefresh = HandleAutoUpdate_ChangedSharing(SHNotData.dwItem1)

    Case SHCNEConstants.SHCNE_RENAMEFOLDER
      ' ein Ordner wurde umbenannt oder verschoben
      ' <SHNotData.dwItem1> ist die alte pIDL,
      ' <SHNotData.dwItem2> ist die neue pIDL
      CopyMemory VarPtr(SHNotData), wParam, LenB(SHNotData)

      ' prüfen, ob eine Verschiebung vorliegt
      ' -> die Parent-Items wären dann verschieden
      #If Debuging Then
        SplitFullyQualifiedPIDL debugger, SHNotData.dwItem1, IParent, pIDLToParent
      #Else
        SplitFullyQualifiedPIDL SHNotData.dwItem1, IParent, pIDLToParent
      #End If
      pIDLParent_ToDesktop1 = GetpIDL(IParent)
      If pIDLParent_ToDesktop1 = 0 Then
        pIDLParent_ToDesktop1 = ILClone(SHNotData.dwItem1)
        MakeParentItemIDList pIDLParent_ToDesktop1
      End If
      Set IParent = Nothing
      #If Debuging Then
        SplitFullyQualifiedPIDL debugger, SHNotData.dwItem2, IParent, pIDLToParent
      #Else
        SplitFullyQualifiedPIDL SHNotData.dwItem2, IParent, pIDLToParent
      #End If
      pIDLParent_ToDesktop2 = GetpIDL(IParent)
      If pIDLParent_ToDesktop2 = 0 Then
        pIDLParent_ToDesktop2 = ILClone(SHNotData.dwItem2)
        MakeParentItemIDList pIDLParent_ToDesktop2
      End If

      If ILIsEqual(pIDLParent_ToDesktop1, pIDLParent_ToDesktop2) Then
        ' eine Umbenennung
        #If Debuging Then
          debugger.AddLogEntry "   AutoUpdate: A folder has been renamed.", LogEntryTypeConstants.letInfo
          debugger.AddLogEntry "      from: " & pIDLToPath_Light(debugger, SHNotData.dwItem1), LogEntryTypeConstants.letOther
          debugger.AddLogEntry "        to: " & pIDLToPath_Light(debugger, SHNotData.dwItem2), LogEntryTypeConstants.letOther
        #End If
        needRefresh = handleAutoUpdate_FolderRenaming(SHNotData, pIDLParent_ToDesktop1, pIDLParent_ToDesktop2)
      Else
        ' eine Verschiebung
        #If Debuging Then
          debugger.AddLogEntry "   AutoUpdate: A folder has been moved.", LogEntryTypeConstants.letInfo
          debugger.AddLogEntry "      from: " & pIDLToPath_Light(debugger, SHNotData.dwItem1), LogEntryTypeConstants.letOther
          debugger.AddLogEntry "        to: " & pIDLToPath_Light(debugger, SHNotData.dwItem2), LogEntryTypeConstants.letOther
        #End If
        If pIDLToParent <> 0 And Not (IParent Is Nothing) Then
          #If Debuging Then
            pIDLToParent = SimplePIDLToRealPIDL(debugger, IParent, pIDLToParent)
          #Else
            pIDLToParent = SimplePIDLToRealPIDL(IParent, pIDLToParent)
          #End If
          needRefresh = HandleAutoUpdate_FolderMoving(SHNotData, IParent, pIDLToParent)
          #If Debuging Then
            FreeItemIDList debugger, "HandleSHNotify #7", pIDLToParent
          #Else
            FreeItemIDList pIDLToParent
          #End If
        End If
      End If
      Set IParent = Nothing

      #If Debuging Then
        FreeItemIDList debugger, "HandleSHNotify #8", pIDLParent_ToDesktop1
        FreeItemIDList debugger, "HandleSHNotify #9", pIDLParent_ToDesktop2
      #Else
        FreeItemIDList pIDLParent_ToDesktop1
        FreeItemIDList pIDLParent_ToDesktop2
      #End If

    Case SHCNEConstants.SHCNE_RENAMEITEM
      ' eine Datei oder ein Archiv (unter WinMe/WinXP) wurde umbenannt oder verschoben oder
      ' ein Drucker wurde umbenannt
      ' <SHNotData.dwItem1> ist die alte pIDL,
      ' <SHNotData.dwItem2> ist die neue pIDL
      CopyMemory VarPtr(SHNotData), wParam, LenB(SHNotData)

      ' prüfen, ob eine Verschiebung vorliegt
      ' -> die Parent-Items wären dann verschieden
      #If Debuging Then
        SplitFullyQualifiedPIDL debugger, SHNotData.dwItem1, IParent, pIDLToParent
      #Else
        SplitFullyQualifiedPIDL SHNotData.dwItem1, IParent, pIDLToParent
      #End If
      pIDLParent_ToDesktop1 = GetpIDL(IParent)
      If pIDLParent_ToDesktop1 = 0 Then
        pIDLParent_ToDesktop1 = ILClone(SHNotData.dwItem1)
        MakeParentItemIDList pIDLParent_ToDesktop1
      End If
      Set IParent = Nothing
      #If Debuging Then
        SplitFullyQualifiedPIDL debugger, SHNotData.dwItem2, IParent, pIDLToParent
      #Else
        SplitFullyQualifiedPIDL SHNotData.dwItem2, IParent, pIDLToParent
      #End If
      pIDLParent_ToDesktop2 = GetpIDL(IParent)
      If pIDLParent_ToDesktop2 = 0 Then
        pIDLParent_ToDesktop2 = ILClone(SHNotData.dwItem2)
        MakeParentItemIDList pIDLParent_ToDesktop2
      End If

      If ILIsEqual(pIDLParent_ToDesktop1, pIDLParent_ToDesktop2) Then
        ' eine Umbenennung
        #If Debuging Then
          debugger.AddLogEntry "   AutoUpdate: A file or a printer has been renamed.", LogEntryTypeConstants.letInfo
          debugger.AddLogEntry "      from: " & pIDLToPath_Light(debugger, SHNotData.dwItem1), LogEntryTypeConstants.letOther
          debugger.AddLogEntry "        to: " & pIDLToPath_Light(debugger, SHNotData.dwItem2), LogEntryTypeConstants.letOther
        #End If
        needRefresh = HandleAutoUpdate_FileRenaming(SHNotData, pIDLParent_ToDesktop1, pIDLParent_ToDesktop2)
      Else
        ' eine Verschiebung
        #If Debuging Then
          debugger.AddLogEntry "   AutoUpdate: A file has been moved.", LogEntryTypeConstants.letInfo
          debugger.AddLogEntry "      from: " & pIDLToPath_Light(debugger, SHNotData.dwItem1), LogEntryTypeConstants.letOther
          debugger.AddLogEntry "        to: " & pIDLToPath_Light(debugger, SHNotData.dwItem2), LogEntryTypeConstants.letOther
        #End If
        If pIDLToParent <> 0 And Not (IParent Is Nothing) Then
          #If Debuging Then
            pIDLToParent = SimplePIDLToRealPIDL(debugger, IParent, pIDLToParent)
          #Else
            pIDLToParent = SimplePIDLToRealPIDL(IParent, pIDLToParent)
          #End If
          needRefresh = HandleAutoUpdate_FileMoving(SHNotData, IParent, pIDLToParent)
          #If Debuging Then
            FreeItemIDList debugger, "HandleSHNotify #10", pIDLToParent
          #Else
            FreeItemIDList pIDLToParent
          #End If
        End If
      End If
      Set IParent = Nothing

      #If Debuging Then
        FreeItemIDList debugger, "HandleSHNotify #11", pIDLParent_ToDesktop1
        FreeItemIDList debugger, "HandleSHNotify #12", pIDLParent_ToDesktop2
      #Else
        FreeItemIDList pIDLParent_ToDesktop1
        FreeItemIDList pIDLParent_ToDesktop2
      #End If

    Case SHCNEConstants.SHCNE_RMDIR
      ' ein Ordner wurde gelöscht
      CopyMemory VarPtr(SHNotData), wParam, LenB(SHNotData)
      #If Debuging Then
        debugger.AddLogEntry "   AutoUpdate: A folder has been removed.", LogEntryTypeConstants.letInfo
        debugger.AddLogEntry "      " & pIDLToPath_Light(debugger, SHNotData.dwItem1), LogEntryTypeConstants.letOther
        If SHNotData.dwItem2 Then
          debugger.AddLogEntry "      to: " & pIDLToPath_Light(debugger, SHNotData.dwItem2), LogEntryTypeConstants.letOther
        End If
      #End If
      needRefresh = HandleAutoUpdate_FolderDeletion(SHNotData.dwItem1)
      If SHNotData.dwItem2 Then
        #If Debuging Then
          SplitFullyQualifiedPIDL debugger, SHNotData.dwItem2, IParent, pIDLToParent
        #Else
          SplitFullyQualifiedPIDL SHNotData.dwItem2, IParent, pIDLToParent
        #End If
        pIDL = GetpIDL(IParent)
        If pIDL = 0 Then
          pIDL = ILClone(SHNotData.dwItem2)
          MakeParentItemIDList pIDL
        End If
        needRefresh = needRefresh Or HandleAutoUpdate_UpdatedRecycler(pIDL, IParent, True)
        If pIDLToParent <> SHNotData.dwItem2 Then
          #If Debuging Then
            FreeItemIDList debugger, "HandleSHNotify #13", pIDLToParent
          #Else
            FreeItemIDList pIDLToParent
          #End If
        End If
        #If Debuging Then
          FreeItemIDList debugger, "HandleSHNotify #14", pIDL
        #Else
          FreeItemIDList pIDL
        #End If
        Set IParent = Nothing
      End If

    Case SHCNEConstants.SHCNE_SERVERDISCONNECT
      #If Debuging Then
        debugger.AddLogEntry "   AutoUpdate: A server connection has been disconnected.", LogEntryTypeConstants.letInfo
        CopyMemory VarPtr(SHNotData), wParam, LenB(SHNotData)
        debugger.AddLogEntry "      " & pIDLToDisplayName_Light(debugger, SHNotData.dwItem1), LogEntryTypeConstants.letOther
      #End If
      ' <SHNotData.dwItem1> ist eine pIDL

    Case SHCNEConstants.SHCNE_UPDATEDIR
      ' der Inhalt eines Ordners wurde allgemein verändert, aber nicht umbenannt (hoffentlich!)
      ' <SHNotData.dwItem1> ist die pIDL des Ordners
      CopyMemory VarPtr(SHNotData), wParam, LenB(SHNotData)
      #If Debuging Then
        debugger.AddLogEntry "   AutoUpdate: A folder's content has changed.", LogEntryTypeConstants.letInfo
        debugger.AddLogEntry "      " & pIDLToPath_Light(debugger, SHNotData.dwItem1), LogEntryTypeConstants.letOther
      #End If

      #If Debuging Then
        If IsPartOfRecycler(debugger, SHNotData.dwItem1, isRecycler) Then
      #Else
        If IsPartOfRecycler(SHNotData.dwItem1, isRecycler) Then
      #End If
        If isRecycler Then
          Set IParent = Nothing
          #If Debuging Then
            debugger.AddLogEntry "HandleSHNotify/SHCNE_UPDATEDIR: Calling GetISHFolderInterfaceFQ()", LogEntryTypeConstants.letInfo
            debugger.AddLogEntry "   0x" & Hex(SHNotData.dwItem1), LogEntryTypeConstants.letOther
            Set IParent = GetISHFolderInterfaceFQ(debugger, SHNotData.dwItem1)
          #Else
            Set IParent = GetISHFolderInterfaceFQ(SHNotData.dwItem1)
          #End If
          needRefresh = HandleAutoUpdate_UpdatedRecycler(SHNotData.dwItem1, IParent)
        Else
          #If Debuging Then
            SplitFullyQualifiedPIDL debugger, SHNotData.dwItem1, IParent, pIDLToParent
          #Else
            SplitFullyQualifiedPIDL SHNotData.dwItem1, IParent, pIDLToParent
          #End If
          pIDL = GetpIDL(IParent)
          If pIDL = 0 Then
            pIDL = ILClone(SHNotData.dwItem1)
            MakeParentItemIDList pIDL
          End If
          needRefresh = HandleAutoUpdate_UpdatedRecycler(pIDL, IParent)
          #If Debuging Then
            FreeItemIDList debugger, "HandleSHNotify #15", pIDL
          #Else
            FreeItemIDList pIDL
          #End If
        End If
      Else
        #If Debuging Then
          SplitFullyQualifiedPIDL debugger, SHNotData.dwItem1, IParent, pIDLToParent
          If pIDLToParent <> 0 And Not (IParent Is Nothing) Then
            pIDLToParent = SimplePIDLToRealPIDL(debugger, IParent, pIDLToParent)
          End If
        #Else
          SplitFullyQualifiedPIDL SHNotData.dwItem1, IParent, pIDLToParent
          If pIDLToParent <> 0 And Not (IParent Is Nothing) Then
            pIDLToParent = SimplePIDLToRealPIDL(IParent, pIDLToParent)
          End If
        #End If
        pIDL = GetpIDL(IParent)
        If pIDL = 0 Then
          pIDL = ILClone(SHNotData.dwItem1)
          MakeParentItemIDList pIDL
        End If
        If pIDLToParent Then
          pIDL2 = ILAppendID(ILClone(pIDL), pIDLToParent, 1)
          needRefresh = HandleAutoUpdate_UpdatedFolder(pIDL2)
          #If Debuging Then
            FreeItemIDList debugger, "HandleSHNotify #16", pIDLToParent
          #Else
            FreeItemIDList pIDLToParent
          #End If
        End If
        #If Debuging Then
          FreeItemIDList debugger, "HandleSHNotify #17", pIDL
          FreeItemIDList debugger, "HandleSHNotify #18", pIDL2
        #Else
          FreeItemIDList pIDL
          FreeItemIDList pIDL2
        #End If
      End If
      Set IParent = Nothing

    Case SHCNEConstants.SHCNE_UPDATEIMAGE
      ' ein Item hat ein anderes Icon der SystemImageList bekommen oder ein OverlayIcon hat
      ' sich geändert
      ' <SHNotData.dwItem1> ist eine Struktur der Form:
      '   Type UIData
      '     Size As Integer
      '     OldIconIndex As Long
      '     ? As Long
      '   End Type
      ' <SHNotData.dwItem2> ist eine Struktur? der Form:
      '   Type UIData
      '     ? As Long?
      '     ? As Integer?
      '     OldIconIndex As Long
      '   End Type
      CopyMemory VarPtr(SHNotData), wParam, LenB(SHNotData)

      CopyMemory VarPtr(oldIconIndex), SHNotData.dwItem1 + 2, LenB(oldIconIndex)
      ' geändertes Icon
      #If Debuging Then
        debugger.AddLogEntry "   AutoUpdate: An icon has been changed.", LogEntryTypeConstants.letInfo
        debugger.AddLogEntry "      " & oldIconIndex, LogEntryTypeConstants.letOther
      #End If
      needRefresh = HandleAutoUpdate_ChangedIcon(oldIconIndex, SHNotData.dwItem2)

    Case SHCNEConstants.SHCNE_UPDATEITEM
      ' ein FS-Item wurde allgemein verändert, aber nicht umbenannt (hoffentlich!) oder ein
      ' Objekt wurde umbenannt oder die Freigabe für einen Drucker wurde geändert
      ' <SHNotData.dwItem1> ist die (neue) pIDL
      CopyMemory VarPtr(SHNotData), wParam, LenB(SHNotData)

      #If Debuging Then
        If IsPartOfRecycler(debugger, SHNotData.dwItem1, isRecycler) Then
      #Else
        If IsPartOfRecycler(SHNotData.dwItem1, isRecycler) Then
      #End If
        If isRecycler Then
          #If Debuging Then
            debugger.AddLogEntry "HandleSHNotify/SHCNE_UPDATEITEM: Calling GetISHFolderInterfaceFQ()", LogEntryTypeConstants.letInfo
            debugger.AddLogEntry "   0x" & Hex(SHNotData.dwItem1), LogEntryTypeConstants.letOther
            Set IParent = GetISHFolderInterfaceFQ(debugger, SHNotData.dwItem1)
          #Else
            Set IParent = GetISHFolderInterfaceFQ(SHNotData.dwItem1)
          #End If
          needRefresh = HandleAutoUpdate_UpdatedRecycler(SHNotData.dwItem1, IParent)
        Else
          #If Debuging Then
            SplitFullyQualifiedPIDL debugger, SHNotData.dwItem1, IParent, pIDLToParent
          #Else
            SplitFullyQualifiedPIDL SHNotData.dwItem1, IParent, pIDLToParent
          #End If
          pIDL = GetpIDL(IParent)
          If pIDL = 0 Then
            pIDL = ILClone(SHNotData.dwItem1)
            MakeParentItemIDList pIDL
          End If
          needRefresh = HandleAutoUpdate_UpdatedRecycler(pIDL, IParent)
          #If Debuging Then
            FreeItemIDList debugger, "HandleSHNotify #19", pIDL
          #Else
            FreeItemIDList pIDL
          #End If
        End If
        Set IParent = Nothing
        dontChangeRecyclerChildrenProp = True
      Else
        #If Debuging Then
          SplitFullyQualifiedPIDL debugger, SHNotData.dwItem1, IParent, pIDLToParent
        #Else
          SplitFullyQualifiedPIDL SHNotData.dwItem1, IParent, pIDLToParent
        #End If
        If (pIDLToParent <> 0) And Not (IParent Is Nothing) Then
          If IsPartOfFileSystem(IParent, pIDLToParent) Then
            ' aktualisierter FS-Item
            #If Debuging Then
              debugger.AddLogEntry "   AutoUpdate: A FS item has been changed.", LogEntryTypeConstants.letInfo
              debugger.AddLogEntry "      " & pIDLToPath_Light(debugger, SHNotData.dwItem1), LogEntryTypeConstants.letOther
            #End If
            #If Debuging Then
              pIDLToParent = SimplePIDLToRealPIDL(debugger, IParent, pIDLToParent)
            #Else
              pIDLToParent = SimplePIDLToRealPIDL(IParent, pIDLToParent)
            #End If
            pIDL = GetpIDL(IParent)
            If pIDL = 0 Then
              pIDL = ILClone(SHNotData.dwItem1)
              MakeParentItemIDList pIDL
            End If
            If pIDLToParent Then
              pIDL = ILAppendID(pIDL, pIDLToParent, 1)
              needRefresh = handleAutoUpdate_UpdatedFSItem(pIDL)
              #If Debuging Then
                FreeItemIDList debugger, "HandleSHNotify #20", pIDL
                FreeItemIDList debugger, "HandleSHNotify #21", pIDLToParent
              #Else
                FreeItemIDList pIDL
                FreeItemIDList pIDLToParent
              #End If
            End If
          #If Debuging Then
            ElseIf IsPrinter_Light(debugger, SHNotData.dwItem1) Then
              ' Drucker mit geänderter Freigabe
              debugger.AddLogEntry "   AutoUpdate: A printer share has been changed.", LogEntryTypeConstants.letInfo
              debugger.AddLogEntry "      " & pIDLToDisplayName_Light(debugger, SHNotData.dwItem1), LogEntryTypeConstants.letOther
          #Else
            ElseIf IsPrinter_Light(SHNotData.dwItem1) Then
              ' Drucker mit geänderter Freigabe
          #End If
            needRefresh = handleAutoUpdate_ChangedPrinterSharing(SHNotData.dwItem1)
          Else
            ' umbenannter NonFS-Item oder erstellte/getrennte DialUp-Verbindung
            #If Debuging Then
              debugger.AddLogEntry "   AutoUpdate: A NonFS item has been renamed or a dial-up connection has been established or disconnected.", LogEntryTypeConstants.letInfo
              debugger.AddLogEntry "      " & pIDLToDisplayName_Light(debugger, SHNotData.dwItem1), LogEntryTypeConstants.letOther
            #End If
            needRefresh = handleAutoUpdate_RenamedNonFSItem(SHNotData.dwItem1)
          End If
        End If
        Set IParent = Nothing
      End If

    Case Else
      #If Debuging Then
        debugger.AddLogEntry "   AutoUpdate: Unknown event 0x" & Hex(lParam), LogEntryTypeConstants.letWarning
      #End If
  End Select

  If needRefresh Then Me.Refresh
End Sub

' füllt das Control
Private Sub InitTreeView()
  Dim DispName As String
  Dim hRootItem As Long
  Dim IParent As IVBShellFolder
  Dim IShIconOverlay As IVBShellIconOverlay
  Dim itemAttr As SFGAOConstants
  Dim overlayIndex As Long
  Dim pIDLToParent As Long
  Dim Size_LargeIcons As Long
  Dim Size_SmallIcons As Long
  Dim useStandardIcon As Boolean

  If disableLoadFileSystem Then Exit Sub

  ' das Dateisystem wird erst zur Runtime geladen (wegen Performance und SubClassing)
  If inIDE Then
    LoadIDEItems
    Exit Sub
  End If

  #If Debuging Then
    debugger.AddLogEntry "Entered InitTreeView()", LogEntryTypeConstants.letInfo
  #End If

  ' TreeView leeren
  ClearTVWItems

  If propShowRoot Then RaiseItemLoadingSubItems -1

  If propAlwaysUseSmallIcons Then
    useLargeIcons = False
  Else
    Size_LargeIcons = GetSystemMetrics(SM_CXICON)
    Size_SmallIcons = GetSystemMetrics(SM_CXSMICON)
    useLargeIcons = (propItemHeight >= Size_LargeIcons * 0.75)
  End If

  If propShowIcons Then
    If Not propUseCustomIcons Then
      ' ImageList setzen
      prophImageList = GetSysImageList(useLargeIcons)
      UpdateDefaultIconIndices
    End If
    SendMessageAsLong prophWnd, TVM_SETIMAGELIST, TVSIL_NORMAL, prophImageList
  Else
    ' ImageList "befreien"
    SendMessageAsLong prophWnd, TVM_SETIMAGELIST, TVSIL_NORMAL, 0
    prophImageList = 0
  End If
  If propCheckBoxes Then
    If IIf(useLargeIcons, prophLargeStateImageList, prophSmallStateImageList) Then
      SendMessageAsLong prophWnd, TVM_SETIMAGELIST, TVSIL_STATE, IIf(useLargeIcons, prophLargeStateImageList, prophSmallStateImageList)
    End If
  End If

  ' pIDL des RootPaths ermitteln
  RootItemInfo.pIDLToDesktop = UserPathTopIDL(propRootPath)
  If RootItemInfo.pIDLToDesktop = 0 Then RootItemInfo.pIDLToDesktop = ILClone(pIDL_Desktop)

  ' <IRoot> setzen
  Set IRoot = Nothing
  #If Debuging Then
    debugger.AddLogEntry "InitTreeView: Calling GetISHFolderInterfaceFQ()", LogEntryTypeConstants.letInfo
    debugger.AddLogEntry "   0x" & Hex(RootItemInfo.pIDLToDesktop), LogEntryTypeConstants.letOther
    Set IRoot = GetISHFolderInterfaceFQ(debugger, RootItemInfo.pIDLToDesktop)
    SplitFullyQualifiedPIDL debugger, RootItemInfo.pIDLToDesktop, IParent, pIDLToParent
  #Else
    Set IRoot = GetISHFolderInterfaceFQ(RootItemInfo.pIDLToDesktop)
    SplitFullyQualifiedPIDL RootItemInfo.pIDLToDesktop, IParent, pIDLToParent
  #End If

  ' Typ des RootItems ermitteln
  #If Debuging Then
    itemAttr = GetAttributes(debugger, IParent, pIDLToParent, SFGAOConstants.SFGAO_FILESYSTEM Or SFGAOConstants.SFGAO_FOLDER Or SFGAOConstants.SFGAO_SHARE Or SFGAOConstants.SFGAO_LINK Or SFGAOConstants.SFGAO_GHOSTED Or SFGAOConstants.SFGAO_NEWCONTENT)
  #Else
    itemAttr = GetAttributes(IParent, pIDLToParent, SFGAOConstants.SFGAO_FILESYSTEM Or SFGAOConstants.SFGAO_FOLDER Or SFGAOConstants.SFGAO_SHARE Or SFGAOConstants.SFGAO_LINK Or SFGAOConstants.SFGAO_GHOSTED Or SFGAOConstants.SFGAO_NEWCONTENT)
  #End If
  If itemAttr And SFGAOConstants.SFGAO_FILESYSTEM Then
    If itemAttr And SFGAOConstants.SFGAO_FOLDER Then
      ' ist es wirklich ein Ordner?
      #If Debuging Then
        If FileExists_pIDL(debugger, IParent, pIDLToParent) Then
      #Else
        If FileExists_pIDL(IParent, pIDLToParent) Then
      #End If
        RootItemInfo.itemType = CByte(ItemTypeConstants.itFSFile)
      Else
        RootItemInfo.itemType = CByte(ItemTypeConstants.itFSFolder)
      End If
    Else
      RootItemInfo.itemType = CByte(ItemTypeConstants.itFSFile)
    End If
  Else
    If itemAttr And SFGAOConstants.SFGAO_FOLDER Then
      RootItemInfo.itemType = CByte(ItemTypeConstants.itNonFSFolder)
    Else
      RootItemInfo.itemType = CByte(ItemTypeConstants.itNonFSFile)
    End If
  End If

  If (prophImageList <> 0) And Not propUseCustomIcons Then
    ' Overlays laden
    If propShownOverlays <> ShownOverlaysConstants.soNone Then
      On Error Resume Next
      IParent.QueryInterface IID_IShellIconOverlay, IShIconOverlay
      If IShIconOverlay Is Nothing Then
        If itemAttr And SFGAOConstants.SFGAO_SHARE Then overlayIndex = 1
        If itemAttr And SFGAOConstants.SFGAO_LINK Then overlayIndex = 2
      Else
        If IShIconOverlay.GetOverlayIndex(pIDLToParent, overlayIndex) = S_OK Then
          ' jetzt haben wir den Overlay-Index für den Item, aber soll dieser Overlay überhaupt
          ' angezeigt werden?
          If ShouldShowOverlay(propShownOverlays, overlayIndex) Then
            ' ja!
          Else
            ' kein Overlay
            overlayIndex = 0
          End If
        Else
          ' kein Overlay
          overlayIndex = 0
        End If
        Set IShIconOverlay = Nothing
      End If
    End If
  End If

  #If Debuging Then
    DispName = pIDLToDisplayName(debugger, IParent, pIDLToParent, SHGDNConstants.SHGDN_INFOLDER)
    If (prophImageList <> 0) And Not propUseCustomIcons Then
      #If NoCallbacks Then
        useStandardIcon = propUseStandardIcons
        If propUseStandardIconsForSlowItems And Not useStandardIcon Then
          useStandardIcon = IsSlowItem(debugger, IParent, pIDLToParent, RootItemInfo.pIDLToDesktop, True)
        End If
      #End If
    End If
  #Else
    DispName = pIDLToDisplayName(IParent, pIDLToParent, SHGDNConstants.SHGDN_INFOLDER)
    If (prophImageList <> 0) And Not propUseCustomIcons Then
      #If NoCallbacks Then
        useStandardIcon = propUseStandardIcons
        If propUseStandardIconsForSlowItems And Not useStandardIcon Then
          useStandardIcon = IsSlowItem(IParent, pIDLToParent, RootItemInfo.pIDLToDesktop, True)
        End If
      #End If
    End If
  #End If

  If propShowRoot Then
    ' hinzufügen
    hRootItem = InsertTVWItem_Common(DispName, TVI_ROOT, RootItemInfo, overlayIndex, itemAttr And SFGAOConstants.SFGAO_GHOSTED, itemAttr And SFGAOConstants.SFGAO_NEWCONTENT, useStandardIcon)
    #If UseSubClassing Then
      '
    #Else
      ' Childs hinzufügen
      #If Debuging Then
        If HasSubItems(debugger, RootItemInfo.pIDLToDesktop, Me) Then
      #Else
        If HasSubItems(RootItemInfo.pIDLToDesktop, Me) Then
      #End If
        InsertSubItems hRootItem
      End If
    #End If

    RaiseItemLoadedSubItems -1
  Else
    #If Debuging Then
      If HasSubItems(debugger, RootItemInfo.pIDLToDesktop, Me) Then
    #Else
      If HasSubItems(RootItemInfo.pIDLToDesktop, Me) Then
    #End If
      InsertSubItems 0
    End If
  End If
  Set IParent = Nothing

  #If Debuging Then
    debugger.AddLogEntry "Leaving InitTreeView()", LogEntryTypeConstants.letInfo
  #End If
End Sub

' fügt die Ordnerstruktur <FolderPath> als SubItem(s) von <hArchiveItem> ein und gibt das
' Handle des letzten Segments zurück
Private Function InsertArchiveFolder(ByVal hArchiveItem As Long, ByVal FolderPath As String, Optional ByVal hRelativeItem As Long = -2) As Long
  Dim DispName As String
  Dim hasChilds As Boolean
  Dim hItem As Long
  Dim hParentItem As Long
  Dim i As Integer
  Dim ImgNormal As Long
  Dim ImgSelected As Long
  Dim ItemText As String
  Dim ret As Long
  Dim Segment As String
  Dim TVWItemInfo As ITEMINFO_ARCHIVE

  If FolderPath = "" Then
    ret = -1
    GoTo Ende
  End If
  If Not ItemExists(hArchiveItem, AllowNullConstants.anNever) Then
    ret = -1
    GoTo Ende
  End If
  If hRelativeItem = -2 Then hRelativeItem = hArchiveItem
  If Not ItemExists(hRelativeItem, AllowNullConstants.anNever) Then
    ret = -1
    GoTo Ende
  End If

  ' ist immer gleich
  With TVWItemInfo
    .itemType = CByte(ItemTypeConstants.itArchiveFolder)
    .hItem_ArchiveFile = hArchiveItem
    .attributes = FILE_ATTRIBUTE_DIRECTORY
  End With

  ' zunächst alle Segmente von <FolderPath> durchgehen und prüfen, ob sie bereits eingefügt
  ' wurden
  ' dadurch erhalten wir das erste einzufügende Segment
  hParentItem = hRelativeItem
  ' mit dem 1. Child (falls vorhanden) des Archivs anfangen
  hItem = ItemGetFirstSubItem(hRelativeItem)
  While (FolderPath <> "") And (IsWithinRange(hItem, -1, 0) = False)
    ' 1. Segment ermitteln...
    Segment = GetFirstFolders(FolderPath, 1)
    ' ...und im TreeView suchen
    ItemText = GetItemText(hItem)

    If LCase$(ItemText) = LCase$(Segment) Then
      ' gefunden
      hParentItem = hItem
      ' 1. Child ermitteln
      FolderPath = RemoveFirstFolders(FolderPath, 1)
      If FolderPath <> "" Then hItem = ItemGetFirstSubItem(hParentItem)
    Else
      ' nächsten Item auf dieser Ebene ermitteln
      hItem = ItemGetNextItem(hItem)
    End If
  Wend

  ' <hParentItem> ist jetzt das letzte Segment, welches nicht eingefügt werden muß
  If FolderPath <> "" Then
    ' die Icons sind immer gleich
    If (prophImageList <> 0) And Not propUseCustomIcons Then
      ImgNormal = IIf(useLargeIcons, DEFICON_FOLDER_LARGE, DEFICON_FOLDER_SMALL)
      If propUseSelectedIcons Then
        ImgSelected = IIf(useLargeIcons, DEFICON_OPENFOLDER_LARGE, DEFICON_OPENFOLDER_SMALL)
      Else
        ImgSelected = ImgNormal
      End If
    End If

    ' die restlichen Segmente einfügen
    While CountSegments(FolderPath) >= 1
      Segment = GetFirstFolders(FolderPath, 1)
      FolderPath = RemoveFirstFolders(FolderPath, 1)

      DispName = Segment
      If ShouldShowItem_Archive(Me, DispName, TVWItemInfo.attributes) Then
        ret = InsertTVWItem_Archive(DispName, hParentItem, hasChilds, TVWItemInfo, ImgNormal, ImgSelected)
        hParentItem = ret
      Else
        ret = 0
        GoTo Ende
      End If
    Wend
  Else
    ' wir brauchen gar keine Items einfügen
    ret = hParentItem
  End If

Ende:
  If ret = 0 Then ret = -1
  InsertArchiveFolder = ret
End Function

' fügt dem TreeView alle SubItems des Archivs <hParentItem> hinzu
Private Sub InsertArchiveSubItems(ByVal hParentItem As Long)
  Const WM_SETREDRAW = &HB
  Dim ArchivePath As String
  Dim ext As String
  Dim hFile As Long
  Dim hItem As Long
  Dim IParent As IVBShellFolder
  Dim pIDLToParent As Long
  Dim TVWItemInfo_Parent As ITEMINFO_COMMON

  If Not ItemExists(hParentItem, AllowNullConstants.anNever) Then Exit Sub

  ' zunächst alle evt. vorhandenen SubItems entfernen
  RemoveSubItems hParentItem

  GetItemInfo_Common hParentItem, TVWItemInfo_Parent
  #If Debuging Then
    SplitFullyQualifiedPIDL debugger, TVWItemInfo_Parent.pIDLToDesktop, IParent, pIDLToParent
  #Else
    SplitFullyQualifiedPIDL TVWItemInfo_Parent.pIDLToDesktop, IParent, pIDLToParent
  #End If
  If IParent Is Nothing Then Exit Sub
  If pIDLToParent = 0 Then Exit Sub

  #If Debuging Then
    ArchivePath = pIDLToPath(debugger, IParent, pIDLToParent)
  #Else
    ArchivePath = pIDLToPath(IParent, pIDLToParent)
  #End If
  hFile = CreateFileAsLong(ArchivePath, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)

  If hFile <> INVALID_HANDLE_VALUE Then
    ' zunächst den Typ des Archivs ermitteln
    ext = GetFileNameExtension(ArchivePath)
    Select Case LCase$(ext)
      Case "zip"
        If propExpandArchives And ExpandArchivesConstants.eaZIP Then
          ' Archiv lesen
          LoadArchive_ZIP hParentItem, hFile
        End If
      Case "rar"
        If propExpandArchives And ExpandArchivesConstants.eaRAR Then
          ' Archiv lesen
          LoadArchive_RAR hParentItem, hFile
        End If
      Case "iso"
        If propExpandArchives And ExpandArchivesConstants.eaISO Then
          ' Archiv lesen
          LoadArchive_ISO hParentItem, hFile
        End If
      Case "ace"
        If propExpandArchives And ExpandArchivesConstants.eaACE Then
          ' Archiv lesen
          LoadArchive_ACE hParentItem, hFile
        End If
      #If NewArchiveSupport Then
        Case "tar"
          If propExpandArchives And ExpandArchivesConstants.eaTAR Then
            ' Archiv lesen
            LoadArchive_TAR hParentItem, hFile
          End If
      #End If
      Case "bin"
        If propExpandArchives And ExpandArchivesConstants.eaBIN Then
          ' Archiv lesen
          LoadArchive_BIN hParentItem, hFile
        End If
      Case "cab"
        If propExpandArchives And ExpandArchivesConstants.eaCAB Then
          ' Archiv lesen
          LoadArchive_CAB hParentItem, hFile
        End If
      Case "jar"
        If propExpandArchives And ExpandArchivesConstants.eaJAR Then
          ' Archiv lesen
          LoadArchive_ZIP hParentItem, hFile
        End If
    End Select
    CloseHandle hFile

    SortAllSubItems hParentItem
  End If
  Set IParent = Nothing
End Sub

Private Function InsertIDEItem(ByVal DisplayName As String, ByVal hParentItem As Long, ByVal hasChilds As Boolean) As Long
  Dim Data As TVINSERTSTRUCT
  Dim hItem As Long
  Dim ret As Long

  If prophWnd = 0 Then
    ret = -1
    GoTo Ende
  End If
  If Not ItemExists(hParentItem, AllowNullConstants.anAlways, True) Then
    ret = -1
    GoTo Ende
  End If

  SetItemChildrenProperty hParentItem, True

  With Data
    .hParent = hParentItem
    With .itemex
      .cChildren = Abs(hasChilds)
      .pszText = DisplayName
      .cchTextMax = Len(.pszText)
      .Mask = TVIF_CHILDREN Or TVIF_TEXT
    End With
    ret = SendMessage(prophWnd, TVM_INSERTITEM, 0, Data)

    If ret = 0 Then ret = -1
  End With

Ende:
  InsertIDEItem = ret
End Function

Private Sub InsertNewAndRemoveOldSubItems(ByVal hParentItem As Long, Optional ByVal insertNew As Boolean = True, Optional ByVal removeOld As Boolean = True, Optional ByVal alsoUpdateSubItems As Boolean = True)
  Dim cItems As Long
  Dim EnumFlags As SHCONTFConstants
  Dim hasToSort As Boolean
  Dim hSubItem As Long
  Dim i As Long
  Dim IEnum As IVBEnumIDList
  Dim IItem As IVBShellFolder
  Dim items() As ITEMMAPENTRY_COMMON
  Dim items_UBound As Long
  Dim pIDLParent_ToDesktop As Long
  Dim pIDLSubItem_ToDesktop As Long
  Dim pIDLSubItem_ToParent As Long
  Dim shouldShowIt As Boolean
  Dim tmp As Long
  Dim TVWItemInfo_Parent As ITEMINFO_COMMON

  If Not ItemExists(hParentItem, AllowNullConstants.anIfNoShowRoot) Then Exit Sub
  If ItemIsArchiveToExpand(hParentItem) Then Exit Sub

  If Not GetItemInfo_Common(hParentItem, TVWItemInfo_Parent) Then Exit Sub
  pIDLParent_ToDesktop = TVWItemInfo_Parent.pIDLToDesktop
  #If Debuging Then
    debugger.AddLogEntry "InsertNewAndRemoveOldSubItems: Calling GetISHFolderInterfaceFQ()", LogEntryTypeConstants.letInfo
    debugger.AddLogEntry "   pIDLParent_ToDesktop: 0x" & Hex(pIDLParent_ToDesktop), LogEntryTypeConstants.letOther
    Set IItem = GetISHFolderInterfaceFQ(debugger, pIDLParent_ToDesktop)
  #Else
    Set IItem = GetISHFolderInterfaceFQ(pIDLParent_ToDesktop)
  #End If

  blockAutoUpdate = blockAutoUpdate + 1
  If Not (IItem Is Nothing) Then
    If removeOld Then
      #If Debuging Then
        debugger.AddLogEntry "InsertNewAndRemoveOldSubItems: Started removing old items...", LogEntryTypeConstants.letInfo
        debugger.AddLogEntry "   hParentItem: 0x" & Hex(hParentItem), LogEntryTypeConstants.letOther
      #End If
      ' zunächst alle SubItems von <hParentItem> durchgehen und validieren
      If hParentItem Then
        hSubItem = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_CHILD, hParentItem)
      Else
        hSubItem = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_ROOT, 0)
      End If
      While hSubItem <> 0
        ' nächsten Item ermitteln
        tmp = hSubItem
        hSubItem = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_NEXT, hSubItem)

        pIDLSubItem_ToDesktop = GetFullyQualifiedPIDL(tmp, True)
        pIDLSubItem_ToParent = ILFindLastID(pIDLSubItem_ToDesktop)
        #If Debuging Then
          If IsDrive(debugger, IItem, pIDLSubItem_ToParent) Then
            shouldShowIt = ShouldShowItem(debugger, Me, IItem, pIDLSubItem_ToParent, pIDLSubItem_ToDesktop, hParentItem)
          Else
            shouldShowIt = ValidateItem(IItem, pIDLSubItem_ToParent)
            If shouldShowIt Then
              shouldShowIt = ShouldShowItem(debugger, Me, IItem, pIDLSubItem_ToParent, pIDLSubItem_ToDesktop, hParentItem)
            End If
          End If
        #Else
          If IsDrive(IItem, pIDLSubItem_ToParent) Then
            shouldShowIt = ShouldShowItem(Me, IItem, pIDLSubItem_ToParent, pIDLSubItem_ToDesktop, hParentItem)
          Else
            shouldShowIt = ValidateItem(IItem, pIDLSubItem_ToParent)
            If shouldShowIt Then
              shouldShowIt = ShouldShowItem(Me, IItem, pIDLSubItem_ToParent, pIDLSubItem_ToDesktop, hParentItem)
            End If
          End If
        #End If

        If Not shouldShowIt Then
          ' der Item existiert nicht mehr
          RemoveTVWItem tmp
        Else
          If insertNew Then
            cItems = cItems + 1
            If cItems > items_UBound Then
              items_UBound = items_UBound + 100
              ReDim Preserve items(1 To items_UBound) As ITEMMAPENTRY_COMMON
            End If
            With items(cItems)
              .hItem = tmp
              .pIDLToDesktop = pIDLSubItem_ToDesktop
            End With
          End If

          If alsoUpdateSubItems Then
            If ItemHasSubItems(tmp) Then
              InsertNewAndRemoveOldSubItems tmp, insertNew, removeOld, alsoUpdateSubItems
            End If
          End If
        End If
      Wend
      #If Debuging Then
        debugger.AddLogEntry "InsertNewAndRemoveOldSubItems: Ended removing old items", LogEntryTypeConstants.letInfo
        debugger.AddLogEntry "   hParentItem: 0x" & Hex(hParentItem), LogEntryTypeConstants.letOther
      #End If

    ElseIf insertNew Then
      ' build item map
      If hParentItem Then
        hSubItem = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_CHILD, hParentItem)
      Else
        hSubItem = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_ROOT, 0)
      End If
      While hSubItem <> 0
        ' nächsten Item ermitteln
        tmp = hSubItem
        hSubItem = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_NEXT, hSubItem)

        pIDLSubItem_ToDesktop = GetFullyQualifiedPIDL(tmp, True)
        cItems = cItems + 1
        If cItems > items_UBound Then
          items_UBound = items_UBound + 100
          ReDim Preserve items(1 To items_UBound) As ITEMMAPENTRY_COMMON
        End If
        With items(cItems)
          .hItem = tmp
          .pIDLToDesktop = pIDLSubItem_ToDesktop
        End With
      Wend
    End If

    If insertNew Then
      ' jetzt alle SubItems von <IItem> durchgehen und neue hinzufügen
      ' on Vista SHCONTF_DRIVES doesn't work anymore, so let shouldShowItem() do the work
      If propDrivesOnly And Not ver_Win_Vista Then
        #If Debuging Then
          If IsMyComputer(debugger, IItem) Then
        #Else
          If IsMyComputer(IItem) Then
        #End If
          EnumFlags = SHCONTFConstants.SHCONTF_DRIVES
        Else
          EnumFlags = SHCONTFConstants.SHCONTF_FOLDERS Or SHCONTFConstants.SHCONTF_NONFOLDERS
          If propFileAttributes And AttributesConstants.attHidden Then EnumFlags = EnumFlags Or SHCONTFConstants.SHCONTF_INCLUDEHIDDEN
          If propFolderAttributes And AttributesConstants.attHidden Then EnumFlags = EnumFlags Or SHCONTFConstants.SHCONTF_INCLUDEHIDDEN
        End If
      Else
        EnumFlags = SHCONTFConstants.SHCONTF_FOLDERS Or SHCONTFConstants.SHCONTF_NONFOLDERS
        If propFileAttributes And AttributesConstants.attHidden Then EnumFlags = EnumFlags Or SHCONTFConstants.SHCONTF_INCLUDEHIDDEN
        If propFolderAttributes And AttributesConstants.attHidden Then EnumFlags = EnumFlags Or SHCONTFConstants.SHCONTF_INCLUDEHIDDEN
      End If

      IItem.EnumObjects 0, EnumFlags, IEnum
      If Not (IEnum Is Nothing) Then
        #If Debuging Then
          debugger.AddLogEntry "InsertNewAndRemoveOldSubItems: Started inserting new items...", LogEntryTypeConstants.letInfo
          debugger.AddLogEntry "   hParentItem: 0x" & Hex(hParentItem), LogEntryTypeConstants.letOther
          debugger.AddLogEntry "InsertNewAndRemoveOldSubItems: Starting enumeration...", LogEntryTypeConstants.letInfo
        #End If
        While IEnum.Next(1, pIDLSubItem_ToParent, 0) = S_OK
          #If Debuging Then
            debugger.AddLogEntry "InsertNewAndRemoveOldSubItems: Enumerated relative pIDL 0x" & Hex(pIDLSubItem_ToParent), LogEntryTypeConstants.letInfo
          #End If
          If pIDLSubItem_ToParent Then
            pIDLSubItem_ToDesktop = ILAppendID(ILClone(pIDLParent_ToDesktop), pIDLSubItem_ToParent, 1)
            #If Debuging Then
              shouldShowIt = ShouldShowItem(debugger, Me, IItem, pIDLSubItem_ToParent, pIDLSubItem_ToDesktop, hParentItem)
              debugger.AddLogEntry "   ShouldShowItem() returned: " & shouldShowIt, LogEntryTypeConstants.letOther
            #Else
              shouldShowIt = ShouldShowItem(Me, IItem, pIDLSubItem_ToParent, pIDLSubItem_ToDesktop, hParentItem)
            #End If
          Else
            shouldShowIt = False
          End If

          If shouldShowIt Then
            ' Do we already have a tree item for this pIDL?
            #If Debuging Then
              debugger.AddLogEntry "InsertNewAndRemoveOldSubItems: Checking item for existance in treeview", LogEntryTypeConstants.letInfo
              debugger.AddLogEntry "   pIDLSubItem_ToDesktop: 0x" & Hex(pIDLSubItem_ToDesktop), LogEntryTypeConstants.letOther
            #End If

            hSubItem = 0
            For i = 1 To cItems
              If ILIsEqual(items(i).pIDLToDesktop, pIDLSubItem_ToDesktop) Then
                hSubItem = items(i).hItem
                Exit For
              End If
            Next i

            #If Debuging Then
              debugger.AddLogEntry "InsertNewAndRemoveOldSubItems: Checked item for existance in treeview", LogEntryTypeConstants.letInfo
              debugger.AddLogEntry "   hSubItem: 0x" & Hex(hSubItem), LogEntryTypeConstants.letOther
            #End If

            If hSubItem = 0 Then
              If AddSubItem(hParentItem, pIDLSubItem_ToParent, , True, True) <> -1 Then
                hasToSort = True
              End If
            End If
          End If

          #If Debuging Then
            FreeItemIDList debugger, "InsertNewAndRemoveOldSubItems #1", pIDLSubItem_ToDesktop
            FreeItemIDList debugger, "InsertNewAndRemoveOldSubItems #2", pIDLSubItem_ToParent
          #Else
            FreeItemIDList pIDLSubItem_ToDesktop
            FreeItemIDList pIDLSubItem_ToParent
          #End If
        Wend
        #If Debuging Then
          debugger.AddLogEntry "InsertNewAndRemoveOldSubItems: Ended inserting new items", LogEntryTypeConstants.letInfo
          debugger.AddLogEntry "   hParentItem: 0x" & Hex(hParentItem), LogEntryTypeConstants.letOther
        #End If
        Erase items

        If hasToSort Then
          SortItems hParentItem
        End If
      End If
      Set IEnum = Nothing
    End If
  End If
  Set IItem = Nothing
  blockAutoUpdate = blockAutoUpdate - 1
End Sub

' fügt dem TreeView alle SubItems von <RootPath> hinzu, ohne <RootPath> selbst hinzuzufügen
Private Sub InsertRootItems()
  Const WM_SETREDRAW = &HB
  Dim DispName As String
  Dim EnumFlags As SHCONTFConstants
  Dim IEnum As IVBEnumIDList
  Dim IItem As IVBShellFolder
  Dim isFirstItem As Boolean
  Dim IShIconOverlay As IVBShellIconOverlay
  Dim itemAttr As SFGAOConstants
  Dim overlayIndex As Long
  Dim pIDLSubItem_ToDesktop As Long
  Dim pIDLSubItem_ToParent As Long
  Dim shouldShowIt As Boolean
  Dim TVWItemInfo As ITEMINFO_COMMON
  Dim useStandardIcon As Boolean

  If prophWnd = 0 Then Exit Sub

  RaiseItemLoadingSubItems 0

  ' Aufzählung initiieren
  blockAutoUpdate = blockAutoUpdate + 1
  #If Debuging Then
    debugger.AddLogEntry "InsertRootItems: Calling GetISHFolderInterfaceFQ()", LogEntryTypeConstants.letInfo
    debugger.AddLogEntry "   0x" & Hex(RootItemInfo.pIDLToDesktop), LogEntryTypeConstants.letOther
    Set IItem = GetISHFolderInterfaceFQ(debugger, RootItemInfo.pIDLToDesktop)
  #Else
    Set IItem = GetISHFolderInterfaceFQ(RootItemInfo.pIDLToDesktop)
  #End If
  If Not (IItem Is Nothing) Then
    ' on Vista SHCONTF_DRIVES doesn't work anymore, so let shouldShowItem() do the work
    If propDrivesOnly And Not ver_Win_Vista Then
      #If Debuging Then
        If IsMyComputer(debugger, IItem) Then
      #Else
        If IsMyComputer(IItem) Then
      #End If
        EnumFlags = SHCONTFConstants.SHCONTF_DRIVES
      Else
        EnumFlags = SHCONTFConstants.SHCONTF_FOLDERS
        If propIncludedItems And IncludedItemsConstants.iiFSFiles Then EnumFlags = EnumFlags Or SHCONTFConstants.SHCONTF_NONFOLDERS
        If propIncludedItems And IncludedItemsConstants.iiNonFSFiles Then EnumFlags = EnumFlags Or SHCONTFConstants.SHCONTF_NONFOLDERS
        If propFileAttributes And AttributesConstants.attHidden Then EnumFlags = EnumFlags Or SHCONTFConstants.SHCONTF_INCLUDEHIDDEN
        If propFolderAttributes And AttributesConstants.attHidden Then EnumFlags = EnumFlags Or SHCONTFConstants.SHCONTF_INCLUDEHIDDEN
      End If
    Else
      EnumFlags = SHCONTFConstants.SHCONTF_FOLDERS
      If propIncludedItems And IncludedItemsConstants.iiFSFiles Then EnumFlags = EnumFlags Or SHCONTFConstants.SHCONTF_NONFOLDERS
      If propIncludedItems And IncludedItemsConstants.iiNonFSFiles Then EnumFlags = EnumFlags Or SHCONTFConstants.SHCONTF_NONFOLDERS
      If propFileAttributes And AttributesConstants.attHidden Then EnumFlags = EnumFlags Or SHCONTFConstants.SHCONTF_INCLUDEHIDDEN
      If propFolderAttributes And AttributesConstants.attHidden Then EnumFlags = EnumFlags Or SHCONTFConstants.SHCONTF_INCLUDEHIDDEN
    End If

    IItem.EnumObjects prophWndShellUIParentWindow, EnumFlags, IEnum
    If Not (IEnum Is Nothing) Then
      If (prophImageList <> 0) And Not propUseCustomIcons Then
        If propShownOverlays <> ShownOverlaysConstants.soNone Then
          IItem.QueryInterface IID_IShellIconOverlay, IShIconOverlay
        End If
      End If
      isFirstItem = True
      While IEnum.Next(1, pIDLSubItem_ToParent, 0) = S_OK
        If pIDLSubItem_ToParent Then
          pIDLSubItem_ToDesktop = ILAppendID(ILClone(RootItemInfo.pIDLToDesktop), pIDLSubItem_ToParent, 1)
          #If Debuging Then
            shouldShowIt = ShouldShowItem(debugger, Me, IItem, pIDLSubItem_ToParent, pIDLSubItem_ToDesktop, 0)
          #Else
            shouldShowIt = ShouldShowItem(Me, IItem, pIDLSubItem_ToParent, pIDLSubItem_ToDesktop, 0)
          #End If
        Else
          shouldShowIt = False
        End If

        If Not shouldShowIt Then
          #If Debuging Then
            FreeItemIDList debugger, "InsertRootItems #1", pIDLSubItem_ToDesktop
          #Else
            FreeItemIDList pIDLSubItem_ToDesktop
          #End If
        Else
          ' Item anzeigen
          TVWItemInfo.pIDLToDesktop = pIDLSubItem_ToDesktop

          ' Typ des Items ermitteln
          #If Debuging Then
            itemAttr = GetAttributes(debugger, IItem, pIDLSubItem_ToParent, SFGAOConstants.SFGAO_FILESYSTEM Or SFGAOConstants.SFGAO_FOLDER Or SFGAOConstants.SFGAO_SHARE Or SFGAOConstants.SFGAO_LINK Or SFGAOConstants.SFGAO_SHARE Or SFGAOConstants.SFGAO_LINK Or SFGAOConstants.SFGAO_GHOSTED Or SFGAOConstants.SFGAO_NEWCONTENT)
          #Else
            itemAttr = GetAttributes(IItem, pIDLSubItem_ToParent, SFGAOConstants.SFGAO_FILESYSTEM Or SFGAOConstants.SFGAO_FOLDER Or SFGAOConstants.SFGAO_SHARE Or SFGAOConstants.SFGAO_LINK Or SFGAOConstants.SFGAO_SHARE Or SFGAOConstants.SFGAO_LINK Or SFGAOConstants.SFGAO_GHOSTED Or SFGAOConstants.SFGAO_NEWCONTENT)
          #End If
          If itemAttr And SFGAOConstants.SFGAO_FILESYSTEM Then
            If itemAttr And SFGAOConstants.SFGAO_FOLDER Then
              ' ist es wirklich ein Ordner?
              #If Debuging Then
                If FileExists_pIDL(debugger, IItem, pIDLSubItem_ToParent) Then
              #Else
                If FileExists_pIDL(IItem, pIDLSubItem_ToParent) Then
              #End If
                TVWItemInfo.itemType = CByte(ItemTypeConstants.itFSFile)
              Else
                TVWItemInfo.itemType = CByte(ItemTypeConstants.itFSFolder)
              End If
            Else
              TVWItemInfo.itemType = CByte(ItemTypeConstants.itFSFile)
            End If
          Else
            If itemAttr And SFGAOConstants.SFGAO_FOLDER Then
              TVWItemInfo.itemType = CByte(ItemTypeConstants.itNonFSFolder)
            Else
              TVWItemInfo.itemType = CByte(ItemTypeConstants.itNonFSFile)
            End If
          End If

          If (prophImageList <> 0) And Not propUseCustomIcons Then
            ' Overlays laden
            overlayIndex = 0
            If propShownOverlays <> ShownOverlaysConstants.soNone Then
              On Error Resume Next
              If IShIconOverlay Is Nothing Then
                If itemAttr And SFGAOConstants.SFGAO_SHARE Then overlayIndex = 1
                If itemAttr And SFGAOConstants.SFGAO_LINK Then overlayIndex = 2
              Else
                If IShIconOverlay.GetOverlayIndex(pIDLSubItem_ToParent, overlayIndex) = S_OK Then
                  ' jetzt haben wir den Overlay-Index für den Item, aber soll dieser Overlay
                  ' überhaupt angezeigt werden?
                  If ShouldShowOverlay(propShownOverlays, overlayIndex) Then
                    ' ja!
                  Else
                    ' kein Overlay
                    overlayIndex = 0
                  End If
                Else
                  ' kein Overlay
                  overlayIndex = 0
                End If
              End If
            End If
          End If

          #If Debuging Then
            DispName = pIDLToDisplayName(debugger, IItem, pIDLSubItem_ToParent, SHGDNConstants.SHGDN_INFOLDER)
            If (prophImageList <> 0) And Not propUseCustomIcons Then
              #If NoCallbacks Then
                useStandardIcon = propUseStandardIcons
                If propUseStandardIconsForSlowItems And Not useStandardIcon Then
                  useStandardIcon = IsSlowItem(debugger, IItem, pIDLSubItem_ToParent, TVWItemInfo.pIDLToDesktop, True)
                End If
              #End If
            End If
          #Else
            DispName = pIDLToDisplayName(IItem, pIDLSubItem_ToParent, SHGDNConstants.SHGDN_INFOLDER)
            If (prophImageList <> 0) And Not propUseCustomIcons Then
              #If NoCallbacks Then
                useStandardIcon = propUseStandardIcons
                If propUseStandardIconsForSlowItems And Not useStandardIcon Then
                  useStandardIcon = IsSlowItem(IItem, pIDLSubItem_ToParent, TVWItemInfo.pIDLToDesktop, True)
                End If
              #End If
            End If
          #End If

          ' Item hinzufügen
          InsertTVWItem_Common DispName, 0, TVWItemInfo, overlayIndex, itemAttr And SFGAOConstants.SFGAO_GHOSTED, itemAttr And SFGAOConstants.SFGAO_NEWCONTENT, useStandardIcon, True, isFirstItem
          isFirstItem = False
        End If
        #If Debuging Then
          FreeItemIDList debugger, "InsertRootItems #2", pIDLSubItem_ToParent
        #Else
          FreeItemIDList pIDLSubItem_ToParent
        #End If
      Wend
      Set IShIconOverlay = Nothing

      SortItems 0
    End If
    Set IEnum = Nothing
  End If
  Set IItem = Nothing
  blockAutoUpdate = blockAutoUpdate - 1

  RaiseItemLoadedSubItems 0
End Sub

' fügt dem TreeView alle SubItems von <hParentItem> hinzu
Private Sub InsertSubItems(ByVal hParentItem As Long)
  Const WM_SETREDRAW = &HB
  #If Debuging Then
    Dim cItems As Long
  #End If
  Dim DispName As String
  Dim EnumFlags As SHCONTFConstants
  Dim IEnum As IVBEnumIDList
  Dim IItem As IVBShellFolder
  Dim isFirstItem As Boolean
  Dim IShIconOverlay As IVBShellIconOverlay
  Dim itemAttr As SFGAOConstants
  Dim overlayIndex As Long
  Dim pIDLSubItem_ToDesktop As Long
  Dim pIDLSubItem_ToParent As Long
  Dim shouldShowIt As Boolean
  Dim TVWItemInfo As ITEMINFO_COMMON
  Dim TVWItemInfo_Parent As ITEMINFO_COMMON
  Dim useStandardIcon As Boolean

  If prophWnd = 0 Then Exit Sub
  If Not ItemExists(hParentItem, AllowNullConstants.anIfNoShowRoot) Then Exit Sub

  #If Debuging Then
    debugger.AddLogEntry "InsertSubItems: Adding sub-items for " & ItemHandleToTreePath(hParentItem), LogEntryTypeConstants.letInfo
  #End If

  RaiseItemLoadingSubItems hParentItem

  ' Daten über <hParentItem> sammeln
  If Not GetItemInfo_Common(hParentItem, TVWItemInfo_Parent) Then GoTo Ende

  #If Debuging Then
    If IsArchiveToExpandFQ(debugger, TVWItemInfo_Parent.pIDLToDesktop, propExpandArchives) Then
  #Else
    If IsArchiveToExpandFQ(TVWItemInfo_Parent.pIDLToDesktop, propExpandArchives) Then
  #End If
    InsertArchiveSubItems hParentItem
    GoTo Ende
  End If

  ' Aufzählung initiieren
  blockAutoUpdate = blockAutoUpdate + 1
  #If Debuging Then
    debugger.AddLogEntry "InsertSubItems: Calling GetISHFolderInterfaceFQ()", LogEntryTypeConstants.letInfo
    debugger.AddLogEntry "   0x" & Hex(TVWItemInfo_Parent.pIDLToDesktop), LogEntryTypeConstants.letOther
    Set IItem = GetISHFolderInterfaceFQ(debugger, TVWItemInfo_Parent.pIDLToDesktop)
  #Else
    Set IItem = GetISHFolderInterfaceFQ(TVWItemInfo_Parent.pIDLToDesktop)
  #End If
  If Not (IItem Is Nothing) Then
    ' on Vista SHCONTF_DRIVES doesn't work anymore, so let shouldShowItem() do the work
    If propDrivesOnly And Not ver_Win_Vista Then
      #If Debuging Then
        If IsMyComputer(debugger, IItem) Then
      #Else
        If IsMyComputer(IItem) Then
      #End If
        EnumFlags = SHCONTFConstants.SHCONTF_DRIVES
      Else
        EnumFlags = SHCONTFConstants.SHCONTF_FOLDERS
        If propIncludedItems And IncludedItemsConstants.iiFSFiles Then EnumFlags = EnumFlags Or SHCONTFConstants.SHCONTF_NONFOLDERS
        If propIncludedItems And IncludedItemsConstants.iiNonFSFiles Then EnumFlags = EnumFlags Or SHCONTFConstants.SHCONTF_NONFOLDERS
        If propFileAttributes And AttributesConstants.attHidden Then EnumFlags = EnumFlags Or SHCONTFConstants.SHCONTF_INCLUDEHIDDEN
        If propFolderAttributes And AttributesConstants.attHidden Then EnumFlags = EnumFlags Or SHCONTFConstants.SHCONTF_INCLUDEHIDDEN
      End If
    Else
      EnumFlags = SHCONTFConstants.SHCONTF_FOLDERS
      If propIncludedItems And IncludedItemsConstants.iiFSFiles Then EnumFlags = EnumFlags Or SHCONTFConstants.SHCONTF_NONFOLDERS
      If propIncludedItems And IncludedItemsConstants.iiNonFSFiles Then EnumFlags = EnumFlags Or SHCONTFConstants.SHCONTF_NONFOLDERS
      If propFileAttributes And AttributesConstants.attHidden Then EnumFlags = EnumFlags Or SHCONTFConstants.SHCONTF_INCLUDEHIDDEN
      If propFolderAttributes And AttributesConstants.attHidden Then EnumFlags = EnumFlags Or SHCONTFConstants.SHCONTF_INCLUDEHIDDEN
    End If

    IItem.EnumObjects prophWndShellUIParentWindow, EnumFlags, IEnum
    If Not (IEnum Is Nothing) Then
      If (prophImageList <> 0) And Not propUseCustomIcons Then
        If propShownOverlays <> ShownOverlaysConstants.soNone Then
          IItem.QueryInterface IID_IShellIconOverlay, IShIconOverlay
        End If
      End If
      isFirstItem = True
      #If Debuging Then
        debugger.AddLogEntry "InsertSubItems: Starting enumeration...", LogEntryTypeConstants.letInfo
      #End If
      While IEnum.Next(1, pIDLSubItem_ToParent, 0) = S_OK
        #If Debuging Then
          debugger.AddLogEntry "InsertSubItems: Enumerated relative pIDL 0x" & Hex(pIDLSubItem_ToParent), LogEntryTypeConstants.letInfo
        #End If
        If pIDLSubItem_ToParent Then
          pIDLSubItem_ToDesktop = ILAppendID(ILClone(TVWItemInfo_Parent.pIDLToDesktop), pIDLSubItem_ToParent, 1)
          #If Debuging Then
            shouldShowIt = ShouldShowItem(debugger, Me, IItem, pIDLSubItem_ToParent, pIDLSubItem_ToDesktop, hParentItem)
            debugger.AddLogEntry "   ShouldShowItem() returned: " & shouldShowIt, LogEntryTypeConstants.letOther
          #Else
            shouldShowIt = ShouldShowItem(Me, IItem, pIDLSubItem_ToParent, pIDLSubItem_ToDesktop, hParentItem)
          #End If
        Else
          shouldShowIt = False
        End If

        If Not shouldShowIt Then
          #If Debuging Then
            FreeItemIDList debugger, "InsertSubItems #1", pIDLSubItem_ToDesktop
          #Else
            FreeItemIDList pIDLSubItem_ToDesktop
          #End If
        Else
          ' Item anzeigen
          TVWItemInfo.pIDLToDesktop = pIDLSubItem_ToDesktop

          ' Typ des Items ermitteln
          #If Debuging Then
            itemAttr = GetAttributes(debugger, IItem, pIDLSubItem_ToParent, SFGAOConstants.SFGAO_FILESYSTEM Or SFGAOConstants.SFGAO_FOLDER Or SFGAOConstants.SFGAO_SHARE Or SFGAOConstants.SFGAO_LINK Or SFGAOConstants.SFGAO_GHOSTED Or SFGAOConstants.SFGAO_NEWCONTENT)
          #Else
            itemAttr = GetAttributes(IItem, pIDLSubItem_ToParent, SFGAOConstants.SFGAO_FILESYSTEM Or SFGAOConstants.SFGAO_FOLDER Or SFGAOConstants.SFGAO_SHARE Or SFGAOConstants.SFGAO_LINK Or SFGAOConstants.SFGAO_GHOSTED Or SFGAOConstants.SFGAO_NEWCONTENT)
          #End If
          If itemAttr And SFGAOConstants.SFGAO_FILESYSTEM Then
            If itemAttr And SFGAOConstants.SFGAO_FOLDER Then
              ' ist es wirklich ein Ordner?
              #If Debuging Then
                If FileExists_pIDL(debugger, IItem, pIDLSubItem_ToParent) Then
              #Else
                If FileExists_pIDL(IItem, pIDLSubItem_ToParent) Then
              #End If
                TVWItemInfo.itemType = CByte(ItemTypeConstants.itFSFile)
              Else
                TVWItemInfo.itemType = CByte(ItemTypeConstants.itFSFolder)
              End If
            Else
              TVWItemInfo.itemType = CByte(ItemTypeConstants.itFSFile)
            End If
          Else
            If itemAttr And SFGAOConstants.SFGAO_FOLDER Then
              TVWItemInfo.itemType = CByte(ItemTypeConstants.itNonFSFolder)
            Else
              TVWItemInfo.itemType = CByte(ItemTypeConstants.itNonFSFile)
            End If
          End If

          If (prophImageList <> 0) And Not propUseCustomIcons Then
            ' Overlays laden
            overlayIndex = 0
            If propShownOverlays <> ShownOverlaysConstants.soNone Then
              On Error Resume Next
              If IShIconOverlay Is Nothing Then
                If itemAttr And SFGAOConstants.SFGAO_SHARE Then overlayIndex = 1
                If itemAttr And SFGAOConstants.SFGAO_LINK Then overlayIndex = 2
              Else
                If IShIconOverlay.GetOverlayIndex(pIDLSubItem_ToParent, overlayIndex) = S_OK Then
                  ' jetzt haben wir den Overlay-Index für den Item, aber soll dieser Overlay
                  ' überhaupt angezeigt werden?
                  If ShouldShowOverlay(propShownOverlays, overlayIndex) Then
                    ' ja!
                  Else
                    ' kein Overlay
                    overlayIndex = 0
                  End If
                Else
                  ' kein Overlay
                  overlayIndex = 0
                End If
              End If
            End If
          End If

          #If Debuging Then
            DispName = pIDLToDisplayName(debugger, IItem, pIDLSubItem_ToParent, SHGDNConstants.SHGDN_INFOLDER)
            If (prophImageList <> 0) And Not propUseCustomIcons Then
              #If NoCallbacks Then
                useStandardIcon = propUseStandardIcons
                If propUseStandardIconsForSlowItems And Not useStandardIcon Then
                  useStandardIcon = IsSlowItem(debugger, IItem, pIDLSubItem_ToParent, TVWItemInfo.pIDLToDesktop, True)
                End If
              #End If
            End If
          #Else
            DispName = pIDLToDisplayName(IItem, pIDLSubItem_ToParent, SHGDNConstants.SHGDN_INFOLDER)
            If (prophImageList <> 0) And Not propUseCustomIcons Then
              #If NoCallbacks Then
                useStandardIcon = propUseStandardIcons
                If propUseStandardIconsForSlowItems And Not useStandardIcon Then
                  useStandardIcon = IsSlowItem(IItem, pIDLSubItem_ToParent, TVWItemInfo.pIDLToDesktop, True)
                End If
              #End If
            End If
          #End If

          ' Item hinzufügen
          InsertTVWItem_Common DispName, hParentItem, TVWItemInfo, overlayIndex, itemAttr And SFGAOConstants.SFGAO_GHOSTED, itemAttr And SFGAOConstants.SFGAO_NEWCONTENT, useStandardIcon, True, isFirstItem
          isFirstItem = False
          #If Debuging Then
            cItems = cItems + 1
          #End If
        End If
        #If Debuging Then
          FreeItemIDList debugger, "InsertSubItems #2", pIDLSubItem_ToParent
        #Else
          FreeItemIDList pIDLSubItem_ToParent
        #End If
      Wend
      #If Debuging Then
        debugger.AddLogEntry "InsertSubItems: Ended enumeration (inserted " & cItems & " items)", LogEntryTypeConstants.letInfo
      #End If
      Set IShIconOverlay = Nothing

      SortItems hParentItem
    End If
    Set IEnum = Nothing
  End If
  Set IItem = Nothing
  blockAutoUpdate = blockAutoUpdate - 1

  #If Debuging Then
    debugger.AddLogEntry "InsertSubItems: Ended adding sub-items for " & ItemHandleToTreePath(hParentItem), LogEntryTypeConstants.letInfo
  #End If

Ende:
  RaiseItemLoadedSubItems hParentItem
End Sub

' fügt <DisplayName> in das TreeView ein und gibt das Handle darauf zurück
' für Archiv-Items
Private Function InsertTVWItem_Archive(ByVal DisplayName As String, ByVal hParentItem As Long, ByVal hasChilds As Boolean, TVWItemInfo As ITEMINFO_ARCHIVE, ByVal ImgNormal As Long, ByVal ImgSelected As Long, Optional ByVal sort As Boolean = False, Optional ByVal hSortBase As Long = -1) As Long
  Dim Data As TVINSERTSTRUCT
  Dim GoUp As Boolean
  Dim hItem As Long
  Dim isFirstComparison As Boolean
  Dim pTVWItemInfo As Long
  Dim ret As Long
  Dim StateIcon As Long
  Dim tmp As Long
  Dim usehSortBase As Boolean

  If inIDE Then Exit Function
  If prophWnd = 0 Then
    ret = -1
    GoTo Ende
  End If
  If Not ItemExists(hParentItem, AllowNullConstants.anNever) Then
    ret = -1
    GoTo Ende
  End If
  If Not ItemExists(TVWItemInfo.hItem_ArchiveFile, AllowNullConstants.anNever) Then
    ret = -1
    GoTo Ende
  End If

  ' Speicher für die Item-Infos reservieren
  pTVWItemInfo = CoTaskMemAlloc(LenB(TVWItemInfo))
  If pTVWItemInfo = 0 Then
    ret = -1
    GoTo Ende
  End If

  #If ShowItemTypes Then
    With TVWItemInfo
      Select Case .itemType
        Case ItemTypeConstants.itArchiveFile
          DisplayName = DisplayName & "   - ArchiveFile"
        Case ItemTypeConstants.itArchiveFolder
          DisplayName = DisplayName & "   - ArchiveFolder"
      End Select
    End With
  #End If

  SetItemChildrenProperty hParentItem, True

  With Data
    .hParent = hParentItem
    With .itemex
      .cChildren = Abs(hasChilds)
      .pszText = DisplayName
      .cchTextMax = Len(.pszText)

      CopyMemory pTVWItemInfo, VarPtr(TVWItemInfo), LenB(TVWItemInfo)
      .lParam = pTVWItemInfo
      .Mask = TVIF_CHILDREN Or TVIF_TEXT Or TVIF_PARAM Or TVIF_STATE

      If (prophImageList <> 0) And Not propUseCustomIcons Then
        .iImage = ImgNormal
        .iSelectedImage = ImgSelected
        .Mask = .Mask Or TVIF_IMAGE Or TVIF_SELECTEDIMAGE
      End If

      If TVWItemInfo.attributes And FILE_ATTRIBUTE_HIDDEN Then
        If propShowHiddenDimmed Then .state = .state Or TVIS_CUT
      End If

      StateIcon = ItemStateIconIndex(hParentItem)
      .state = .state Or INDEXTOSTATEIMAGEMASK(StateIcon)
      .StateMask = .state Or TVIS_STATEIMAGEMASK
    End With

    If sort Then
      usehSortBase = ItemExists(hSortBase, AllowNullConstants.anIfNoShowRoot)
      hItem = IIf(usehSortBase, hSortBase, ItemGetFirstSubItem(hParentItem))
      If hItem <> -1 Then
        tmp = IIf(usehSortBase, hSortBase, TVI_FIRST)
        If usehSortBase Then isFirstComparison = True
        Do
          ret = ISort_CBSortArchiveItems(GetItemlParam(hItem), .itemex.lParam, True)
          If ret = 0 Then ret = StrComp(GetItemText(hItem), DisplayName, VbCompareMethod.vbTextCompare)
          If ret > 0 Then
            ' 2. Item zuerst
            If isFirstComparison Then
              isFirstComparison = False
              GoUp = True
            End If
            If GoUp Then
              ' weitersuchen (nach oben)
              tmp = hItem
              hItem = ItemGetPreviousItem(hItem)
              If hItem = -1 Then .hInsertAfter = TVI_FIRST
            Else
              .hInsertAfter = tmp
              Exit Do
            End If
          Else
            ' 1. Item zuerst
            If GoUp Then
              .hInsertAfter = hItem
              Exit Do
            Else
              ' weitersuchen (nach unten)
              tmp = hItem
              hItem = ItemGetNextItem(hItem)
              If hItem = -1 Then .hInsertAfter = TVI_LAST
            End If
          End If
        Loop Until hItem = -1
      End If
    End If
    ret = SendMessage(prophWnd, TVM_INSERTITEM, 0, Data)

    If ret = 0 Then
      ' Speicher wieder freigeben
      CoTaskMemFree pTVWItemInfo

      ret = -1
    Else
      ' jetzt noch das Handle des Items speichern
      CopyMemory pTVWItemInfo + 4, VarPtr(ret), LenB(ret)
    End If
  End With

Ende:
  InsertTVWItem_Archive = ret
End Function

' fügt <DisplayName> in das TreeView ein und gibt das Handle darauf zurück
' für "normale" Items
Private Function InsertTVWItem_Common(DisplayName As String, hParentItem As Long, TVWItemInfo As ITEMINFO_COMMON, overlayIndex As Long, Ghosted As Boolean, Bold As Boolean, useStandardIcon As Boolean, Optional useCache As Boolean = False, Optional invalidateCaches As Boolean = False) As Long
  Dim Data As TVINSERTSTRUCT
'  Dim GoUp As Boolean
  Dim hItem As Long
'  Dim isFirstComparison As Boolean
  Dim pTVWItemInfo As Long
  Dim ret As Long
  Dim StateIcon As Long
  Dim tmp As Long
'  Dim usehSortBase As Boolean
  Static Cache_hParentIsOk As Boolean
  Static Cache_hParentStateIcon As Long

'  If prophWnd = 0 Then
'    ret = -1
'    GoTo Ende
'  End If
  If useCache Then
    If invalidateCaches Then Cache_hParentIsOk = ItemExists(hParentItem, AllowNullConstants.anAlways, True)
    If Not Cache_hParentIsOk Then
      ret = -1
      GoTo Ende
    End If
  Else
    If Not ItemExists(hParentItem, AllowNullConstants.anAlways, True) Then
      ret = -1
      GoTo Ende
    End If
  End If
'  If Not NoItemData Then
'    If TVWItemInfo.IParent Is Nothing Then
'      ret = -1
'      GoTo Ende
'    End If
'
    ' Speicher für die Item-Infos reservieren
    pTVWItemInfo = CoTaskMemAlloc(LenB(TVWItemInfo))
    If pTVWItemInfo = 0 Then
      ret = -1
      GoTo Ende
    End If
'  End If

  #If ShowItemTypes Then
    With TVWItemInfo
      Select Case .itemType
        Case ItemTypeConstants.itFSFile
          DisplayName = DisplayName & "   - FSFile"
        Case ItemTypeConstants.itFSFolder
          DisplayName = DisplayName & "   - FSFolder"
        Case ItemTypeConstants.itNonFSFile
          DisplayName = DisplayName & "   - NonFSFile"
        Case ItemTypeConstants.itNonFSFolder
          DisplayName = DisplayName & "   - NonFSFolder"
      End Select
    End With
  #End If

  If useCache Then
    If invalidateCaches Then SetItemChildrenProperty hParentItem, True
  Else
    SetItemChildrenProperty hParentItem, True
  End If

  #If Debuging Then
    #If LogItemInsertions Then
      debugger.AddLogEntry "Inserting item: " & DisplayName, LogEntryTypeConstants.letInfo
      debugger.AddLogEntry "   pIDL: 0x" & Hex(TVWItemInfo.pIDLToDesktop), LogEntryTypeConstants.letOther
    #End If
  #End If

  With Data
    .hParent = hParentItem
    With .itemex
      #If NoCallbacks Then
        .cChildren = 1
      #Else
        .cChildren = I_CHILDRENCALLBACK
      #End If
      .pszText = DisplayName
      .cchTextMax = Len(.pszText)

'      If Not NoItemData Then
        CopyMemory pTVWItemInfo, VarPtr(TVWItemInfo), LenB(TVWItemInfo)
        .lParam = pTVWItemInfo
'      End If
      .Mask = TVIF_CHILDREN Or TVIF_TEXT Or TVIF_PARAM Or TVIF_STATE

      If (prophImageList <> 0) And Not propUseCustomIcons Then
        #If NoCallbacks Then
          #If Debuging Then
            .iImage = GetSysIconIndex_Light(debugger, TVWItemInfo.pIDLToDesktop, useStandardIcon, useLargeIcons)
            .iSelectedImage = GetSysIconIndex_Light(debugger, TVWItemInfo.pIDLToDesktop, True, useStandardIcon, useLargeIcons)
          #Else
            .iImage = GetSysIconIndex_Light(TVWItemInfo.pIDLToDesktop, useStandardIcon, useLargeIcons)
            .iSelectedImage = GetSysIconIndex_Light(TVWItemInfo.pIDLToDesktop, True, useStandardIcon, useLargeIcons)
          #End If
        #Else
          .iImage = I_IMAGECALLBACK
          .iSelectedImage = I_IMAGECALLBACK
        #End If
        .Mask = .Mask Or TVIF_IMAGE Or TVIF_SELECTEDIMAGE
      End If

      If Bold Then .state = .state Or TVIS_BOLD
      If Ghosted Then
        If propShowHiddenDimmed Then .state = .state Or TVIS_CUT
      End If

      If overlayIndex > 0 Then
        .state = .state Or INDEXTOOVERLAYMASK(overlayIndex)
        .StateMask = .StateMask Or TVIS_OVERLAYMASK
      End If

      If propCheckBoxes Then
        If hParentItem Then
          If useCache Then
            If invalidateCaches Then Cache_hParentStateIcon = ItemStateIconIndex(hParentItem)
            StateIcon = Cache_hParentStateIcon
          Else
            StateIcon = ItemStateIconIndex(hParentItem)
          End If
          .state = .state Or INDEXTOSTATEIMAGEMASK(StateIcon)
          .StateMask = .StateMask Or TVIS_STATEIMAGEMASK
        End If
      End If
      .StateMask = .StateMask Or .state
    End With

'    If NoItemData Then sort = False
'    If sort Then
'      usehSortBase = ItemExists(hSortBase, AllowNullConstants.anIfNoShowRoot)
'      hItem = IIf(usehSortBase, hSortBase, ItemGetFirstSubItem(hParentItem))
'      If hItem <> -1 Then
'        tmp = IIf(usehSortBase, hSortBase, TVI_FIRST)
'        If usehSortBase Then isFirstComparison = True
'        Do
'          If ISort_CBSortCommonItems(getItemlParam(hItem), .itemex.lParam) >= 0 Then
'            ' 2. Item zuerst
'            If isFirstComparison Then
'              isFirstComparison = False
'              GoUp = True
'            End If
'            If GoUp Then
'              ' weitersuchen (nach oben)
'              tmp = hItem
'              hItem = ItemGetPreviousItem(hItem)
'              If hItem = -1 Then .hInsertAfter = TVI_FIRST
'            Else
'              .hInsertAfter = tmp
'              Exit Do
'            End If
'          Else
'            ' 1. Item zuerst
'            If GoUp Then
'              .hInsertAfter = hItem
'              Exit Do
'            Else
'              ' weitersuchen (nach unten)
'              tmp = hItem
'              hItem = ItemGetNextItem(hItem)
'              If hItem = -1 Then .hInsertAfter = TVI_LAST
'            End If
'          End If
'        Loop Until hItem = -1
'      End If
'    End If
    ret = SendMessage(prophWnd, TVM_INSERTITEM, 0, Data)

    If ret = 0 Then
      ' Speicher wieder freigeben
      #If Debuging Then
        FreeItemIDList debugger, "InsertTVWItem_Common", TVWItemInfo.pIDLToDesktop
      #Else
        FreeItemIDList TVWItemInfo.pIDLToDesktop
      #End If
      CoTaskMemFree pTVWItemInfo

      ret = -1
    #If Debuging Then
      #If LogItemInsertions Then
        Else
          debugger.AddLogEntry "   Handle: 0x" & Hex(ret), LogEntryTypeConstants.letOther
      #End If
    #End If
    End If
  End With

Ende:
  InsertTVWItem_Common = ret
End Function

' sucht den Item heraus, dessen <ItemInfo.pIDLToDesktop> mit <pIDLToDesktop> identisch ist
' beginnt bei <hStartItem>
Private Function InternalpIDLToItemHandle(pIDLToDesktop As Long, Optional ByVal hStartItem As Long = -2, Optional ByVal insertIfNecessary As Boolean = False, Optional ByVal freepIDL As Boolean = False, Optional ByVal diveIntoSubItems As Boolean = True, Optional ByVal forceInsertion As Boolean = False, Optional ByVal exactMatchRequired As Boolean = True) As Long
  Dim bFoundIt As Boolean
  Dim bIsChild As Boolean
  Dim cParent As Long
  Dim cSubItem As Long
  Dim h As Long
  Dim hItem As Long
  Dim hItem2 As Long
  Dim pIDL As Long
  Dim pIDLToParent As Long
  Dim pIDLToRelative As Long
  Dim ret As Long
  Dim TVWItemInfo As ITEMINFO_COMMON

  If pIDLToDesktop = 0 Then
    ret = -1
    GoTo Ende
  End If

  If hStartItem = -2 Then
    hStartItem = IIf(propShowRoot, GetRootOrFirstItem, 0)
  End If

  If Not ItemExists(hStartItem, AllowNullConstants.anIfNoShowRoot) Then
    ret = -1
    GoTo Ende
  End If

  blockAutoUpdate = blockAutoUpdate + 1
  If hStartItem Then
    hItem = hStartItem
  Else
    ' Root-Item testen
    If ILIsEqual(RootItemInfo.pIDLToDesktop, pIDLToDesktop) Then
      ' gefunden?
      #If Debuging Then
        If LCase$(pIDLToDisplayName(debugger, IDesktop, pIDLToDesktop, SHGDNConstants.SHGDN_NORMAL Or SHGDNConstants.SHGDN_FORPARSING)) = LCase$(pIDLToDisplayName(debugger, IDesktop, RootItemInfo.pIDLToDesktop, SHGDNConstants.SHGDN_NORMAL Or SHGDNConstants.SHGDN_FORPARSING)) Then
      #Else
        If LCase$(pIDLToDisplayName(IDesktop, pIDLToDesktop, SHGDNConstants.SHGDN_NORMAL Or SHGDNConstants.SHGDN_FORPARSING)) = LCase$(pIDLToDisplayName(IDesktop, RootItemInfo.pIDLToDesktop, SHGDNConstants.SHGDN_NORMAL Or SHGDNConstants.SHGDN_FORPARSING)) Then
      #End If
        ' gefunden!
        ret = 0
        bFoundIt = True
      End If
    ElseIf ILIsParent(RootItemInfo.pIDLToDesktop, pIDLToDesktop, 0) Then
      bIsChild = True
    Else
      bIsChild = False
    End If

    If Not bFoundIt Then
      If bIsChild Then
        ' wir sind auf der richtigen Spur
        hItem = GetRootOrFirstItem
        If IsWithinRange(hItem, -1, 0) And insertIfNecessary Then
          ' das TreeView ist momentan leer
          ' möglicherweise fällt der Item bei IEnumIDList durch
          #If Debuging Then
            cParent = CountItemIDs(debugger, RootItemInfo.pIDLToDesktop)
            cSubItem = CountItemIDs(debugger, pIDLToDesktop)
            debugger.AddLogEntry "InternalpIDLToItemHandle: Calling GetLastItemIDs() #1", LogEntryTypeConstants.letInfo
            debugger.AddLogEntry "   0x" & Hex(pIDLToDesktop), LogEntryTypeConstants.letOther
            pIDLToRelative = GetLastItemIDs(debugger, pIDLToDesktop, cSubItem - cParent)
          #Else
            cParent = CountItemIDs(RootItemInfo.pIDLToDesktop)
            cSubItem = CountItemIDs(pIDLToDesktop)
            pIDLToRelative = GetLastItemIDs(pIDLToDesktop, cSubItem - cParent)
          #End If
          If pIDLToRelative Then
            #If Debuging Then
              pIDLToParent = CopyFirstItemIDs(debugger, pIDLToRelative, 1)
            #Else
              pIDLToParent = CopyFirstItemIDs(pIDLToRelative, 1)
            #End If
            #If Debuging Then
              debugger.AddLogEntry "InternalpIDLToItemHandle: Calling AddSubItem() #1", LogEntryTypeConstants.letInfo
              debugger.AddLogEntry "   0x0", LogEntryTypeConstants.letOther
            #End If
            hItem2 = AddSubItem(0, pIDLToParent, True, forceInsertion, True)
            If hItem2 <> -1 Then
              SortItems 0
            End If
            hItem = hItem2
          End If
        End If
      End If
    End If
  End If

  #If Debuging Then
    debugger.AddLogEntry "InternalpIDLToItemHandle: pIDLToDesktop=0x" & Hex(pIDLToDesktop), LogEntryTypeConstants.letInfo
  #End If
  Do While Not IsWithinRange(hItem, -1, 0) And Not bFoundIt
    ' Item prüfen
    #If Debuging Then
      debugger.AddLogEntry "InternalpIDLToItemHandle: Calling GetFullyQualifiedPIDL() #1", LogEntryTypeConstants.letInfo
      debugger.AddLogEntry "   Item 0x" & Hex(hItem), LogEntryTypeConstants.letOther
    #End If
    pIDL = GetFullyQualifiedPIDL(hItem)
    #If Debuging Then
      debugger.AddLogEntry "   pIDL 0x" & Hex(pIDL), LogEntryTypeConstants.letOther
    #End If
    If pIDL = 0 Then
      hItem = 0
      Exit Do
    End If
    If ILIsEqual(pIDLToDesktop, pIDL) Then
      ' gefunden?
      #If Debuging Then
        If LCase$(pIDLToDisplayName(debugger, IDesktop, pIDLToDesktop, SHGDNConstants.SHGDN_NORMAL Or SHGDNConstants.SHGDN_FORPARSING)) = LCase$(pIDLToDisplayName(debugger, IDesktop, pIDL, SHGDNConstants.SHGDN_NORMAL Or SHGDNConstants.SHGDN_FORPARSING)) Then
      #Else
        If LCase$(pIDLToDisplayName(IDesktop, pIDLToDesktop, SHGDNConstants.SHGDN_NORMAL Or SHGDNConstants.SHGDN_FORPARSING)) = LCase$(pIDLToDisplayName(IDesktop, pIDL, SHGDNConstants.SHGDN_NORMAL Or SHGDNConstants.SHGDN_FORPARSING)) Then
      #End If
        ' gefunden!
        ret = hItem
        bFoundIt = True
      End If
    ElseIf ILIsParent(pIDL, pIDLToDesktop, 0) Then
      bIsChild = True
    Else
      bIsChild = False
    End If
    #If Debuging Then
      debugger.AddLogEntry "InternalpIDLToItemHandle...", LogEntryTypeConstants.letInfo
      debugger.AddLogEntry "   bFoundIt=" & bFoundIt, LogEntryTypeConstants.letOther
      debugger.AddLogEntry "   bIsChild=" & bIsChild, LogEntryTypeConstants.letOther
    #End If

    If Not bFoundIt Then
      If bIsChild Then
        ' wir sind auf der richtigen Spur
        ret = hItem
        If diveIntoSubItems Then
          hItem2 = Me.ItemGetFirstSubItem(hItem)
          If (hItem2 = -1) And insertIfNecessary Then
            InsertSubItems hItem
            h = hItem
            hItem2 = Me.ItemGetFirstSubItem(hItem)
            If hItem2 = -1 Then
              ' möglicherweise fällt der Item bei IEnumIDList durch
              If GetItemInfo_Common(hItem, TVWItemInfo) Then
                #If Debuging Then
                  cParent = CountItemIDs(debugger, TVWItemInfo.pIDLToDesktop)
                  cSubItem = CountItemIDs(debugger, pIDLToDesktop)
                  debugger.AddLogEntry "InternalpIDLToItemHandle: Calling GetLastItemIDs() #2", LogEntryTypeConstants.letInfo
                  debugger.AddLogEntry "   0x" & Hex(pIDLToDesktop), LogEntryTypeConstants.letOther
                  pIDLToRelative = GetLastItemIDs(debugger, pIDLToDesktop, cSubItem - cParent)
                #Else
                  cParent = CountItemIDs(TVWItemInfo.pIDLToDesktop)
                  cSubItem = CountItemIDs(pIDLToDesktop)
                  pIDLToRelative = GetLastItemIDs(pIDLToDesktop, cSubItem - cParent)
                #End If
                If pIDLToRelative Then
                  #If Debuging Then
                    pIDLToParent = CopyFirstItemIDs(debugger, pIDLToRelative)
                  #Else
                    pIDLToParent = CopyFirstItemIDs(pIDLToRelative)
                  #End If
                  #If Debuging Then
                    debugger.AddLogEntry "InternalpIDLToItemHandle: Calling AddSubItem() #2", LogEntryTypeConstants.letInfo
                    debugger.AddLogEntry "   0x" & Hex(hItem), LogEntryTypeConstants.letOther
                  #End If
                  hItem2 = AddSubItem(hItem, pIDLToParent, (pIDLToParent <> pIDL), forceInsertion, True)
                  If hItem2 <> -1 Then
                    SortItems hItem
                  End If
                  hItem = hItem2
                End If
              End If
            Else
              hItem = hItem2
            End If
            If Me.ItemGetFirstSubItem(h) = -1 Then SetItemChildrenProperty h, False
          Else
            hItem = hItem2
          End If
        Else
          ' wir sollen nicht in die Sub-Items vordringen
          hItem = -1
        End If
      Else
        ' wir sind auf dem Holzweg -> Nachbar-Item testen
        hItem2 = Me.ItemGetNextItem(hItem)
        If (hItem2 = -1) And insertIfNecessary Then
          hItem = Me.ItemGetParentItem(hItem)
          ' möglicherweise fällt der Item bei IEnumIDList durch
          If GetItemInfo_Common(hItem, TVWItemInfo) Then
            #If Debuging Then
              cParent = CountItemIDs(debugger, TVWItemInfo.pIDLToDesktop)
              cSubItem = CountItemIDs(debugger, pIDLToDesktop)
              debugger.AddLogEntry "InternalpIDLToItemHandle: Calling GetLastItemIDs() #3", LogEntryTypeConstants.letInfo
              debugger.AddLogEntry "   0x" & Hex(pIDLToDesktop), LogEntryTypeConstants.letOther
              pIDLToRelative = GetLastItemIDs(debugger, pIDLToDesktop, cSubItem - cParent)
            #Else
              cParent = CountItemIDs(TVWItemInfo.pIDLToDesktop)
              cSubItem = CountItemIDs(pIDLToDesktop)
              pIDLToRelative = GetLastItemIDs(pIDLToDesktop, cSubItem - cParent)
            #End If
            If pIDLToRelative Then
              #If Debuging Then
                pIDLToParent = CopyFirstItemIDs(debugger, pIDLToRelative)
              #Else
                pIDLToParent = CopyFirstItemIDs(pIDLToRelative)
              #End If
              #If Debuging Then
                debugger.AddLogEntry "InternalpIDLToItemHandle: Calling AddSubItem() #3", LogEntryTypeConstants.letInfo
                debugger.AddLogEntry "   0x" & Hex(hItem), LogEntryTypeConstants.letOther
              #End If
              hItem2 = AddSubItem(hItem, pIDLToParent, (pIDLToParent <> pIDL), forceInsertion, True)
              If hItem2 <> -1 Then
                SortItems hItem
              End If
              hItem = hItem2
            End If
          End If
        Else
          hItem = hItem2
        End If
      End If
    End If
  Loop
  blockAutoUpdate = blockAutoUpdate - 1

Ende:
  #If Debuging Then
    debugger.AddLogEntry "End of InternalpIDLToItemHandle...", LogEntryTypeConstants.letInfo
  #End If
  If exactMatchRequired Then
    If Not IsWithinRange(ret, -1, 0) Then
      If ILIsEqual(pIDLToDesktop, GetFullyQualifiedPIDL(ret)) = 0 Then ret = -1
    End If
  End If

  If freepIDL Then
    #If Debuging Then
      FreeItemIDList debugger, "InternalpIDLToItemHandle", pIDLToDesktop
    #Else
      FreeItemIDList pIDLToDesktop
    #End If
  End If
  If ret = 0 Then ret = -1
  InternalpIDLToItemHandle = ret
  #If Debuging Then
    debugger.AddLogEntry "Leaving InternalpIDLToItemHandle (ret=0x" & Hex(ret) & ")", LogEntryTypeConstants.letInfo
  #End If
End Function

Private Sub InvalidateIcons(ByVal hItem As Long, ByVal iconIndex As Long, Optional ByVal invalidateOverlay As Boolean = False, Optional ByVal updateSubItems As Boolean = True)
  Dim ext As String
  Dim IParent As IVBShellFolder
  Dim IShIconOverlay As IVBShellIconOverlay
  Dim itemAttr As SFGAOConstants
  Dim itemData As TVITEMEX
  Dim itemType As Byte
  Dim overlayIndex As Long
  Dim pIDLToParent As Long
  Dim state As Long
  Dim TVWItemInfo As ITEMINFO_COMMON

  If hItem Then
    With itemData
      .hItem = hItem
      .Mask = TVIF_HANDLE Or TVIF_IMAGE Or TVIF_PARAM Or TVIF_SELECTEDIMAGE Or TVIF_STATE
      .StateMask = TVIS_OVERLAYMASK
      SendMessage prophWnd, TVM_GETITEM, 0, itemData
      CopyMemory VarPtr(itemType), .lParam, LenB(itemType)
      .Mask = TVIF_HANDLE

      #If Debuging Then
        debugger.AddLogEntry "Invalidating icons/overlay for item 0x" & Hex(hItem) & " (" & ItemHandleToTreePath(hItem) & ")", LogEntryTypeConstants.letInfo
      #End If

      If (iconIndex = -1) Or (.iImage = iconIndex) Or (.iSelectedImage = iconIndex) Then
        .Mask = .Mask Or TVIF_IMAGE Or TVIF_SELECTEDIMAGE
        If itemType = ItemTypeConstants.itArchiveFile Then
          ext = GetFileNameExtension(GetItemText(hItem))
          .iImage = GetSysIconIndexFromExt(ext, False, useLargeIcons)
          .iSelectedImage = .iImage
        ElseIf itemType = ItemTypeConstants.itArchiveFolder Then
          .iImage = IIf(useLargeIcons, DEFICON_FOLDER_LARGE, DEFICON_FOLDER_SMALL)
          If propUseSelectedIcons Then
            .iSelectedImage = IIf(useLargeIcons, DEFICON_OPENFOLDER_LARGE, DEFICON_OPENFOLDER_SMALL)
          Else
            .iSelectedImage = .iImage
          End If
        Else
          .iImage = I_IMAGECALLBACK
          .iSelectedImage = I_IMAGECALLBACK
        End If
      End If

      If invalidateOverlay And (itemType <> ItemTypeConstants.itArchiveFile) And (itemType <> ItemTypeConstants.itArchiveFolder) Then
        CopyMemory VarPtr(TVWItemInfo), .lParam, LenB(TVWItemInfo)
        #If Debuging Then
          SplitFullyQualifiedPIDL debugger, TVWItemInfo.pIDLToDesktop, IParent, pIDLToParent
        #Else
          SplitFullyQualifiedPIDL TVWItemInfo.pIDLToDesktop, IParent, pIDLToParent
        #End If
        If (pIDLToParent <> 0) And Not (IParent Is Nothing) Then
          ' Overlays laden
          If propShownOverlays <> ShownOverlaysConstants.soNone Then
            On Error Resume Next
            IParent.QueryInterface IID_IShellIconOverlay, IShIconOverlay
            If IShIconOverlay Is Nothing Then
              #If Debuging Then
                itemAttr = GetAttributes(debugger, IParent, pIDLToParent, SFGAOConstants.SFGAO_SHARE Or SFGAOConstants.SFGAO_LINK)
              #Else
                itemAttr = GetAttributes(IParent, pIDLToParent, SFGAOConstants.SFGAO_SHARE Or SFGAOConstants.SFGAO_LINK)
              #End If
              If itemAttr And SFGAOConstants.SFGAO_SHARE Then overlayIndex = 1
              If itemAttr And SFGAOConstants.SFGAO_LINK Then overlayIndex = 2
            Else
              If IShIconOverlay.GetOverlayIndex(pIDLToParent, overlayIndex) = S_OK Then
                ' jetzt haben wir den Overlay-Index für den Item, aber soll dieser Overlay überhaupt
                ' angezeigt werden?
                If ShouldShowOverlay(propShownOverlays, overlayIndex) Then
                  ' ja!
                Else
                  ' kein Overlay
                  overlayIndex = 0
                End If
              Else
                ' kein Overlay
                overlayIndex = 0
              End If
              Set IShIconOverlay = Nothing
            End If
          End If

          .Mask = .Mask Or TVIF_STATE
          ' TODO: Mglw. funktioniert das Entfernen eines Overlays unter Vista so nicht
          .state = state And Not TVIS_OVERLAYMASK
          If overlayIndex > 0 Then
            .state = state Or INDEXTOOVERLAYMASK(overlayIndex)
          End If
        End If
        Set IParent = Nothing
      End If
      SendMessage prophWnd, TVM_SETITEM, 0, itemData
    End With
  End If

  If updateSubItems Then
    If hItem Then
      If ver_Comctl32_580 Then
        state = SendMessageAsLong(prophWnd, TVM_GETITEMSTATE, hItem, TVIS_EXPANDED)
      Else
        With itemData
          .hItem = hItem
          .StateMask = TVIS_EXPANDED
          .Mask = TVIF_HANDLE Or TVIF_STATE
          SendMessage prophWnd, TVM_GETITEM, 0, itemData
          state = .state
        End With
      End If

      If state And TVIS_EXPANDED Then
        hItem = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_CHILD, hItem)
        While hItem
          InvalidateIcons hItem, iconIndex

          hItem = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_NEXT, hItem)
        Wend
      Else
        ItemCollapse hItem, True
      End If
    Else
      hItem = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_ROOT, 0)
      While hItem
        InvalidateIcons hItem, iconIndex

        hItem = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_NEXT, hItem)
      Wend
    End If
  End If
End Sub

' lädt die SysImageList neu und aktualisiert alle Icons und Overlays
Private Sub InvalidateImgLstAndIcons()
  Const WM_SETREDRAW = &HB
  Dim Size_LargeIcons As Long
  Dim Size_SmallIcons As Long

  SendMessageAsLong prophWnd, WM_SETREDRAW, 0, 0

  If propAlwaysUseSmallIcons Then
    useLargeIcons = False
  Else
    Size_LargeIcons = GetSystemMetrics(SM_CXICON)
    Size_SmallIcons = GetSystemMetrics(SM_CXSMICON)
    useLargeIcons = (propItemHeight >= Size_LargeIcons * 0.75)
  End If

  If (prophImageList <> 0) And Not propUseCustomIcons Then
    ' zunächst die neue SysImageList laden
    SendMessageAsLong prophWnd, TVM_SETIMAGELIST, TVSIL_NORMAL, 0
    prophImageList = 0

    ' SysImageList initialisieren
    ' -> mit Standard-Icons und den Overlays
    If ver_Win_NTBased Then FileIconInit 0
    UpdateDefaultIconIndices

    prophImageList = GetSysImageList(useLargeIcons)
    SendMessageAsLong prophWnd, TVM_SETIMAGELIST, TVSIL_NORMAL, prophImageList
  End If

  If propCheckBoxes Then
    SendMessageAsLong prophWnd, TVM_SETIMAGELIST, TVSIL_STATE, IIf(useLargeIcons, prophLargeStateImageList, prophSmallStateImageList)
  End If
  SendMessageAsLong prophWnd, WM_SETREDRAW, 1, 0

  If (prophImageList <> 0) And Not propUseCustomIcons Then
    InvalidateIcons 0, -1, True
  End If
End Sub

Private Sub InvalidateOverlays(ByVal hItem As Long)
  Dim IParent As IVBShellFolder
  Dim IShIconOverlay As IVBShellIconOverlay
  Dim itemAttr As SFGAOConstants
  Dim itemData As TVITEMEX
  Dim itemType As Byte
  Dim overlayIndex As Long
  Dim pIDLToParent As Long
  Dim state As Long
  Dim TVWItemInfo As ITEMINFO_COMMON

  If hItem Then
    With itemData
      .hItem = hItem
      .Mask = TVIF_HANDLE Or TVIF_PARAM Or TVIF_STATE
      .StateMask = TVIS_OVERLAYMASK
      SendMessage prophWnd, TVM_GETITEM, 0, itemData
      CopyMemory VarPtr(itemType), .lParam, LenB(itemType)
      .Mask = TVIF_HANDLE Or TVIF_STATE

      #If Debuging Then
        debugger.AddLogEntry "Invalidating overlay for item 0x" & Hex(hItem) & " (" & ItemHandleToTreePath(hItem) & ")", LogEntryTypeConstants.letInfo
      #End If

      If (itemType <> ItemTypeConstants.itArchiveFile) And (itemType <> ItemTypeConstants.itArchiveFolder) Then
        CopyMemory VarPtr(TVWItemInfo), .lParam, LenB(TVWItemInfo)
        #If Debuging Then
          SplitFullyQualifiedPIDL debugger, TVWItemInfo.pIDLToDesktop, IParent, pIDLToParent
        #Else
          SplitFullyQualifiedPIDL TVWItemInfo.pIDLToDesktop, IParent, pIDLToParent
        #End If
        If (pIDLToParent <> 0) And Not (IParent Is Nothing) Then
          ' Overlays laden
          If propShownOverlays <> ShownOverlaysConstants.soNone Then
            On Error Resume Next
            IParent.QueryInterface IID_IShellIconOverlay, IShIconOverlay
            If IShIconOverlay Is Nothing Then
              #If Debuging Then
                itemAttr = GetAttributes(debugger, IParent, pIDLToParent, SFGAOConstants.SFGAO_SHARE Or SFGAOConstants.SFGAO_LINK)
              #Else
                itemAttr = GetAttributes(IParent, pIDLToParent, SFGAOConstants.SFGAO_SHARE Or SFGAOConstants.SFGAO_LINK)
              #End If
              If itemAttr And SFGAOConstants.SFGAO_SHARE Then overlayIndex = 1
              If itemAttr And SFGAOConstants.SFGAO_LINK Then overlayIndex = 2
            Else
              If IShIconOverlay.GetOverlayIndex(pIDLToParent, overlayIndex) = S_OK Then
                ' jetzt haben wir den Overlay-Index für den Item, aber soll dieser Overlay überhaupt
                ' angezeigt werden?
                If ShouldShowOverlay(propShownOverlays, overlayIndex) Then
                  ' ja!
                Else
                  ' kein Overlay
                  overlayIndex = 0
                End If
              Else
                ' kein Overlay
                overlayIndex = 0
              End If
              Set IShIconOverlay = Nothing
            End If
          End If

          ' TODO: Mglw. funktioniert das Entfernen eines Overlays unter Vista so nicht
          .state = state And Not TVIS_OVERLAYMASK
          If overlayIndex > 0 Then
            .state = state Or INDEXTOOVERLAYMASK(overlayIndex)
          End If
        End If
        Set IParent = Nothing
        SendMessage prophWnd, TVM_SETITEM, 0, itemData
      End If
    End With
  End If

  If hItem Then
    If ver_Comctl32_580 Then
      state = SendMessageAsLong(prophWnd, TVM_GETITEMSTATE, hItem, TVIS_EXPANDED)
    Else
      With itemData
        .hItem = hItem
        .StateMask = TVIS_EXPANDED
        .Mask = TVIF_HANDLE Or TVIF_STATE
        SendMessage prophWnd, TVM_GETITEM, 0, itemData
        state = .state
      End With
    End If

    If state And TVIS_EXPANDED Then
      hItem = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_CHILD, hItem)
      While hItem
        InvalidateOverlays hItem

        hItem = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_NEXT, hItem)
      Wend
    Else
      ItemCollapse hItem, True
    End If
  Else
    hItem = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_ROOT, 0)
    While hItem
      InvalidateOverlays hItem

      hItem = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_NEXT, hItem)
    Wend
  End If
End Sub

Private Function IsCustomDrawNecessary() As Boolean
  Dim ret As Boolean
'  Static iCalls As Long
'  Static ret As Boolean
'
'  If iCalls <> 0 Then
'    isCustomDrawNecessary = ret
'    Exit Function
'  End If
'  iCalls = (iCalls + 1) Mod 15
'  ret = False

  If inIDE Then Exit Function

  If propExpandArchives <> 0 Then
    If (propArchiveItemBackColor <> propBackColor) Or (propArchiveItemForeColor <> propForeColor) Then ret = True
  End If
  If (propCompressedBackColor <> propBackColor) Or (propCompressedForeColor <> propForeColor) Then ret = True
  If (propEncryptedBackColor <> propBackColor) Or (propEncryptedForeColor <> propForeColor) Then ret = True
  If propHotTracking Then
    ' HotBackColor und HotForeColor gelten nur für HotTracking
    If propHotBackColor <> -1 Or propHotForeColor <> -1 Then ret = True
  End If
  If propSelectedItemBackColor <> -1 Or propSelectedItemForeColor <> -1 Then ret = True
  If propFavoritesStyle Then ret = True
  #If BkImageSupport Then
    If Not (propBkImage Is Nothing) Then ret = True
  #End If

  If prophImageList Then
    If propDimSelectedItemsIcon And Not propCheckBoxes Then ret = True
    If Not ret Then
'    If Not ret And Not ver_Comctl32_610 Then
      If SendMessageAsLong(prophWnd, TVM_GETBKCOLOR, 0, 0) <> ImageList_GetBkColor(prophImageList) Then ret = True
    End If
  End If

  IsCustomDrawNecessary = ret
End Function

' prüft, ob der Punkt [<x>;<y>] im Scrollbereich liegt
Private Function IsWithinScrollingRegion(ByVal x As Long, ByVal y As Long) As Boolean
  Dim ret As Boolean

  With rcClient
    ret = (x <= propDragScrollZoneLeft)
    ret = ret Or (x >= (.Right - .Left) - propDragScrollZoneRight)
    ret = ret Or (y <= propDragScrollZoneTop)
    IsWithinScrollingRegion = ret Or (y >= (.Bottom - .Top) - propDragScrollZoneBottom)
  End With
End Function

' prüft, ob <hItem> existiert
Private Function ItemExists(ByVal hItem As Long, allowNull As AllowNullConstants, Optional allowTVIConsts As Boolean = False, Optional fastCheck As Boolean = True) As Boolean
  Dim itemData As TVITEMEX
  Dim ret As Boolean

  If hItem = -1 Then Exit Function

  If hItem = 0 Then
    If allowNull = AllowNullConstants.anAlways Then
      ret = True
    ElseIf allowNull = AllowNullConstants.anNever Then
      ret = False
    Else
      If allowNull And AllowNullConstants.anIfEmpty Then ret = (GetRootOrFirstItem = -1)

      If Not ret Then
        If allowNull And AllowNullConstants.anIfNoShowRoot Then ret = Not propShowRoot
      End If
    End If
  Else
    Select Case hItem
      Case TVI_FIRST, TVI_LAST, TVI_SORT, TVI_ROOT
        hItem = 0
        ret = allowTVIConsts
      Case Else
        If fastCheck Then
          ret = True
        Else
          With itemData
            .hItem = hItem
            .Mask = TVIF_HANDLE Or TVIF_INTEGRAL
            SendMessage prophWnd, TVM_GETITEM, 0, itemData
            ret = (.iIntegral <> 0)
          End With
        End If
    End Select
  End If

  ItemExists = ret
End Function

' gibt den Pfad von <hItem> zurück
Private Function ItemHandleToPath(ByVal hItem As Long) As String
  Dim c As Long
  Dim i As Long
  Dim IParent As IVBShellFolder
  Dim pIDL As Long
  Dim pIDLToParent As Long
  Dim ret As String
  Dim TVWItemInfo_Archive As ITEMINFO_ARCHIVE
  Dim TVWItemInfo_Common As ITEMINFO_COMMON
  Dim txt As String

  If inIDE Then Exit Function
  If Not ItemExists(hItem, AllowNullConstants.anIfNoShowRoot) Then Exit Function

  If ItemIsPartOfArchive(hItem) Then
    ' ein Archiv-Item
    ' -> der Pfad besteht aus 2 Teilen:
    '    - Pfad des Archivs
    '    - Pfad im Archiv
    GetItemInfo_Archive hItem, TVWItemInfo_Archive
    With TVWItemInfo_Archive
      ' Pfad des Archivs ermitteln
      ret = ItemHandleToPath(.hItem_ArchiveFile)
      ' Pfad im Archiv ermitteln
      While (hItem <> .hItem_ArchiveFile) And (IsWithinRange(hItem, -1, 0) = False)
        txt = "\" & GetItemText(hItem) & txt
        ' eine Ebene höher gehen
        hItem = ItemGetParentItem(hItem)
      Wend
      ItemHandleToPath = ret & ArchiveSignature & Mid$(txt, 2)
    End With
  Else
    ' ein normaler Item
    GetItemInfo_Common hItem, TVWItemInfo_Common
    With TVWItemInfo_Common
      Select Case .itemType
        Case ItemTypeConstants.itFSFile, ItemTypeConstants.itFSFolder
          ' wenn der Item über den Arbeitsplatz läuft, den kompletten Pfad nehmen
          #If Debuging Then
            If pIDLIncludesMyComputer(debugger, .pIDLToDesktop) Then
              ' kompletten Pfad nehmen
              SplitFullyQualifiedPIDL debugger, .pIDLToDesktop, IParent, pIDLToParent
              If (pIDLToParent <> 0) And Not (IParent Is Nothing) Then
                ItemHandleToPath = pIDLToPath(debugger, IParent, pIDLToParent)
              End If
              Set IParent = Nothing
            Else
              ' nur DisplayNames nehmen
              c = CountItemIDs(debugger, .pIDLToDesktop)
              ret = pIDLToDisplayName_Light(debugger, pIDL_Desktop)
              For i = 1 To c
                pIDL = CopyFirstItemIDs(debugger, .pIDLToDesktop, i)
                If pIDL = 0 Then Exit For
                SplitFullyQualifiedPIDL debugger, pIDL, IParent, pIDLToParent
                If (pIDLToParent <> 0) And Not (IParent Is Nothing) Then
                  ret = AddBackslash(ret, False) & pIDLToDisplayName(debugger, IParent, pIDLToParent, SHGDNConstants.SHGDN_INFOLDER)
                  ' don't free pIDL earlier because pIDLToParent is a part of it ;)
                  FreeItemIDList debugger, "ItemHandleToPath #1", pIDL
                Else
                  FreeItemIDList debugger, "ItemHandleToPath #2", pIDL
                  Exit For
                End If
                Set IParent = Nothing
              Next
              ItemHandleToPath = ret
            End If
          #Else
            If pIDLIncludesMyComputer(.pIDLToDesktop) Then
              ' kompletten Pfad nehmen
              SplitFullyQualifiedPIDL .pIDLToDesktop, IParent, pIDLToParent
              If (pIDLToParent <> 0) And Not (IParent Is Nothing) Then
                ItemHandleToPath = pIDLToPath(IParent, pIDLToParent)
              End If
              Set IParent = Nothing
            Else
              ' nur DisplayNames nehmen
              c = CountItemIDs(.pIDLToDesktop)
              ret = pIDLToDisplayName_Light(pIDL_Desktop)
              For i = 1 To c
                pIDL = CopyFirstItemIDs(.pIDLToDesktop, i)
                If pIDL = 0 Then Exit For
                SplitFullyQualifiedPIDL pIDL, IParent, pIDLToParent
                If (pIDLToParent <> 0) And Not (IParent Is Nothing) Then
                  ret = AddBackslash(ret, False) & pIDLToDisplayName(IParent, pIDLToParent, SHGDNConstants.SHGDN_INFOLDER)
                  ' don't free pIDL earlier because pIDLToParent is a part of it ;)
                  FreeItemIDList pIDL
                Else
                  FreeItemIDList pIDL
                  Exit For
                End If
                Set IParent = Nothing
              Next
              ItemHandleToPath = ret
            End If
          #End If
        Case ItemTypeConstants.itNonFSFile, ItemTypeConstants.itNonFSFolder
          ' CLSIDs nehmen, evt. mit DisplayNames mischen
          #If Debuging Then
            SplitFullyQualifiedPIDL debugger, .pIDLToDesktop, IParent, pIDLToParent
            If (pIDLToParent <> 0) And Not (IParent Is Nothing) Then
              txt = pIDLToPath(debugger, IParent, pIDLToParent)
              If Not IsFTPURL(txt) And Not IncludesCLSID(txt) Then
                txt = AddBackslash(ItemHandleToPath(ItemGetParentItem(hItem)), False) & pIDLToDisplayName(debugger, IParent, pIDLToParent, SHGDNConstants.SHGDN_INFOLDER)
              End If
              ItemHandleToPath = txt
            End If
            Set IParent = Nothing
          #Else
            SplitFullyQualifiedPIDL .pIDLToDesktop, IParent, pIDLToParent
            If (pIDLToParent <> 0) And Not (IParent Is Nothing) Then
              txt = pIDLToPath(IParent, pIDLToParent)
              If Not IsFTPURL(txt) And Not IncludesCLSID(txt) Then
                txt = AddBackslash(ItemHandleToPath(ItemGetParentItem(hItem)), False) & pIDLToDisplayName(IParent, pIDLToParent, SHGDNConstants.SHGDN_INFOLDER)
              End If
              ItemHandleToPath = txt
            End If
            Set IParent = Nothing
          #End If
      End Select
    End With
  End If
End Function

' gibt zurück, ob <hItem> mindestens 2 SubItems hat
Private Function ItemHasMoreThan1SubItems(ByVal hItem As Long) As Boolean
  Dim hSubItem As Long

  If Not ItemExists(hItem, AllowNullConstants.anIfNoShowRoot) Then Exit Function

  If ItemHasSubItems(hItem) Then
    ' jetzt die SubItems durchgehen
    hSubItem = ItemGetFirstSubItem(hItem)
    If Not IsWithinRange(hSubItem, -1, 0) Then
      ' jetzt nach einen 2. SubItem suchen
      hSubItem = ItemGetNextItem(hSubItem)
      ItemHasMoreThan1SubItems = Not IsWithinRange(hSubItem, -1, 0)
    End If
  End If
End Function

' gibt zurück, ob <hItem> SubItems hat
Private Function ItemHasSubItems(hItem As Long) As Boolean
  ItemHasSubItems = Not IsWithinRange(ItemGetFirstSubItem(hItem), -1, 0)
End Function

' gibt den Item unter <pt> zurück
' <Flags> enthält TVHT_*-Konstanten
Private Function ItemHitTest(pt As POINT, ByRef Flags As Long) As Long
  Dim Data As TVHITTESTINFO
  Dim ret As Long

  If prophWnd Then
    With Data
      .pt = pt
      ret = SendMessage(prophWnd, TVM_HITTEST, 0, Data)
      If (Flags And .Flags) = 0 Then ret = -1
      Flags = .Flags
    End With
  End If

  If ret = 0 Then ret = -1
  ItemHitTest = ret
End Function

' lädt alle(!) Items im Archiv <hArchiveFile>
' für ACE-Archive
Private Sub LoadArchive_ACE(ByVal hArchiveItem As Long, ByVal hArchiveFile As Long)
  Dim ArchiveSize As Long
  Dim attributes As Long
  Dim Blocksize As Integer
  Dim bufferSize As Integer
  Dim bytBuffer As Byte
  Dim BytesRead As Long
  Dim DispName As String
  Dim ext As String
  Dim FilePath As String
  Dim FilePointerPos As Long
  Dim FileSize As Long
  Dim hasChilds As Boolean
  Dim hItem As Long
  Dim hParentItem As Long
  Dim i As Integer
  Dim ImgNormal As Long
  Dim ImgNormal_Folder As Long
  Dim ImgSelected As Long
  Dim ImgSelected_Folder As Long
  Dim PathOnly As String
  Dim pos As Long
  Dim strBuffer As String
  Dim tmp As Long
  Dim TVWItemInfo As ITEMINFO_ARCHIVE

  If Not ItemExists(hArchiveItem, AllowNullConstants.anIfNoShowRoot) Then Exit Sub
  If hArchiveFile = INVALID_HANDLE_VALUE Then Exit Sub

  ArchiveSize = GetFileSize(hArchiveFile, 0)

  If (prophImageList <> 0) And Not propUseCustomIcons Then
    ImgNormal_Folder = IIf(useLargeIcons, DEFICON_FOLDER_LARGE, DEFICON_FOLDER_SMALL)
    If propUseSelectedIcons Then
      ImgSelected_Folder = IIf(useLargeIcons, DEFICON_OPENFOLDER_LARGE, DEFICON_OPENFOLDER_SMALL)
    Else
      ImgSelected_Folder = ImgNormal_Folder
    End If
  End If

  ' Signatur der Datei prüfen - die Zeichenfolge "**ACE**" suchen, von dieser Position 7 Bytes
  ' zurück muß der Archiv-Header beginnen
  ' um den Speicherbedarf zu verringern, nicht die gesamte Datei auf einmal, sondern in 1024-
  ' Bytes-Schritten einlesen
  Do
    strBuffer = String$(1024, Chr$(0))
    SetFilePointer hArchiveFile, FilePointerPos, 0, FILE_BEGIN
    ReadFileAsLong hArchiveFile, ByVal strBuffer, Len(strBuffer), BytesRead, 0
    If InStr(strBuffer, "**ACE**") Then
      ' jetzt noch prüfen, ob 3 Bytes vorher eine 0 steht
      pos = FilePointerPos + InStr(strBuffer, "**ACE**") - 3 - 1
      SetFilePointer hArchiveFile, pos, 0, FILE_BEGIN
      ReadFileAsLong hArchiveFile, ByVal VarPtr(bytBuffer), LenB(bytBuffer), BytesRead, 0
      If bytBuffer = 0 Then
        ' wir haben den Archiv-Header gefunden!
        FilePointerPos = pos - 4
        Exit Do
      Else
        ' Fehlschlag
        FilePointerPos = FilePointerPos + Len(strBuffer)
      End If
    Else
      FilePointerPos = FilePointerPos + Len(strBuffer)
    End If
  Loop Until FilePointerPos >= ArchiveSize
  If FilePointerPos >= ArchiveSize Then Exit Sub

  ' Bytes 3 und 4 enthalten die Größe des Archiv-Headers
  SetFilePointer hArchiveFile, FilePointerPos + 2, 0, FILE_BEGIN
  ReadFileAsLong hArchiveFile, ByVal VarPtr(Blocksize), LenB(Blocksize), BytesRead, 0
  FilePointerPos = FilePointerPos + 4 + Blocksize

  Do
    ' prüfen, ob dies ein File-Header ist
    ' -> Byte 5 muß 1 sein
    SetFilePointer hArchiveFile, FilePointerPos + 4, 0, FILE_BEGIN
    ReadFileAsLong hArchiveFile, ByVal VarPtr(bytBuffer), LenB(bytBuffer), BytesRead, 0
    If bytBuffer <> 1 Then Exit Do

    ' Bytes 3 und 4 enthalten die Größe des File-Headers
    SetFilePointer hArchiveFile, FilePointerPos + 2, 0, FILE_BEGIN
    ReadFileAsLong hArchiveFile, ByVal VarPtr(Blocksize), LenB(Blocksize), BytesRead, 0

    ' Bytes 8-11 enthalten die Größe der komprimierten Datei
    SetFilePointer hArchiveFile, FilePointerPos + 7, 0, FILE_BEGIN
    ReadFileAsLong hArchiveFile, ByVal VarPtr(FileSize), LenB(FileSize), BytesRead, 0

    ' Bytes 20-23 enthalten die Attribute der Datei
    SetFilePointer hArchiveFile, FilePointerPos + 19, 0, FILE_BEGIN
    ReadFileAsLong hArchiveFile, ByVal VarPtr(attributes), LenB(attributes), BytesRead, 0

    ' Bytes 34 und 35 enthalten die Länge des Dateinamens (mit Pfad)
    SetFilePointer hArchiveFile, FilePointerPos + 33, 0, FILE_BEGIN
    ReadFileAsLong hArchiveFile, ByVal VarPtr(bufferSize), LenB(bufferSize), BytesRead, 0
    FilePath = Space(bufferSize)
    ReadFileAsLong hArchiveFile, ByVal FilePath, Len(FilePath), BytesRead, 0
    FilePointerPos = FilePointerPos + 4 + Blocksize + FileSize

    ' das ACE-Format nutzt den OEM-Zeichensatz
    OemToChar FilePath, FilePath

    With TVWItemInfo
      hParentItem = 0
      ' Pfad anpassen und Item hinzufügen
      FilePath = RemoveBackslash(Replace(FilePath, "/", "\"), False)

      If attributes And FILE_ATTRIBUTE_DIRECTORY Then
        ' ein Ordner
        ' -> in 1 Schritt hinzufügen
        hItem = InsertArchiveFolder(hArchiveItem, FilePath)
        SetArchiveItemAttributes hItem, attributes
      Else
        ' eine Datei
        ' -> zunächst den Pfad abtrennen und hinzufügen
        PathOnly = GetPathName(FilePath)
        If PathOnly <> "" Then hParentItem = InsertArchiveFolder(hArchiveItem, PathOnly)
        If hParentItem = 0 Then
          hParentItem = hArchiveItem
        Else
          SetItemType hParentItem, ItemTypeConstants.itArchiveFolder
          SetItemIcons hParentItem, ImgNormal_Folder, ImgSelected_Folder
        End If

        DispName = GetFileName(FilePath)
        tmp = ItemGetFirstSubItem(hParentItem)
        Do While ItemExists(tmp, AllowNullConstants.anNever)
          If LCase$(GetItemText(tmp)) = LCase$(DispName) Then Exit Do
          tmp = ItemGetNextItem(tmp)
        Loop

        If Not ItemExists(tmp, AllowNullConstants.anNever) Then
          ' jetzt die Datei hinzufügen
          .hItem_ArchiveFile = hArchiveItem
          .attributes = attributes
          .itemType = CByte(ItemTypeConstants.itArchiveFile)
          hasChilds = False

          ext = "." & GetFileNameExtension(FilePath)
          If (prophImageList <> 0) And Not propUseCustomIcons Then
            ImgNormal = GetSysIconIndexFromExt(ext, False, useLargeIcons)
            If .itemType = ItemTypeConstants.itArchiveFile Then
              ImgSelected = ImgNormal
            Else
              ImgSelected = GetSysIconIndexFromExt(ext, propUseSelectedIcons, useLargeIcons)
            End If
          End If

          If ShouldShowItem_Archive(Me, DispName, TVWItemInfo.attributes) Then
            InsertTVWItem_Archive DispName, hParentItem, hasChilds, TVWItemInfo, ImgNormal, ImgSelected
          End If
        End If
      End If
    End With
  Loop Until FilePointerPos >= ArchiveSize
End Sub

' lädt alle(!) Items im Archiv <hArchiveFile>
' für BIN-Archive
Private Sub LoadArchive_BIN(ByVal hArchiveItem As Long, ByVal hArchiveFile As Long)
  Dim ArchiveSize As Long
  Dim bytBuffer As Byte
  Dim BytesRead As Long
  Dim FilePointerPos As Long
  Dim FilePointerPosJoliet As Long
  Dim FilePointerPosPVD As Long
  Dim Signatur As String
  Dim strBuffer As String
  Dim tmp As Long

  If Not ItemExists(hArchiveItem, AllowNullConstants.anIfNoShowRoot) Then Exit Sub
  If hArchiveFile = INVALID_HANDLE_VALUE Then Exit Sub

  ArchiveSize = GetFileSize(hArchiveFile, 0)

  ' Volume Descriptos suchen und dabei die Signatur der Datei prüfen - jeder Volume Descriptor startet
  ' mit einer Typ-ID gefolgt von "CD001"
  ' -> dies ist scheinbar die einzige Möglichkeit, zu prüfen, ob dies ein iso-Archiv ist
  ' um den Speicherbedarf zu verringern, nicht die gesamte Datei auf einmal, sondern in 2048-
  ' Byte-Schritten einlesen
  ' Signatur in String konvertieren
  Signatur = "CD001"
  FilePointerPos = 0
  FilePointerPosJoliet = -1
  FilePointerPosPVD = -1
  ' die ersten 16 Blöcke (à 2048 Bytes) überspringen (sind lt. Standard immer leer)
  SetFilePointer hArchiveFile, 32768, 0, FILE_BEGIN
  strBuffer = String$(2048, Chr$(0))
  Do
    ReadFileAsLong hArchiveFile, ByVal strBuffer, 2048, BytesRead, 0
    tmp = InStr(strBuffer, Signatur)
    If tmp Then
      ' wir haben einen Volume Descriptor gefunden - Typ ermitteln
      SetFilePointer hArchiveFile, FilePointerPos + tmp - 2, 0, FILE_BEGIN
      ReadFileAsLong hArchiveFile, ByVal VarPtr(bytBuffer), 1, BytesRead, 0
      Select Case bytBuffer
        Case 1
          ' Primary Volume Descriptor gefunden
          FilePointerPosPVD = FilePointerPos + tmp - 2
        Case 2
          ' Supplementary Volume Descriptor (= JOLIET) gefunden
          FilePointerPosJoliet = FilePointerPos + tmp - 2
          Exit Do
        Case 255
          ' Volume Descriptor Set Terminator gefunden
          Exit Do
      End Select
    End If
    FilePointerPos = FilePointerPos + 2048
    SetFilePointer hArchiveFile, FilePointerPos, 0, FILE_BEGIN
  Loop Until FilePointerPos >= ArchiveSize

  If FilePointerPosJoliet >= 0 Then
    LoadISOVolumeDescriptor_2353 hArchiveItem, hArchiveFile, FilePointerPosJoliet
  ElseIf FilePointerPosPVD >= 0 Then
    LoadISOVolumeDescriptor_2353 hArchiveItem, hArchiveFile, FilePointerPosPVD
  End If
End Sub

' lädt alle(!) Items im Archiv <hArchiveFile>
' für CAB-Archive
Private Sub LoadArchive_CAB(ByVal hArchiveItem As Long, ByVal hArchiveFile As Long)
  Dim attributes As Integer
  Dim BytesRead As Long
  Dim CountEntries As Integer
  Dim DispName As String
  Dim ext As String
  Dim FilePath As String
  Dim FilePointerPos As Long
  Dim hasChilds As Boolean
  Dim hParentItem As Long
  Dim i As Integer
  Dim ImgNormal As Long
  Dim ImgNormal_Folder As Long
  Dim ImgSelected As Long
  Dim ImgSelected_Folder As Long
  Dim PathOnly As String
  Dim strBuffer As String
  Dim tmp As Long
  Dim TVWItemInfo As ITEMINFO_ARCHIVE

  If Not ItemExists(hArchiveItem, AllowNullConstants.anIfNoShowRoot) Then Exit Sub
  If hArchiveFile = INVALID_HANDLE_VALUE Then Exit Sub

  If (prophImageList <> 0) And Not propUseCustomIcons Then
    ImgNormal_Folder = IIf(useLargeIcons, DEFICON_FOLDER_LARGE, DEFICON_FOLDER_SMALL)
    If propUseSelectedIcons Then
      ImgSelected_Folder = IIf(useLargeIcons, DEFICON_OPENFOLDER_LARGE, DEFICON_OPENFOLDER_SMALL)
    Else
      ImgSelected_Folder = ImgNormal_Folder
    End If
  End If

  ' Signatur der Datei prüfen - die ersten 4 Bytes müssen "MSCF" ergeben
  FilePointerPos = 0
  SetFilePointer hArchiveFile, FilePointerPos, 0, FILE_BEGIN
  strBuffer = String$(4, Chr$(0))
  ReadFileAsLong hArchiveFile, ByVal strBuffer, Len(strBuffer), BytesRead, 0
  If strBuffer <> "MSCF" Then Exit Sub

  ' Bytes 29 und 30 enthalten die Anzahl der Einträge
  FilePointerPos = 28
  SetFilePointer hArchiveFile, FilePointerPos, 0, FILE_BEGIN
  ReadFileAsLong hArchiveFile, ByVal VarPtr(CountEntries), LenB(CountEntries), BytesRead, 0

  ' Bytes 17-20 enthalten den Offset des 1. Eintrags
  FilePointerPos = 16
  SetFilePointer hArchiveFile, FilePointerPos, 0, FILE_BEGIN
  ReadFileAsLong hArchiveFile, ByVal VarPtr(FilePointerPos), LenB(FilePointerPos), BytesRead, 0

  For i = 1 To CountEntries
    ' Bytes 15 und 16 enthalten die Attribute des Eintrags
    SetFilePointer hArchiveFile, FilePointerPos + 14, 0, FILE_BEGIN
    ReadFileAsLong hArchiveFile, ByVal VarPtr(attributes), LenB(attributes), BytesRead, 0

    ' bei Byte 17 fängt der Dateiname (mit Pfad) an
    FilePath = String$(MAX_PATH, Chr$(0))
    ReadFileAsLong hArchiveFile, ByVal FilePath, Len(FilePath), BytesRead, 0
    FilePath = Left$(FilePath, lstrlenA(FilePath))
    FilePointerPos = FilePointerPos + 16 + Len(FilePath) + 1   ' abschließende 0

    With TVWItemInfo
      hParentItem = 0
      ' Pfad anpassen und Item hinzufügen
      FilePath = RemoveBackslash(Replace(FilePath, "/", "\"), False)

      ' zunächst den Pfad abtrennen und hinzufügen
      PathOnly = GetPathName(FilePath)
      If PathOnly <> "" Then hParentItem = InsertArchiveFolder(hArchiveItem, PathOnly)
      If hParentItem = 0 Then
        hParentItem = hArchiveItem
      Else
        SetItemType hParentItem, ItemTypeConstants.itArchiveFolder
        SetItemIcons hParentItem, ImgNormal_Folder, ImgSelected_Folder
      End If

      DispName = GetFileName(FilePath)
      tmp = ItemGetFirstSubItem(hParentItem)
      Do While ItemExists(tmp, AllowNullConstants.anNever)
        If LCase$(GetItemText(tmp)) = LCase$(DispName) Then Exit Do
        tmp = ItemGetNextItem(tmp)
      Loop

      If Not ItemExists(tmp, AllowNullConstants.anNever) Then
        ' jetzt die Datei hinzufügen
        .hItem_ArchiveFile = hArchiveItem
        .attributes = attributes
        .itemType = CByte(ItemTypeConstants.itArchiveFile)
        hasChilds = False

        ext = "." & GetFileNameExtension(FilePath)
        If (prophImageList <> 0) And Not propUseCustomIcons Then
          ImgNormal = GetSysIconIndexFromExt(ext, False, useLargeIcons)
          If .itemType = ItemTypeConstants.itArchiveFile Then
            ImgSelected = ImgNormal
          Else
            ImgSelected = GetSysIconIndexFromExt(ext, propUseSelectedIcons, useLargeIcons)
          End If
        End If

        If ShouldShowItem_Archive(Me, DispName, TVWItemInfo.attributes) Then
          InsertTVWItem_Archive DispName, hParentItem, hasChilds, TVWItemInfo, ImgNormal, ImgSelected
        End If
      End If
    End With
  Next
End Sub

' lädt alle(!) Items im Archiv <hArchiveFile>
' für ISO-Archive
Private Sub LoadArchive_ISO(ByVal hArchiveItem As Long, ByVal hArchiveFile As Long)
  Dim ArchiveSize As Long
  Dim bytBuffer As Byte
  Dim BytesRead As Long
  Dim FilePointerPos As Long
  Dim FilePointerPosJoliet As Long
  Dim FilePointerPosPVD As Long
  Dim Signatur As String
  Dim strBuffer As String
  Dim tmp As Long

  If Not ItemExists(hArchiveItem, AllowNullConstants.anIfNoShowRoot) Then Exit Sub
  If hArchiveFile = INVALID_HANDLE_VALUE Then Exit Sub

  ArchiveSize = GetFileSize(hArchiveFile, 0)

  ' Volume Descriptos suchen und dabei die Signatur der Datei prüfen - jeder Volume Descriptor startet
  ' mit einer Typ-ID gefolgt von "CD001"
  ' -> dies ist scheinbar die einzige Möglichkeit, zu prüfen, ob dies ein iso-Archiv ist
  ' um den Speicherbedarf zu verringern, nicht die gesamte Datei auf einmal, sondern in 2048-
  ' Byte-Schritten einlesen
  ' Signatur in String konvertieren
  Signatur = "CD001"
  FilePointerPos = 0
  FilePointerPosJoliet = -1
  FilePointerPosPVD = -1
  ' die ersten 16 Blöcke (à 2048 Bytes) überspringen (sind lt. Standard immer leer)
  SetFilePointer hArchiveFile, 32768, 0, FILE_BEGIN
  strBuffer = String$(2048, Chr$(0))
  Do
    ReadFileAsLong hArchiveFile, ByVal strBuffer, 2048, BytesRead, 0
    tmp = InStr(strBuffer, Signatur)
    If tmp Then
      ' wir haben einen Volume Descriptor gefunden - Typ ermitteln
      SetFilePointer hArchiveFile, FilePointerPos + tmp - 2, 0, FILE_BEGIN
      ReadFileAsLong hArchiveFile, ByVal VarPtr(bytBuffer), 1, BytesRead, 0
      Select Case bytBuffer
        Case 1
          ' Primary Volume Descriptor gefunden
          FilePointerPosPVD = FilePointerPos + tmp - 2
        Case 2
          ' Supplementary Volume Descriptor (= JOLIET) gefunden
          FilePointerPosJoliet = FilePointerPos + tmp - 2
          Exit Do
        Case 255
          ' Volume Descriptor Set Terminator gefunden
          Exit Do
      End Select
    End If
    FilePointerPos = FilePointerPos + 2048
    SetFilePointer hArchiveFile, FilePointerPos, 0, FILE_BEGIN
  Loop Until FilePointerPos >= ArchiveSize

  If FilePointerPosJoliet >= 0 Then
    LoadISOVolumeDescriptor_2048 hArchiveItem, hArchiveFile, FilePointerPosJoliet
  ElseIf FilePointerPosPVD >= 0 Then
    LoadISOVolumeDescriptor_2048 hArchiveItem, hArchiveFile, FilePointerPosPVD
  End If
End Sub

' lädt alle(!) Items im Archiv <hArchiveFile>
' für RAR-Archive
Private Sub LoadArchive_RAR(ByVal hArchiveItem As Long, ByVal hArchiveFile As Long)
  Dim ArchiveSize As Long
  Dim attributes As Long
  Dim BlockFlags As Integer
  Dim Blocksize As Integer
  Dim BlockType As Byte
  Dim BytesRead As Long
  Dim DispName As String
  Dim ext As String
  Dim FileNameLength As Integer
  Dim FilePath As String
  Dim FilePointerPos As Long
  Dim hasChilds As Boolean
  Dim hItem As Long
  Dim hParentItem As Long
  Dim i As Integer
  Dim ImgNormal As Long
  Dim ImgNormal_Folder As Long
  Dim ImgSelected As Long
  Dim ImgSelected_Folder As Long
  Dim PathOnly As String
  Dim Signatur As String
  Dim strBuffer As String
  Dim SubBlockSize As Long
  Dim tmp As Long
  Dim TVWItemInfo As ITEMINFO_ARCHIVE

  If Not ItemExists(hArchiveItem, AllowNullConstants.anIfNoShowRoot) Then Exit Sub
  If hArchiveFile = INVALID_HANDLE_VALUE Then Exit Sub

  ArchiveSize = GetFileSize(hArchiveFile, 0)

  If (prophImageList <> 0) And Not propUseCustomIcons Then
    ImgNormal_Folder = IIf(useLargeIcons, DEFICON_FOLDER_LARGE, DEFICON_FOLDER_SMALL)
    If propUseSelectedIcons Then
      ImgSelected_Folder = IIf(useLargeIcons, DEFICON_OPENFOLDER_LARGE, DEFICON_OPENFOLDER_SMALL)
    Else
      ImgSelected_Folder = ImgNormal_Folder
    End If
  End If

  ' Signatur der Datei prüfen - am Anfang des Archivs muß ein Marker-Block stehen
  ' -> er ist 7 Bytes lang und hat immer die Sequenz &H52, &H61, &H72, &H21, &H1A, &H07, &H00
  Signatur = Chr$(&H52) & Chr$(&H61) & Chr$(&H72) & Chr$(&H21) & Chr$(&H1A) & Chr$(&H7) & Chr$(&H0)
  strBuffer = String$(7, Chr$(0))
  SetFilePointer hArchiveFile, FilePointerPos, 0, FILE_BEGIN
  ReadFileAsLong hArchiveFile, ByVal strBuffer, Len(strBuffer), BytesRead, 0
  If strBuffer <> Signatur Then Exit Sub

  Do
    ' den nächsten Block prüfen
    ' Byte 3 enthält den Header-Typ
    SetFilePointer hArchiveFile, FilePointerPos + 2, 0, FILE_BEGIN
    ReadFileAsLong hArchiveFile, ByVal VarPtr(BlockType), LenB(BlockType), BytesRead, 0

    ' Bytes 4 und 5 enthalten die Flags des Blockes
    ReadFileAsLong hArchiveFile, ByVal VarPtr(BlockFlags), LenB(BlockFlags), BytesRead, 0

    ' Bytes 6 und 7 enthalten die Größe des Blockes
    ReadFileAsLong hArchiveFile, ByVal VarPtr(Blocksize), LenB(Blocksize), BytesRead, 0

    SubBlockSize = 0
    If BlockFlags And &H8000 Then
      ' Bytes 8-11 enthalten die Größe des SubBlocks
      ReadFileAsLong hArchiveFile, ByVal VarPtr(SubBlockSize), LenB(SubBlockSize), BytesRead, 0
    End If

    If BlockType = &H74 Then
      ' ein File-Header
      ' Bytes 27 und 28 enthalten die Länge des Dateinamens (mit Pfad)
      SetFilePointer hArchiveFile, FilePointerPos + 26, 0, FILE_BEGIN
      ReadFileAsLong hArchiveFile, ByVal VarPtr(FileNameLength), LenB(FileNameLength), BytesRead, 0

      ' Bytes 29-32 enthalten die Attribute des Eintrags
      ReadFileAsLong hArchiveFile, ByVal VarPtr(attributes), LenB(attributes), BytesRead, 0

      If BlockFlags And &H100 Then
        ' für die SubBlock-Größe werden an dieser Stelle 4 weitere Bytes belegt
        ' (und danach noch 4 Bytes für die Original-Größe der Daten)
        ' -> TODO: alles auf 64 Bit umstellen
        Exit Do
      End If

      ' bei Byte 33 fängt der Dateiname (mit Pfad) an
      FilePath = String$(FileNameLength, Chr$(0))
      ReadFileAsLong hArchiveFile, ByVal FilePath, Len(FilePath), BytesRead, 0

      ' das RAR-Format nutzt den OEM-Zeichensatz
      OemToChar FilePath, FilePath

      With TVWItemInfo
        hParentItem = 0
        ' Pfad anpassen und Item hinzufügen
        FilePath = RemoveBackslash(Replace(FilePath, "/", "\"), False)

        If attributes And FILE_ATTRIBUTE_DIRECTORY Then
          ' ein Ordner
          ' -> in 1 Schritt hinzufügen
          hItem = InsertArchiveFolder(hArchiveItem, FilePath)
          SetArchiveItemAttributes hItem, attributes
        Else
          ' eine Datei
          ' -> zunächst den Pfad abtrennen und hinzufügen
          PathOnly = GetPathName(FilePath)
          If PathOnly <> "" Then hParentItem = InsertArchiveFolder(hArchiveItem, PathOnly)
          If hParentItem = 0 Then
            hParentItem = hArchiveItem
          Else
            SetItemType hParentItem, ItemTypeConstants.itArchiveFolder
            SetItemIcons hParentItem, ImgNormal_Folder, ImgSelected_Folder
          End If

          DispName = GetFileName(FilePath)
          tmp = ItemGetFirstSubItem(hParentItem)
          Do While ItemExists(tmp, AllowNullConstants.anNever)
            If LCase$(GetItemText(tmp)) = LCase$(DispName) Then Exit Do
            tmp = ItemGetNextItem(tmp)
          Loop

          If Not ItemExists(tmp, AllowNullConstants.anNever) Then
            ' jetzt die Datei hinzufügen
            .hItem_ArchiveFile = hArchiveItem
            .attributes = attributes
            .itemType = CByte(ItemTypeConstants.itArchiveFile)
            hasChilds = False

            ext = "." & GetFileNameExtension(FilePath)
            If (prophImageList <> 0) And Not propUseCustomIcons Then
              ImgNormal = GetSysIconIndexFromExt(ext, False, useLargeIcons)
              If .itemType = ItemTypeConstants.itArchiveFile Then
                ImgSelected = ImgNormal
              Else
                ImgSelected = GetSysIconIndexFromExt(ext, propUseSelectedIcons, useLargeIcons)
              End If
            End If

            If ShouldShowItem_Archive(Me, DispName, TVWItemInfo.attributes) Then
              InsertTVWItem_Archive DispName, hParentItem, hasChilds, TVWItemInfo, ImgNormal, ImgSelected
            End If
          End If
        End If
      End With
    End If

    FilePointerPos = FilePointerPos + Blocksize + SubBlockSize
  Loop Until FilePointerPos >= ArchiveSize
End Sub

#If NewArchiveSupport Then
' lädt alle(!) Items im Archiv <hArchiveFile>
' für TAR-Archive
Private Sub LoadArchive_TAR(ByVal hArchiveItem As Long, ByVal hArchiveFile As Long)
  Dim ArchiveSize As Long
  Dim attributes As Integer
  Dim BytesRead As Long
  Dim DispName As String
  Dim ext As String
  Dim FilePath As String
  Dim FilePointerPos As Long
  Dim FileSize As Long
  Dim FileType As Byte
  Dim hasChilds As Boolean
  Dim hItem As Long
  Dim hParentItem As Long
  Dim ImgNormal As Long
  Dim ImgNormal_Folder As Long
  Dim ImgSelected As Long
  Dim ImgSelected_Folder As Long
  Dim longFileName As Boolean
  Dim NumberOfBlocks As Long
  Dim PathOnly As String
  Dim skipFile As Boolean
  Dim strBuffer As String
  Dim tmp As Long
  Dim TVWItemInfo As ITEMINFO_ARCHIVE

  If Not ItemExists(hArchiveItem, AllowNullConstants.anIfNoShowRoot) Then Exit Sub
  If hArchiveFile = INVALID_HANDLE_VALUE Then Exit Sub

  ArchiveSize = GetFileSize(hArchiveFile, 0)

  If (prophImageList <> 0) And Not propUseCustomIcons Then
    ImgNormal_Folder = IIf(useLargeIcons, DEFICON_FOLDER_LARGE, DEFICON_FOLDER_SMALL)
    If propUseSelectedIcons Then
      ImgSelected_Folder = IIf(useLargeIcons, DEFICON_OPENFOLDER_LARGE, DEFICON_OPENFOLDER_SMALL)
    Else
      ImgSelected_Folder = ImgNormal_Folder
    End If
  End If

  ' tar-Archive haben keine Signatur

  Do
    ' Byte 156 enthält den Dateityp
    SetFilePointer hArchiveFile, FilePointerPos + 156, 0, FILE_BEGIN
    ReadFileAsLong hArchiveFile, ByVal VarPtr(FileType), LenB(FileType), BytesRead, 0
    skipFile = True
    longFileName = False
    Select Case FileType
      Case 53     ' Asc("5")
        attributes = FILE_ATTRIBUTE_DIRECTORY
        skipFile = False
      Case 0
        attributes = 0
        skipFile = False
      Case 48     ' Asc("0")
        attributes = 0
        skipFile = False
      Case 49     ' Asc("1")
        attributes = 0
        skipFile = False
      Case 50     ' Asc("2")
        attributes = 0
        skipFile = False
      Case 51     ' Asc("3")
        attributes = 0
        skipFile = False
      Case 52     ' Asc("4")
        attributes = 0
        skipFile = False
      Case 54     ' Asc("6")
        attributes = 0
        skipFile = False
      Case 55     ' Asc("7")
        attributes = 0
        skipFile = False
      Case 68     ' Asc("D")
        ' TODO: The tar specs sound like this record contains multiple filenames?!
        attributes = FILE_ATTRIBUTE_DIRECTORY
        skipFile = False
      Case 75     ' Asc("K")
        ' TODO...
        skipFile = False
      Case 76     ' Asc("L")
        ' Header besteht aus 3 Blöcken:
        '   (1) Info, dass es ein langer Dateiname ist + Länge des Dateinamens
        '   (2) Dateiname
        '   (3) normaler Header (mit abgeschnittenem Dateinamen) + Content-Länge
        longFileName = True

        ' Byte 156 des 3. Blocks enthält den Dateityp
        SetFilePointer hArchiveFile, FilePointerPos + 1180, 0, FILE_BEGIN
        ReadFileAsLong hArchiveFile, ByVal VarPtr(FileType), LenB(FileType), BytesRead, 0
        Select Case FileType
          Case 53     ' Asc("5")
            attributes = FILE_ATTRIBUTE_DIRECTORY
            skipFile = False
          Case 0
            attributes = 0
            skipFile = False
          Case 48     ' Asc("0")
            attributes = 0
            skipFile = False
          Case 49     ' Asc("1")
            attributes = 0
            skipFile = False
          Case 50     ' Asc("2")
            attributes = 0
            skipFile = False
          Case 51     ' Asc("3")
            attributes = 0
            skipFile = False
          Case 52     ' Asc("4")
            attributes = 0
            skipFile = False
          Case 54     ' Asc("6")
            attributes = 0
            skipFile = False
          Case 55     ' Asc("7")
            attributes = 0
            skipFile = False
          Case 68     ' Asc("D")
            ' TODO: The tar specs sound like this record contains multiple filenames?!
            attributes = FILE_ATTRIBUTE_DIRECTORY
            skipFile = False
          Case 75     ' Asc("K")
            ' TODO...
            skipFile = False
          Case 76     ' Asc("L")
            ' TODO: Can this happen?
            skipFile = False
          Case 77     ' Asc("M")
            ' continued multivolume file - skip
          Case 78     ' Asc("N")
            ' TODO...
            skipFile = False
          Case 83     ' Asc("S")
            ' TODO...
            skipFile = False
          Case 86     ' Asc("V")
            ' volume header - skip
        End Select
      Case 77     ' Asc("M")
        ' continued multivolume file - skip
      Case 78     ' Asc("N")
        ' TODO...
        skipFile = False
      Case 83     ' Asc("S")
        ' TODO...
        skipFile = False
      Case 86     ' Asc("V")
        ' volume header - skip
    End Select

    If Not skipFile Then
      If longFileName Then
        ' der nächste Block enthält den Dateinamen
        ' Länge des Dateinamens auslesen (Bytes 124-135 als Oktalzahl)
        strBuffer = String$(12, Chr$(0))
        SetFilePointer hArchiveFile, FilePointerPos + 124, 0, FILE_BEGIN
        ReadFileAsLong hArchiveFile, ByVal strBuffer, Len(strBuffer), BytesRead, 0
        tmp = OctalToDecimal(Left$(strBuffer, lstrlenA(strBuffer)))

        SetFilePointer hArchiveFile, FilePointerPos + 512, 0, FILE_BEGIN
        FilePath = String$(tmp, Chr$(0))
      Else
        ' die ersten 100 Bytes eines Headers enthalten den Dateinamen
        FilePath = String$(100, Chr$(0))
        SetFilePointer hArchiveFile, FilePointerPos, 0, FILE_BEGIN
      End If
      ReadFileAsLong hArchiveFile, ByVal FilePath, Len(FilePath), BytesRead, 0
      FilePath = Left$(FilePath, lstrlenA(FilePath))
      If FilePath = "" Then Exit Do

      ' Bytes 124-135 enthalten die Dateigröße als Oktalzahl
      strBuffer = String$(12, Chr$(0))
      If longFileName Then
        SetFilePointer hArchiveFile, FilePointerPos + 1146, 0, FILE_BEGIN
      Else
        SetFilePointer hArchiveFile, FilePointerPos + 124, 0, FILE_BEGIN
      End If
      ReadFileAsLong hArchiveFile, ByVal strBuffer, Len(strBuffer), BytesRead, 0
      FileSize = OctalToDecimal(Left$(strBuffer, lstrlenA(strBuffer)))

      With TVWItemInfo
        hParentItem = 0
        ' Pfad anpassen und Item hinzufügen
        FilePath = RemoveBackslash(Replace(FilePath, "/", "\"), False)

        If attributes And FILE_ATTRIBUTE_DIRECTORY Then
          ' ein Ordner
          ' -> in 1 Schritt hinzufügen
          hItem = InsertArchiveFolder(hArchiveItem, FilePath)
          SetArchiveItemAttributes hItem, attributes
        Else
          ' eine Datei
          ' -> zunächst den Pfad abtrennen und hinzufügen
          PathOnly = GetPathName(FilePath)
          If PathOnly <> "" Then hParentItem = InsertArchiveFolder(hArchiveItem, PathOnly)
          If hParentItem = 0 Then
            hParentItem = hArchiveItem
          Else
            SetItemType hParentItem, ItemTypeConstants.itArchiveFolder
            SetItemIcons hParentItem, ImgNormal_Folder, ImgSelected_Folder
          End If

          DispName = GetFileName(FilePath)
          tmp = ItemGetFirstSubItem(hParentItem)
          Do While ItemExists(tmp, AllowNullConstants.anNever)
            If LCase$(GetItemText(tmp)) = LCase$(DispName) Then Exit Do
            tmp = ItemGetNextItem(tmp)
          Loop

          If Not ItemExists(tmp, AllowNullConstants.anNever) Then
            ' jetzt die Datei hinzufügen
            .hItem_ArchiveFile = hArchiveItem
            .attributes = attributes
            .itemType = CByte(ItemTypeConstants.itArchiveFile)
            hasChilds = False

            ext = "." & GetFileNameExtension(FilePath)
            If (prophImageList <> 0) And Not propUseCustomIcons Then
              ImgNormal = GetSysIconIndexFromExt(ext, False, useLargeIcons)
              If .itemType = ItemTypeConstants.itArchiveFile Then
                ImgSelected = ImgNormal
              Else
                ImgSelected = GetSysIconIndexFromExt(ext, propUseSelectedIcons, useLargeIcons)
              End If
            End If

            If ShouldShowItem_Archive(Me, DispName, TVWItemInfo.attributes) Then
              InsertTVWItem_Archive DispName, hParentItem, hasChilds, TVWItemInfo, ImgNormal, ImgSelected
            End If
          End If
        End If
      End With
    End If

    ' ein Block ist 512 Bytes groß
    ' auch die Daten selbst bestehen aus 512-Bytes-Blöcken
    NumberOfBlocks = 1     ' Header
    If longFileName Then NumberOfBlocks = NumberOfBlocks + 2
    If FileSize > 0 Then
      NumberOfBlocks = NumberOfBlocks + (FileSize \ 512)        ' Content
      If FileSize Mod 512 > 0 Then NumberOfBlocks = NumberOfBlocks + 1
    End If
    FilePointerPos = FilePointerPos + NumberOfBlocks * 512
  Loop Until FilePointerPos >= ArchiveSize
End Sub
#End If

' lädt alle(!) Items im Archiv <hArchiveFile>
' für JAR- und ZIP-Archive
Private Sub LoadArchive_ZIP(ByVal hArchiveItem As Long, ByVal hArchiveFile As Long)
  Dim ArchiveSize As Long
  Dim attributes As Integer
  Dim BytesRead As Long
  Dim DispName As String
  Dim ext As String
  Dim ExtraDataLength As Integer
  Dim FileCommentLength As Integer
  Dim FileNameLength As Integer
  Dim FilePath As String
  Dim FilePointerPos As Long
  Dim hasChilds As Boolean
  Dim hItem As Long
  Dim hParentItem As Long
  Dim i As Integer
  Dim ImgNormal As Long
  Dim ImgNormal_Folder As Long
  Dim ImgSelected As Long
  Dim ImgSelected_Folder As Long
  Dim lngBuffer As Long
  Dim PathOnly As String
  Dim Signatur As String
  Dim strBuffer As String
  Dim tmp As Long
  Dim TVWItemInfo As ITEMINFO_ARCHIVE

  If Not ItemExists(hArchiveItem, AllowNullConstants.anIfNoShowRoot) Then Exit Sub
  If hArchiveFile = INVALID_HANDLE_VALUE Then Exit Sub

  ArchiveSize = GetFileSize(hArchiveFile, 0)

  If (prophImageList <> 0) And Not propUseCustomIcons Then
    ImgNormal_Folder = IIf(useLargeIcons, DEFICON_FOLDER_LARGE, DEFICON_FOLDER_SMALL)
    If propUseSelectedIcons Then
      ImgSelected_Folder = IIf(useLargeIcons, DEFICON_OPENFOLDER_LARGE, DEFICON_OPENFOLDER_SMALL)
    Else
      ImgSelected_Folder = ImgNormal_Folder
    End If
  End If

  ' Signatur der Datei prüfen - am Ende der Datei muß irgendwo &H06054B50 (Signatur des
  ' Archiv-Headers) stehen
  ' -> dies ist anscheinend die einzige Möglichkeit, zu prüfen, ob dies ein zip-Archiv ist
  ' um den Speicherbedarf zu verringern, nicht die gesamte Datei auf einmal, sondern in 2048-
  ' Byte-Schritten einlesen
  Signatur = Chr(&H50) & Chr(&H4B) & Chr(&H5) & Chr(&H6)
  ' von hinten anfangen, da der Archiv-Header am Dateiende steht
  FilePointerPos = ArchiveSize
  Do
    strBuffer = Space(2048)
    FilePointerPos = FilePointerPos - Len(strBuffer)
    If FilePointerPos < 0 Then FilePointerPos = 0
    SetFilePointer hArchiveFile, FilePointerPos, 0, FILE_BEGIN
    ReadFileAsLong hArchiveFile, ByVal strBuffer, Len(strBuffer), BytesRead, 0
    If InStr(strBuffer, Signatur) Then
      ' wir haben ihn gefunden - genaue Position ermitteln
      FilePointerPos = FilePointerPos + InStrRev(strBuffer, Signatur)
      Exit Do
    End If
  Loop Until Trim(strBuffer) = ""
  If Trim(strBuffer) = "" Then Exit Sub

  ' Bytes 17-20 enthalten die Position des 1. TOC-Eintrags
  FilePointerPos = FilePointerPos - 1 + 16
  SetFilePointer hArchiveFile, FilePointerPos, 0, FILE_BEGIN
  ReadFileAsLong hArchiveFile, ByVal VarPtr(FilePointerPos), LenB(FilePointerPos), BytesRead, 0

  Do
    ' prüfen, ob dies ein TOC-Eintrag ist
    ' -> Signatur muß &H2014B50 sein
    SetFilePointer hArchiveFile, FilePointerPos, 0, FILE_BEGIN
    ReadFileAsLong hArchiveFile, ByVal VarPtr(lngBuffer), LenB(lngBuffer), BytesRead, 0
    If lngBuffer <> &H2014B50 Then Exit Do

    ' Bytes 29 und 30 enthalten die Länge des Dateinamens (mit Pfad)
    SetFilePointer hArchiveFile, FilePointerPos + 28, 0, FILE_BEGIN
    ReadFileAsLong hArchiveFile, ByVal VarPtr(FileNameLength), LenB(FileNameLength), BytesRead, 0

    ' Bytes 31 und 32 enthalten die Länge der Extra-Daten
    ReadFileAsLong hArchiveFile, ByVal VarPtr(ExtraDataLength), LenB(ExtraDataLength), BytesRead, 0

    ' Bytes 33 und 34 enthalten die Länge des File-Kommentars
    ReadFileAsLong hArchiveFile, ByVal VarPtr(FileCommentLength), LenB(FileCommentLength), BytesRead, 0

    ' Bytes 39-42 enthalten die Attribute des Eintrags
    SetFilePointer hArchiveFile, FilePointerPos + 38, 0, FILE_BEGIN
    ReadFileAsLong hArchiveFile, ByVal VarPtr(attributes), LenB(attributes), BytesRead, 0

    ' bei Byte 47 fängt der Dateiname (mit Pfad) an
    FilePath = String$(FileNameLength, Chr$(0))
    SetFilePointer hArchiveFile, FilePointerPos + 46, 0, FILE_BEGIN
    ReadFileAsLong hArchiveFile, ByVal FilePath, Len(FilePath), BytesRead, 0
    FilePointerPos = FilePointerPos + 46 + FileNameLength + ExtraDataLength + FileCommentLength

    ' das ZIP-Format nutzt einen OEM-Zeichensatz
    OemToChar FilePath, FilePath

    With TVWItemInfo
      hParentItem = 0
      ' Pfad anpassen und Item hinzufügen
      FilePath = RemoveBackslash(Replace(FilePath, "/", "\"), False)

      If attributes And FILE_ATTRIBUTE_DIRECTORY Then
        ' ein Ordner
        ' -> in 1 Schritt hinzufügen
        hItem = InsertArchiveFolder(hArchiveItem, FilePath)
        SetArchiveItemAttributes hItem, attributes
      Else
        ' eine Datei
        ' -> zunächst den Pfad abtrennen und hinzufügen
        PathOnly = GetPathName(FilePath)
        If PathOnly <> "" Then hParentItem = InsertArchiveFolder(hArchiveItem, PathOnly)
        If hParentItem = 0 Then
          hParentItem = hArchiveItem
        Else
          SetItemType hParentItem, ItemTypeConstants.itArchiveFolder
          SetItemIcons hParentItem, ImgNormal_Folder, ImgSelected_Folder
        End If

        DispName = GetFileName(FilePath)
        tmp = ItemGetFirstSubItem(hParentItem)
        Do While ItemExists(tmp, AllowNullConstants.anNever)
          If LCase$(GetItemText(tmp)) = LCase$(DispName) Then Exit Do
          tmp = ItemGetNextItem(tmp)
        Loop

        If Not ItemExists(tmp, AllowNullConstants.anNever) Then
          ' jetzt die Datei hinzufügen
          .hItem_ArchiveFile = hArchiveItem
          .attributes = attributes
          .itemType = CByte(ItemTypeConstants.itArchiveFile)
          hasChilds = False

          ext = "." & GetFileNameExtension(FilePath)
          If (prophImageList <> 0) And Not propUseCustomIcons Then
            ImgNormal = GetSysIconIndexFromExt(ext, False, useLargeIcons)
            If .itemType = ItemTypeConstants.itArchiveFile Then
              ImgSelected = ImgNormal
            Else
              ImgSelected = GetSysIconIndexFromExt(ext, propUseSelectedIcons, useLargeIcons)
            End If
          End If

          If ShouldShowItem_Archive(Me, DispName, TVWItemInfo.attributes) Then
            InsertTVWItem_Archive DispName, hParentItem, hasChilds, TVWItemInfo, ImgNormal, ImgSelected
          End If
        End If
      End If
    End With
  Loop Until FilePointerPos >= ArchiveSize
End Sub

' füllt das Control mit ein paar Beispiel-Items
Private Sub LoadIDEItems()
  Dim hItem As Long
  Dim hItem2 As Long

  ClearTVWItems

  'hItem = insertTVWItem_Common("TimoSoft ExplorerTreeView-Control", 0, True, DummyItemInfo, 0, 0, 0, True)
  hItem = InsertIDEItem("TimoSoft ExplorerTreeView-Control", 0, True)
  'insertTVWItem_Common "Version " & Version, hItem, False, DummyItemInfo, 0, 0, 0, True
  InsertIDEItem "Version " & Version, hItem, False
  If Not IsRelease Then
    'insertTVWItem_Common "Dies ist eine Debuging-Version", hItem, False, DummyItemInfo, 0, 0, 0, True
    InsertIDEItem "Dies ist eine Debuging-Version", hItem, False
  End If

  ItemExpand hItem

  'insertTVWItem_Common "", 0, False, DummyItemInfo, 0, 0, 0, True
  InsertIDEItem "", 0, False
  'hItem2 = insertTVWItem_Common("Weitere ActiveX-Controls von Timo ""TimoSoft"" Kunze:", 0, True, DummyItemInfo, 0, 0, 0, True)
  hItem2 = InsertIDEItem("Weitere ActiveX-Controls von Timo ""TimoSoft"" Kunze:", 0, True)
  'InsertTVWItem_Common "Animation", hItem2, False, DummyItemInfo, 0, True
  InsertIDEItem "Animation", hItem2, False
  'insertTVWItem_Common "ButtonControls", hItem2, False, DummyItemInfo, 0, 0, 0, True
  InsertIDEItem "ButtonControls", hItem2, False
  'insertTVWItem_Common "ComboListBoxControls", hItem2, False, DummyItemInfo, 0, 0, 0, True
  InsertIDEItem "ComboListBoxControls", hItem2, False
  'insertTVWItem_Common "DateTimeControls", hItem2, False, DummyItemInfo, 0, 0, 0, True
  InsertIDEItem "DateTimeControls", hItem2, False
  'insertTVWItem_Common "EditControls", hItem2, False, DummyItemInfo, 0, 0, 0, True
  InsertIDEItem "EditControls", hItem2, False
  'insertTVWItem_Common "ExplorerListView", hItem2, False, DummyItemInfo, 0, 0, 0, True
  InsertIDEItem "ExplorerListView", hItem2, False
  'insertTVWItem_Common "FlatButton", hItem2, False, DummyItemInfo, 0, 0, 0, True
  InsertIDEItem "FlatButton", hItem2, False
  'insertTVWItem_Common "LabelControls", hItem2, False, DummyItemInfo, 0, 0, 0, True
  InsertIDEItem "LabelControls", hItem2, False
  'insertTVWItem_Common "Pager", hItem2, False, DummyItemInfo, 0, 0, 0, True
  InsertIDEItem "Pager", hItem2, False
  'insertTVWItem_Common "ProgressBar", hItem2, False, DummyItemInfo, 0, 0, 0, True
  InsertIDEItem "ProgressBar", hItem2, False
  'insertTVWItem_Common "ShellBrowserControls", hItem2, False, DummyItemInfo, 0, 0, 0, True
  InsertIDEItem "ShellBrowserControls", hItem2, False
  'insertTVWItem_Common "StatusBar", hItem2, False, DummyItemInfo, 0, 0, 0, True
  InsertIDEItem "StatusBar", hItem2, False
  'insertTVWItem_Common "TabStrip", hItem2, False, DummyItemInfo, 0, 0, 0, True
  InsertIDEItem "TabStrip", hItem2, False
  'insertTVWItem_Common "ToolBarControls", hItem2, False, DummyItemInfo, 0, 0, 0, True
  InsertIDEItem "ToolBarControls", hItem2, False
  'insertTVWItem_Common "TrackBar", hItem2, False, DummyItemInfo, 0, 0, 0, True
  InsertIDEItem "TrackBar", hItem2, False

  ItemExpand hItem2

  FirstVisibleItem = hItem
  ItemEnsureVisible hItem
End Sub

' lädt alle(!) Items des ISO-Directory-Records, welcher im Archiv <hArchiveFile> an Position <FilePointerPos> beginnt
' für Primary Volume Descriptors und Supplementary Volume Descriptors (= JOLIET)
Private Function LoadISODirectoryRecord_2048(ByVal hArchiveItem As Long, ByVal hParentItem As Long, ByVal hArchiveFile As Long, ByVal FilePointerPos As Long, ByVal Blocksize As Long, ByRef LBAOfDirRecord As Long, Optional ByRef DirRecLength As Long = 0) As Long
  Dim attributes As Long
  Dim bytBuffer As Byte
  Dim BytesRead As Long
  Dim DispName As String
  Dim ext As String
  Dim FilePath As String
  Dim hasChilds As Boolean
  Dim ImgNormal As Long
  Dim ImgNormal_Folder As Long
  Dim ImgSelected As Long
  Dim ImgSelected_Folder As Long
  Dim intBuffer As Integer
  Dim lngBuffer As Long
  Dim pos As Long
  Dim ret As Long
  Dim strBuffer As String
  Dim TVWItemInfo As ITEMINFO_ARCHIVE

  ret = hParentItem
  If hArchiveFile = INVALID_HANDLE_VALUE Then GoTo Ende

  If (prophImageList <> 0) And Not propUseCustomIcons Then
    ImgNormal_Folder = IIf(useLargeIcons, DEFICON_FOLDER_LARGE, DEFICON_FOLDER_SMALL)
    If propUseSelectedIcons Then
      ImgSelected_Folder = IIf(useLargeIcons, DEFICON_OPENFOLDER_LARGE, DEFICON_OPENFOLDER_SMALL)
    Else
      ImgSelected_Folder = ImgNormal_Folder
    End If
  End If

  ' Länge ermitteln
  SetFilePointer hArchiveFile, FilePointerPos, 0, FILE_BEGIN
  ReadFileAsLong hArchiveFile, ByVal VarPtr(bytBuffer), LenB(bytBuffer), BytesRead, 0
  DirRecLength = bytBuffer
  If DirRecLength = 0 Then GoTo Ende

  ' Position des Extended Attribute Records ermitteln
  ' einige Daten überspringen
  SetFilePointer hArchiveFile, 1, 0, FILE_CURRENT
  ReadFileAsLong hArchiveFile, ByVal VarPtr(lngBuffer), LenB(lngBuffer), BytesRead, 0
  LBAOfDirRecord = lngBuffer

  ' Dateiflags ermitteln
  ' einige Daten überspringen
  SetFilePointer hArchiveFile, 19, 0, FILE_CURRENT
  ReadFileAsLong hArchiveFile, ByVal VarPtr(bytBuffer), LenB(bytBuffer), BytesRead, 0
  If (bytBuffer And &H2) Then attributes = attributes Or FILE_ATTRIBUTE_DIRECTORY

  ' Länge des Dateinamens ermitteln
  ' einige Daten überspringen
  SetFilePointer hArchiveFile, 6, 0, FILE_CURRENT
  ReadFileAsLong hArchiveFile, ByVal VarPtr(bytBuffer), LenB(bytBuffer), BytesRead, 0
  If bytBuffer Then
    ' Dateinamen ermitteln
    strBuffer = String$(bytBuffer, Chr$(0))
    ReadFileAsLong hArchiveFile, ByVal strBuffer, Len(strBuffer), BytesRead, 0
    If (bytBuffer = 1) And ((strBuffer = Chr$(0)) Or (strBuffer = Chr$(1))) Then strBuffer = ""
    If Left$(strBuffer, 1) = Chr$(0) Then strBuffer = StrConv(Mid$(strBuffer, 2) & Chr$(0), VbStrConv.vbFromUnicode)
    FilePath = Left$(strBuffer, lstrlenA(strBuffer))
  End If

  If FilePath <> "" Then
    With TVWItemInfo
      ' Pfad anpassen und Item hinzufügen
      FilePath = RemoveBackslash(FilePath, False)

      If attributes And FILE_ATTRIBUTE_DIRECTORY Then
        ' ein Ordner
        ' -> in 1 Schritt hinzufügen
        ret = InsertArchiveFolder(hArchiveItem, FilePath, IIf(hParentItem, hParentItem, -2))
        SetArchiveItemAttributes ret, attributes

        ' SubItems auslesen
        LoadISODirectoryRecords_2048 hArchiveItem, ret, hArchiveFile, LBAOfDirRecord * Blocksize, Blocksize
      Else
        ' eine Datei
        If hParentItem = 0 Then
          hParentItem = hArchiveItem
        Else
          SetItemType hParentItem, ItemTypeConstants.itArchiveFolder
          SetItemIcons hParentItem, ImgNormal_Folder, ImgSelected_Folder
        End If

        ' das ISO-Format ist so ausgelegt, dass wir uns sicher sein können, hier keine Ordner mehr
        ' einfügen zu müssen
        ' dafür müssen wir ggf. die Version der Datei (";n") abtrennen
        pos = InStrRev(FilePath, ";")
        If pos Then FilePath = Left$(FilePath, pos - 1)
        DispName = GetFileName(FilePath)
        .hItem_ArchiveFile = hArchiveItem
        .attributes = attributes
        .itemType = CByte(ItemTypeConstants.itArchiveFile)
        hasChilds = False

        ext = "." & GetFileNameExtension(FilePath)
        If (prophImageList <> 0) And Not propUseCustomIcons Then
          ImgNormal = GetSysIconIndexFromExt(ext, False, useLargeIcons)
          If .itemType = ItemTypeConstants.itArchiveFile Then
            ImgSelected = ImgNormal
          Else
            ImgSelected = GetSysIconIndexFromExt(ext, propUseSelectedIcons, useLargeIcons)
          End If
        End If

        If ShouldShowItem_Archive(Me, DispName, TVWItemInfo.attributes) Then
          InsertTVWItem_Archive DispName, hParentItem, hasChilds, TVWItemInfo, ImgNormal, ImgSelected
        End If
      End If
    End With
  End If

Ende:
  LoadISODirectoryRecord_2048 = ret
End Function

' lädt alle(!) Items des ISO-Directory-Records, welcher im Archiv <hArchiveFile> an Position <FilePointerPos> beginnt
' für Primary Volume Descriptors und Supplementary Volume Descriptors (= JOLIET)
Private Function LoadISODirectoryRecord_2353(ByVal hArchiveItem As Long, ByVal hParentItem As Long, ByVal hArchiveFile As Long, ByVal FilePointerPos As Long, ByVal Blocksize As Long, ByRef LBAOfDirRecord As Long, Optional ByRef DirRecLength As Long = 0) As Long
  Dim attributes As Long
  Dim bytBuffer As Byte
  Dim BytesRead As Long
  Dim DispName As String
  Dim ext As String
  Dim FilePath As String
  Dim hasChilds As Boolean
  Dim ImgNormal As Long
  Dim ImgNormal_Folder As Long
  Dim ImgSelected As Long
  Dim ImgSelected_Folder As Long
  Dim intBuffer As Integer
  Dim lngBuffer As Long
  Dim pos As Long
  Dim ret As Long
  Dim strBuffer As String
  Dim TVWItemInfo As ITEMINFO_ARCHIVE

  ret = hParentItem
  If hArchiveFile = INVALID_HANDLE_VALUE Then GoTo Ende

  If (prophImageList <> 0) And Not propUseCustomIcons Then
    ImgNormal_Folder = IIf(useLargeIcons, DEFICON_FOLDER_LARGE, DEFICON_FOLDER_SMALL)
    If propUseSelectedIcons Then
      ImgSelected_Folder = IIf(useLargeIcons, DEFICON_OPENFOLDER_LARGE, DEFICON_OPENFOLDER_SMALL)
    Else
      ImgSelected_Folder = ImgNormal_Folder
    End If
  End If

  ' Länge ermitteln
  SetFilePointer hArchiveFile, FilePointerPos, 0, FILE_BEGIN
  ReadFileAsLong hArchiveFile, ByVal VarPtr(bytBuffer), LenB(bytBuffer), BytesRead, 0
  DirRecLength = bytBuffer
  If DirRecLength = 0 Then GoTo Ende

  ' Position des Extended Attribute Records ermitteln
  ' einige Daten überspringen
  SetFilePointer hArchiveFile, 1, 0, FILE_CURRENT
  ReadFileAsLong hArchiveFile, ByVal VarPtr(lngBuffer), LenB(lngBuffer), BytesRead, 0
  LBAOfDirRecord = lngBuffer

  ' Dateiflags ermitteln
  ' einige Daten überspringen
  SetFilePointer hArchiveFile, 19, 0, FILE_CURRENT
  ReadFileAsLong hArchiveFile, ByVal VarPtr(bytBuffer), LenB(bytBuffer), BytesRead, 0
  If (bytBuffer And &H2) Then attributes = attributes Or FILE_ATTRIBUTE_DIRECTORY

  ' Länge des Dateinamens ermitteln
  ' einige Daten überspringen
  SetFilePointer hArchiveFile, 6, 0, FILE_CURRENT
  ReadFileAsLong hArchiveFile, ByVal VarPtr(bytBuffer), LenB(bytBuffer), BytesRead, 0
  If bytBuffer Then
    ' Dateinamen ermitteln
    strBuffer = String$(bytBuffer, Chr$(0))
    ReadFileAsLong hArchiveFile, ByVal strBuffer, Len(strBuffer), BytesRead, 0
    If (bytBuffer = 1) And ((strBuffer = Chr$(0)) Or (strBuffer = Chr$(1))) Then strBuffer = ""
    If Left$(strBuffer, 1) = Chr$(0) Then strBuffer = StrConv(Mid$(strBuffer, 2) & Chr$(0), VbStrConv.vbFromUnicode)
    FilePath = Left$(strBuffer, lstrlenA(strBuffer))
  End If

  If FilePath <> "" Then
    With TVWItemInfo
      ' Pfad anpassen und Item hinzufügen
      FilePath = RemoveBackslash(FilePath, False)

      If attributes And FILE_ATTRIBUTE_DIRECTORY Then
        ' ein Ordner
        ' -> in 1 Schritt hinzufügen
        ret = InsertArchiveFolder(hArchiveItem, FilePath, IIf(hParentItem, hParentItem, -2))
        SetArchiveItemAttributes ret, attributes

        ' SubItems auslesen                                           16 * 2353
        LoadISODirectoryRecords_2353 hArchiveItem, ret, hArchiveFile, 37648 + (LBAOfDirRecord - 16) * 2352, 2352
      Else
        ' eine Datei
        If hParentItem = 0 Then
          hParentItem = hArchiveItem
        Else
          SetItemType hParentItem, ItemTypeConstants.itArchiveFolder
          SetItemIcons hParentItem, ImgNormal_Folder, ImgSelected_Folder
        End If

        ' das ISO-Format ist so ausgelegt, dass wir uns sicher sein können, hier keine Ordner mehr
        ' einfügen zu müssen
        ' dafür müssen wir ggf. die Version der Datei (";n") abtrennen
        pos = InStrRev(FilePath, ";")
        If pos Then FilePath = Left$(FilePath, pos - 1)
        DispName = GetFileName(FilePath)
        .hItem_ArchiveFile = hArchiveItem
        .attributes = attributes
        .itemType = CByte(ItemTypeConstants.itArchiveFile)
        hasChilds = False

        ext = "." & GetFileNameExtension(FilePath)
        If (prophImageList <> 0) And Not propUseCustomIcons Then
          ImgNormal = GetSysIconIndexFromExt(ext, False, useLargeIcons)
          If .itemType = ItemTypeConstants.itArchiveFile Then
            ImgSelected = ImgNormal
          Else
            ImgSelected = GetSysIconIndexFromExt(ext, propUseSelectedIcons, useLargeIcons)
          End If
        End If

        If ShouldShowItem_Archive(Me, DispName, TVWItemInfo.attributes) Then
          InsertTVWItem_Archive DispName, hParentItem, hasChilds, TVWItemInfo, ImgNormal, ImgSelected
        End If
      End If
    End With
  End If

Ende:
  LoadISODirectoryRecord_2353 = ret
End Function

' lädt alle(!) SubItems des ISO-Directory-Records, welcher im Archiv <hArchiveFile> an Position <FilePointerPos> beginnt
' für Primary Volume Descriptors und Supplementary Volume Descriptors (= JOLIET)
Private Sub LoadISODirectoryRecords_2048(ByVal hArchiveItem As Long, ByVal hParentItem As Long, ByVal hArchiveFile As Long, ByVal FilePointerPos As Long, ByVal Blocksize As Long)
  Dim DirRecLength As Long
  Dim LBAOfDirRecord As Long

  Do
    FilePointerPos = FilePointerPos + DirRecLength
    LBAOfDirRecord = 0
    LoadISODirectoryRecord_2048 hArchiveItem, hParentItem, hArchiveFile, FilePointerPos, Blocksize, LBAOfDirRecord, DirRecLength
  Loop While LBAOfDirRecord > 0
End Sub

' lädt alle(!) SubItems des ISO-Directory-Records, welcher im Archiv <hArchiveFile> an Position <FilePointerPos> beginnt
' für Primary Volume Descriptors und Supplementary Volume Descriptors (= JOLIET)
Private Sub LoadISODirectoryRecords_2353(ByVal hArchiveItem As Long, ByVal hParentItem As Long, ByVal hArchiveFile As Long, ByVal FilePointerPos As Long, ByVal Blocksize As Long)
  Dim DirRecLength As Long
  Dim LBAOfDirRecord As Long

  Do
    FilePointerPos = FilePointerPos + DirRecLength
    LBAOfDirRecord = 0
    LoadISODirectoryRecord_2353 hArchiveItem, hParentItem, hArchiveFile, FilePointerPos, Blocksize, LBAOfDirRecord, DirRecLength
  Loop While LBAOfDirRecord > 0
End Sub

' lädt alle(!) Items des ISO-Volumes, welches im Archiv <hArchiveFile> an Position <FilePointerPos> beginnt
' für Primary Volume Descriptors und Supplementary Volume Descriptors (= JOLIET)
Private Sub LoadISOVolumeDescriptor_2048(ByVal hArchiveItem As Long, ByVal hArchiveFile As Long, ByVal FilePointerPos As Long)
  Dim Blocksize As Long
  Dim BytesRead As Long
  Dim hParentItem As Long
  Dim intBuffer As Integer
  Dim LBAOfDirRecord As Long

  If hArchiveFile = INVALID_HANDLE_VALUE Then Exit Sub

  ' Länge eines logischen Blockes ermitteln
  ' einige Daten überspringen
  SetFilePointer hArchiveFile, FilePointerPos + 128, 0, FILE_BEGIN
  ReadFileAsLong hArchiveFile, ByVal VarPtr(intBuffer), LenB(intBuffer), BytesRead, 0
  Blocksize = intBuffer

  ' Directory Record des root-Ordner auslesen
  ' einige Daten überspringen
  FilePointerPos = FilePointerPos + 156
  hParentItem = LoadISODirectoryRecord_2048(hArchiveItem, 0, hArchiveFile, FilePointerPos, Blocksize, LBAOfDirRecord)
  ' SubItems auslesen
  LoadISODirectoryRecords_2048 hArchiveItem, hParentItem, hArchiveFile, LBAOfDirRecord * Blocksize, Blocksize
End Sub

' lädt alle(!) Items des ISO-Volumes, welches im Archiv <hArchiveFile> an Position <FilePointerPos> beginnt
' für Primary Volume Descriptors und Supplementary Volume Descriptors (= JOLIET)
Private Sub LoadISOVolumeDescriptor_2353(ByVal hArchiveItem As Long, ByVal hArchiveFile As Long, ByVal FilePointerPos As Long)
  Dim hParentItem As Long
  Dim LBAOfDirRecord As Long

  If hArchiveFile = INVALID_HANDLE_VALUE Then Exit Sub

  ' Directory Record des root-Ordner auslesen
  ' einige Daten überspringen
  FilePointerPos = FilePointerPos + 156
  hParentItem = LoadISODirectoryRecord_2353(hArchiveItem, 0, hArchiveFile, FilePointerPos, 2352, LBAOfDirRecord)
  ' SubItems auslesen                                                   16 * 2353
  LoadISODirectoryRecords_2353 hArchiveItem, hParentItem, hArchiveFile, 37648 + (LBAOfDirRecord - 16) * 2352, 2352
End Sub

' wird aufgerufen, wenn das Event "BeforeInsertItem" ausgelöst werden soll
Friend Sub RaiseBeforeInsertItem(ByVal pIDLToDesktop As Long, ByVal hParentItem As Long, Cancel As Boolean)
  RaiseEvent BeforeInsertItem(pIDLToDesktop, hParentItem, Cancel)
End Sub

' wird aufgerufen, wenn das Event "BeforeShowEditContextMenu" ausgelöst werden soll
Private Sub RaiseBeforeShowEditContextMenu(ByRef Cancel As Boolean)
  RaiseEvent BeforeShowEditContextMenu(Cancel)
End Sub

' wird aufgerufen, wenn das Event "Click" ausgelöst werden soll
Private Sub RaiseClick()
  Dim hItem As Long
  Dim pt As POINT

  RaiseEvent Click

  GetCursorPos pt
  ScreenToClient prophWnd, pt
  hItem = ItemHitTest(pt, TVHT_ONITEM)
  RaiseItemClick hItem

  raisedMouseDown = raisedMouseDown And (Not MouseButtonConstants.vbLeftButton)
  releasedMouseButton = (raisedMouseDown = 0)
End Sub

' wird aufgerufen, wenn das Event "CompleteDrag" ausgelöst werden soll
Private Sub RaiseCompleteDrag(ByVal Canceled As Boolean, ByVal performedEffect As Long)
  RaiseEvent CompleteDrag(Canceled, performedEffect)
End Sub

' wird aufgerufen, wenn das Event "DblClick" ausgelöst werden soll
Private Sub RaiseDblClick()
  Dim hItem As Long
  Dim pt As POINT

  RaiseEvent DblClick

  GetCursorPos pt
  ScreenToClient prophWnd, pt
  hItem = ItemHitTest(pt, TVHT_ONITEM)
  RaiseItemDblClick hItem
End Sub

' wird aufgerufen, wenn das Event "DestroyLabelEditControl" ausgelöst werden soll
Private Sub RaiseDestroyLabelEditControl()
  RaiseEvent DestroyLabelEditControl
End Sub

' wird aufgerufen, wenn das Event "DragEnter" ausgelöst werden soll
Private Sub RaiseDragEnter(ByRef Data As DataObject, ByRef Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long, ByVal hItem As Long)
  RaiseEvent DragEnter(Data, Effect, Button, Shift, x, y)

  RaiseItemDragEnter hItem, Data, Effect, Button, Shift, x, y

  ' evt. Scrollen
  If IsWithinScrollingRegion(x, y) Then SetTimer prophWnd, TimerID_DragScroll, propDragScrollTime, 0
End Sub

' wird aufgerufen, wenn das Event "DragLeave" ausgelöst werden soll
Private Sub RaiseDragLeave(ByVal hItem As Long)
  RaiseItemDragLeave hItem

  RaiseEvent DragLeave
End Sub

' wird aufgerufen, wenn das Event "DragMove" ausgelöst werden soll
Private Sub RaiseDragMove(ByRef Data As DataObject, ByRef Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long, ByVal hItem As Long)
  Static oldPt As POINT

  RaiseEvent DragMove(Data, Effect, Button, Shift, x, y)

  If hItem <> prophDropTargetItem Then
    RaiseItemDragLeave prophDropTargetItem
    RaiseItemDragEnter hItem, Data, Effect, Button, Shift, x, y
  Else
    RaiseItemDragMove hItem, Data, Effect, Button, Shift, x, y
  End If

  ' evt. Scrollen
  With oldPt
    If IsWithinScrollingRegion(x, y) Then
      If Not IsWithinScrollingRegion(.x, .y) Then SetTimer prophWnd, TimerID_DragScroll, propDragScrollTime, 0
    Else
      KillTimer prophWnd, TimerID_DragScroll
    End If
    .x = x
    .y = y
  End With
End Sub

' wird aufgerufen, wenn das Event "Drop" ausgelöst werden soll
Private Sub RaiseDrop(ByRef Data As DataObject, ByRef Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long, ByVal hItem As Long)
  RaiseItemDrop hItem, Data, Effect, Button, Shift, x, y

  RaiseEvent Drop(Data, Effect, Button, Shift, x, y)
End Sub

' wird aufgerufen, wenn das Event "EditClick" ausgelöst werden soll
Private Sub RaiseEditClick()
  RaiseEvent EditClick
End Sub

' wird aufgerufen, wenn das Event "EditDblClick" ausgelöst werden soll
Private Sub RaiseEditDblClick()
  RaiseEvent EditDblClick
End Sub

' wird aufgerufen, wenn das Event "EditKeyDown" ausgelöst werden soll
Private Sub RaiseEditKeyDown(ByRef KeyCode As Integer, ByVal Shift As Integer)
  RaiseEvent EditKeyDown(KeyCode, Shift)
End Sub

' wird aufgerufen, wenn das Event "EditKeyPress" ausgelöst werden soll
Private Sub RaiseEditKeyPress(ByRef KeyAscii As Integer)
  RaiseEvent EditKeyPress(KeyAscii)
End Sub

' wird aufgerufen, wenn das Event "EditKeyUp" ausgelöst werden soll
Private Sub RaiseEditKeyUp(ByRef KeyCode As Integer, ByVal Shift As Integer)
  RaiseEvent EditKeyUp(KeyCode, Shift)
End Sub

' wird aufgerufen, wenn das Event "EditMClick" ausgelöst werden soll
Private Sub RaiseEditMClick()
  RaiseEvent EditMClick
End Sub

' wird aufgerufen, wenn das Event "EditMDblClick" ausgelöst werden soll
Private Sub RaiseEditMDblClick()
  RaiseEvent EditMDblClick
End Sub

' wird aufgerufen, wenn das Event "EditMouseDown" ausgelöst werden soll
Private Sub RaiseEditMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim pt As POINT

  If Not propHot_Edit Then
    GetCursorPos pt
    ScreenToClient hWndEdit, pt
    WinMainUC 1, 0, hWndEdit, EXTVM_MYMOUSEHOVER, -1, MakeDWord(pt.x, pt.y)
  End If

  RaiseEvent EditMouseDown(Button, Shift, x, y)
  curMouseButtons_Edit = curMouseButtons_Edit Or Button
  ' verschachtelte Click-Events verhindern
  If raisedMouseDown_Edit = 0 Then raisedMouseDown_Edit = raisedMouseDown_Edit Or Button
End Sub

' wird aufgerufen, wenn das Event "EditMouseEnter" ausgelöst werden soll
Private Sub RaiseEditMouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  MouseOverControl_Edit = True

  RaiseEvent EditMouseEnter(Button, Shift, x, y)
End Sub

' wird aufgerufen, wenn das Event "EditMouseHover" ausgelöst werden soll
Private Sub RaiseEditMouseHover(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  Hot_Edit = True

  RaiseEvent EditMouseHover(Button, Shift, x, y)
End Sub

' wird aufgerufen, wenn das Event "EditMouseLeave" ausgelöst werden soll
Private Sub RaiseEditMouseLeave(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  Dragging_Edit = False
  Hot_Edit = False
  MouseOverControl_Edit = False
  releasedMouseButton_Edit = False

  RaiseEvent EditMouseLeave(Button, Shift, x, y)

  checkForMouseLeave_Edit = False
  TimeOfPossibleMouseLeave_Edit = -1
End Sub

' wird aufgerufen, wenn das Event "EditMouseMove" ausgelöst werden soll
Private Sub RaiseEditMouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  RaiseEvent EditMouseMove(Button, Shift, x, y)
  curMouseButtons_Edit = Button
End Sub

' wird aufgerufen, wenn das Event "EditMouseUp" ausgelöst werden soll
Private Sub RaiseEditMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim pt As POINT

  If Button = 0 Then Exit Sub

  If raisedMouseDown_Edit And Button Then
    GetCursorPos pt
    If IsWithinVisibleClientArea(hWndEdit, rcWnd_Edit, pt) Then
      Select Case Button
        Case MouseButtonConstants.vbLeftButton
          RaiseEditClick
        Case MouseButtonConstants.vbMiddleButton
          RaiseEditMClick
        Case MouseButtonConstants.vbRightButton
          RaiseEditRClick
      End Select
    End If
    raisedMouseDown_Edit = raisedMouseDown_Edit And (Not Button)
  End If
  releasedMouseButton_Edit = (raisedMouseDown_Edit = 0)

  RaiseEvent EditMouseUp(Button, Shift, x, y)
  curMouseButtons_Edit = curMouseButtons_Edit And (Not Button)
End Sub

' wird aufgerufen, wenn das Event "EditRClick" ausgelöst werden soll
Private Sub RaiseEditRClick()
  RaiseEvent EditRClick
End Sub

' wird aufgerufen, wenn das Event "EditRDblClick" ausgelöst werden soll
Private Sub RaiseEditRDblClick()
  RaiseEvent EditRDblClick
End Sub

' wird aufgerufen, wenn das Event "InitLabelEditControl" ausgelöst werden soll
Private Sub RaiseInitLabelEditControl()
  checkForMouseLeave_Edit = False
  curMouseButtons_Edit = 0
  Dragging_Edit = False
  Hot_Edit = False
  MouseOverControl_Edit = False
  With ptLastMousePos_Edit
    .x = 0
    .y = 0
  End With
  raisedMouseDown_Edit = 0
  releasedMouseButton_Edit = False
  TimeOfPossibleMouseLeave_Edit = -1
  TimeOfFirstMouseMove_Edit = -1
  TimeOfLastMouseMove_Edit = -1

  RaiseEvent InitLabelEditControl
End Sub

' wird aufgerufen, wenn das Event "ItemBeforeExecuteShellContextMenuItem" ausgelöst werden soll
Private Sub RaiseItemBeforeExecuteShellContextMenuItem(ByVal hItem As Long, ByVal ItemID As Long, ByVal ItemVerb As String, ByVal ItemDescription As String, ByRef Cancel As Boolean, ByRef WindowMode As WindowModeConstants)
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub

  RaiseEvent ItemBeforeExecuteShellContextMenuItem(hItem, ItemID, ItemVerb, ItemDescription, Cancel, WindowMode)
End Sub

' wird aufgerufen, wenn das Event "ItemBeforeGetInfoTip" ausgelöst werden soll
Private Sub RaiseItemBeforeGetInfoTip(ByVal hItem As Long, ByRef InfoTipStyle As InfoTipStyleConstants, ByRef Cancel As Boolean)
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub

  RaiseEvent ItemBeforeGetInfoTip(hItem, InfoTipStyle, Cancel)
End Sub

' wird aufgerufen, wenn das Event "ItemBeforeInitShellContextMenu" ausgelöst werden soll
Private Sub RaiseItemBeforeInitShellContextMenu(ByVal hItem As Long, ByRef ContextMenuStyle As ShellContextMenuStyleConstants, ByRef Cancel As Boolean)
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub

  RaiseEvent ItemBeforeInitShellContextMenu(hItem, ContextMenuStyle, Cancel)
End Sub

' wird aufgerufen, wenn das Event "ItemBeforeRenaming" ausgelöst werden soll
Private Sub RaiseItemBeforeRenaming(ByVal hItem As Long, ByVal OldName As String, ByRef NewName As String, ByRef Cancel As Boolean)
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub

  RaiseEvent ItemBeforeRenaming(hItem, OldName, NewName, Cancel)
End Sub

' wird aufgerufen, wenn das Event "ItemBeforeShowShellContextMenu" ausgelöst werden soll
Private Sub RaiseItemBeforeShowShellContextMenu(ByVal hItem As Long, ByRef Cancel As Boolean)
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub

  RaiseEvent ItemBeforeShowShellContextMenu(hItem, Cancel)
End Sub

' wird aufgerufen, wenn das Event "ItemBeforeShowToolTip" ausgelöst werden soll
Private Sub RaiseItemBeforeShowToolTip(ByVal hItem As Long, ByVal ToolTipType As ToolTipTypeConstants, ByRef ToolTipText As String, ByRef Cancel As Boolean)
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub

  SetWindowPos hWndToolTip, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE

  RaiseEvent ItemBeforeShowToolTip(hItem, ToolTipType, ToolTipText, Cancel)
End Sub

' wird aufgerufen, wenn das Event "ItemBeforeStartRenaming" ausgelöst werden soll
Private Sub RaiseItemBeforeStartRenaming(ByVal hItem As Long, ByRef Cancel As Boolean)
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub

  RaiseEvent ItemBeforeStartRenaming(hItem, Cancel)
End Sub

' wird aufgerufen, wenn das Event "ItemBeforeStateIconChange" ausgelöst werden soll
Private Sub RaiseItemBeforeStateIconChange(ByVal hItem As Long, ByVal OldStateIcon As Long, ByRef NewStateIcon As Long, ByRef Cancel As Boolean)
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub

  RaiseEvent ItemBeforeStateIconChange(hItem, OldStateIcon, NewStateIcon, Cancel)
End Sub

' wird aufgerufen, wenn das Event "ItemClick" ausgelöst werden soll
Private Sub RaiseItemClick(ByVal hItem As Long)
  Const MK_CONTROL = &H8
  Const MK_LBUTTON = &H1
  Const MK_MBUTTON = &H10
  Const MK_RBUTTON = &H2
  Const MK_SHIFT = &H4
  Const VK_CONTROL = &H11
  Const VK_LBUTTON = MK_LBUTTON
  Const VK_MBUTTON = &H4
  Const VK_RBUTTON = MK_RBUTTON
  Const VK_SHIFT = &H10
  Const WM_LBUTTONUP = &H202
  Dim lParam As Long
  Dim pt As POINT
  Dim wParam As Long

  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub

  RaiseEvent ItemClick(hItem)

  If propAllowDragDrop Then
    ' WM_LBUTTONUP wird nur an das TreeView gesendet, wenn kein Item angeklickt wurde
    ' -> simulieren
    wParam = 0
    If IsKeyHoldDown(VK_CONTROL) Then wParam = wParam Or MK_CONTROL
    If IsKeyHoldDown(VK_LBUTTON) Then wParam = wParam Or MK_LBUTTON
    If IsKeyHoldDown(VK_MBUTTON) Then wParam = wParam Or MK_MBUTTON
    If IsKeyHoldDown(VK_RBUTTON) Then wParam = wParam Or MK_RBUTTON
    If IsKeyHoldDown(VK_SHIFT) Then wParam = wParam Or MK_SHIFT
    With pt
      GetCursorPos pt
      ScreenToClient prophWnd, pt
      lParam = MakeDWord(.x, .y)

      SendMessageAsLong prophWnd, WM_LBUTTONUP, wParam, lParam
    End With
  End If
End Sub

' wird aufgerufen, wenn das Event "ItemCollapsed" ausgelöst werden soll
Private Sub RaiseItemCollapsed(ByVal hItem As Long)
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub
  If propFavoritesStyle Then InvalidateRectAsLong prophWnd, 0, 1

  RaiseEvent ItemCollapsed(hItem)
End Sub

' wird aufgerufen, wenn das Event "ItemCollapsing" ausgelöst werden soll
Private Sub RaiseItemCollapsing(ByVal hItem As Long, ByRef Cancel As Boolean)
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub

  RaiseEvent ItemCollapsing(hItem, Cancel)
End Sub

' wird aufgerufen, wenn das Event "ItemCompleteDrag" ausgelöst werden soll
Private Sub RaiseItemCompleteDrag(ByVal hItem As Long, ByVal Canceled As Boolean, ByVal performedEffect As Long)
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then hItem = -1

  RaiseEvent ItemCompleteDrag(hItem, Canceled, performedEffect)
End Sub

' wird aufgerufen, wenn das Event "ItemDblClick" ausgelöst werden soll
Private Sub RaiseItemDblClick(ByVal hItem As Long)
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub

  RaiseEvent ItemDblClick(hItem)
End Sub

' wird aufgerufen, wenn das Event "ItemDestroyShellContextMenu" ausgelöst werden soll
Private Sub RaiseItemDestroyShellContextMenu(ByVal hItem As Long)
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then hItem = -1

  RaiseEvent ItemDestroyShellContextMenu(hItem)
End Sub

' wird aufgerufen, wenn das Event "ItemDragEnter" ausgelöst werden soll
Private Sub RaiseItemDragEnter(ByVal hItem As Long, ByRef Data As DataObject, ByRef Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub

  RaiseEvent ItemDragEnter(hItem, Data, Effect, Button, Shift, x, y)
End Sub

' wird aufgerufen, wenn das Event "ItemDragLeave" ausgelöst werden soll
Private Sub RaiseItemDragLeave(ByVal hItem As Long)
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub

  RaiseEvent ItemDragLeave(hItem)
End Sub

' wird aufgerufen, wenn das Event "ItemDragMove" ausgelöst werden soll
Private Sub RaiseItemDragMove(ByVal hItem As Long, ByRef Data As DataObject, ByRef Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub

  RaiseEvent ItemDragMove(hItem, Data, Effect, Button, Shift, x, y)
End Sub

' wird aufgerufen, wenn das Event "ItemDrop" ausgelöst werden soll
Private Sub RaiseItemDrop(ByVal hItem As Long, ByRef Data As DataObject, ByRef Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub

  RaiseEvent ItemDrop(hItem, Data, Effect, Button, Shift, x, y)
End Sub

' wird aufgerufen, wenn das Event "ItemExecutedShellContextMenuItem" ausgelöst werden soll
Private Sub RaiseItemExecutedShellContextMenuItem(ByVal hItem As Long, ByVal ItemID As Long, ByVal ItemVerb As String, ByVal ItemDescription As String)
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then hItem = -1

  RaiseEvent ItemExecutedShellContextMenuItem(hItem, ItemID, ItemVerb, ItemDescription)
End Sub

' wird aufgerufen, wenn das Event "ItemExpanded" ausgelöst werden soll
Private Sub RaiseItemExpanded(ByVal hItem As Long)
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub
  If propFavoritesStyle Then InvalidateRectAsLong prophWnd, 0, 1

  RaiseEvent ItemExpanded(hItem)
End Sub

' wird aufgerufen, wenn das Event "ItemExpanding" ausgelöst werden soll
Private Sub RaiseItemExpanding(ByVal hItem As Long, ByRef Cancel As Boolean)
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub

  RaiseEvent ItemExpanding(hItem, Cancel)
End Sub

' wird aufgerufen, wenn das Event "ItemInitShellContextMenu" ausgelöst werden soll
Private Sub RaiseItemInitShellContextMenu(ByVal hItem As Long)
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub

  RaiseEvent ItemInitShellContextMenu(hItem)
End Sub

' wird aufgerufen, wenn das Event "ItemLoadedSubItems" ausgelöst werden soll
Private Sub RaiseItemLoadedSubItems(ByVal hItem As Long)
'  If Not ItemExists(hItem, AllowNullConstants.anIfNoShowRoot) Then Exit Sub

  RaiseEvent ItemLoadedSubItems(hItem)
End Sub

' wird aufgerufen, wenn das Event "ItemLoadingSubItems" ausgelöst werden soll
Private Sub RaiseItemLoadingSubItems(ByVal hItem As Long)
'  If Not ItemExists(hItem, AllowNullConstants.anIfNoShowRoot) Then Exit Sub

  RaiseEvent ItemLoadingSubItems(hItem)
End Sub

' wird aufgerufen, wenn das Event "ItemMClick" ausgelöst werden soll
Private Sub RaiseItemMClick(ByVal hItem As Long)
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub

  RaiseEvent ItemMClick(hItem)
End Sub

' wird aufgerufen, wenn das Event "ItemMDblClick" ausgelöst werden soll
Private Sub RaiseItemMDblClick(ByVal hItem As Long)
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub

  RaiseEvent ItemMDblClick(hItem)
End Sub

' wird aufgerufen, wenn das Event "ItemMouseDown" ausgelöst werden soll
Private Sub RaiseItemMouseDown(ByVal hItem As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub

  RaiseEvent ItemMouseDown(hItem, Button, Shift, x, y)
End Sub

' wird aufgerufen, wenn das Event "ItemMouseEnter" ausgelöst werden soll
Private Sub RaiseItemMouseEnter(ByVal hItem As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub

  RaiseEvent ItemMouseEnter(hItem, Button, Shift, x, y)
End Sub

' wird aufgerufen, wenn das Event "ItemMouseLeave" ausgelöst werden soll
Private Sub RaiseItemMouseLeave(ByVal hItem As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub

  RaiseEvent ItemMouseLeave(hItem, Button, Shift, x, y)
End Sub

' wird aufgerufen, wenn das Event "ItemMouseMove" ausgelöst werden soll
Private Sub RaiseItemMouseMove(ByVal hItem As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub

  RaiseEvent ItemMouseMove(hItem, Button, Shift, x, y)
End Sub

' wird aufgerufen, wenn das Event "ItemMouseUp" ausgelöst werden soll
Private Sub RaiseItemMouseUp(ByVal hItem As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub

  RaiseEvent ItemMouseUp(hItem, Button, Shift, x, y)
End Sub

' wird aufgerufen, wenn das Event "ItemRClick" ausgelöst werden soll
Private Sub RaiseItemRClick(ByVal hItem As Long)
  Const MK_CONTROL = &H8
  Const MK_LBUTTON = &H1
  Const MK_MBUTTON = &H10
  Const MK_RBUTTON = &H2
  Const MK_SHIFT = &H4
  Const VK_CONTROL = &H11
  Const VK_LBUTTON = MK_LBUTTON
  Const VK_MBUTTON = &H4
  Const VK_RBUTTON = MK_RBUTTON
  Const VK_SHIFT = &H10
  Const WM_RBUTTONUP = &H205
  Dim lParam As Long
  Dim pt As POINT
  Dim wParam As Long

  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub

  RaiseEvent ItemRClick(hItem)

  If propAllowDragDrop Then
    ' WM_RBUTTONUP wird nur an das TreeView gesendet, wenn kein Item angeklickt wurde
    ' -> simulieren
    wParam = 0
    If IsKeyHoldDown(VK_CONTROL) Then wParam = wParam Or MK_CONTROL
    If IsKeyHoldDown(VK_LBUTTON) Then wParam = wParam Or MK_LBUTTON
    If IsKeyHoldDown(VK_MBUTTON) Then wParam = wParam Or MK_MBUTTON
    If IsKeyHoldDown(VK_RBUTTON) Then wParam = wParam Or MK_RBUTTON
    If IsKeyHoldDown(VK_SHIFT) Then wParam = wParam Or MK_SHIFT
    With pt
      GetCursorPos pt
      ScreenToClient prophWnd, pt
      lParam = MakeDWord(.x, .y)

      SendMessageAsLong prophWnd, WM_RBUTTONUP, wParam, lParam
    End With
  End If
End Sub

' wird aufgerufen, wenn das Event "ItemRDblClick" ausgelöst werden soll
Private Sub RaiseItemRDblClick(ByVal hItem As Long)
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub

  RaiseEvent ItemRDblClick(hItem)
End Sub

' wird aufgerufen, wenn das Event "ItemRenamed" ausgelöst werden soll
Private Sub RaiseItemRenamed(ByVal hItem As Long, ByVal OldName As String, ByVal NewName As String)
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub

  RaiseEvent ItemRenamed(hItem, OldName, NewName)
End Sub

' wird aufgerufen, wenn das Event "ItemShellContextMenuItemEnter" ausgelöst werden soll
Private Sub RaiseItemShellContextMenuItemEnter(ByVal hItem As Long, ByVal hMenu As Long, ByVal ItemID As Long, ByVal ItemText As String, ByVal ItemVerb As String, ByVal ItemDescription As String)
  Static hPrevItem As Long
  Static hPrevMenu As Long
  Static PrevItemDescr As String
  Static PrevItemID As Long
  Static PrevItemText As String
  Static PrevItemVerb As String

  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub

  If hPrevItem = hItem Then
    If (hPrevMenu <> hMenu) Or (PrevItemID <> ItemID) Then RaiseItemShellContextMenuItemLeave hPrevItem, hPrevMenu, PrevItemID, PrevItemText, PrevItemVerb, PrevItemDescr
  Else
    hPrevItem = 0
  End If

  If IsMenu(hMenu) = 0 Then Exit Sub

  If (hPrevMenu <> hMenu) Or (PrevItemID <> ItemID) Then
    RaiseEvent ItemShellContextMenuItemEnter(hItem, hMenu, ItemID, ItemText, ItemVerb, ItemDescription)
  End If

  hPrevItem = hItem
  hPrevMenu = hMenu
  PrevItemID = ItemID
  PrevItemText = ItemText
  PrevItemVerb = ItemVerb
  PrevItemDescr = ItemDescription
End Sub

' wird aufgerufen, wenn das Event "ItemShellContextMenuItemLeave" ausgelöst werden soll
Private Sub RaiseItemShellContextMenuItemLeave(ByVal hItem As Long, ByVal hMenu As Long, ByVal ItemID As Long, ByVal ItemText As String, ByVal ItemVerb As String, ByVal ItemDescription As String)
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub
  If IsMenu(hMenu) = 0 Then Exit Sub

  RaiseEvent ItemShellContextMenuItemLeave(hItem, hMenu, ItemID, ItemText, ItemVerb, ItemDescription)
End Sub

' wird aufgerufen, wenn das Event "ItemStateIconChanged" ausgelöst werden soll
Private Sub RaiseItemStateIconChanged(ByVal hItem As Long, ByVal OldStateIcon As Long, ByVal NewStateIcon As Long)
  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub

  RaiseEvent ItemStateIconChanged(hItem, OldStateIcon, NewStateIcon)
End Sub
'
'' wird aufgerufen, wenn das Event "ItemUpdatedSubItems" ausgelöst werden soll
'Private Sub raiseItemUpdatedSubItems(ByVal hItem As Long, ByRef DeletedItems() As Long, ByRef AddedItems() As Long, ByRef ChangedItems() As Long)
'  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub
'
'  RaiseEvent ItemUpdatedSubItems(hItem, DeletedItems, AddedItems, ChangedItems)
'End Sub
'
'' wird aufgerufen, wenn das Event "ItemUpdatingSubItems" ausgelöst werden soll
'Private Sub raiseItemUpdatingSubItems(ByVal hItem As Long, ByRef DontDeleteOld As Boolean, ByRef DontAddNew As Boolean, ByRef DontChangeExisting As Boolean)
'  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub
'
'  RaiseEvent ItemUpdatingSubItems(hItem, DontDeleteOld, DontAddNew, DontChangeExisting)
'End Sub

' wird aufgerufen, wenn das Event "KeyDown" ausgelöst werden soll
Private Sub RaiseKeyDown(ByRef KeyCode As Integer, ByVal Shift As Integer)
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub

' wird aufgerufen, wenn das Event "KeyPress" ausgelöst werden soll
Private Sub RaiseKeyPress(ByRef KeyAscii As Integer)
  RaiseEvent KeyPress(KeyAscii)
End Sub

' wird aufgerufen, wenn das Event "KeyUp" ausgelöst werden soll
Private Sub RaiseKeyUp(ByRef KeyCode As Integer, ByVal Shift As Integer)
  RaiseEvent KeyUp(KeyCode, Shift)
End Sub

' wird aufgerufen, wenn das Event "MClick" ausgelöst werden soll
Private Sub RaiseMClick()
  Dim hItem As Long
  Dim pt As POINT

  RaiseEvent MClick

  GetCursorPos pt
  ScreenToClient prophWnd, pt
  hItem = ItemHitTest(pt, TVHT_ONITEM)
  RaiseItemMClick hItem

'  raisedMouseDown = raisedMouseDown And (Not MouseButtonConstants.vbMiddleButton)
'  releasedMouseButton = (raisedMouseDown = 0)
End Sub

' wird aufgerufen, wenn das Event "MDblClick" ausgelöst werden soll
Private Sub RaiseMDblClick()
  Dim hItem As Long
  Dim pt As POINT

  RaiseEvent MDblClick

  GetCursorPos pt
  ScreenToClient prophWnd, pt
  hItem = ItemHitTest(pt, TVHT_ONITEM)
  RaiseItemMDblClick hItem
End Sub

' wird aufgerufen, wenn das Event "MouseDown" ausgelöst werden soll
Private Sub RaiseMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hItem As Long)
  Dim pt As POINT

  If Not propHot Then
    GetCursorPos pt
    ScreenToClient prophWnd, pt
    WinMainUC 1, 0, prophWnd, EXTVM_MYMOUSEHOVER, -1, MakeDWord(pt.x, pt.y)
  End If

  RaiseEvent MouseDown(Button, Shift, x, y)
  RaiseItemMouseDown hItem, Button, Shift, x, y
  curMouseButtons = curMouseButtons Or Button
  ' verschachtelte Click-Events verhindern
  If raisedMouseDown = 0 Then raisedMouseDown = raisedMouseDown Or Button
End Sub

' wird aufgerufen, wenn das Event "MouseEnter" ausgelöst werden soll
Private Sub RaiseMouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  MouseOverControl = True

  RaiseEvent MouseEnter(Button, Shift, x, y)
  RaiseItemMouseEnter hItemUnderMouse, Button, Shift, x, y
End Sub

' wird aufgerufen, wenn das Event "MouseHover" ausgelöst werden soll
Private Sub RaiseMouseHover(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  Hot = True

  RaiseEvent MouseHover(Button, Shift, x, y)
End Sub

' wird aufgerufen, wenn das Event "MouseLeave" ausgelöst werden soll
Private Sub RaiseMouseLeave(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  Dragging = False
  Hot = False
  MouseOverControl = False
  releasedMouseButton = False

  RaiseItemMouseLeave hItemUnderMouse, Button, Shift, x, y
  hItemUnderMouse = -1
  RaiseEvent MouseLeave(Button, Shift, x, y)

  checkForMouseLeave = False
  TimeOfPossibleMouseLeave = -1
End Sub

' wird aufgerufen, wenn das Event "MouseMove" ausgelöst werden soll
Private Sub RaiseMouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hItem As Long)
  RaiseEvent MouseMove(Button, Shift, x, y)
  curMouseButtons = Button

  If hItem <> hItemUnderMouse Then
    RaiseItemMouseLeave hItemUnderMouse, Button, Shift, x, y
    hItemUnderMouse = hItem
    RaiseItemMouseEnter hItemUnderMouse, Button, Shift, x, y
  Else
    RaiseItemMouseMove hItem, Button, Shift, x, y
  End If
End Sub

' wird aufgerufen, wenn das Event "MouseUp" ausgelöst werden soll
Private Sub RaiseMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hItem As Long)
  Dim pt As POINT

  If Button = 0 Then Exit Sub

  If raisedMouseDown And Button Then
    GetCursorPos pt
    If IsWithinVisibleClientArea(prophWnd, rcWnd, pt) Then
      Select Case Button
        Case MouseButtonConstants.vbMiddleButton
          RaiseMClick
          raisedMouseDown = raisedMouseDown And (Not Button)
          releasedMouseButton = (raisedMouseDown = 0)
      End Select
    End If
  End If

  RaiseItemMouseUp hItem, Button, Shift, x, y

  RaiseEvent MouseUp(Button, Shift, x, y)
  curMouseButtons = curMouseButtons And (Not Button)
End Sub

' wird aufgerufen, wenn das Event "RClick" ausgelöst werden soll
Private Sub RaiseRClick()
  Dim Button As MouseButtonConstants
  Dim hItem As Long
  Dim pt As POINT
  Dim Shift As ShiftConstants

  RaiseEvent RClick

  With pt
    GetCursorPos pt
    ScreenToClient prophWnd, pt
    hItem = ItemHitTest(pt, TVHT_ONITEM)
    RaiseItemRClick hItem

    GetShiftAndMouseButtonBits 0, Shift, Button, True
    RaiseMouseUp curMouseButtons And Not Button, Shift, .x, .y, hItem
  End With

  raisedMouseDown = raisedMouseDown And (Not MouseButtonConstants.vbRightButton)
  releasedMouseButton = (raisedMouseDown = 0)
End Sub

' wird aufgerufen, wenn das Event "RDblClick" ausgelöst werden soll
Private Sub RaiseRDblClick()
  Dim hItem As Long
  Dim pt As POINT

  RaiseEvent RDblClick

  GetCursorPos pt
  ScreenToClient prophWnd, pt
  hItem = ItemHitTest(pt, TVHT_ONITEM)
  RaiseItemRDblClick hItem
End Sub

' wird aufgerufen, wenn das Event "SelChanged" ausgelöst werden soll
Private Sub RaiseSelChanged(ByVal hOldItem As Long, ByVal hNewItem As Long, ByVal CausedBy As CausedByConstants, Optional ByVal timered As Boolean = True)
  If timered And propSelChangedDelayTime > 0 Then
    If CausedBy = CausedByConstants.cbKeyboard Then
      SelChangedDelay_hNewItem = hNewItem
      If TimeOfRealSelChanged = -1 Then SelChangedDelay_hOldItem = hOldItem
      TimeOfRealSelChanged = GetTickCount
      Exit Sub
    End If
  End If

  If Not ItemExists(hNewItem, AllowNullConstants.anNever) Then Exit Sub

  If (CausedBy <> CausedByConstants.cbMouse) And (CausedBy <> CausedByConstants.cbKeyboard) Then CausedBy = IIf(SelChangeInternal, CausedByConstants.cbInternal, CausedByConstants.cbUnknown)
  SelChangeInternal = False
  If propFavoritesStyle Then InvalidateRectAsLong prophWnd, 0, 1

  If Not dontSavePath Then propPath = ItemHandleToPath(hNewItem)

  If Not suppressSelChangeEvents Then
    If Not ItemExists(hOldItem, AllowNullConstants.anAlways) Then hOldItem = -1
    RaiseEvent SelChanged(hOldItem, hNewItem, CausedBy)
  End If
End Sub

' wird aufgerufen, wenn das Event "SelChanging" ausgelöst werden soll
Private Sub RaiseSelChanging(ByVal hOldItem As Long, ByVal hNewItem As Long, ByVal CausedBy As CausedByConstants, ByRef Cancel As Boolean)
  If Not ItemExists(hNewItem, AllowNullConstants.anNever) Then Exit Sub

  If (CausedBy <> CausedByConstants.cbMouse) And (CausedBy <> CausedByConstants.cbKeyboard) Then CausedBy = IIf(SelChangeInternal, CausedByConstants.cbInternal, CausedByConstants.cbUnknown)
  SelChangeInternal = False

  If Not suppressSelChangeEvents Then
    If Not ItemExists(hOldItem, AllowNullConstants.anAlways) Then hOldItem = -1
    RaiseEvent SelChanging(hOldItem, hNewItem, CausedBy, Cancel)
  End If
End Sub

Private Sub RebuildShellIconCache()
  Const WM_SETREDRAW = &HB
  Dim cxIcons As Long

  If bPendingShellIconCacheRebuild Then Exit Sub

  bPendingShellIconCacheRebuild = True
  SendMessageAsLong prophWnd, WM_SETREDRAW, 0, 0
  If ver_Win_NTBased Then
    cxIcons = GetShellIconSize
    SetShellIconSize cxIcons + 1
    SetShellIconSize cxIcons
  End If
  SetTimer prophWnd, TimerID_ResetIconCacheRebuildFlag, 5000, 0
  SendMessageAsLong prophWnd, WM_SETREDRAW, 1, 0
End Sub

' lädt für <hItem> und die Sub-Items die pIDL neu
' die pIDLs müssen von ILIsEqual noch als gleich erkannt werden
Private Sub ReloadTVWItemPIDL(hItem As Long)
  Dim EnumFlags As SHCONTFConstants
  Dim hItem2 As Long
  Dim IEnum As IVBEnumIDList
  Dim IParent As IVBShellFolder
  Dim isRootItem As Boolean
  Dim pIDLParent_ToDesktop As Long
  Dim pIDLSubItem_ToDesktop As Long
  Dim pIDLSubItem_ToParent As Long
  Dim pIDLToDesktop As Long

  If hItem Then
    If propShowRoot Then isRootItem = (GetRootOrFirstItem = hItem)
  Else
    isRootItem = True
  End If

  pIDLToDesktop = GetFullyQualifiedPIDL(hItem)
  If pIDLToDesktop = 0 Then Exit Sub
  pIDLParent_ToDesktop = ILClone(pIDLToDesktop)
  MakeParentItemIDList pIDLParent_ToDesktop
  #If Debuging Then
    debugger.AddLogEntry "ReloadTVWItemPIDL: Calling GetISHFolderInterfaceFQ()", LogEntryTypeConstants.letInfo
    debugger.AddLogEntry "   0x" & Hex(pIDLParent_ToDesktop), LogEntryTypeConstants.letOther
    Set IParent = GetISHFolderInterfaceFQ(debugger, pIDLParent_ToDesktop)
  #Else
    Set IParent = GetISHFolderInterfaceFQ(pIDLParent_ToDesktop)
  #End If

  If Not (IParent Is Nothing) Then
    EnumFlags = SHCONTFConstants.SHCONTF_FOLDERS Or SHCONTFConstants.SHCONTF_NONFOLDERS Or SHCONTFConstants.SHCONTF_INCLUDEHIDDEN
    IParent.EnumObjects 0, EnumFlags, IEnum
    If Not (IEnum Is Nothing) Then
      Do While IEnum.Next(1, pIDLSubItem_ToParent, 0) = S_OK
        If pIDLSubItem_ToParent Then
          pIDLSubItem_ToDesktop = ILAppendID(ILClone(pIDLParent_ToDesktop), pIDLSubItem_ToParent, 1)
          #If Debuging Then
            FreeItemIDList debugger, "ReloadTVWItemPIDL #1", pIDLSubItem_ToParent
            debugger.AddLogEntry "ReloadTVWItemPIDL: Calling InternalpIDLToItemHandle()", letInfo
            debugger.AddLogEntry "   0x" & Hex(pIDLSubItem_ToDesktop), LogEntryTypeConstants.letOther
          #Else
            FreeItemIDList pIDLSubItem_ToParent
          #End If
          hItem2 = InternalpIDLToItemHandle(pIDLSubItem_ToDesktop)
        Else
          hItem2 = -1
        End If

        If hItem2 = hItem Then
          #If Debuging Then
            debugger.AddLogEntry "ReloadTVWItemPIDL: Calling ReloadTVWItemProps()", letInfo
            debugger.AddLogEntry "   Item 0x" & Hex(hItem), LogEntryTypeConstants.letOther
            debugger.AddLogEntry "   pIDL 0x" & Hex(pIDLSubItem_ToDesktop), LogEntryTypeConstants.letOther
          #End If
          ReloadTVWItemProps hItem, pIDLSubItem_ToDesktop, , False
          #If Debuging Then
            FreeItemIDList debugger, "ReloadTVWItemPIDL #2", pIDLSubItem_ToDesktop
          #Else
            FreeItemIDList pIDLSubItem_ToDesktop
          #End If
          Exit Do
        End If
        #If Debuging Then
          FreeItemIDList debugger, "ReloadTVWItemPIDL #3", pIDLSubItem_ToDesktop
        #Else
          FreeItemIDList pIDLSubItem_ToDesktop
        #End If
      Loop
    End If
    Set IEnum = Nothing
  End If
  Set IParent = Nothing
  #If Debuging Then
    FreeItemIDList debugger, "ReloadTVWItemPIDL #4", pIDLParent_ToDesktop
    debugger.AddLogEntry "Leaving ReloadTVWItemPIDL", LogEntryTypeConstants.letInfo
  #Else
    FreeItemIDList pIDLParent_ToDesktop
  #End If
End Sub

' lädt für <hItem> alle Eigenschaften neu
' -> geht davon aus, daß die pIDLs aktuell sind
Private Sub ReloadTVWItemProps(hItem As Long, Optional ByVal newpIDL As Long = 0, Optional ByVal skipShouldShowIt As Boolean = False, Optional ByVal updateSubItemsOnNewPIDL As Boolean = True)
  Dim hasChilds As Boolean
  Dim hSubItem As Long
  Dim IParent As IVBShellFolder
  Dim IShIconOverlay As IVBShellIconOverlay
  Dim isRootItem As Boolean
  Dim itemAttr As SFGAOConstants
  Dim itemData As TVITEMEX
  Dim itemType As Byte
  Dim oldPIDLToFree As Long
  Dim overlayIndex As Long
  Dim pIDL As Long
  Dim pIDLToParent As Long
  Dim shouldShowIt As Boolean
  Dim tmp As Long
  Dim TVWItemInfo As ITEMINFO_COMMON

  If hItem Then
    If propShowRoot Then isRootItem = (GetRootOrFirstItem = hItem)
  Else
    isRootItem = True
  End If

  ' die alten Daten ermitteln
  If hItem Then
    With itemData
      .hItem = hItem
      .Mask = TVIF_HANDLE Or TVIF_PARAM
      SendMessage prophWnd, TVM_GETITEM, 0, itemData
      CopyMemory VarPtr(itemType), .lParam, LenB(itemType)
      If (itemType = ItemTypeConstants.itArchiveFile) Or (itemType = ItemTypeConstants.itArchiveFolder) Then Exit Sub

      CopyMemory VarPtr(TVWItemInfo), .lParam, LenB(TVWItemInfo)
    End With
  Else
    TVWItemInfo = RootItemInfo
  End If

  If newpIDL Then
    oldPIDLToFree = TVWItemInfo.pIDLToDesktop
    TVWItemInfo.pIDLToDesktop = ILClone(newpIDL)
  End If
  #If Debuging Then
    SplitFullyQualifiedPIDL debugger, TVWItemInfo.pIDLToDesktop, IParent, pIDLToParent
  #Else
    SplitFullyQualifiedPIDL TVWItemInfo.pIDLToDesktop, IParent, pIDLToParent
  #End If

  ' Daten verändern
  If skipShouldShowIt Then
    shouldShowIt = True
  Else
    #If Debuging Then
      debugger.AddLogEntry "ReloadTVWItemProps: Calling ShouldShowItem()...", LogEntryTypeConstants.letInfo
      debugger.AddLogEntry "   Item=" & pIDLToDisplayName(debugger, IParent, pIDLToParent), LogEntryTypeConstants.letOther
      debugger.AddLogEntry "   Parent Item=0x" & Hex(Me.ItemGetParentItem(hItem)), LogEntryTypeConstants.letOther
      shouldShowIt = ShouldShowItem(debugger, Me, IParent, pIDLToParent, TVWItemInfo.pIDLToDesktop, Me.ItemGetParentItem(hItem))
      debugger.AddLogEntry "   Result=" & shouldShowIt, LogEntryTypeConstants.letOther
    #Else
      shouldShowIt = ShouldShowItem(Me, IParent, pIDLToParent, TVWItemInfo.pIDLToDesktop, Me.ItemGetParentItem(hItem))
    #End If
  End If

  If shouldShowIt Then
    ' Item updaten
    ' Typ des Items ermitteln
    #If Debuging Then
      itemAttr = GetAttributes(debugger, IParent, pIDLToParent, SFGAOConstants.SFGAO_FILESYSTEM Or SFGAOConstants.SFGAO_FOLDER Or SFGAOConstants.SFGAO_SHARE Or SFGAOConstants.SFGAO_LINK Or SFGAOConstants.SFGAO_NEWCONTENT Or SFGAOConstants.SFGAO_GHOSTED)
    #Else
      itemAttr = GetAttributes(IParent, pIDLToParent, SFGAOConstants.SFGAO_FILESYSTEM Or SFGAOConstants.SFGAO_FOLDER Or SFGAOConstants.SFGAO_SHARE Or SFGAOConstants.SFGAO_LINK Or SFGAOConstants.SFGAO_NEWCONTENT Or SFGAOConstants.SFGAO_GHOSTED)
    #End If
    If itemAttr And SFGAOConstants.SFGAO_FILESYSTEM Then
      If itemAttr And SFGAOConstants.SFGAO_FOLDER Then
        ' ist es wirklich ein Ordner?
        #If Debuging Then
          If FileExists_pIDL(debugger, IParent, pIDLToParent) Then
        #Else
          If FileExists_pIDL(IParent, pIDLToParent) Then
        #End If
          TVWItemInfo.itemType = CByte(ItemTypeConstants.itFSFile)
        Else
          TVWItemInfo.itemType = CByte(ItemTypeConstants.itFSFolder)
        End If
      Else
        TVWItemInfo.itemType = CByte(ItemTypeConstants.itFSFile)
      End If
    Else
      If itemAttr And SFGAOConstants.SFGAO_FOLDER Then
        TVWItemInfo.itemType = CByte(ItemTypeConstants.itNonFSFolder)
      Else
        TVWItemInfo.itemType = CByte(ItemTypeConstants.itNonFSFile)
      End If
    End If
    If hItem = GetInternetExplorerItemHandle(False) Then
      hasChilds = ItemHasChildren(hItem)
    Else
      #If Debuging Then
        If IsSlowItem(debugger, IParent, pIDLToParent, TVWItemInfo.pIDLToDesktop) And Not propCheckSlowItemsForSubItems Then
      #Else
        If IsSlowItem(IParent, pIDLToParent, TVWItemInfo.pIDLToDesktop) And Not propCheckSlowItemsForSubItems Then
      #End If
        ' Dateien haben in der Regel keine SubItems
        ' vielleicht ist es ein Archiv
        #If Debuging Then
          If IsArchiveToExpand(debugger, IParent, pIDLToParent, propExpandArchives) Then
        #Else
          If IsArchiveToExpand(IParent, pIDLToParent, propExpandArchives) Then
        #End If
          hasChilds = True
        Else
          hasChilds = ((TVWItemInfo.itemType = ItemTypeConstants.itFSFolder) Or (TVWItemInfo.itemType = ItemTypeConstants.itNonFSFolder))
        End If
      Else
        #If Debuging Then
          hasChilds = HasSubItems(debugger, TVWItemInfo.pIDLToDesktop, Me)
        #Else
          hasChilds = HasSubItems(TVWItemInfo.pIDLToDesktop, Me)
        #End If
      End If
    End If

    If (prophImageList <> 0) And Not propUseCustomIcons Then
      ' Overlays laden
      If propShownOverlays <> ShownOverlaysConstants.soNone Then
        On Error Resume Next
        IParent.QueryInterface IID_IShellIconOverlay, IShIconOverlay
        If IShIconOverlay Is Nothing Then
          If itemAttr And SFGAOConstants.SFGAO_SHARE Then overlayIndex = 1
          If itemAttr And SFGAOConstants.SFGAO_LINK Then overlayIndex = 2
        Else
          If IShIconOverlay.GetOverlayIndex(pIDLToParent, overlayIndex) = S_OK Then
            ' jetzt haben wir den Overlay-Index für den Item, aber soll dieser Overlay überhaupt
            ' angezeigt werden?
            If ShouldShowOverlay(propShownOverlays, overlayIndex) Then
              ' ja!
            Else
              ' kein Overlay
              overlayIndex = 0
            End If
          Else
            ' kein Overlay
            overlayIndex = 0
          End If
          Set IShIconOverlay = Nothing
        End If
      End If
    End If

    If Not hasChilds Then
      If ItemHasSubItems(hItem) Then RemoveSubItems hItem
    End If

    ' geänderte Daten übertragen
    If hItem Then
      With itemData
        .hItem = hItem
        .cChildren = Abs(hasChilds)
        #If Debuging Then
          .pszText = pIDLToDisplayName(debugger, IParent, pIDLToParent, SHGDNConstants.SHGDN_INFOLDER)
        #Else
          .pszText = pIDLToDisplayName(IParent, pIDLToParent, SHGDNConstants.SHGDN_INFOLDER)
        #End If
        .cchTextMax = Len(.pszText)
        CopyMemory .lParam, VarPtr(TVWItemInfo), LenB(TVWItemInfo)
        .Mask = TVIF_CHILDREN Or TVIF_HANDLE Or TVIF_TEXT Or TVIF_STATE
        If (prophImageList <> 0) And Not propUseCustomIcons Then
          .iImage = I_IMAGECALLBACK
          .iSelectedImage = I_IMAGECALLBACK
          .Mask = .Mask Or TVIF_IMAGE Or TVIF_SELECTEDIMAGE
        End If

        .state = 0
        If itemAttr And SFGAOConstants.SFGAO_NEWCONTENT Then .state = .state Or TVIS_BOLD
        If itemAttr And SFGAOConstants.SFGAO_GHOSTED Then
          If propShowHiddenDimmed Then .state = .state Or TVIS_CUT
        End If

        ' TODO: Mglw. funktioniert das Entfernen eines Overlays unter Vista so nicht
        If overlayIndex > 0 Then
          .state = .state Or INDEXTOOVERLAYMASK(overlayIndex)
        End If
        .StateMask = TVIS_CUT Or TVIS_BOLD Or TVIS_OVERLAYMASK

        SendMessage prophWnd, TVM_SETITEM, 0, itemData
      End With
    End If

    If isRootItem Then
      RootItemInfo = TVWItemInfo
    End If

    If (newpIDL <> 0) And updateSubItemsOnNewPIDL Then
      hSubItem = Me.ItemGetFirstSubItem(hItem)
      While hSubItem <> -1
        tmp = Me.ItemGetNextItem(hSubItem)

        pIDL = GetFullyQualifiedPIDL(hSubItem, True)
        If pIDL Then
          pIDL = ILAppendID(ILClone(newpIDL), ILFindLastID(pIDL), 1)
          #If Debuging Then
            If Not ValidateItemFQ(debugger, pIDL) Then
          #Else
            If Not ValidateItemFQ(pIDL) Then
          #End If
            RemoveTVWItem hSubItem, True
          Else
            ReloadTVWItemProps hSubItem, pIDL, , updateSubItemsOnNewPIDL
          End If
          #If Debuging Then
            FreeItemIDList debugger, "ReloadTVWItemProps #1", pIDL
          #Else
            FreeItemIDList pIDL
          #End If
        End If

        hSubItem = tmp
      Wend
    End If
  Else
    ' Item entfernen
    If hItem Then RemoveTVWItem hItem
  End If
  Set IParent = Nothing

  If oldPIDLToFree Then
    #If Debuging Then
      FreeItemIDList debugger, "ReloadTVWItemProps #2", oldPIDLToFree
    #Else
      FreeItemIDList oldPIDLToFree
    #End If
  End If
End Sub

' entfernt bei allen SubItems von <hParentItem> die Icons
Private Sub RemoveAllIcons(hParentItem As Long, Optional ByVal AlsoRemoveOverlay As Boolean = False, Optional ByVal skipExistanceCheck As Boolean = False)
  Dim hSubItem As Long

  If Not skipExistanceCheck Then
    If Not ItemExists(hParentItem, AllowNullConstants.anIfNoShowRoot) Then Exit Sub
  End If

  hSubItem = ItemGetFirstSubItem(hParentItem)
  While Not IsWithinRange(hSubItem, -1, 0)
    SetItemIcons hSubItem, 0, 0
    If AlsoRemoveOverlay Then SetItemOverlay hSubItem, 0

    If ItemIsExpanded(hSubItem) Then
      If ItemHasSubItems(hSubItem) Then RemoveAllIcons hSubItem, AlsoRemoveOverlay, True
    Else
      ItemCollapse hSubItem, True
    End If
    hSubItem = ItemGetNextItem(hSubItem)
  Wend
End Sub

' entfernt alle SubItems von <hItem>, außer die, deren Handle in <protectedItems> enthalten ist
Private Sub RemoveSubItems(ByVal hItem As Long, Optional ByVal protectedItems As String = "")
  Dim hSubItem As Long
  Dim tmp As Long

  If Not ItemExists(hItem, AllowNullConstants.anAlways) Then Exit Sub

  If protectedItems = "" Then
    ' alle SubItems entfernen
    If hItem Then
      ItemCollapse hItem, True
    Else
      ClearTVWItems
    End If
  Else
    ' jeden Item einzeln prüfen
    If hItem Then
      hSubItem = ItemGetFirstSubItem(hItem)
    Else
      hSubItem = GetRootOrFirstItem
    End If
    While Not IsWithinRange(hSubItem, -1, 0)
      ' nächsten Item ermitteln
      tmp = ItemGetNextItem(hSubItem)

      If InStr(protectedItems, CStr(hSubItem)) = 0 Then
        ' Item ist nicht geschützt - entfernen
        RemoveTVWItem hSubItem
      End If

      hSubItem = tmp
    Wend
  End If
End Sub

' entfernt <hItem>
Private Sub RemoveTVWItem(ByVal hItem As Long, Optional ByVal skipExistanceCheck As Boolean = False)
  Dim hParentItem As Long
  Dim hSelectedItem As Long

  If Not skipExistanceCheck Then
    If Not ItemExists(hItem, AllowNullConstants.anIfNoShowRoot) Then Exit Sub
  End If

  #If Debuging Then
    debugger.AddLogEntry "Removing item 0x" & Hex(hItem) & " (" & ItemHandleToTreePath(hItem) & ")", LogEntryTypeConstants.letInfo
  #End If

  hParentItem = Me.ItemGetParentItem(hItem)
  hSelectedItem = Me.SelectedItem
  If hSelectedItem = hItem Then
    ' den Parent-Item markieren
    SelChangeInternal = True
    Me.SelectedItem = hParentItem
    SelChangeInternal = False
  End If

  If ItemExists(hParentItem, AllowNullConstants.anNever) Then
    If Not ItemHasMoreThan1SubItems(hParentItem) Then
      ' wegen "ItemWasExpandedOnce"
      ItemCollapse hParentItem, True
    Else
      SendMessageAsLong prophWnd, TVM_DELETEITEM, 0, hItem
    End If
    ' hasChilds-Eigenschaft neu setzen
    SetItemChildrenProperty hParentItem, ItemHasSubItems(hParentItem)
  Else
    SendMessageAsLong prophWnd, TVM_DELETEITEM, 0, hItem
  End If
End Sub

' aktiviert <PathToSelect> und erweitert die Items entsprechend
' <PathToSelect> kann sein: - normaler Pfad (evt. mit DisplayNames)
'                           - FTP-URL
'                           - CLSID-Pfad (evt. mit DisplayNames)
'                           - DisplayName-Pfad (beginnend mit "Desktop")
' alle Formate können mit Archiv-Items kombiniert sein
Private Sub SelectPath(ByVal PathToSelect As String, Optional ByVal expand As Boolean = True)
  Dim ArchivePart As String
  Dim cSegments As Long
  Dim hItem As Long
  Dim i As Integer
  Dim isFTPPath As Boolean
  Dim pIDLToSelect As Long
  Dim pos As Integer
  Dim Segment As String
  Dim Sep As String
  Dim txt As String

  If inIDE Then Exit Sub
  If prophWnd = 0 Then Exit Sub
  If PathToSelect = "" Then Exit Sub

  #If Debuging Then
    debugger.AddLogEntry "Selecting " & propPath, LogEntryTypeConstants.letInfo
  #End If

  isFTPPath = IsFTPURL(PathToSelect)
  Sep = IIf(isFTPPath, "/", "\")

  ' ggf. Archiv-Pfad abtrennen
  If InStr(PathToSelect, ArchiveSignature) Then
    pos = InStr(PathToSelect, ArchiveSignature)
    ArchivePart = Mid$(PathToSelect, pos + Len(ArchiveSignature))
    PathToSelect = Left$(PathToSelect, pos - 1)
  End If

  ' ggf. CLSID für den Desktop abtrennen
  If IncludesCLSID(PathToSelect) Then
    Segment = GetFirstFolders(PathToSelect, 1, Sep)
    #If Debuging Then
      If LCase$(CLSIDToDisplayName(Segment)) = LCase$(pIDLToDisplayName_Light(debugger, pIDL_Desktop)) Then
    #Else
      If LCase$(CLSIDToDisplayName(Segment)) = LCase$(pIDLToDisplayName_Light(pIDL_Desktop)) Then
    #End If
      PathToSelect = RemoveFirstFolders(PathToSelect, 1, Sep)
    End If
  End If

  PathToSelect = RemoveBackslash(PathToSelect, False, Sep)

  #If Debuging Then
    debugger.AddLogEntry "SelectPath: PathToSelect=" & PathToSelect, LogEntryTypeConstants.letInfo
  #End If

  ' anhand von <PathToSelect> <pIDLToSelect> aufbauen
  pIDLToSelect = UserPathTopIDL(PathToSelect)

  #If Debuging Then
    debugger.AddLogEntry "SelectPath: pIDLToSelect=0x" & Hex(pIDLToSelect) & " (" & pIDLToPath_Light(debugger, pIDLToSelect) & ")", LogEntryTypeConstants.letInfo
  #End If

  If isFTPPath Then
    If propShowRoot Then
      If ILIsEqual(pIDL_Desktop, RootItemInfo.pIDLToDesktop) Then
        ' der Root-Item wurde evt. noch nicht erweitert
        hItem = GetRootOrFirstItem
        ItemExpand hItem
      End If
    End If
    hItem = Me.AddFTPPath(PathToSelect)
  Else
    #If Debuging Then
      debugger.AddLogEntry "SelectPath: Calling InternalpIDLToItemHandle()", letInfo
      debugger.AddLogEntry "   0x" & Hex(pIDLToSelect), LogEntryTypeConstants.letOther
    #End If
    hItem = InternalpIDLToItemHandle(pIDLToSelect, IIf(propShowRoot, GetRootOrFirstItem, 0), True, , , True, False)
  End If
  #If Debuging Then
    FreeItemIDList debugger, "SelectPath", pIDLToSelect
  #Else
    FreeItemIDList pIDLToSelect
  #End If
  SelChangeInternal = True
  Me.SelectedItem = hItem
  Me.ItemEnsureVisible hItem
  If expand Or (ArchivePart <> "") Then Me.ItemExpand hItem

  ' jetzt den Archiv-Pfad verarbeiten
  If ArchivePart <> "" Then
    hItem = Me.SelectedItem
    If hItem <> -1 Then
      cSegments = CountSegments(ArchivePart)
      For i = 1 To cSegments
        Segment = GetFirstFolders(ArchivePart, i)
        Segment = GetLastFolders(Segment, 1)
        ' jetzt alle SubItems von <hItem> durchgehen
        hItem = Me.ItemGetFirstSubItem(hItem)
        Do While Not IsWithinRange(hItem, -1, 0)
          txt = GetItemText(hItem)
          If LCase$(txt) = LCase$(Segment) Then
            ' gefunden! -> anwählen und erweitern
            SelChangeInternal = True
            Me.SelectedItem = hItem
            Me.ItemExpand hItem
            Exit Do
          End If
          hItem = Me.ItemGetNextItem(hItem)
        Loop
      Next
    End If
  End If

Ende:
  If SelectedItem = -1 Then
    ' wenigstens die Root anwählen
    hItem = GetRootOrFirstItem
    If hItem Then
      SelChangeInternal = True
      Me.SelectedItem = hItem
      Me.ItemExpand hItem
    End If
  End If
End Sub

Private Sub SetArchiveItemAttributes(ByVal hItem As Long, ByVal attributes As Long)
  Dim itemData As TVITEMEX
  Dim TVWItemInfo As ITEMINFO_ARCHIVE

  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub
  If Not ItemIsPartOfArchive(hItem) Then Exit Sub

  With itemData
    .hItem = hItem
    .Mask = TVIF_HANDLE Or TVIF_PARAM Or TVIF_STATE
    SendMessage prophWnd, TVM_GETITEM, 0, itemData
    CopyMemory VarPtr(TVWItemInfo), .lParam, LenB(TVWItemInfo)

    TVWItemInfo.attributes = attributes
    .state = 0
    If TVWItemInfo.attributes And FILE_ATTRIBUTE_HIDDEN Then
      If propShowHiddenDimmed Then .state = .state Or TVIS_CUT
    End If
    .StateMask = TVIS_CUT

    CopyMemory .lParam, VarPtr(TVWItemInfo), LenB(TVWItemInfo)

    SendMessage prophWnd, TVM_SETITEM, 0, itemData
  End With
End Sub

' setzt die Farben des TreeViews
Private Sub SetColors()
  Dim RGBClr As Long

  If prophWnd Then
    #If BkImageSupport Then
      RGBClr = IIf(propBkImage Is Nothing, TranslateColor(propBackColor), -1)
    #Else
      RGBClr = TranslateColor(propBackColor)
    #End If
    SendMessageAsLong prophWnd, TVM_SETBKCOLOR, 0, RGBClr

    RGBClr = TranslateColor(propLineColor)
    SendMessageAsLong prophWnd, TVM_SETLINECOLOR, 0, RGBClr
    RGBClr = TranslateColor(propForeColor)
    SendMessageAsLong prophWnd, TVM_SETTEXTCOLOR, 0, RGBClr
    mustCustomDraw = IsCustomDrawNecessary
  End If
End Sub

Private Sub SetItemChildrenProperty(ByVal hItem As Long, ByVal hasChilds As Boolean)
  Dim itemData As TVITEMEX

  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub

  With itemData
    .hItem = hItem
    .cChildren = Abs(hasChilds)
    .Mask = TVIF_HANDLE Or TVIF_CHILDREN
    SendMessage prophWnd, TVM_SETITEM, 0, itemData
  End With
End Sub

Private Sub SetItemIcon(hItem As Long, iconIndex As Long, Selected As Boolean)
  Dim itemData As TVITEMEX

  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub
  If prophImageList = 0 Then Exit Sub

  With itemData
    .hItem = hItem
    If Selected Then
      .Mask = TVIF_HANDLE Or TVIF_SELECTEDIMAGE
      .iSelectedImage = iconIndex
    Else
      .Mask = TVIF_HANDLE Or TVIF_IMAGE
      .iImage = iconIndex
    End If
    SendMessage prophWnd, TVM_SETITEM, 0, itemData
  End With
End Sub

Private Sub SetItemIcons(hItem As Long, ImgNormal As Long, ImgSelected As Long)
  Dim itemData As TVITEMEX

  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub
  If prophImageList = 0 Then Exit Sub

  With itemData
    .hItem = hItem
    .Mask = TVIF_HANDLE Or TVIF_IMAGE Or TVIF_SELECTEDIMAGE
    .iImage = ImgNormal
    .iSelectedImage = ImgSelected
    SendMessage prophWnd, TVM_SETITEM, 0, itemData
  End With
End Sub

Private Sub SetItemOverlay(hItem As Long, overlayIndex As Long)
  If prophImageList = 0 Then Exit Sub

  If overlayIndex > 0 Then
    SetItemState hItem, INDEXTOOVERLAYMASK(overlayIndex), TVIS_OVERLAYMASK
  Else
    ' TODO: Mglw. funktioniert das Entfernen eines Overlays unter Vista so nicht
    SetItemState hItem, 0, TVIS_OVERLAYMASK
  End If
End Sub

Private Sub SetItemState(ByVal hItem As Long, ByVal ItemState As Long, ByVal StateMask As Long)
  Dim itemData As TVITEMEX

  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub

  With itemData
    .hItem = hItem
    .Mask = TVIF_HANDLE Or TVIF_STATE
    .state = ItemState
    .StateMask = StateMask
    SendMessage prophWnd, TVM_SETITEM, 0, itemData
  End With
End Sub

Private Sub SetItemText(ByVal hItem As Long, ByVal ItemText As String)
  Dim itemData As TVITEMEX

  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub

  With itemData
    .hItem = hItem
    .pszText = ItemText
    .cchTextMax = Len(.pszText)
    .Mask = TVIF_HANDLE Or TVIF_TEXT
    SendMessage prophWnd, TVM_SETITEM, 0, itemData
  End With
End Sub

Private Sub SetItemType(ByVal hItem As Long, ByVal itemType As Byte)
  Dim itemData As TVITEMEX

  If Not ItemExists(hItem, AllowNullConstants.anIfNoShowRoot) Then Exit Sub

  If hItem Then
    With itemData
      .hItem = hItem
      .Mask = TVIF_HANDLE Or TVIF_PARAM
      SendMessage prophWnd, TVM_GETITEM, 0, itemData
      CopyMemory .lParam, VarPtr(itemType), LenB(itemType)
      SendMessage prophWnd, TVM_SETITEM, 0, itemData
    End With
  Else
    RootItemInfo.itemType = itemType
  End If

  If propShowRoot Then
    If GetRootOrFirstItem = hItem Then RootItemInfo.itemType = itemType
  End If
End Sub

' installiert einen Hook, um Veränderungen an der Shell abzufangen
Private Sub SHNotify_Register()
  Dim Data As SHCHANGENOTIFYENTRY

  SHNotify_UnRegister

  If inIDE Then Exit Sub

  #If UseAutoUpdate Then
    Data.pIDL = pIDL_Desktop
    Data.fRecursive = 1

    ' den Hook installieren
    hSHNotify = SHChangeNotifyRegister(prophWnd, SHCNRFConstants.SHCNRF_InterruptLevel Or SHCNRFConstants.SHCNRF_RecursiveInterrupt Or SHCNRFConstants.SHCNRF_ShellLevel, SHCNEConstants.SHCNE_ALLEVENTS Or SHCNEConstants.SHCNE_INTERRUPT, WM_SHNOTIFY, 1, Data)
  #End If
End Sub

' deinstalliert einen mit "SHNotify_Register" installierten Hook
Private Sub SHNotify_UnRegister()
  If hSHNotify Then
    If SHChangeNotifyUnRegister(hSHNotify) Then
      hSHNotify = 0
    End If
  End If
End Sub

' sortiert alle SubItems von <hParentItem> durch rekursive Aufrufe
Private Sub SortAllSubItems(ByVal hParentItem As Long, Optional ByVal skipExistanceCheck As Boolean = False)
  Dim hItem As Long

  If inIDE Then Exit Sub
  If Not skipExistanceCheck Then
    If Not ItemExists(hParentItem, AllowNullConstants.anIfNoShowRoot) Then Exit Sub
  End If

  SortItems hParentItem

  hItem = ItemGetFirstSubItem(hParentItem)
  While Not IsWithinRange(hItem, -1, 0)
    If ItemHasSubItems(hItem) Then SortAllSubItems hItem, True

    ' nächsten Item ermitteln
    hItem = ItemGetNextItem(hItem)
  Wend
End Sub

' sortiert alle direkten SubItems von <hParentItem>
Private Sub SortItems(ByVal hParentItem As Long)
  Dim Data As TVSORTCB
  Dim pIDLParent As Long

  #If NoSorting Then
    Exit Sub
  #End If

  If inIDE Then Exit Sub
  If Not ItemExists(hParentItem, AllowNullConstants.anIfNoShowRoot) Then Exit Sub

  With Data
    .hParent = hParentItem
    .lParam = ObjPtr(Me)
    If ItemIsPartOfArchive(hParentItem) Or ItemIsArchiveToExpand(hParentItem) Then
      .lpfnCompare = FuncPtr(AddressOf basCallback.Callback_Sort_Archive)
    Else
      #If Debuging Then
        If Not (cachedIShFolderForSorting Is Nothing) Then
          debugger.AddLogEntry "SortItems: Nested sorts! Cache will make trouble!", LogEntryTypeConstants.letError
        End If
      #End If
      .lpfnCompare = FuncPtr(AddressOf basCallback.Callback_Sort_Common)
      Set cachedIShFolderForSorting = Nothing
      pIDLParent = GetFullyQualifiedPIDL(hParentItem, True)
      #If Debuging Then
        Set cachedIShFolderForSorting = GetISHFolderInterfaceFQ(debugger, pIDLParent)
      #Else
        Set cachedIShFolderForSorting = GetISHFolderInterfaceFQ(pIDLParent)
      #End If
    End If
  End With
  #If Debuging Then
    debugger.AddLogEntry "SortItems: Starting sort...", LogEntryTypeConstants.letInfo
  #End If
  SendMessage prophWnd, TVM_SORTCHILDRENCB, 0, Data
  Set cachedIShFolderForSorting = Nothing
  #If Debuging Then
    debugger.AddLogEntry "SortItems: Ended sort", LogEntryTypeConstants.letInfo
  #End If
End Sub

' initiiert einen Drag'n'Drop-Vorgang für <hItem>
Private Sub StartDragDrop(ByVal hItem As Long)
  Const DRAGDROP_S_CANCEL = &H40101
  Const DRAGDROP_S_DROP = &H40100
  Dim arraySize As Long
  Dim dataFmt As FORMATETC
  Dim hr As Long
  Dim IDataObj As IVBDataObject
  Dim IDragSourceHelper As IVBDragSourceHelper
  Dim IDragSourceHelper2 As IVBDragSourceHelper2
  Dim IDrpSrc As IVBDropSource
  Dim IParent As IVBShellFolder
  Dim pBinaryData As Long
  Dim performedEffect As Long
  Dim pIDLToParent As Long
  Dim pt As POINT
  Dim stgData As STGMEDIUM
  Dim supportedEffects As DROPEFFECTConstants
  Dim TVWItemInfo As ITEMINFO_COMMON

  If Not ItemExists(hItem, AllowNullConstants.anNever) Then Exit Sub

  ' Daten über <hItem> sammeln
  If Not GetItemInfo_Common(hItem, TVWItemInfo) Then Exit Sub
  #If Debuging Then
    SplitFullyQualifiedPIDL debugger, TVWItemInfo.pIDLToDesktop, IParent, pIDLToParent
  #Else
    SplitFullyQualifiedPIDL TVWItemInfo.pIDLToDesktop, IParent, pIDLToParent
  #End If
  If IParent Is Nothing Then Exit Sub
  If pIDLToParent = 0 Then Exit Sub
  hDraggedItem = hItem

  With TVWItemInfo
    ' Interface "IDataObject" ermitteln
    IParent.GetUIObjectOf prophWndShellUIParentWindow, 1, pIDLToParent, IID_IDataObject, 0, IDataObj

    If Not (IDataObj Is Nothing) Then
      ' DROPEFFECT_COPY equals SFGAO_CANCOPY and so on...
      supportedEffects = DROPEFFECTConstants.DROPEFFECT_COPY Or DROPEFFECTConstants.DROPEFFECT_LINK Or DROPEFFECTConstants.DROPEFFECT_MOVE
      IParent.GetAttributesOf 1, pIDLToParent, supportedEffects
      supportedEffects = supportedEffects And (DROPEFFECTConstants.DROPEFFECT_COPY Or DROPEFFECTConstants.DROPEFFECT_LINK Or DROPEFFECTConstants.DROPEFFECT_MOVE)

      If supportedEffects <> DROPEFFECTConstants.DROPEFFECT_NONE Then
        If useSHDoDragDrop Then
          #If Debuging Then
            debugger.AddLogEntry "Will call SHDoDragDrop()", LogEntryTypeConstants.letInfo
          #End If
          If ver_Win_Vista Then
            hr = SHDoDragDrop(0, IDataObj, Nothing, supportedEffects, VarPtr(performedEffect))
          Else
            hr = SHDoDragDrop(prophWnd, IDataObj, Nothing, supportedEffects, VarPtr(performedEffect))
          End If
        ElseIf IDSData.pThis Then
          CopyMemory VarPtr(IDrpSrc), VarPtr(IDSData.pThis), 4
          If Not (IDrpSrc Is Nothing) Then
            IDrpSrc.AddRef
            ' fallback to IDragSourceHelper & DoDragDrop
            CoCreateInstance IID_IDragDropHelper, Nothing, CLSCTXConsts.CLSCTX_INPROC_SERVER, IID_IDragSourceHelper, IDragSourceHelper
            If Not (IDragSourceHelper Is Nothing) Then
              GetCursorPos pt
              ScreenToClient prophWnd, pt
              IDragSourceHelper.QueryInterface IID_IDragSourceHelper2, IDragSourceHelper2
              If Not (IDragSourceHelper2 Is Nothing) Then
                IDragSourceHelper2.SetFlags DSH_FLAGSConstants.DSH_ALLOWDROPDESCRIPTIONTEXT
                IDragSourceHelper2.InitializeFromWindow prophWnd, pt, IDataObj
              Else
                IDragSourceHelper.InitializeFromWindow prophWnd, pt, IDataObj
              End If
            End If
            hr = DoDragDrop(IDataObj, IDrpSrc, supportedEffects, VarPtr(performedEffect))
            Set IDragSourceHelper = Nothing
            Set IDragSourceHelper2 = Nothing
            Set IDrpSrc = Nothing
          End If
        End If

        With dataFmt
          .cfFormat = LoWord(RegisterClipboardFormat("Performed DropEffect"))
          .dwAspect = DVASPECTConstants.DVASPECT_CONTENT
          .lindex = -1
        End With
        If IDataObj.GetData(dataFmt, stgData) = S_OK Then
          arraySize = GlobalSize(stgData.Data)
          pBinaryData = GlobalLock(stgData.Data)
          CopyMemory VarPtr(performedEffect), pBinaryData, IIf(arraySize < LenB(performedEffect), arraySize, LenB(performedEffect))
          GlobalUnlock stgData.Data
          ReleaseStgMedium VarPtr(stgData)
        End If

        Select Case hr
          Case DRAGDROP_S_CANCEL
            RaiseItemCompleteDrag hDraggedItem, True, performedEffect
            RaiseCompleteDrag True, performedEffect
          Case DRAGDROP_S_DROP
            RaiseItemCompleteDrag hDraggedItem, False, performedEffect
            RaiseCompleteDrag False, performedEffect
          #If Debuging Then
            Case Else
              debugger.AddLogEntry "(SH)DoDragDrop failed with 0x" & Hex(hr), LogEntryTypeConstants.letError
          #End If
        End Select
      End If

      Set IDataObj = Nothing
    End If
  End With

  hDraggedItem = 0
  InvalidateRectAsLong prophWnd, 0, 1
  Set IParent = Nothing
End Sub

' startet das Abfangen von Nachrichten für das UserControl und das TreeView
Private Sub SubClass()
  If inIDE Then Exit Sub

  FreeItemMem 0

  #If UseSubClassing Then
    StartListening Me, UserControl.hWnd
    StartListening Me, prophWnd
  #End If

  #If UseHooking Then
    If Not CBTHookStarted Then
      installCBTHook Me
      CBTHookStarted = True
    End If
  #End If

  SetTimer prophWnd, TimerID_SubClassParent, 30, 0
  SetTimer prophWnd, TimerID_MouseEvents, 10, 0

  ' für AutoUpdate
  SHNotify_Register
End Sub

' startet das Abfangen von Nachrichten für das LabelEdit-Control
Private Sub SubClass_Edit()
  Dim hMod As Long
  Dim IParent As IVBShellFolder
  Dim TVWItemInfo As ITEMINFO_COMMON

  If inIDE Then Exit Sub

  If hWndEdit Then
    #If UseSubClassing Then
      StartListening Me, hWndEdit
    #End If

    GetWindowRect hWndEdit, rcWnd_Edit
    OffsetRect rcWnd_Edit, -rcWnd_Edit.Left, -rcWnd_Edit.Top

    SetTimer hWndEdit, TimerID_MouseEvents, 10, 0

    If propLimitEditInput Then
      If GetItemInfo_Common(hEditetItem, TVWItemInfo) Then
        #If Debuging Then
          Set IParent = GetParentInterface(debugger, TVWItemInfo.pIDLToDesktop)
        #Else
          Set IParent = GetParentInterface(TVWItemInfo.pIDLToDesktop)
        #End If
        ' SHLimitInputEdit existiert erst ab WinXP
        hMod = LoadLibrary("shell32.dll")
        If hMod Then
          If GetProcAddress(hMod, "SHLimitInputEdit") Then SHLimitInputEdit hWndEdit, IParent
          FreeLibrary hMod
        End If
        Set IParent = Nothing
      End If
    End If
  End If
End Sub

' beendet das Abfangen von Nachrichten für alle Fenster, für die SubClassing aktiviert wurde
Private Sub UnSubClass()
  SHNotify_UnRegister

  KillTimer prophWnd, TimerID_DragExpand
  KillTimer prophWnd, TimerID_DragScroll
  KillTimer prophWnd, TimerID_StateIconChange
  KillTimer prophWnd, TimerID_MouseEvents

  If CBTHookStarted Then
    removeCBTHook Me
    CBTHookStarted = False
  End If

  UnSubClass_Edit

  StopListening Me, UserControl.hWnd
  StopListening Me, prophWnd
  'StopListening Me, hWnd_ParentForm
End Sub

' beendet das Abfangen von Nachrichten für das LabelEdit-Control
Private Sub UnSubClass_Edit()
  KillTimer hWndEdit, TimerID_MouseEvents

  StopListening Me, hWndEdit
End Sub

' prüft für alle Archive unter den SubItems von <hParentItem>, ob sie noch zu den erweiterbaren
' Archiven gehören
Private Sub UpdateAllArchives(ByVal hParentItem As Long, Optional ByVal skipExistanceCheck As Boolean = False)
  Dim hSubItem As Long

  If Not skipExistanceCheck Then
    If Not ItemExists(hParentItem, AllowNullConstants.anIfNoShowRoot) Then Exit Sub
  End If

  hSubItem = ItemGetFirstSubItem(hParentItem)
  While Not IsWithinRange(hSubItem, -1, 0)
    Select Case True
      Case ItemIsFSFile(hSubItem)
        #If Debuging Then
          If IsArchiveToExpandFQ(debugger, GetFullyQualifiedPIDL(hSubItem), propExpandArchives) Then
        #Else
          If IsArchiveToExpandFQ(GetFullyQualifiedPIDL(hSubItem), propExpandArchives) Then
        #End If
          SetItemChildrenProperty hSubItem, True
        Else
          RemoveSubItems hSubItem
          SetItemChildrenProperty hSubItem, False
        End If

        hSubItem = ItemGetNextItem(hSubItem)
      Case ItemIsPartOfArchive(hSubItem)
        hSubItem = -1
      Case Else
        If ItemHasSubItems(hSubItem) Then UpdateAllArchives hSubItem, True

        hSubItem = ItemGetNextItem(hSubItem)
    End Select
  Wend
End Sub

Private Function UpdateSettingsFile(ByVal File As String, Optional ByRef AdditionalData As Variant = Empty) As ErrorTypeConstants
  Dim arrProps_ExCmb() As Byte
  Dim arrProps_ExLvw() As Byte
  Dim arrProps_ExTvw() As Byte
  Dim BytesRead As Long
  Dim BytesWritten As Long
  Dim FilePointerPos As Long
  Dim hFile As Long
  Dim intBuffer As Integer
  Dim Length_ExCmb As Long
  Dim Length_ExLvw As Long
  Dim Length_ExTvw As Long
  Dim lngBuffer As Long
  Dim Props As PropertyBag
  Dim ret As ErrorTypeConstants
  Dim ret2 As Long
  Dim Start_ExCmb As Long
  Dim Start_ExLvw As Long
  Dim Start_ExTvw As Long
  Dim strBuffer As String

  On Error GoTo Fehler

  AdditionalData = Empty

  ' Datei öffnen
  hFile = CreateFileAsLong(File, GENERIC_READ Or GENERIC_WRITE, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
  If hFile = INVALID_HANDLE_VALUE Then GoTo CreateFileError

  ' Signatur speichern
  strBuffer = FileTypeSignature_SettingsFile
  ret2 = WriteFileAsLong(hFile, ByVal strBuffer, Len(strBuffer), BytesWritten, 0)
  If ret2 = 0 Then GoTo WriteFileError
  ' Formatversion speichern
  strBuffer = FileTypeVersion_SettingsFile
  ret2 = WriteFileAsLong(hFile, ByVal strBuffer, Len(strBuffer), BytesWritten, 0)
  If ret2 = 0 Then GoTo WriteFileError

  ' Signatur und Version des Controls, dass die Datei erstellt hat, überspringen
  SetFilePointer hFile, 11, 0, FILE_CURRENT

  ' Signatur dieses Controls speichern - letzte Dateiänderung
  strBuffer = FileCreatedBy_SettingsFile
  ret2 = WriteFileAsLong(hFile, ByVal strBuffer, Len(strBuffer), BytesWritten, 0)
  If ret2 = 0 Then GoTo WriteFileError
  ' Version dieses Controls speichern - letzte Dateiänderung
  intBuffer = MakeWord(App.Minor, App.Major)
  lngBuffer = MakeDWord(App.Revision, intBuffer)
  ret2 = WriteFileAsLong(hFile, ByVal VarPtr(lngBuffer), LenB(lngBuffer), BytesWritten, 0)
  If ret2 = 0 Then GoTo WriteFileError
  ' Buildnumber speichern - letzte Dateiänderung
  intBuffer = BuildNumber
  ret2 = WriteFileAsLong(hFile, ByVal VarPtr(intBuffer), LenB(intBuffer), BytesWritten, 0)
  If ret2 = 0 Then GoTo WriteFileError

  ' Version von ExplorerTreeView speichern, die zum Lesen benötigt wird
  intBuffer = MakeWord(NeededVersion_SettingsFile_Read_Minor_ExTvw, NeededVersion_SettingsFile_Read_Major_ExTvw)
  lngBuffer = MakeDWord(NeededVersion_SettingsFile_Read_Revision_ExTvw, intBuffer)
  ret2 = WriteFileAsLong(hFile, ByVal VarPtr(lngBuffer), LenB(lngBuffer), BytesWritten, 0)
  If ret2 = 0 Then GoTo WriteFileError
  ' Buildnumber speichern
  intBuffer = NeededVersion_SettingsFile_Read_Build_ExTvw
  ret2 = WriteFileAsLong(hFile, ByVal VarPtr(intBuffer), LenB(intBuffer), BytesWritten, 0)
  If ret2 = 0 Then GoTo WriteFileError

  ' Version von ExplorerListView speichern, die zum Lesen benötigt wird
  intBuffer = MakeWord(NeededVersion_SettingsFile_Read_Minor_ExLvw, NeededVersion_SettingsFile_Read_Major_ExLvw)
  lngBuffer = MakeDWord(NeededVersion_SettingsFile_Read_Revision_ExLvw, intBuffer)
  ret2 = WriteFileAsLong(hFile, ByVal VarPtr(lngBuffer), LenB(lngBuffer), BytesWritten, 0)
  If ret2 = 0 Then GoTo WriteFileError
  ' Buildnumber speichern
  intBuffer = NeededVersion_SettingsFile_Read_Build_ExLvw
  ret2 = WriteFileAsLong(hFile, ByVal VarPtr(intBuffer), LenB(intBuffer), BytesWritten, 0)
  If ret2 = 0 Then GoTo WriteFileError

  ' Version von ExplorerComboBox speichern, die zum Lesen benötigt wird
  intBuffer = MakeWord(NeededVersion_SettingsFile_Read_Minor_ExCmb, NeededVersion_SettingsFile_Read_Major_ExCmb)
  lngBuffer = MakeDWord(NeededVersion_SettingsFile_Read_Revision_ExCmb, intBuffer)
  ret2 = WriteFileAsLong(hFile, ByVal VarPtr(lngBuffer), LenB(lngBuffer), BytesWritten, 0)
  If ret2 = 0 Then GoTo WriteFileError
  ' Buildnumber speichern
  intBuffer = NeededVersion_SettingsFile_Read_Build_ExCmb
  ret2 = WriteFileAsLong(hFile, ByVal VarPtr(intBuffer), LenB(intBuffer), BytesWritten, 0)
  If ret2 = 0 Then GoTo WriteFileError

  ' 12 Bytes freilassen
  SetFilePointer hFile, 12, 0, FILE_CURRENT

  ' Version von ExplorerTreeView speichern, die zum Schreiben benötigt wird
  intBuffer = MakeWord(NeededVersion_SettingsFile_Write_Minor_ExTvw, NeededVersion_SettingsFile_Write_Major_ExTvw)
  lngBuffer = MakeDWord(NeededVersion_SettingsFile_Write_Revision_ExTvw, intBuffer)
  ret2 = WriteFileAsLong(hFile, ByVal VarPtr(lngBuffer), LenB(lngBuffer), BytesWritten, 0)
  If ret2 = 0 Then GoTo WriteFileError
  ' Buildnumber speichern
  intBuffer = NeededVersion_SettingsFile_Write_Build_ExTvw
  ret2 = WriteFileAsLong(hFile, ByVal VarPtr(intBuffer), LenB(intBuffer), BytesWritten, 0)
  If ret2 = 0 Then GoTo WriteFileError

  ' Version von ExplorerListView speichern, die zum Schreiben benötigt wird
  intBuffer = MakeWord(NeededVersion_SettingsFile_Write_Minor_ExLvw, NeededVersion_SettingsFile_Write_Major_ExLvw)
  lngBuffer = MakeDWord(NeededVersion_SettingsFile_Write_Revision_ExLvw, intBuffer)
  ret2 = WriteFileAsLong(hFile, ByVal VarPtr(lngBuffer), LenB(lngBuffer), BytesWritten, 0)
  If ret2 = 0 Then GoTo WriteFileError
  ' Buildnumber speichern
  intBuffer = NeededVersion_SettingsFile_Write_Build_ExLvw
  ret2 = WriteFileAsLong(hFile, ByVal VarPtr(intBuffer), LenB(intBuffer), BytesWritten, 0)
  If ret2 = 0 Then GoTo WriteFileError

  ' Version von ExplorerComboBox speichern, die zum Schreiben benötigt wird
  intBuffer = MakeWord(NeededVersion_SettingsFile_Write_Minor_ExCmb, NeededVersion_SettingsFile_Write_Major_ExCmb)
  lngBuffer = MakeDWord(NeededVersion_SettingsFile_Write_Revision_ExCmb, intBuffer)
  ret2 = WriteFileAsLong(hFile, ByVal VarPtr(lngBuffer), LenB(lngBuffer), BytesWritten, 0)
  If ret2 = 0 Then GoTo WriteFileError
  ' Buildnumber speichern
  intBuffer = NeededVersion_SettingsFile_Write_Build_ExCmb
  ret2 = WriteFileAsLong(hFile, ByVal VarPtr(intBuffer), LenB(intBuffer), BytesWritten, 0)
  If ret2 = 0 Then GoTo WriteFileError

  ' 12 Bytes freilassen
  SetFilePointer hFile, 12, 0, FILE_CURRENT

  ' 32 Bytes freilassen
  SetFilePointer hFile, 32, 0, FILE_CURRENT

  ' Speicherplatz für die Adresse des Startbytes der Eigenschaften freilassen (4 Bytes)
  SetFilePointer hFile, 4, 0, FILE_CURRENT

  ' Länge der Eigenschaften speichern
  Set Props = New PropertyBag
  UserControl_WriteProperties Props
  arrProps_ExTvw = Props.Contents
  Length_ExTvw = ArraySize_AsByte(arrProps_ExTvw)
  ret2 = WriteFileAsLong(hFile, ByVal VarPtr(Length_ExTvw), LenB(Length_ExTvw), BytesWritten, 0)
  If ret2 = 0 Then GoTo WriteFileError

  ' Startbyte der ExplorerListView-Daten einlesen
  ret2 = ReadFileAsLong(hFile, ByVal VarPtr(Start_ExLvw), LenB(Start_ExLvw), BytesRead, 0)
  If ret2 = 0 Then GoTo ReadFileError
  ' Länge der ExplorerListView-Daten einlesen
  ret2 = ReadFileAsLong(hFile, ByVal VarPtr(Length_ExLvw), LenB(Length_ExLvw), BytesRead, 0)
  If ret2 = 0 Then GoTo ReadFileError
  ' Startbyte der ExplorerComboBox-Daten einlesen
  ret2 = ReadFileAsLong(hFile, ByVal VarPtr(Start_ExCmb), LenB(Start_ExCmb), BytesRead, 0)
  If ret2 = 0 Then GoTo ReadFileError
  ' Länge der ExplorerComboBox-Daten einlesen
  ret2 = ReadFileAsLong(hFile, ByVal VarPtr(Length_ExCmb), LenB(Length_ExCmb), BytesRead, 0)
  If ret2 = 0 Then GoTo ReadFileError

  If Length_ExLvw Then
    ' ExplorerListView-Daten einlesen
    SetFilePointer hFile, Start_ExLvw, 0, FILE_BEGIN
    ReDim arrProps_ExLvw(Length_ExLvw - 1)
    ret2 = ReadFileAsLong(hFile, ByVal VarPtr(arrProps_ExLvw(LBound(arrProps_ExLvw))), ArraySize_AsByte(arrProps_ExLvw), BytesRead, 0)
    If ret2 = 0 Then GoTo ReadFileError
  End If
  If Length_ExCmb Then
    ' ExplorerComboBox-Daten einlesen
    SetFilePointer hFile, Start_ExCmb, 0, FILE_BEGIN
    ReDim arrProps_ExCmb(Length_ExCmb - 1)
    ret2 = ReadFileAsLong(hFile, ByVal VarPtr(arrProps_ExCmb(LBound(arrProps_ExCmb))), ArraySize_AsByte(arrProps_ExCmb), BytesRead, 0)
    If ret2 = 0 Then GoTo ReadFileError
  End If

  ' ExplorerTreeView-Eigenschaften speichern
  Start_ExTvw = 165
  SetFilePointer hFile, Start_ExTvw, 0, FILE_BEGIN
  ret2 = WriteFileAsLong(hFile, ByVal VarPtr(arrProps_ExTvw(LBound(arrProps_ExTvw))), ArraySize_AsByte(arrProps_ExTvw), BytesWritten, 0)
  Erase arrProps_ExTvw
  If ret2 = 0 Then GoTo WriteFileError
  If Length_ExLvw Then
    ' ExplorerListView-Eigenschaften speichern
    Start_ExLvw = SetFilePointer(hFile, 0, 0, FILE_CURRENT)
    ret2 = WriteFileAsLong(hFile, ByVal VarPtr(arrProps_ExLvw(LBound(arrProps_ExLvw))), ArraySize_AsByte(arrProps_ExLvw), BytesWritten, 0)
    Erase arrProps_ExLvw
    If ret2 = 0 Then GoTo WriteFileError
  End If
  If Length_ExCmb Then
    ' ExplorerComboBox-Eigenschaften speichern
    Start_ExCmb = SetFilePointer(hFile, 0, 0, FILE_CURRENT)
    ret2 = WriteFileAsLong(hFile, ByVal VarPtr(arrProps_ExCmb(LBound(arrProps_ExCmb))), ArraySize_AsByte(arrProps_ExCmb), BytesWritten, 0)
    Erase arrProps_ExCmb
    If ret2 = 0 Then GoTo WriteFileError
  End If

  ' Startbyte der ExplorerTreeView-Eigenschaften speichern
  SetFilePointer hFile, 125, 0, FILE_BEGIN
  ret2 = WriteFileAsLong(hFile, ByVal VarPtr(Start_ExTvw), LenB(Start_ExTvw), BytesWritten, 0)
  If ret2 = 0 Then GoTo WriteFileError
  ' Startbyte der ExplorerListView-Eigenschaften speichern
  SetFilePointer hFile, 133, 0, FILE_BEGIN
  ret2 = WriteFileAsLong(hFile, ByVal VarPtr(Start_ExLvw), LenB(Start_ExLvw), BytesWritten, 0)
  If ret2 = 0 Then GoTo WriteFileError
  ' Startbyte der ExplorerComboBox-Eigenschaften speichern
  SetFilePointer hFile, 141, 0, FILE_BEGIN
  ret2 = WriteFileAsLong(hFile, ByVal VarPtr(Start_ExCmb), LenB(Start_ExCmb), BytesWritten, 0)
  If ret2 = 0 Then GoTo WriteFileError

  ' Datei schließen
  CloseHandle hFile

  UpdateSettingsFile = ErrorTypeConstants.etNone
  Exit Function

CreateFileError:
  ' CreateFile ist fehlgeschlagen
  CloseHandle hFile
  UpdateSettingsFile = ErrorTypeConstants.etCanNotOpenFile
  Exit Function

Fehler:
  ' unbekannter Fehler
  CloseHandle hFile
  AdditionalData = Err.Number
  UpdateSettingsFile = ErrorTypeConstants.etUnknown
  Exit Function

ReadFileError:
  ' ReadFile ist fehlgeschlagen
  CloseHandle hFile
  UpdateSettingsFile = ErrorTypeConstants.etCanNotReadFile
  Exit Function

WriteFileError:
  ' WriteFile ist fehlgeschlagen
  CloseHandle hFile
  UpdateSettingsFile = ErrorTypeConstants.etCanNotWriteFile
End Function

' transformiert <UserPath> in eine pIDL
' <UserPath> kann alles sein, was von der <Path>-Eigenschaft akzeptiert wird
Private Function UserPathTopIDL(ByVal userPath As String, Optional ByVal returnOnlyExactMatch As Boolean = False) As Long
  Dim i As Integer
  Dim pIDL As Long
  Dim ret As Long
  Dim txt As String

  userPath = Trim$(userPath)
  If userPath = "" Then Exit Function

  Select Case True
    Case IncludesCLSID(userPath)
      ' ein CLSID-Pfad - er darf nicht mit der CLSID des Desktops starten
      txt = GetFirstFolders(userPath, 1)
      #If Debuging Then
        If LCase$(CLSIDToDisplayName(txt)) = LCase$(pIDLToDisplayName_Light(debugger, pIDL_Desktop)) Then
      #Else
        If LCase$(CLSIDToDisplayName(txt)) = LCase$(pIDLToDisplayName_Light(pIDL_Desktop)) Then
      #End If
        userPath = RemoveFirstFolders(userPath, 1)
      End If

      userPath = RemoveBackslash(userPath, False)
      ret = DisplayNameTopIDL(IDesktop, userPath)

      If ret = 0 Then
        ' möglicherweise sind DisplayNames angehängt
        i = 1
        Do
          txt = GetFirstFolders(userPath, i)

          pIDL = DisplayNameTopIDL(IDesktop, txt)
          If pIDL Then
            #If Debuging Then
              FreeItemIDList debugger, "UserPathTopIDL #1", ret
            #Else
              FreeItemIDList ret
            #End If
            ret = pIDL
          End If

          i = i + 1
        Loop Until (pIDL = 0) Or (i > CountSegments(userPath))
        i = i - 2

        ' <ret> ist die letzte pIDL, die noch ermittelt werden konnte, <i> ist die Anzahl der
        ' verarbeiteten Segmente
        userPath = GetLastFolders(userPath, CountSegments(userPath) - i)

        #If Debuging Then
          ret = AddDisplayNamesTopIDL(debugger, ret, userPath, returnOnlyExactMatch)
        #Else
          ret = AddDisplayNamesTopIDL(ret, userPath, returnOnlyExactMatch)
        #End If
      End If

    Case IsDirectory(userPath)
      ' ein "normaler" Pfad
      ret = PathTopIDL(userPath)

    Case IsFTPURL(userPath)
      ' eine FTP-URL
      ret = PathTopIDL(userPath)

    #If Debuging Then
      Case IsCSIDL(debugger, userPath)
        ' eine "echte" CSIDL
        ret = CSIDLTopIDL(userPath)

      Case IsCSIDL(debugger, userPath, True)
        ' ein Eintrag aus dem Kontextmenü
        ret = CSIDLTopIDL(MenuItemToCSIDL(userPath))
    #Else
      Case IsCSIDL(userPath)
        ' eine "echte" CSIDL
        ret = CSIDLTopIDL(userPath)

      Case IsCSIDL(userPath, True)
        ' ein Eintrag aus dem Kontextmenü
        ret = CSIDLTopIDL(MenuItemToCSIDL(userPath))
    #End If

    Case Else
      ' evt. ein Pfad aus DisplayNames
      ' -> nur gültig für den Desktop
      '    -> müssen mit "Desktop" beginnen
      #If Debuging Then
        txt = pIDLToDisplayName_Light(debugger, pIDL_Desktop)
      #Else
        txt = pIDLToDisplayName_Light(pIDL_Desktop)
      #End If
      If LCase$(Left$(userPath, Len(txt))) = LCase$(txt) Then
        ' Pfad aus DisplayNames
        ret = ILClone(pIDL_Desktop)
        userPath = RemoveFirstFolders(userPath, 1)
      Else
        ' möglicherweise ein Pfad, an den DisplayNames angehängt sind
        i = 1
        Do
          pIDL = PathTopIDL(AddBackslash(GetFirstFolders(userPath, i), True))
          If pIDL Then
            #If Debuging Then
              FreeItemIDList debugger, "UserPathTopIDL #2", ret
            #Else
              FreeItemIDList ret
            #End If
            ret = pIDL
          End If

          i = i + 1
        Loop Until (pIDL = 0) Or (i > CountSegments(userPath))
        i = i - 2

        ' <ret> ist die letzte pIDL, die noch ermittelt werden konnte, <i> ist die Anzahl der
        ' verarbeiteten Segmente
        userPath = GetLastFolders(userPath, CountSegments(userPath) - i)
      End If
      #If Debuging Then
        ret = AddDisplayNamesTopIDL(debugger, ret, userPath, returnOnlyExactMatch)
      #Else
        ret = AddDisplayNamesTopIDL(ret, userPath, returnOnlyExactMatch)
      #End If
  End Select

  UserPathTopIDL = ret
End Function

' wird bei eingehenden Nachrichten aufgerufen
' für LabelEdit-Control
Private Function WinMainEdit(ByVal iListener As Long, ByVal oldWinMain As Long, ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Const WM_CHAR = &H102
  Const WM_CONTEXTMENU = &H7B
  Const WM_KEYDOWN = &H100
  Const WM_KEYUP = &H101
  Const WM_LBUTTONDBLCLK = &H203
  Const WM_LBUTTONDOWN = &H201
  Const WM_LBUTTONUP = &H202
  Const WM_MBUTTONDBLCLK = &H209
  Const WM_MBUTTONDOWN = &H207
  Const WM_MBUTTONUP = &H208
  Const WM_MOUSEMOVE = &H200
  Const WM_NCDESTROY = &H82
  Const WM_NCMOUSEMOVE = &HA0
  Const WM_RBUTTONDBLCLK = &H206
  Const WM_RBUTTONDOWN = &H204
  Const WM_RBUTTONUP = &H205
  Const WM_SIZE = &H5
  Const WM_SYSKEYDOWN = &H104
  Const WM_SYSKEYUP = &H105
  Const WM_TIMER = &H113
  Dim boolTmp As Boolean
  Dim Button As MouseButtonConstants
  Dim Cancel As Boolean
  Dim lngTmp As Long
  Dim pt As POINT
  Dim Shift As ShiftConstants
  Dim tmp As Integer

  Select Case MSG
    Case WM_CHAR
      ' eine ANSI-Taste wurde gedrückt und wieder losgelassen
      tmp = CInt(wParam)
      RaiseEditKeyPress tmp
      wParam = tmp
      If wParam = 0 Then Exit Function

    Case WM_CONTEXTMENU
      ' das Kontextmenü soll angezeigt werden
      TimeOfLastWM_CONTEXTMENU_Edit = GetTickCount
      RaiseBeforeShowEditContextMenu Cancel
      If Cancel Then Exit Function

    Case WM_KEYDOWN, WM_SYSKEYDOWN
      ' eine Taste wurde gedrückt
      GetShiftAndMouseButtonBits 0, Shift, 0, True
      tmp = CInt(wParam)
      RaiseEditKeyDown tmp, Shift
      wParam = tmp
      If wParam = 0 Then Exit Function

    Case WM_KEYUP, WM_SYSKEYUP
      ' eine Taste wurde losgelassen
      GetShiftAndMouseButtonBits 0, Shift, 0, True
      tmp = CInt(wParam)
      RaiseEditKeyUp tmp, Shift
      wParam = tmp
      If wParam = 0 Then Exit Function

    Case WM_LBUTTONDBLCLK, WM_RBUTTONDBLCLK, WM_MBUTTONDBLCLK
      ' ein Doppelklick
      If MSG = WM_LBUTTONDBLCLK Then RaiseEditDblClick
      If MSG = WM_MBUTTONDBLCLK Then RaiseEditMDblClick
      If MSG = WM_RBUTTONDBLCLK Then RaiseEditRDblClick

    Case WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN
      ' eine Maustaste wurde gedrückt
      GetShiftAndMouseButtonBits wParam, Shift, Button
      If MSG = WM_LBUTTONDOWN Then Button = MouseButtonConstants.vbLeftButton
      If MSG = WM_MBUTTONDOWN Then Button = MouseButtonConstants.vbMiddleButton
      If MSG = WM_RBUTTONDOWN Then Button = MouseButtonConstants.vbRightButton
      RaiseEditMouseDown Button, Shift, LoWord(lParam), HiWord(lParam)

    Case WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP
      ' eine Maustaste wurde losgelassen
      GetShiftAndMouseButtonBits wParam, Shift, Button
      If MSG = WM_LBUTTONUP Then Button = MouseButtonConstants.vbLeftButton
      If MSG = WM_MBUTTONUP Then Button = MouseButtonConstants.vbMiddleButton
      If MSG = WM_RBUTTONUP Then Button = MouseButtonConstants.vbRightButton
      RaiseEditMouseUp Button, Shift, LoWord(lParam), HiWord(lParam)

    Case WM_MOUSEMOVE
      ' der Mauszeiger wurde über dem Control bewegt
      With ptLastMousePos_Edit
        TimeOfLastMouseMove_Edit = GetTickCount
        .x = LoWord(lParam)
        .y = HiWord(lParam)
        ptLastMousePos = ptLastMousePos_Edit
        ClientToScreen hWndEdit, ptLastMousePos
        ScreenToClient prophWnd, ptLastMousePos

        If Not checkForMouseLeave_Edit Then
          TimeOfPossibleMouseLeave_Edit = -1
          checkForMouseLeave_Edit = True
        End If
        GetShiftAndMouseButtonBits wParam, Shift, Button
        If Not propMouseOverControl_Edit Then
          RaiseEditMouseEnter Button, Shift, .x, .y

          TimeOfFirstMouseMove_Edit = GetTickCount
        End If

        pt = ptLastMousePos_Edit
        Dragging_Edit = (Button <> 0) And (Not IsWithinVisibleClientArea(hWndEdit, rcWnd_Edit, pt, False))

        RaiseEditMouseMove Button, Shift, .x, .y
      End With

    Case WM_NCDESTROY
      UnSubClass_Edit

    Case WM_NCMOUSEMOVE
      ' der Mauszeiger wurde über dem Control bewegt
      With ptLastMousePos_Edit
        TimeOfLastMouseMove_Edit = GetTickCount
        .x = LoWord(lParam)
        .y = HiWord(lParam)
        ptLastMousePos = ptLastMousePos_Edit
        ScreenToClient prophWnd, ptLastMousePos
        ScreenToClient hWndEdit, ptLastMousePos_Edit

        If Not checkForMouseLeave_Edit Then
          TimeOfPossibleMouseLeave_Edit = -1
          checkForMouseLeave_Edit = True
        End If
        GetShiftAndMouseButtonBits wParam, Shift, Button
        If Not propMouseOverControl_Edit Then
          RaiseEditMouseEnter Button, Shift, .x, .y

          TimeOfFirstMouseMove_Edit = GetTickCount
        End If

        RaiseEditMouseMove Button, Shift, .x, .y
      End With

    Case WM_SIZE
      GetWindowRect hWndEdit, rcWnd_Edit
      OffsetRect rcWnd_Edit, -rcWnd_Edit.Left, -rcWnd_Edit.Top

    Case WM_TIMER
      Select Case wParam
        Case TimerID_MouseEvents
          If checkForMouseLeave_Edit Then
            If TimeOfPossibleMouseLeave_Edit = -1 Then
              ' prüfen, ob die aktuelle Position des Mauszeigers von seiner Position beim letzten
              ' WM_MOUSEMOVE abweicht
              ' -> das wäre ein Hinweis auf das Verlassen des Controls
              GetCursorPos pt
              ScreenToClient hWndEdit, pt
              If Not IsSamePoint(pt, ptLastMousePos_Edit) Then
                boolTmp = True
              Else
                ' möglicherweise gehört <pt> inzwischen zu einem anderen Fenster
                If Not IsWithinVisibleClientArea(hWndEdit, rcWnd_Edit, pt, False, hWnd_EditContextMenu, hWnd_ScrollBarContextMenu, , , , lngTmp) Then
                  boolTmp = releasedMouseButton_Edit
                End If
              End If

              If boolTmp And Not blockMouseLeave_Edit Then TimeOfPossibleMouseLeave_Edit = GetTickCount
            ElseIf TimeOfLastMouseMove_Edit < TimeOfPossibleMouseLeave_Edit Then
              ' seit der 1. Überprüfung wurde der Mauszeiger nicht mehr über dem Control bewegt
              ' -> sollte wieder alles auf ein Verlassen des Controls hindeuten, die Auslösung von
              ' MouseLeave einleiten
              GetCursorPos pt
              ScreenToClient hWndEdit, pt
              If Not IsSamePoint(pt, ptLastMousePos_Edit) Then
                boolTmp = True
              Else
                ' möglicherweise gehört <pt> inzwischen zu einem anderen Fenster
                If Not IsWithinVisibleClientArea(hWndEdit, rcWnd_Edit, pt, False, hWnd_EditContextMenu, hWnd_ScrollBarContextMenu, , , , lngTmp) Then
                  boolTmp = releasedMouseButton_Edit
                End If
              End If

              If boolTmp And Not blockMouseLeave_Edit Then
                With ptLastMousePos_Edit
                  WinMainUC 1, 0, hWndEdit, EXTVM_MYMOUSELEAVE, -1, MakeDWord(.x, .y)
                End With
              End If
            Else
              ' wir haben inzwischen wieder eine WM_MOUSEMOVE empfangen -> es war wohl ein Fehlalarm
              ' Ein Fehlalarm kann entstehen, wenn der User die Maus bewegt hat, aber WM_MOUSEMOVE erst
              ' nach WM_TIMER gesendet wurde.
              TimeOfPossibleMouseLeave_Edit = -1
              releasedMouseButton_Edit = False
            End If
          End If

          If propMouseOverControl_Edit And Not propHot_Edit Then
            ' vor dem MouseHover-Event HoverTime ablaufen lassen
            lngTmp = GetTickCount - GetSystemHoverTime
            If lngTmp >= TimeOfFirstMouseMove_Edit Then
              ' HoverTime ist vorbei -> Auslösung des MouseHover-Events einleiten
              GetCursorPos pt
              ScreenToClient hWndEdit, pt
              WinMainUC 1, 0, hWndEdit, EXTVM_MYMOUSEHOVER, -1, MakeDWord(pt.x, pt.y)
            End If
          End If
          Exit Function
      End Select
  End Select

DoDefault:
  If iListener = 1 Then
    If oldWinMain Then WinMainEdit = CallWindowProc(oldWinMain, hWnd, MSG, wParam, lParam)
  End If
End Function

' wird bei eingehenden Nachrichten aufgerufen
' für die oberste Form
Private Function WinMainParent(ByVal iListener As Long, ByVal oldWinMain As Long, ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Const WM_WININICHANGE = &H1A
  Const WM_SETTINGCHANGE = WM_WININICHANGE

  Select Case MSG
    Case WM_SETTINGCHANGE
      ' die Systemeinstellungen für die Darstellung wurden geändert
      ' -> könnte mit AutoUpdate zusammenhängen
      ' -> alle Items updaten und SysImageList neu laden
      If wParam = SPI_SETNONCLIENTMETRICS Then
        If lParam = 0 Then
          InvalidateImgLstAndIcons
        ElseIf LCase(GetStrFromPointer(lParam)) = "windowmetrics" Then
          InvalidateImgLstAndIcons
        End If
      End If
  End Select

DoDefault:
  If iListener = 1 Then
    If oldWinMain Then WinMainParent = CallWindowProc(oldWinMain, hWnd, MSG, wParam, lParam)
  End If
End Function

' wird bei eingehenden Nachrichten aufgerufen
' für das TreeView
Private Function WinMainTree(ByVal iListener As Long, ByVal oldWinMain As Long, ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Const MA_NOACTIVATE = 3
  Const VK_LEFT = &H25
  Const WM_CHAR = &H102
  Const WM_CONTEXTMENU = &H7B
  Const WM_DRAWITEM = &H2B
  Const WM_HSCROLL = &H114
  Const WM_INITMENUPOPUP = &H117
  Const WM_KEYDOWN = &H100
  Const WM_KEYUP = &H101
  Const WM_LBUTTONDBLCLK = &H203
  Const WM_LBUTTONDOWN = &H201
  Const WM_LBUTTONUP = &H202
  Const WM_MBUTTONDBLCLK = &H209
  Const WM_MBUTTONDOWN = &H207
  Const WM_MBUTTONUP = &H208
  Const WM_MEASUREITEM = &H2C
  Const WM_MENUCHAR = &H120
  Const WM_MENUSELECT = &H11F
  Const WM_MOUSEACTIVATE = &H21
  Const WM_MOUSEMOVE = &H200
  #If BkImageSupport Then
    Const WM_MOUSEWHEEL = &H20A
  #End If
  Const WM_NCMOUSEMOVE = &HA0
  Const WM_NOTIFY = &H4E
  Const WM_RBUTTONDBLCLK = &H206
  Const WM_RBUTTONDOWN = &H204
  Const WM_RBUTTONUP = &H205
  Const WM_SETCURSOR = &H20
  Const WM_SETFOCUS = &H7
  Const WM_SIZE = &H5
  Const WM_SYSCOLORCHANGE = &H15
  Const WM_SYSKEYDOWN = &H104
  Const WM_SYSKEYUP = &H105
  Const WM_TIMER = &H113
  Const WM_UNINITMENUPOPUP = &H125
  Const WM_VSCROLL = &H115
  Dim boolTmp As Boolean
  Dim Button As MouseButtonConstants
  Dim Cancel As Boolean
  Dim ClipRect As RECT
  Dim CmdID As Long
  Dim Data As NMHDR
'  Dim DragImg As SHDRAGIMAGE
  Dim f As Long
  Dim FrameInfo As OLEINPLACEFRAMEINFO
  Dim hClickedItem As Long
  Dim hCursor As Long
  Dim hImgLst As Long
  Dim hItem As Long
  Dim HTInfo As TVHITTESTINFO
  Dim intTmp As Integer
  Dim IOleInPlaceActiveObject As IVBOleInPlaceActiveObject
  Dim IOleInPlaceFrame As IVBOleInPlaceFrame
  Dim IOleInPlaceSite As IVBOleInPlaceSite
  Dim IOleInPlaceUIWindow As IVBOleInPlaceUIWindow
  Dim IOleObject As IVBOleObject
  Dim ItemDescr As String
  Dim ItemText As String
  Dim ItemVerb As String
  Dim lngTmp As Long
  Dim NumStateIcons As Long
  Dim pt As POINT
  Dim rc As RECT
  Dim Shift As ShiftConstants
  #If Debuging Then
    Dim SHNotData As SHNOTIFY
  #End If
  Dim ToolTipData As NMTTDISPINFO
  Static hItem_StateIconChange As Long
  Static NewStateIcon As Long
  Static OldStateIcon As Long

  Select Case MSG
'    Case DI_GETDRAGIMAGE
'      CopyMemory VarPtr(DragImg), lParam, LenB(DragImg)

    Case WM_CHAR
      ' eine ANSI-Taste wurde gedrückt und wieder losgelassen
      intTmp = CInt(wParam)
      RaiseKeyPress intTmp
      wParam = intTmp
      If wParam = 0 Then Exit Function

    Case WM_CONTEXTMENU
      ' Kontextmenü anzeigen
      TimeOfLastWM_CONTEXTMENU_Shell = GetTickCount
      If propShowShellContextMenus Then
        With pt
          .x = LoWord(lParam)
          .y = HiWord(lParam)
          If (.x = -1) And (.y = -1) Then
            ' Nachricht wurde per Tastatur erzeugt
            ' -> wird in WinMainUC behandelt
          ElseIf propAllowDragDrop Then
            ' Nachricht wurde per Maus erzeugt
            ScreenToClient prophWnd, pt
            hItem = ItemHitTest(pt, TVHT_ONITEM)
            ClientToScreen prophWnd, pt
            ItemShowShellContextMenu hItem, .x, .y
            Exit Function
          End If
        End With
      End If

    Case WM_DRAWITEM, WM_INITMENUPOPUP, WM_MEASUREITEM, WM_MENUCHAR
      ' ein MenuItem muß gezeichnet werden
      If IContextMenu3 Is Nothing Then
        If Not (IContextMenu2 Is Nothing) Then IContextMenu2.HandleMenuMsg MSG, wParam, lParam
      Else
        IContextMenu3.HandleMenuMsg2 MSG, wParam, lParam, 0
      End If

    #If BkImageSupport Then
      Case WM_HSCROLL, WM_MOUSEWHEEL, WM_VSCROLL
        ' für Hintergrundbild
        If Not (propBkImage Is Nothing) Then InvalidateRectAsLong prophWnd, 0, 1
    #End If

    Case WM_KEYDOWN, WM_SYSKEYDOWN
      ' eine Taste wurde gedrückt
      GetShiftAndMouseButtonBits 0, Shift, 0, True
      intTmp = CInt(wParam)
      RaiseKeyDown intTmp, Shift
      wParam = intTmp

      If IsKeyHoldDown(VK_LEFT) Then
        If propSingleExpand = SingleExpandConstants.seWinXPStyle Then
          hItem = Me.SelectedItem
          If Me.ItemIsExpanded(hItem) Then
            ' collapse
            Me.ItemCollapse hItem
            ' eat the message
            wParam = 0
          End If
        End If
      End If
      If wParam = 0 Then Exit Function

      If propCheckBoxes Then
        If (intTmp = KeyCodeConstants.vbKeySpace) And ((Shift And ShiftConstants.vbCtrlMask) = 0) Then
          hItem = SelectedItem
          OldStateIcon = ItemStateIconIndex(hItem)
          hImgLst = SendMessageAsLong(prophWnd, TVM_GETIMAGELIST, TVSIL_STATE, 0)
          If hImgLst Then
            NumStateIcons = ImageList_GetImageCount(hImgLst)
            NewStateIcon = ((OldStateIcon + 1) Mod NumStateIcons)
            If NewStateIcon = 0 Then NewStateIcon = 1
            intTmp = NewStateIcon

            RaiseItemBeforeStateIconChange hItem, OldStateIcon, NewStateIcon, Cancel
            If Not Cancel Then
              If oldWinMain Then WinMainTree = CallWindowProc(oldWinMain, hWnd, MSG, wParam, lParam)
              If NewStateIcon <> intTmp Then ItemStateIconIndex(hItem) = NewStateIcon
              NewStateIcon = ItemStateIconIndex(hItem)
              RaiseItemStateIconChanged hItem, OldStateIcon, NewStateIcon
            End If
            Exit Function
          End If
        End If
      End If

    Case WM_KEYUP, WM_SYSKEYUP
      ' eine Taste wurde losgelassen
      GetShiftAndMouseButtonBits 0, Shift, 0, True
      intTmp = CInt(wParam)
      RaiseKeyUp intTmp, Shift
      wParam = intTmp
      If wParam = 0 Then Exit Function

    Case WM_LBUTTONDBLCLK, WM_MBUTTONDBLCLK, WM_RBUTTONDBLCLK
      ' ein Doppelklick
      If MSG = WM_LBUTTONDBLCLK Then RaiseDblClick
      If MSG = WM_MBUTTONDBLCLK Then RaiseMDblClick
      If MSG = WM_RBUTTONDBLCLK Then RaiseRDblClick

    Case WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN
      ' eine Maustaste wurde gedrückt
      GetShiftAndMouseButtonBits wParam, Shift, Button
      If MSG = WM_LBUTTONDOWN Then Button = MouseButtonConstants.vbLeftButton
      If MSG = WM_MBUTTONDOWN Then Button = MouseButtonConstants.vbMiddleButton
      If MSG = WM_RBUTTONDOWN Then Button = MouseButtonConstants.vbRightButton
      With pt
        .x = LoWord(lParam)
        .y = HiWord(lParam)
        hItem = ItemHitTest(pt, TVHT_ONITEM)

        RaiseMouseDown Button, Shift, .x, .y, hItem

        If (MSG = WM_LBUTTONDOWN) And (propSingleExpand = SingleExpandConstants.seWinXPStyle) Then
          With HTInfo
            .pt = pt
            hClickedItem = SendMessage(prophWnd, TVM_HITTEST, 0, HTInfo)
            f = .Flags
          End With

          If Me.ItemIsExpanded(hClickedItem) And (hClickedItem = Me.SelectedItem) Then
            If (TVHT_ONITEMICON Or TVHT_ONITEMLABEL) And f Then
              Exit Function
            ElseIf propFullRowSelect Then
              If (TVHT_ONITEMINDENT Or TVHT_ONITEMRIGHT) And f Then Exit Function
            End If
          End If

          ' we should always behave as we do if FullRowSelect = True
          If ((TVHT_ONITEMINDENT Or TVHT_ONITEMICON Or TVHT_ONITEMLABEL Or TVHT_ONITEMRIGHT) And f) = 0 Then
            hClickedItem = 0
          End If
          If Not propFullRowSelect Then
            If (TVHT_ONITEMINDENT Or TVHT_ONITEMRIGHT) And f Then
              If hClickedItem Then
                SendMessageAsLong prophWnd, TVM_SELECTITEM, TVGN_CARET, hClickedItem
              End If
            End If
          End If
        End If

        If MSG = WM_LBUTTONDOWN Then
          hItem = HitTest(.x, .y, TVHT_ONITEMSTATEICON)
          hItem_StateIconChange = hItem
          If Not propAllowDragDrop Then
            If ItemExists(hItem, AllowNullConstants.anNever) Then
              If propCheckBoxes Then
                OldStateIcon = ItemStateIconIndex(hItem)
                hImgLst = SendMessageAsLong(prophWnd, TVM_GETIMAGELIST, TVSIL_STATE, 0)
                If hImgLst Then
                  NumStateIcons = ImageList_GetImageCount(hImgLst)
                  NewStateIcon = ((OldStateIcon + 1) Mod NumStateIcons)
                  If NewStateIcon = 0 Then NewStateIcon = 1
                  intTmp = NewStateIcon

                  RaiseItemBeforeStateIconChange hItem, OldStateIcon, NewStateIcon, Cancel
                  If Not Cancel Then
                    If oldWinMain Then WinMainTree = CallWindowProc(oldWinMain, hWnd, MSG, wParam, lParam)
                    If NewStateIcon <> intTmp Then ItemStateIconIndex(hItem) = NewStateIcon
                    NewStateIcon = ItemStateIconIndex(hItem)
                    RaiseItemStateIconChanged hItem, OldStateIcon, NewStateIcon
                  End If
                  Exit Function
                End If
              End If
            End If
          End If
        End If
      End With

    Case WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP
      ' eine Maustaste wurde losgelassen
      If iListener = 1 Then
        If oldWinMain Then WinMainTree = CallWindowProc(oldWinMain, hWnd, MSG, wParam, lParam)
      End If

      GetShiftAndMouseButtonBits wParam, Shift, Button
      If MSG = WM_LBUTTONUP Then Button = MouseButtonConstants.vbLeftButton
      If MSG = WM_MBUTTONUP Then Button = MouseButtonConstants.vbMiddleButton
      If MSG = WM_RBUTTONUP Then Button = MouseButtonConstants.vbRightButton
      With pt
        .x = LoWord(lParam)
        .y = HiWord(lParam)
        hItem = ItemHitTest(pt, TVHT_ONITEM)

        RaiseMouseUp Button, Shift, .x, .y, hItem

        If (MSG = WM_LBUTTONUP) And (propAllowDragDrop = True) Then
          hItem = ItemHitTest(pt, TVHT_ONITEMSTATEICON)
          If hItem = hItem_StateIconChange Then   'If hItem = hCollapsingItem Then
            If ItemExists(hItem, AllowNullConstants.anNever) Then
              If propCheckBoxes Then
                OldStateIcon = ItemStateIconIndex(hItem)
                hImgLst = SendMessageAsLong(prophWnd, TVM_GETIMAGELIST, TVSIL_STATE, 0)
                If hImgLst Then
                  NumStateIcons = ImageList_GetImageCount(hImgLst)
                  NewStateIcon = ((OldStateIcon + 1) Mod NumStateIcons)
                  If NewStateIcon = 0 Then NewStateIcon = 1
                  intTmp = NewStateIcon

                  RaiseItemBeforeStateIconChange hItem, OldStateIcon, NewStateIcon, Cancel
                  If Not Cancel Then
                    If iListener Then
                      If oldWinMain Then WinMainTree = CallWindowProc(oldWinMain, hWnd, MSG, wParam, lParam)
                    End If
                    'hItem_StateIconChange = hItem
                    SetTimer prophWnd, TimerID_StateIconChange, 10, 0
                  End If
                End If
              End If
            End If
          End If
        End If
      End With
      'hCollapsingItem = 0
      Exit Function

    Case WM_MENUSELECT
      ' ein MenuItem wurde markiert
      ' entweder handelt es sich um das Shell-Kontextmenü oder um das Kontextmenü der ScrollBar
      If lParam = prophShellContextMenu Then
        CmdID = LoWord(wParam)

        If (IsSeparator(lParam, CmdID) = False) And (CmdID > 0) Then
          ItemDescr = String$(MAX_PATH, Chr$(0))
          ItemVerb = String$(MAX_PATH, Chr$(0))
          If Not (IContextMenu Is Nothing) Then
            IContextMenu.GetCommandString CmdID - 1, GCSConstants.GCS_HELPTEXT Or GCSConstants.GCS_UNICODE, 0, StrPtr(ItemDescr), Len(ItemDescr)
            IContextMenu.GetCommandString CmdID - 1, GCSConstants.GCS_VERB Or GCSConstants.GCS_UNICODE, 0, StrPtr(ItemVerb), Len(ItemVerb)
          End If
          ItemDescr = Left$(ItemDescr, lstrlenA(ItemDescr))
          ItemVerb = Left$(ItemVerb, lstrlenA(ItemVerb))

          ItemText = GetMenuItemText(lParam, CmdID)
        End If

        ' ein Item wurde betreten
        RaiseItemShellContextMenuItemEnter hItemWithContMenu, lParam, CmdID, ItemText, ItemVerb, ItemDescr
      End If

    Case WM_MOUSEACTIVATE
      ' das TreeView wurde mit einem Mausklick aktiviert
      If (GetFocus <> prophWnd) And (hWndEdit = 0) Then
        SetFocusAPI UserControl.hWnd
        WinMainTree = MA_NOACTIVATE
        Exit Function
      End If

    Case WM_MOUSEMOVE
      ' der Mauszeiger wurde über dem TreeView bewegt
      With ptLastMousePos
        TimeOfLastMouseMove = GetTickCount
        .x = LoWord(lParam)
        .y = HiWord(lParam)
        hItem = ItemHitTest(ptLastMousePos, TVHT_ONITEM)

        If Not checkForMouseLeave Then
          TimeOfPossibleMouseLeave = -1
          checkForMouseLeave = True
        End If
        GetShiftAndMouseButtonBits wParam, Shift, Button
        If Not propMouseOverControl Then
          RaiseMouseEnter Button, Shift, .x, .y

          TimeOfFirstMouseMove = GetTickCount
        End If

        pt = ptLastMousePos
        ' wird vom TreeView scheinbar eh nicht unterstützt
        Dragging = (Button <> 0) And (Not IsWithinVisibleClientArea(prophWnd, rcWnd, pt, False))

        RaiseMouseMove Button, Shift, .x, .y, hItem
      End With

    Case WM_NCMOUSEMOVE
      ' der Mauszeiger wurde über dem TreeView bewegt
      With ptLastMousePos
        TimeOfLastMouseMove = GetTickCount
        .x = LoWord(lParam)
        .y = HiWord(lParam)
        ScreenToClient prophWnd, ptLastMousePos

        If Not checkForMouseLeave Then
          TimeOfPossibleMouseLeave = -1
          checkForMouseLeave = True
        End If
        GetShiftAndMouseButtonBits wParam, Shift, Button
        If Not propMouseOverControl Then
          RaiseMouseEnter Button, Shift, .x, .y

          TimeOfFirstMouseMove = GetTickCount
        End If

        RaiseMouseMove Button, Shift, .x, .y, -1
      End With

    Case WM_NOTIFY
      CopyMemory VarPtr(Data), lParam, LenB(Data)
      Select Case Data.code
        Case TTN_GETDISPINFO, TTN_GETDISPINFOW
          ' das ToolTip-Control braucht den ToolTipText
          CopyMemory VarPtr(ToolTipData), lParam, LenB(ToolTipData)
          With ToolTipData
            If oldWinMain Then WinMainTree = CallWindowProc(oldWinMain, hWnd, MSG, wParam, lParam)
            If canceledToolTip Then
              ' keinen ToolTip anzeigen
              ' dazu einfach die alten Daten zurückschreiben
              CopyMemory lParam, VarPtr(ToolTipData), LenB(ToolTipData)
              WinMainTree = 0
            End If
          End With
          Exit Function
      End Select

    Case EXTVM_SETAUTOUPDATE
      ' hidden feature: enable/disable AutoUpdate
      WinMainTree = Abs(Not propNoAutoUpdate)
      propNoAutoUpdate = Not CBool(wParam)
      Exit Function

    Case WM_SETCURSOR
      ' Möglichkeit, den Mauszeiger anzupassen
      If Not propMouseOverControl_Edit Then
        If prophShellContextMenu Then
          hCursor = LoadImage(0, MAKEINTRESOURCE(OCR_NORMAL), IMAGE_CURSOR, 0, 0, LR_DEFAULTCOLOR Or LR_DEFAULTSIZE Or LR_SHARED)
        ElseIf propReplaceHandCursor Then
          If propMouseIcon Is Nothing Then
            hCursor = MousePointerConstTohCursor(propMousePointer)
          Else
            hCursor = propMouseIcon.Handle
          End If
        End If

        If hCursor Then
          SetCursor hCursor

          WinMainTree = 1
          Exit Function
        End If
      End If

    Case WM_SETFOCUS
      ' das TreeView hat den Fokus erhalten
'      If hWnd = prophWnd Then dontExpand = (propSingleExpand <> SingleExpandConstants.seNone)
      If oldWinMain Then WinMainTree = CallWindowProc(oldWinMain, hWnd, MSG, wParam, lParam)
      dontExpand = False
      ' ?????
      If Not isSettingFocus Then
        UserControl.SetFocus
      Else
        If hWnd = prophWnd Then
          Set IOleObject = Me
          IOleObject.GetClientSite IOleInPlaceSite
          If Not (IOleInPlaceSite Is Nothing) Then
            IOleInPlaceSite.GetWindowContext IOleInPlaceFrame, IOleInPlaceUIWindow, rc, ClipRect, FrameInfo
            If IPAOData.pThis Then
              CopyMemory VarPtr(IOleInPlaceActiveObject), VarPtr(IPAOData.pThis), 4
              If Not (IOleInPlaceActiveObject Is Nothing) Then
                IOleInPlaceActiveObject.AddRef
                If Not (IOleInPlaceFrame Is Nothing) Then
                  IOleInPlaceFrame.SetActiveObject IOleInPlaceActiveObject, vbNullString
                  If Not (IOleInPlaceUIWindow Is Nothing) Then
                    IOleInPlaceUIWindow.SetActiveObject IOleInPlaceActiveObject, vbNullString
                  End If
                End If
                Set IOleInPlaceActiveObject = Nothing
              End If
            End If
          End If
        End If
        isSettingFocus = False
        Exit Function
      End If

    Case EXTVM_SETROOTPATH
      If lParam Then
        If wParam Then
          propRootPath = String$(wParam, Chr$(0))
          lstrcpyAsLong2 propRootPath, lParam
          #If Debuging Then
            debugger.AddLogEntry "Received EXTVM_SETROOTPATH", LogEntryTypeConstants.letInfo
            debugger.AddLogEntry "   " & propRootPath, LogEntryTypeConstants.letOther
            debugger.AddLogEntry "   Will reload items now!", LogEntryTypeConstants.letOther
          #End If
          ReloadItems
        End If

        HeapFree GetProcessHeap, 0, lParam
      End If
      Exit Function

    Case WM_SHNOTIFY
      ' für AutoUpdate
      If Not propNoAutoUpdate Then
        #If Debuging Then
          CopyMemory VarPtr(SHNotData), wParam, LenB(SHNotData)
          If (lParam = SHCNEConstants.SHCNE_CREATE) Or (lParam = SHCNEConstants.SHCNE_UPDATEITEM) Then
            If LCase$(AddBackslash(pIDLToPath_Light(debugger, SHNotData.dwItem1))) = LCase$(debugger.LogFileName) Then
              ' hack - ignore AutoUpdate events that are caused by the debugger
              GoTo SkipAutoUpdate
            End If
          ElseIf lParam = SHCNEConstants.SHCNE_UPDATEDIR Then
            If LCase$(AddBackslash(pIDLToPath_Light(debugger, SHNotData.dwItem1))) = LCase$(debugger.LogFileDir) Then
              ' hack - ignore AutoUpdate events that are caused by the debugger
              GoTo SkipAutoUpdate
            End If
          End If
          debugger.AddLogEntry "Calling HandleSHNotify()", LogEntryTypeConstants.letInfo
          debugger.AddLogEntry "     Event: 0x" & Hex(lParam), LogEntryTypeConstants.letOther
          debugger.AddLogEntry "   dwItem1: 0x" & Hex(SHNotData.dwItem1), LogEntryTypeConstants.letOther
          debugger.AddLogEntry "   dwItem2: 0x" & Hex(SHNotData.dwItem2), LogEntryTypeConstants.letOther
        #End If

        HandleSHNotify MSG, wParam, lParam

        #If Debuging Then
          debugger.AddLogEntry "Returned from HandleSHNotify()", LogEntryTypeConstants.letInfo
          debugger.AddLogEntry "     Event: 0x" & Hex(lParam), LogEntryTypeConstants.letOther
          debugger.AddLogEntry "   dwItem1: 0x" & Hex(SHNotData.dwItem1), LogEntryTypeConstants.letOther
          debugger.AddLogEntry "   dwItem2: 0x" & Hex(SHNotData.dwItem2), LogEntryTypeConstants.letOther
        #End If
        hItem = SelectedItem
        #If Debuging Then
          debugger.AddLogEntry "Updating Path property after AutoUpdate", LogEntryTypeConstants.letInfo
          debugger.AddLogEntry "   SelectedItem: 0x" & Hex(hItem), LogEntryTypeConstants.letOther
        #End If
        If hItem <> -1 Then
          propPath = ItemHandleToPath(hItem)
        End If
SkipAutoUpdate:
      End If

    Case WM_SIZE
      #If BkImageSupport Then
        ' für Hintergrundbild
        If Not (propBkImage Is Nothing) Then InvalidateRectAsLong prophWnd, 0, 1
      #Else
        GetClientRect prophWnd, rcClient
      #End If

    Case WM_SYSCOLORCHANGE
      ' Systemfarben wurden geändert
      SetColors
      Refresh

    Case WM_TIMER
      Select Case wParam
        Case TimerID_DragExpand
          KillTimer prophWnd, TimerID_DragExpand
          ItemExpand prophDropTargetItem
          Exit Function
        Case TimerID_DragScroll
          GetCursorPos pt
          ScreenToClient prophWnd, pt
          With pt
            If IsWithinScrollingRegion(.x, .y) Then
              If Not (IDropTargetHelper Is Nothing) Then IDropTargetHelper.Show 0
              If .x <= propDragScrollZoneLeft Then
                SendMessageAsLong prophWnd, WM_HSCROLL, SB_LINELEFT, 0
              ElseIf .x >= (rcClient.Right - rcClient.Left) - propDragScrollZoneRight Then
                SendMessageAsLong prophWnd, WM_HSCROLL, SB_LINERIGHT, 0
              End If

              If .y <= propDragScrollZoneTop Then
                SendMessageAsLong prophWnd, WM_VSCROLL, SB_LINEUP, 0
              ElseIf .y >= (rcClient.Bottom - rcClient.Top) - propDragScrollZoneBottom Then
                SendMessageAsLong prophWnd, WM_VSCROLL, SB_LINEDOWN, 0
              End If
              If Not (IDropTargetHelper Is Nothing) Then IDropTargetHelper.Show 1
            End If
          End With
          Exit Function

        Case TimerID_MouseEvents
          If TimeOfRealSelChanged <> -1 Then
            If GetTickCount - TimeOfRealSelChanged >= propSelChangedDelayTime Then
              ' ein SelChanged-Event steht aus
              TimeOfRealSelChanged = -1
              RaiseSelChanged SelChangedDelay_hOldItem, SelChangedDelay_hNewItem, CausedByConstants.cbKeyboard, False
            End If
          End If

          If checkForMouseLeave Then
            If TimeOfPossibleMouseLeave = -1 Then
              ' prüfen, ob die aktuelle Position des Mauszeigers von seiner Position beim letzten
              ' WM_MOUSEMOVE abweicht
              ' -> das wäre ein Hinweis auf das Verlassen des Controls
              If Not propDragging_Edit Then
                GetCursorPos pt
                ScreenToClient prophWnd, pt
                If Not IsSamePoint(pt, ptLastMousePos) Then
                  boolTmp = True
                Else
                  ' möglicherweise gehört <pt> inzwischen zu einem anderen Fenster
                  If Not IsWithinVisibleClientArea(prophWnd, rcWnd, pt, False, hWndEdit, hWnd_ShellContextMenu, hWnd_EditContextMenu, hWnd_ScrollBarContextMenu, , lngTmp) Then
                    boolTmp = releasedMouseButton
                  End If
                End If

                If boolTmp And Not blockMouseLeave Then TimeOfPossibleMouseLeave = GetTickCount
              End If
            ElseIf TimeOfLastMouseMove < TimeOfPossibleMouseLeave Then
              ' seit der 1. Überprüfung wurde der Mauszeiger nicht mehr über dem Control bewegt
              ' -> sollte wieder alles auf ein Verlassen des Controls hindeuten, die Auslösung von
              ' MouseLeave einleiten
              If Not propDragging_Edit Then
                GetCursorPos pt
                ScreenToClient prophWnd, pt
                If Not IsSamePoint(pt, ptLastMousePos) Then
                  boolTmp = True
                Else
                  ' möglicherweise gehört <pt> inzwischen zu einem anderen Fenster
                  If Not IsWithinVisibleClientArea(prophWnd, rcWnd, pt, False, hWndEdit, hWnd_ShellContextMenu, hWnd_EditContextMenu, hWnd_ScrollBarContextMenu, , lngTmp) Then
                    boolTmp = releasedMouseButton
                  End If
                End If

                If boolTmp And Not blockMouseLeave Then
                  With ptLastMousePos
                    WinMainUC 1, 0, prophWnd, EXTVM_MYMOUSELEAVE, -1, MakeDWord(.x, .y)
                  End With
                End If
              End If
            Else
              ' wir haben inzwischen wieder eine WM_MOUSEMOVE empfangen -> es war wohl ein Fehlalarm
              ' Ein Fehlalarm kann entstehen, wenn der User die Maus bewegt hat, aber WM_MOUSEMOVE erst
              ' nach WM_TIMER gesendet wurde.
              TimeOfPossibleMouseLeave = -1
              releasedMouseButton = False
            End If
          End If

          If propMouseOverControl And propEnabled And Not propHot Then
            ' vor dem MouseHover-Event <propHoverTime> ablaufen lassen
            lngTmp = GetTickCount - IIf(propHoverTime = -1, GetSystemHoverTime, propHoverTime)
            If lngTmp >= TimeOfFirstMouseMove Then
              ' <propHoverTime> ist vorbei -> Auslösung des MouseHover-Events einleiten
              GetCursorPos pt
              ScreenToClient prophWnd, pt
              WinMainUC 1, 0, prophWnd, EXTVM_MYMOUSEHOVER, -1, MakeDWord(pt.x, pt.y)
            End If
          End If
          Exit Function

        Case TimerID_ResetIconCacheRebuildFlag
          bPendingShellIconCacheRebuild = False
          KillTimer prophWnd, TimerID_ResetIconCacheRebuildFlag
          Exit Function

        Case TimerID_StateIconChange
          ItemStateIconIndex(hItem_StateIconChange) = NewStateIcon

          If ItemStateIconIndex(hItem_StateIconChange) = NewStateIcon Then
            KillTimer prophWnd, TimerID_StateIconChange
            RaiseItemStateIconChanged hItem_StateIconChange, OldStateIcon, NewStateIcon
          End If
          Exit Function

        Case TimerID_SubClassEdit
          KillTimer prophWnd, TimerID_SubClassEdit

          #If UseSubClassing Then
            SubClass_Edit
          #End If
          Exit Function

        Case TimerID_SubClassParent
          If IsWindowVisible(prophWnd) Then
            KillTimer prophWnd, TimerID_SubClassParent

            If ver_Win_98OrNT4SP4 Then
              hWnd_ParentForm = GetAncestor(prophWnd, GA_ROOT)
            Else
              lngTmp = prophWnd
              Do
                hWnd_ParentForm = lngTmp
                lngTmp = GetParent(hWnd_ParentForm)
              Loop Until lngTmp = 0
            End If
            #If UseSubClassing Then
              'StartListening Me, hWnd_ParentForm
            #End If

            disableLoadFileSystem = False
            InitTreeView
            SelectPath propPath
          End If
          Exit Function
      End Select

    Case WM_UNINITMENUPOPUP
      ' ein Menü wurde zerstört
      If wParam = prophShellContextMenu Then
        prophShellContextMenu = 0
        RaiseItemDestroyShellContextMenu hItemWithContMenu
      End If
  End Select

DoDefault:
  If iListener = 1 Then
    If oldWinMain Then WinMainTree = CallWindowProc(oldWinMain, hWnd, MSG, wParam, lParam)
  End If
End Function

' wird bei eingehenden Nachrichten aufgerufen
' für das UserControl
Private Function WinMainUC(ByVal iListener As Long, ByVal oldWinMain As Long, ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Const COLOR_BTNFACE = 15
  Const COLOR_HIGHLIGHT = 13
  Const COLOR_WINDOW = 5
  Const DT_EDITCONTROL = &H2000
  Const DT_NOPREFIX = &H800
  Const DT_SINGLELINE = &H20
  Const DT_RTLREADING = &H20000
  Const DT_VCENTER = &H4
  Const TRANSPARENT = 1
  Const WM_CONTEXTMENU = &H7B
  #If BkImageSupport Then
    Const WM_CTLCOLOREDIT = &H133
    Const WM_CTLCOLORSTATIC = &H138
  #End If
  Const WM_NOTIFY = &H4E
  Const WM_SETFOCUS = &H7
  Const WM_SETTEXT = &HC
  #If Debuging Then
    Dim AsyncDrawData As NMTVASYNCDRAW
  #End If
  Dim BackClr As Long
  Dim Button As MouseButtonConstants
  Dim bytTmp As Byte
  Dim Cancel As Boolean
  Dim CDrawData As NMTVCUSTOMDRAW
  Dim changedData As Boolean
  Dim Data As NMHDR
  Dim DispInfo As NMTVDISPINFO
  Dim DrawingHeight As Long
  Dim DrawingWidth As Long
  Dim ForeClr As Long
  Dim gotItemInfo As Boolean
  Dim hasFocus As Boolean
  Dim hBrush As Long
  Dim hCaretItem As Long
  Dim hItem As Long
  Dim hLastSubItem As Long
  Dim hParentItem As Long
  Dim hPen As Long
  Dim hPenOld As Long
  Dim iconBoundingRect As RECT
  Dim IconHeight As Long
  Dim iconIndex As Long
  Dim IconWidth As Long
  #If Debuging Then
    Dim ImgLstDrawParams As IMAGELISTDRAWPARAMS
  #End If
  Dim InfoTipData As NMTVGETINFOTIP
  Dim InfoTipStyle As InfoTipStyleConstants
  Dim IParent As IVBShellFolder
  Dim isCut As Boolean
  Dim isDropHilited As Boolean
  Dim isHot As Boolean
  Dim isHilited As Boolean
  Dim isSelected As Boolean
  Dim itemAttr As SFGAOConstants
  Dim itemData As TVITEMEX
  Dim labelBoundingRect As RECT
  Dim NewName As String
  Dim newpIDL As Long
  Dim NotData As NMTREEVIEW
  Dim OldName As String
  Dim pIDLToParent As Long
  Dim pt As POINT
  Dim rc As RECT
  Dim rcGroupBox As RECT
  Dim rcItem As RECT
  Dim selectionBoundingRect As RECT
  Dim Shift As ShiftConstants
  Dim Style As ExWindowStyleConsts
  Dim textDrawStyle As Long
  Dim tmp As Long
  Dim ToolTipType As ToolTipTypeConstants
  Dim TVWItemInfo_Common As ITEMINFO_COMMON
  Dim txt As String
  Dim useStandardIcon As Boolean
  Dim x As Long
  Dim y As Long

  Select Case MSG
    Case WM_CONTEXTMENU
      ' Kontextmenü anzeigen
      TimeOfLastWM_CONTEXTMENU_Shell = GetTickCount
      If propShowShellContextMenus Then
        With pt
          .x = LoWord(lParam)
          .y = HiWord(lParam)
          If (.x = -1) And (.y = -1) Then
            ' Nachricht wurde per Tastatur erzeugt
            If propProcessContextMenuKeys Then
              ' Kontextmenü für den aktuellen Item anzeigen
              hItem = Me.SelectedItem
              GetItemRect hItem, rc

              ' Menü mittig anzeigen
              With rc
                pt.x = (.Left + .Right) / 2
                pt.y = (.Top + .Bottom) / 2
              End With
              ClientToScreen prophWnd, pt

              ItemShowShellContextMenu hItem, .x, .y
              Exit Function
            End If
          Else
            ' Nachricht wurde per Maus erzeugt
            If Not propAllowDragDrop Then
              ScreenToClient prophWnd, pt
              hItem = ItemHitTest(pt, TVHT_ONITEM)
              ClientToScreen prophWnd, pt
              ItemShowShellContextMenu hItem, .x, .y
            End If
            Exit Function
          End If
        End With
      End If

    #If BkImageSupport Then
      Case WM_CTLCOLOREDIT, WM_CTLCOLORSTATIC
        If Not (propBkImage Is Nothing) Then
          If bkImageBrush Then
            DeleteObject bkImageBrush
          End If
          If bkImageTexture Then
            DeleteObject bkImageTexture
          End If

          bkImageTexture = GetBackgroundBitmap(wParam)
          bkImageBrush = CreatePatternBrush(bkImageTexture)
          WinMainUC = bkImageBrush
          Exit Function
        End If
    #End If

    Case EXTVM_MYMOUSEHOVER
      ' der Mauszeiger befindet sich seit mindestens <propHoverTime> ms über dem Control
      Select Case hWnd
        Case prophWnd
          If Not propHot Then
            GetShiftAndMouseButtonBits wParam, Shift, Button
            With pt
              .x = LoWord(lParam)
              .y = HiWord(lParam)
              hItemUnderMouse = ItemHitTest(pt, TVHT_ONITEM)
              RaiseMouseHover curMouseButtons, Shift, .x, .y
            End With
          End If
        Case hWndEdit
          If Not propHot_Edit Then
            GetShiftAndMouseButtonBits wParam, Shift, Button
            RaiseEditMouseHover curMouseButtons_Edit, Shift, LoWord(lParam), HiWord(lParam)
          End If
      End Select

    Case EXTVM_MYMOUSELEAVE
      ' der Mauszeiger wurde aus dem Control bewegt
      Select Case hWnd
        Case prophWnd
          If propMouseOverControl Then
            GetShiftAndMouseButtonBits wParam, Shift, Button
            RaiseMouseLeave curMouseButtons, Shift, LoWord(lParam), HiWord(lParam)
          Else
            checkForMouseLeave = False
            TimeOfPossibleMouseLeave = -1
          End If
        Case hWndEdit
          If propMouseOverControl_Edit Then
            GetShiftAndMouseButtonBits wParam, Shift, Button
            RaiseEditMouseLeave curMouseButtons_Edit, Shift, LoWord(lParam), HiWord(lParam)
          Else
            checkForMouseLeave_Edit = False
            TimeOfPossibleMouseLeave_Edit = -1
          End If
      End Select

    Case WM_NOTIFY
      CopyMemory VarPtr(Data), lParam, LenB(Data)
      Select Case Data.code
        #If Debuging Then
          Case NM_STARTWAIT
            debugger.AddLogEntry "received NM_STARTWAIT", LogEntryTypeConstants.letInfo
          Case NM_ENDWAIT
            debugger.AddLogEntry "received NM_ENDWAIT", LogEntryTypeConstants.letInfo
          Case NM_BTNCLK
            debugger.AddLogEntry "received NM_BTNCLK", LogEntryTypeConstants.letInfo
        #End If
        Case NM_CLICK
          RaiseClick

        Case NM_CUSTOMDRAW
          CopyMemory VarPtr(CDrawData), lParam, LenB(CDrawData)
          With CDrawData.nmcd
            Select Case .dwDrawStage
              Case CDDS_PREPAINT
                ' wenn CustomDraw nötig ist, zurückgeben, dass wir benachrichtigt werden wollen, wenn
                ' ein Item dargestellt werden muss
                If mustCustomDraw Then
                  WinMainUC = CDRF_NOTIFYITEMDRAW
                  #If BkImageSupport Then
                    If Not (propBkImage Is Nothing) Then
                      hDC_BkImageTexture = CreateCompatibleDC(.hDC)
                      hPrevBMP_BkImageTexture = SelectObject(hDC_BkImageTexture, bkImageTexture)
                      WinMainUC = WinMainUC Or CDRF_NOTIFYPOSTPAINT
                    ElseIf propFavoritesStyle Then
                  #Else
                    If FavoritesStyle Then
                  #End If
                    WinMainUC = WinMainUC Or CDRF_NOTIFYPOSTPAINT
                  End If
                  Exit Function
                End If

              Case CDDS_POSTPAINT
                #If BkImageSupport Then
                  If Not (propBkImage Is Nothing) Then
                    SelectObject hDC_BkImageTexture, hPrevBMP_BkImageTexture
                    DeleteDC hDC_BkImageTexture
                  End If
                #End If
                If propFavoritesStyle Then
                  ' retrieve the group's start and end position
                  hCaretItem = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_CARET, 0)
                  If hCaretItem Then
                    itemData.hItem = hCaretItem
                    itemData.StateMask = TVIS_EXPANDED Or TVIS_EXPANDPARTIAL
                    itemData.Mask = TVIF_HANDLE Or TVIF_STATE
                    SendMessage prophWnd, TVM_GETITEM, 0, itemData
                    If itemData.state And (TVIS_EXPANDED Or TVIS_EXPANDPARTIAL) Then
                      hParentItem = hCaretItem
                      hItem = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_CHILD, hParentItem)
                      While hItem
                        hLastSubItem = hItem
                        hItem = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_NEXT, hItem)
                      Wend
                    Else
                      hParentItem = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_PARENT, hCaretItem)
                      If hParentItem Then
                        hItem = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_CHILD, hParentItem)
                        While hItem
                          hLastSubItem = hItem
                          hItem = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_NEXT, hItem)
                        Wend
                      Else
                        hParentItem = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_ROOT, 0)
                        hLastSubItem = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_LASTVISIBLE, 0)
                      End If
                    End If

                    ' calculate the rectangle
                    rcItem.Left = hParentItem
                    SendMessage prophWnd, TVM_GETITEMRECT, 0, rcItem
                    rcGroupBox = rcItem
                    rcItem.Left = hLastSubItem
                    SendMessage prophWnd, TVM_GETITEMRECT, 0, rcItem
                    rcGroupBox.Bottom = rcItem.Bottom

                    ' now draw the rectangle
                    hPen = CreatePen(PS_SOLID, 1, TranslateColor(SystemColorConstants.vbButtonShadow))
                    If hPen Then
                      hPenOld = SelectObject(.hDC, hPen)
                      MoveToEx .hDC, rcGroupBox.Left, rcGroupBox.Top, pt
                      LineTo .hDC, rcGroupBox.Right - 1, rcGroupBox.Top
                      LineTo .hDC, rcGroupBox.Right - 1, rcGroupBox.Bottom
                      LineTo .hDC, rcGroupBox.Left, rcGroupBox.Bottom
                      LineTo .hDC, rcGroupBox.Left, rcGroupBox.Top - 1
                      SelectObject .hDC, hPenOld
                      DeleteObject hPen
                    End If
                  End If
                End If
                WinMainUC = CDRF_DODEFAULT
                Exit Function

              Case CDDS_ITEMPREPAINT
                BackClr = CLR_INVALID
                ForeClr = CLR_INVALID
                itemAttr = -1

                hasFocus = ((.uItemState And CDIS_FOCUS) = CDIS_FOCUS)
                tmp = GetItemState_Fast(.dwItemSpec, TVIS_CUT Or TVIS_SELECTED Or TVIS_DROPHILITED)
                isCut = (tmp And TVIS_CUT)
                isSelected = (tmp And TVIS_SELECTED)
                isDropHilited = (tmp And TVIS_DROPHILITED)
                If isDropHilited Then
                  isHilited = True
                ElseIf isSelected And hasFocus Then
                  isHilited = (SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_DROPHILITE, 0) = 0)
                End If
                isHot = ((.uItemState And CDIS_HOT) = CDIS_HOT)

                #If BkImageSupport Then
                  If Not (propBkImage Is Nothing) Then
                    If Not isSelected And Not isDropHilited Then
                      BackClr = -2     ' flag for ourselves
                    ElseIf (CDrawData.clrTextBk <> GetSysColor(COLOR_HIGHLIGHT)) And (CDrawData.clrTextBk <> GetSysColor(COLOR_BTNFACE)) Then
                      BackClr = -2     ' flag for ourselves
                    End If
                  End If
                #End If

                If prophImageList Then
                  If propDimSelectedItemsIcon And (Not propCheckBoxes) And isSelected And Not isCut Then
                    If propFullRowSelect Then
                      WinMainUC = WinMainUC Or CDRF_NOTIFYPOSTPAINT
                    Else
                      itemData.hItem = .dwItemSpec
                      itemData.Mask = TVIF_HANDLE Or TVIF_SELECTEDIMAGE Or TVIF_STATE
                      SendMessage prophWnd, TVM_GETITEM, 0, itemData

                      rcItem.Left = .dwItemSpec
                      SendMessage prophWnd, TVM_GETITEMRECT, 1, rcItem
                      ImageList_GetIconSize prophImageList, VarPtr(IconWidth), VarPtr(IconHeight)
                      With rcItem
                        x = .Left - IconWidth - 3
                        If propItemHeight - IconHeight > 0 Then
                          y = .Top + Fix(((propItemHeight - IconHeight) / 2) - 0.5)
                        Else
                          y = .Top
                        End If
                      End With

                      If ver_Comctl32_610 Then
                        WinMainUC = WinMainUC Or CDRF_DOERASE Or CDRF_NOTIFYPOSTPAINT
                      Else
                        tmp = SendMessageAsLong(prophWnd, TVM_GETBKCOLOR, 0, 0)
                        hBrush = IIf(tmp = -1, GetSysColorBrush(COLOR_WINDOW), CreateSolidBrush(tmp))
                        If hBrush Then
                          ' Icon ausfüllen
                          With rc
                            .Left = x
                            .Top = y
                            .Right = .Left + IconWidth
                            .Bottom = .Top + IconHeight
                          End With
                          FillRect .hDC, rc, hBrush

                          If tmp <> -1 Then DeleteObject hBrush
                        End If
                        ImageList_DrawEx prophImageList, itemData.iSelectedImage, .hDC, x, y, IconWidth, IconHeight, CLR_HILIGHT, CLR_HILIGHT, ILD_SELECTED Or ILD_TRANSPARENT Or (itemData.state And TVIS_OVERLAYMASK)
                      End If

                      WinMainUC = WinMainUC Or TVCDRF_NOIMAGES
                    End If
                  ElseIf SendMessageAsLong(prophWnd, TVM_GETBKCOLOR, 0, 0) <> ImageList_GetBkColor(prophImageList) Then
                    itemData.hItem = .dwItemSpec
                    itemData.Mask = TVIF_HANDLE Or TVIF_IMAGE Or TVIF_SELECTEDIMAGE Or TVIF_STATE
                    SendMessage prophWnd, TVM_GETITEM, 0, itemData

                    rcItem.Left = .dwItemSpec
                    SendMessage prophWnd, TVM_GETITEMRECT, 1, rcItem
                    ImageList_GetIconSize prophImageList, VarPtr(IconWidth), VarPtr(IconHeight)
                    With rcItem
                      x = .Left - IconWidth - 3
                      If propItemHeight - IconHeight > 0 Then
                        y = .Top + Fix(((propItemHeight - IconHeight) / 2) - 0.5)
                      Else
                        y = .Top
                      End If
                    End With

                    If ver_Comctl32_610 Then
                      WinMainUC = WinMainUC Or CDRF_DOERASE Or CDRF_NOTIFYPOSTPAINT
                    Else
                      tmp = SendMessageAsLong(prophWnd, TVM_GETBKCOLOR, 0, 0)
                      hBrush = IIf(tmp = -1, GetSysColorBrush(COLOR_WINDOW), CreateSolidBrush(tmp))
                      If hBrush Then
                        ' Icon ausfüllen
                        With rc
                          .Left = x
                          .Top = y
                          .Right = .Left + IconWidth
                          .Bottom = .Top + IconHeight
                        End With
                        FillRect .hDC, rc, hBrush

                        If tmp <> -1 Then DeleteObject hBrush
                      End If
                      If isCut Then
                        ImageList_DrawEx prophImageList, IIf(isSelected, itemData.iSelectedImage, itemData.iImage), .hDC, x, y, IconWidth, IconHeight, CLR_NONE, CLR_NONE, ILD_SELECTED Or ILD_TRANSPARENT Or (itemData.state And TVIS_OVERLAYMASK)
                      ElseIf isSelected And propDimSelectedItemsIcon Then
                        ImageList_DrawEx prophImageList, IIf(isSelected, itemData.iSelectedImage, itemData.iImage), .hDC, x, y, IconWidth, IconHeight, CLR_HILIGHT, CLR_HILIGHT, ILD_SELECTED Or ILD_TRANSPARENT Or (itemData.state And TVIS_OVERLAYMASK)
                      Else
                        ImageList_DrawEx prophImageList, IIf(isSelected, itemData.iSelectedImage, itemData.iImage), .hDC, x, y, IconWidth, IconHeight, CLR_NONE, CLR_NONE, ILD_TRANSPARENT Or (itemData.state And TVIS_OVERLAYMASK)
                      End If
                    End If

                    WinMainUC = WinMainUC Or TVCDRF_NOIMAGES
                  End If
                End If

                If propSelectedItemBackColor <> -1 Then
                  If isHilited And (hDraggedItem = 0 Or hDraggedItem = .dwItemSpec) Then
                    BackClr = TranslateColor(propSelectedItemBackColor)
                  End If
                End If
                If propSelectedItemForeColor <> -1 Then
                  If isHilited And (hDraggedItem = 0 Or hDraggedItem = .dwItemSpec) Then
                    ForeClr = TranslateColor(propSelectedItemForeColor)
                  End If
                End If
                If (propHotTracking = True) And (propHotBackColor <> -1) Then
                  If isHot Then
                    ' andere HotTracking-Farbe nutzen
                    If Not isHilited Then
                      BackClr = TranslateColor(propHotBackColor)
                    End If
                  End If
                End If
                If (propHotTracking = True) And (propHotForeColor <> -1) Then
                  If isHot Then
                    ' andere HotTracking-Farbe nutzen
                    If Not isHilited Then
                      ForeClr = TranslateColor(propHotForeColor)
                    End If
                  End If
                End If

                If (propCompressedBackColor <> propBackColor) Or (propArchiveItemBackColor <> propBackColor) Then
                  gotItemInfo = True
                  If GetItemInfo_Common(.dwItemSpec, TVWItemInfo_Common) Then
                    #If Debuging Then
                      SplitFullyQualifiedPIDL debugger, TVWItemInfo_Common.pIDLToDesktop, IParent, pIDLToParent
                    #Else
                      SplitFullyQualifiedPIDL TVWItemInfo_Common.pIDLToDesktop, IParent, pIDLToParent
                    #End If
                    If (pIDLToParent <> 0) And Not (IParent Is Nothing) Then
                      #If Debuging Then
                        itemAttr = GetAttributes(debugger, IParent, pIDLToParent, SFGAOConstants.SFGAO_COMPRESSED Or SFGAOConstants.SFGAO_ENCRYPTED)
                      #Else
                        itemAttr = GetAttributes(IParent, pIDLToParent, SFGAOConstants.SFGAO_COMPRESSED Or SFGAOConstants.SFGAO_ENCRYPTED)
                      #End If
                      If itemAttr And SFGAOConstants.SFGAO_COMPRESSED Then
                        ' ein komprimierter Item
                        If Not isHilited Then
                          If propCompressedBackColor <> -1 Then
                            BackClr = TranslateColor(propCompressedBackColor)
                          End If
                        End If
                      End If
                    End If
                  Else
                    ' ein Archive-Item
                    If Not isHilited Then
                      If propArchiveItemBackColor <> -1 Then
                        BackClr = TranslateColor(propArchiveItemBackColor)
                      End If
                    End If
                  End If
                End If

                If (propCompressedForeColor <> propForeColor) Or (propArchiveItemForeColor <> propForeColor) Then
                  If Not gotItemInfo Then
                    gotItemInfo = True
                    GetItemInfo_Common .dwItemSpec, TVWItemInfo_Common
                    #If Debuging Then
                      SplitFullyQualifiedPIDL debugger, TVWItemInfo_Common.pIDLToDesktop, IParent, pIDLToParent
                    #Else
                      SplitFullyQualifiedPIDL TVWItemInfo_Common.pIDLToDesktop, IParent, pIDLToParent
                    #End If
                  End If
                  If (pIDLToParent <> 0) And Not (IParent Is Nothing) Then
                    If itemAttr = -1 Then
                      #If Debuging Then
                        itemAttr = GetAttributes(debugger, IParent, pIDLToParent, SFGAOConstants.SFGAO_COMPRESSED Or SFGAOConstants.SFGAO_ENCRYPTED)
                      #Else
                        itemAttr = GetAttributes(IParent, pIDLToParent, SFGAOConstants.SFGAO_COMPRESSED Or SFGAOConstants.SFGAO_ENCRYPTED)
                      #End If
                    End If
                    If itemAttr And SFGAOConstants.SFGAO_COMPRESSED Then
                      ' ein komprimierter Item
                      If Not isHilited Then
                        If propCompressedForeColor = -1 Then
                          ' Systemfarbe nutzen
                          If ShouldUseAlternativeClr Then
                            ForeClr = GetAlternativeClr
                          Else
                            ' Systemeinstellung ist: komprimierte Items nicht hervorheben
                          End If
                        Else
                          ForeClr = TranslateColor(propCompressedForeColor)
                        End If
                      End If
                    End If
                  Else
                    ' ein Archive-Item
                    If Not isHilited Then
                      If propArchiveItemForeColor = -1 Then
                        ' Systemfarbe nutzen
                        If ShouldUseAlternativeClr Then
                          ForeClr = GetAlternativeClr
                        Else
                          ' Systemeinstellung ist: komprimierte Items nicht hervorheben
                        End If
                      Else
                        ForeClr = TranslateColor(propArchiveItemForeColor)
                      End If
                    End If
                  End If
                End If

                If propEncryptedBackColor <> propBackColor Then
                  If Not gotItemInfo Then
                    gotItemInfo = True
                    GetItemInfo_Common .dwItemSpec, TVWItemInfo_Common
                    #If Debuging Then
                      SplitFullyQualifiedPIDL debugger, TVWItemInfo_Common.pIDLToDesktop, IParent, pIDLToParent
                    #Else
                      SplitFullyQualifiedPIDL TVWItemInfo_Common.pIDLToDesktop, IParent, pIDLToParent
                    #End If
                  End If
                  If (pIDLToParent <> 0) And Not (IParent Is Nothing) Then
                    If itemAttr = -1 Then
                      #If Debuging Then
                        itemAttr = GetAttributes(debugger, IParent, pIDLToParent, SFGAOConstants.SFGAO_COMPRESSED Or SFGAOConstants.SFGAO_ENCRYPTED)
                      #Else
                        itemAttr = GetAttributes(IParent, pIDLToParent, SFGAOConstants.SFGAO_COMPRESSED Or SFGAOConstants.SFGAO_ENCRYPTED)
                      #End If
                    End If
                    If itemAttr And SFGAOConstants.SFGAO_ENCRYPTED Then
                      ' ein verschlüsselter Item
                      If Not isHilited Then
                        If propEncryptedBackColor <> -1 Then BackClr = TranslateColor(propEncryptedBackColor)
                      End If
                    End If
                  End If
                End If

                If propEncryptedForeColor <> propForeColor Then
                  If Not gotItemInfo Then
                    gotItemInfo = True
                    GetItemInfo_Common .dwItemSpec, TVWItemInfo_Common
                    #If Debuging Then
                      SplitFullyQualifiedPIDL debugger, TVWItemInfo_Common.pIDLToDesktop, IParent, pIDLToParent
                    #Else
                      SplitFullyQualifiedPIDL TVWItemInfo_Common.pIDLToDesktop, IParent, pIDLToParent
                    #End If
                  End If
                  If (pIDLToParent <> 0) And Not (IParent Is Nothing) Then
                    If itemAttr = -1 Then
                      #If Debuging Then
                        itemAttr = GetAttributes(debugger, IParent, pIDLToParent, SFGAOConstants.SFGAO_COMPRESSED Or SFGAOConstants.SFGAO_ENCRYPTED)
                      #Else
                        itemAttr = GetAttributes(IParent, pIDLToParent, SFGAOConstants.SFGAO_COMPRESSED Or SFGAOConstants.SFGAO_ENCRYPTED)
                      #End If
                    End If
                    If itemAttr And SFGAOConstants.SFGAO_ENCRYPTED Then
                      ' ein verschlüsselter Item
                      If Not isHilited Then
                        If propEncryptedForeColor = -1 Then
                          ' Systemfarbe nutzen
                          ForeClr = GetAlternativeEncryptionClr
                        Else
                          ForeClr = TranslateColor(propEncryptedForeColor)
                        End If
                      End If
                    End If
                  End If
                End If

                ' Hintergrundfarbe übertragen
                If BackClr <> CLR_INVALID Then
                  If BackClr <> CDrawData.clrTextBk Then
                    CDrawData.clrTextBk = BackClr
                    changedData = True
                    WinMainUC = WinMainUC Or CDRF_NEWFONT
                  End If
                End If
                ' Schriftfarbe übertragen
                If ForeClr <> CLR_INVALID Then
                  If ForeClr <> CDrawData.clrText Then
                    CDrawData.clrText = ForeClr
                    changedData = True
                    WinMainUC = WinMainUC Or CDRF_NEWFONT
                  End If
                End If

                #If BkImageSupport Then
                  If Not (propBkImage Is Nothing) Then
                    If propDimSelectedItemsIcon Or isCut Or (BackClr = -2) Then
                      WinMainUC = WinMainUC Or CDRF_NOTIFYPOSTPAINT
                    End If
                  End If
                #End If

                Set IParent = Nothing
                If changedData Then
                  CopyMemory lParam, VarPtr(CDrawData), LenB(CDrawData)
                End If
                Exit Function

              Case CDDS_ITEMPOSTPAINT
                ' repaint the icon and text
                tmp = GetItemState_Fast(.dwItemSpec, TVIS_CUT Or TVIS_SELECTED Or TVIS_DROPHILITED Or TVIS_OVERLAYMASK)
                isCut = (tmp And TVIS_CUT)
                isSelected = (tmp And TVIS_SELECTED)

                #If BkImageSupport Then
                  If (ver_Comctl32_610 Or propFullRowSelect) And (CDrawData.clrTextBk = -2) Then
                    selectionBoundingRect.Left = .dwItemSpec
                    SendMessage prophWnd, TVM_GETITEMRECT, 0, selectionBoundingRect
                    BitBlt .hDC, selectionBoundingRect.Left, selectionBoundingRect.Top, selectionBoundingRect.Right - selectionBoundingRect.Left, selectionBoundingRect.Bottom - selectionBoundingRect.Top, hDC_BkImageTexture, selectionBoundingRect.Left, selectionBoundingRect.Top, SRCCOPY
                  End If
                #End If

                selectionBoundingRect.Left = .dwItemSpec
                SendMessage prophWnd, TVM_GETITEMRECT, 1, selectionBoundingRect
                labelBoundingRect = selectionBoundingRect
                labelBoundingRect.Bottom = labelBoundingRect.Top + propItemHeight
                iconBoundingRect = labelBoundingRect

                itemData.hItem = .dwItemSpec
                itemData.pszText = String(300, Chr$(0))
                itemData.cchTextMax = 299
                ' TODO: support TVIF_EXPANDEDIMAGE
                itemData.Mask = TVIF_HANDLE Or TVIF_IMAGE Or TVIF_SELECTEDIMAGE Or TVIF_TEXT
                SendMessage prophWnd, TVM_GETITEM, 0, itemData
                txt = Left$(itemData.pszText, lstrlenA(itemData.pszText))

                If prophImageList Then
                  ImageList_GetIconSize prophImageList, VarPtr(IconWidth), VarPtr(IconHeight)
                  iconBoundingRect.Right = labelBoundingRect.Left - 3
                  iconBoundingRect.Left = iconBoundingRect.Right - IconWidth

                  ' draw the icon
                  #If BkImageSupport Then
                    If CDrawData.clrTextBk = -2 Then
                      BitBlt .hDC, iconBoundingRect.Left, iconBoundingRect.Top, IconWidth, IconHeight, hDC_BkImageTexture, iconBoundingRect.Left, iconBoundingRect.Top, SRCCOPY
                    End If
                  #End If
                  If isCut Then
                    ImageList_DrawEx prophImageList, IIf(isSelected, itemData.iSelectedImage, itemData.iImage), .hDC, iconBoundingRect.Left, iconBoundingRect.Top, IconWidth, IconHeight, CLR_NONE, CLR_NONE, ILD_SELECTED Or ILD_TRANSPARENT Or (tmp And TVIS_OVERLAYMASK)
                  ElseIf isSelected And propDimSelectedItemsIcon Then
                    ImageList_DrawEx prophImageList, IIf(isSelected, itemData.iSelectedImage, itemData.iImage), .hDC, iconBoundingRect.Left, iconBoundingRect.Top, IconWidth, IconHeight, CLR_HILIGHT, CLR_HILIGHT, ILD_SELECTED Or ILD_TRANSPARENT Or (tmp And TVIS_OVERLAYMASK)
                  Else
                    ImageList_DrawEx prophImageList, IIf(isSelected, itemData.iSelectedImage, itemData.iImage), .hDC, iconBoundingRect.Left, iconBoundingRect.Top, IconWidth, IconHeight, CLR_NONE, CLR_NONE, ILD_TRANSPARENT Or (tmp And TVIS_OVERLAYMASK)
                  End If
                End If

                #If BkImageSupport Then
                  If CDrawData.clrTextBk = -2 Then
                    ' draw the text
                    BitBlt .hDC, labelBoundingRect.Left, labelBoundingRect.Top, labelBoundingRect.Right - labelBoundingRect.Left, labelBoundingRect.Bottom - labelBoundingRect.Top, hDC_BkImageTexture, labelBoundingRect.Left, labelBoundingRect.Top, SRCCOPY
                    If Len(txt) > 0 Then
                      textDrawStyle = DT_EDITCONTROL Or DT_NOPREFIX Or DT_SINGLELINE Or DT_VCENTER
                      Style = GetWindowLongPtr(prophWnd, GWL_EXSTYLE)
                      If Style And ExWindowStyleConsts.WS_EX_RTLREADING Then
                        textDrawStyle = textDrawStyle Or DT_RTLREADING
                      End If

                      InflateRect labelBoundingRect, -2, 0
                      tmp = SetBkMode(.hDC, TRANSPARENT)
                      ForeClr = SetTextColor(.hDC, CDrawData.clrText)
                      DrawText .hDC, itemData.pszText, lstrlenA(itemData.pszText), labelBoundingRect, textDrawStyle
                      SetBkMode .hDC, tmp
                      SetTextColor .hDC, ForeClr
                    End If
                  End If
                #End If
            End Select
          End With

'        Case NM_DBLCLK
'          RaiseDblClick

        Case NM_RCLICK
          ' es wurde mit der rechten Maustaste ins TreeView geklickt
          RaiseRClick

'        Case NM_RDBLCLK
'          RaiseRDblClick

        #If Debuging Then
          Case TVN_ASYNCDRAW
            CopyMemory VarPtr(AsyncDrawData), lParam, LenB(AsyncDrawData)
            With AsyncDrawData
              CopyMemory VarPtr(ImgLstDrawParams), .pimldp, LenB(ImgLstDrawParams)
              CopyMemory VarPtr(TVWItemInfo_Common), .lParam, LenB(TVWItemInfo_Common)
              debugger.AddLogEntry "received TVN_ASYNCDRAW", LogEntryTypeConstants.letInfo
              debugger.AddLogEntry "   Item 0x" & Hex(.hItem) & " (" & pIDLToDisplayName_Light(debugger, TVWItemInfo_Common.pIDLToDesktop) & ")", LogEntryTypeConstants.letOther
              debugger.AddLogEntry "   hr=" & Hex(.hr), LogEntryTypeConstants.letOther
              debugger.AddLogEntry "   dwRetFlags=" & Hex(.dwRetFlags), LogEntryTypeConstants.letOther
              debugger.AddLogEntry "   iRetImageIndex=" & .iRetImageIndex, LogEntryTypeConstants.letOther
            End With
            debugger.AddLogEntry "   ImgLstDrawParams:", LogEntryTypeConstants.letOther
            With ImgLstDrawParams
              debugger.AddLogEntry "      cbSize=" & .cbSize, LogEntryTypeConstants.letOther
              debugger.AddLogEntry "      himl=0x" & Hex(.himl), LogEntryTypeConstants.letOther
              debugger.AddLogEntry "      hDCDst=0x" & Hex(.hDCDst), LogEntryTypeConstants.letOther
              debugger.AddLogEntry "      i=" & .i, LogEntryTypeConstants.letOther
              debugger.AddLogEntry "      fState=0x" & Hex(.fState), LogEntryTypeConstants.letOther
              debugger.AddLogEntry "      fStyle=0x" & Hex(.fStyle), LogEntryTypeConstants.letOther
              debugger.AddLogEntry "      rgbBk=0x" & Hex(.rgbBk), LogEntryTypeConstants.letOther
              debugger.AddLogEntry "      rgbFg=0x" & Hex(.rgbFg), LogEntryTypeConstants.letOther
            End With
        #End If

        Case TVN_BEGINDRAG
          ' ein Item wird gedraggt
          CopyMemory VarPtr(NotData), lParam, LenB(NotData)
          GetShiftAndMouseButtonBits 0, Shift, Button, True
          dragButton = Shift Or Button
          StartDragDrop NotData.itemNew.hItem

        Case TVN_BEGINLABELEDIT
          ' ein Item wird umbenannt
          CopyMemory VarPtr(DispInfo), lParam, LenB(DispInfo)
          With DispInfo.Item
            If GetItemInfo_Common(.hItem, TVWItemInfo_Common) Then
              #If Debuging Then
                SplitFullyQualifiedPIDL debugger, TVWItemInfo_Common.pIDLToDesktop, IParent, pIDLToParent
              #Else
                SplitFullyQualifiedPIDL TVWItemInfo_Common.pIDLToDesktop, IParent, pIDLToParent
              #End If
              If (pIDLToParent <> 0) And Not (IParent Is Nothing) Then
                #If Debuging Then
                  If Not CanBeRenamed(debugger, IParent, pIDLToParent) Then
                #Else
                  If Not CanBeRenamed(IParent, pIDLToParent) Then
                #End If
                  WinMainUC = 1
                  Set IParent = Nothing
                  Exit Function
                End If
              Else
                Set IParent = Nothing
                Exit Function
              End If
            Else
              ' Archiv-Items können nicht umbenannt werden
              WinMainUC = 1
              Exit Function
            End If

            RaiseItemBeforeStartRenaming .hItem, Cancel
            If Cancel Then
              WinMainUC = 1
              Set IParent = Nothing
              Exit Function
            Else
              prophEditetItem = .hItem
              SetFocusAPI hWndEdit

              RaiseInitLabelEditControl

              ' TextBox mit dem Renaming-Namen des Items füllen
              #If Debuging Then
                txt = pIDLToDisplayName(debugger, IParent, pIDLToParent, SHGDNConstants.SHGDN_INFOLDER Or SHGDNConstants.SHGDN_FOREDITING)
              #Else
                txt = pIDLToDisplayName(IParent, pIDLToParent, SHGDNConstants.SHGDN_INFOLDER Or SHGDNConstants.SHGDN_FOREDITING)
              #End If
              txt = StrConv(txt, VbStrConv.vbFromUnicode)
              SendMessageAsLong hWndEdit, WM_SETTEXT, 0, StrPtr(txt)
              Set IParent = Nothing
            End If
          End With

        Case TVN_BEGINRDRAG
          ' ein Item wird mit der rechten Maustaste gedraggt
          CopyMemory VarPtr(NotData), lParam, LenB(NotData)
          GetShiftAndMouseButtonBits 0, Shift, Button, True
          dragButton = Shift Or Button
          StartDragDrop NotData.itemNew.hItem

        Case TVN_DELETEITEM
          CopyMemory VarPtr(NotData), lParam, LenB(NotData)

          With NotData.itemOld
            If hItemWithContMenu = .hItem Then
              hItemWithContMenu = 0
            End If
            #If Debuging Then
              #If LogItemDeletions Then
                debugger.AddLogEntry "An item is being removed.", LogEntryTypeConstants.letInfo
                debugger.AddLogEntry "   0x" & Hex(.hItem) & " (" & GetItemText(.hItem) & ")", LogEntryTypeConstants.letOther
              #End If
            #End If
            FreeItemMem .hItem
          End With

        Case TVN_ENDLABELEDIT
          ' ein Item wurde umbenannt
          If propMouseOverControl_Edit Then WinMainUC 1, 0, hWndEdit, EXTVM_MYMOUSELEAVE, 0, 0

          CopyMemory VarPtr(DispInfo), lParam, LenB(DispInfo)
          With DispInfo.Item
            RaiseDestroyLabelEditControl

            If .pszText = 0 Then
              ' solche Fälle können nur gecancelt werden - wozu also den User informieren
              prophEditetItem = 0
              WinMainUC = 1
              Exit Function
            End If

            GetItemInfo_Common .hItem, TVWItemInfo_Common
            #If Debuging Then
              SplitFullyQualifiedPIDL debugger, TVWItemInfo_Common.pIDLToDesktop, IParent, pIDLToParent
            #Else
              SplitFullyQualifiedPIDL TVWItemInfo_Common.pIDLToDesktop, IParent, pIDLToParent
            #End If
            If (pIDLToParent <> 0) And Not (IParent Is Nothing) Then
              NewName = GetStrFromPointer(.pszText)
              #If Debuging Then
                OldName = pIDLToDisplayName(debugger, IParent, pIDLToParent, SHGDNConstants.SHGDN_INFOLDER Or SHGDNConstants.SHGDN_FOREDITING)
              #Else
                OldName = pIDLToDisplayName(IParent, pIDLToParent, SHGDNConstants.SHGDN_INFOLDER Or SHGDNConstants.SHGDN_FOREDITING)
              #End If
            End If

            If NewName <> OldName Then
              RaiseItemBeforeRenaming .hItem, OldName, NewName, Cancel
            Else
              Cancel = True
            End If
            If Not Cancel Then
              newpIDL = RenamepIDL(prophWndShellUIParentWindow, IParent, pIDLToParent, NewName)
              If newpIDL = 0 Then Cancel = True

              #If Debuging Then
                If IsDrive(debugger, IParent, pIDLToParent) Then
                  ' da es ein Laufwerk ist, stimmt der DisplayName nicht (es fehlt der Laufwerksbuchstabe)
                  txt = pIDLToDisplayName(debugger, IParent, pIDLToParent, SHGDNConstants.SHGDN_INFOLDER)
              #Else
                If IsDrive(IParent, pIDLToParent) Then
                  ' da es ein Laufwerk ist, stimmt der DisplayName nicht (es fehlt der Laufwerksbuchstabe)
                  txt = pIDLToDisplayName(IParent, pIDLToParent, SHGDNConstants.SHGDN_INFOLDER)
              #End If
                txt = StrConv(txt, VbStrConv.vbFromUnicode)
                .pszText = StrPtr(txt)
                CopyMemory lParam, VarPtr(DispInfo), LenB(DispInfo)
              Else
                #If Debuging Then
                  FreeItemIDList debugger, "WinMainUC", newpIDL
                #Else
                  FreeItemIDList newpIDL
                #End If
              End If

              RaiseItemRenamed .hItem, OldName, NewName
            End If
            Set IParent = Nothing
          End With

          WinMainUC = Abs(Not Cancel)

          prophEditetItem = 0
          Exit Function

        Case TVN_GETDISPINFO
          CopyMemory VarPtr(DispInfo), lParam, LenB(DispInfo)
          With DispInfo.Item
            If Not propUseCustomIcons Then
              If (.Mask And (TVIF_IMAGE Or TVIF_SELECTEDIMAGE)) <> 0 Then
                GetItemInfo_Common .hItem, TVWItemInfo_Common
                #If Debuging Then
                  SplitFullyQualifiedPIDL debugger, TVWItemInfo_Common.pIDLToDesktop, IParent, pIDLToParent
                #Else
                  SplitFullyQualifiedPIDL TVWItemInfo_Common.pIDLToDesktop, IParent, pIDLToParent
                #End If
                If (pIDLToParent = 0) Or (IParent Is Nothing) Then
                  #If Debuging Then
                    debugger.AddLogEntry "TVN_GETDISPINFO: pIDL splitting failed (1)", LogEntryTypeConstants.letError
                  #End If
                  Exit Function
                End If
                gotItemInfo = True

                ' try IShellIcon
                hParentItem = SendMessageAsLong(prophWnd, TVM_GETNEXTITEM, TVGN_PARENT, .hItem)
                If hParentItem = 0 Then
                  hParentItem = TVI_ROOT
                End If
                If hParentItem = cachedISHIcon_hItem Then
                  ' use the current cachedISHIcon
                Else
                  ' other parent item than the last time
                  Set cachedISHIcon = Nothing
                  IParent.QueryInterface IID_IShellIcon, cachedISHIcon
                  cachedISHIcon_hItem = hParentItem
                End If

                If Not (cachedISHIcon Is Nothing) Then
                  #If Debuging Then
                    debugger.AddLogEntry "TVN_GETDISPINFO: using IShellIcon for 0x" & Hex(.hItem), LogEntryTypeConstants.letInfo
                  #End If
                  changedData = True
                  If .Mask And TVIF_IMAGE Then
                    #If Debuging Then
                      useStandardIcon = propUseStandardIcons
                      If propUseStandardIconsForSlowItems And Not useStandardIcon Then
                        useStandardIcon = IsSlowItem(debugger, IParent, pIDLToParent, TVWItemInfo_Common.pIDLToDesktop, True)
                      End If
                      .iImage = FastGetSysIconIndex(debugger, cachedISHIcon, IParent, pIDLToParent, TVWItemInfo_Common.pIDLToDesktop, False, useStandardIcon, useLargeIcons)
                    #Else
                      useStandardIcon = propUseStandardIcons
                      If propUseStandardIconsForSlowItems And Not useStandardIcon Then
                        useStandardIcon = IsSlowItem(IParent, pIDLToParent, TVWItemInfo_Common.pIDLToDesktop, True)
                      End If
                      .iImage = FastGetSysIconIndex(cachedISHIcon, IParent, pIDLToParent, TVWItemInfo_Common.pIDLToDesktop, False, useStandardIcon, useLargeIcons)
                    #End If
                  End If
                  If .Mask And TVIF_SELECTEDIMAGE Then
                    If propUseSelectedIcons Then
                      #If Debuging Then
                        useStandardIcon = propUseStandardIcons
                        If propUseStandardIconsForSlowItems And Not useStandardIcon Then
                          useStandardIcon = IsSlowItem(debugger, IParent, pIDLToParent, TVWItemInfo_Common.pIDLToDesktop, True)
                        End If
                        .iSelectedImage = FastGetSysIconIndex(debugger, cachedISHIcon, IParent, pIDLToParent, TVWItemInfo_Common.pIDLToDesktop, True, useStandardIcon, useLargeIcons)
                      #Else
                        useStandardIcon = propUseStandardIcons
                        If propUseStandardIconsForSlowItems And Not useStandardIcon Then
                          useStandardIcon = IsSlowItem(IParent, pIDLToParent, TVWItemInfo_Common.pIDLToDesktop, True)
                        End If
                        .iSelectedImage = FastGetSysIconIndex(cachedISHIcon, IParent, pIDLToParent, TVWItemInfo_Common.pIDLToDesktop, True, useStandardIcon, useLargeIcons)
                      #End If
                    Else
                      .iSelectedImage = .iImage
                    End If
                  End If
                End If

                If Not changedData Then
                  ' try SHGetFileInfo
                  changedData = True

                  If .Mask And TVIF_IMAGE Then
                    #If Debuging Then
                      useStandardIcon = propUseStandardIcons
                      If propUseStandardIconsForSlowItems And Not useStandardIcon Then
                        useStandardIcon = IsSlowItem(debugger, IParent, pIDLToParent, TVWItemInfo_Common.pIDLToDesktop, True)
                      End If
                      .iImage = GetSysIconIndex(debugger, IParent, pIDLToParent, TVWItemInfo_Common.pIDLToDesktop, False, useStandardIcon, useLargeIcons)
                    #Else
                      useStandardIcon = propUseStandardIcons
                      If propUseStandardIconsForSlowItems And Not useStandardIcon Then
                        useStandardIcon = IsSlowItem(IParent, pIDLToParent, TVWItemInfo_Common.pIDLToDesktop, True)
                      End If
                      .iImage = GetSysIconIndex(IParent, pIDLToParent, TVWItemInfo_Common.pIDLToDesktop, False, useStandardIcon, useLargeIcons)
                    #End If
                  End If
                  If .Mask And TVIF_SELECTEDIMAGE Then
                    If propUseSelectedIcons Then
                      #If Debuging Then
                        useStandardIcon = propUseStandardIcons
                        If propUseStandardIconsForSlowItems And Not useStandardIcon Then
                          useStandardIcon = IsSlowItem(debugger, IParent, pIDLToParent, TVWItemInfo_Common.pIDLToDesktop, True)
                        End If
                        .iSelectedImage = GetSysIconIndex(debugger, IParent, pIDLToParent, TVWItemInfo_Common.pIDLToDesktop, True, useStandardIcon, useLargeIcons)
                      #Else
                        useStandardIcon = propUseStandardIcons
                        If propUseStandardIconsForSlowItems And Not useStandardIcon Then
                          useStandardIcon = IsSlowItem(IParent, pIDLToParent, TVWItemInfo_Common.pIDLToDesktop, True)
                        End If
                        .iSelectedImage = GetSysIconIndex(IParent, pIDLToParent, TVWItemInfo_Common.pIDLToDesktop, True, useStandardIcon, useLargeIcons)
                      #End If
                    Else
                      .iSelectedImage = .iImage
                    End If
                  End If
                End If
              End If
            End If

            If (.Mask And TVIF_CHILDREN) <> 0 Then
              changedData = True
              If Not gotItemInfo Then
                GetItemInfo_Common .hItem, TVWItemInfo_Common
                #If Debuging Then
                  SplitFullyQualifiedPIDL debugger, TVWItemInfo_Common.pIDLToDesktop, IParent, pIDLToParent
                #Else
                  SplitFullyQualifiedPIDL TVWItemInfo_Common.pIDLToDesktop, IParent, pIDLToParent
                #End If
                If (pIDLToParent = 0) Or (IParent Is Nothing) Then
                  #If Debuging Then
                    debugger.AddLogEntry "TVN_GETDISPINFO: pIDL splitting failed (2)", LogEntryTypeConstants.letError
                  #End If
                  Exit Function
                End If
                gotItemInfo = True
              End If

              ' addFTPPath, addSubItem
              If Not propCheckSlowItemsForSubItems Then
                #If Debuging Then
                  If IsSlowItem(debugger, IParent, pIDLToParent, TVWItemInfo_Common.pIDLToDesktop) Then
                #Else
                  If IsSlowItem(IParent, pIDLToParent, TVWItemInfo_Common.pIDLToDesktop) Then
                #End If
                  .cChildren = 1
                Else
                  #If Debuging Then
                    .cChildren = Abs(HasSubItems(debugger, TVWItemInfo_Common.pIDLToDesktop, Me))
                  #Else
                    .cChildren = Abs(HasSubItems(TVWItemInfo_Common.pIDLToDesktop, Me))
                  #End If
                End If
              Else
                #If Debuging Then
                  .cChildren = Abs(HasSubItems(debugger, TVWItemInfo_Common.pIDLToDesktop, Me))
                #Else
                  .cChildren = Abs(HasSubItems(TVWItemInfo_Common.pIDLToDesktop, Me))
                #End If
              End If
            End If

            Set IParent = Nothing
            If changedData Then
              .Mask = .Mask Or TVIF_DI_SETITEM
              CopyMemory lParam, VarPtr(DispInfo), LenB(DispInfo)
              Exit Function
            End If
          End With

        Case TVN_GETINFOTIP
          ' ein ToolTip wird angefordert
          CopyMemory VarPtr(InfoTipData), lParam, LenB(InfoTipData)
          With InfoTipData
            InfoTipStyle = InfoTipStyleConstants.itsDefault
            If Not propFastInfoTipsOnly Then
              InfoTipStyle = InfoTipStyle Or InfoTipStyleConstants.itsShowSlowTip
            End If
            txt = GetToolTipText(.hItem, InfoTipStyle, ToolTipType)
            Cancel = (txt = "")

            If Not Cancel Then RaiseItemBeforeShowToolTip .hItem, ToolTipType, txt, Cancel
            Cancel = (txt = "")

            canceledToolTip = Cancel
            If Not Cancel Then
              If txt = GetItemText(.hItem) Then
                If ItemTextIsCut(.hItem) Then
                  If propShownToolTips And ShownToolTipsConstants.sttHiddenItems Then txt = ""
                End If
              End If

              If txt <> "" Then
                lstrcpyAsLong1 .pszText, txt
                .cchTextMax = lstrlenAsLong(.pszText)
                CopyMemory lParam, VarPtr(InfoTipData), LenB(InfoTipData)
              End If
            End If
          End With
          WinMainUC = 0
          Exit Function

        Case TVN_ITEMEXPANDED
          ' ein Item wurde erweitert oder geschlossen
          CopyMemory VarPtr(NotData), lParam, LenB(NotData)
          With NotData
            Select Case .Action
              Case TVE_EXPAND
                RaiseItemExpanded .itemNew.hItem
              Case TVE_COLLAPSE
                RaiseItemCollapsed .itemNew.hItem
            End Select
          End With

        Case TVN_ITEMEXPANDING
          ' ein Item wird erweitert oder geschlossen
          If Not dontExpand Then
            CopyMemory VarPtr(NotData), lParam, LenB(NotData)
            With NotData
              Select Case .Action
                Case TVE_EXPAND
                  RaiseItemExpanding .itemNew.hItem, Cancel

                  If Not Cancel Then
                    If Not ItemHasSubItems(.itemNew.hItem) Then InsertSubItems .itemNew.hItem
                    If Not ItemHasSubItems(.itemNew.hItem) Then
                      dontChangeRecyclerChildrenProp = True
                      SetItemChildrenProperty .itemNew.hItem, False
                    End If
                  End If
                Case TVE_COLLAPSE
                  RaiseItemCollapsing .itemNew.hItem, Cancel
              End Select
            End With
          Else
            Cancel = True
          End If

          If Cancel Then
            WinMainUC = 1
            Exit Function
          End If

        Case TVN_SELCHANGED
          ' ein anderer Item wurde angewählt
          CopyMemory VarPtr(NotData), lParam, LenB(NotData)
          With NotData
            If propDimSelectedItemsIcon Then
              GetItemRect .itemNew.hItem, rcItem, True
              InvalidateRectAsLong prophWnd, VarPtr(rcItem), 1
            End If
            RaiseSelChanged .itemOld.hItem, .itemNew.hItem, .Action
          End With

        Case TVN_SELCHANGING
          ' ein anderer Item wird angewählt
          CopyMemory VarPtr(NotData), lParam, LenB(NotData)
          With NotData
            RaiseSelChanging .itemOld.hItem, .itemNew.hItem, .Action, Cancel
            If Cancel Then
              WinMainUC = 1
              Exit Function
            End If
          End With

        Case TVN_SINGLEEXPAND
          ' ein Item soll per Klick auf ihn getoggelt werden
          If propSingleExpand = SingleExpandConstants.seWinXPStyle Then
            CopyMemory VarPtr(NotData), lParam, LenB(NotData)
            With NotData
              If Me.ItemGetParentItem(.itemNew.hItem) <> Me.ItemGetParentItem(.itemOld.hItem) Then
                ' we're handling 2 different branches - skip handling the old caret
                WinMainUC = TVNRET_SKIPOLD
              End If
              If Me.ItemIsExpanded(.itemNew.hItem) Then
                ' the new caret already is expanded
                WinMainUC = WinMainUC Or TVNRET_SKIPNEW
              End If
              If WinMainUC <> TVNRET_DEFAULT Then Exit Function
            End With
          End If
      End Select

    Case WM_SETFOCUS
      ' das Control hat den Fokus erhalten
      If oldWinMain Then WinMainUC = CallWindowProc(oldWinMain, hWnd, MSG, wParam, lParam)
      If hWnd = UserControl.hWnd Then
        isSettingFocus = True
        SetFocusAPI prophWnd
      End If
      Exit Function
  End Select

DoDefault:
  If iListener = 1 Then
    If oldWinMain Then WinMainUC = CallWindowProc(oldWinMain, hWnd, MSG, wParam, lParam)
  End If
End Function

Private Function WriteSettingsFile(ByVal File As String, Optional ByRef AdditionalData As Variant = Empty) As ErrorTypeConstants
  Dim arrProps() As Byte
  Dim BytesWritten As Long
  Dim hFile As Long
  Dim intBuffer As Integer
  Dim lngBuffer As Long
  Dim Props As PropertyBag
  Dim ret As Long
  Dim strBuffer As String

  On Error GoTo Fehler

  AdditionalData = Empty

  ' Datei erstellen
  hFile = CreateFileAsLong(File, GENERIC_WRITE, 0, 0, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
  If hFile = INVALID_HANDLE_VALUE Then GoTo CreateFileError

  ' Signatur speichern
  strBuffer = FileTypeSignature_SettingsFile
  ret = WriteFileAsLong(hFile, ByVal strBuffer, Len(strBuffer), BytesWritten, 0)
  If ret = 0 Then GoTo WriteFileError
  ' Formatversion speichern
  strBuffer = FileTypeVersion_SettingsFile
  ret = WriteFileAsLong(hFile, ByVal strBuffer, Len(strBuffer), BytesWritten, 0)
  If ret = 0 Then GoTo WriteFileError

  ' Signatur dieses Controls speichern - Dateierstellung
  strBuffer = FileCreatedBy_SettingsFile
  ret = WriteFileAsLong(hFile, ByVal strBuffer, Len(strBuffer), BytesWritten, 0)
  If ret = 0 Then GoTo WriteFileError
  ' Version dieses Controls speichern - Dateierstellung
  intBuffer = MakeWord(App.Minor, App.Major)
  lngBuffer = MakeDWord(App.Revision, intBuffer)
  ret = WriteFileAsLong(hFile, ByVal VarPtr(lngBuffer), LenB(lngBuffer), BytesWritten, 0)
  If ret = 0 Then GoTo WriteFileError
  ' Buildnumber speichern - Dateierstellung
  intBuffer = BuildNumber
  ret = WriteFileAsLong(hFile, ByVal VarPtr(intBuffer), LenB(intBuffer), BytesWritten, 0)
  If ret = 0 Then GoTo WriteFileError

  ' Signatur dieses Controls speichern - letzte Dateiänderung
  ret = WriteFileAsLong(hFile, ByVal strBuffer, Len(strBuffer), BytesWritten, 0)
  If ret = 0 Then GoTo WriteFileError
  ' Version dieses Controls speichern - letzte Dateiänderung
  ret = WriteFileAsLong(hFile, ByVal VarPtr(lngBuffer), LenB(lngBuffer), BytesWritten, 0)
  If ret = 0 Then GoTo WriteFileError
  ' Buildnumber speichern - letzte Dateiänderung
  ret = WriteFileAsLong(hFile, ByVal VarPtr(intBuffer), LenB(intBuffer), BytesWritten, 0)
  If ret = 0 Then GoTo WriteFileError

  ' Version von ExplorerTreeView speichern, die zum Lesen benötigt wird
  intBuffer = MakeWord(NeededVersion_SettingsFile_Read_Minor_ExTvw, NeededVersion_SettingsFile_Read_Major_ExTvw)
  lngBuffer = MakeDWord(NeededVersion_SettingsFile_Read_Revision_ExTvw, intBuffer)
  ret = WriteFileAsLong(hFile, ByVal VarPtr(lngBuffer), LenB(lngBuffer), BytesWritten, 0)
  If ret = 0 Then GoTo WriteFileError
  ' Buildnumber speichern
  intBuffer = NeededVersion_SettingsFile_Read_Build_ExTvw
  ret = WriteFileAsLong(hFile, ByVal VarPtr(intBuffer), LenB(intBuffer), BytesWritten, 0)
  If ret = 0 Then GoTo WriteFileError

  ' Version von ExplorerListView speichern, die zum Lesen benötigt wird
  intBuffer = MakeWord(NeededVersion_SettingsFile_Read_Minor_ExLvw, NeededVersion_SettingsFile_Read_Major_ExLvw)
  lngBuffer = MakeDWord(NeededVersion_SettingsFile_Read_Revision_ExLvw, intBuffer)
  ret = WriteFileAsLong(hFile, ByVal VarPtr(lngBuffer), LenB(lngBuffer), BytesWritten, 0)
  If ret = 0 Then GoTo WriteFileError
  ' Buildnumber speichern
  intBuffer = NeededVersion_SettingsFile_Read_Build_ExLvw
  ret = WriteFileAsLong(hFile, ByVal VarPtr(intBuffer), LenB(intBuffer), BytesWritten, 0)
  If ret = 0 Then GoTo WriteFileError

  ' Version von ExplorerComboBox speichern, die zum Lesen benötigt wird
  intBuffer = MakeWord(NeededVersion_SettingsFile_Read_Minor_ExCmb, NeededVersion_SettingsFile_Read_Major_ExCmb)
  lngBuffer = MakeDWord(NeededVersion_SettingsFile_Read_Revision_ExCmb, intBuffer)
  ret = WriteFileAsLong(hFile, ByVal VarPtr(lngBuffer), LenB(lngBuffer), BytesWritten, 0)
  If ret = 0 Then GoTo WriteFileError
  ' Buildnumber speichern
  intBuffer = NeededVersion_SettingsFile_Read_Build_ExCmb
  ret = WriteFileAsLong(hFile, ByVal VarPtr(intBuffer), LenB(intBuffer), BytesWritten, 0)
  If ret = 0 Then GoTo WriteFileError

  ' 12 Bytes freilassen
  SetFilePointer hFile, 12, 0, FILE_CURRENT

  ' Version von ExplorerTreeView speichern, die zum Schreiben benötigt wird
  intBuffer = MakeWord(NeededVersion_SettingsFile_Write_Minor_ExTvw, NeededVersion_SettingsFile_Write_Major_ExTvw)
  lngBuffer = MakeDWord(NeededVersion_SettingsFile_Write_Revision_ExTvw, intBuffer)
  ret = WriteFileAsLong(hFile, ByVal VarPtr(lngBuffer), LenB(lngBuffer), BytesWritten, 0)
  If ret = 0 Then GoTo WriteFileError
  ' Buildnumber speichern
  intBuffer = NeededVersion_SettingsFile_Write_Build_ExTvw
  ret = WriteFileAsLong(hFile, ByVal VarPtr(intBuffer), LenB(intBuffer), BytesWritten, 0)
  If ret = 0 Then GoTo WriteFileError

  ' Version von ExplorerListView speichern, die zum Schreiben benötigt wird
  intBuffer = MakeWord(NeededVersion_SettingsFile_Write_Minor_ExLvw, NeededVersion_SettingsFile_Write_Major_ExLvw)
  lngBuffer = MakeDWord(NeededVersion_SettingsFile_Write_Revision_ExLvw, intBuffer)
  ret = WriteFileAsLong(hFile, ByVal VarPtr(lngBuffer), LenB(lngBuffer), BytesWritten, 0)
  If ret = 0 Then GoTo WriteFileError
  ' Buildnumber speichern
  intBuffer = NeededVersion_SettingsFile_Write_Build_ExLvw
  ret = WriteFileAsLong(hFile, ByVal VarPtr(intBuffer), LenB(intBuffer), BytesWritten, 0)
  If ret = 0 Then GoTo WriteFileError

  ' Version von ExplorerComboBox speichern, die zum Schreiben benötigt wird
  intBuffer = MakeWord(NeededVersion_SettingsFile_Write_Minor_ExCmb, NeededVersion_SettingsFile_Write_Major_ExCmb)
  lngBuffer = MakeDWord(NeededVersion_SettingsFile_Write_Revision_ExCmb, intBuffer)
  ret = WriteFileAsLong(hFile, ByVal VarPtr(lngBuffer), LenB(lngBuffer), BytesWritten, 0)
  If ret = 0 Then GoTo WriteFileError
  ' Buildnumber speichern
  intBuffer = NeededVersion_SettingsFile_Write_Build_ExCmb
  ret = WriteFileAsLong(hFile, ByVal VarPtr(intBuffer), LenB(intBuffer), BytesWritten, 0)
  If ret = 0 Then GoTo WriteFileError

  ' 12 Bytes freilassen
  SetFilePointer hFile, 12, 0, FILE_CURRENT

  ' 32 Bytes freilassen
  SetFilePointer hFile, 32, 0, FILE_CURRENT

  ' Speicherplatz für die Adresse des Startbytes der Eigenschaften freilassen (4 Bytes)
  SetFilePointer hFile, 4, 0, FILE_CURRENT

  ' Länge der Eigenschaften speichern
  Set Props = New PropertyBag
  UserControl_WriteProperties Props
  arrProps = Props.Contents
  lngBuffer = ArraySize_AsByte(arrProps)
  ret = WriteFileAsLong(hFile, ByVal VarPtr(lngBuffer), LenB(lngBuffer), BytesWritten, 0)
  If ret = 0 Then GoTo WriteFileError

  ' 16 Bytes überspringen (Adressen des Startbytes und Längen der Daten für ExplorerListView und
  ' ExplorerComboBox)
  SetFilePointer hFile, 16, 0, FILE_CURRENT

  ' 16 Bytes freilassen
  lngBuffer = SetFilePointer(hFile, 16, 0, FILE_CURRENT)

  ' Eigenschaften speichern
  ret = WriteFileAsLong(hFile, ByVal VarPtr(arrProps(LBound(arrProps))), ArraySize_AsByte(arrProps), BytesWritten, 0)
  Erase arrProps
  If ret = 0 Then GoTo WriteFileError

  ' Startbyte der Eigenschaften speichern
  SetFilePointer hFile, 125, 0, FILE_BEGIN
  ret = WriteFileAsLong(hFile, ByVal VarPtr(lngBuffer), LenB(lngBuffer), BytesWritten, 0)
  If ret = 0 Then GoTo WriteFileError

  ' Datei schließen
  CloseHandle hFile

  WriteSettingsFile = ErrorTypeConstants.etNone
  Exit Function

CreateFileError:
  ' CreateFile ist fehlgeschlagen
  CloseHandle hFile
  WriteSettingsFile = ErrorTypeConstants.etCanNotOpenFile
  Exit Function

Fehler:
  ' unbekannter Fehler
  CloseHandle hFile
  AdditionalData = Err.Number
  WriteSettingsFile = ErrorTypeConstants.etUnknown
  Exit Function

WriteFileError:
  ' WriteFile ist fehlgeschlagen
  CloseHandle hFile
  WriteSettingsFile = ErrorTypeConstants.etCanNotWriteFile
End Function


' Methoden für Fokus

' wird bei eingehenden Nachrichten aufgerufen
' empfängt auch Nachrichten wie das Drücken der Pfeiltasten
Friend Function TranslateAccelerator(MsgData As MSG) As Long
  Const WM_KEYDOWN = &H100
  Const WM_KEYUP = &H101
  Const WM_SYSKEYDOWN = &H104
  Const WM_SYSKEYUP = &H105
  Dim hWndEdit As Long
  Dim Key As KeyCodeConstants
  Dim Shift As ShiftConstants

  TranslateAccelerator = S_FALSE
  If MsgData.message = WM_KEYDOWN Or MsgData.message = WM_KEYUP Or MsgData.message = WM_SYSKEYDOWN Or MsgData.message = WM_SYSKEYUP Then
    ' warum ist das eigentlich nötig??
    Key = MsgData.wParam And &HFFFF&

    Select Case Key
      Case KeyCodeConstants.vbKeyUp, KeyCodeConstants.vbKeyDown, KeyCodeConstants.vbKeyLeft, KeyCodeConstants.vbKeyRight, KeyCodeConstants.vbKeyPageDown, KeyCodeConstants.vbKeyPageUp, KeyCodeConstants.vbKeyHome, KeyCodeConstants.vbKeyEnd, KeyCodeConstants.vbKeyReturn
        hWndEdit = Me.hWndEdit
        If hWndEdit Then
          ' die Nachrichten an das Edit-Control weiterleiten
          SendMessageAsLong hWndEdit, MsgData.message, MsgData.wParam, MsgData.lParam
        Else
          SendMessageAsLong prophWnd, MsgData.message, MsgData.wParam, MsgData.lParam
        End If
        TranslateAccelerator = S_OK
    End Select
  End If
End Function


#If Debuging Then
  Private Sub TestBed()
    Dim IAugmentedSHF2 As IVBAugmentedShellFolder2
    Dim ISHF As IVBShellFolder
    Dim ISHFCommon As IVBShellFolder
    Dim ISHFUser As IVBShellFolder
    Dim pIDLCommon As Long
    Dim pIDLUser As Long

    CoCreateInstance CLSID_AugmentedShellFolder2, Nothing, CLSCTXConsts.CLSCTX_INPROC_SERVER Or CLSCTXConsts.CLSCTX_INPROC_HANDLER, IID_IAugmentedShellFolder2, IAugmentedSHF2
    If Not (IAugmentedSHF2 Is Nothing) Then
      pIDLUser = CSIDLTopIDL(CSIDLConstants.CSIDL_STARTMENU)
      Set ISHFUser = GetISHFolderInterfaceFQ(debugger, pIDLUser)
      If Not (ISHFUser Is Nothing) Then
        IAugmentedSHF2.AddNameSpace 0, ISHFUser, pIDLUser, SHCONTFConstants.SHCONTF_FOLDERS Or SHCONTFConstants.SHCONTF_NONFOLDERS
        Set ISHFUser = Nothing
      End If

      pIDLCommon = CSIDLTopIDL(CSIDLConstants.CSIDL_COMMON_STARTMENU)
      Set ISHFCommon = GetISHFolderInterfaceFQ(debugger, pIDLCommon)
      If Not (ISHFCommon Is Nothing) Then
        IAugmentedSHF2.AddNameSpace 0, ISHFCommon, pIDLCommon, SHCONTFConstants.SHCONTF_FOLDERS Or SHCONTFConstants.SHCONTF_NONFOLDERS
        Set ISHFCommon = Nothing
      End If

      FreeItemIDList debugger, "TestBed #1", pIDLUser
      FreeItemIDList debugger, "TestBed #2", pIDLCommon

      IAugmentedSHF2.QueryInterface IID_IShellFolder, ISHF
      If Not (ISHF Is Nothing) Then
        Set ISHF = Nothing
      End If
      Set IAugmentedSHF2 = Nothing
    Else
      Debug.Print "IAugmentedShellFolder2 not supported"
    End If
  End Sub
#End If
