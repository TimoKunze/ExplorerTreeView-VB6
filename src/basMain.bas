Attribute VB_Name = "basMain"
Option Explicit

  ' öffentliche Enums

  Public Enum CommonControlsConsts
    #If Debuging Then
      ICC_ANIMATE_CLASS = &H80
      ICC_BAR_CLASSES = &H4
      ICC_COOL_CLASSES = &H400
      ICC_DATE_CLASSES = &H100
      ICC_HOTKEY_CLASS = &H40
      ICC_INTERNET_CLASSES = &H800
      ICC_LINK_CLASS = &H8000
      ICC_LISTVIEW_CLASSES = &H1
      ICC_NATIVEFNTCTL_CLASS = &H2000
      ICC_PAGESCROLLER_CLASS = &H1000
      ICC_PROGRESS_CLASS = &H20
      ICC_STANDARD_CLASSES = &H4000
      ICC_TAB_CLASSES = &H8
      ICC_UPDOWN_CLASS = &H10
      ICC_USEREX_CLASSES = &H200
      ICC_WIN95_CLASSES = &HFF
    #End If
    ICC_TREEVIEW_CLASSES = &H2
  End Enum


  ' lokale Konstanten

  ' Größen von Datenstrukturen
  Private Const Size_OSVERSIONINFO = 148
  #If Debuging Then
    Private Const Size_OSVERSIONINFOEX = 156
  #End If

  ' Konstanten für GetVersionEx
  #If Debuging Then
    Private Const VER_PLATFORM_WIN32s = 0
  #End If
  Private Const VER_PLATFORM_WIN32_NT = 2
  Private Const VER_PLATFORM_WIN32_WINDOWS = 1

  ' Konstanten für GetWindowInfo
  Private Const WS_ACTIVECAPTION = &H1


  ' globale Konstanten

  Global Const BuildNumber = 881
  Global Const CLR_INVALID = -1

  ' Ressourcen
  Global Const resCreateFileError = 100
  Global Const resNoSettingsFound = 103
  Global Const resOverwriteData = 110
  Global Const resReadFileError = 102
  Global Const resSettingsAlreadyExists = 106
  Global Const resTooNewFileFormat_Read = 104
  Global Const resTooNewFileFormat_Write = 105
  Global Const resWriteFileError = 101

  ' Fehler
  Global Const ERROR_FILE_NOT_FOUND = &H2

  ' Konstanten für CheckMenuRadioItem
  Global Const MF_BYCOMMAND = &H0
  Global Const MF_BYPOSITION = &H400

  ' Konstanten für CreatePen
  #If Debuging Then
    Global Const PS_ALTERNATE = 8
    Global Const PS_COSMETIC = &H0
    Global Const PS_DASH = 1
    Global Const PS_DASHDOT = 3
    Global Const PS_DASHDOTDOT = 4
    Global Const PS_DOT = 2
    Global Const PS_ENDCAP_FLAT = &H200
    Global Const PS_ENDCAP_MASK = &HF00
    Global Const PS_ENDCAP_ROUND = &H0
    Global Const PS_ENDCAP_SQUARE = &H100
    Global Const PS_GEOMETRIC = &H10000
    Global Const PS_INSIDEFRAME = 6
    Global Const PS_JOIN_BEVEL = &H1000
    Global Const PS_JOIN_MASK = &HF000
    Global Const PS_JOIN_MITER = &H2000
    Global Const PS_JOIN_ROUND = &H0
    Global Const PS_NULL = 5
    Global Const PS_STYLE_MASK = &HF
    Global Const PS_TYPE_MASK = &HF0000
    Global Const PS_USERSTYLE = 7
  #End If
  Global Const PS_SOLID = 0

  ' Konstanten für GetSystemMetrics
  #If Debuging Then
    Global Const SM_CXSCREEN = 0
    Global Const SM_CYSCREEN = 1
    Global Const SM_CYHSCROLL = 3
    Global Const SM_CYCAPTION = 4
    Global Const SM_CXBORDER = 5
    Global Const SM_CYBORDER = 6
    Global Const SM_CXDLGFRAME = 7
    Global Const SM_CYDLGFRAME = 8
    Global Const SM_CYVTHUMB = 9
    Global Const SM_CXHTHUMB = 10
    Global Const SM_CYICON = 12
    Global Const SM_CXCURSOR = 13
    Global Const SM_CYCURSOR = 14
    Global Const SM_CYMENU = 15
    Global Const SM_CXFULLSCREEN = 16
    Global Const SM_CYFULLSCREEN = 17
    Global Const SM_CYKANJIWINDOW = 18
    Global Const SM_MOUSEPRESENT = 19
    Global Const SM_CYVSCROLL = 20
    Global Const SM_CXHSCROLL = 21
    Global Const SM_DEBUG = 22
    Global Const SM_RESERVED1 = 24
    Global Const SM_RESERVED2 = 25
    Global Const SM_RESERVED3 = 26
    Global Const SM_RESERVED4 = 27
    Global Const SM_CXMIN = 28
    Global Const SM_CYMIN = 29
    Global Const SM_CXSIZE = 30
    Global Const SM_CYSIZE = 31
    Global Const SM_CXFRAME = 32
    Global Const SM_CYFRAME = 33
    Global Const SM_CXMINTRACK = 34
    Global Const SM_CYMINTRACK = 35
    Global Const SM_CXDOUBLECLK = 36
    Global Const SM_CYDOUBLECLK = 37
    Global Const SM_CXICONSPACING = 38
    Global Const SM_CYICONSPACING = 39
    Global Const SM_MENUDROPALIGNMENT = 40
    Global Const SM_PENWINDOWS = 41
    Global Const SM_DBCSENABLED = 42
    Global Const SM_CMOUSEBUTTONS = 43
    Global Const SM_CXFIXEDFRAME = SM_CXDLGFRAME
    Global Const SM_CYFIXEDFRAME = SM_CYDLGFRAME
    Global Const SM_CXSIZEFRAME = SM_CXFRAME
    Global Const SM_CYSIZEFRAME = SM_CYFRAME
    Global Const SM_SECURE = 44
    Global Const SM_CXEDGE = 45
    Global Const SM_CYEDGE = 46
    Global Const SM_CXMINSPACING = 47
    Global Const SM_CYMINSPACING = 48
    Global Const SM_CYSMICON = 50
    Global Const SM_CYSMCAPTION = 51
    Global Const SM_CXSMSIZE = 52
    Global Const SM_CYSMSIZE = 53
    Global Const SM_CXMENUSIZE = 54
    Global Const SM_CYMENUSIZE = 55
    Global Const SM_ARRANGE = 56
    Global Const SM_CXMINIMIZED = 57
    Global Const SM_CYMINIMIZED = 58
    Global Const SM_CXMAXTRACK = 59
    Global Const SM_CYMAXTRACK = 60
    Global Const SM_CXMAXIMIZED = 61
    Global Const SM_CYMAXIMIZED = 62
    Global Const SM_NETWORK = 63
    Global Const SM_CLEANBOOT = 67
    Global Const SM_CXDRAG = 68
    Global Const SM_CYDRAG = 69
    Global Const SM_SHOWSOUNDS = 70
    Global Const SM_CXMENUCHECK = 71
    Global Const SM_CYMENUCHECK = 72
    Global Const SM_SLOWMACHINE = 73
    Global Const SM_MIDEASTENABLED = 74
    Global Const SM_MOUSEWHEELPRESENT = 75
    Global Const SM_XVIRTUALSCREEN = 76
    Global Const SM_YVIRTUALSCREEN = 77
    Global Const SM_CXVIRTUALSCREEN = 78
    Global Const SM_CYVIRTUALSCREEN = 79
    Global Const SM_CMONITORS = 80
    Global Const SM_SAMEDISPLAYFORMAT = 81
    Global Const SM_IMMENABLED = 82
    Global Const SM_CXFOCUSBORDER = 83
    Global Const SM_CYFOCUSBORDER = 84
    Global Const SM_TABLETPC = 86
    Global Const SM_MEDIACENTER = 87
    Global Const SM_REMOTESESSION = &H1000
    Global Const SM_SHUTTINGDOWN = &H2000
    Global Const SM_REMOTECONTROL = &H2001
  #End If
  Global Const SM_CXVSCROLL = 2
  Global Const SM_CXICON = 11
  Global Const SM_SWAPBUTTON = 23
  Global Const SM_CXSMICON = 49

  ' Konstanten für LoadImage
  #If Debuging Then
    Global Const IMAGE_BITMAP = 0
    Global Const IMAGE_ENHMETAFILE = 3
    Global Const IMAGE_ICON = 1
    Global Const LR_COLOR = &H2
    Global Const LR_COPYDELETEORG = &H8
    Global Const LR_COPYFROMRESOURCE = &H4000
    Global Const LR_COPYRETURNORG = &H4
    Global Const LR_CREATEDIBSECTION = &H2000
    Global Const LR_LOADFROMFILE = &H10
    Global Const LR_LOADMAP3DCOLORS = &H1000
    Global Const LR_LOADTRANSPARENT = &H20
    Global Const LR_MONOCHROME = &H1
    Global Const LR_VGACOLOR = &H80
    Global Const OBM_CLOSE = 32754
    Global Const OBM_COMBO = 32738
    Global Const OBM_DNARROW = 32752
    Global Const OBM_DNARROWD = 32742
    Global Const OBM_DNARROWI = 32736
    Global Const OBM_LFARROW = 32750
    Global Const OBM_LFARROWD = 32740
    Global Const OBM_LFARROWI = 32734
    Global Const OBM_MNARROW = 32739
    Global Const OBM_REDUCE = 32749
    Global Const OBM_REDUCED = 32746
    Global Const OBM_RESTORE = 32747
    Global Const OBM_RESTORED = 32744
    Global Const OBM_RGARROW = 32751
    Global Const OBM_RGARROWD = 32741
    Global Const OBM_RGARROWI = 32735
    Global Const OBM_UPARROW = 32753
    Global Const OBM_UPARROWD = 32743
    Global Const OBM_UPARROWI = 32737
    Global Const OBM_ZOOM = 32748
    Global Const OBM_ZOOMD = 32745
    Global Const OCR_MEDIAIN = 32663     ' nicht offiziell!
    Global Const OCR_SCROLLALL = 32654     ' nicht offiziell!
    Global Const OCR_SCROLLE = 32658     ' nicht offiziell!
    Global Const OCR_SCROLLN = 32655     ' nicht offiziell!
    Global Const OCR_SCROLLNE = 32660     ' nicht offiziell!
    Global Const OCR_SCROLLNS = 32652     ' nicht offiziell!
    Global Const OCR_SCROLLNW = 32659     ' nicht offiziell!
    Global Const OCR_SCROLLS = 32656     ' nicht offiziell!
    Global Const OCR_SCROLLSE = 32662     ' nicht offiziell!
    Global Const OCR_SCROLLSW = 32661     ' nicht offiziell!
    Global Const OCR_SCROLLW = 32657     ' nicht offiziell!
    Global Const OCR_SCROLLWE = 32653     ' nicht offiziell!
    Global Const OIC_BANG = 32515
    Global Const OIC_HAND = 32513
    Global Const OIC_NOTE = 32516
    Global Const OIC_QUES = 32514
    Global Const OIC_SAMPLE = 32512
    Global Const OIC_WINLOGO = 32517
    Global Const OIC_ERROR = OIC_HAND
    Global Const OIC_INFORMATION = OIC_NOTE
    Global Const OIC_WARNING = OIC_BANG
  #End If
  Global Const IMAGE_CURSOR = 2
  Global Const LR_DEFAULTCOLOR = &H0
  Global Const LR_DEFAULTSIZE = &H40
  Global Const LR_SHARED = &H8000
  Global Const OCR_APPSTARTING = 32650
  Global Const OCR_CROSS = 32515
  Global Const OCR_HAND = 32649
  Global Const OCR_HELP = 32651     ' nicht offiziell!
  Global Const OCR_IBEAM = 32513
  Global Const OCR_ICOCUR = 32647
  Global Const OCR_NO = 32648
  Global Const OCR_NORMAL = 32512
  Global Const OCR_SIZEALL = 32646
  Global Const OCR_SIZENESW = 32643
  Global Const OCR_SIZENS = 32645
  Global Const OCR_SIZENWSE = 32642
  Global Const OCR_SIZEWE = 32644
  Global Const OCR_UP = 32516
  Global Const OCR_WAIT = 32514

  ' Konstanten für MENUITEMINFO
  #If Debuging Then
    Global Const MFS_CHECKED = &H8
    Global Const MFS_DEFAULT = &H1000
    Global Const MFS_DISABLED = &H3
    Global Const MFS_ENABLED = &H0
    Global Const MFS_GRAYED = &H3
    Global Const MFS_HILITE = &H80
    Global Const MFS_UNCHECKED = &H0
    Global Const MFS_UNHILITE = &H0
    Global Const MFT_BITMAP = &H4
    Global Const MFT_MENUBREAK = &H40
    Global Const MFT_OWNERDRAW = &H100
    Global Const MFT_RADIOCHECK = &H200
    Global Const MFT_RIGHTJUSTIFY = &H4000
    Global Const MFT_RIGHTORDER = &H2000
  #End If
  Global Const MFT_MENUBARBREAK = &H20
  Global Const MFT_SEPARATOR = &H800
  Global Const MFT_STRING = &H0
  #If Debuging Then
    Global Const MIIM_BITMAP = &H80   ' ab Win2000
    Global Const MIIM_CHECKMARKS = &H8
    Global Const MIIM_DATA = &H20
    Global Const MIIM_STATE = &H1
    Global Const MIIM_SUBMENU = &H4
  #End If
  Global Const MIIM_FTYPE = &H100
  Global Const MIIM_ID = &H2
  Global Const MIIM_STRING = &H40
  Global Const MIIM_TYPE = &H10

  ' Konstanten für SHGetValue
  #If Debuging Then
    Global Const HKEY_CURRENT_CONFIG = &H80000005
    Global Const HKEY_DYN_DATA = &H80000006
    Global Const HKEY_LOCAL_MACHINE = &H80000002
    Global Const HKEY_PERFORMANCE_DATA = &H80000004
    Global Const HKEY_USERS = &H80000003
  #End If
  Global Const HKEY_CLASSES_ROOT = &H80000000
  Global Const HKEY_CURRENT_USER = &H80000001
  #If Debuging Then
    Global Const REG_DWORD_BIG_ENDIAN = 5
    Global Const REG_EXPAND_SZ = 2
    Global Const REG_FULL_RESOURCE_DESCRIPTOR = 9
    Global Const REG_LINK = 6
    Global Const REG_MULTI_SZ = 7
    Global Const REG_NONE = 0
    Global Const REG_QWORD = 11
    Global Const REG_QWORD_LITTLE_ENDIAN = 11
    Global Const REG_RESOURCE_LIST = 8
    Global Const REG_RESOURCE_REQUIREMENTS_LIST = 10
  #End If
  Global Const REG_BINARY = 3
  Global Const REG_DWORD = 4
  Global Const REG_DWORD_LITTLE_ENDIAN = 4
  Global Const REG_SZ = 1

  ' Konstanten für SystemParametersInfo
  #If Debuging Then
    Global Const SPI_GETBEEP = &H1
    Global Const SPI_SETBEEP = &H2
    Global Const SPI_GETMOUSE = &H3
    Global Const SPI_SETMOUSE = &H4
    Global Const SPI_GETBORDER = &H5
    Global Const SPI_SETBORDER = &H6
    Global Const SPI_GETKEYBOARDSPEED = &HA
    Global Const SPI_SETKEYBOARDSPEED = &HB
    Global Const SPI_LANGDRIVER = &HC
    Global Const SPI_ICONHORIZONTALSPACING = &HD
    Global Const SPI_GETSCREENSAVETIMEOUT = &HE
    Global Const SPI_SETSCREENSAVETIMEOUT = &HF
    Global Const SPI_GETSCREENSAVEACTIVE = &H10
    Global Const SPI_SETSCREENSAVEACTIVE = &H11
    Global Const SPI_GETGRIDGRANULARITY = &H12
    Global Const SPI_SETGRIDGRANULARITY = &H13
    Global Const SPI_SETDESKWALLPAPER = &H14
    Global Const SPI_SETDESKPATTERN = &H15
    Global Const SPI_GETKEYBOARDDELAY = &H16
    Global Const SPI_SETKEYBOARDDELAY = &H17
    Global Const SPI_ICONVERTICALSPACING = &H18
    Global Const SPI_GETICONTITLEWRAP = &H19
    Global Const SPI_SETICONTITLEWRAP = &H1A
    Global Const SPI_GETMENUDROPALIGNMENT = &H1B
    Global Const SPI_SETMENUDROPALIGNMENT = &H1C
    Global Const SPI_SETDOUBLECLKWIDTH = &H1D
    Global Const SPI_SETDOUBLECLKHEIGHT = &H1E
    Global Const SPI_GETICONTITLELOGFONT = &H1F
    Global Const SPI_SETDOUBLECLICKTIME = &H20
    Global Const SPI_SETMOUSEBUTTONSWAP = &H21
    Global Const SPI_SETICONTITLELOGFONT = &H22
    Global Const SPI_GETFASTTASKSWITCH = &H23
    Global Const SPI_SETFASTTASKSWITCH = &H24
    Global Const SPI_SETDRAGFULLWINDOWS = &H25
    Global Const SPI_GETDRAGFULLWINDOWS = &H26
    Global Const SPI_GETNONCLIENTMETRICS = &H29
    Global Const SPI_GETMINIMIZEDMETRICS = &H2B
    Global Const SPI_SETMINIMIZEDMETRICS = &H2C
    Global Const SPI_GETICONMETRICS = &H2D
    Global Const SPI_SETWORKAREA = &H2F
    Global Const SPI_GETWORKAREA = &H30
    Global Const SPI_SETPENWINDOWS = &H31
    Global Const SPI_GETHIGHCONTRAST = &H42
    Global Const SPI_SETHIGHCONTRAST = &H43
    Global Const SPI_GETKEYBOARDPREF = &H44
    Global Const SPI_SETKEYBOARDPREF = &H45
    Global Const SPI_GETSCREENREADER = &H46
    Global Const SPI_SETSCREENREADER = &H47
    Global Const SPI_GETANIMATION = &H48
    Global Const SPI_SETANIMATION = &H49
    Global Const SPI_GETFONTSMOOTHING = &H4A
    Global Const SPI_SETFONTSMOOTHING = &H4B
    Global Const SPI_SETDRAGWIDTH = &H4C
    Global Const SPI_SETDRAGHEIGHT = &H4D
    Global Const SPI_SETHANDHELD = &H4E
    Global Const SPI_GETLOWPOWERTIMEOUT = &H4F
    Global Const SPI_GETPOWEROFFTIMEOUT = &H50
    Global Const SPI_SETLOWPOWERTIMEOUT = &H51
    Global Const SPI_SETPOWEROFFTIMEOUT = &H52
    Global Const SPI_GETLOWPOWERACTIVE = &H53
    Global Const SPI_GETPOWEROFFACTIVE = &H54
    Global Const SPI_SETLOWPOWERACTIVE = &H55
    Global Const SPI_SETPOWEROFFACTIVE = &H56
    Global Const SPI_SETCURSORS = &H57
    Global Const SPI_SETICONS = &H58
    Global Const SPI_GETDEFAULTINPUTLANG = &H59
    Global Const SPI_SETDEFAULTINPUTLANG = &H5A
    Global Const SPI_SETLANGTOGGLE = &H5B
    Global Const SPI_GETWINDOWSEXTENSION = &H5C
    Global Const SPI_SETMOUSETRAILS = &H5D
    Global Const SPI_GETMOUSETRAILS = &H5E
    Global Const SPI_SETSCREENSAVERRUNNING = &H61
    Global Const SPI_SCREENSAVERRUNNING = SPI_SETSCREENSAVERRUNNING
    Global Const SPI_GETFILTERKEYS = &H32
    Global Const SPI_SETFILTERKEYS = &H33
    Global Const SPI_GETTOGGLEKEYS = &H34
    Global Const SPI_SETTOGGLEKEYS = &H35
    Global Const SPI_GETMOUSEKEYS = &H36
    Global Const SPI_SETMOUSEKEYS = &H37
    Global Const SPI_GETSHOWSOUNDS = &H38
    Global Const SPI_SETSHOWSOUNDS = &H39
    Global Const SPI_GETSTICKYKEYS = &H3A
    Global Const SPI_SETSTICKYKEYS = &H3B
    Global Const SPI_GETACCESSTIMEOUT = &H3C
    Global Const SPI_SETACCESSTIMEOUT = &H3D
    Global Const SPI_GETSERIALKEYS = &H3E
    Global Const SPI_SETSERIALKEYS = &H3F
    Global Const SPI_GETSOUNDSENTRY = &H40
    Global Const SPI_SETSOUNDSENTRY = &H41
    Global Const SPI_GETSNAPTODEFBUTTON = &H5F
    Global Const SPI_SETSNAPTODEFBUTTON = &H60
    Global Const SPI_GETMOUSEHOVERWIDTH = &H62
    Global Const SPI_SETMOUSEHOVERWIDTH = &H63
    Global Const SPI_GETMOUSEHOVERHEIGHT = &H64
    Global Const SPI_SETMOUSEHOVERHEIGHT = &H65
    Global Const SPI_SETMOUSEHOVERTIME = &H67
    Global Const SPI_GETWHEELSCROLLLINES = &H68
    Global Const SPI_SETWHEELSCROLLLINES = &H69
    Global Const SPI_GETMENUSHOWDELAY = &H6A
    Global Const SPI_SETMENUSHOWDELAY = &H6B
    Global Const SPI_GETSHOWIMEUI = &H6E
    Global Const SPI_SETSHOWIMEUI = &H6F
    Global Const SPI_GETMOUSESPEED = &H70
    Global Const SPI_SETMOUSESPEED = &H71
    Global Const SPI_GETSCREENSAVERRUNNING = &H72
    Global Const SPI_GETDESKWALLPAPER = &H73
    Global Const SPI_GETACTIVEWINDOWTRACKING = &H1000
    Global Const SPI_SETACTIVEWINDOWTRACKING = &H1001
    Global Const SPI_GETMENUANIMATION = &H1002
    Global Const SPI_SETMENUANIMATION = &H1003
    Global Const SPI_GETCOMBOBOXANIMATION = &H1004
    Global Const SPI_SETCOMBOBOXANIMATION = &H1005
    Global Const SPI_GETLISTBOXSMOOTHSCROLLING = &H1006
    Global Const SPI_SETLISTBOXSMOOTHSCROLLING = &H1007
    Global Const SPI_GETGRADIENTCAPTIONS = &H1008
    Global Const SPI_SETGRADIENTCAPTIONS = &H1009
    Global Const SPI_GETKEYBOARDCUES = &H100A
    Global Const SPI_SETKEYBOARDCUES = &H100B
    Global Const SPI_GETMENUUNDERLINES = SPI_GETKEYBOARDCUES
    Global Const SPI_SETMENUUNDERLINES = SPI_SETKEYBOARDCUES
    Global Const SPI_GETACTIVEWNDTRKZORDER = &H100C
    Global Const SPI_SETACTIVEWNDTRKZORDER = &H100D
    Global Const SPI_GETHOTTRACKING = &H100E
    Global Const SPI_SETHOTTRACKING = &H100F
    Global Const SPI_GETMENUFADE = &H1012
    Global Const SPI_SETMENUFADE = &H1013
    Global Const SPI_GETSELECTIONFADE = &H1014
    Global Const SPI_SETSELECTIONFADE = &H1015
    Global Const SPI_GETTOOLTIPANIMATION = &H1016
    Global Const SPI_SETTOOLTIPANIMATION = &H1017
    Global Const SPI_GETTOOLTIPFADE = &H1018
    Global Const SPI_SETTOOLTIPFADE = &H1019
    Global Const SPI_GETCURSORSHADOW = &H101A
    Global Const SPI_SETCURSORSHADOW = &H101B
    Global Const SPI_GETMOUSESONAR = &H101C
    Global Const SPI_SETMOUSESONAR = &H101D
    Global Const SPI_GETMOUSECLICKLOCK = &H101E
    Global Const SPI_SETMOUSECLICKLOCK = &H101F
    Global Const SPI_GETMOUSEVANISH = &H1020
    Global Const SPI_SETMOUSEVANISH = &H1021
    Global Const SPI_GETFLATMENU = &H1022
    Global Const SPI_SETFLATMENU = &H1023
    Global Const SPI_GETDROPSHADOW = &H1024
    Global Const SPI_SETDROPSHADOW = &H1025
    Global Const SPI_GETBLOCKSENDINPUTRESETS = &H1026
    Global Const SPI_SETBLOCKSENDINPUTRESETS = &H1027
    Global Const SPI_GETUIEFFECTS = &H103E
    Global Const SPI_SETUIEFFECTS = &H103F
    Global Const SPI_GETFOREGROUNDLOCKTIMEOUT = &H2000
    Global Const SPI_SETFOREGROUNDLOCKTIMEOUT = &H2001
    Global Const SPI_GETACTIVEWNDTRKTIMEOUT = &H2002
    Global Const SPI_SETACTIVEWNDTRKTIMEOUT = &H2003
    Global Const SPI_GETFOREGROUNDFLASHCOUNT = &H2004
    Global Const SPI_SETFOREGROUNDFLASHCOUNT = &H2005
    Global Const SPI_GETCARETWIDTH = &H2006
    Global Const SPI_SETCARETWIDTH = &H2007
    Global Const SPI_GETMOUSECLICKLOCKTIME = &H2008
    Global Const SPI_SETMOUSECLICKLOCKTIME = &H2009
    Global Const SPI_GETFONTSMOOTHINGTYPE = &H200A
    Global Const SPI_SETFONTSMOOTHINGTYPE = &H200B
    Global Const SPI_GETFONTSMOOTHINGCONTRAST = &H200C
    Global Const SPI_SETFONTSMOOTHINGCONTRAST = &H200D
    Global Const SPI_GETFOCUSBORDERWIDTH = &H200E
    Global Const SPI_SETFOCUSBORDERWIDTH = &H200F
    Global Const SPI_GETFOCUSBORDERHEIGHT = &H2010
    Global Const SPI_SETFOCUSBORDERHEIGHT = &H2011
    Global Const SPI_GETFONTSMOOTHINGORIENTATION = &H2012
    Global Const SPI_SETFONTSMOOTHINGORIENTATION = &H2013
    Global Const SPIF_UPDATEINIFILE = &H1
    Global Const SPIF_SENDWININICHANGE = &H2
    Global Const SPIF_SENDCHANGE = SPIF_SENDWININICHANGE
  #End If
  Global Const SPI_SETNONCLIENTMETRICS = &H2A
  Global Const SPI_SETICONMETRICS = &H2E
  Global Const SPI_GETMOUSEHOVERTIME = &H66

  ' Konstanten für TrackPopupMenuEx
  #If Debuging Then
    Global Const TPM_BOTTOMALIGN = &H20
    Global Const TPM_CENTERALIGN = &H4
    Global Const TPM_HORIZONTAL = &H0
    Global Const TPM_HORNEGANIMATION = &H800
    Global Const TPM_HORPOSANIMATION = &H400
    Global Const TPM_NOANIMATION = &H4000
    Global Const TPM_NONOTIFY = &H80
    Global Const TPM_RECURSE = &H1
    Global Const TPM_RIGHTALIGN = &H8
    Global Const TPM_TOPALIGN = &H0
    Global Const TPM_VCENTERALIGN = &H10
    Global Const TPM_VERNEGANIMATION = &H2000
    Global Const TPM_VERPOSANIMATION = &H1000
    Global Const TPM_VERTICAL = &H40
  #End If
  Global Const TPM_LEFTALIGN = &H0
  Global Const TPM_LEFTBUTTON = &H0
  Global Const TPM_RETURNCMD = &H100
  Global Const TPM_RIGHTBUTTON = &H2


  ' lokale Types

  Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
  End Type

  Private Type DLLVERSIONINFO
    cbSize As Long
    dwMajor As Long
    dwMinor As Long
    dwBuildNumber As Long
    dwPlatformId As Long
  End Type

  Private Type ICONINFO
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
  End Type

  Private Type INITCOMMONCONTROLSEXData
    dwSize As Long
    dwICC As Long
  End Type

  Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
    ' ab Windows NT4 SP6
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
  End Type

  Private Type WINDOWINFO
    cbSize As Long
    rcWindow As RECT
    rcClient As RECT
    dwStyle As Long
    dwExStyle As Long
    dwWindowStatus As Long
    cxWindowBorders As Long
    cyWindowBorders As Long
    atomWindowType As Integer
    wCreatorVersion As Integer
  End Type


  ' globale Types

  Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As Long
    cch As Long

'    hBitmap As Long     ' ab Windows 98
  End Type

  Type NMHDR
    hwndFrom As Long
    idfrom As Long
    code As Long
  End Type


  ' globale Variablen

  Global ver_Comctl32_580 As Boolean
  Global ver_Comctl32_600 As Boolean
  Global ver_Comctl32_610 As Boolean
  Global ver_Shell32_50 As Boolean
  Global ver_Shlwapi_50 As Boolean
  Global ver_Win_98Or2k As Boolean
  Global ver_Win_98OrNT4SP3 As Boolean
  Global ver_Win_98OrNT4SP4 As Boolean
  Global ver_Win_2k As Boolean
  Global ver_Win_MEOrXP As Boolean
  Global ver_Win_NTBased As Boolean
  Global ver_Win_XP As Boolean
  Global ver_Win_Vista As Boolean
  Global ver_Win_Seven As Boolean


  ' lokale APIs

  Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDC_Dst As Long, ByVal xDst As Long, ByVal yDst As Long, ByVal cxDst As Long, ByVal cyDst As Long, ByVal hDC_Src As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal Mode As Long) As Long
  Private Declare Function CreateBitmapAsLong Lib "gdi32.dll" Alias "CreateBitmap" (ByVal cx As Long, ByVal cy As Long, ByVal planes As Long, ByVal bitCount As Long, ByVal pBits As Long) As Long
  Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hDC As Long, ByVal Width As Long, ByVal Height As Long) As Long
  Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
  Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
  Private Declare Function DllGetVersion_comctl32 Lib "comctl32.dll" Alias "DllGetVersion" (Data As DLLVERSIONINFO) As Long
  Private Declare Function DllGetVersion_shell32 Lib "shell32.dll" Alias "DllGetVersion" (Data As DLLVERSIONINFO) As Long
  Private Declare Function DllGetVersion_shlwapi Lib "shlwapi.dll" Alias "DllGetVersion" (Data As DLLVERSIONINFO) As Long
  Private Declare Function DrawIconEx Lib "user32.dll" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
  Private Declare Function GetActiveWindow Lib "user32.dll" () As Long
  Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal VirtKey As Long) As Integer
  Private Declare Function GetBitmapBits Lib "gdi32.dll" (ByVal hBMP As Long, ByVal cbBuffer As Long, lpvBits As Any) As Long
  Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
  Private Declare Function GetIconInfo Lib "user32.dll" (ByVal hIcon As Long, icoInfo As ICONINFO) As Long
  Private Declare Function GetObjectAPI Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal bufferSize As Long, Data As Any) As Long
  Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" (Data As Any) As Long
  Private Declare Function GetWindowInfo Lib "user32.dll" (ByVal hWnd As Long, ByRef Data As WINDOWINFO) As Long
  Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (Data As INITCOMMONCONTROLSEXData) As Long
  Private Declare Function ImageList_Add Lib "comctl32.dll" (ByVal himl As Long, ByVal hbmImage As Long, ByVal hbmMask As Long) As Long
  Private Declare Function ImageList_GetIconSize Lib "comctl32.dll" (ByVal himl As Long, ByVal cx As Long, ByVal cy As Long) As Long
  Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal Clr As Long, ByVal hPal As Long, ret As Long) As Long
  Private Declare Function PathAddBackslash Lib "shlwapi.dll" Alias "PathAddBackslashA" (ByVal path As String) As Long
  Private Declare Function PathCommonPrefix Lib "shlwapi.dll" Alias "PathCommonPrefixA" (ByVal Path1 As String, ByVal Path2 As String, ByVal CommonPrefix As String) As Long
  Private Declare Function PathRemoveBackslash Lib "shlwapi.dll" Alias "PathRemoveBackslashA" (ByVal path As String) As Long
  Private Declare Function PathRemoveFileSpec Lib "shlwapi.dll" Alias "PathRemoveFileSpecA" (ByVal path As String) As Long
  Private Declare Sub PathStripPath Lib "shlwapi.dll" Alias "PathStripPathA" (ByVal path As String)
  Private Declare Function PtInRect Lib "user32.dll" (rc As RECT, ByVal x As Long, ByVal y As Long) As Long
  Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long
  Private Declare Function RemoveMenu Lib "user32.dll" (ByVal hMenu As Long, ByVal pos As Long, ByVal Flags As Long) As Long
  Private Declare Function ShowCursor Lib "user32.dll" (ByVal fShow As Long) As Long
  Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal x As Long, ByVal y As Long) As Long


  ' globale APIs

  Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, pt As POINT) As Long
  Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal Size As Long)
  Declare Function CreatePen Lib "gdi32" (ByVal PenStyle As Long, ByVal Width As Long, ByVal Clr As Long) As Long
  Declare Function CreatePopupMenu Lib "user32" () As Long
  Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
  Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
  Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
  Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, rc As RECT) As Long
  Declare Function GetCursorPos Lib "user32" (pt As POINT) As Long
  Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
  Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal ItemID As Long, ByVal isPosition As Long, Data As MENUITEMINFO) As Long
  Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal ProcName As String) As Long
  Declare Function GetSystemMetrics Lib "user32" (ByVal Index As Long) As Long
  Declare Function IsMenu Lib "user32" (ByVal hMenu As Long) As Long
  Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
  Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hinst As Long, ByVal Img As Long, ByVal ImgType As Long, ByVal DesiredWidth As Long, ByVal DesiredHeight As Long, ByVal Flags As Long) As Long
  Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpFileName As String) As Long
  Declare Function lstrcpyAsLong2 Lib "kernel32.dll" Alias "lstrcpyA" (ByVal Dest As String, ByVal pSrc As Long) As Long
  Declare Function lstrlenA Lib "kernel32.dll" (ByVal str As String) As Long
  Declare Function lstrlenAsLong Lib "kernel32.dll" Alias "lstrlenA" (ByVal pStr As Long) As Long
  Declare Function lstrlenW Lib "kernel32.dll" (ByVal pStr As Long) As Long
  Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, oldPoint As POINT) As Long
  Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, pt As POINT) As Long
  Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
  Declare Function SendMessageAsLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
  Declare Function SHGetValue Lib "shlwapi" Alias "SHGetValueA" (ByVal hKey As Long, ByVal SubKey As String, ByVal Value As String, ByVal DataType As Long, Data As Any, ByVal Datasize As Long) As Long
  Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal ParamInfo As Long, ByVal ParamData1 As Long, ByVal ParamData2 As Any, ByVal Flags As Long) As Long
  Declare Function TrackPopupMenuExAsLong Lib "user32" Alias "TrackPopupMenuEx" (ByVal hMenu As Long, ByVal Flags As Long, ByVal x As Long, ByVal y As Long, ByVal hWnd As Long, ByVal Data As Long) As Long


' globale Methoden

' hängt einen Backslash an <txt> an (wenn noch keiner dran ist)
' vorher wird geprüft, ob <txt> ein Verzeichnis ist
Function AddBackslash(ByVal txt As String, Optional ByVal OnlyIfDirectory As Boolean = True, Optional ByVal Backslash As String = "\") As String
  If OnlyIfDirectory Then
    If Not IsDirectory(txt) Then
      If Not (Backslash = "/" And IsFTPURL(txt)) Then
        AddBackslash = Trim$(Left$(txt, lstrlenA(txt)))
        Exit Function
      End If
    End If
  End If
  If Backslash = "\" Then
    txt = txt & String$(2, Chr$(0))
    PathAddBackslash txt
    txt = Trim$(Left$(txt, lstrlenA(txt)))
  Else
    If Right$(txt, Len(Backslash)) <> Backslash Then txt = txt & Backslash
  End If

  AddBackslash = txt
End Function
'
'Function AddCursorToImageList(ByVal himl As Long, ByVal hCursor As Long, ByVal idMerge As Long, ptHotSpot As POINT) As Long
'  Dim hBMPImage As Long
'  Dim hBMPMask As Long
'
'  If MergeIcons(hCursor, -1, hBMPImage, hBMPMask, ptHotSpot) Then
'    AddCursorToImageList = ImageList_Add(himl, hBMPImage, hBMPMask)
'    If hBMPImage Then
'      DeleteObject hBMPImage
'    End If
'    If hBMPMask Then
'      DeleteObject hBMPMask
'    End If
'  Else
'    AddCursorToImageList = -1
'  End If
'End Function

' ermittelt die Größe des Arrays <Arr> in Bytes (für ByteArrays)
Function ArraySize_AsByte(arr() As Byte) As Long
  Dim ret As Long

  If Not IsArray(arr) Then Exit Function

  ret = LenB(arr(LBound(arr))) * countElements_AsByte(arr())
  ArraySize_AsByte = ret
End Function

' ermittelt die Anzahl der Elemente des Arrays <Arr> (für ByteArrays)
Function countElements_AsByte(arr() As Byte) As Long
  Dim ret As Long

  If Not IsArray(arr) Then Exit Function

  ret = UBound(arr) - LBound(arr)
  countElements_AsByte = ret + 1
End Function

' gibt die Anzahl der Segmente zurück, aus der <Path> besteht
Function CountSegments(ByVal path As String, Optional ByVal Sep As String = "\") As Integer
  If IsUNC(path) Then path = Mid$(path, 3)
  path = AddBackslash(path, False, Sep)
  If LCase$(Left$(path, Len("ftp://"))) = "ftp://" Then path = Mid$(path, 7)
  CountSegments = countStrings(path, Sep)
End Function

' gibt zurück, wie oft <Zeichen> in <txt> im Intervall [<Start>; <Ende>] vorkommt
Function countStrings(ByVal txt As String, ByVal char As String, Optional ByVal Start As Integer = -1, Optional ByVal Ende As Integer = -1) As Integer
  Dim c As Long
  Dim pos As Integer

  If Start = -1 Then Start = 1
  If Ende = -1 Then Ende = Len(txt)

  txt = Mid$(txt, Start, Ende - Start + 1)
  pos = 0
  Do
    pos = InStrB(pos + 1, txt, char)
    If pos Then c = c + 1
  Loop While pos

  countStrings = c
End Function
'
'Function CreateColorBitmap(ByVal cx As Long, ByVal cy As Long) As Long
'  Dim hDC As Long
'
'  hDC = GetDC(0)
'  CreateColorBitmap = CreateCompatibleBitmap(hDC, cx, cy)
'  ReleaseDC 0, hDC
'End Function
'
'Sub DAD_ShowCursor(ByVal fShow As Boolean)
'  Static s_fCursorHidden As Boolean
'
'  If fShow Then
'    If s_fCursorHidden Then
'      ShowCursor 1
'      s_fCursorHidden = False
'    End If
'  Else
'    If Not s_fCursorHidden Then
'      ShowCursor 0
'      s_fCursorHidden = True
'    End If
'  End If
'End Sub

'
'' malt in <hDC> einen Rand <rc> in der Farbe <Clr>
'Sub drawBorder(ByVal hDC As Long, rc As RECT, ByVal Clr As Long, Optional ByVal Thickness As Long = 1)
'  Dim hPen As Long
'  Dim hPen_Old As Long
'  Dim pt As POINT
'
'  hPen = CreatePen(PS_SOLID, Thickness, Clr)
'  If hPen Then
'    hPen_Old = SelectObject(hDC, hPen)
'
'    With rc
'      MoveToEx hDC, .Left, .Top, pt
'      LineTo hDC, .Right - 1, .Top
'      LineTo hDC, .Right - 1, .Bottom - 1
'      LineTo hDC, .Left, .Bottom - 1
'      LineTo hDC, .Left, .Top
'    End With
'
'    SelectObject hDC, hPen_Old
'    DeleteObject hPen
'  End If
'End Sub

' gibt die Ziffern nach dem Komma zurück
Function Frac(ByVal Num As Variant) As Variant
  If Not IsNumeric(Num) Then Exit Function

  Frac = Num - Int(Num)
End Function

' gibt einen Pointer auf <Func> zurück (Workaround, da "a = AddressOf b" nicht geht)
Function FuncPtr(ByVal Func As Long) As Long
  FuncPtr = Func
End Function

' gibt die Systemfarbe für komprimierte Items zurück
Function GetAlternativeClr() As Long
  Dim Data As Long
  Dim Datasize As Long
  Dim ret As Long

  Datasize = LenB(Data)
  ret = SHGetValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer", "AltColor", VarPtr(REG_BINARY), ByVal VarPtr(Data), VarPtr(Datasize))
  ' blau ist doch hoffentlich überall der Standardwert
  If ret = ERROR_FILE_NOT_FOUND Then Data = ColorConstants.vbBlue

  GetAlternativeClr = Data
End Function

' gibt die Systemfarbe für verschlüsselte Items zurück
Function GetAlternativeEncryptionClr() As Long
  Dim Data As Long
  Dim Datasize As Long
  Dim ret As Long

  Datasize = LenB(Data)
  ret = SHGetValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer", "AltEncryptionColor", VarPtr(REG_BINARY), ByVal VarPtr(Data), VarPtr(Datasize))
  ' grün ist doch hoffentlich überall der Standardwert
  If ret = ERROR_FILE_NOT_FOUND Then Data = RGB(19, 146, 13)

  GetAlternativeEncryptionClr = Data
End Function

' gibt den gemeinsamen Anfang von <Path1> und <Path2> zurück
Function GetCommonPrefix(ByVal Path1 As String, ByVal Path2 As String, Optional ByVal OnlyIfDirectorys As Boolean = True) As String
  Dim i As Integer
  Dim ret As String
  Dim txt As String

  If OnlyIfDirectorys Then
    ' wenn es keine Pfade sind, erledigt sich die Sache von selbst
    ret = String$(MAX_PATH, Chr$(0))
    PathCommonPrefix Path1, Path2, ret
  Else
    If IsDirectory(Path1) And IsDirectory(Path2) Then
      ' da beides Pfade sind, können wir das API nehmen
      ret = String$(MAX_PATH, Chr$(0))
      PathCommonPrefix Path1, Path2, ret
    Else
      ' jetzt müssen wir "per Hand" suchen
      For i = 1 To CountSegments(GetLongestString(Path1, Path2))
        txt = GetFirstFolders(Path1, i)
        If LCase$(txt) <> LCase$(GetFirstFolders(Path2, i)) Then
          Exit For
        Else
          ret = txt
        End If
      Next
    End If
  End If
  GetCommonPrefix = Left$(ret, lstrlenA(ret))
End Function
'
'Sub GetCursorLowerRight(ByVal hCursor As Long, x As Long, y As Long, ptHotSpot As POINT)
'  Dim bmp As BITMAP
'  Dim curMask(0 To 127) As Integer
'  Dim dwMask As Long
'  Dim icoInfo As ICONINFO
'  Dim i As Long
'  Dim j As Long
'  Dim xFine As Long
'
'  xFine = 16
'
'  GetIconInfo hCursor, icoInfo
'  GetObjectAPI icoInfo.hbmMask, LenB(bmp), bmp
'  GetBitmapBits icoInfo.hbmMask, 256, ByVal VarPtr(curMask(0))
'  ptHotSpot.x = icoInfo.xHotspot
'  ptHotSpot.y = icoInfo.yHotspot
'  If icoInfo.hbmColor Then
'    i = bmp.bmWidth * bmp.bmHeight / 16 - 1
'  Else
'    i = bmp.bmWidth * (bmp.bmHeight / 2) / 16 - 1
'  End If
'
'  If i >= 256 Then
'    i = 255
'  End If
'
'  For i = i To 0 Step -1
'    If curMask(i) <> &HFFFF Then
'      ' this is only accurate to 16 pixels... which is a big gap..
'      ' so let's try to be a bit more accurate.
'      For j = 0 To 15
'        If j < 8 Then
'          dwMask = 2 ^ (8 + j)
'        Else
'          dwMask = 2 ^ (j - 8)
'        End If
'
'        If (curMask(i) And dwMask) = 0 Then
'          Exit For
'        End If
'        xFine = xFine - 1
'      Next j
'      Exit For
'    End If
'  Next i
'
'  If icoInfo.hbmColor Then
'    DeleteObject icoInfo.hbmColor
'  End If
'  If icoInfo.hbmMask Then
'    DeleteObject icoInfo.hbmMask
'  End If
'
'  ' Compute the pointer height
'  ' use width in both directions because the cursor is square, but the
'  ' height might be doubleheight if it's mono
'  y = ((i + 1) * 16) / bmp.bmWidth
'  x = ((i * 16) Mod bmp.bmWidth) + xFine + 2   ' hang it off a little
'End Sub

' gibt den Dateinamen in einem Pfad zurück
Function GetFileName(ByVal path As String) As String
  PathStripPath path
  GetFileName = Left$(path, lstrlenA(path))
End Function

' gibt von <Path> die ersten <c> Segmente zurück
Function GetFirstFolders(ByVal path As String, ByVal c As Integer, Optional ByVal Sep As String = "\") As String
  Dim Count As Integer
  Dim i As Integer
  Dim pos As Integer

  Count = CountSegments(path, Sep)
  If c > Count Then c = Count

  On Error Resume Next
  If IsUNC(path) Then pos = 2
  path = AddBackslash(path, False, Sep)
  If Left$(path, Len("ftp://")) = "ftp://" Then c = c + 2
  For i = 1 To c
    pos = InStr(pos + 1, path, Sep)
    If pos = 0 Then Exit For
  Next
  GetFirstFolders = Left$(path, pos - 1)
End Function

Sub GetImgLstIconSize(ByVal hImgLst As Long, ByRef IconWidth As Long, ByRef IconHeight As Long)
  If hImgLst = 0 Then Exit Sub

  ImageList_GetIconSize hImgLst, VarPtr(IconWidth), VarPtr(IconHeight)
End Sub

' gibt von <Path> die letzten <c> Segmente zurück
Function GetLastFolders(ByVal path As String, ByVal c As Integer, Optional ByVal Sep As String = "\") As String
  Dim i As Integer
  Dim pos As Integer

  path = RemoveBackslash(path, False, Sep)
  For i = 1 To c
    pos = InStrRev(path, Sep, pos - 1)
    If pos = 0 Then Exit For
  Next
  If IsFTPURL(path) Then
    If pos <= Len("ftp://") Then pos = 0
  End If

  GetLastFolders = Mid$(path, pos + 1)
End Function

' gibt den längeren der beiden Strings zurück
Function GetLongestString(ByVal str1 As String, ByVal str2 As String) As String
  If Len(str1) >= Len(str2) Then
    GetLongestString = str1
  Else
    GetLongestString = str2
  End If
End Function

Function GetMenuItemText(ByVal hMenu As Long, ByVal ItemID As Long, Optional ByVal isPosition As Boolean = False) As String
  Dim buffer() As Byte
  Dim bufferSize As Long
  Dim Data As MENUITEMINFO
  Dim ret As String

  If IsMenu(hMenu) Then
    If Not IsSeparator(hMenu, ItemID, isPosition) Then
      With Data
        .cbSize = LenB(Data)
        .cch = 0
        If ver_Win_98Or2k Then
          .fMask = MIIM_STRING
        Else
          .fMask = MIIM_TYPE
        End If
        GetMenuItemInfo hMenu, ItemID, Abs(isPosition), Data
        If .cch > 0 Then
          bufferSize = .cch
          .cch = .cch + 1
          ReDim buffer(0 To .cch) As Byte
          .dwTypeData = VarPtr(buffer(0))
          GetMenuItemInfo hMenu, ItemID, Abs(isPosition), Data

          ret = Left$(StrConv(buffer, VbStrConv.vbUnicode), bufferSize)
        End If
      End With
    End If
  End If

  GetMenuItemText = ret
End Function

' gibt den Pfad in einem Dateipfad zurück
Function GetPathName(ByVal path As String) As String
  PathRemoveFileSpec path
  GetPathName = Left$(path, lstrlenA(path))
End Function

' gibt den Standardwert eines RegSchlüssels zurück
Function getRegDefaultValue(ByVal hRootKey As Long, ByVal SubKey As String) As String
  Dim Data As String

  Data = Space(MAX_PATH)
  SHGetValue hRootKey, SubKey, "", VarPtr(REG_SZ), ByVal Data, VarPtr(Len(Data))
  Data = Trim(removeNull(Data))

  getRegDefaultValue = Data
End Function

' ermittelt aus <wParam> die Bits für die Shift-Tasten und die Maustasten
' zusätzlich wird der Status der Alt-Taste abgefragt
' für <useAsynchKeyState> = True werden die Werte ermittelt
Sub GetShiftAndMouseButtonBits(ByVal wParam As Long, Shift As ShiftConstants, MButton As MouseButtonConstants, Optional ByVal useAsynchKeyState As Boolean = False)
  Const MK_ALT = &H20
  Const MK_CONTROL = &H8
  Const MK_LBUTTON = &H1
  Const MK_MBUTTON = &H10
  Const MK_RBUTTON = &H2
  Const MK_SHIFT = &H4
  Const VK_CONTROL = &H11
  Const VK_LBUTTON = MK_LBUTTON
  Const VK_MBUTTON = &H4
  Const VK_MENU = &H12
  Const VK_RBUTTON = MK_RBUTTON
  Const VK_SHIFT = &H10

  If wParam = -1 Then useAsynchKeyState = True

  If useAsynchKeyState Then
    If IsKeyHoldDown(VK_CONTROL) Then Shift = Shift Or ShiftConstants.vbCtrlMask
    If IsKeyHoldDown(VK_LBUTTON) Then MButton = MButton Or MouseButtonConstants.vbLeftButton
    If IsKeyHoldDown(VK_MBUTTON) Then MButton = MButton Or MouseButtonConstants.vbMiddleButton
    If IsKeyHoldDown(VK_MENU) Then Shift = Shift Or ShiftConstants.vbAltMask
    If IsKeyHoldDown(VK_RBUTTON) Then MButton = MButton Or MouseButtonConstants.vbRightButton
    If IsKeyHoldDown(VK_SHIFT) Then Shift = Shift Or ShiftConstants.vbShiftMask
  Else
    If wParam And MK_ALT Then Shift = Shift Or ShiftConstants.vbAltMask
    If wParam And MK_CONTROL Then Shift = Shift Or ShiftConstants.vbCtrlMask
    If wParam And MK_LBUTTON Then MButton = MButton Or MouseButtonConstants.vbLeftButton
    If wParam And MK_MBUTTON Then MButton = MButton Or MouseButtonConstants.vbMiddleButton
    If wParam And MK_RBUTTON Then MButton = MButton Or MouseButtonConstants.vbRightButton
    If wParam And MK_SHIFT Then Shift = Shift Or ShiftConstants.vbShiftMask
  End If
End Sub

' gibt den String an der Adresse <Ptr> zurück
Function GetStrFromPointer(ByVal Ptr As Long, Optional ByVal remNull As Boolean = True) As String
  Dim Length As Long
  Dim ret As String

  Length = lstrlenAsLong(Ptr)
  If Length Then
    ret = String$(Length, Chr$(0))
    lstrcpyAsLong2 ret, Ptr

    If remNull Then ret = Left$(ret, lstrlenA(ret))
  End If

  GetStrFromPointer = ret
End Function

Function GetSystemHoverTime() As Long
  Dim ret As Long

  If SystemParametersInfo(SPI_GETMOUSEHOVERTIME, 0, VarPtr(ret), 0) = 0 Then ret = 1

  GetSystemHoverTime = ret
End Function

' ermittelt die Versionen von Windows und diversen Dateien
Sub GetVersions()
  Dim DLLVerData As DLLVERSIONINFO
  Dim hModule As Long
  Dim OSVerData As OSVERSIONINFOEX
  Dim txt As String

  With DLLVerData
    hModule = LoadLibrary("comctl32.dll")
    If hModule Then
      If GetProcAddress(hModule, "DllGetVersion") Then
        With DLLVerData
          .cbSize = LenB(DLLVerData)
          DllGetVersion_comctl32 DLLVerData
          If .dwMajor = 5 Then
            ver_Comctl32_580 = (.dwMinor >= 80)
          ElseIf .dwMajor >= 6 Then
            ver_Comctl32_580 = True
            ver_Comctl32_600 = True
            ver_Comctl32_610 = (.dwMinor >= 10)
          End If
        End With
      Else
        ' eine ältere Version als 4.71
      End If
      FreeLibrary hModule
    End If

    hModule = LoadLibrary("shell32.dll")
    If hModule Then
      If GetProcAddress(hModule, "DllGetVersion") Then
        With DLLVerData
          .cbSize = LenB(DLLVerData)
          DllGetVersion_shell32 DLLVerData
          ver_Shell32_50 = (.dwMajor >= 5)
        End With
      Else
        ' eine ältere Version als 4.71
      End If
      FreeLibrary hModule
    End If

    hModule = LoadLibrary("shlwapi.dll")
    If hModule Then
      If GetProcAddress(hModule, "DllGetVersion") Then
        With DLLVerData
          .cbSize = LenB(DLLVerData)
          DllGetVersion_shlwapi DLLVerData
          ver_Shlwapi_50 = (.dwMajor >= 5)
        End With
      Else
        ' eine ältere Version als 4.71
      End If
      FreeLibrary hModule
    Else
      MsgBox "shlwapi.dll not found!", vbCritical
    End If
  End With

  With OSVerData
    .dwOSVersionInfoSize = Size_OSVERSIONINFO
    GetVersionEx OSVerData

    Select Case .dwPlatformId
      Case VER_PLATFORM_WIN32_NT
        ver_Win_NTBased = True
        If .dwMajorVersion = 4 Then
          If .dwMinorVersion = 0 Then
            ' auf SP3 bzw. SP4 oder höher prüfen
            txt = removeNull(.szCSDVersion)
            If txt <> "" Then
              txt = Right(txt, 5)
              While Not isValidNumber(txt, 1, 99999)
                txt = Mid(txt, 2)
              Wend
              ver_Win_98OrNT4SP3 = (CInt(txt) >= 3)
              ver_Win_98OrNT4SP4 = (CInt(txt) >= 4)
            End If
          Else
            ver_Win_98OrNT4SP3 = True
            ver_Win_98OrNT4SP4 = True
          End If
        ElseIf .dwMajorVersion = 5 Then
          ver_Win_98OrNT4SP3 = (.dwMajorVersion > 4)
          ver_Win_98OrNT4SP4 = (.dwMajorVersion > 4)
          ver_Win_MEOrXP = (.dwMinorVersion >= 1)
          ver_Win_2k = (.dwMinorVersion >= 0)
          ver_Win_98Or2k = ver_Win_2k
          ver_Win_XP = ver_Win_MEOrXP
        Else
          ver_Win_98OrNT4SP3 = (.dwMajorVersion > 4)
          ver_Win_98OrNT4SP4 = (.dwMajorVersion > 4)
          ver_Win_2k = (.dwMajorVersion >= 5)
          ver_Win_98Or2k = ver_Win_2k
          ver_Win_MEOrXP = (.dwMajorVersion > 5)
          ver_Win_XP = ver_Win_MEOrXP
          ver_Win_Vista = (.dwMajorVersion >= 6)
          If .dwMajorVersion = 6 Then
            ver_Win_Seven = (.dwMinorVersion >= 1)
          Else
            ver_Win_Seven = (.dwMajorVersion >= 7)
          End If
        End If
      Case VER_PLATFORM_WIN32_WINDOWS
        ver_Win_NTBased = False
        If .dwMajorVersion = 4 Then
          ver_Win_98OrNT4SP3 = (.dwMinorVersion >= 10)
          ver_Win_MEOrXP = (.dwMinorVersion >= 90)
        Else
          ver_Win_98OrNT4SP3 = (.dwMajorVersion > 4)
          ver_Win_MEOrXP = (.dwMajorVersion > 4)
        End If
        ver_Win_98OrNT4SP4 = ver_Win_98OrNT4SP3
        ver_Win_98Or2k = ver_Win_98OrNT4SP3
    End Select
  End With
End Sub

' gibt die letzten 8 Bits von <value> zurück
Function HiByte(ByVal Value As Integer) As Byte
  Dim ret As Byte

  CopyMemory VarPtr(ret), VarPtr(Value) + LenB(ret), LenB(ret)
  HiByte = ret
End Function

' gibt die letzten 16 Bits von <value> zurück
Function HiWord(ByVal Value As Long) As Integer
  Dim ret As Integer

  CopyMemory VarPtr(ret), VarPtr(Value) + LenB(ret), LenB(ret)
  HiWord = ret
End Function

' prüft, ob <txt> eine CLSID enthält
Function IncludesCLSID(ByVal txt As String) As Boolean
  Dim pos As Integer

  pos = InStr(txt, "::{")
  If pos Then
    ' das 38. Zeichen danach muß "}" sein
    IncludesCLSID = (Mid$(txt, pos + 2 + 37, 1) = "}")
  End If
End Function

' Makro "INDEXTOOVERLAYMASK" für TreeView- und ListView-Items (s. MSDN)
Function INDEXTOOVERLAYMASK(ByVal Index As Long) As Long
  INDEXTOOVERLAYMASK = LShiftDWord(Index, 8)
End Function

' Makro "INDEXTOSTATEIMAGEMASK" für TreeView- und ListView-Items (s. MSDN)
Function INDEXTOSTATEIMAGEMASK(ByVal Index As Long) As Long
  INDEXTOSTATEIMAGEMASK = LShiftDWord(Index, 12)
End Function

' initiiert die angeforderten CommonControls
Sub InitCommonCtls(ByVal Ctls As CommonControlsConsts)
  Dim Data As INITCOMMONCONTROLSEXData

  Data.dwICC = Ctls
  Data.dwSize = LenB(Data)
  InitCommonControlsEx Data
End Sub

' prüft, ob die Taste <Key> gedrückt ist
Function IsKeyHoldDown(ByVal Key As Long) As Boolean
  Const MK_LBUTTON = &H1
  Const MK_RBUTTON = &H2
  Const VK_LBUTTON = MK_LBUTTON
  Const VK_RBUTTON = MK_RBUTTON

  If Key = VK_LBUTTON Then
    If GetSystemMetrics(SM_SWAPBUTTON) Then Key = VK_RBUTTON
  ElseIf Key = VK_RBUTTON Then
    If GetSystemMetrics(SM_SWAPBUTTON) Then Key = VK_LBUTTON
  End If

  IsKeyHoldDown = (GetAsyncKeyState(Key) And &H8000&)
End Function

Function IsSamePoint(pt1 As POINT, pt2 As POINT, Optional ByVal xTolerance As Integer = 0, Optional ByVal yTolerance As Integer = 0) As Boolean
  Dim ret As Boolean

  With pt2
    ret = IsWithinRange(pt1.x, .x - xTolerance, .x + xTolerance)
    If ret Then ret = IsWithinRange(pt1.y, .y - yTolerance, .y + yTolerance)
  End With

  IsSamePoint = ret
End Function

' prüft, ob <ItemID> im Menü <hMenu> ein Separator ist
Function IsSeparator(ByVal hMenu As Long, ByVal ItemID As Long, Optional ByVal isPosition As Boolean = False) As Boolean
  Dim Data As MENUITEMINFO

  If IsMenu(hMenu) Then
    With Data
      .cbSize = LenB(Data)
      If ver_Win_98Or2k Then
        .fMask = MIIM_FTYPE
      Else
        .fMask = MIIM_TYPE
      End If
      GetMenuItemInfo hMenu, ItemID, Abs(isPosition), Data
      IsSeparator = (.fType And MFT_SEPARATOR)
    End With
  End If
End Function

' prüft, ob die Tab-Taste gedrückt ist
Function isTabKeyHoldDown() As Boolean
  Const VK_TAB = &H9

  isTabKeyHoldDown = IsKeyHoldDown(VK_TAB)
End Function

' prüft, ob <txt> eine Zahl zwischen <LoBound> und <UpBound> ist
Function isValidNumber(ByVal txt As String, ByVal LoBound As Long, ByVal UpBound As Long) As Boolean
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
    ret = IsWithinRange(CLng(txt), LoBound, UpBound)
  End If

  isValidNumber = ret
End Function

' prüft, ob <hWnd> aktiv ist
' WIRD DERZEIT NICHT VERWENDET
Function isWindowActive(ByVal hWnd As Long) As Boolean
  Dim Data As WINDOWINFO
  Dim ret As Boolean

  With Data
    .cbSize = LenB(Data)
    If ver_Win_98OrNT4SP3 Then
      GetWindowInfo hWnd, Data
    Else
      If GetActiveWindow = hWnd Then .dwWindowStatus = WS_ACTIVECAPTION
    End If

    ret = (.dwWindowStatus = WS_ACTIVECAPTION)
  End With

  isWindowActive = ret
End Function

' prüft, ob <Number> im Intervall [<LoBound>; <UpBound>] liegt
Function IsWithinRange(ByVal Number As Long, ByVal LoBound As Long, ByVal UpBound As Long) As Boolean
  IsWithinRange = ((Number >= LoBound) And (Number <= UpBound))
End Function

' prüft, ob der Punkt <pt> in <rc> liegt
Function IsWithinRectangle(rc As RECT, pt As POINT) As Boolean
  With pt
    IsWithinRectangle = PtInRect(rc, .x, .y)
  End With
End Function

Function IsWithinVisibleClientArea(ByVal hWnd As Long, rcToCheck As RECT, pt As POINT, Optional ByVal PtIsRelativeToScreen As Boolean = True, Optional ByVal hPossibleChildWnd1 As Long = 0, Optional ByVal hPossibleChildWnd2 As Long = 0, Optional ByVal hPossibleChildWnd3 As Long = 0, Optional ByVal hPossibleChildWnd4 As Long = 0, Optional ByVal hPossibleChildWnd5 As Long = 0, Optional ByRef WndFromPtRetVal As Long = 0) As Boolean
  Dim ret As Boolean

  If PtIsRelativeToScreen Then
    With pt
      WndFromPtRetVal = WindowFromPoint(.x, .y)
      If (WndFromPtRetVal = hWnd) Or (WndFromPtRetVal = hPossibleChildWnd1) Or (WndFromPtRetVal = hPossibleChildWnd2) Or (WndFromPtRetVal = hPossibleChildWnd3) Or (WndFromPtRetVal = hPossibleChildWnd4) Or (WndFromPtRetVal = hPossibleChildWnd5) Then
        ScreenToClient hWnd, pt
        ret = IsWithinRectangle(rcToCheck, pt)
      End If
    End With
  Else
    If IsWithinRectangle(rcToCheck, pt) Then
      With pt
        ClientToScreen hWnd, pt
        WndFromPtRetVal = WindowFromPoint(.x, .y)
        ret = (WndFromPtRetVal = hWnd) Or (WndFromPtRetVal = hPossibleChildWnd1) Or (WndFromPtRetVal = hPossibleChildWnd2) Or (WndFromPtRetVal = hPossibleChildWnd3) Or (WndFromPtRetVal = hPossibleChildWnd4) Or (WndFromPtRetVal = hPossibleChildWnd5)
      End With
    End If
  End If

  IsWithinVisibleClientArea = ret
End Function

' gibt die ersten 8 Bits von <value> zurück
Function LoByte(ByVal Value As Integer) As Byte
  Dim ret As Byte

  CopyMemory VarPtr(ret), VarPtr(Value), LenB(ret)
  LoByte = ret
End Function

' gibt die ersten 16 Bits von <value> zurück
Function LoWord(ByVal Value As Long) As Integer
  Dim ret As Integer

  CopyMemory VarPtr(ret), VarPtr(Value), LenB(ret)
  LoWord = ret
End Function

' verschiebt die Bytes von <DWord> um <Bytes> Stellen nach links
Function LShiftDWord(ByVal DWord As Long, ByVal Bytes As Integer) As Long
  LShiftDWord = DWord * (2 ^ Bytes)
End Function

' fügt <Lo> und <Hi> zu einer 32-Bit-Zahl zusammen
Function MakeDWord(ByVal Lo As Integer, ByVal Hi As Integer) As Long
  Dim ret As Long

  CopyMemory VarPtr(ret), VarPtr(Lo), LenB(Lo)
  CopyMemory VarPtr(ret) + LenB(Lo), VarPtr(Hi), LenB(Hi)

  MakeDWord = ret
End Function

' Makro "MAKEINTRESOURCE" (s. MSDN)
Function MAKEINTRESOURCE(ByVal Value As Long) As Long
  MAKEINTRESOURCE = MakeDWord(Value, 0)
End Function

' fügt <Lo> und <Hi> zu einer 16-Bit-Zahl zusammen
Function MakeWord(ByVal Lo As Byte, ByVal Hi As Byte) As Integer
  Dim ret As Integer

  CopyMemory VarPtr(ret), VarPtr(Lo), LenB(Lo)
  CopyMemory VarPtr(ret) + LenB(Lo), VarPtr(Hi), LenB(Hi)

  MakeWord = ret
End Function
'
'Function MergeIcons(ByVal hCursor As Long, ByVal idMerge As Long, hBMPImage As Long, hBMPMask As Long, ptHotSpot As POINT) As Boolean
'  Const DI_MASK = &H1
'  Const DI_NORMAL = &H3
'  Const IMAGE_BITMAP = 0
'  Const SM_CXCURSOR = 13
'  Const SM_CYCURSOR = 14
'  Const SRCCOPY = &HCC0020
'  Dim bmp As BITMAP
'  Dim cxBitmap As Long
'  Dim cyBitmap As Long
'  Dim cxCursor As Long
'  Dim cyCursor As Long
'  Dim hBMP As Long
'  Dim hBMPImageWork As Long
'  Dim hBMPMaskWork As Long
'  Dim hBMPTemp1 As Long
'  Dim hBMPTemp2 As Long
'  Dim hDCBitmap As Long
'  Dim hDCCursor As Long
'  Dim xDraw As Long
'  Dim yDraw As Long
'
'  cxCursor = GetSystemMetrics(SM_CXCURSOR)
'  cyCursor = GetSystemMetrics(SM_CYCURSOR)
'
'  ' find the lower corner of the cursor and put it there.
'  ' do this whether or not we have an idMerge because it will set the hotspot
'  GetCursorLowerRight hCursor, xDraw, yDraw, ptHotSpot
'  If idMerge <> -1 Then
'    hBMP = LoadImage(App.hInstance, idMerge, IMAGE_BITMAP, 0, 0, 0)
'    If hBMP Then
'      GetObjectAPI hBMP, LenB(bmp), bmp
'      cxBitmap = bmp.bmWidth
'      cyBitmap = bmp.bmHeight / 2
'
'      If xDraw + cxBitmap > cxCursor Then
'        xDraw = cxCursor - cxBitmap
'      End If
'      If yDraw + cyBitmap > cyCursor Then
'        yDraw = cyCursor - cyBitmap
'      End If
'    End If
'  Else
'    hBMP = 0
'  End If
'
'  hDCCursor = CreateCompatibleDC(0)
'
'  hBMPMaskWork = CreateBitmapAsLong(cxCursor, cyCursor, 1, 1, 0)
'  hBMPImageWork = CreateColorBitmap(cxCursor, cyCursor)
'
'  If (hDCCursor <> 0) And (hBMPMaskWork <> 0) And (hBMPImageWork <> 0) Then
'    hBMPTemp1 = SelectObject(hDCCursor, hBMPImageWork)
'    DrawIconEx hDCCursor, 0, 0, hCursor, 0, 0, 0, 0, DI_NORMAL
'
'    If hBMP Then
'      hDCBitmap = CreateCompatibleDC(0)
'      hBMPTemp2 = SelectObject(hDCBitmap, hBMP)
'
'      ' blt the two bitmaps onto the color and mask bitmaps for the cursor
'      BitBlt hDCCursor, xDraw, yDraw, cxBitmap, cyBitmap, hDCBitmap, 0, 0, SRCCOPY
'    End If
'    SelectObject hDCCursor, hBMPMaskWork
'
'    DrawIconEx hDCCursor, 0, 0, hCursor, 0, 0, 0, 0, DI_MASK
'
'    If hBMP Then
'      BitBlt hDCCursor, xDraw, yDraw, cxBitmap, cyBitmap, hDCBitmap, 0, cyBitmap, SRCCOPY
'
'      ' select back in the old bitmaps
'      SelectObject hDCBitmap, hBMPTemp2
'      DeleteDC hDCBitmap
'      DeleteObject hBMP
'    End If
'    ' select back in the old bitmaps
'    SelectObject hDCCursor, hBMPTemp1
'  End If
'
'  If hDCCursor Then
'    DeleteDC hDCCursor
'  End If
'
'  hBMPImage = hBMPImageWork
'  hBMPMask = hBMPMaskWork
'  MergeIcons = ((hBMPImage <> 0) And (hBMPMask <> 0))
'End Function

' lädt den Mauszeiger <MousePointer> und gibt ein Handle darauf zurück
Function MousePointerConstTohCursor(ByVal MousePointer As MousePointerConstants) As Long
  Dim Flag As Long

  Select Case MousePointer
    Case MousePointerConstants.vbArrow, MousePointerConstants.vbDefault
      Flag = OCR_NORMAL
    Case MousePointerConstants.vbArrowHourglass
      Flag = OCR_APPSTARTING
    Case MousePointerConstants.vbArrowQuestion
      Flag = OCR_HELP
    Case MousePointerConstants.vbCrosshair
      Flag = OCR_CROSS
    Case MousePointerConstants.vbHourglass
      Flag = OCR_WAIT
    Case MousePointerConstants.vbIbeam
      Flag = OCR_IBEAM
    Case MousePointerConstants.vbIconPointer
      Flag = OCR_ICOCUR
    Case MousePointerConstants.vbNoDrop
      Flag = OCR_NO
    Case MousePointerConstants.vbSizeAll
      Flag = OCR_SIZEALL
    Case MousePointerConstants.vbSizeNESW
      Flag = OCR_SIZENESW
    Case MousePointerConstants.vbSizeNS
      Flag = OCR_SIZENS
    Case MousePointerConstants.vbSizeNWSE
      Flag = OCR_SIZENWSE
    Case MousePointerConstants.vbSizePointer
      Flag = OCR_SIZEALL
    Case MousePointerConstants.vbSizeWE
      Flag = OCR_SIZEWE
    Case MousePointerConstants.vbUpArrow
      Flag = OCR_UP
  End Select

  MousePointerConstTohCursor = LoadImage(0, MAKEINTRESOURCE(Flag), IMAGE_CURSOR, 0, 0, LR_DEFAULTCOLOR Or LR_DEFAULTSIZE Or LR_SHARED)
End Function

' converts a octal number to a decimal number
Function OctalToDecimal(ByVal OctVal As String) As Double
  Dim i As Integer
  Dim ret As Double

  For i = 1 To Len(OctVal)
    ret = ret + CDbl(Val(Mid$(OctVal, i, 1))) * 8 ^ (Len(OctVal) - i)
  Next

  OctalToDecimal = ret
End Function

' Makro "OVERLAYMASKTOINDEX" für TreeView- und ListView-Items (Gegenstück zu INDEXTOOVERLAYMASK)
Function OVERLAYMASKTOINDEX(ByVal Mask As Long) As Long
  OVERLAYMASKTOINDEX = RShiftDWord(Mask, 8)
End Function

' entfernt den Backslash am Ende von <txt> (wenn einer dran ist)
' vorher wird geprüft, ob <txt> ein Verzeichnis ist
Function RemoveBackslash(ByVal txt As String, Optional ByVal OnlyIfDirectory As Boolean = True, Optional ByVal Backslash As String = "\") As String
  If OnlyIfDirectory Then
    If Not IsDirectory(txt) Then
      If Not (Backslash = "/" And IsFTPURL(txt)) Then
        RemoveBackslash = Trim$(Left$(txt, lstrlenA(txt)))
        Exit Function
      End If
    End If
  End If
  If Backslash = "\" Then
    txt = txt & Chr(0)
    PathRemoveBackslash txt
    txt = Trim$(Left$(txt, lstrlenA(txt)))
  Else
    If Right$(txt, Len(Backslash)) = Backslash Then txt = Left$(txt, Len(txt) - Len(Backslash))
  End If

  RemoveBackslash = txt
End Function

' ersetzt in <hMenu> alle doppelten Separators durch einen einzelnen
Sub RemoveDoubleSeparators(ByVal hMenu As Long)
  Dim i As Integer
  Dim iLastSep As Integer

  iLastSep = -2
  If IsMenu(hMenu) Then
    While IsSeparator(hMenu, 0, True)
      RemoveMenu hMenu, 0, MF_BYPOSITION
    Wend

    While i < GetMenuItemCount(hMenu)
      If IsSeparator(hMenu, i, True) Then
        If i = iLastSep + 1 Then
          ' 2 Separators hintereinander
          RemoveMenu hMenu, i, MF_BYPOSITION
        Else
          iLastSep = i
          i = i + 1
        End If
      Else
        i = i + 1
      End If
    Wend

    i = i - 1
    While IsSeparator(hMenu, i, True)
      RemoveMenu hMenu, i, MF_BYPOSITION
      i = i - 1
    Wend
  End If
End Sub

' entfernt von <Pfad> die ersten <c> Segmente
Function RemoveFirstFolders(ByVal path As String, ByVal c As Integer, Optional ByVal Sep As String = "\") As String
  Dim i As Integer
  Dim pos As Integer

  If IsUNC(path) Then path = Mid$(path, 3)
  path = AddBackslash(path, False, Sep)
  If Left$(path, Len("ftp://")) = "ftp://" Then c = c + 2
  For i = 1 To c
    pos = InStr(path, Sep)
    If pos = 0 Then Exit For
    path = Mid$(path, pos + 1)
  Next

  RemoveFirstFolders = path
End Function

' entfernt von <Pfad> die letzten <c> Segmente
Function RemoveLastFolders(ByVal path As String, ByVal c As Integer, Optional ByVal Sep As String = "\") As String
  Dim i As Integer
  Dim pos As Integer

  path = RemoveBackslash(path, , Sep)
  For i = 1 To c
    pos = InStrRev(path, Sep)
    If pos = 0 Then
      ' <Path> hat nur 1 Segment
      path = ""
      Exit For
    End If
    path = Left$(path, pos - 1)
    If path = "ftp:/" Then path = ""
  Next
  path = AddBackslash(path, , Sep)

  RemoveLastFolders = path
End Function

Function removeNull(txt As String) As String
  removeNull = Left$(txt, lstrlenA(txt))
End Function

' verschiebt die Bytes von <DWord> um <Bytes> Stellen nach rechts
Function RShiftDWord(ByVal DWord As Long, ByVal Bytes As Integer) As Long
  RShiftDWord = DWord \ (2 ^ Bytes)
End Function

' markiert den Text in <TxtBox>
Sub selectText(ByVal TxtBox As TextBox, Optional ByVal OnlyIfTabKeyIsHoldDown As Boolean = True)
  Const EM_SETSEL = &HB1

  If OnlyIfTabKeyIsHoldDown Then
    If isTabKeyHoldDown Then PostMessage TxtBox.hWnd, EM_SETSEL, 0, Len(TxtBox)
  Else
    PostMessage TxtBox.hWnd, EM_SETSEL, 0, Len(TxtBox)
  End If
End Sub

' gibt zurück, ob komprimierte Items in einer anderen Farbe dargestellt werden sollen oder nicht
' (ermittelt die Systemeinstellung)
Function ShouldUseAlternativeClr() As Boolean
  Dim Data As Long
  Dim Datasize As Long
  Dim ret As Long

  Datasize = LenB(Data)
  ret = SHGetValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowCompColor", VarPtr(REG_DWORD), ByVal VarPtr(Data), VarPtr(Datasize))
  If ret = ERROR_FILE_NOT_FOUND Then Data = 1

  ShouldUseAlternativeClr = Data
End Function

' Makro "STATEIMAGEMASKTOINDEX" für TreeView- und ListView-Items (Gegenstück zu INDEXTOSTATEIMAGEMASK)
Function STATEIMAGEMASKTOINDEX(ByVal Mask As Long) As Long
  STATEIMAGEMASKTOINDEX = RShiftDWord(Mask, 12)
End Function

' transformiert eine OLE-Farbe in eine RGB-Farbe
Function TranslateColor(ByVal Clr As OLE_COLOR, Optional ByVal hPal As Long = 0) As Long
  If OleTranslateColor(Clr, hPal, TranslateColor) Then TranslateColor = CLR_INVALID
End Function
