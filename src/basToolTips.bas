Attribute VB_Name = "basToolTips"
Option Explicit

  ' lokale Konstanten

  Global Const TTN_FIRST = (-520)
  #If Debuging Then
'    Global Const WM_USER = &H400
  #End If

  ' Klassenname für ToolTip
  #If Debuging Then
    Global Const TOOLTIPS_CLASS = "tooltips_class32"
  #End If

  ' Konstanten für TTM_SETDELAYEVENT
  #If Debuging Then
    Global Const TTDT_AUTOMATIC = 0
    Global Const TTDT_AUTOPOP = 2
    Global Const TTDT_INITIAL = 3
    Global Const TTDT_RESHOW = 1
  #End If

  ' Konstanten für TTM_SETTITLE
  #If Debuging Then
    Global Const TTI_ERROR = 3
    Global Const TTI_INFO = 1
    Global Const TTI_NONE = 0
    Global Const TTI_WARNING = 2
  #End If

  ' Konstanten für TTTOOLINFO
  #If Debuging Then
    Global Const TTF_ABSOLUTE = &H80
    Global Const TTF_CENTERTIP = &H2
    Global Const TTF_DI_SETITEM = &H8000
    Global Const TTF_IDISHWND = &H1
    Global Const TTF_PARSELINKS = &H1000
    Global Const TTF_RTLREADING = &H4
    Global Const TTF_SUBCLASS = &H10
    Global Const TTF_TRACK = &H20
    Global Const TTF_TRANSPARENT = &H100
  #End If

  ' CommonControls-Messages
  #If Debuging Then
    Global Const CCM_FIRST = &H2000
    Global Const CCM_DPISCALE = (CCM_FIRST + &HC)
    Global Const CCM_GETCOLORSCHEME = (CCM_FIRST + 3)
    Global Const CCM_GETDROPTARGET = (CCM_FIRST + 4)
    Global Const CCM_GETUNICODEFORMAT = (CCM_FIRST + 6)
    Global Const CCM_GETVERSION = (CCM_FIRST + &H8)
    Global Const CCM_SETBKCOLOR = (CCM_FIRST + 1)
    Global Const CCM_SETCOLORSCHEME = (CCM_FIRST + 2)
    Global Const CCM_SETNOTIFYWINDOW = (CCM_FIRST + &H9)
    Global Const CCM_SETUNICODEFORMAT = (CCM_FIRST + 5)
    Global Const CCM_SETVERSION = (CCM_FIRST + &H7)
    Global Const CCM_SETWINDOWTHEME = (CCM_FIRST + &HB)
  #End If

  ' ToolTip-Messages
  #If Debuging Then
    Global Const TTM_ACTIVATE = (WM_USER + 1)
    Global Const TTM_ADDTOOLW = (WM_USER + 50)
    Global Const TTM_ADDTOOL = (WM_USER + 4)
    Global Const TTM_ADJUSTRECT = (WM_USER + 31)
    Global Const TTM_DELTOOLW = (WM_USER + 51)
    Global Const TTM_DELTOOL = (WM_USER + 5)
    Global Const TTM_ENUMTOOLSW = (WM_USER + 58)
    Global Const TTM_ENUMTOOLS = (WM_USER + 14)
    Global Const TTM_GETBUBBLESIZE = (WM_USER + 30)
    Global Const TTM_GETCURRENTTOOLW = (WM_USER + 59)
    Global Const TTM_GETCURRENTTOOL = (WM_USER + 15)
    Global Const TTM_GETDELAYTIME = (WM_USER + 21)
    Global Const TTM_GETMARGIN = (WM_USER + 27)
    Global Const TTM_GETMAXTIPWIDTH = (WM_USER + 25)
    Global Const TTM_GETTEXTW = (WM_USER + 56)
    Global Const TTM_GETTEXT = (WM_USER + 11)
    Global Const TTM_GETTIPBKCOLOR = (WM_USER + 22)
    Global Const TTM_GETTIPTEXTCOLOR = (WM_USER + 23)
    Global Const TTM_GETTITLE = (WM_USER + 35)
    Global Const TTM_GETTOOLCOUNT = (WM_USER + 13)
    Global Const TTM_GETTOOLINFOW = (WM_USER + 53)
    Global Const TTM_GETTOOLINFO = (WM_USER + 8)
    Global Const TTM_HITTESTW = (WM_USER + 55)
    Global Const TTM_HITTEST = (WM_USER + 10)
    Global Const TTM_NEWTOOLRECTW = (WM_USER + 52)
    Global Const TTM_NEWTOOLRECT = (WM_USER + 6)
    Global Const TTM_POP = (WM_USER + 28)
    Global Const TTM_POPUP = (WM_USER + 34)
    Global Const TTM_RELAYEVENT = (WM_USER + 7)
    Global Const TTM_SETDELAYTIME = (WM_USER + 3)
    Global Const TTM_SETMARGIN = (WM_USER + 26)
    Global Const TTM_SETMAXTIPWIDTH = (WM_USER + 24)
    Global Const TTM_SETTIPBKCOLOR = (WM_USER + 19)
    Global Const TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
    Global Const TTM_SETTITLEW = (WM_USER + 33)
    Global Const TTM_SETTITLE = (WM_USER + 32)
    Global Const TTM_SETTOOLINFOW = (WM_USER + 54)
    Global Const TTM_SETTOOLINFO = (WM_USER + 9)
    Global Const TTM_SETWINDOWTHEME = CCM_SETWINDOWTHEME
    Global Const TTM_TRACKACTIVATE = (WM_USER + 17)
    Global Const TTM_TRACKPOSITION = (WM_USER + 18)
    Global Const TTM_UPDATE = (WM_USER + 29)
    Global Const TTM_UPDATETIPTEXTW = (WM_USER + 57)
    Global Const TTM_UPDATETIPTEXT = (WM_USER + 12)
    Global Const TTM_WINDOWFROMPOINT = (WM_USER + 16)
  #End If

  ' ToolTip-Notifications
  Global Const TTN_GETDISPINFOW = (TTN_FIRST - 10)
  Global Const TTN_GETDISPINFO = (TTN_FIRST - 0)
  #If Debuging Then
    Global Const TTN_LINKCLICK = (TTN_FIRST - 3)
    Global Const TTN_NEEDTEXTW = TTN_GETDISPINFOW
    Global Const TTN_NEEDTEXT = TTN_GETDISPINFO
    Global Const TTN_POP = (TTN_FIRST - 2)
    Global Const TTN_SHOW = (TTN_FIRST - 1)
  #End If

  ' ToolTip-Styles
  #If Debuging Then
    Global Const TTS_ALWAYSTIP = &H1
    Global Const TTS_BALLOON = &H40
    Global Const TTS_CLOSE = &H80
    Global Const TTS_NOANIMATE = &H10
    Global Const TTS_NOFADE = &H20
    Global Const TTS_NOPREFIX = &H2
  #End If


  ' lokale Types

  #If Debuging Then
    Type COLORSCHEME
      dwSize As Long
      clrBtnHighlight As Long
      clrBtnShadow As Long
    End Type

    Type NMCUSTOMDRAW
      hdr As NMHDR
      dwDrawStage As Long
      hDC As Long
      rc As RECT
      dwItemSpec As Long
      uItemState As Long
      lItemlParam As Long
    End Type

    Type NMTTCUSTOMDRAW
      nmcd As NMCUSTOMDRAW
      uDrawFlags As Long
    End Type

    Type TTGETTITLE
      dwSize As Long
      uTitleBitmap As Long
      cch As Long
      pszTitle As String
    End Type

    Type TTTOOLINFO
      cbSize As Long
      uFlags As Long
      hWnd As Long
      uId As Long
      rc As RECT
      hinst As Long
      lpszText As Long
      lParam As Long
      lpReserved As Long
    End Type

    Type TTHITTESTINFO
      hWnd As Long
      pt As POINT
      ti As TTTOOLINFO
    End Type
  #End If

  Type NMTTDISPINFO
    hdr As NMHDR
    lpszText As Long
    szText(0 To 79) As Byte
    hinst As Long
    uFlags As Long
    lParam As Long
  End Type
