Attribute VB_Name = "basShell"
Option Explicit

  ' öffentliche Enums

  Public Enum CSIDLConstants
    CSIDL_First = &H0

    CSIDL_DESKTOP = &H0
    CSIDL_INTERNET = &H1
    CSIDL_PROGRAMS = &H2
    CSIDL_CONTROLS = &H3
    CSIDL_PRINTERS = &H4
    CSIDL_PERSONAL = &H5
    CSIDL_FAVORITES = &H6
    CSIDL_STARTUP = &H7
    CSIDL_RECENT = &H8
    CSIDL_SENDTO = &H9
    CSIDL_BITBUCKET = &HA
    CSIDL_STARTMENU = &HB
    CSIDL_MYDOCUMENTS = &HC
    CSIDL_MYMUSIC = &HD
    CSIDL_MYVIDEO = &HE
    CSIDL_DESKTOPDIRECTORY = &H10
    CSIDL_DRIVES = &H11
    CSIDL_NETWORK = &H12
    CSIDL_NETHOOD = &H13
    CSIDL_FONTS = &H14
    CSIDL_TEMPLATES = &H15
    CSIDL_COMMON_STARTMENU = &H16
    CSIDL_COMMON_PROGRAMS = &H17
    CSIDL_COMMON_STARTUP = &H18
    CSIDL_COMMON_DESKTOPDIRECTORY = &H19
    CSIDL_APPDATA = &H1A
    CSIDL_PRINTHOOD = &H1B
    CSIDL_LOCAL_APPDATA = &H1C
    CSIDL_ALTSTARTUP = &H1D
    CSIDL_COMMON_ALTSTARTUP = &H1E
    CSIDL_COMMON_FAVORITES = &H1F
    CSIDL_INTERNET_CACHE = &H20
    CSIDL_COOKIES = &H21
    CSIDL_HISTORY = &H22
    CSIDL_COMMON_APPDATA = &H23
    CSIDL_WINDOWS = &H24
    CSIDL_SYSTEM = &H25
    CSIDL_PROGRAM_FILES = &H26
    CSIDL_MYPICTURES = &H27
    CSIDL_PROFILE = &H28
    CSIDL_SYSTEMX86 = &H29
    CSIDL_PROGRAM_FILESX86 = &H2A
    CSIDL_PROGRAM_FILES_COMMON = &H2B
    CSIDL_PROGRAM_FILES_COMMONX86 = &H2C
    CSIDL_COMMON_TEMPLATES = &H2D
    CSIDL_COMMON_DOCUMENTS = &H2E
    CSIDL_COMMON_ADMINTOOLS = &H2F
    CSIDL_ADMINTOOLS = &H30
    CSIDL_CONNECTIONS = &H31
    CSIDL_COMMON_MUSIC = &H35
    CSIDL_COMMON_PICTURES = &H36
    CSIDL_COMMON_VIDEO = &H37
    CSIDL_RESOURCES = &H38
    CSIDL_RESOURCES_LOCALIZED = &H39
    CSIDL_COMMON_OEM_LINKS = &H3A
    CSIDL_CDBURN_AREA = &H3B
    CSIDL_COMPUTERSNEARME = &H3D

    CSIDL_Last = &H3D

    CSIDL_FLAG_CREATE = &H8000
    CSIDL_FLAG_DONT_VERIFY = &H4000
    CSIDL_FLAG_MASK = &HFF
    CSIDL_FLAG_NO_ALIAS = &H1000
    CSIDL_FLAG_PER_USER_INIT = &H800
  End Enum

  Public Enum SHCNRFConstants
    ' interrupt level notifications from the file system (???)
    SHCNRF_InterruptLevel = &H1
    ' shell level notifications from the shell
    SHCNRF_ShellLevel = &H2
    ' interrupt events from the whole subtree (???)
    SHCNRF_RecursiveInterrupt = &H1000
    ' messages use shared memory -> use SHChangeNotification_Lock() and SHChangeNotification_Unlock()
    SHCNRF_NewDelivery = &H8000
  End Enum

  Public Enum SHCNEConstants
    ' Kombination aller Events
    SHCNE_ALLEVENTS = &H7FFFFFFF
    ' eine Dateityp-Verknüpfung wurde geändert
    SHCNE_ASSOCCHANGED = &H8000000
    ' die Attribute eines Items wurden geändert
    SHCNE_ATTRIBUTES = &H800&
    ' ein NonFolder-Item wurde erzeugt
    SHCNE_CREATE = &H2&
    ' ein NonFolder-Item wurde gelöscht
    SHCNE_DELETE = &H4&
    ' Kombination aller Disk-Events
    SHCNE_DISKEVENTS = &H2381F
    ' ein Laufwerk wurde hinzugefügt
    SHCNE_DRIVEADD = &H100&
    ' ein Laufwerk wurde hinzugefügt, und die Shell
    ' sollte ein Fenster für es erzeugen
    SHCNE_DRIVEADDGUI = &H10000
    ' ein Laufwerk wurde entfernt
    SHCNE_DRIVEREMOVED = &H80&
    SHCNE_EXTENDED_EVENT = &H4000000
    ' der freie Speicherplatz eines Laufwerkes hat
    ' sich geändert
    SHCNE_FREESPACE = &H40000
    ' Kombination aller globaler Events
    SHCNE_GLOBALEVENTS = &HC0581E0
    ' das Event wurde aufgrund eines System-Interrupts
    ' ausgelöst
    SHCNE_INTERRUPT = &H80000000
    ' ein Medium wurde in ein Laufwerk eingelegt
    SHCNE_MEDIAINSERTED = &H20&
    ' ein Medium wurde aus einem Laufwerk entnommen
    SHCNE_MEDIAREMOVED = &H40&
    ' ein Folder-Item wurde erzeugt
    SHCNE_MKDIR = &H8
    ' ein Ordner wurde für den Netz-Zugriff freigegeben
    SHCNE_NETSHARE = &H200&
    ' ein Ordner wurde für den Netz-Zugriff gesperrt
    SHCNE_NETUNSHARE = &H400&
    ' ein Ordner wurde umbenannt
    SHCNE_RENAMEFOLDER = &H20000
    ' ein NonFolder-Item wurde umbenannt
    SHCNE_RENAMEITEM = &H1&
    ' ein Folder-Item wurde gelöscht
    SHCNE_RMDIR = &H10&
    ' der Computer wurde vom Server getrennt
    SHCNE_SERVERDISCONNECT = &H4000&
    ' der Inhalt eines Ordners wurde geändert
    SHCNE_UPDATEDIR = &H1000&
    ' ein Bild in der SysImageList wurde geändert
    SHCNE_UPDATEIMAGE = &H8000&
    ' der Inhalt eines NonFolder-Items wurde geändert
    SHCNE_UPDATEITEM = &H2000&
  End Enum

  Public Enum SHCNEEConstants
    ' die Anordnung der Items wurde geändert
    ' dwItem2 ist die pIDL des betroffenen Items
    SHCNEE_ORDERCHANGED = 2
    ' eine MSI-Installation wurde gestartet
    ' dwItem2 ist der Produkt-Code
    SHCNEE_MSI_CHANGE = 4
    ' eine MSI-Deinstallation wurde gestartet
    ' dwItem2 ist der Produkt-Code
    SHCNEE_MSI_UNINSTALL = 5

    ' selbst entdeckte Events
    SHCNEE_PROGRAMSTARTED = 6
    ' ein Programm wurde gestartet
    ' dwItem2 ist die pIDL der Programmdatei
  End Enum


  ' lokale Konstanten

  Private Const strIID_IContextMenu = "{000214E4-0000-0000-C000-000000000046}"
  Private Const strIID_IContextMenu2 = "{000214F4-0000-0000-C000-000000000046}"
  Private Const strIID_IContextMenu3 = "{BCFCE0A0-EC17-11D0-8D10-00A0C90F2719}"
  Private Const strIID_IDataObject = "{0000010E-0000-0000-C000-000000000046}"
  Private Const strIID_IDragDropHelper = "{4657278A-411B-11d2-839A-00C04FD918D0}"
  Private Const strIID_IDragSourceHelper = "{DE5BF786-477A-11d2-839D-00C04FD918D0}"
  Private Const strIID_IDragSourceHelper2 = "{83E07D0D-0C5F-4163-BF1A-60B274051E40}"
  Private Const strIID_IDropTarget = "{00000122-0000-0000-C000-000000000046}"
  Private Const strIID_IDropTargetHelper = "{4657278B-411B-11d2-839A-00C04FD918D0}"
  Private Const strIID_IPersistFolder = "{000214EA-0000-0000-C000-000000000046}"
  Private Const strIID_IPersistFolder2 = "{1AC3D9F0-175C-11D1-95BE-00609797EA4F}"
  Private Const strIID_IQueryInfo = "{00021500-0000-0000-C000-000000000046}"
  Private Const strIID_IShellFolder = "{000214E6-0000-0000-C000-000000000046}"
  Private Const strIID_IShellIcon = "{000214E5-0000-0000-C000-000000000046}"
  Private Const strIID_IShellIconOverlay = "{7D688A70-C613-11D0-999B-00C04FD655E1}"
  Private Const strIID_IShellLinkA = "{000214EE-0000-0000-C000-000000000046}"
  Private Const strIID_IShellLinkW = "{000214F9-0000-0000-C000-000000000046}"
  #If Debuging Then
    Private Const strCLSID_AugmentedShellFolder = "{91EA3F8B-C99B-11d0-9815-00C04FD91972}"
    Private Const strCLSID_AugmentedShellFolder2 = "{6413BA2C-B461-11d1-A18A-080036B11A03}"
    Private Const strIID_IAugmentedShellFolder = "{91EA3F8C-C99B-11d0-9815-00C04FD91972}"
    Private Const strIID_IAugmentedShellFolder2 = "{8DB3B3F4-6CFE-11d1-8AE9-00C04FD918D0}"
    Private Const strIID_IDelegateFolder = "{ADD8BA80-002B-11D0-8F0F-00C04FD7D062}"
    Private Const strIID_IEnumUICommand = "{869447DA-9F84-4E2A-B92D-00642DC8A911}"
    Private Const strIID_IShellFolder2 = "{93F2F68C-1D1B-11D3-A30E-00C04F79ABD1}"
    Private Const strIID_IUICommand = "{4026DFB9-7691-4142-B71C-DCF08EA4DD9C}"
    Private Const strIID_IUICommandTarget = "{2CB95001-FC47-4064-89B3-328F2FE60F44}"
    Private Const strIID_IUIElement = "{EC6FE84F-DC14-4FBB-889F-EA50FE27FE0F}"
    Private Const strIID_IThumbnailProvider = "{e357fccd-a995-4576-b01f-234630154e96}"
  #End If

  ' Konstanten für GetDriveType
  #If Debuging Then
    Private Const DRIVE_CDROM = 5
    Private Const DRIVE_RAMDISK = 6
    Private Const DRIVE_REMOVABLE = 2
  #End If
  Private Const DRIVE_FIXED = 3
  Private Const DRIVE_NO_ROOT_DIR = 1
  Private Const DRIVE_REMOTE = 4

  ' Konstanten für RegCreateKeyEx
  Private Const KEY_QUERY_VALUE = &H1
  Private Const KEY_SET_VALUE = &H2
  #If Debuging Then
    Private Const HKEY_CURRENT_CONFIG = &H80000005
    Private Const HKEY_CURRENT_USER = &H80000001
    Private Const HKEY_DYN_DATA = &H80000006
    Private Const HKEY_LOCAL_MACHINE = &H80000002
    Private Const HKEY_PERF_ROOT = HKEY_LOCAL_MACHINE
    Private Const HKEY_PERFORMANCE_DATA = &H80000004
    Private Const HKEY_USERS = &H80000003
    Private Const READ_CONTROL = &H20000
    Private Const STANDARD_RIGHTS_ALL = &H1F0000
    Private Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
    Private Const STANDARD_RIGHTS_READ = READ_CONTROL
    Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
    Private Const STANDARD_RIGHTS_WRITE = READ_CONTROL
    Private Const SYNCHRONIZE = &H100000
    Private Const KEY_CREATE_LINK = &H20
    Private Const KEY_CREATE_SUB_KEY = &H4
    Private Const KEY_ENUMERATE_SUB_KEYS = &H8
    Private Const KEY_EVENT = &H1
    Private Const KEY_LENGTH_MASK = &HFFFF0000
    Private Const KEY_NOTIFY = &H10
    Private Const KEY_READ = (STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE)
    Private Const KEY_EXECUTE = KEY_READ
    Private Const KEY_WRITE = (STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE)
    Private Const KEY_ALL_ACCESS = (STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE)
    Private Const REG_BINARY = 3
    Private Const REG_DWORD = 4
    Private Const REG_DWORD_BIG_ENDIAN = 5
    Private Const REG_DWORD_LITTLE_ENDIAN = 4
    Private Const REG_EXPAND_SZ = 2
    Private Const REG_LINK = 6
    Private Const REG_MULTI_SZ = 7
    Private Const REG_NONE = 0
    Private Const REG_QWORD = 11
    Private Const REG_QWORD_LITTLE_ENDIAN = 11
  #End If
  Private Const HKEY_CLASSES_ROOT = &H80000000
  Private Const REG_SZ = 1

  ' Konstanten für SHGetDataFromIDList
  #If Debuging Then
    Private Const SHDID_COMPUTER_AUDIO = 19
    Private Const SHDID_COMPUTER_IMAGING = 18
    Private Const SHDID_COMPUTER_SHAREDDOCS = 20
    Private Const SHDID_FS_DIRECTORY = 3
    Private Const SHDID_FS_FILE = 2
    Private Const SHDID_FS_OTHER = 4
    Private Const SHDID_NET_DOMAIN = 13
    Private Const SHDID_NET_OTHER = 17
    Private Const SHDID_NET_RESTOFNET = 16
    Private Const SHDID_NET_SERVER = 14
    Private Const SHDID_NET_SHARE = 15
    Private Const SHDID_ROOT_REGITEM = 1
  #End If
  Private Const SHDID_COMPUTER_CDROM = 10
  Private Const SHDID_COMPUTER_DRIVE35 = 5
  Private Const SHDID_COMPUTER_DRIVE525 = 6
  Private Const SHDID_COMPUTER_FIXED = 8
  Private Const SHDID_COMPUTER_NETDRIVE = 9
  Private Const SHDID_COMPUTER_OTHER = 12
  Private Const SHDID_COMPUTER_RAMDISK = 11
  Private Const SHDID_COMPUTER_REMOVABLE = 7
  #If Debuging Then
    Private Const SHGDFIL_FINDDATA = 1
    Private Const SHGDFIL_NETRESOURCE = 2
  #End If
  Private Const SHGDFIL_DESCRIPTIONID = 3

  ' Konstanten für SHGetFileInfo
  #If Debuging Then
    Private Const SHGFI_ADDOVERLAYS = &H20
    Private Const SHGFI_ATTR_SPECIFIED = &H20000
    Private Const SHGFI_ATTRIBUTES = &H800
    Private Const SHGFI_EXETYPE = &H2000
    Private Const SHGFI_ICON = &H100
    Private Const SHGFI_ICONLOCATION = &H1000
    Private Const SHGFI_LINKOVERLAY = &H8000
    Private Const SHGFI_OVERLAYINDEX = &H40
    Private Const SHGFI_SELECTED = &H10000
    Private Const SHGFI_SHELLICONSIZE = &H4
    Private Const SHGFI_TYPENAME = &H400
  #End If
  Private Const SHGFI_DISPLAYNAME = &H200
  Private Const SHGFI_LARGEICON = &H0
  Private Const SHGFI_OPENICON = &H2
  Private Const SHGFI_PIDL = &H8
  Private Const SHGFI_SMALLICON = &H1
  Private Const SHGFI_SYSICONINDEX = &H4000
  Private Const SHGFI_USEFILEATTRIBUTES = &H10

  ' Konstanten für SHGetIconOverlayIndex
  Private Const IDO_SHGIOI_LINK = &HFFFFFFE
  Private Const IDO_SHGIOI_SHARE = &HFFFFFFF
  Private Const IDO_SHGIOI_SLOWFILE = &HFFFFFFFD

  ' Konstanten für WideCharToMultiByte
  #If Debuging Then
    Private Const CP_MACCP = 2
    Private Const CP_OEMCP = 1
    Private Const CP_SYMBOL = 42
    Private Const CP_THREAD_ACP = 3
    Private Const CP_UTF7 = 65000
    Private Const CP_UTF8 = 65001
  #End If
  Private Const CP_ACP = 0
  #If Debuging Then
    Private Const WC_COMPOSITECHECK = &H200
    Private Const WC_DEFAULTCHAR = &H40
    Private Const WC_DISCARDNS = &H10
    Private Const WC_NO_BEST_FIT_CHARS = &H400
    Private Const WC_SEPCHARS = &H20
  #End If


  ' globale Konstanten

  Global Const ERROR_SUCCESS = 0
  Global Const MAX_PATH = 260
  Global Const NOERROR = 0
  Global Const S_OK = &H0
  Global Const WM_USER = &H400
  Global Const WM_SHNOTIFY = WM_USER + 1

  ' Konstanten für GetFileAttributes
  #If Debuging Then
    Global Const FILE_ATTRIBUTE_DEVICE = &H40
    Global Const FILE_ATTRIBUTE_NOT_CONTENT_INDEXED = &H2000
    Global Const FILE_ATTRIBUTE_OFFLINE = &H1000
    Global Const FILE_ATTRIBUTE_REPARSE_POINT = &H400
    Global Const FILE_ATTRIBUTE_SPARSE_FILE = &H200
    Global Const FILE_ATTRIBUTE_TEMPORARY = &H100
  #End If
  Global Const FILE_ATTRIBUTE_ARCHIVE = &H20
  Global Const FILE_ATTRIBUTE_COMPRESSED = &H800
  Global Const FILE_ATTRIBUTE_DIRECTORY = &H10
  Global Const FILE_ATTRIBUTE_ENCRYPTED = &H4000
  Global Const FILE_ATTRIBUTE_HIDDEN = &H2
  Global Const FILE_ATTRIBUTE_NORMAL = &H80
  Global Const FILE_ATTRIBUTE_READONLY = &H1
  Global Const FILE_ATTRIBUTE_SYSTEM = &H4


  ' lokale Types

  Private Type REPARSE_DATA_BUFFER
    ReparseTagLo As Integer
    ReparseTagHi As Integer
    ReparseDataLength As Integer
    Reserved As Integer
    SubstituteNameOffset As Integer
    SubstituteNameLength As Integer
    PrintNameOffset As Integer
    PrintNameLength As Integer
    PathBuffer As Integer
  End Type

  Private Type SHDESCRIPTIONID
    dwDescriptionId As Long
    CLSID As UUID
  End Type

  Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
  End Type


  ' globale Types

  ' für AutoUpdate
  Type SHCHANGENOTIFYENTRY
    pIDL As Long
    fRecursive As Long
  End Type

  Type SHNOTIFY
    dwItem1 As Long
    dwItem2 As Long
  End Type


  ' globale Variablen

  Global IDesktop As IVBShellFolder
  Global IID_IContextMenu As UUID
  Global IID_IContextMenu2 As UUID
  Global IID_IContextMenu3 As UUID
  Global IID_IDataObject As UUID
  Global IID_IDragDropHelper As UUID
  Global IID_IDragSourceHelper As UUID
  Global IID_IDragSourceHelper2 As UUID
  Global IID_IDropTarget As UUID
  Global IID_IDropTargetHelper As UUID
  Global IID_IPersistFolder As UUID
  Global IID_IPersistFolder2 As UUID
  Global IID_IQueryInfo As UUID
  Global IID_IShellFolder As UUID
  Global IID_IShellIcon As UUID
  Global IID_IShellIconOverlay As UUID
  Global IID_IShellLinkA As UUID
  Global IID_IShellLinkW As UUID
  #If Debuging Then
    Global CLSID_AugmentedShellFolder As UUID
    Global CLSID_AugmentedShellFolder2 As UUID
    Global IID_IAugmentedShellFolder As UUID
    Global IID_IAugmentedShellFolder2 As UUID
    Global IID_IDelegateFolder As UUID
    Global IID_IEnumUICommand As UUID
    Global IID_IShellFolder2 As UUID
    Global IID_IUICommand As UUID
    Global IID_IUICommandTarget As UUID
    Global IID_IUIElement As UUID
    Global IID_IThumbnailProvider As UUID
  #End If

  Global DEFICON_BLANKDOC_SMALL As Long
  Global DEFICON_BLANKDOC_LARGE As Long
  Global DEFICON_DOC_SMALL As Long
  Global DEFICON_DOC_LARGE As Long
  Global DEFICON_APP_SMALL As Long
  Global DEFICON_APP_LARGE As Long
  Global DEFICON_FOLDER_SMALL As Long
  Global DEFICON_FOLDER_LARGE As Long
  Global DEFICON_OPENFOLDER_SMALL As Long
  Global DEFICON_OPENFOLDER_LARGE As Long
  Global OVERLAY_LINK As Long
  Global OVERLAY_SHARE As Long
  Global OVERLAY_SLOWFILE As Long


  ' lokale APIs

  Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
  Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
  Private Declare Function ConvertSidToStringSid Lib "advapi32.dll" Alias "ConvertSidToStringSidA" (sid As Byte, pStringSid As Long) As Long
  Private Declare Function CreateFileAsLong Lib "kernel32.dll" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
  Private Declare Function DeviceIoControl Lib "kernel32.dll" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, ByVal lpInBuffer As Long, ByVal nInBufferSize As Long, ByVal lpOutBuffer As Long, ByVal nOutBufferSize As Long, ByRef lpBytesReturned As Long, ByVal lpOverlapped As Long) As Long
  Private Declare Function GetDriveType Lib "kernel32.dll" Alias "GetDriveTypeA" (ByVal Drive As String) As Long
  Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal File As String) As Long
  Private Declare Function GetLogicalDriveStrings Lib "kernel32.dll" Alias "GetLogicalDriveStringsA" (ByVal bufferSize As Long, ByVal buffer As String) As Long
  Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long
  Private Declare Function GetVolumeInformation Lib "kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, ByRef lpVolumeSerialNumber As Long, ByRef lpMaximumComponentLength As Long, ByRef lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
  Private Declare Function ILCloneFirst Lib "shell32.dll" Alias "#19" (ByVal pIDL As Long) As Long
  Private Declare Function ILCreateFromPath Lib "shell32.dll" Alias "#157" (ByVal path As String) As Long
  Private Declare Function ILCreateFromPathAsLong Lib "shell32.dll" Alias "#157" (ByVal pPath As Long) As Long
  Private Declare Sub ILFree Lib "shell32.dll" Alias "#155" (ByVal pMem As Long)
  Private Declare Function ILGetNext Lib "shell32" Alias "#153" (ByVal pIDL As Long) As Long
  Private Declare Function ILGetSize Lib "shell32" Alias "#152" (ByVal pIDL As Long) As Long
  Private Declare Function LocalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long
  Private Declare Function LookupAccountName Lib "advapi32.dll" Alias "LookupAccountNameA" (ByVal lpSystemName As Long, ByVal lpAccountName As String, pSID As Byte, cbSid As Long, ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, peUse As Long) As Long
  Private Declare Function lstrcpyn Lib "kernel32.dll" Alias "lstrcpynW" (ByVal lpString1 As Long, ByVal lpString2 As Long, ByVal iMaxLength As Long) As Long
  Private Declare Function OpenPrinterAsLong Lib "winspool.drv" Alias "OpenPrinterA" (ByVal PrinterName As String, hPrinter As Long, ByVal PrinterData As Long) As Long
  Private Declare Function PathFileExists Lib "shlwapi" Alias "PathFileExistsA" (ByVal File As String) As Long
  Private Declare Function PathFindExtension Lib "shlwapi" Alias "PathFindExtensionA" (ByVal File As Long) As Long
  Private Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal path As String) As Long
  Private Declare Function PathIsUNC Lib "shlwapi" Alias "PathIsUNCA" (ByVal path As String) As Long
  Private Declare Function PathIsURL Lib "shlwapi" Alias "PathIsURLA" (ByVal path As String) As Long
  Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
  Private Declare Function RegCreateKeyExAsLong Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal SubKey As String, ByVal Reserved As Long, ByVal KeyClass As String, ByVal Flags As Long, ByVal AccessRightsMask As Long, ByVal SecurityAttributes As Long, hKeyResult As Long, DoneAction As Long) As Long
  Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByRef lpData As Any, ByRef lpcbData As Long) As Long
  Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal ValueName As String, ByVal Reserved As Long, ByVal dwType As Long, Data As Any, ByVal Datasize As Long) As Long
  Private Declare Function SendMessageTimeout Lib "user32.dll" Alias "SendMessageTimeoutA" (ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal fuFlags As Long, ByVal uTimeout As Long, ByRef lpdwResult As Long) As Long
  Private Declare Function SHBindToParent Lib "shell32.dll" (ByVal pIDL As Long, IID As UUID, IFolder As IVBShellFolder, ByVal pIDLRelative As Long) As Long
  Private Declare Function SHGetDataFromIDList Lib "shell32" Alias "SHGetDataFromIDListA" (ByVal IRelative As IVBShellFolder, ByVal pIDLToRelative As Long, ByVal Mode As Long, Data As Any, ByVal Datasize As Long) As Long
  Private Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" (ByVal File As String, ByVal FileAttributes As Long, Data As SHFILEINFO, ByVal Datasize As Long, ByVal Flags As Long) As Long
  Private Declare Function SHGetFileInfoAsLong Lib "shell32" Alias "SHGetFileInfoA" (ByVal pIDL As Long, ByVal FileAttributes As Long, Data As SHFILEINFO, ByVal Datasize As Long, ByVal Flags As Long) As Long
  Private Declare Function SHGetFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal CSIDL As Long, ByVal hToken As Long, ByVal reserviert As Long, pIDLToDesktop As Long) As Long
  Private Declare Function SHGetIconOverlayIndexAsLong Lib "shell32" Alias "SHGetIconOverlayIndexA" (ByVal IconPath As Long, ByVal iconIndex As Long) As Long
  Private Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pIDLToDesktop As Long, ByVal path As String) As Long
  Private Declare Function SHGetRealIDL Lib "shell32.dll" Alias "#98" (ByVal IRelative As IVBShellFolder, ByVal pIDLSimple As Long, ByRef pIDLReal As Long) As Long
  Private Declare Function SHSetValue Lib "shlwapi" Alias "SHSetValueA" (ByVal hKey As Long, ByVal SubKey As String, ByVal Value As String, ByVal DataType As Long, Data As Any, ByVal Datasize As Long) As Long
  'Private Declare Function StrRetToBSTR Lib "shlwapi.dll" (Data As STRRET, ByVal pIDL As Long, ByVal pBuffer As Long) As Long
  Private Declare Function StrRetToBuf Lib "shlwapi.dll" Alias "StrRetToBufA" (Data As STRRET, ByVal pIDL As Long, ByVal buffer As String, ByVal bufferSize As Long) As Long
  Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal Flags As Long, ByVal WideCharStr As Long, ByVal WideCharSize As Long, ByVal MultiByteStr As String, ByVal MultiByteSize As Long, ByVal DefaultChar As String, ByVal UsedDefaultChar As Long) As Long

  ' Win95
  Private Declare Sub SHGetSpecialFolderLocation Lib "shell32" (ByVal hwndOwner As Long, ByVal CSIDL As Long, ByRef pIDL As Long)


  ' globale APIs

  Declare Function CLSIDFromString Lib "ole32.dll" (ByVal pString As Long, CLSID As UUID) As Long
  Declare Function CoTaskMemAlloc Lib "ole32.dll" (ByVal cb As Long) As Long
  Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pMem As Long)
  Declare Function FileIconInit Lib "shell32" Alias "#660" (ByVal FullInit As Long) As Long
  Declare Function GetTickCount Lib "kernel32" () As Long
  Declare Function ILAppendID Lib "shell32.dll" Alias "#154" (ByVal pIDL As Long, ByVal pItemID As Long, ByVal fAppend As Long) As Long
  Declare Function ILClone Lib "shell32.dll" Alias "#18" (ByVal pIDL As Long) As Long
  Declare Function ILFindChild Lib "shell32.dll" Alias "#24" (ByVal pIDLParent As Long, ByVal pIDLChild As Long) As Long
  Declare Function ILFindLastID Lib "shell32.dll" Alias "#16" (ByVal pIDL As Long) As Long
  Declare Function ILIsEqual Lib "shell32.dll" Alias "#21" (ByVal pIDL1 As Long, ByVal pIDL2 As Long) As Long
  Declare Function ILIsParent Lib "shell32.dll" Alias "#23" (ByVal pIDLParent As Long, ByVal pIDLBelow As Long, ByVal fImmediate As Long) As Long
  Declare Function ILRemoveLastID Lib "shell32.dll" Alias "#17" (ByVal pIDL As Long) As Long
  Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByVal pDest As Long, ByVal Length As Long)


' globale Methoden

' hängt <DispNames> an <pIDLToDesktop> an
#If Debuging Then
  Function AddDisplayNamesTopIDL(debugger As clsDebugger, ByVal pIDLToDesktop As Long, ByVal DispNames As String, Optional ByVal returnOnlyExactMatch As Boolean = False) As Long
#Else
  Function AddDisplayNamesTopIDL(ByVal pIDLToDesktop As Long, ByVal DispNames As String, Optional ByVal returnOnlyExactMatch As Boolean = False) As Long
#End If
  Dim i As Integer
  Dim IParent As IVBShellFolder
  Dim pIDL As Long
  Dim Segment As String
  Dim tmp As Long

  If pIDLToDesktop = 0 Then Exit Function

  For i = 1 To CountSegments(DispNames)
    #If Debuging Then
      debugger.AddLogEntry "AddDisplayNamesTopIDL: Calling GetISHFolderInterfaceFQ()", LogEntryTypeConstants.letInfo
      debugger.AddLogEntry "   0x" & Hex(pIDLToDesktop), LogEntryTypeConstants.letOther
      Set IParent = GetISHFolderInterfaceFQ(debugger, pIDLToDesktop)
    #Else
      Set IParent = GetISHFolderInterfaceFQ(pIDLToDesktop)
    #End If

    Segment = GetFirstFolders(DispNames, i)
    Segment = GetLastFolders(Segment, 1)

    #If Debuging Then
      pIDL = FilterSubItems(debugger, IParent, pIDLToDesktop, Segment, True)
    #Else
      pIDL = FilterSubItems(IParent, pIDLToDesktop, Segment, True)
    #End If
    Set IParent = Nothing
    If pIDL Then
      tmp = pIDLToDesktop
      If pIDL Then
        pIDLToDesktop = ILAppendID(ILClone(pIDLToDesktop), pIDL, 1)
      Else
        pIDLToDesktop = 0
      End If
      #If Debuging Then
        FreeItemIDList debugger, "AddDisplayNamesTopIDL #1", pIDL
        FreeItemIDList debugger, "AddDisplayNamesTopIDL #2", tmp
      #Else
        FreeItemIDList pIDL
        FreeItemIDList tmp
      #End If
    Else
      If returnOnlyExactMatch Then
        #If Debuging Then
          FreeItemIDList debugger, "AddDisplayNamesTopIDL #3", pIDLToDesktop
        #Else
          FreeItemIDList pIDLToDesktop
        #End If
      End If
      Exit For
    End If
  Next

  AddDisplayNamesTopIDL = pIDLToDesktop
End Function

' prüft, ob <pIDLToParent> umbenannt werden kann
#If Debuging Then
  Function CanBeRenamed(debugger As clsDebugger, IParent As IVBShellFolder, pIDLToParent As Long) As Boolean
#Else
  Function CanBeRenamed(IParent As IVBShellFolder, pIDLToParent As Long) As Boolean
#End If
  Dim ret As Boolean

  If IParent Is Nothing Then Exit Function
  If pIDLToParent = 0 Then Exit Function

  ret = HasAttribute(IParent, pIDLToParent, SFGAOConstants.SFGAO_CANRENAME)
  If ret Then
    ' WORKAROUND: den Desktop kann man lt. Windows umbenennen
    #If Debuging Then
      If IsDesktop(debugger, IParent, pIDLToParent) Then
    #Else
      If IsDesktop(IParent, pIDLToParent) Then
    #End If
      ret = True
    End If
  End If

  CanBeRenamed = ret
End Function

' wandelt <CLSID> in einen DisplayName um
Function CLSIDToDisplayName(ByVal CLSID As String) As String
  Dim ret As String

  ' zunächst versuchen, den DisplayName unter "HKEY_CURRENT_USER\Software\Microsoft\Windows\
  ' CurrentVersion\Explorer\CLSID\" zu bekommen
  ret = getRegDefaultValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\" & Mid$(CLSID, 3))
  If ret = "" Then ret = getRegDefaultValue(HKEY_CLASSES_ROOT, "CLSID\" & Mid$(CLSID, 3))

  CLSIDToDisplayName = ret
End Function

#If Debuging Then
  Function CopyFirstItemIDs(debugger As clsDebugger, pIDL As Long, Optional ByVal Count As Integer = 1, Optional ByVal freepIDL As Boolean = False) As Long
#Else
  Function CopyFirstItemIDs(pIDL As Long, Optional ByVal Count As Integer = 1, Optional ByVal freepIDL As Boolean = False) As Long
#End If
  Dim c As Long
  Dim i As Long
  Dim newpIDL As Long

  #If Debuging Then
    If pIDL = 0 Then
      debugger.AddLogEntry "CopyFirstItemIDs: pIDL = 0 - leaving function", LogEntryTypeConstants.letWarning
      Exit Function
    End If
    If Count < 0 Then
      debugger.AddLogEntry "CopyFirstItemIDs: Count < 0 - handling it as 0", LogEntryTypeConstants.letWarning
    End If
  #End If
  If Count <= 0 Then GoTo FreeMem

  If Count = 1 Then
    CopyFirstItemIDs = ILCloneFirst(pIDL)
    GoTo FreeMem
  End If

  newpIDL = ILClone(pIDL)
  #If Debuging Then
    c = CountItemIDs(debugger, newpIDL)
  #Else
    c = CountItemIDs(newpIDL)
  #End If
  For i = 1 To c - Count
    ILRemoveLastID newpIDL
  Next i
  CopyFirstItemIDs = newpIDL

FreeMem:
  If freepIDL Then
    #If Debuging Then
      FreeItemIDList debugger, "CopyFirstItemIDs", pIDL
    #Else
      FreeItemIDList pIDL
    #End If
  End If
End Function

#If Debuging Then
  Function CopyItemIDList(debugger As clsDebugger, pIDL As Long, Optional ByVal freepIDL As Boolean = False) As Long
#Else
  Function CopyItemIDList(pIDL As Long, Optional ByVal freepIDL As Boolean = False) As Long
#End If
  #If Debuging Then
    If pIDL = 0 Then
      debugger.AddLogEntry "CopyItemIDList: pIDL = 0 - leaving function", LogEntryTypeConstants.letWarning
      Exit Function
    End If
  #End If

  CopyItemIDList = ILClone(pIDL)

FreeMem:
  If freepIDL Then
    #If Debuging Then
      FreeItemIDList debugger, "CopyItemIDList", pIDL
    #Else
      FreeItemIDList pIDL
    #End If
  End If
End Function

' kopiert die letzten <Count> ItemIDs von <pIDL>
#If Debuging Then
  Function CopyLastItemIDs(debugger As clsDebugger, pIDL As Long, Optional ByVal Count As Integer = 1, Optional ByVal freepIDL As Boolean = False) As Long
    Dim p As Long

    debugger.AddLogEntry "CopyLastItemIDs: Calling GetLastItemIDs()", LogEntryTypeConstants.letInfo
    debugger.AddLogEntry "   0x" & Hex(pIDL), LogEntryTypeConstants.letOther
    p = GetLastItemIDs(debugger, pIDL, Count)
    If p Then
      CopyLastItemIDs = CopyItemIDList(debugger, p)
    End If
    If freepIDL Then
      FreeItemIDList debugger, "CopyLastItemIDs", pIDL
    End If
  End Function
#Else
  Function CopyLastItemIDs(pIDL As Long, Optional ByVal Count As Integer = 1, Optional ByVal freepIDL As Boolean = False) As Long
    Dim p As Long

    p = GetLastItemIDs(pIDL, Count)
    If p Then
      CopyLastItemIDs = CopyItemIDList(p)
    End If
    If freepIDL Then
      FreeItemIDList pIDL
    End If
  End Function
#End If

#If Debuging Then
  Function CountItemIDs(debugger As clsDebugger, ByVal pIDL As Long) As Long
#Else
  Function CountItemIDs(ByVal pIDL As Long) As Long
#End If
  Dim i As Long

  #If Debuging Then
    If pIDL = 0 Then
      debugger.AddLogEntry "CountItemIDs: pIDL = 0 - leaving function", LogEntryTypeConstants.letWarning
      Exit Function
    End If
  #End If

  ' the final iteration will find the 2 NULL bytes
  i = -1
  While pIDL
    i = i + 1
    pIDL = ILGetNext(pIDL)
  Wend

  CountItemIDs = i
End Function

' konvertiert <CSIDL> in einen MenuItem
Function CSIDLToMenuItem(ByVal CSIDL As CSIDLConstants) As String
  Dim ret As String

  Select Case CSIDL
    Case CSIDLConstants.CSIDL_ADMINTOOLS
      ret = "001Administrator-Tools"
    Case CSIDLConstants.CSIDL_ALTSTARTUP
      ret = "004alternativer Autostart"
    Case CSIDLConstants.CSIDL_APPDATA
      ret = "006Anwendungsdaten"
    Case CSIDLConstants.CSIDL_BITBUCKET
      ret = "043Papierkorb"
    Case CSIDLConstants.CSIDL_CDBURN_AREA
      ret = "012CD Burning"
    Case CSIDLConstants.CSIDL_COMMON_ADMINTOOLS
      ret = "002Administrator-Tools (All Users)"
    Case CSIDLConstants.CSIDL_COMMON_ALTSTARTUP
      ret = "005alternativer Autostart (All Users)"
    Case CSIDLConstants.CSIDL_COMMON_APPDATA
      ret = "007Anwendungsdaten (All Users)"
    Case CSIDLConstants.CSIDL_COMMON_DESKTOPDIRECTORY
      ret = "016Desktop (Ordner, All Users)"
    Case CSIDLConstants.CSIDL_COMMON_DOCUMENTS
      ret = "018Dokumente (All Users)"
    Case CSIDLConstants.CSIDL_COMMON_FAVORITES
      ret = "032Favoriten (All Users)"
    Case CSIDLConstants.CSIDL_COMMON_MUSIC
      ret = "028Eigene Musik (All Users)"
    Case CSIDLConstants.CSIDL_COMMON_OEM_LINKS
      ret = "042OEM-Software-Links (All Users)"
    Case CSIDLConstants.CSIDL_COMMON_PICTURES
      ret = "024Eigene Bilder (All Users)"
    Case CSIDLConstants.CSIDL_COMMON_PROGRAMS
      ret = "047Programme (Startmenü, All Users)"
    Case CSIDLConstants.CSIDL_COMMON_STARTMENU
      ret = "051Startmenü (All Users)"
    Case CSIDLConstants.CSIDL_COMMON_STARTUP
      ret = "011Autostart (All Users)"
    Case CSIDLConstants.CSIDL_COMMON_TEMPLATES
      ret = "020Dokumentvorlagen (All Users)"
    Case CSIDLConstants.CSIDL_COMMON_VIDEO
      ret = "030Eigene Videos (All Users)"
    Case CSIDLConstants.CSIDL_COMPUTERSNEARME
      ret = "008Arbeitsgruppencomputer"
    Case CSIDLConstants.CSIDL_CONNECTIONS
      ret = "041Netzwerk- und DFÜ-Verbindungen"
    Case CSIDLConstants.CSIDL_CONTROLS
      ret = "054Systemsteuerung"
    Case CSIDLConstants.CSIDL_COOKIES
      ret = "013Cookies"
    Case CSIDLConstants.CSIDL_DESKTOP
      ret = "014Desktop"
    Case CSIDLConstants.CSIDL_DESKTOPDIRECTORY
      ret = "015Desktop (Ordner)"
    Case CSIDLConstants.CSIDL_DRIVES
      ret = "009Arbeitsplatz"
    Case CSIDLConstants.CSIDL_FAVORITES
      ret = "031Favoriten"
    Case CSIDLConstants.CSIDL_FONTS
      ret = "048Schriftarten"
    Case CSIDLConstants.CSIDL_HISTORY
      ret = "056Verlauf"
    Case CSIDLConstants.CSIDL_INTERNET
      ret = "035Internet"
    Case CSIDLConstants.CSIDL_INTERNET_CACHE
      ret = "036Internet-Cache"
    Case CSIDLConstants.CSIDL_LOCAL_APPDATA
      ret = "037lokale Anwendungsdaten"
    Case CSIDLConstants.CSIDL_MYDOCUMENTS
      ret = "025Eigene Dateien"
    Case CSIDLConstants.CSIDL_MYMUSIC
      ret = "027Eigene Musik"
    Case CSIDLConstants.CSIDL_MYPICTURES
      ret = "023Eigene Bilder"
    Case CSIDLConstants.CSIDL_MYVIDEO
      ret = "029Eigene Videos"
    Case CSIDLConstants.CSIDL_NETHOOD
      ret = "040Netzwerkumgebung (Ordner)"
    Case CSIDLConstants.CSIDL_NETWORK
      ret = "039Netzwerkumgebung"
    Case CSIDLConstants.CSIDL_PERSONAL
      ret = "026Eigene Dateien (Ordner)"
    Case CSIDLConstants.CSIDL_PRINTERS
      ret = "021Drucker"
    Case CSIDLConstants.CSIDL_PRINTHOOD
      ret = "022Druckumgebung"
    Case CSIDLConstants.CSIDL_PROFILE
      ret = "003aktuelles Benutzerprofil"
    Case CSIDLConstants.CSIDL_PROGRAM_FILES
      ret = "044Programme"
    Case CSIDLConstants.CSIDL_PROGRAM_FILES_COMMON
      ret = "033Gemeinsame Dateien"
    Case CSIDLConstants.CSIDL_PROGRAM_FILES_COMMONX86
      ret = "034Gemeinsame Dateien (RISC-Systeme)"
    Case CSIDLConstants.CSIDL_PROGRAM_FILESX86
      ret = "045Programme (RISC-Systeme)"
    Case CSIDLConstants.CSIDL_PROGRAMS
      ret = "046Programme (Startmenü)"
    Case CSIDLConstants.CSIDL_RECENT
      ret = "017Dokumente"
    Case CSIDLConstants.CSIDL_RESOURCES
      ret = "055System-Resourcen (Themes etc.)"
    Case CSIDLConstants.CSIDL_RESOURCES_LOCALIZED
      ret = "038lokalisierte System-Resourcen (Themes etc.)"
    Case CSIDLConstants.CSIDL_SENDTO
      ret = "049Senden an"
    Case CSIDLConstants.CSIDL_STARTMENU
      ret = "050Startmenü"
    Case CSIDLConstants.CSIDL_STARTUP
      ret = "010Autostart"
    Case CSIDLConstants.CSIDL_SYSTEM
      ret = "052System"
    Case CSIDLConstants.CSIDL_SYSTEMX86
      ret = "053System (RISC-Systeme)"
    Case CSIDLConstants.CSIDL_TEMPLATES
      ret = "019Dokumentvorlagen"
    Case CSIDLConstants.CSIDL_WINDOWS
      ret = "057Windows"
  End Select

  CSIDLToMenuItem = ret
End Function

Function CSIDLTopIDL(ByVal CSIDL As CSIDLConstants) As Long
  Dim pIDLToDesktop As Long

  If ver_Shell32_50 Then
    SHGetFolderLocation 0, CSIDL, 0, 0, pIDLToDesktop
  Else
    SHGetSpecialFolderLocation 0, CSIDL, pIDLToDesktop
  End If

  CSIDLTopIDL = pIDLToDesktop
End Function

' ermittelt die pIDL von <ParsingName>
' die pIDL ist relativ zu <IRelative>
Function DisplayNameTopIDL(IRelative As IVBShellFolder, ParsingName As String) As Long
  Dim ret As Long

  If IRelative Is Nothing Then Exit Function

  IRelative.ParseDisplayName 0, 0, ParsingName, Len(ParsingName), ret, 0

  DisplayNameTopIDL = ret
End Function

' gibt den Index des Icons für <pIDLToParent> zurück
#If Debuging Then
  Function FastGetSysIconIndex(debugger As clsDebugger, ISHIcon As IVBShellIcon, IParent As IVBShellFolder, pIDLToParent As Long, pIDLToDesktop As Long, Selected As Boolean, StandardIcons As Boolean, Optional LargeIcons As Boolean = False) As Long
#Else
  Function FastGetSysIconIndex(ISHIcon As IVBShellIcon, IParent As IVBShellFolder, pIDLToParent As Long, pIDLToDesktop As Long, Selected As Boolean, StandardIcons As Boolean, Optional LargeIcons As Boolean = False) As Long
#End If
  Dim ret As Long

  If LargeIcons Or StandardIcons Then
    ' we can't use IShellIcon here
    #If Debuging Then
      ret = GetSysIconIndex(debugger, IParent, pIDLToParent, pIDLToDesktop, Selected, StandardIcons, LargeIcons)
    #Else
      ret = GetSysIconIndex(IParent, pIDLToParent, pIDLToDesktop, Selected, StandardIcons, LargeIcons)
    #End If
  Else
    If Selected Then
      If ISHIcon.GetIconOf(pIDLToParent, GILConstants.GIL_FORSHELL Or GILConstants.GIL_OPENICON, ret) <> NOERROR Then
        ' fall back to SHGetFileInfo
        #If Debuging Then
          ret = GetSysIconIndex(debugger, IParent, pIDLToParent, pIDLToDesktop, Selected, StandardIcons, LargeIcons)
        #Else
          ret = GetSysIconIndex(IParent, pIDLToParent, pIDLToDesktop, Selected, StandardIcons, LargeIcons)
        #End If
      End If
    Else
      If ISHIcon.GetIconOf(pIDLToParent, GILConstants.GIL_FORSHELL, ret) <> NOERROR Then
        ' fall back to SHGetFileInfo
        #If Debuging Then
          ret = GetSysIconIndex(debugger, IParent, pIDLToParent, pIDLToDesktop, Selected, StandardIcons, LargeIcons)
        #Else
          ret = GetSysIconIndex(IParent, pIDLToParent, pIDLToDesktop, Selected, StandardIcons, LargeIcons)
        #End If
      End If
    End If
  End If

  FastGetSysIconIndex = ret
End Function

' prüft, ob <File> existiert
Function FileExists(ByVal File As String) As Boolean
  ' <PathFileExists> macht keinen Unterschied zwischen Ordner und Datei
  If IsDirectory(File) Then Exit Function

  FileExists = PathFileExists(File)
End Function

' prüft, ob <pIDLToRelative> existiert
#If Debuging Then
  Function FileExists_pIDL(debugger As clsDebugger, IRelative As IVBShellFolder, pIDLToRelative As Long, Optional ByVal freepIDL As Boolean = False) As Boolean
    If IsDrive(debugger, IRelative, pIDLToRelative) Then
      ' damit verhindern wir unnütze Laufwerkszugriffe
      If freepIDL Then
        FreeItemIDList debugger, "FileExists_pIDL", pIDLToRelative
      End If
    Else
      FileExists_pIDL = FileExists(pIDLToPath(debugger, IRelative, pIDLToRelative, freepIDL))
    End If
  End Function
#Else
  Function FileExists_pIDL(IRelative As IVBShellFolder, pIDLToRelative As Long, Optional ByVal freepIDL As Boolean = False) As Boolean
    If IsDrive(IRelative, pIDLToRelative) Then
      ' damit verhindern wir unnütze Laufwerkszugriffe
      If freepIDL Then
        FreeItemIDList pIDLToRelative
      End If
    Else
      FileExists_pIDL = FileExists(pIDLToPath(IRelative, pIDLToRelative, freepIDL))
    End If
  End Function
#End If

' sucht unter den SubItems von <IParent> den heraus, der den Suchkriterien entspricht und gibt
' dessen pIDL (relativ zu <IParent>) zurück
' <isRelativeToParent> gibt an, ob <SearchFor> relativ zum Parent-Item ist
#If Debuging Then
  Function FilterSubItems(debugger As clsDebugger, IParent As IVBShellFolder, ByVal pIDL_Parent_ToDesktop As Long, ByVal SearchFor As String, ByVal isDispName As Boolean) As Long
#Else
  Function FilterSubItems(IParent As IVBShellFolder, ByVal pIDL_Parent_ToDesktop As Long, ByVal SearchFor As String, ByVal isDispName As Boolean) As Long
#End If
  Dim EnumFlags As SHCONTFConstants
  Dim IEnum As IVBEnumIDList
  Dim pIDLSubItem_ToParent As Long
  Dim ret As Long
  Dim txt As String

  If IParent Is Nothing Then Exit Function
  If pIDL_Parent_ToDesktop = 0 Then Exit Function

  If isDispName Then
    EnumFlags = SHCONTFConstants.SHCONTF_FOLDERS Or SHCONTFConstants.SHCONTF_NONFOLDERS Or SHCONTFConstants.SHCONTF_INCLUDEHIDDEN
    IParent.EnumObjects 0, EnumFlags, IEnum

    If Not (IEnum Is Nothing) Then
      While (IEnum.Next(1, pIDLSubItem_ToParent, 0) = S_OK) And (ret = 0)
        #If Debuging Then
          txt = pIDLToDisplayName(debugger, IParent, pIDLSubItem_ToParent, SHGDNConstants.SHGDN_INFOLDER)
        #Else
          txt = pIDLToDisplayName(IParent, pIDLSubItem_ToParent, SHGDNConstants.SHGDN_INFOLDER)
        #End If
        If LCase$(txt) = LCase$(SearchFor) Then
          ret = pIDLSubItem_ToParent
        Else
          #If Debuging Then
            FreeItemIDList debugger, "FilterSubItems", pIDLSubItem_ToParent
          #Else
            FreeItemIDList pIDLSubItem_ToParent
          #End If
        End If
      Wend
    End If
    Set IEnum = Nothing
  Else
    IParent.ParseDisplayName 0, 0, SearchFor, Len(SearchFor), ret, 0
  End If

  FilterSubItems = ret
End Function

' sucht unter den SubItems von <IParent> den heraus, der den Suchkriterien entspricht und gibt
' dessen pIDL (relativ zu <IParent>) zurück
' <isRelativeToParent> gibt an, ob <pIDL_SearchFor> relativ zum Parent-Item ist
#If Debuging Then
  Function FilterSubItems_pIDL(debugger As clsDebugger, ByVal IParent As IVBShellFolder, pIDLParent_ToDesktop As Long, pIDL_SearchFor As Long, ByVal isRelativeToParent As Boolean, Optional ByVal freepIDLs As Boolean = True) As Long
#Else
  Function FilterSubItems_pIDL(ByVal IParent As IVBShellFolder, pIDLParent_ToDesktop As Long, pIDL_SearchFor As Long, ByVal isRelativeToParent As Boolean, Optional ByVal freepIDLs As Boolean = True) As Long
#End If
  Dim EnumFlags As SHCONTFConstants
  Dim IEnum As IVBEnumIDList
  Dim pIDLSubItem_ToDesktop As Long
  Dim pIDLSubItem_ToParent As Long

  If IParent Is Nothing Then GoTo FreeMem
  If pIDLParent_ToDesktop = 0 Then GoTo FreeMem
  If pIDL_SearchFor = 0 Then GoTo FreeMem

  EnumFlags = SHCONTFConstants.SHCONTF_FOLDERS Or SHCONTFConstants.SHCONTF_NONFOLDERS Or SHCONTFConstants.SHCONTF_INCLUDEHIDDEN
  IParent.EnumObjects 0, EnumFlags, IEnum
  If Not (IEnum Is Nothing) Then
    Do While IEnum.Next(1, pIDLSubItem_ToParent, 0) = S_OK
      If isRelativeToParent Then
        If ILIsEqual(pIDLSubItem_ToParent, pIDL_SearchFor) Then
          FilterSubItems_pIDL = pIDLSubItem_ToParent
          Exit Do
        End If
      ElseIf pIDLSubItem_ToParent Then
        pIDLSubItem_ToDesktop = ILAppendID(ILClone(pIDLParent_ToDesktop), pIDLSubItem_ToParent, 1)
        If ILIsEqual(pIDLSubItem_ToDesktop, pIDL_SearchFor) Then
          FilterSubItems_pIDL = pIDLSubItem_ToParent
          #If Debuging Then
            FreeItemIDList debugger, "FilterSubItems_pIDL #1", pIDLSubItem_ToDesktop
          #Else
            FreeItemIDList pIDLSubItem_ToDesktop
          #End If
          Exit Do
        End If
        #If Debuging Then
          FreeItemIDList debugger, "FilterSubItems_pIDL #2", pIDLSubItem_ToDesktop
        #Else
          FreeItemIDList pIDLSubItem_ToDesktop
        #End If
      End If
      #If Debuging Then
        FreeItemIDList debugger, "FilterSubItems_pIDL #3", pIDLSubItem_ToParent
      #Else
        FreeItemIDList pIDLSubItem_ToParent
      #End If
    Loop
  End If
  Set IEnum = Nothing

FreeMem:
  If freepIDLs Then
    #If Debuging Then
      FreeItemIDList debugger, "FilterSubItems_pIDL #4", pIDLParent_ToDesktop
      FreeItemIDList debugger, "FilterSubItems_pIDL #5", pIDL_SearchFor
    #Else
      FreeItemIDList pIDLParent_ToDesktop
      FreeItemIDList pIDL_SearchFor
    #End If
  End If
End Function

#If Debuging Then
  Sub FreeItemIDList(debugger As clsDebugger, callingFunction As String, ByRef pIDL As Long)
#Else
  Sub FreeItemIDList(ByRef pIDL As Long)
#End If
  #If Debuging Then
    If pIDL = 0 Then
      debugger.AddLogEntry "FreeItemIDList (called from " & callingFunction & "): pIDL = 0 - leaving sub", LogEntryTypeConstants.letWarning
      Exit Sub
    Else
      #If LogFreeItemIDList Then
        debugger.AddLogEntry "FreeItemIDList (called from " & callingFunction & ")", LogEntryTypeConstants.letInfo
        debugger.AddLogEntry "   0x" & Hex(pIDL), LogEntryTypeConstants.letOther
      #End If
    End If
  #End If

  If ver_Win_2k Then
    CoTaskMemFree pIDL
  Else
    ILFree pIDL
  End If
  pIDL = 0
End Sub

' gibt alle Festplatten-Laufwerke zurück
' Form: "C:\|D:\"
Function GetAllHardDisks() As String
  Dim allDrives() As String
  Dim buffer As String
  Dim i As Integer
  Dim ret As String

  buffer = Space$(MAX_PATH)
  GetLogicalDriveStrings Len(buffer), buffer
  allDrives = Split(buffer, Chr$(0))

  For i = LBound(allDrives) To UBound(allDrives)
    If GetDriveType(allDrives(i)) = DRIVE_FIXED Then ret = ret & AddBackslash(allDrives(i)) & "|"
  Next
  If Right$(ret, 1) = "|" Then ret = Left$(ret, Len(ret) - 1)

  GetAllHardDisks = ret
End Function

' gibt alle Papierkörbe zurück
' Form: "C:\Recycler\S-1-...|C:\Recycler\S-1-...|D:\Recycler\S-1-...|D:\Recycler\S-1-..."
#If Debuging Then
  Function GetAllRecycleBins(debugger As clsDebugger) As String
#Else
  Function GetAllRecycleBins() As String
#End If
  Dim allHDs() As String
  Dim buffer As String
  Dim bUserSID() As Byte
  Dim i As Integer
  Dim pSID As Long
  Dim ret As String
  Dim SIDType As Long
  Dim txt As String
  Dim userName As String
  Dim userSID As String

  userName = String$(MAX_PATH, Chr$(0))
  GetUserName userName, MAX_PATH
  If ver_Win_NTBased Then
    ReDim bUserSID(1 To MAX_PATH) As Byte
    buffer = String$(MAX_PATH, Chr$(0))
    LookupAccountName 0, userName, bUserSID(1), MAX_PATH, buffer, MAX_PATH, SIDType
    If ver_Win_2k Then
      ConvertSidToStringSid bUserSID(1), pSID
      userSID = GetStrFromPointer(pSID)
      LocalFree pSID
    End If
    ' problem: what do we do on NT4?
  End If

  allHDs = Split(GetAllHardDisks, "|")
  For i = LBound(allHDs) To UBound(allHDs)
    ' alle Papierkörbe auf diesem Laufwerk suchen
    ' dazu alle SubItems von "<HD>\recycler\" aufzählen
    txt = UCase$(GetFileSystem(allHDs(i)))
    buffer = ""
    If txt = "NTFS" Then
      If ver_Win_Vista Then
        buffer = allHDs(i) & "$recycle.bin\" & userSID
      Else
        buffer = allHDs(i) & "recycler\" & userSID
      End If
    ElseIf txt Like "FAT*" Then
      If ver_Win_Vista Then
        buffer = allHDs(i) & "$recycle.bin\"
      Else
        buffer = allHDs(i) & "recycled\"
      End If
    ElseIf txt = "UDF" Then
      If ver_Win_Vista Then
        buffer = allHDs(i) & "$recycle.bin\"
      #If Debuging Then
        Else
          debugger.AddLogEntry "GetAllRecycleBins() - UDF hard disk on a pre-Vista system!", LogEntryTypeConstants.letWarning
          debugger.AddLogEntry "   " & allHDs(i), LogEntryTypeConstants.letOther
      #End If
      End If
    ElseIf txt Like "EXT*" Then
      If ver_Win_Vista Then
        buffer = allHDs(i) & "$recycle.bin\"
      Else
        buffer = allHDs(i) & "recycled\"
      End If
    #If Debuging Then
      Else
        debugger.AddLogEntry "GetAllRecycleBins() - unsupported filesystem", LogEntryTypeConstants.letWarning
        debugger.AddLogEntry "   " & allHDs(i), LogEntryTypeConstants.letOther
        debugger.AddLogEntry "   " & txt, LogEntryTypeConstants.letOther
    #End If
    End If
    txt = ""
    If IsDirectory(buffer) Then txt = buffer & "|"

    ret = ret & txt
  Next
  If Right$(ret, 1) = "|" Then ret = Left$(ret, Len(ret) - 1)

  GetAllRecycleBins = ret
End Function

#If Debuging Then
  Function GetAttributes(debugger As clsDebugger, IParent As IVBShellFolder, pIDLToParent As Long, Mask As SFGAOConstants) As Long
#Else
  Function GetAttributes(IParent As IVBShellFolder, pIDLToParent As Long, Mask As SFGAOConstants) As Long
#End If
  #If Debuging Then
    If IParent Is Nothing Then
      debugger.AddLogEntry "GetAttributes: IParent = Nothing - leaving function", LogEntryTypeConstants.letWarning
      Exit Function
    End If
    If pIDLToParent = 0 Then
      debugger.AddLogEntry "GetAttributes: pIDLToParent = 0 - leaving function", LogEntryTypeConstants.letWarning
      Exit Function
    End If
  #End If

  IParent.GetAttributesOf 1, pIDLToParent, Mask
  GetAttributes = Mask
End Function

#If Debuging Then
  ' ermittelt den DisplayName des Ordners <CSIDL>
  Function GetCSIDLDisplayName(debugger As clsDebugger, ByVal CSIDL As CSIDLConstants) As String
    GetCSIDLDisplayName = pIDLToDisplayName_Light(debugger, CSIDLTopIDL(CSIDL), True)
  End Function
#Else
  ' ermittelt den DisplayName des Ordners <CSIDL>
  Function GetCSIDLDisplayName(ByVal CSIDL As CSIDLConstants) As String
    GetCSIDLDisplayName = pIDLToDisplayName_Light(CSIDLTopIDL(CSIDL), True)
  End Function
#End If

Sub GetDefaultOverlays()
  If ver_Shell32_50 Then
    OVERLAY_LINK = SHGetIconOverlayIndexAsLong(0, IDO_SHGIOI_LINK)
    OVERLAY_SHARE = SHGetIconOverlayIndexAsLong(0, IDO_SHGIOI_SHARE)
    OVERLAY_SLOWFILE = SHGetIconOverlayIndexAsLong(0, IDO_SHGIOI_SLOWFILE)
  Else
    ' Standardwerte nutzen
    OVERLAY_LINK = 2
    OVERLAY_SHARE = 1
    OVERLAY_SLOWFILE = 4
  End If
End Sub

#If Debuging Then
  Function GetFileAttribs(debugger As clsDebugger, IParent As IVBShellFolder, pIDLToParent As Long, Mask As Long) As Long
#Else
  Function GetFileAttribs(IParent As IVBShellFolder, pIDLToParent As Long, Mask As Long) As Long
#End If
  Dim l As Long
  Dim path As String
  Dim ret As Long
  Dim ret2 As Long

  If IsPartOfFileSystem(IParent, pIDLToParent) Then
    #If Debuging Then
      path = pIDLToPath(debugger, IParent, pIDLToParent)
    #Else
      path = pIDLToPath(IParent, pIDLToParent)
    #End If
    If path <> "" Then
      ' HACK: Für die Eigenen Dateien auf dem Desktop gibt pIDLToPath() den realen Pfad (bspw. D:\)
      '       zurück. Die Eigenen Dateien haben eigentlich keine Attribute, der reale Pfad aber schon.
      l = 0
      If IParent Is IDesktop Then
        l = Mask
      #If Debuging Then
        ElseIf IsMyComputer(debugger, IParent) Then
      #Else
        ElseIf IsMyComputer(IParent) Then
      #End If
        l = Mask
      End If
      If l Then
        l = 0
        If CountSegments(path) <= 1 Then
          If GetDriveType(Left(path, 2)) > 1 Then l = Mask
        End If
      End If

      If l Then
        ret2 = -1
      Else
        ret2 = GetFileAttributes(path)
      End If
      If ret2 <> -1 Then ret = (ret2 And Mask)
    End If
  End If

  GetFileAttribs = ret
End Function

' gibt die Dateiendung von <File> zurück
Function GetFileNameExtension(ByVal File As String) As String
  Dim pExt As Long

  File = StrConv(File, VbStrConv.vbFromUnicode)
  pExt = PathFindExtension(StrPtr(File))
  GetFileNameExtension = Mid$(GetStrFromPointer(pExt), 2)
End Function

' gibt die Dateiendung von <pIDLToParent> zurück
#If Debuging Then
  Function GetFileNameExtension_pIDL(debugger As clsDebugger, IParent As IVBShellFolder, pIDLToParent As Long, Optional ByVal skipCheckForFile As Boolean = False)
#Else
  Function GetFileNameExtension_pIDL(IParent As IVBShellFolder, pIDLToParent As Long, Optional ByVal skipCheckForFile As Boolean = False)
#End If
  Dim path As String
  Dim ret As String

  If IParent Is Nothing Then Exit Function
  If pIDLToParent = 0 Then Exit Function

  If skipCheckForFile Then
    #If Debuging Then
      path = pIDLToPath(debugger, IParent, pIDLToParent)
    #Else
      path = pIDLToPath(IParent, pIDLToParent)
    #End If
  Else
    #If Debuging Then
      If IsFSFolder(IParent, pIDLToParent) Then
        ' ist es wirklich ein Ordner?
        If FileExists_pIDL(debugger, IParent, pIDLToParent) Then
          path = pIDLToPath(debugger, IParent, pIDLToParent)
        End If
      ElseIf IsFSFile(debugger, IParent, pIDLToParent) Then
        path = pIDLToPath(debugger, IParent, pIDLToParent)
      End If
    #Else
      If IsFSFolder(IParent, pIDLToParent) Then
        ' ist es wirklich ein Ordner?
        If FileExists_pIDL(IParent, pIDLToParent) Then
          path = pIDLToPath(IParent, pIDLToParent)
        End If
      ElseIf IsFSFile(IParent, pIDLToParent) Then
        path = pIDLToPath(IParent, pIDLToParent)
      End If
    #End If
  End If

  If path <> "" Then
    ret = GetFileNameExtension(path)
  End If

  GetFileNameExtension_pIDL = ret
End Function

Function GetFileSystem(strDrive As String) As String
  Dim fsFlags As Long
  Dim fsName As String
  Dim maxComponentLength As Long
  Dim volName As String

  volName = String$(MAX_PATH, Chr$(0))
  fsName = String$(MAX_PATH, Chr$(0))
  GetVolumeInformation strDrive, volName, MAX_PATH, 0, maxComponentLength, fsFlags, fsName, MAX_PATH
  GetFileSystem = Left$(fsName, lstrlenA(fsName))
End Function

#If Debuging Then
  Function GetISHFolderInterface(debugger As clsDebugger, IParent As IVBShellFolder, pIDLToParent As Long) As IVBShellFolder
#Else
  Function GetISHFolderInterface(IParent As IVBShellFolder, pIDLToParent As Long) As IVBShellFolder
#End If
  Dim ret As IVBShellFolder

  #If Debuging Then
    If IParent Is Nothing Then
      debugger.AddLogEntry "GetIShFolderInterface: IParent = Nothing - leaving function", LogEntryTypeConstants.letWarning
      Exit Function
    End If
    If pIDLToParent = 0 Then
      debugger.AddLogEntry "GetIShFolderInterface: pIDLToParent = 0 - leaving function", LogEntryTypeConstants.letWarning
      Exit Function
    End If
  #End If

  #If Debuging Then
    If IsDesktop(debugger, IParent, pIDLToParent) Then
  #Else
    If IsDesktop(IParent, pIDLToParent) Then
  #End If
    Set GetISHFolderInterface = IDesktop
  Else
    IParent.BindToObject pIDLToParent, 0, IID_IShellFolder, ret
    Set GetISHFolderInterface = ret
    Set ret = Nothing
  End If
End Function

#If Debuging Then
  Function GetISHFolderInterfaceFQ(debugger As clsDebugger, pIDLToDesktop As Long) As IVBShellFolder
#Else
  Function GetISHFolderInterfaceFQ(pIDLToDesktop As Long) As IVBShellFolder
#End If
  Dim dummy As IVBShellFolder
  Dim pIDL_Desktop As Long
  Dim pIDLToParent As Long
  Dim ret As IVBShellFolder

  #If Debuging Then
    If pIDLToDesktop = 0 Then
      debugger.AddLogEntry "GetISHFolderInterfaceFQ: pIDLToDesktop = 0 - leaving function", LogEntryTypeConstants.letWarning
      Exit Function
    End If
  #End If

  pIDL_Desktop = CSIDLTopIDL(CSIDLConstants.CSIDL_DESKTOP)
  If ILIsEqual(pIDLToDesktop, pIDL_Desktop) Then
    Set GetISHFolderInterfaceFQ = IDesktop
  Else
    #If Debuging Then
      SplitFullyQualifiedPIDL debugger, pIDLToDesktop, dummy, pIDLToParent
    #Else
      SplitFullyQualifiedPIDL pIDLToDesktop, dummy, pIDLToParent
    #End If
    If Not (dummy Is Nothing) Then
      dummy.BindToObject pIDLToParent, 0, IID_IShellFolder, ret
    End If
    Set GetISHFolderInterfaceFQ = ret
    Set dummy = Nothing
    Set ret = Nothing
  End If
  #If Debuging Then
    FreeItemIDList debugger, "GetISHFolderInterfaceFQ", pIDL_Desktop
  #Else
    FreeItemIDList pIDL_Desktop
  #End If
End Function

' gibt die Größe (in Bytes) von <pIDL> zurück
#If Debuging Then
  Function GetItemIDListSize(debugger As clsDebugger, pIDL As Long, Optional Count As Integer = -1) As Long
#Else
  Function GetItemIDListSize(pIDL As Long, Optional Count As Integer = -1) As Long
#End If
  Dim i As Integer
  Dim itemIDSize As Integer
  Dim totalSize As Long

  #If Debuging Then
    If pIDL = 0 Then
      debugger.AddLogEntry "GetItemIDListSize: pIDL = 0 - leaving function", LogEntryTypeConstants.letWarning
      Exit Function
    End If
  #End If

  If Count <= -1 Then
    totalSize = ILGetSize(pIDL)
  Else
    For i = 1 To Count
      CopyMemory VarPtr(itemIDSize), pIDL + totalSize, LenB(itemIDSize)
      If itemIDSize = 0 Then Exit For
      totalSize = totalSize + itemIDSize
    Next
  End If

  GetItemIDListSize = totalSize
End Function

' gibt den InfoTip für <pIDLToDesktop> zurück
#If Debuging Then
  Function GetItemInfo(debugger As clsDebugger, hWndShellUIParentWindow As Long, pIDLToDesktop As Long, ByVal Flags As QITipFlags) As String
#Else
  Function GetItemInfo(hWndShellUIParentWindow As Long, pIDLToDesktop As Long, ByVal Flags As QITipFlags) As String
#End If
  Dim bufferSize As Long
  Dim IParent As IVBShellFolder
  Dim IQueryInfo As IVBQueryInfo
  Dim Length As Long
  Dim pIDLToParent As Long
  Dim pInfoTip As Long
  Dim ret As String

  #If Debuging Then
    If pIDLToDesktop = 0 Then
      debugger.AddLogEntry "GetItemInfo: pIDLToDesktop = 0 - leaving function", LogEntryTypeConstants.letWarning
      Exit Function
    End If
  #End If

  #If Debuging Then
    SplitFullyQualifiedPIDL debugger, pIDLToDesktop, IParent, pIDLToParent
  #Else
    SplitFullyQualifiedPIDL pIDLToDesktop, IParent, pIDLToParent
  #End If
  If (pIDLToParent <> 0) And Not (IParent Is Nothing) Then
    If IParent.GetUIObjectOf(hWndShellUIParentWindow, 1, pIDLToParent, IID_IQueryInfo, 0, IQueryInfo) >= 0 Then
      If Not (IQueryInfo Is Nothing) Then
        If IQueryInfo.GetInfoTip(Flags, pInfoTip) = S_OK Then
          Length = lstrlenW(pInfoTip) + 1
          bufferSize = WideCharToMultiByte(CP_ACP, 0, pInfoTip, Length, "", 0, "", 0)
          If bufferSize Then
            ret = String$(bufferSize, Chr$(0))
            WideCharToMultiByte CP_ACP, 0, pInfoTip, Length, ret, bufferSize, "", 0
            ' On Vista and newer, many info tips contain Unicode characters that cannot be converted.
            ' We have told WideCharToMultiByte to replace them with Chr$(0). Now remove those Chr$(0).
            GetItemInfo = Replace$(Left$(ret, bufferSize - 1), Chr$(0), "")
          End If
          CoTaskMemFree pInfoTip
        End If

        Set IQueryInfo = Nothing
      End If
    End If
  End If
  Set IParent = Nothing
End Function

' gibt die letzten <Count> ItemIDs in <pIDL> zurück
#If Debuging Then
  Function GetLastItemIDs(debugger As clsDebugger, pIDL As Long, Optional Count As Integer = 1) As Long
#Else
  Function GetLastItemIDs(pIDL As Long, Optional Count As Integer = 1) As Long
#End If
  Dim Count_all As Integer
  Dim i As Integer
  Dim itemIDSize As Integer
  Dim totalSize As Long

  #If Debuging Then
    If pIDL = 0 Then
      debugger.AddLogEntry "GetLastItemIDs: pIDL = 0 - leaving function", LogEntryTypeConstants.letWarning
      Exit Function
    End If
    If Count < 0 Then
      debugger.AddLogEntry "GetLastItemIDs: Count < 0 - changing it to 1", LogEntryTypeConstants.letWarning
    End If
  #End If
  If Count < 0 Then Count = 1

  If Count = 1 Then
    GetLastItemIDs = ILFindLastID(pIDL)
    Exit Function
  End If

  If Count = 0 Then
    ' die abschließenden 2 Null-Bytes zurückgeben
    GetLastItemIDs = pIDL + ILGetSize(pIDL) - 2
    Exit Function
  End If

  #If Debuging Then
    Count_all = CountItemIDs(debugger, pIDL)
  #Else
    Count_all = CountItemIDs(pIDL)
  #End If
  If Count >= Count_all Then
    GetLastItemIDs = pIDL
    Exit Function
  End If

  For i = 1 To Count_all - Count
    CopyMemory VarPtr(itemIDSize), pIDL + totalSize, LenB(itemIDSize)
    totalSize = totalSize + itemIDSize
  Next
  GetLastItemIDs = pIDL + totalSize
End Function

#If Debuging Then
  Function GetLinkTarget(debugger As clsDebugger, IParent As IVBShellFolder, pIDLToParent As Long) As String
#Else
  Function GetLinkTarget(IParent As IVBShellFolder, pIDLToParent As Long) As String
#End If
  Const FILE_FLAG_BACKUP_SEMANTICS = &H2000000
  Const FILE_FLAG_OPEN_REPARSE_POINT = &H200000
  Const FILE_SHARE_READ = &H1
  Const FSCTL_GET_REPARSE_POINT = &H900A8
  Const INVALID_HANDLE_VALUE = -1
  Const MAXIMUM_REPARSE_DATA_BUFFER_SIZE = 16 * 1024
  Const OPEN_EXISTING = 3
  Dim bBuffer(0 To MAXIMUM_REPARSE_DATA_BUFFER_SIZE - 1) As Byte
  Dim bytesReturned As Long
  Dim DummyA As WIN32_FIND_DATAA
  Dim DummyW As WIN32_FIND_DATAW
  Dim hReparsePoint As Long
  Dim IShLinkA As IVBShellLinkA
  Dim IShLinkW As IVBShellLinkW
  Dim path As String
  Dim posNull As Long
  Dim reparseDetails As REPARSE_DATA_BUFFER
  Dim ret As String

  If IParent Is Nothing Then Exit Function
  If pIDLToParent = 0 Then Exit Function

  IParent.GetUIObjectOf 0, 1, pIDLToParent, IID_IShellLinkW, 0, IShLinkW
  If Not (IShLinkW Is Nothing) Then
    ret = String$(MAX_PATH, Chr$(0))
    IShLinkW.GetPath ret, Len(ret), DummyW, SLGPConstants.SLGP_UNCPRIORITY
    GetLinkTarget = Left$(ret, lstrlenA(ret))

    Set IShLinkW = Nothing
  Else
    IParent.GetUIObjectOf 0, 1, pIDLToParent, IID_IShellLinkA, 0, IShLinkA
    If Not (IShLinkA Is Nothing) Then
      ret = String$(MAX_PATH, Chr$(0))
      IShLinkA.GetPath ret, Len(ret), DummyA, SLGPConstants.SLGP_UNCPRIORITY
      GetLinkTarget = Left$(ret, lstrlenA(ret))

      Set IShLinkA = Nothing
    Else
      #If Debuging Then
        path = pIDLToPath(debugger, IParent, pIDLToParent)
      #Else
        path = pIDLToPath(IParent, pIDLToParent)
      #End If
      hReparsePoint = CreateFileAsLong(path, 0, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS Or FILE_FLAG_OPEN_REPARSE_POINT, 0)
      If hReparsePoint <> INVALID_HANDLE_VALUE Then
        If DeviceIoControl(hReparsePoint, FSCTL_GET_REPARSE_POINT, 0, 0, VarPtr(bBuffer(0)), MAXIMUM_REPARSE_DATA_BUFFER_SIZE, bytesReturned, 0) Then
          CopyMemory VarPtr(reparseDetails), VarPtr(bBuffer(0)), LenB(reparseDetails)
          ret = String$(reparseDetails.SubstituteNameLength / 2 + 1, 0)
          lstrcpyn StrPtr(ret), VarPtr(bBuffer(16)), reparseDetails.SubstituteNameLength / 2 + 1
          If bBuffer(16) = Asc("\") Then
            ' starts with "\??\"
            GetLinkTarget = Mid$(ret, 5, lstrlenA(ret) - 4)
          Else
            GetLinkTarget = Left$(ret, lstrlenA(ret))
          End If
        End If

        CloseHandle hReparsePoint
      End If
    End If
  End If
End Function

' gibt das IShellFolder-Interface des Parent-Objekts von <pIDLToDesktop> zurück
' <pIDLToDesktop> muß relativ zum Desktop sein
#If Debuging Then
  Function GetParentInterface(debugger As clsDebugger, pIDLToDesktop As Long) As IVBShellFolder
#Else
  Function GetParentInterface(pIDLToDesktop As Long) As IVBShellFolder
#End If
  Dim itemIDSize As Integer
  Dim pIDL_Desktop As Long
  Dim pIDLToParent As Long
  Dim ret As IVBShellFolder

  #If Debuging Then
    If pIDLToDesktop = 0 Then
      debugger.AddLogEntry "GetParentInterface: pIDLToDesktop = 0 - leaving function", LogEntryTypeConstants.letWarning
      Exit Function
    End If
  #End If

  pIDL_Desktop = CSIDLTopIDL(CSIDLConstants.CSIDL_DESKTOP)
  If ILIsEqual(pIDLToDesktop, pIDL_Desktop) Then
    Set ret = IDesktop
    pIDLToParent = pIDLToDesktop
  Else
    If ver_Shell32_50 Then
      SHBindToParent pIDLToDesktop, IID_IShellFolder, ret, 0
    Else
      pIDLToParent = ILClone(pIDLToDesktop)
      ILRemoveLastID pIDLToParent
      CopyMemory VarPtr(itemIDSize), pIDLToParent, LenB(itemIDSize)
      If itemIDSize = 0 Then
        Set ret = IDesktop
      Else
        IDesktop.BindToObject pIDLToParent, 0, IID_IShellFolder, ret
      End If
      #If Debuging Then
        FreeItemIDList debugger, "GetParentInterface #1", pIDLToParent
      #Else
        FreeItemIDList pIDLToParent
      #End If
    End If
    ' the next line shouldn't be necessary, but Nero Scout may crash the control via AutoUpdate if it is missing
    ' 30.12.2006: deactivated, because printing a document crashes the control via AutoUpdate otherwise
    'ret.AddRef
  End If
  #If Debuging Then
    FreeItemIDList debugger, "GetParentInterface #2", pIDL_Desktop
  #Else
    FreeItemIDList pIDL_Desktop
  #End If

  Set GetParentInterface = ret
  Set ret = Nothing
End Function

Function GetpIDL(ISHFolder As IVBShellFolder) As Long
  Dim IPF As IVBPersistFolder
  Dim IPF2 As IVBPersistFolder2
  Dim pIDL As Long

  If ISHFolder Is Nothing Then Exit Function

  ISHFolder.QueryInterface IID_IPersistFolder, IPF
  If Not (IPF Is Nothing) Then
    IPF.QueryInterface IID_IPersistFolder2, IPF2
  Else
    ISHFolder.QueryInterface IID_IPersistFolder2, IPF2
  End If
  If Not (IPF2 Is Nothing) Then
    IPF2.GetCurFolder pIDL
    GetpIDL = pIDL
  End If
  Set IPF = Nothing
  Set IPF2 = Nothing
End Function

Function GetShellIconSize() As Long
  Dim Data As String
  Dim ret As Long

  Data = String$(MAX_PATH, Chr$(0))
  ret = SHGetValue(HKEY_CURRENT_USER, "Control Panel\Desktop\WindowMetrics", "Shell Icon Size", VarPtr(REG_SZ), ByVal Data, VarPtr(Len(Data)))
  ' 32 Pixel ist doch hoffentlich überall der Standardwert
  If ret = ERROR_FILE_NOT_FOUND Then Data = 32
  GetShellIconSize = CLng(Left$(Data, lstrlenA(Data)))
End Function

' gibt den Index des Icons für <pIDLToDesktop> zurück
#If Debuging Then
  Function GetSysIconIndex(debugger As clsDebugger, IParent As IVBShellFolder, pIDLToParent As Long, pIDLToDesktop As Long, Selected As Boolean, StandardIcons As Boolean, Optional LargeIcons As Boolean = False) As Long
#Else
  Function GetSysIconIndex(IParent As IVBShellFolder, pIDLToParent As Long, pIDLToDesktop As Long, Selected As Boolean, StandardIcons As Boolean, Optional LargeIcons As Boolean = False) As Long
#End If
  Dim atts As SFGAOConstants
  Dim Data As SHFILEINFO
  Dim ext As String
  Dim Flags As Long
  Dim UseStandardIcons As Boolean

  If pIDLToDesktop = 0 Then Exit Function

  Data.iIcon = -1
  Flags = SHGFI_SYSICONINDEX
  Flags = Flags Or IIf(LargeIcons, SHGFI_LARGEICON, SHGFI_SMALLICON)

  If Selected Or StandardIcons Then
    #If Debuging Then
      If IParent Is Nothing Then
        debugger.AddLogEntry "GetSysIconIndex: IParent = Nothing - HasAttribute() will fail!", LogEntryTypeConstants.letWarning
      End If
      If pIDLToParent = 0 Then
        debugger.AddLogEntry "GetSysIconIndex: pIDLToParent = 0 - HasAttribute() will fail!", LogEntryTypeConstants.letWarning
      End If
    #End If

    #If Debuging Then
      atts = GetAttributes(debugger, IParent, pIDLToParent, SFGAOConstants.SFGAO_FOLDER Or SFGAOConstants.SFGAO_FILESYSTEM)
    #Else
      atts = GetAttributes(IParent, pIDLToParent, SFGAOConstants.SFGAO_FOLDER Or SFGAOConstants.SFGAO_FILESYSTEM)
    #End If
    UseStandardIcons = ((atts And SFGAOConstants.SFGAO_FILESYSTEM) = SFGAOConstants.SFGAO_FILESYSTEM)
    #If Debuging Then
      UseStandardIcons = UseStandardIcons And Not IsDesktoppIDL(debugger, pIDLToDesktop)
      UseStandardIcons = UseStandardIcons And Not IsDrive(debugger, IParent, pIDLToParent)
    #Else
      UseStandardIcons = UseStandardIcons And Not IsDesktoppIDL(pIDLToDesktop)
      UseStandardIcons = UseStandardIcons And Not IsDrive(IParent, pIDLToParent)
    #End If
  End If

  If UseStandardIcons Then
    If atts And SFGAOConstants.SFGAO_FOLDER Then
      If LargeIcons Then
        Data.iIcon = IIf(Selected, DEFICON_OPENFOLDER_LARGE, DEFICON_FOLDER_LARGE)
      Else
        Data.iIcon = IIf(Selected, DEFICON_OPENFOLDER_SMALL, DEFICON_FOLDER_SMALL)
      End If
    Else
      #If Debuging Then
        ext = GetFileNameExtension_pIDL(debugger, IParent, pIDLToParent, True)
      #Else
        ext = GetFileNameExtension_pIDL(IParent, pIDLToParent, True)
      #End If
      If ext <> "" Then
        Flags = Flags Or SHGFI_USEFILEATTRIBUTES
        SHGetFileInfo "." & ext, FILE_ATTRIBUTE_NORMAL, Data, LenB(Data), Flags
      End If
    End If
  End If

  If Data.iIcon = -1 Then
    Flags = Flags Or SHGFI_PIDL
    If Selected Then
      If atts And SFGAOConstants.SFGAO_FOLDER Then
        Flags = Flags Or SHGFI_OPENICON
      End If
    End If
    SHGetFileInfoAsLong pIDLToDesktop, 0, Data, LenB(Data), Flags
  End If

  GetSysIconIndex = Data.iIcon
End Function

' gibt den Index des Icons für <pIDLToDesktop> zurück
#If Debuging Then
  Function GetSysIconIndex_Light(debugger As clsDebugger, pIDLToDesktop As Long, Selected As Boolean, StandardIcons As Boolean, Optional LargeIcons As Boolean = False) As Long
#Else
  Function GetSysIconIndex_Light(pIDLToDesktop As Long, Selected As Boolean, StandardIcons As Boolean, Optional LargeIcons As Boolean = False) As Long
#End If
  Dim Data As SHFILEINFO
  Dim Flags As Long
  Dim IParent As IVBShellFolder
  Dim pIDLToParent As Long

  If pIDLToDesktop = 0 Then Exit Function

  If Selected Or StandardIcons Then
    #If Debuging Then
      SplitFullyQualifiedPIDL debugger, pIDLToDesktop, IParent, pIDLToParent
      If (pIDLToParent = 0) Or (IParent Is Nothing) Then
        debugger.AddLogEntry "GetSysIconIndex_Light: pIDL splitting failed", LogEntryTypeConstants.letWarning
      End If
      GetSysIconIndex_Light = GetSysIconIndex(debugger, IParent, pIDLToParent, pIDLToDesktop, Selected, StandardIcons, LargeIcons)
    #Else
      SplitFullyQualifiedPIDL pIDLToDesktop, IParent, pIDLToParent
      GetSysIconIndex_Light = GetSysIconIndex(IParent, pIDLToParent, pIDLToDesktop, Selected, StandardIcons, LargeIcons)
    #End If
    Set IParent = Nothing
  Else
    Flags = SHGFI_SYSICONINDEX Or SHGFI_PIDL Or IIf(LargeIcons, SHGFI_LARGEICON, SHGFI_SMALLICON)
    SHGetFileInfoAsLong pIDLToDesktop, 0, Data, LenB(Data), Flags
    GetSysIconIndex_Light = Data.iIcon
  End If
End Function

' gibt den Index des Icons für <ext> zurück
Function GetSysIconIndexFromExt(ByVal ext As String, Selected As Boolean, Optional LargeIcons As Boolean = False, Optional ByVal getFolderIcons As Boolean = False) As Long
  Dim attr As Long
  Dim Data As SHFILEINFO
  Dim Flags As Long

  If Not getFolderIcons Then
    If Left$(ext, 1) <> "." Then ext = "." & ext
  Else
    attr = FILE_ATTRIBUTE_DIRECTORY
    ext = "folder"
  End If

  Flags = SHGFI_USEFILEATTRIBUTES Or SHGFI_SYSICONINDEX
  If Selected Then
    Flags = Flags Or SHGFI_OPENICON
  End If
  Flags = Flags Or IIf(LargeIcons, SHGFI_LARGEICON, SHGFI_SMALLICON)
  SHGetFileInfo ext, attr, Data, LenB(Data), Flags

  GetSysIconIndexFromExt = Data.iIcon
End Function

' gibt die System-ImageList zurück
Function GetSysImageList(Optional LargeIcons As Boolean = False) As Long
  Dim Data As SHFILEINFO
  Dim Flags As Long

  ' unter WinNT4 wird angeblich eine uninitialisierte System-ImageList zurückgegeben
  If ver_Win_NTBased Then FileIconInit 1

  Flags = SHGFI_SYSICONINDEX Or SHGFI_USEFILEATTRIBUTES
  Flags = Flags Or IIf(LargeIcons, SHGFI_LARGEICON, SHGFI_SMALLICON)

  GetSysImageList = SHGetFileInfo(".txt", FILE_ATTRIBUTE_NORMAL, Data, LenB(Data), Flags)
End Function

Sub GetUUIDs()
  CLSIDFromString StrPtr(strIID_IContextMenu), IID_IContextMenu
  CLSIDFromString StrPtr(strIID_IContextMenu2), IID_IContextMenu2
  CLSIDFromString StrPtr(strIID_IContextMenu3), IID_IContextMenu3
  CLSIDFromString StrPtr(strIID_IDataObject), IID_IDataObject
  CLSIDFromString StrPtr(strIID_IDragDropHelper), IID_IDragDropHelper
  CLSIDFromString StrPtr(strIID_IDragSourceHelper), IID_IDragSourceHelper
  CLSIDFromString StrPtr(strIID_IDragSourceHelper2), IID_IDragSourceHelper2
  CLSIDFromString StrPtr(strIID_IDropTarget), IID_IDropTarget
  CLSIDFromString StrPtr(strIID_IDropTargetHelper), IID_IDropTargetHelper
  CLSIDFromString StrPtr(strIID_IPersistFolder), IID_IPersistFolder
  CLSIDFromString StrPtr(strIID_IPersistFolder2), IID_IPersistFolder2
  CLSIDFromString StrPtr(strIID_IQueryInfo), IID_IQueryInfo
  CLSIDFromString StrPtr(strIID_IShellFolder), IID_IShellFolder
  CLSIDFromString StrPtr(strIID_IShellIcon), IID_IShellIcon
  CLSIDFromString StrPtr(strIID_IShellIconOverlay), IID_IShellIconOverlay
  CLSIDFromString StrPtr(strIID_IShellLinkA), IID_IShellLinkA
  CLSIDFromString StrPtr(strIID_IShellLinkW), IID_IShellLinkW
  #If Debuging Then
    CLSIDFromString StrPtr(strCLSID_AugmentedShellFolder), CLSID_AugmentedShellFolder
    CLSIDFromString StrPtr(strCLSID_AugmentedShellFolder2), CLSID_AugmentedShellFolder2
    CLSIDFromString StrPtr(strIID_IAugmentedShellFolder), IID_IAugmentedShellFolder
    CLSIDFromString StrPtr(strIID_IAugmentedShellFolder2), IID_IAugmentedShellFolder2
    CLSIDFromString StrPtr(strIID_IDelegateFolder), IID_IDelegateFolder
    CLSIDFromString StrPtr(strIID_IEnumUICommand), IID_IEnumUICommand
    CLSIDFromString StrPtr(strIID_IShellFolder2), IID_IShellFolder2
    CLSIDFromString StrPtr(strIID_IUICommand), IID_IUICommand
    CLSIDFromString StrPtr(strIID_IUICommandTarget), IID_IUICommandTarget
    CLSIDFromString StrPtr(strIID_IUIElement), IID_IUIElement
    CLSIDFromString StrPtr(strIID_IThumbnailProvider), IID_IThumbnailProvider
  #End If
End Sub

' prüft, ob <pIDLToParent> das Attribut <Attr> hat
Function HasAttribute(IParent As IVBShellFolder, pIDLToParent As Long, ByVal attr As SFGAOConstants) As Boolean
  Dim tmp As SFGAOConstants

  If IParent Is Nothing Then Exit Function
  If pIDLToParent = 0 Then Exit Function

  tmp = attr
  IParent.GetAttributesOf 1, pIDLToParent, tmp
  HasAttribute = ((tmp And attr) = attr)
End Function

' prüft, ob <pIDLToDesktop> SubItems hat
' anhand der Filterkriterien von <Ctl> wird geprüft, ob <pIDLToDesktop> SubItems hat, die angezeigt
' werden sollen
#If Debuging Then
  Function HasSubItems(debugger As clsDebugger, pIDLToDesktop As Long, Ctl As ExplorerTreeView) As Boolean
#Else
  Function HasSubItems(pIDLToDesktop As Long, Ctl As ExplorerTreeView) As Boolean
#End If
  Dim EnumFlags As SHCONTFConstants
  Dim ext As String
  Dim IEnum As IVBEnumIDList
  Dim IItem As IVBShellFolder
  Dim IParent As IVBShellFolder
  Dim itemAttr As SFGAOConstants
  Dim pIDL_Desktop As Long
  Dim pIDLToParent As Long
  Dim pIDLSubItem_ToDesktop As Long
  Dim pIDLSubItem_ToParent As Long
  Dim ret As Boolean

  If pIDLToDesktop = 0 Then Exit Function
  If Ctl Is Nothing Then Exit Function

  ' übergebene Daten aufbereiten
  pIDL_Desktop = CSIDLTopIDL(CSIDLConstants.CSIDL_DESKTOP)
  If ILIsEqual(pIDLToDesktop, pIDL_Desktop) Then
    pIDLToParent = pIDLToDesktop
    Set IParent = IDesktop
  Else
    #If Debuging Then
      SplitFullyQualifiedPIDL debugger, pIDLToDesktop, IParent, pIDLToParent
    #Else
      SplitFullyQualifiedPIDL pIDLToDesktop, IParent, pIDLToParent
    #End If
  End If
  #If Debuging Then
    FreeItemIDList debugger, "HasSubItems #1", pIDL_Desktop
  #Else
    FreeItemIDList pIDL_Desktop
  #End If

  If IParent Is Nothing Then Exit Function
  If pIDLToParent = 0 Then Exit Function

  #If Debuging Then
    itemAttr = GetAttributes(debugger, IParent, pIDLToParent, SFGAOConstants.SFGAO_FOLDER Or SFGAOConstants.SFGAO_FILESYSTEM)
  #Else
    itemAttr = GetAttributes(IParent, pIDLToParent, SFGAOConstants.SFGAO_FOLDER Or SFGAOConstants.SFGAO_FILESYSTEM)
  #End If
  If itemAttr And SFGAOConstants.SFGAO_FOLDER Then
    If Ctl.DontCheckFolderExpandibility Then
      ret = True
      GoTo Ende
    End If

    ' jetzt solange alle SubItems durchgehen und prüfen bis ein Item gefunden wird, der nicht herausgefiltert
    ' wurde

    ' Aufzählung initiieren
    #If Debuging Then
      Set IItem = GetISHFolderInterface(debugger, IParent, pIDLToParent)
    #Else
      Set IItem = GetISHFolderInterface(IParent, pIDLToParent)
    #End If
    If Not (IItem Is Nothing) Then
      ret = False
      ' on Vista SHCONTF_DRIVES doesn't work anymore, so let shouldShowItem() do the work
      If Ctl.DrivesOnly And Not ver_Win_Vista Then
        #If Debuging Then
          If IsMyComputer(debugger, IItem) Then
        #Else
          If IsMyComputer(IItem) Then
        #End If
          EnumFlags = SHCONTFConstants.SHCONTF_DRIVES
        Else
          EnumFlags = SHCONTFConstants.SHCONTF_FOLDERS
          If Ctl.IncludedItems And IncludedItemsConstants.iiFSFiles Then EnumFlags = EnumFlags Or SHCONTFConstants.SHCONTF_NONFOLDERS
          If Ctl.IncludedItems And IncludedItemsConstants.iiNonFSFiles Then EnumFlags = EnumFlags Or SHCONTFConstants.SHCONTF_NONFOLDERS
          If Ctl.FileAttributes And AttributesConstants.attHidden Then EnumFlags = EnumFlags Or SHCONTFConstants.SHCONTF_INCLUDEHIDDEN
          If Ctl.FolderAttributes And AttributesConstants.attHidden Then EnumFlags = EnumFlags Or SHCONTFConstants.SHCONTF_INCLUDEHIDDEN
        End If
      Else
        EnumFlags = SHCONTFConstants.SHCONTF_FOLDERS
        If Ctl.IncludedItems And IncludedItemsConstants.iiFSFiles Then EnumFlags = EnumFlags Or SHCONTFConstants.SHCONTF_NONFOLDERS
        If Ctl.IncludedItems And IncludedItemsConstants.iiNonFSFiles Then EnumFlags = EnumFlags Or SHCONTFConstants.SHCONTF_NONFOLDERS
        If Ctl.FileAttributes And AttributesConstants.attHidden Then EnumFlags = EnumFlags Or SHCONTFConstants.SHCONTF_INCLUDEHIDDEN
        If Ctl.FolderAttributes And AttributesConstants.attHidden Then EnumFlags = EnumFlags Or SHCONTFConstants.SHCONTF_INCLUDEHIDDEN
      End If

      IItem.EnumObjects 0, EnumFlags, IEnum
      If Not (IEnum Is Nothing) Then
        While (IEnum.Next(1, pIDLSubItem_ToParent, 0) = S_OK) And (ret = False)
          If pIDLSubItem_ToParent Then
            pIDLSubItem_ToDesktop = ILAppendID(ILClone(pIDLToDesktop), pIDLSubItem_ToParent, 1)
            #If Debuging Then
              ret = ShouldShowItem(debugger, Ctl, IItem, pIDLSubItem_ToParent, pIDLSubItem_ToDesktop, -1)
              FreeItemIDList debugger, "HasSubItems #2", pIDLSubItem_ToDesktop
              FreeItemIDList debugger, "HasSubItems #3", pIDLSubItem_ToParent
            #Else
              ret = ShouldShowItem(Ctl, IItem, pIDLSubItem_ToParent, pIDLSubItem_ToDesktop, -1)
              FreeItemIDList pIDLSubItem_ToDesktop
              FreeItemIDList pIDLSubItem_ToParent
            #End If
          End If
        Wend
      End If
      Set IEnum = Nothing
    Else
      ' wahrscheinlich ein Laufwerk, bei dem das Medium fehlt
    End If
    Set IItem = Nothing
  ElseIf Ctl.ExpandArchives <> 0 Then
    If itemAttr And SFGAOConstants.SFGAO_FILESYSTEM Then
      ' falls es ein Archiv ist, prüfen ob dieser Archivtyp als Ordner behandelt werden soll
      #If Debuging Then
        ext = GetFileNameExtension_pIDL(debugger, IParent, pIDLToParent)
      #Else
        ext = GetFileNameExtension_pIDL(IParent, pIDLToParent)
      #End If
      Select Case LCase$(ext)
        Case "zip"
          ret = (Ctl.ExpandArchives And 16)   ' ExpandArchivesConstants.eaZIP
        Case "rar"
          ret = (Ctl.ExpandArchives And 8)   ' ExpandArchivesConstants.eaRAR
        Case "iso"
          ret = (Ctl.ExpandArchives And 32)   ' ExpandArchivesConstants.eaISO
        Case "ace"
          ret = (Ctl.ExpandArchives And 1)   ' ExpandArchivesConstants.eaACE
        #If NewArchiveSupport Then
          Case "tar"
            ret = (Ctl.ExpandArchives And 128)   ' ExpandArchivesConstants.eaTAR
        #End If
        Case "cab"
          ret = (Ctl.ExpandArchives And 2)   ' ExpandArchivesConstants.eaCAB
        Case "jar"
          ret = (Ctl.ExpandArchives And 4)   ' ExpandArchivesConstants.eaJAR
        Case "bin"
          ret = (Ctl.ExpandArchives And 64)   ' ExpandArchivesConstants.eaBIN
      End Select
    End If
  End If

Ende:
  Set IParent = Nothing
  HasSubItems = ret
End Function

' prüft, ob <pIDLToParent> ein Archiv ist, welches als Ordner behandelt werden soll
#If Debuging Then
  Function IsArchiveToExpand(debugger As clsDebugger, IParent As IVBShellFolder, pIDLToParent As Long, ExpandArchives As Long) As Boolean
#Else
  Function IsArchiveToExpand(IParent As IVBShellFolder, pIDLToParent As Long, ExpandArchives As Long) As Boolean
#End If
  Dim ext As String
  Dim itemAttr As SFGAOConstants
  Dim ret As Boolean

  If IParent Is Nothing Then Exit Function
  If pIDLToParent = 0 Then Exit Function
  If ExpandArchives = 0 Then Exit Function

  #If Debuging Then
    itemAttr = GetAttributes(debugger, IParent, pIDLToParent, SFGAOConstants.SFGAO_FOLDER Or SFGAOConstants.SFGAO_FILESYSTEM)
  #Else
    itemAttr = GetAttributes(IParent, pIDLToParent, SFGAOConstants.SFGAO_FOLDER Or SFGAOConstants.SFGAO_FILESYSTEM)
  #End If
  If ((itemAttr And SFGAOConstants.SFGAO_FOLDER) = SFGAOConstants.SFGAO_FOLDER) And ((itemAttr And SFGAOConstants.SFGAO_FILESYSTEM) = SFGAOConstants.SFGAO_FILESYSTEM) Then
    ' ist es wirklich ein Ordner?
    #If Debuging Then
      If Not FileExists_pIDL(debugger, IParent, pIDLToParent) Then
    #Else
      If Not FileExists_pIDL(IParent, pIDLToParent) Then
    #End If
      Exit Function
    End If
  ElseIf Not (((itemAttr And SFGAOConstants.SFGAO_FOLDER) = 0) And ((itemAttr And SFGAOConstants.SFGAO_FILESYSTEM) = SFGAOConstants.SFGAO_FILESYSTEM)) Then
    Exit Function
  End If

  #If Debuging Then
    ext = GetFileNameExtension_pIDL(debugger, IParent, pIDLToParent, True)
  #Else
    ext = GetFileNameExtension_pIDL(IParent, pIDLToParent, True)
  #End If
  Select Case LCase$(ext)
    Case "zip"
      ret = ((ExpandArchives And 16) = 16)   ' ExpandArchivesConstants.eaZIP
    Case "rar"
      ret = ((ExpandArchives And 8) = 8)   ' ExpandArchivesConstants.eaRAR
    Case "iso"
      ret = ((ExpandArchives And 32) = 32)   ' ExpandArchivesConstants.eaISO
    Case "ace"
      ret = ((ExpandArchives And 1) = 1)   ' ExpandArchivesConstants.eaACE
    #If NewArchiveSupport Then
      Case "tar"
        ret = ((ExpandArchives And 128) = 128)   ' ExpandArchivesConstants.eaTAR
    #End If
    Case "cab"
      ret = ((ExpandArchives And 2) = 2)   ' ExpandArchivesConstants.eaCAB
    Case "jar"
      ret = ((ExpandArchives And 4) = 4)   ' ExpandArchivesConstants.eaJAR
    Case "bin"
      ret = ((ExpandArchives And 64) = 64)   ' ExpandArchivesConstants.eaBIN
  End Select

  IsArchiveToExpand = ret
End Function

' prüft, ob <pIDLToParent> ein Archiv ist, welches als Ordner behandelt werden soll
#If Debuging Then
  Function IsArchiveToExpandFQ(debugger As clsDebugger, pIDLToDesktop As Long, ExpandArchives As Long) As Boolean
    Dim IParent As IVBShellFolder
    Dim pIDLToParent As Long

    If pIDLToDesktop = 0 Then Exit Function
    If ExpandArchives = 0 Then Exit Function

    SplitFullyQualifiedPIDL debugger, pIDLToDesktop, IParent, pIDLToParent
    IsArchiveToExpandFQ = IsArchiveToExpand(debugger, IParent, pIDLToParent, ExpandArchives)
    Set IParent = Nothing
  End Function
#Else
  Function IsArchiveToExpandFQ(pIDLToDesktop As Long, ExpandArchives As Long) As Boolean
    Dim IParent As IVBShellFolder
    Dim pIDLToParent As Long

    If pIDLToDesktop = 0 Then Exit Function
    If ExpandArchives = 0 Then Exit Function

    SplitFullyQualifiedPIDL pIDLToDesktop, IParent, pIDLToParent
    IsArchiveToExpandFQ = IsArchiveToExpand(IParent, pIDLToParent, ExpandArchives)
    Set IParent = Nothing
  End Function
#End If

' prüfen, ob <txt> eine CSIDL ist
#If Debuging Then
  Function IsCSIDL(debugger As clsDebugger, ByVal txt As String, Optional ByVal allowMenuItems As Boolean = False) As Boolean
#Else
  Function IsCSIDL(ByVal txt As String, Optional ByVal allowMenuItems As Boolean = False) As Boolean
#End If
  Dim loTmp As Long

  txt = Trim$(txt)
  If txt = "" Then Exit Function

  On Error Resume Next
  loTmp = CLng(txt)
  If Err Then
    ' ein String?!
    If allowMenuItems Then IsCSIDL = (MenuItemToCSIDL(txt) <> -1)
  Else
    ' eine Zahl!
    #If Debuging Then
      IsCSIDL = (GetCSIDLDisplayName(debugger, loTmp) <> "")
    #Else
      IsCSIDL = (GetCSIDLDisplayName(loTmp) <> "")
    #End If
  End If
End Function

' prüft, ob <pIDLToParent> der Desktop ist
#If Debuging Then
  Function IsDesktop(debugger As clsDebugger, IParent As IVBShellFolder, pIDLToParent As Long) As Boolean
#Else
  Function IsDesktop(IParent As IVBShellFolder, pIDLToParent As Long) As Boolean
#End If
  Dim IPF As IVBPersistFolder
  Dim IPF2 As IVBPersistFolder2
  Dim pIDL As Long
  Dim pIDL_Desktop As Long
  Dim ret As Boolean

  If IParent Is Nothing Then Exit Function
  If pIDLToParent = 0 Then Exit Function

  pIDL_Desktop = CSIDLTopIDL(CSIDLConstants.CSIDL_DESKTOP)
  ret = ILIsEqual(pIDL_Desktop, pIDLToParent)
  If ret Then
    ' compare the interfaces
    IParent.QueryInterface IID_IPersistFolder, IPF
    If Not (IPF Is Nothing) Then
      IPF.QueryInterface IID_IPersistFolder2, IPF2
    Else
      IParent.QueryInterface IID_IPersistFolder2, IPF2
    End If
    If Not (IPF2 Is Nothing) Then
      IPF2.GetCurFolder pIDL
      ret = ILIsEqual(pIDL, pIDL_Desktop)
      #If Debuging Then
        FreeItemIDList debugger, "IsDesktop #1", pIDL
      #Else
        FreeItemIDList pIDL
      #End If
    End If
  End If
  Set IPF = Nothing
  Set IPF2 = Nothing
  #If Debuging Then
    FreeItemIDList debugger, "IsDesktop #2", pIDL_Desktop
  #Else
    FreeItemIDList pIDL_Desktop
  #End If

  IsDesktop = ret
End Function

#If Debuging Then
  Function IsDesktoppIDL(debugger As clsDebugger, pIDLToDesktop As Long) As Boolean
#Else
  Function IsDesktoppIDL(pIDLToDesktop As Long) As Boolean
#End If
  Dim pIDL_Desktop As Long

  #If Debuging Then
    If pIDLToDesktop = 0 Then
      debugger.AddLogEntry "IsDesktoppIDL: pIDLToDesktop = 0 - leaving function", LogEntryTypeConstants.letWarning
      Exit Function
    End If
  #End If

  pIDL_Desktop = CSIDLTopIDL(CSIDLConstants.CSIDL_DESKTOP)
  IsDesktoppIDL = ILIsEqual(pIDLToDesktop, pIDL_Desktop)
  #If Debuging Then
    FreeItemIDList debugger, "IsDesktoppIDL", pIDL_Desktop
  #Else
    FreeItemIDList pIDL_Desktop
  #End If
End Function

' prüft, ob <txt> ein Verzeichnis ist
Function IsDirectory(ByVal txt As String) As Boolean
  If txt = "" Then Exit Function

  IsDirectory = PathIsDirectory(txt)
End Function

' prüft, ob <pIDLToRelative> ein Laufwerk ist
#If Debuging Then
  Function IsDrive(debugger As clsDebugger, IRelative As IVBShellFolder, pIDLToRelative As Long) As Boolean
#Else
  Function IsDrive(IRelative As IVBShellFolder, pIDLToRelative As Long) As Boolean
#End If
  Dim Data As SHDESCRIPTIONID
  Dim path As String
  Dim ret As Boolean

  If IRelative Is Nothing Then Exit Function
  If pIDLToRelative = 0 Then Exit Function

  If SHGetDataFromIDList(IRelative, pIDLToRelative, SHGDFIL_DESCRIPTIONID, Data, LenB(Data)) = NOERROR Then
    Select Case Data.dwDescriptionId
      Case SHDID_COMPUTER_DRIVE35, SHDID_COMPUTER_DRIVE525, SHDID_COMPUTER_REMOVABLE, SHDID_COMPUTER_FIXED, SHDID_COMPUTER_NETDRIVE, SHDID_COMPUTER_CDROM, SHDID_COMPUTER_RAMDISK, SHDID_COMPUTER_OTHER
        IsDrive = True
    End Select
  ' TODO: wir gehen davon aus, dass <IRelative> der direkte Parent-Item von <pIDLToRelative> ist
  #If Debuging Then
    ElseIf IsMyComputer(debugger, IRelative) And CountItemIDs(debugger, pIDLToRelative) = 1 Then
      path = Left$(pIDLToPath(debugger, IRelative, pIDLToRelative), 2)
  #Else
    ElseIf IsMyComputer(IRelative) And CountItemIDs(pIDLToRelative) = 1 Then
      path = Left$(pIDLToPath(IRelative, pIDLToRelative), 2)
  #End If
    IsDrive = (GetDriveType(path) > 1)
  End If
End Function

' prüft, ob <pIDLToParent> eine Datei des Dateisystems ist
#If Debuging Then
  Function IsFSFile(debugger As clsDebugger, IParent As IVBShellFolder, pIDLToParent As Long) As Boolean
#Else
  Function IsFSFile(IParent As IVBShellFolder, pIDLToParent As Long) As Boolean
#End If
  Dim itemAttr As SFGAOConstants

  #If Debuging Then
    itemAttr = GetAttributes(debugger, IParent, pIDLToParent, SFGAOConstants.SFGAO_FOLDER Or SFGAOConstants.SFGAO_FILESYSTEM)
  #Else
    itemAttr = GetAttributes(IParent, pIDLToParent, SFGAOConstants.SFGAO_FOLDER Or SFGAOConstants.SFGAO_FILESYSTEM)
  #End If
  IsFSFile = ((itemAttr And SFGAOConstants.SFGAO_FOLDER) = 0) And ((itemAttr And SFGAOConstants.SFGAO_FILESYSTEM) = SFGAOConstants.SFGAO_FILESYSTEM)
End Function

' prüft, ob <pIDLToParent> ein Ordner des Dateisystems ist
Function IsFSFolder(IParent As IVBShellFolder, pIDLToParent As Long) As Boolean
  IsFSFolder = HasAttribute(IParent, pIDLToParent, SFGAOConstants.SFGAO_FOLDER Or SFGAOConstants.SFGAO_FILESYSTEM)
End Function

' prüft, ob <txt> eine ftp-URL ist
Function IsFTPURL(ByVal txt As String) As Boolean
  Dim pos As Long
  Dim ret As Boolean

  If Trim$(txt) = "" Then Exit Function

  ret = PathIsURL(txt)
  If ret Then
    ' Man könnte ab shlwapi.dll v5.0 (IE5) auch UrlGetPart() nutzen.
    ret = (Left$(txt, Len("ftp://")) = "ftp://")
  End If

  IsFTPURL = ret
End Function

' prüft, ob <pIDLToParent> auf einen Netzwerkordner verlinkt
#If Debuging Then
  Function IsLinkToNetworkFolder(debugger As clsDebugger, IParent As IVBShellFolder, pIDLToParent As Long) As Boolean
#Else
  Function IsLinkToNetworkFolder(IParent As IVBShellFolder, pIDLToParent As Long) As Boolean
#End If
  Dim ret As Boolean
  Dim target As String

  If HasAttribute(IParent, pIDLToParent, SFGAOConstants.SFGAO_LINK Or SFGAOConstants.SFGAO_FOLDER) Then
    #If Debuging Then
      target = GetLinkTarget(debugger, IParent, pIDLToParent)
    #Else
      target = GetLinkTarget(IParent, pIDLToParent)
    #End If
    ret = IsUNC(target)
    If Not ret Then ret = IsURL(target)
  End If

  IsLinkToNetworkFolder = ret
End Function

' prüft, ob <IShFolder> der Arbeitsplatz ist
#If Debuging Then
  Function IsMyComputer(debugger As clsDebugger, ISHFolder As IVBShellFolder) As Boolean
#Else
  Function IsMyComputer(ISHFolder As IVBShellFolder) As Boolean
#End If
  Dim IPF As IVBPersistFolder
  Dim IPF2 As IVBPersistFolder2
  Dim pIDL As Long
  Dim pIDLMyComputer As Long

  If ISHFolder Is Nothing Then Exit Function

  pIDLMyComputer = CSIDLTopIDL(CSIDLConstants.CSIDL_DRIVES)
  ISHFolder.QueryInterface IID_IPersistFolder, IPF
  If Not (IPF Is Nothing) Then
    IPF.QueryInterface IID_IPersistFolder2, IPF2
  Else
    ISHFolder.QueryInterface IID_IPersistFolder2, IPF2
  End If
  If Not (IPF2 Is Nothing) Then
    IPF2.GetCurFolder pIDL
    IsMyComputer = ILIsEqual(pIDL, pIDLMyComputer)
    #If Debuging Then
      FreeItemIDList debugger, "IsMyComputer #1", pIDL
    #Else
      FreeItemIDList pIDL
    #End If
  End If
  Set IPF = Nothing
  Set IPF2 = Nothing
  #If Debuging Then
    FreeItemIDList debugger, "IsMyComputer #2", pIDLMyComputer
  #Else
    FreeItemIDList pIDLMyComputer
  #End If
End Function

' prüft, ob <pIDLToParent> ein Netzlaufwerk ist
#If Debuging Then
  Function IsNetworkDrive(debugger As clsDebugger, IParent As IVBShellFolder, pIDLToParent As Long) As Boolean
#Else
  Function IsNetworkDrive(IParent As IVBShellFolder, pIDLToParent As Long) As Boolean
#End If
  Dim DrvType As Long
  Dim path As String

  #If Debuging Then
    path = pIDLToPath(debugger, IParent, pIDLToParent)
  #Else
    path = pIDLToPath(IParent, pIDLToParent)
  #End If
  'If Not (Path Like "::{????????-????-????-????-????????????}*") Then
    path = GetFirstFolders(path, 1)
    path = AddBackslash(GetPathName(path), False)
    DrvType = GetDriveType(path)

    IsNetworkDrive = ((DrvType = DRIVE_REMOTE) Or (DrvType = DRIVE_NO_ROOT_DIR))
  'End If
End Function

' prüft, ob <pIDLToParent> zum Dateisystem gehört
Function IsPartOfFileSystem(IParent As IVBShellFolder, pIDLToParent As Long) As Boolean
  IsPartOfFileSystem = HasAttribute(IParent, pIDLToParent, SFGAOConstants.SFGAO_FILESYSTEM)
End Function

#If Debuging Then
  Function IsPartOfRecycler(debugger As clsDebugger, pIDLToDesktop As Long, Optional ByRef isRecycler As Boolean = False, Optional freepIDL As Boolean = False) As Boolean
#Else
  Function IsPartOfRecycler(pIDLToDesktop As Long, Optional ByRef isRecycler As Boolean = False, Optional freepIDL As Boolean = False) As Boolean
#End If
  Dim arrBitBuckets() As String
  Dim DispName As String
  Dim i As Long
  Dim pIDLBitBucket As Long
  Dim ret As Boolean

  isRecycler = False
  pIDLBitBucket = CSIDLTopIDL(CSIDLConstants.CSIDL_BITBUCKET)
  If ILIsEqual(pIDLBitBucket, pIDLToDesktop) Then
    ret = True
  ElseIf ILIsParent(pIDLBitBucket, pIDLToDesktop, 0) Then
    ret = True
  End If
  #If Debuging Then
    FreeItemIDList debugger, "IsPartOfRecycler #1", pIDLBitBucket
  #Else
    FreeItemIDList pIDLBitBucket
  #End If

  If Not ret Then
    #If Debuging Then
      arrBitBuckets = Split(GetAllRecycleBins(debugger), "|")
    #Else
      arrBitBuckets = Split(GetAllRecycleBins, "|")
    #End If
    If Not IsEmpty(arrBitBuckets) Then
      For i = LBound(arrBitBuckets) To UBound(arrBitBuckets)
        pIDLBitBucket = PathTopIDL(arrBitBuckets(i))

        If ILIsEqual(pIDLBitBucket, pIDLToDesktop) Then
          isRecycler = True
          ret = True
        ElseIf ILIsParent(pIDLBitBucket, pIDLToDesktop, 0) Then
          ret = True
        End If

        #If Debuging Then
          FreeItemIDList debugger, "IsPartOfRecycler #2", pIDLBitBucket
        #Else
          FreeItemIDList pIDLBitBucket
        #End If
        If ret Then Exit For
      Next i
    End If
  End If

  If freepIDL Then
    #If Debuging Then
      FreeItemIDList debugger, "IsPartOfRecycler #3", pIDLToDesktop
    #Else
      FreeItemIDList pIDLToDesktop
    #End If
  End If
  IsPartOfRecycler = ret
End Function

' prüft, ob <pIDLToDesktop> ein Drucker ist
#If Debuging Then
  Function IsPrinter_Light(debugger As clsDebugger, pIDLToDesktop As Long) As Boolean
#Else
  Function IsPrinter_Light(pIDLToDesktop As Long) As Boolean
#End If
  Dim DispName As String
  Dim hPrinter As Long
  Dim pIDL As Long

  If pIDLToDesktop = 0 Then Exit Function

  pIDL = CSIDLTopIDL(CSIDLConstants.CSIDL_PRINTERS)
  If ILIsParent(pIDL, pIDLToDesktop, 0) Then
    #If Debuging Then
      DispName = pIDLToDisplayName_Light(debugger, pIDLToDesktop)
    #Else
      DispName = pIDLToDisplayName_Light(pIDLToDesktop)
    #End If

    OpenPrinterAsLong DispName, hPrinter, 0
    IsPrinter_Light = (hPrinter <> 0)
    ClosePrinter hPrinter
  End If
  #If Debuging Then
    FreeItemIDList debugger, "IsPrinter_Light", pIDL
  #Else
    FreeItemIDList pIDL
  #End If
End Function

' prüft, ob bei <pIDLToParent> mit sehr langsamem Zugriff zu rechnen ist
' -> ist erfüllt, wenn <pIDLToParent> auf einem auswechselbaren Datenträger oder einem Netzlaufwerk liegt,
'    es auf eine Netzwerkressource verlinkt (bei Ordnern) oder im Netzwerk liegt
' Subitems der Netzwerkumgebung werden immer als langsam eingestuft
#If Debuging Then
  Function IsSlowItem(debugger As clsDebugger, IParent As IVBShellFolder, pIDLToParent As Long, pIDLToDesktop As Long, Optional excludeZipArchives As Boolean = False) As Boolean
#Else
  Function IsSlowItem(IParent As IVBShellFolder, pIDLToParent As Long, pIDLToDesktop As Long, Optional excludeZipArchives As Boolean = False) As Boolean
#End If
  Dim attr As SFGAOConstants
  Dim path As String
  Dim pIDLNetHood As Long
  Dim ret As Boolean

  If IParent Is Nothing Then Exit Function
  If pIDLToParent = 0 Then Exit Function

  pIDLNetHood = CSIDLTopIDL(CSIDLConstants.CSIDL_NETWORK)
  ret = ILIsParent(pIDLNetHood, pIDLToDesktop, 0)
  #If Debuging Then
    FreeItemIDList debugger, "IsSlowItem", pIDLNetHood
  #Else
    FreeItemIDList pIDLNetHood
  #End If
  If Not ret Then
    #If Debuging Then
      ret = IsNetworkDrive(debugger, IParent, pIDLToParent)
    #Else
      ret = IsNetworkDrive(IParent, pIDLToParent)
    #End If
  End If
  If Not ret Then
    attr = SFGAOConstants.SFGAO_REMOVABLE Or SFGAOConstants.SFGAO_FOLDER Or SFGAOConstants.SFGAO_STREAM
    IParent.GetAttributesOf 1, pIDLToParent, attr
    ret = ((attr And SFGAOConstants.SFGAO_REMOVABLE) = SFGAOConstants.SFGAO_REMOVABLE)
    If Not ret And Not excludeZipArchives Then
      ' check for zip archives
      ret = ((attr And (SFGAOConstants.SFGAO_FOLDER Or SFGAOConstants.SFGAO_STREAM)) = (SFGAOConstants.SFGAO_FOLDER Or SFGAOConstants.SFGAO_STREAM))
    End If
  End If
  If Not ret Then
    #If Debuging Then
      ret = IsLinkToNetworkFolder(debugger, IParent, pIDLToParent)
    #Else
      ret = IsLinkToNetworkFolder(IParent, pIDLToParent)
    #End If
  End If
  If Not ret Then
    #If Debuging Then
      path = pIDLToPath(debugger, IParent, pIDLToParent)
    #Else
      path = pIDLToPath(IParent, pIDLToParent)
    #End If
    path = AddBackslash(path, False)
    ret = IsUNC(path)
    If Not ret Then ret = IsURL(path)
  End If

  IsSlowItem = ret
End Function

' prüft, ob <txt> ein Netzwerkpfad ist
Function IsUNC(ByVal txt As String) As Boolean
  If Trim$(txt) = "" Then Exit Function

  IsUNC = PathIsUNC(txt)
End Function

' prüft, ob <txt> eine URL ist
Function IsURL(ByVal txt As String) As Boolean
  If Trim$(txt) = "" Then Exit Function

  IsURL = PathIsURL(txt)
End Function

' gibt die pIDL des Parent-Objektes von <pIDL> zurück
' Rückgabewerte: TRUE  - ItemID erfolgreich abgetrennt
'                FALSE - <pIDL> scheint kein Parent-Objekt zu haben
Function MakeParentItemIDList(pIDL As Long) As Boolean
  If pIDL = 0 Then Exit Function

  MakeParentItemIDList = ILRemoveLastID(pIDL)
End Function

' wandelt <MenuItem> in eine CSIDL um
Function MenuItemToCSIDL(ByVal MenuItem As String) As CSIDLConstants
  On Error Resume Next
  ' vielleicht ist es schon eine CSIDL
  MenuItemToCSIDL = CLng(MenuItem)
  If Err = 0 Then Exit Function

  Select Case LCase$(MenuItem)
    Case "administrator-tools"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_ADMINTOOLS
    Case "administrator-tools (all users)"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_COMMON_ADMINTOOLS
    Case "aktuelles benutzerprofil"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_PROFILE
    Case "anwendungsdaten"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_APPDATA
    Case "anwendungsdaten (all users)"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_COMMON_APPDATA
    Case "arbeitsgruppencomputer"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_COMPUTERSNEARME
    Case "arbeitsplatz"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_DRIVES
    Case "autostart"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_STARTUP
    Case "autostart (all users)"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_COMMON_STARTUP
    Case "cd burning"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_CDBURN_AREA
    Case "cookies"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_COOKIES
    Case "desktop"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_DESKTOP
    Case "desktop (ordner)"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_DESKTOPDIRECTORY
    Case "desktop (ordner, all users)"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_COMMON_DESKTOPDIRECTORY
    Case "dokumente"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_RECENT
    Case "dokumente (all users)"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_COMMON_DOCUMENTS
    Case "dokumentvorlagen"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_TEMPLATES
    Case "dokumentvorlagen (all users)"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_COMMON_TEMPLATES
    Case "drucker"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_PRINTERS
    Case "druckumgebung"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_PRINTHOOD
    Case "eigene bilder"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_MYPICTURES
    Case "eigene bilder (all users)"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_COMMON_PICTURES
    Case "eigene dateien"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_MYDOCUMENTS
    Case "eigene dateien (ordner)"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_PERSONAL
    Case "eigene musik"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_MYMUSIC
    Case "eigene musik (all users)"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_COMMON_MUSIC
    Case "eigene videos"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_MYVIDEO
    Case "eigene videos (all users)"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_COMMON_VIDEO
    Case "favoriten"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_FAVORITES
    Case "favoriten (all users)"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_COMMON_FAVORITES
    Case "gemeinsame dateien"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_PROGRAM_FILES_COMMON
    Case "gemeinsame dateien (risc-systeme)"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_PROGRAM_FILES_COMMONX86
    Case "globaler autostart"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_ALTSTARTUP
    Case "globaler autostart (all users)"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_COMMON_ALTSTARTUP
    Case "internet"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_INTERNET
    Case "internet-cache"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_INTERNET_CACHE
    Case "lokale anwendungsdaten"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_LOCAL_APPDATA
    Case "lokalisierte system-resourcen (themes etc.)"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_RESOURCES_LOCALIZED
    Case "netzwerk- und dfü-verbindungen"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_CONNECTIONS
    Case "netzwerkumgebung"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_NETWORK
    Case "netzwerkumgebung (ordner)"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_NETHOOD
    Case "oem-software-links (all users)"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_COMMON_OEM_LINKS
    Case "papierkorb"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_BITBUCKET
    Case "programme"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_PROGRAM_FILES
    Case "programme (risc-systeme)"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_PROGRAM_FILESX86
    Case "programme (startmenü)"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_PROGRAMS
    Case "programme (startmenü, all users)"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_COMMON_PROGRAMS
    Case "schriftarten"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_FONTS
    Case "senden an"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_SENDTO
    Case "startmenü"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_STARTMENU
    Case "startmenü (all users)"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_COMMON_STARTMENU
    Case "system"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_SYSTEM
    Case "system (risc-systeme)"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_SYSTEMX86
    Case "system-resourcen (themes etc.)"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_RESOURCES
    Case "systemsteuerung"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_CONTROLS
    Case "verlauf"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_HISTORY
    Case "windows"
      MenuItemToCSIDL = CSIDLConstants.CSIDL_WINDOWS
    Case Else
      MenuItemToCSIDL = -1
  End Select
End Function

' gibt die pIDL von <Path> zurück
' <Path> muß relativ zum Desktop sein
Function PathTopIDL(ByVal path As String) As Long
  Dim ret As Long

  ret = ILCreateFromPathAsLong(StrPtr(path))
  If ret = 0 Then
    ' try the ANSI version
    ret = ILCreateFromPath(path)
  End If
  PathTopIDL = ret
End Function

' prüft, ob <pIDL> die ItemID des Arbeitsplatzes enthält
#If Debuging Then
  Function pIDLIncludesMyComputer(debugger As clsDebugger, pIDL As Long, Optional freepIDL As Boolean = False) As Boolean
#Else
  Function pIDLIncludesMyComputer(pIDL As Long, Optional freepIDL As Boolean = False) As Boolean
#End If
  Dim pIDLMyComputer As Long

  If pIDL = 0 Then Exit Function

  pIDLMyComputer = CSIDLTopIDL(CSIDLConstants.CSIDL_DRIVES)
  pIDLIncludesMyComputer = ILIsParent(pIDLMyComputer, pIDL, 0)
  #If Debuging Then
    FreeItemIDList debugger, "pIDLIncludesMyComputer #1", pIDLMyComputer
    If freepIDL Then FreeItemIDList debugger, "pIDLIncludesMyComputer #2", pIDL
  #Else
    FreeItemIDList pIDLMyComputer
    If freepIDL Then FreeItemIDList pIDL
  #End If
End Function

' ermittelt den DisplayName von <pIDLToRelative>
' <Flags> gibt die Art an (Pfad, für Adressleiste...)
#If Debuging Then
  Function pIDLToDisplayName(debugger As clsDebugger, IRelative As IVBShellFolder, pIDLToRelative As Long, Optional ByVal Flags As SHGDNConstants = 0, Optional ByVal freepIDL As Boolean = False) As String
#Else
  Function pIDLToDisplayName(IRelative As IVBShellFolder, pIDLToRelative As Long, Optional ByVal Flags As SHGDNConstants = 0, Optional ByVal freepIDL As Boolean = False) As String
#End If
  Dim DispName As STRRET

  #If Debuging Then
    If IRelative Is Nothing Then
      debugger.AddLogEntry "pIDLToDisplayName: IRelative = Nothing - leaving function", LogEntryTypeConstants.letWarning
      Exit Function
    End If
    If pIDLToRelative = 0 Then
      debugger.AddLogEntry "pIDLToDisplayName: pIDLToRelative = 0 - leaving function", LogEntryTypeConstants.letWarning
      Exit Function
    End If
  #End If

  DispName.uType = STRRET_OFFSET
  If IRelative.GetDisplayNameOf(pIDLToRelative, Flags, DispName) = S_OK Then
    pIDLToDisplayName = STRRETToString(DispName, pIDLToRelative)
  End If

  If freepIDL Then
    #If Debuging Then
      FreeItemIDList debugger, "pIDLToDisplayName", pIDLToRelative
    #Else
      FreeItemIDList pIDLToRelative
    #End If
  End If
End Function

' ermittelt den DisplayName von <pIDLToDesktop>
' es wird immer der FriendlyName zurückgegeben
#If Debuging Then
  Function pIDLToDisplayName_Light(debugger As clsDebugger, pIDLToDesktop As Long, Optional freepIDL As Boolean = False) As String
#Else
  Function pIDLToDisplayName_Light(pIDLToDesktop As Long, Optional freepIDL As Boolean = False) As String
#End If
  Dim Data As SHFILEINFO

  #If Debuging Then
    If pIDLToDesktop = 0 Then
      debugger.AddLogEntry "pIDLToDisplayName_Light: pIDLToDesktop = 0 - leaving function", LogEntryTypeConstants.letWarning
      Exit Function
    End If
  #End If

  SHGetFileInfoAsLong pIDLToDesktop, 0, Data, LenB(Data), SHGFI_PIDL Or SHGFI_DISPLAYNAME
  pIDLToDisplayName_Light = Left$(Data.szDisplayName, lstrlenA(Data.szDisplayName))
  If freepIDL Then
    #If Debuging Then
      FreeItemIDList debugger, "pIDLToDisplayName_Light", pIDLToDesktop
    #Else
      FreeItemIDList pIDLToDesktop
    #End If
  End If
End Function

' gibt den Parsing-Pfad für <pIDLToRelative> zurück
' für Items des Dateisystems ist dieser gleichzeitig der volle Pfad
#If Debuging Then
  Function pIDLToPath(debugger As clsDebugger, IRelative As IVBShellFolder, pIDLToRelative As Long, Optional freepIDL As Boolean = False) As String
#Else
  Function pIDLToPath(IRelative As IVBShellFolder, pIDLToRelative As Long, Optional freepIDL As Boolean = False) As String
#End If
  #If Debuging Then
    If IRelative Is Nothing Then
      debugger.AddLogEntry "pIDLToPath: IRelative = Nothing - leaving function", LogEntryTypeConstants.letWarning
      Exit Function
    End If
    If pIDLToRelative = 0 Then
      debugger.AddLogEntry "pIDLToPath: pIDLToRelative = 0 - leaving function", LogEntryTypeConstants.letWarning
      Exit Function
    End If
  #End If

  #If Debuging Then
    pIDLToPath = pIDLToDisplayName(debugger, IRelative, pIDLToRelative, SHGDNConstants.SHGDN_FORPARSING Or SHGDNConstants.SHGDN_NORMAL)
    If freepIDL Then
      FreeItemIDList debugger, "pIDLToPath", pIDLToRelative
    End If
  #Else
    pIDLToPath = pIDLToDisplayName(IRelative, pIDLToRelative, SHGDNConstants.SHGDN_FORPARSING Or SHGDNConstants.SHGDN_NORMAL)
    If freepIDL Then
      FreeItemIDList pIDLToRelative
    End If
  #End If
End Function

' gibt den Parsing-Pfad für <pIDLToDesktop> zurück
' für Items des Dateisystems ist dieser gleichzeitig der volle Pfad
#If Debuging Then
  Function pIDLToPath_Light(debugger As clsDebugger, pIDLToDesktop As Long, Optional freepIDL As Boolean = False) As String
#Else
  Function pIDLToPath_Light(pIDLToDesktop As Long, Optional freepIDL As Boolean = False) As String
#End If
  Dim ret As String

  If pIDLToDesktop = 0 Then Exit Function

  ret = String$(MAX_PATH, Chr$(0))
  If SHGetPathFromIDList(pIDLToDesktop, ret) Then
    pIDLToPath_Light = Left$(ret, lstrlenA(ret))
  Else
    ' <SHGetPathFromIDList> funktioniert nur für FS-Items
    #If Debuging Then
      pIDLToPath_Light = pIDLToPath(debugger, IDesktop, pIDLToDesktop)
    #Else
      pIDLToPath_Light = pIDLToPath(IDesktop, pIDLToDesktop)
    #End If
  End If

  If freepIDL Then
    #If Debuging Then
      FreeItemIDList debugger, "pIDLToPath_Light", pIDLToDesktop
    #Else
      FreeItemIDList pIDLToDesktop
    #End If
  End If
End Function

' benennt <pIDLToParent> um und gibt die neue pIDL zurück
Function RenamepIDL(hWndShellUIParentWindow As Long, IParent As IVBShellFolder, pIDLToParent As Long, ByVal NewName As String) As Long
  Dim newpIDL As Long
  Dim ret As Long

  If IParent Is Nothing Then Exit Function
  If pIDLToParent = 0 Then Exit Function

  NewName = StrConv(NewName, VbStrConv.vbUnicode)
  ret = IParent.SetNameOf(hWndShellUIParentWindow, pIDLToParent, NewName, MakeDWord(SHGDNConstants.SHGDN_NORMAL, SHGDNConstants.SHGDN_NORMAL), newpIDL)

  ' bei Laufwerken bleibt die pIDL gleich
  If newpIDL = 0 Then newpIDL = pIDLToParent
  If ret = 0 Then RenamepIDL = newpIDL
End Function

Sub SetShellIconSize(ByVal cxIcons As Long)
  Const HWND_BROADCAST = &HFFFF&
  Const SMTO_ABORTIFHUNG = &H2
  Const SMTO_NORMAL = &H0
  Const WM_WININICHANGE = &H1A
  Const WM_SETTINGCHANGE = WM_WININICHANGE
  Dim hKey As Long

'  SHSetValue HKEY_CURRENT_USER, "Control Panel\Desktop\WindowMetrics", "Shell Icon Size", REG_SZ, ByVal CStr(cxIcons), Len(CStr(cxIcons))
  If RegCreateKeyExAsLong(HKEY_CURRENT_USER, "Control Panel\Desktop\WindowMetrics", 0, "", 0, KEY_SET_VALUE, 0, hKey, 0) = ERROR_SUCCESS Then
    RegSetValueEx hKey, "Shell Icon Size", 0, REG_SZ, ByVal StrPtr(StrConv(CStr(cxIcons), VbStrConv.vbFromUnicode)), Len(CStr(cxIcons))
    RegCloseKey hKey
  End If

'  SendMessageTimeout HWND_BROADCAST, WM_SETTINGCHANGE, SPI_SETNONCLIENTMETRICS, StrPtr("WindowMetrics"), SMTO_NORMAL Or SMTO_ABORTIFHUNG, 10000, 0
End Sub

' gibt zurück, ob das Laufwerk <pIDLToMyComputer> angezeigt wird oder nicht
#If Debuging Then
  Function ShouldShowDrive(debugger As clsDebugger, IMyComputer As IVBShellFolder, pIDLToMyComputer As Long) As Boolean
#Else
  Function ShouldShowDrive(IMyComputer As IVBShellFolder, pIDLToMyComputer As Long) As Boolean
#End If
  Dim allDrives As Long
  Dim Flags As SHCONTFConstants
  Dim hKey As Long
  Dim IDDrive As Long
  Dim IEnum As IVBEnumIDList
  Dim Path1 As String
  Dim Path2 As String
  Dim pIDLMyComputer_ToDesktop As Long
  Dim pIDLSubItem_ToParent As Long
  Dim ret As Boolean
  Dim ret2 As Long
  Dim Size As Long

  If IMyComputer Is Nothing Then Exit Function
  If pIDLToMyComputer = 0 Then Exit Function

  #If Debuging Then
    Path1 = pIDLToPath(debugger, IMyComputer, pIDLToMyComputer)
  #Else
    Path1 = pIDLToPath(IMyComputer, pIDLToMyComputer)
  #End If

  ' zunächst prüfen, ob das Laufwerk ausgeblendet ist
  Size = LenB(allDrives)
  If RegCreateKeyExAsLong(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", 0, "", 0, KEY_QUERY_VALUE, 0, hKey, 0) = ERROR_SUCCESS Then
    ret = RegQueryValueEx(hKey, "NoDrives", 0, REG_DWORD_LITTLE_ENDIAN, ByVal VarPtr(allDrives), Size)
    If ret2 <> ERROR_SUCCESS Then allDrives = 0
    RegCloseKey hKey
  End If

  IDDrive = 2 ^ (Asc(UCase$(Left$(Path1, 1))) - 65)
  ret = ((allDrives And IDDrive) = 0)

  If ret Then
    ret = False
    ' es kann auch deaktiviert sein, ohne vorher versteckt worden zu sein
    ' -> den Arbeitsplatz nach <pIDLToParent> durchsuchen

    ' Suche initiieren...
    pIDLMyComputer_ToDesktop = CSIDLTopIDL(CSIDLConstants.CSIDL_DRIVES)
    Flags = SHCONTFConstants.SHCONTF_FOLDERS Or SHCONTFConstants.SHCONTF_NONFOLDERS Or SHCONTFConstants.SHCONTF_INCLUDEHIDDEN
    IMyComputer.EnumObjects 0, Flags, IEnum
    If Not (IEnum Is Nothing) Then
      ' ...und starten
      While (IEnum.Next(1, pIDLSubItem_ToParent, 0) = S_OK) And (ret = False)
        #If Debuging Then
          Path2 = pIDLToPath(debugger, IMyComputer, pIDLSubItem_ToParent, True)
        #Else
          Path2 = pIDLToPath(IMyComputer, pIDLSubItem_ToParent, True)
        #End If
        ret = (LCase$(Path1) = LCase$(Path2))
      Wend
    End If
    Set IEnum = Nothing

    #If Debuging Then
      FreeItemIDList debugger, "ShouldShowDrive", pIDLMyComputer_ToDesktop
    #Else
      FreeItemIDList pIDLMyComputer_ToDesktop
    #End If
  End If

  ShouldShowDrive = ret
End Function

' prüft anhand der Filterkriterien von <Ctl>, ob <pIDLToParent> angezeigt werden soll
#If Debuging Then
  Function ShouldShowItem(debugger As clsDebugger, Ctl As ExplorerTreeView, IParent As IVBShellFolder, pIDLToParent As Long, pIDLToDesktop As Long, hParentItem As Long) As Boolean
#Else
  Function ShouldShowItem(Ctl As ExplorerTreeView, IParent As IVBShellFolder, pIDLToParent As Long, pIDLToDesktop As Long, hParentItem As Long) As Boolean
#End If
  Dim arrFilters() As String
  Dim attributes As AttributesConstants
  Dim FileFilters As String
  Dim FileName As String
  Dim FolderFilters As String
  Dim i As Integer
  '---zum Vermeiden mehrfacher Aufrufe einer Funktion---
  Dim FileAttr As Long
  Dim itemAttr As SFGAOConstants
  Dim ItemIsCompressed As Boolean
  Dim ItemIsFile As Boolean
  Dim ItemIsFolder As Boolean
  Dim ItemIsHidden As Boolean
  Dim ItemIsPartOfFS As Boolean
  Dim ItemIsSystemItem As Boolean
  '---zum Vermeiden mehrfacher Aufrufe einer Funktion---
  Dim ret As Boolean
  Dim showFSFiles As Boolean
  Dim showFSFolders As Boolean
  Dim showIt As Boolean
  Dim showNonFSFiles As Boolean
  Dim showNonFSFolders As Boolean

  If IParent Is Nothing Then Exit Function
  If pIDLToParent = 0 Then Exit Function
  If Ctl Is Nothing Then Exit Function
  If hParentItem = -1 Then
    #If Debuging Then
      debugger.AddLogEntry "ShouldShowItem() has been called with hParentItem=0xFFFFFFFF", LogEntryTypeConstants.letInfo
    #End If
    ' don't apply filter rules to the root item
    ret = True
    GoTo Ende
  End If

  FileFilters = IIf(Ctl.UseFileFilters, Ctl.FileFilters, "")
  FolderFilters = IIf(Ctl.UseFolderFilters, Ctl.FolderFilters, "")
  showFSFiles = (Ctl.IncludedItems And IncludedItemsConstants.iiFSFiles)
  showFSFolders = (Ctl.IncludedItems And IncludedItemsConstants.iiFSFolders)
  showNonFSFiles = (Ctl.IncludedItems And IncludedItemsConstants.iiNonFSFiles)
  showNonFSFolders = (Ctl.IncludedItems And IncludedItemsConstants.iiNonFSFolders)

  ' soll dieser Typ überhaupt angezeigt werden?
  #If Debuging Then
    itemAttr = GetAttributes(debugger, IParent, pIDLToParent, SFGAOConstants.SFGAO_FILESYSTEM Or SFGAOConstants.SFGAO_FOLDER)
  #Else
    itemAttr = GetAttributes(IParent, pIDLToParent, SFGAOConstants.SFGAO_FILESYSTEM Or SFGAOConstants.SFGAO_FOLDER)
  #End If
  If itemAttr And SFGAOConstants.SFGAO_FILESYSTEM Then
    ItemIsPartOfFS = True
    If itemAttr And SFGAOConstants.SFGAO_FOLDER Then
      ' ist es wirklich ein Ordner?
      #If Debuging Then
        If FileExists_pIDL(debugger, IParent, pIDLToParent) Then
      #Else
        If FileExists_pIDL(IParent, pIDLToParent) Then
      #End If
        If Not showFSFiles Then
          ret = False
          GoTo Ende
        End If
        ItemIsFile = True
      Else
        If Not showFSFolders Then
          ret = False
          GoTo Ende
        End If
        ItemIsFolder = True
      End If
    Else
      If Not showFSFiles Then
        ret = False
        GoTo Ende
      End If
      ItemIsFile = True
    End If
  Else
    If itemAttr And SFGAOConstants.SFGAO_FOLDER Then
      ItemIsFolder = True
      ' da die Filter nur für Items des Dateisystems gelten, können NonFS-Items gleich behandelt
      ' werden
      #If Debuging Then
        ret = IIf(IsMyComputer(debugger, IParent) And Ctl.DrivesOnly, False, showNonFSFolders)
      #Else
        ret = IIf(IsMyComputer(IParent) And Ctl.DrivesOnly, False, showNonFSFolders)
      #End If
      GoTo Ende
    Else
      ItemIsFile = True
      ' da die Filter nur für Items des Dateisystems gelten, können NonFS-Items gleich behandelt
      ' werden
      #If Debuging Then
        ret = IIf(IsMyComputer(debugger, IParent) And Ctl.DrivesOnly, False, showNonFSFiles)
      #Else
        ret = IIf(IsMyComputer(IParent) And Ctl.DrivesOnly, False, showNonFSFiles)
      #End If
      GoTo Ende
    End If
  End If

  '------------------------------------------------------------------------------------------
  ' <pIDLToParent> ist ein Element des Dateisystems
  ' -> die Unterteilung in FS- und NonFS-Items kann zu Gunsten der Performance wegfallen
  '------------------------------------------------------------------------------------------

  #If Debuging Then
    If IsDrive(debugger, IParent, pIDLToParent) Then
      ret = ShouldShowDrive(debugger, IParent, pIDLToParent)
      GoTo Ende
    ElseIf IsMyComputer(debugger, IParent) Then
  #Else
    If IsDrive(IParent, pIDLToParent) Then
      ret = ShouldShowDrive(IParent, pIDLToParent)
      GoTo Ende
    ElseIf IsMyComputer(IParent) Then
  #End If
    If Ctl.DrivesOnly Then
      ret = False
      GoTo Ende
    End If
  End If

  If ItemIsFile Or ItemIsFolder Then
    ' Filter durchlaufen
    arrFilters = Split(IIf(ItemIsFile, FileFilters, FolderFilters), "|")
    If UBound(arrFilters) >= 0 Then
      #If Debuging Then
        FileName = pIDLToDisplayName(debugger, IParent, pIDLToParent, SHGDNConstants.SHGDN_INFOLDER Or SHGDNConstants.SHGDN_FORPARSING)
      #Else
        FileName = pIDLToDisplayName(IParent, pIDLToParent, SHGDNConstants.SHGDN_INFOLDER Or SHGDNConstants.SHGDN_FORPARSING)
      #End If
      For i = LBound(arrFilters) To UBound(arrFilters)
        If Left$(arrFilters(i), 1) = "/" Then
          showIt = Not (FileName Like Mid$(arrFilters(i), 2))
          If Not showIt Then Exit For
        Else
          showIt = FileName Like arrFilters(i)
          If showIt Then Exit For
        End If
      Next
      If Not showIt Then
        ' dieser Item wurde herausgefiltert
        ret = False
        GoTo Ende
      End If
    End If
  End If

  attributes = IIf(ItemIsFile, Ctl.FileAttributes, Ctl.FolderAttributes)

  ' jetzt die Attribute prüfen
  If attributes = (AttributesConstants.attReadOnly Or AttributesConstants.attHidden Or AttributesConstants.attArchive Or AttributesConstants.attSystem Or AttributesConstants.attEncrypted Or AttributesConstants.attCompressed) Then
    ' es wird kein Attribut herausgefiltert
    ' -> Item anzeigen
    ret = True
    GoTo Ende
  End If

  ' diese Attribute erfordern bei Laufwerken anscheinend einen Laufwerkszugriff,
  ' deshalb holen wir sie erst jetzt
  #If Debuging Then
    itemAttr = GetAttributes(debugger, IParent, pIDLToParent, SFGAOConstants.SFGAO_HIDDEN Or SFGAOConstants.SFGAO_COMPRESSED Or SFGAOConstants.SFGAO_READONLY Or SFGAOConstants.SFGAO_ENCRYPTED)
  #Else
    itemAttr = GetAttributes(IParent, pIDLToParent, SFGAOConstants.SFGAO_HIDDEN Or SFGAOConstants.SFGAO_COMPRESSED Or SFGAOConstants.SFGAO_READONLY Or SFGAOConstants.SFGAO_ENCRYPTED)
  #End If

  ' wenn der Item versteckt oder das System-Attribut hat und solche Items nicht angezeigt werden sollen,
  ' ohne Rücksicht auf die anderen Attribute False zurückgeben
  If itemAttr And SFGAOConstants.SFGAO_HIDDEN Then
    If (attributes And AttributesConstants.attHidden) = 0 Then
      ret = False
      GoTo Ende
    End If
    ItemIsHidden = True
  End If
  #If Debuging Then
    FileAttr = GetFileAttribs(debugger, IParent, pIDLToParent, FILE_ATTRIBUTE_SYSTEM Or FILE_ATTRIBUTE_ARCHIVE)
  #Else
    FileAttr = GetFileAttribs(IParent, pIDLToParent, FILE_ATTRIBUTE_SYSTEM Or FILE_ATTRIBUTE_ARCHIVE)
  #End If
  If FileAttr And FILE_ATTRIBUTE_SYSTEM Then
    If (attributes And AttributesConstants.attSystem) = 0 Then
      ret = False
      GoTo Ende
    End If
    ItemIsSystemItem = True
  End If
  ' das gleiche gilt für komprimierte Items
  If itemAttr And SFGAOConstants.SFGAO_COMPRESSED Then
    If (attributes And AttributesConstants.attCompressed) = 0 Then
      ret = False
      GoTo Ende
    End If
    ItemIsCompressed = True
  End If

  ' wenn 1 Attribut des Items angezeigt werden soll, sofort True zurückgeben
  If ItemIsHidden Then
    If attributes And AttributesConstants.attHidden Then
      ret = True
      GoTo Ende
    End If
  End If
  If ItemIsSystemItem Then
    If attributes And AttributesConstants.attSystem Then
      ret = True
      GoTo Ende
    End If
  End If
  If ItemIsCompressed Then
    If attributes And AttributesConstants.attCompressed Then
      ret = True
      GoTo Ende
    End If
  End If
  If itemAttr And SFGAOConstants.SFGAO_READONLY Then
    If attributes And AttributesConstants.attReadOnly Then
      ret = True
      GoTo Ende
    End If
  End If
  If itemAttr And SFGAOConstants.SFGAO_ENCRYPTED Then
    If attributes And AttributesConstants.attEncrypted Then
      ret = True
      GoTo Ende
    End If
  End If
  If FileAttr And FILE_ATTRIBUTE_ARCHIVE Then
    If attributes And AttributesConstants.attArchive Then ret = True
  Else
    ' der Item hat keine Attribute, die herausgefiltert werden könnten
    ' -> Item anzeigen
    ret = True
  End If

Ende:
  If ret Then
    If ItemIsPartOfFS Then
      If ItemIsFolder Then
        If Ctl.FireBeforeInsertItem And FireBeforeInsertItemConstants.fbiiForFSFolders Then
          ret = False
          Ctl.RaiseBeforeInsertItem pIDLToDesktop, hParentItem, ret
          ret = Not ret
        End If
      ElseIf ItemIsFile Then
        If Ctl.FireBeforeInsertItem And FireBeforeInsertItemConstants.fbiiForFSFiles Then
          ret = False
          Ctl.RaiseBeforeInsertItem pIDLToDesktop, hParentItem, ret
          ret = Not ret
        End If
      End If
    Else
      If ItemIsFolder Then
        If Ctl.FireBeforeInsertItem And FireBeforeInsertItemConstants.fbiiForNonFSFolders Then
          ret = False
          Ctl.RaiseBeforeInsertItem pIDLToDesktop, hParentItem, ret
          ret = Not ret
        End If
      ElseIf ItemIsFile Then
        If Ctl.FireBeforeInsertItem And FireBeforeInsertItemConstants.fbiiForNonFSFiles Then
          ret = False
          Ctl.RaiseBeforeInsertItem pIDLToDesktop, hParentItem, ret
          ret = Not ret
        End If
      End If
    End If
  End If

  ShouldShowItem = ret
End Function

' prüft anhand der Filterkriterien von <Ctl>, ob <DispName> angezeigt werden soll
Function ShouldShowItem_Archive(Ctl As ExplorerTreeView, ByVal DispName As String, ByVal ItemAttributes As Long) As Boolean
  Dim arrFilters() As String
  Dim attributes As AttributesConstants
  Dim FileFilters As String
  Dim FolderFilters As String
  Dim i As Integer
  Dim ItemIsArchiveItem As Boolean
  Dim ItemIsCompressed As Boolean
  Dim ItemIsEncrypted As Boolean
  Dim ItemIsFile As Boolean
  Dim ItemIsFolder As Boolean
  Dim ItemIsHidden As Boolean
  Dim ItemIsReadOnly As Boolean
  Dim ItemIsSystemItem As Boolean
  Dim ret As Boolean
  Dim showFSFiles As Boolean
  Dim showFSFolders As Boolean
  Dim showIt As Boolean

  If DispName = "" Then Exit Function
  If Ctl Is Nothing Then Exit Function

  FileFilters = IIf(Ctl.UseFileFilters, Ctl.FileFilters, "")
  FolderFilters = IIf(Ctl.UseFolderFilters, Ctl.FolderFilters, "")
  ItemIsArchiveItem = ItemAttributes And FILE_ATTRIBUTE_ARCHIVE
  ItemIsCompressed = ItemAttributes And FILE_ATTRIBUTE_COMPRESSED
  ItemIsEncrypted = ItemAttributes And FILE_ATTRIBUTE_ENCRYPTED
  ItemIsFolder = ItemAttributes And FILE_ATTRIBUTE_DIRECTORY
  ItemIsFile = Not ItemIsFolder
  ItemIsHidden = ItemAttributes And FILE_ATTRIBUTE_HIDDEN
  ItemIsReadOnly = ItemAttributes And FILE_ATTRIBUTE_READONLY
  ItemIsSystemItem = ItemAttributes And FILE_ATTRIBUTE_SYSTEM

  showFSFiles = (Ctl.IncludedItems And IncludedItemsConstants.iiFSFiles)
  showFSFolders = (Ctl.IncludedItems And IncludedItemsConstants.iiFSFolders)

  ' soll dieser Typ überhaupt angezeigt werden?
  Select Case True
    Case ItemIsFolder
      If Not showFSFolders Then
        ret = False
        GoTo Ende
      End If
    Case ItemIsFile
      If Not showFSFiles Then
        ret = False
        GoTo Ende
      End If
  End Select

  ' Filter durchlaufen
  arrFilters = Split(IIf(ItemIsFile, FileFilters, FolderFilters), "|")
  If UBound(arrFilters) >= 0 Then
    For i = LBound(arrFilters) To UBound(arrFilters)
      If Left$(arrFilters(i), 1) = "/" Then
        If DispName Like Mid$(arrFilters(i), 2) Then showIt = False
      Else
        If DispName Like arrFilters(i) Then showIt = True
      End If
    Next
    If Not showIt Then
      ' dieser Item wurde herausgefiltert
      ret = False
      GoTo Ende
    End If
  End If

  attributes = IIf(ItemIsFile, Ctl.FileAttributes, Ctl.FolderAttributes)

  ' jetzt die Attribute prüfen
  If attributes = (AttributesConstants.attReadOnly Or AttributesConstants.attHidden Or AttributesConstants.attArchive Or AttributesConstants.attSystem Or AttributesConstants.attEncrypted Or AttributesConstants.attCompressed) Then
    ' es wird kein Attribut herausgefiltert
    ' -> Item anzeigen
    ret = True
    GoTo Ende
  End If

  ' wenn der Item versteckt oder das System-Attribut hat und solche Items nicht angezeigt werden sollen,
  ' ohne Rücksicht auf die anderen Attribute False zurückgeben
  If ItemIsHidden Then
    If (attributes And AttributesConstants.attHidden) = 0 Then
      ret = False
      GoTo Ende
    End If
  End If
  If ItemIsSystemItem Then
    If (attributes And AttributesConstants.attSystem) = 0 Then
      ret = False
      GoTo Ende
    End If
  End If
  ' das gleiche gilt für komprimierte Items
  If ItemIsCompressed Then
    If (attributes And AttributesConstants.attCompressed) = 0 Then
      ret = False
      GoTo Ende
    End If
  End If

  ' wenn 1 Attribut des Items angezeigt werden soll, sofort True zurückgeben
  If ItemIsHidden Then
    If attributes And AttributesConstants.attHidden Then
      ret = True
      GoTo Ende
    End If
  End If
  If ItemIsSystemItem Then
    If attributes And AttributesConstants.attSystem Then
      ret = True
      GoTo Ende
    End If
  End If
  If ItemIsCompressed Then
    If attributes And AttributesConstants.attCompressed Then
      ret = True
      GoTo Ende
    End If
  End If
  If ItemIsReadOnly Then
    If attributes And AttributesConstants.attReadOnly Then
      ret = True
      GoTo Ende
    End If
  End If
  If ItemIsEncrypted Then
    If attributes And AttributesConstants.attEncrypted Then
      ret = True
      GoTo Ende
    End If
  End If
  If ItemIsArchiveItem Then
    If attributes And AttributesConstants.attArchive Then ret = True
  Else
    ' der Item hat keine Attribute, die herausgefiltert werden könnten
    ' -> Item anzeigen
    ret = True
  End If

Ende:
  ShouldShowItem_Archive = ret
End Function

' gibt zurück, ob <OverlayIndex> zu den Overlays gehört, die nach <ShownOverlays> angezeigt
' werden sollen
Function ShouldShowOverlay(ShownOverlays As Long, overlayIndex As Long) As Boolean
  Dim ret As Boolean

  ' einfach die 3 Standard-Overlays durchgehen und prüfen, ob die Indizes übereinstimmen
  Select Case overlayIndex
    Case 0
      ret = False
    Case OVERLAY_LINK
      ret = (ShownOverlays And ShownOverlaysConstants.soLink)
    Case OVERLAY_SHARE
      ret = (ShownOverlays And ShownOverlaysConstants.soSharedItem)
    Case OVERLAY_SLOWFILE
      ret = (ShownOverlays And ShownOverlaysConstants.soSlowFile)
    Case Else
      ret = (ShownOverlays And ShownOverlaysConstants.soOthers)
  End Select

  ShouldShowOverlay = ret
End Function

#If Debuging Then
  Function SimplePIDLToRealPIDL(debugger As clsDebugger, IParent As IVBShellFolder, pIDLSimple As Long, Optional ByVal freepIDL As Boolean = False) As Long
#Else
  Function SimplePIDLToRealPIDL(IParent As IVBShellFolder, pIDLSimple As Long, Optional ByVal freepIDL As Boolean = False) As Long
#End If
  Dim ret As Long

  If SHGetRealIDL(IParent, pIDLSimple, ret) >= 0 Then
    SimplePIDLToRealPIDL = ret
    If freepIDL Then
      #If Debuging Then
        FreeItemIDList debugger, "SimplePIDLToRealPIDL", pIDLSimple
      #Else
        FreeItemIDList pIDLSimple
      #End If
    End If
  Else
    SimplePIDLToRealPIDL = ILClone(pIDLSimple)
  End If
End Function

#If Debuging Then
  Sub SplitFullyQualifiedPIDL(debugger As clsDebugger, pIDLToDesktop As Long, IParent As IVBShellFolder, pIDLToParent As Long)
#Else
  Sub SplitFullyQualifiedPIDL(pIDLToDesktop As Long, IParent As IVBShellFolder, pIDLToParent As Long)
#End If
  Dim itemIDSize As Integer
  Dim pIDL_Desktop As Long

  #If Debuging Then
    If pIDLToDesktop = 0 Then
      debugger.AddLogEntry "SplitFullyQualifiedPIDL: pIDLToDesktop = 0 - leaving sub", LogEntryTypeConstants.letWarning
      Exit Sub
    End If
  #End If

  pIDL_Desktop = CSIDLTopIDL(CSIDLConstants.CSIDL_DESKTOP)
  If ILIsEqual(pIDLToDesktop, pIDL_Desktop) Then
    Set IParent = IDesktop
    pIDLToParent = pIDLToDesktop
  Else
    If ver_Shell32_50 Then
      SHBindToParent pIDLToDesktop, IID_IShellFolder, IParent, VarPtr(pIDLToParent)
    Else
      pIDLToParent = ILClone(pIDLToDesktop)
      ILRemoveLastID pIDLToParent
      CopyMemory VarPtr(itemIDSize), pIDLToParent, LenB(itemIDSize)
      If itemIDSize = 0 Then
        Set IParent = IDesktop
      Else
        IDesktop.BindToObject pIDLToParent, 0, IID_IShellFolder, IParent
      End If
      #If Debuging Then
        FreeItemIDList debugger, "SplitFullyQualifiedPIDL #1", pIDLToParent
      #Else
        FreeItemIDList pIDLToParent
      #End If
      pIDLToParent = ILFindLastID(pIDLToDesktop)
    End If
    ' the next line shouldn't be necessary, but Nero Scout may crash the control via AutoUpdate if it is missing
    ' 30.12.2006: deactivated, because printing a document crashes the control via AutoUpdate otherwise
    'IParent.AddRef
  End If
  #If Debuging Then
    FreeItemIDList debugger, "SplitFullyQualifiedPIDL #2", pIDL_Desktop
  #Else
    FreeItemIDList pIDL_Desktop
  #End If
End Sub

Function STRRETToString(Data As STRRET, ByVal pIDL As Long) As String
  Dim ret As String
  Dim tmp As Long

  ret = String$(MAX_PATH, Chr$(0))
'  If ver_Win_XP Then
'    ' StrRetToBSTR seems to allocate a new string for 'ret' without freeing the old one
'    StrRetToBSTR Data, pIDL, VarPtr(ret)
'    STRRETToString = ret
'  ElseIf ver_Shell32_50 Then
  If ver_Shell32_50 Then
    StrRetToBuf Data, pIDL, ret, Len(ret)
    STRRETToString = Left$(ret, lstrlenA(ret))
  Else
    With Data
      Select Case .uType
        Case STRRETConstants.STRRET_CSTR
          lstrcpyAsLong2 ret, VarPtr(.Data(0))
        Case STRRETConstants.STRRET_OFFSET
          If pIDL Then
            CopyMemory VarPtr(tmp), VarPtr(.Data(0)), LenB(tmp)
            lstrcpyAsLong2 ret, pIDL + tmp
          End If
        Case STRRETConstants.STRRET_WSTR
          CopyMemory VarPtr(tmp), VarPtr(.Data(0)), LenB(tmp)
          WideCharToMultiByte CP_ACP, 0, tmp, -1, ret, Len(ret), vbNullString, 0
          CoTaskMemFree tmp
      End Select
    End With
    STRRETToString = Left$(ret, lstrlenA(ret))
  End If
End Function

Sub UpdateDefaultIconIndices()
  Dim Flags As Long
  Dim Data As SHFILEINFO

  Flags = SHGFI_USEFILEATTRIBUTES Or SHGFI_SYSICONINDEX Or SHGFI_SMALLICON
  SHGetFileInfo ".zyxwv12", FILE_ATTRIBUTE_NORMAL, Data, LenB(Data), Flags
  DEFICON_BLANKDOC_SMALL = Data.iIcon
  ' ToDo: DEFICON_DOC_SMALL ist eigentlich das Standardicon für Dokumente
  SHGetFileInfo ".zyxwv12", FILE_ATTRIBUTE_NORMAL, Data, LenB(Data), Flags
  DEFICON_DOC_SMALL = Data.iIcon
  SHGetFileInfo ".exe", FILE_ATTRIBUTE_NORMAL, Data, LenB(Data), Flags
  DEFICON_APP_SMALL = Data.iIcon
  SHGetFileInfo "folder", FILE_ATTRIBUTE_DIRECTORY, Data, LenB(Data), Flags
  DEFICON_FOLDER_SMALL = Data.iIcon
  SHGetFileInfo "folder", FILE_ATTRIBUTE_DIRECTORY, Data, LenB(Data), Flags Or SHGFI_OPENICON
  DEFICON_OPENFOLDER_SMALL = Data.iIcon

  Flags = SHGFI_USEFILEATTRIBUTES Or SHGFI_SYSICONINDEX Or SHGFI_LARGEICON
  SHGetFileInfo ".zyxwv12", FILE_ATTRIBUTE_NORMAL, Data, LenB(Data), Flags
  DEFICON_BLANKDOC_LARGE = Data.iIcon
  SHGetFileInfo ".zyxwv12", FILE_ATTRIBUTE_NORMAL, Data, LenB(Data), Flags
  DEFICON_DOC_LARGE = Data.iIcon
  SHGetFileInfo ".exe", FILE_ATTRIBUTE_NORMAL, Data, LenB(Data), Flags
  DEFICON_APP_LARGE = Data.iIcon
  SHGetFileInfo "folder", FILE_ATTRIBUTE_DIRECTORY, Data, LenB(Data), Flags
  DEFICON_FOLDER_LARGE = Data.iIcon
  SHGetFileInfo "folder", FILE_ATTRIBUTE_DIRECTORY, Data, LenB(Data), Flags Or SHGFI_OPENICON
  DEFICON_OPENFOLDER_LARGE = Data.iIcon

  GetDefaultOverlays
End Sub

' gibt zurück, ob der Item <pIDLToParent> noch existiert
Function ValidateItem(IParent As IVBShellFolder, pIDLToParent As Long) As Boolean
  Dim tmp As SFGAOConstants

  If IParent Is Nothing Then Exit Function
  If pIDLToParent = 0 Then Exit Function

  tmp = SFGAOConstants.SFGAO_VALIDATE
  ValidateItem = (IParent.GetAttributesOf(1, pIDLToParent, tmp) = NOERROR)
End Function

' gibt zurück, ob der Item <pIDLToDesktop> noch existiert
#If Debuging Then
  Function ValidateItemFQ(debugger As clsDebugger, pIDLToDesktop As Long) As Boolean
#Else
  Function ValidateItemFQ(pIDLToDesktop As Long) As Boolean
#End If
  Dim IParent As IVBShellFolder
  Dim pIDLToParent As Long
  Dim tmp As SFGAOConstants

  #If Debuging Then
    SplitFullyQualifiedPIDL debugger, pIDLToDesktop, IParent, pIDLToParent
  #Else
    SplitFullyQualifiedPIDL pIDLToDesktop, IParent, pIDLToParent
  #End If
  If (pIDLToParent <> 0) And Not (IParent Is Nothing) Then
    tmp = SFGAOConstants.SFGAO_VALIDATE
    ValidateItemFQ = (IParent.GetAttributesOf(1, pIDLToParent, tmp) = NOERROR)
  End If
  Set IParent = Nothing
End Function
