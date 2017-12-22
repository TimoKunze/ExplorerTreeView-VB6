VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDebugger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

#If Debuging Then
  Private Const FOREGROUND_BLUE = &H1
  Private Const FOREGROUND_GREEN = &H2
  Private Const FOREGROUND_RED = &H4
  Private Const FOREGROUND_INTENSITY = &H8


  Public Enum ConsoleColorConstants
    conBlack = 0&
    conBlue = FOREGROUND_BLUE
    conGreen = FOREGROUND_GREEN
    conCyan = FOREGROUND_BLUE Or FOREGROUND_GREEN
    conRed = FOREGROUND_RED
    conMagenta = FOREGROUND_RED Or FOREGROUND_BLUE
    conYellow = FOREGROUND_RED Or FOREGROUND_GREEN
    conWhite = FOREGROUND_BLUE Or FOREGROUND_GREEN Or FOREGROUND_RED
    conBlackHi = FOREGROUND_INTENSITY
    conBlueHi = FOREGROUND_BLUE Or FOREGROUND_INTENSITY
    conCyanHi = FOREGROUND_BLUE Or FOREGROUND_GREEN Or FOREGROUND_INTENSITY
    conGreenHi = FOREGROUND_GREEN Or FOREGROUND_INTENSITY
    conRedHi = FOREGROUND_RED Or FOREGROUND_INTENSITY
    conMagentaHi = FOREGROUND_RED Or FOREGROUND_BLUE Or FOREGROUND_INTENSITY
    conYellowHi = FOREGROUND_RED Or FOREGROUND_GREEN Or FOREGROUND_INTENSITY
    conWhiteHi = FOREGROUND_BLUE Or FOREGROUND_GREEN Or FOREGROUND_RED Or FOREGROUND_INTENSITY
  End Enum

  Public Enum LogEntryTypeConstants
    letInfo
    letWarning
    letError
    letOther
  End Enum


  Private Type COORD
    x As Integer
    y As Integer
  End Type

  Private Type SMALL_RECT
    Left As Integer
    Top As Integer
    Right As Integer
    Bottom As Integer
  End Type

  Private Type CONSOLE_SCREEN_BUFFER_INFO
    dwSize As COORD
    dwCursorPosition As COORD
    wAttributes As Integer
    srWindow As SMALL_RECT
    dwMaximumWindowSize As COORD
  End Type


  Private m_hConsole As Long
  Private m_hWnd As Long
  Private m_OriginalOutputMode As Long
  Private m_Prefix As String


  Private Declare Function AllocConsole Lib "kernel32.dll" () As Long
  Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
  Private Declare Function CreateFileAsLong Lib "kernel32.dll" Alias "CreateFileA" (ByVal File As String, ByVal AccessMode As Long, ByVal ShareMode As Long, ByVal Security As Long, ByVal CreationFlags As Long, ByVal FileAttributes As Long, ByVal hTemplateFile As Long) As Long
  Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
  Private Declare Function FreeConsole Lib "kernel32.dll" () As Long
  Private Declare Function GetConsoleMode Lib "kernel32.dll" (ByVal hConsoleHandle As Long, ByRef lpMode As Long) As Long
  Private Declare Function GetConsoleScreenBufferInfo Lib "kernel32.dll" (ByVal hConsoleOutput As Long, ByRef lpConsoleScreenBufferInfo As CONSOLE_SCREEN_BUFFER_INFO) As Long
  Private Declare Function GetConsoleTitle Lib "kernel32.dll" Alias "GetConsoleTitleA" (ByVal lpConsoleTitle As String, ByVal nSize As Long) As Long
  Private Declare Function GetConsoleWindow Lib "kernel32.dll" () As Long
  Private Declare Function GetCurrentProcessId Lib "kernel32.dll" () As Long
  Private Declare Function GetStdHandle Lib "kernel32.dll" (ByVal nStdHandle As Long) As Long
  Private Declare Function GetSystemMenu Lib "user32.dll" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
  Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hWnd As Long, ByRef lpdwProcessId As Long) As Long
  Private Declare Function RemoveMenu Lib "user32.dll" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
  Private Declare Function SetConsoleMode Lib "kernel32.dll" (ByVal hConsoleHandle As Long, ByVal dwMode As Long) As Long
  Private Declare Function SetConsoleTextAttribute Lib "kernel32.dll" (ByVal hConsoleOutput As Long, ByVal wAttributes As Long) As Long
  Private Declare Function SetConsoleTitle Lib "kernel32.dll" Alias "SetConsoleTitleA" (ByVal lpConsoleTitle As String) As Long
  Private Declare Function SetFilePointer Lib "kernel32.dll" (ByVal hFile As Long, ByVal DistToMoveLow As Long, DistToMoveHigh As Long, ByVal Flags As Long) As Long
  Private Declare Function ShowWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
  Private Declare Function WriteConsole Lib "kernel32.dll" Alias "WriteConsoleA" (ByVal hConsoleOutput As Long, ByVal lpBuffer As String, ByVal nNumberOfCharsToWrite As Long, ByRef lpNumberOfCharsWritten As Long, ByVal lpReserved As Long) As Long
  Private Declare Function WriteFileAsLong Lib "kernel32.dll" Alias "WriteFile" (ByVal hFile As Long, buffer As Any, ByVal BytesToWrite As Long, BytesWritten As Long, ByVal OverlappedData As Long) As Long


  Private Sub Class_Initialize()
    Const ENABLE_PROCESSED_OUTPUT = &H1
    Const ENABLE_WRAP_AT_EOL_OUTPUT = &H2
    Const MF_BYCOMMAND As Long = &H0&
    Const SC_CLOSE As Long = &HF060&
    Const STD_OUTPUT_HANDLE = -11
    Const SW_SHOW = 5
    Dim hMenu As Long
    Dim pid As Long
    Dim strTitle As String

    #If NoConsole = 0 Then
'    If IsCompiled Then
'      ' TODO: Doesn't work on NT 4.0, Win9x
'      hWnd = GetConsoleWindow
'    Else
      AllocConsole

      ' find window most likely to be our new console
      strTitle = Me.ConsoleTitle
      m_hWnd = FindConsole(strTitle)

      ' confirm by comparing process IDs
      GetWindowThreadProcessId m_hWnd, pid
      If pid = GetCurrentProcessId Then
        ' disable its Close button, because if we accidently press that during debugging, everything goes kaboom!
        hMenu = GetSystemMenu(m_hWnd, False)
        RemoveMenu hMenu, SC_CLOSE, MF_BYCOMMAND
      End If
'    End If

    If m_hWnd Then
      Me.ConsoleTitle = "ExplorerTreeView Debugger"
      ShowWindow m_hWnd, SW_SHOW

      m_hConsole = GetStdHandle(STD_OUTPUT_HANDLE)
      GetConsoleMode m_hConsole, m_OriginalOutputMode
      SetConsoleMode m_hConsole, ENABLE_PROCESSED_OUTPUT Or ENABLE_WRAP_AT_EOL_OUTPUT
    End If
    #End If
  End Sub

  Private Sub Class_Terminate()
    Const INVALID_HANDLE_VALUE = -1

    #If NoConsole = 0 Then
      If m_hConsole <> INVALID_HANDLE_VALUE Then
        SetConsoleMode m_hConsole, m_OriginalOutputMode
        CloseHandle m_hConsole
      End If
      FreeConsole
    #End If
  End Sub


  Public Property Get ConsoleTitle() As String
    Dim buffer As String
    Dim bufferSize As Long

    bufferSize = 1024
    buffer = String$(bufferSize, Chr$(0))
    bufferSize = GetConsoleTitle(buffer, bufferSize)
    If bufferSize Then
      ConsoleTitle = Left$(buffer, bufferSize)
    End If
  End Property

  Public Property Let ConsoleTitle(ByVal newVal As String)
    SetConsoleTitle newVal
  End Property

  Public Property Get ForeColor() As ConsoleColorConstants
    Const INVALID_HANDLE_VALUE = -1
    Dim csbi As CONSOLE_SCREEN_BUFFER_INFO

    If m_hConsole <> INVALID_HANDLE_VALUE Then
      ' get the current colors, return forecolor
      GetConsoleScreenBufferInfo m_hConsole, csbi
      ForeColor = csbi.wAttributes Mod &H10
    End If
  End Property

  Public Property Let ForeColor(ByVal newVal As ConsoleColorConstants)
    Const INVALID_HANDLE_VALUE = -1
    Dim csbi As CONSOLE_SCREEN_BUFFER_INFO

    If m_hConsole <> INVALID_HANDLE_VALUE Then
      ' attempt to set a new forecolor
      GetConsoleScreenBufferInfo m_hConsole, csbi
      SetConsoleTextAttribute m_hConsole, newVal Or ((csbi.wAttributes \ &H10) * &H10)
    End If
  End Property

  Public Property Get LogFileName() As String
    LogFileName = Me.LogFileDir & "Logfile (" & m_Prefix & ").txt"
  End Property

  Public Property Get LogFileDir() As String
    Dim strLogFileName As String

    strLogFileName = Environ$("UserProfile")
    If Right$(strLogFileName, 1) <> "\" Then strLogFileName = strLogFileName & "\"

    LogFileDir = strLogFileName & "Desktop\"
  End Property

  Public Property Get Prefix() As String
    Prefix = m_Prefix
  End Property

  Public Property Let Prefix(ByVal newVal As String)
    Const CREATE_ALWAYS = 2
    Const FILE_ATTRIBUTE_NORMAL = &H80
    Const FILE_SHARE_READ = &H1
    Const GENERIC_WRITE = &H40000000
    Const INVALID_HANDLE_VALUE = -1
    Dim hFile As Long
    Dim strLogFileName As String

    m_Prefix = newVal

    strLogFileName = Me.LogFileName
    hFile = CreateFileAsLong(strLogFileName, GENERIC_WRITE, FILE_SHARE_READ, 0, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
    If hFile <> INVALID_HANDLE_VALUE Then
      CloseHandle hFile
    End If
  End Property


  Public Sub AddLogEntry(str As String, ByVal logEntryType As LogEntryTypeConstants)
    Const FILE_ATTRIBUTE_NORMAL = &H80
    Const FILE_END = 2
    Const FILE_SHARE_READ = &H1
    Const GENERIC_WRITE = &H40000000
    Const INVALID_HANDLE_VALUE = -1
    Const OPEN_EXISTING = 3
    Dim charsWritten As Long
    Dim hFile As Long
    Dim strConsoleEntry As String
    Dim strLogFileEntry As String
    Dim strLogFileName As String

    #If NoConsole = 0 Then
      If m_hConsole <> INVALID_HANDLE_VALUE Then
        strConsoleEntry = str
        If strConsoleEntry <> "" Then
          If m_Prefix <> "" Then
            strConsoleEntry = "(" & m_Prefix & ") " & strConsoleEntry & vbNewLine
          Else
            strConsoleEntry = strConsoleEntry & vbNewLine
          End If
          Select Case logEntryType
            Case LogEntryTypeConstants.letOther
              Me.ForeColor = ConsoleColorConstants.conWhiteHi
            Case LogEntryTypeConstants.letInfo
              Me.ForeColor = ConsoleColorConstants.conGreenHi
            Case LogEntryTypeConstants.letWarning
              Me.ForeColor = ConsoleColorConstants.conYellowHi
            Case LogEntryTypeConstants.letError
              Me.ForeColor = ConsoleColorConstants.conRedHi
          End Select
          WriteConsole m_hConsole, strConsoleEntry, Len(strConsoleEntry), charsWritten, 0
        End If
      End If
    #End If

    strLogFileName = Me.LogFileName
    hFile = CreateFileAsLong(strLogFileName, GENERIC_WRITE, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If hFile <> INVALID_HANDLE_VALUE Then
      SetFilePointer hFile, 0, 0, FILE_END
      Select Case logEntryType
        Case LogEntryTypeConstants.letOther
          strLogFileEntry = str & vbNewLine
        Case LogEntryTypeConstants.letInfo
          strLogFileEntry = str & vbNewLine
        Case LogEntryTypeConstants.letWarning
          strLogFileEntry = "(WARNING) " & str & vbNewLine
        Case LogEntryTypeConstants.letError
          strLogFileEntry = "(ERROR) " & str & vbNewLine
      End Select
      WriteFileAsLong hFile, ByVal strLogFileEntry, Len(strLogFileEntry), charsWritten, 0

      CloseHandle hFile
    End If
  End Sub


  Private Function FindConsole(ByVal strCaption As String) As Long
    Dim hWnd As Long

    ' search for console window using passed caption and known good console window classnames
    hWnd = FindWindow("ConsoleWindowClass", strCaption)
    If hWnd = 0 Then
      ' try the Win95 console classname
      hWnd = FindWindow("tty", strCaption)
      If hWnd = 0 Then
        ' try no classname at all
        hWnd = FindWindow(vbNullString, strCaption)
      End If
    End If
    FindConsole = hWnd
  End Function

  Private Function IsCompiled() As Boolean
    On Error Resume Next
    Debug.Print 1 / 0
    IsCompiled = (Err.Number = 0)
  End Function
#End If
