#Region
#AutoIt3Wrapper_Icon=Resources\phoenix.ico
#AutoIt3Wrapper_Outfile=Builds\GoogleContactsToMailEnable_v1.0.exe
#AutoIt3Wrapper_Outfile_x64=Builds\GoogleContactsToMailEnable_64-bit(x64)_v1.0.exe
#AutoIt3Wrapper_Res_Comment=By Phoenix125.com
#AutoIt3Wrapper_Res_Description=Google Contacts vCard to Mail Enable Importer
#AutoIt3Wrapper_Res_Fileversion=1.0
#AutoIt3Wrapper_Res_ProductName=Google Contacts To Mail Enable Importer
#AutoIt3Wrapper_Res_ProductVersion=1.0
#AutoIt3Wrapper_Res_CompanyName=http://www.Phoenix125.com
#AutoIt3Wrapper_Res_LegalCopyright=http://www.Phoenix125.com
#AutoIt3Wrapper_Res_SaveSource=y
#AutoIt3Wrapper_Run_AU3Check=n
#AutoIt3Wrapper_AU3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#AutoIt3Wrapper_Tidy_Stop_OnError=n
#AutoIt3Wrapper_Run_Au3Stripper=y
#Au3Stripper_Parameters=/mo
#AutoIt3Wrapper_Res_Icon_Add=Resources\discord.ico
#AutoIt3Wrapper_Res_Icon_Add=Resources\info.ico
#AutoIt3Wrapper_Res_Icon_Add=Resources\forum.ico
#AutoIt3Wrapper_Res_Icon_Add=Resources\manual.ico
#AutoIt3Wrapper_Res_Icon_Add=Resources\refresh.ico
#AutoIt3Wrapper_Res_Icon_Add=Resources\refreshnotice.ico
#EndRegion
Global $aUtilName = "Google Contacts to Mail Enable"
Global $aUtilFileName = StringStripWS($aUtilName, 8)
Global $aFolderTemp = @ScriptDir & "\" & $aUtilFileName & "UtilFiles\"
If @Compiled = 0 Then
Global $aIconFile = @ScriptDir & "\Google Contacts to Mail Enable Icons.exe"
Else
Global $aIconFile = @ScriptFullPath
EndIf
If Not FileExists($aFolderTemp) Then
Do
DirCreate($aFolderTemp)
Until FileExists($aFolderTemp)
EndIf
FileInstall("K:\AutoIT\_MyProgs\GoogleToMailEnable\Resources\phoenixlogo.jpg", $aFolderTemp, 0)
FileInstall("K:\AutoIT\_MyProgs\GoogleToMailEnable\Resources\MailEnableLogo1.jpg", $aFolderTemp, 0)
FileInstall("K:\AutoIT\_MyProgs\GoogleToMailEnable\Resources\GoogleLogo.jpg", $aFolderTemp, 0)
FileInstall("K:\AutoIT\_MyProgs\GoogleToMailEnable\Resources\GoogleExport1.jpg", $aFolderTemp, 0)
Global Const $BS_GROUPBOX = 0x0007
Global Const $BS_BOTTOM = 0x0800
Global Const $BS_CENTER = 0x0300
Global Const $BS_DEFPUSHBUTTON = 0x0001
Global Const $BS_LEFT = 0x0100
Global Const $BS_MULTILINE = 0x2000
Global Const $BS_PUSHBOX = 0x000A
Global Const $BS_PUSHLIKE = 0x1000
Global Const $BS_RIGHT = 0x0200
Global Const $BS_RIGHTBUTTON = 0x0020
Global Const $BS_TOP = 0x0400
Global Const $BS_VCENTER = 0x0C00
Global Const $BS_FLAT = 0x8000
Global Const $BS_ICON = 0x0040
Global Const $BS_BITMAP = 0x0080
Global Const $BS_NOTIFY = 0x4000
Global Const $BS_SPLITBUTTON = 0x0000000C
Global Const $BS_DEFSPLITBUTTON = 0x0000000D
Global Const $BS_COMMANDLINK = 0x0000000E
Global Const $BS_DEFCOMMANDLINK = 0x0000000F
Global Const $BCSIF_GLYPH = 0x0001
Global Const $BCSIF_IMAGE = 0x0002
Global Const $BCSIF_STYLE = 0x0004
Global Const $BCSIF_SIZE = 0x0008
Global Const $BCSS_NOSPLIT = 0x0001
Global Const $BCSS_STRETCH = 0x0002
Global Const $BCSS_ALIGNLEFT = 0x0004
Global Const $BCSS_IMAGE = 0x0008
Global Const $BUTTON_IMAGELIST_ALIGN_LEFT = 0
Global Const $BUTTON_IMAGELIST_ALIGN_RIGHT = 1
Global Const $BUTTON_IMAGELIST_ALIGN_TOP = 2
Global Const $BUTTON_IMAGELIST_ALIGN_BOTTOM = 3
Global Const $BUTTON_IMAGELIST_ALIGN_CENTER = 4
Global Const $BS_3STATE = 0x0005
Global Const $BS_AUTO3STATE = 0x0006
Global Const $BS_AUTOCHECKBOX = 0x0003
Global Const $BS_CHECKBOX = 0x0002
Global Const $BS_RADIOBUTTON = 0x4
Global Const $BS_AUTORADIOBUTTON = 0x0009
Global Const $BS_OWNERDRAW = 0xB
Global Const $GUI_SS_DEFAULT_BUTTON = 0
Global Const $GUI_SS_DEFAULT_CHECKBOX = 0
Global Const $GUI_SS_DEFAULT_GROUP = 0
Global Const $GUI_SS_DEFAULT_RADIO = 0
Global Const $BCM_FIRST = 0x1600
Global Const $BCM_GETIDEALSIZE = ($BCM_FIRST + 0x0001)
Global Const $BCM_GETIMAGELIST = ($BCM_FIRST + 0x0003)
Global Const $BCM_GETNOTE = ($BCM_FIRST + 0x000A)
Global Const $BCM_GETNOTELENGTH = ($BCM_FIRST + 0x000B)
Global Const $BCM_GETSPLITINFO = ($BCM_FIRST + 0x0008)
Global Const $BCM_GETTEXTMARGIN = ($BCM_FIRST + 0x0005)
Global Const $BCM_SETDROPDOWNSTATE = ($BCM_FIRST + 0x0006)
Global Const $BCM_SETIMAGELIST = ($BCM_FIRST + 0x0002)
Global Const $BCM_SETNOTE = ($BCM_FIRST + 0x0009)
Global Const $BCM_SETSHIELD = ($BCM_FIRST + 0x000C)
Global Const $BCM_SETSPLITINFO = ($BCM_FIRST + 0x0007)
Global Const $BCM_SETTEXTMARGIN = ($BCM_FIRST + 0x0004)
Global Const $BM_CLICK = 0xF5
Global Const $BM_GETCHECK = 0xF0
Global Const $BM_GETIMAGE = 0xF6
Global Const $BM_GETSTATE = 0xF2
Global Const $BM_SETCHECK = 0xF1
Global Const $BM_SETDONTCLICK = 0xF8
Global Const $BM_SETIMAGE = 0xF7
Global Const $BM_SETSTATE = 0xF3
Global Const $BM_SETSTYLE = 0xF4
Global Const $BCN_FIRST = -1250
Global Const $BCN_DROPDOWN = ($BCN_FIRST + 0x0002)
Global Const $BCN_HOTITEMCHANGE = ($BCN_FIRST + 0x0001)
Global Const $BN_CLICKED = 0
Global Const $BN_PAINT = 1
Global Const $BN_HILITE = 2
Global Const $BN_UNHILITE = 3
Global Const $BN_DISABLE = 4
Global Const $BN_DOUBLECLICKED = 5
Global Const $BN_SETFOCUS = 6
Global Const $BN_KILLFOCUS = 7
Global Const $BN_PUSHED = $BN_HILITE
Global Const $BN_UNPUSHED = $BN_UNHILITE
Global Const $BN_DBLCLK = $BN_DOUBLECLICKED
Global Const $BST_CHECKED = 0x1
Global Const $BST_INDETERMINATE = 0x2
Global Const $BST_UNCHECKED = 0x0
Global Const $BST_FOCUS = 0x8
Global Const $BST_PUSHED = 0x4
Global Const $BST_DONTCLICK = 0x000080
Global Const $ES_LEFT = 0
Global Const $ES_CENTER = 1
Global Const $ES_RIGHT = 2
Global Const $ES_MULTILINE = 4
Global Const $ES_UPPERCASE = 8
Global Const $ES_LOWERCASE = 16
Global Const $ES_PASSWORD = 32
Global Const $ES_AUTOVSCROLL = 64
Global Const $ES_AUTOHSCROLL = 128
Global Const $ES_NOHIDESEL = 256
Global Const $ES_OEMCONVERT = 1024
Global Const $ES_READONLY = 2048
Global Const $ES_WANTRETURN = 4096
Global Const $ES_NUMBER = 8192
Global Const $EC_ERR = -1
Global Const $ECM_FIRST = 0X1500
Global Const $EM_CANUNDO = 0xC6
Global Const $EM_CHARFROMPOS = 0xD7
Global Const $EM_EMPTYUNDOBUFFER = 0xCD
Global Const $EM_FMTLINES = 0xC8
Global Const $EM_GETCUEBANNER = ($ECM_FIRST + 2)
Global Const $EM_GETFIRSTVISIBLELINE = 0xCE
Global Const $EM_GETHANDLE = 0xBD
Global Const $EM_GETIMESTATUS = 0xD9
Global Const $EM_GETLIMITTEXT = 0xD5
Global Const $EM_GETLINE = 0xC4
Global Const $EM_GETLINECOUNT = 0xBA
Global Const $EM_GETMARGINS = 0xD4
Global Const $EM_GETMODIFY = 0xB8
Global Const $EM_GETPASSWORDCHAR = 0xD2
Global Const $EM_GETRECT = 0xB2
Global Const $EM_GETSEL = 0xB0
Global Const $EM_GETTHUMB = 0xBE
Global Const $EM_GETWORDBREAKPROC = 0xD1
Global Const $EM_HIDEBALLOONTIP = ($ECM_FIRST + 4)
Global Const $EM_LIMITTEXT = 0xC5
Global Const $EM_LINEFROMCHAR = 0xC9
Global Const $EM_LINEINDEX = 0xBB
Global Const $EM_LINELENGTH = 0xC1
Global Const $EM_LINESCROLL = 0xB6
Global Const $EM_POSFROMCHAR = 0xD6
Global Const $EM_REPLACESEL = 0xC2
Global Const $EM_SCROLL = 0xB5
Global Const $EM_SCROLLCARET = 0x00B7
Global Const $EM_SETCUEBANNER = ($ECM_FIRST + 1)
Global Const $EM_SETHANDLE = 0xBC
Global Const $EM_SETIMESTATUS = 0xD8
Global Const $EM_SETLIMITTEXT = $EM_LIMITTEXT
Global Const $EM_SETMARGINS = 0xD3
Global Const $EM_SETMODIFY = 0xB9
Global Const $EM_SETPASSWORDCHAR = 0xCC
Global Const $EM_SETREADONLY = 0xCF
Global Const $EM_SETRECT = 0xB3
Global Const $EM_SETRECTNP = 0xB4
Global Const $EM_SETSEL = 0xB1
Global Const $EM_SETTABSTOPS = 0xCB
Global Const $EM_SETWORDBREAKPROC = 0xD0
Global Const $EM_SHOWBALLOONTIP = ($ECM_FIRST + 3)
Global Const $EM_UNDO = 0xC7
Global Const $EC_LEFTMARGIN = 0x1
Global Const $EC_RIGHTMARGIN = 0x2
Global Const $EC_USEFONTINFO = 0xFFFF
Global Const $EMSIS_COMPOSITIONSTRING = 0x1
Global Const $EIMES_GETCOMPSTRATONCE = 0x1
Global Const $EIMES_CANCELCOMPSTRINFOCUS = 0x2
Global Const $EIMES_COMPLETECOMPSTRKILLFOCUS = 0x4
Global Const $EN_ALIGN_LTR_EC = 0x700
Global Const $EN_ALIGN_RTL_EC = 0x701
Global Const $EN_CHANGE = 0x300
Global Const $EN_ERRSPACE = 0x500
Global Const $EN_HSCROLL = 0X601
Global Const $EN_KILLFOCUS = 0x200
Global Const $EN_MAXTEXT = 0x501
Global Const $EN_SETFOCUS = 0x100
Global Const $EN_UPDATE = 0x400
Global Const $EN_VSCROLL = 0x602
Global Const $GUI_SS_DEFAULT_EDIT = 0x003010c0
Global Const $GUI_SS_DEFAULT_INPUT = 0x00000080
Global Const $OPT_COORDSRELATIVE = 0
Global Const $OPT_COORDSABSOLUTE = 1
Global Const $OPT_COORDSCLIENT = 2
Global Const $OPT_ERRORSILENT = 0
Global Const $OPT_ERRORFATAL = 1
Global Const $OPT_CAPSNOSTORE = 0
Global Const $OPT_CAPSSTORE = 1
Global Const $OPT_MATCHSTART = 1
Global Const $OPT_MATCHANY = 2
Global Const $OPT_MATCHEXACT = 3
Global Const $OPT_MATCHADVANCED = 4
Global Const $CCS_TOP = 0x01
Global Const $CCS_NOMOVEY = 0x02
Global Const $CCS_BOTTOM = 0x03
Global Const $CCS_NORESIZE = 0x04
Global Const $CCS_NOPARENTALIGN = 0x08
Global Const $CCS_NOHILITE = 0x10
Global Const $CCS_ADJUSTABLE = 0x20
Global Const $CCS_NODIVIDER = 0x40
Global Const $CCS_VERT = 0x0080
Global Const $CCS_LEFT = 0x0081
Global Const $CCS_NOMOVEX = 0x0082
Global Const $CCS_RIGHT = 0x0083
Global Const $DT_DRIVETYPE = 1
Global Const $DT_SSDSTATUS = 2
Global Const $DT_BUSTYPE = 3
Global Const $PROXY_IE = 0
Global Const $PROXY_NONE = 1
Global Const $PROXY_SPECIFIED = 2
Global Const $OBJID_WINDOW = 0x00000000
Global Const $OBJID_TITLEBAR = 0xFFFFFFFE
Global Const $OBJID_SIZEGRIP = 0xFFFFFFF9
Global Const $OBJID_CARET = 0xFFFFFFF8
Global Const $OBJID_CURSOR = 0xFFFFFFF7
Global Const $OBJID_ALERT = 0xFFFFFFF6
Global Const $OBJID_SOUND = 0xFFFFFFF5
Global Const $DLG_CENTERONTOP = 0
Global Const $DLG_NOTITLE = 1
Global Const $DLG_NOTONTOP = 2
Global Const $DLG_TEXTLEFT = 4
Global Const $DLG_TEXTRIGHT = 8
Global Const $DLG_MOVEABLE = 16
Global Const $DLG_TEXTVCENTER = 32
Global Const $IDC_UNKNOWN = 0
Global Const $IDC_APPSTARTING = 1
Global Const $IDC_ARROW = 2
Global Const $IDC_CROSS = 3
Global Const $IDC_HAND = 32649
Global Const $IDC_HELP = 4
Global Const $IDC_IBEAM = 5
Global Const $IDC_ICON = 6
Global Const $IDC_NO = 7
Global Const $IDC_SIZE = 8
Global Const $IDC_SIZEALL = 9
Global Const $IDC_SIZENESW = 10
Global Const $IDC_SIZENS = 11
Global Const $IDC_SIZENWSE = 12
Global Const $IDC_SIZEWE = 13
Global Const $IDC_UPARROW = 14
Global Const $IDC_WAIT = 15
Global Const $IDI_APPLICATION = 32512
Global Const $IDI_ASTERISK = 32516
Global Const $IDI_EXCLAMATION = 32515
Global Const $IDI_HAND = 32513
Global Const $IDI_QUESTION = 32514
Global Const $IDI_WINLOGO = 32517
Global Const $IDI_SHIELD = 32518
Global Const $IDI_ERROR = $IDI_HAND
Global Const $IDI_INFORMATION = $IDI_ASTERISK
Global Const $IDI_WARNING = $IDI_EXCLAMATION
Global Const $SD_LOGOFF = 0
Global Const $SD_SHUTDOWN = 1
Global Const $SD_REBOOT = 2
Global Const $SD_FORCE = 4
Global Const $SD_POWERDOWN = 8
Global Const $SD_FORCEHUNG = 16
Global Const $SD_STANDBY = 32
Global Const $SD_HIBERNATE = 64
Global Const $STDIN_CHILD = 1
Global Const $STDOUT_CHILD = 2
Global Const $STDERR_CHILD = 4
Global Const $STDERR_MERGED = 8
Global Const $STDIO_INHERIT_PARENT = 0x10
Global Const $RUN_CREATE_NEW_CONSOLE = 0x00010000
Global Const $UBOUND_DIMENSIONS = 0
Global Const $UBOUND_ROWS = 1
Global Const $UBOUND_COLUMNS = 2
Global Const $MOUSEEVENTF_ABSOLUTE = 0x8000
Global Const $MOUSEEVENTF_MOVE = 0x0001
Global Const $MOUSEEVENTF_LEFTDOWN = 0x0002
Global Const $MOUSEEVENTF_LEFTUP = 0x0004
Global Const $MOUSEEVENTF_RIGHTDOWN = 0x0008
Global Const $MOUSEEVENTF_RIGHTUP = 0x0010
Global Const $MOUSEEVENTF_MIDDLEDOWN = 0x0020
Global Const $MOUSEEVENTF_MIDDLEUP = 0x0040
Global Const $MOUSEEVENTF_WHEEL = 0x0800
Global Const $MOUSEEVENTF_XDOWN = 0x0080
Global Const $MOUSEEVENTF_XUP = 0x0100
Global Const $REG_NONE = 0
Global Const $REG_SZ = 1
Global Const $REG_EXPAND_SZ = 2
Global Const $REG_BINARY = 3
Global Const $REG_DWORD = 4
Global Const $REG_DWORD_LITTLE_ENDIAN = 4
Global Const $REG_DWORD_BIG_ENDIAN = 5
Global Const $REG_LINK = 6
Global Const $REG_MULTI_SZ = 7
Global Const $REG_RESOURCE_LIST = 8
Global Const $REG_FULL_RESOURCE_DESCRIPTOR = 9
Global Const $REG_RESOURCE_REQUIREMENTS_LIST = 10
Global Const $REG_QWORD = 11
Global Const $REG_QWORD_LITTLE_ENDIAN = 11
Global Const $HWND_BOTTOM = 1
Global Const $HWND_NOTOPMOST = -2
Global Const $HWND_TOP = 0
Global Const $HWND_TOPMOST = -1
Global Const $SWP_NOSIZE = 0x0001
Global Const $SWP_NOMOVE = 0x0002
Global Const $SWP_NOZORDER = 0x0004
Global Const $SWP_NOREDRAW = 0x0008
Global Const $SWP_NOACTIVATE = 0x0010
Global Const $SWP_FRAMECHANGED = 0x0020
Global Const $SWP_DRAWFRAME = 0x0020
Global Const $SWP_SHOWWINDOW = 0x0040
Global Const $SWP_HIDEWINDOW = 0x0080
Global Const $SWP_NOCOPYBITS = 0x0100
Global Const $SWP_NOOWNERZORDER = 0x0200
Global Const $SWP_NOREPOSITION = 0x0200
Global Const $SWP_NOSENDCHANGING = 0x0400
Global Const $SWP_DEFERERASE = 0x2000
Global Const $SWP_ASYNCWINDOWPOS = 0x4000
Global Const $KEYWORD_DEFAULT = 1
Global Const $KEYWORD_NULL = 2
Global Const $DECLARED_LOCAL = -1
Global Const $DECLARED_UNKNOWN = 0
Global Const $DECLARED_GLOBAL = 1
Global Const $ASSIGN_CREATE = 0
Global Const $ASSIGN_FORCELOCAL = 1
Global Const $ASSIGN_FORCEGLOBAL = 2
Global Const $ASSIGN_EXISTFAIL = 4
Global Const $BI_ENABLE = 0
Global Const $BI_DISABLE = 1
Global Const $BREAK_ENABLE = 1
Global Const $BREAK_DISABLE = 0
Global Const $CDTRAY_OPEN = "open"
Global Const $CDTRAY_CLOSED = "closed"
Global Const $SEND_DEFAULT = 0
Global Const $SEND_RAW = 1
Global Const $DIR_DEFAULT = 0
Global Const $DIR_EXTENDED= 1
Global Const $DIR_NORECURSE = 2
Global Const $DIR_REMOVE= 1
Global Const $DT_ALL = "ALL"
Global Const $DT_CDROM = "CDROM"
Global Const $DT_REMOVABLE = "REMOVABLE"
Global Const $DT_FIXED = "FIXED"
Global Const $DT_NETWORK = "NETWORK"
Global Const $DT_RAMDISK = "RAMDISK"
Global Const $DT_UNKNOWN = "UNKNOWN"
Global Const $DT_UNDEFINED = 1
Global Const $DT_FAT = "FAT"
Global Const $DT_FAT32 = "FAT32"
Global Const $DT_EXFAT = "exFAT"
Global Const $DT_NTFS = "NTFS"
Global Const $DT_NWFS = "NWFS"
Global Const $DT_CDFS = "CDFS"
Global Const $DT_UDF = "UDF"
Global Const $DMA_DEFAULT = 0
Global Const $DMA_PERSISTENT = 1
Global Const $DMA_AUTHENTICATION = 8
Global Const $DS_UNKNOWN = "UNKNOWN"
Global Const $DS_READY = "READY"
Global Const $DS_NOTREADY = "NOTREADY"
Global Const $DS_INVALID = "INVALID"
Global Const $MOUSE_CLICK_LEFT = "left"
Global Const $MOUSE_CLICK_RIGHT = "right"
Global Const $MOUSE_CLICK_MIDDLE = "middle"
Global Const $MOUSE_CLICK_MAIN = "main"
Global Const $MOUSE_CLICK_MENU = "menu"
Global Const $MOUSE_CLICK_PRIMARY = "primary"
Global Const $MOUSE_CLICK_SECONDARY = "secondary"
Global Const $MOUSE_WHEEL_UP = "up"
Global Const $MOUSE_WHEEL_DOWN = "down"
Global Const $NUMBER_AUTO = 0
Global Const $NUMBER_32BIT = 1
Global Const $NUMBER_64BIT = 2
Global Const $NUMBER_DOUBLE = 3
Global Const $OBJ_NAME = 1
Global Const $OBJ_STRING = 2
Global Const $OBJ_PROGID = 3
Global Const $OBJ_FILE = 4
Global Const $OBJ_MODULE = 5
Global Const $OBJ_CLSID = 6
Global Const $OBJ_IID = 7
Global Const $EXITCLOSE_NORMAL = 0
Global Const $EXITCLOSE_BYEXIT = 1
Global Const $EXITCLOSE_BYCLICK = 2
Global Const $EXITCLOSE_BYLOGOFF = 3
Global Const $EXITCLOSE_BYSHUTDOWN = 4
Global Const $PROCESS_STATS_MEMORY = 0
Global Const $PROCESS_STATS_IO = 1
Global Const $PROCESS_LOW = 0
Global Const $PROCESS_BELOWNORMAL = 1
Global Const $PROCESS_NORMAL = 2
Global Const $PROCESS_ABOVENORMAL = 3
Global Const $PROCESS_HIGH = 4
Global Const $PROCESS_REALTIME = 5
Global Const $RUN_LOGON_NOPROFILE = 0
Global Const $RUN_LOGON_PROFILE = 1
Global Const $RUN_LOGON_NETWORK = 2
Global Const $RUN_LOGON_INHERIT = 4
Global Const $SOUND_NOWAIT = 0
Global Const $SOUND_WAIT = 1
Global Const $SHEX_OPEN = "open"
Global Const $SHEX_EDIT = "edit"
Global Const $SHEX_PRINT = "print"
Global Const $SHEX_PROPERTIES = "properties"
Global Const $TCP_DATA_DEFAULT = 0
Global Const $TCP_DATA_BINARY = 1
Global Const $UDP_OPEN_DEFAULT = 0
Global Const $UDP_OPEN_BROADCAST = 1
Global Const $UDP_DATA_DEFAULT = 0
Global Const $UDP_DATA_BINARY = 1
Global Const $UDP_DATA_ARRAY = 2
Global Const $TIP_NOICON = 0
Global Const $TIP_INFOICON = 1
Global Const $TIP_WARNINGICON = 2
Global Const $TIP_ERRORICON = 3
Global Const $TIP_BALLOON = 1
Global Const $TIP_CENTER = 2
Global Const $TIP_FORCEVISIBLE = 4
Global Const $WINDOWS_NOONTOP = 0
Global Const $WINDOWS_ONTOP = 1
Global Const $WIN_STATE_EXISTS = 1
Global Const $WIN_STATE_VISIBLE  = 2
Global Const $WIN_STATE_ENABLED = 4
Global Const $WIN_STATE_ACTIVE = 8
Global Const $WIN_STATE_MINIMIZED = 16
Global Const $WIN_STATE_MAXIMIZED = 32
Global Const $MB_OK = 0
Global Const $MB_OKCANCEL = 1
Global Const $MB_ABORTRETRYIGNORE = 2
Global Const $MB_YESNOCANCEL = 3
Global Const $MB_YESNO = 4
Global Const $MB_RETRYCANCEL = 5
Global Const $MB_CANCELTRYCONTINUE = 6
Global Const $MB_HELP = 0x4000
Global Const $MB_ICONSTOP = 16
Global Const $MB_ICONERROR = 16
Global Const $MB_ICONHAND = 16
Global Const $MB_ICONQUESTION = 32
Global Const $MB_ICONEXCLAMATION = 48
Global Const $MB_ICONWARNING = 48
Global Const $MB_ICONINFORMATION = 64
Global Const $MB_ICONASTERISK = 64
Global Const $MB_USERICON = 0x00000080
Global Const $MB_DEFBUTTON1 = 0
Global Const $MB_DEFBUTTON2 = 256
Global Const $MB_DEFBUTTON3 = 512
Global Const $MB_DEFBUTTON4 = 768
Global Const $MB_APPLMODAL = 0
Global Const $MB_SYSTEMMODAL = 4096
Global Const $MB_TASKMODAL = 8192
Global Const $MB_DEFAULT_DESKTOP_ONLY = 0x00020000
Global Const $MB_RIGHT = 0x00080000
Global Const $MB_RTLREADING = 0x00100000
Global Const $MB_SETFOREGROUND = 0x00010000
Global Const $MB_TOPMOST = 0x00040000
Global Const $MB_SERVICE_NOTIFICATION = 0x00200000
Global Const $MB_RIGHTJUSTIFIED = $MB_RIGHT
Global Const $IDTIMEOUT = -1
Global Const $IDOK = 1
Global Const $IDCANCEL = 2
Global Const $IDABORT = 3
Global Const $IDRETRY = 4
Global Const $IDIGNORE = 5
Global Const $IDYES = 6
Global Const $IDNO = 7
Global Const $IDCLOSE = 8
Global Const $IDHELP = 9
Global Const $IDTRYAGAIN = 10
Global Const $IDCONTINUE = 11
Global Const $STR_NOCASESENSE = 0
Global Const $STR_CASESENSE = 1
Global Const $STR_NOCASESENSEBASIC = 2
Global Const $STR_STRIPLEADING = 1
Global Const $STR_STRIPTRAILING = 2
Global Const $STR_STRIPSPACES = 4
Global Const $STR_STRIPALL = 8
Global Const $STR_CHRSPLIT = 0
Global Const $STR_ENTIRESPLIT = 1
Global Const $STR_NOCOUNT = 2
Global Const $STR_REGEXPMATCH = 0
Global Const $STR_REGEXPARRAYMATCH = 1
Global Const $STR_REGEXPARRAYFULLMATCH = 2
Global Const $STR_REGEXPARRAYGLOBALMATCH = 3
Global Const $STR_REGEXPARRAYGLOBALFULLMATCH = 4
Global Const $STR_ENDISSTART = 0
Global Const $STR_ENDNOTSTART = 1
Global Const $SB_ANSI = 1
Global Const $SB_UTF16LE = 2
Global Const $SB_UTF16BE = 3
Global Const $SB_UTF8 = 4
Global Const $SE_UTF16 = 0
Global Const $SE_ANSI = 1
Global Const $SE_UTF8 = 2
Global Const $STR_UTF16 = 0
Global Const $STR_UCS2 = 1
#Region Global Variables and Constants
Global Const $_ARRAYCONSTANT_SORTINFOSIZE = 11
Global $__g_aArrayDisplay_SortInfo[$_ARRAYCONSTANT_SORTINFOSIZE]
Global Const $ARRAYDISPLAY_COLALIGNLEFT = 0
Global Const $ARRAYDISPLAY_TRANSPOSE = 1
Global Const $ARRAYDISPLAY_COLALIGNRIGHT = 2
Global Const $ARRAYDISPLAY_COLALIGNCENTER = 4
Global Const $ARRAYDISPLAY_VERBOSE = 8
Global Const $ARRAYDISPLAY_NOROW = 64
Global Const $_ARRAYCONSTANT_tagHDITEM = "uint Mask;int XY;ptr Text;handle hBMP;int TextMax;int Fmt;lparam Param;int Image;int Order;uint Type;ptr pFilter;uint State"
Global Const $_ARRAYCONSTANT_tagLVITEM = "struct;uint Mask;int Item;int SubItem;uint State;uint StateMask;ptr Text;int TextMax;int Image;lparam Param;" &  "int Indent;int GroupID;uint Columns;ptr pColumns;ptr piColFmt;int iGroup;endstruct"
#EndRegion Global Variables and Constants
#Region Functions list
#EndRegion Functions list
Func __ArrayDisplay_Share(Const ByRef $aArray, $sTitle = Default, $sArrayRange = Default, $iFlags = Default, $vUser_Separator = Default, $sHeader = Default, $iMax_ColWidth = Default, $hUser_Function = Default, $bDebug = True)
Local $vTmp, $sMsgBoxTitle = (($bDebug) ? ("DebugArray") : ("ArrayDisplay"))
If $sTitle = Default Then $sTitle = $sMsgBoxTitle
If $sArrayRange = Default Then $sArrayRange = ""
If $iFlags = Default Then $iFlags = 0
If $vUser_Separator = Default Then $vUser_Separator = ""
If $sHeader = Default Then $sHeader = ""
If $iMax_ColWidth = Default Then $iMax_ColWidth = 350
If $hUser_Function = Default Then $hUser_Function = 0
Local $iTranspose = BitAND($iFlags, $ARRAYDISPLAY_TRANSPOSE)
Local $iColAlign = BitAND($iFlags, 6)
Local $iVerbose = BitAND($iFlags, $ARRAYDISPLAY_VERBOSE)
Local $iNoRow = BitAND($iFlags, $ARRAYDISPLAY_NOROW)
Local $iButtonBorder = (($bDebug) ? (40) : (20))
Local $sMsg = "", $iRet = 1
If IsArray($aArray) Then
Local $iDimension = UBound($aArray, $UBOUND_DIMENSIONS), $iRowCount = UBound($aArray, $UBOUND_ROWS), $iColCount = UBound($aArray, $UBOUND_COLUMNS)
If $iDimension > 2 Then
$sMsg = "Larger than 2D array passed to function"
$iRet = 2
EndIf
If $iDimension = 1 Then
$iTranspose = 0
EndIf
Else
$sMsg = "No array variable passed to function"
EndIf
If $sMsg Then
If $iVerbose And MsgBox($MB_SYSTEMMODAL + $MB_ICONERROR + $MB_YESNO,  $sMsgBoxTitle & " Error: " & $sTitle, $sMsg & @CRLF & @CRLF & "Exit the script?") = $IDYES Then
Exit
Else
Return SetError($iRet, 0, 0)
EndIf
EndIf
Local $iCW_ColWidth = Number($vUser_Separator)
Local $sCurr_Separator = Opt("GUIDataSeparatorChar")
If $vUser_Separator = "" Then $vUser_Separator = $sCurr_Separator
Local $iItem_Start = 0, $iItem_End = $iRowCount - 1, $iSubItem_Start = 0, $iSubItem_End = (($iDimension = 2) ? ($iColCount - 1) : (0))
Local $bRange_Flag = False, $avRangeSplit
If $sArrayRange Then
Local $aArray_Range = StringRegExp($sArrayRange & "||", "(?U)(.*)\|", 3)
If $aArray_Range[0] Then
$avRangeSplit = StringSplit($aArray_Range[0], ":")
If @error Then
$iItem_End = Number($avRangeSplit[1])
Else
$iItem_Start = Number($avRangeSplit[1])
If $avRangeSplit[2] <> "" Then
$iItem_End = Number($avRangeSplit[2])
EndIf
EndIf
EndIf
If $iItem_Start < 0 Then $iItem_Start = 0
If $iItem_End > $iRowCount - 1 Then $iItem_End = $iRowCount - 1
If $iItem_Start > $iItem_End Then
$vTmp = $iItem_Start
$iItem_Start = $iItem_End
$iItem_End = $vTmp
EndIf
If $iItem_Start <> 0 Or $iItem_End <> $iRowCount - 1 Then $bRange_Flag = True
If $iDimension = 2 And $aArray_Range[1] Then
$avRangeSplit = StringSplit($aArray_Range[1], ":")
If @error Then
$iSubItem_End = Number($avRangeSplit[1])
Else
$iSubItem_Start = Number($avRangeSplit[1])
If $avRangeSplit[2] <> "" Then
$iSubItem_End = Number($avRangeSplit[2])
EndIf
EndIf
If $iSubItem_Start > $iSubItem_End Then
$vTmp = $iSubItem_Start
$iSubItem_Start = $iSubItem_End
$iSubItem_End = $vTmp
EndIf
If $iSubItem_Start < 0 Then $iSubItem_Start = 0
If $iSubItem_End > $iColCount - 1 Then $iSubItem_End = $iColCount - 1
If $iSubItem_Start <> 0 Or $iSubItem_End <> $iColCount - 1 Then $bRange_Flag = True
EndIf
EndIf
Local $sDisplayData = "[" & $iRowCount & "]"
If $iDimension = 2 Then
$sDisplayData &= " [" & $iColCount & "]"
EndIf
Local $sTipData = ""
If $bRange_Flag Then
If $sTipData Then $sTipData &= " - "
$sTipData &= "Range set"
EndIf
If $iTranspose Then
If $sTipData Then $sTipData &= " - "
$sTipData &= "Transposed"
EndIf
Local $asHeader = StringSplit($sHeader, $sCurr_Separator, $STR_NOCOUNT)
If UBound($asHeader) = 0 Then Local $asHeader[1] = [""]
$sHeader = "Row"
Local $iIndex = $iSubItem_Start
If $iTranspose Then
$sHeader = "Col"
For $j = $iItem_Start To $iItem_End
$sHeader &= $sCurr_Separator & "Row " & $j
Next
Else
If $asHeader[0] Then
For $iIndex = $iSubItem_Start To $iSubItem_End
If $iIndex >= UBound($asHeader) Then ExitLoop
$sHeader &= $sCurr_Separator & $asHeader[$iIndex]
Next
EndIf
For $j = $iIndex To $iSubItem_End
$sHeader &= $sCurr_Separator & "Col " & $j
Next
EndIf
If $iNoRow Then $sHeader = StringTrimLeft($sHeader, 4)
If $iVerbose And ($iItem_End - $iItem_Start + 1) * ($iSubItem_End - $iSubItem_Start + 1) > 10000 Then
SplashTextOn($sMsgBoxTitle, "Preparing display" & @CRLF & @CRLF & "Please be patient", 300, 100)
EndIf
Local Const $_ARRAYCONSTANT_GUI_DOCKBOTTOM = 64
Local Const $_ARRAYCONSTANT_GUI_DOCKBORDERS = 102
Local Const $_ARRAYCONSTANT_GUI_DOCKHEIGHT = 512
Local Const $_ARRAYCONSTANT_GUI_DOCKLEFT = 2
Local Const $_ARRAYCONSTANT_GUI_DOCKRIGHT = 4
Local Const $_ARRAYCONSTANT_GUI_DOCKHCENTER = 8
Local Const $_ARRAYCONSTANT_GUI_EVENT_CLOSE = -3
Local Const $_ARRAYCONSTANT_GUI_FOCUS = 256
Local Const $_ARRAYCONSTANT_SS_CENTER = 0x1
Local Const $_ARRAYCONSTANT_SS_CENTERIMAGE = 0x0200
Local Const $_ARRAYCONSTANT_LVM_GETITEMCOUNT = (0x1000 + 4)
Local Const $_ARRAYCONSTANT_LVM_GETITEMRECT = (0x1000 + 14)
Local Const $_ARRAYCONSTANT_LVM_GETCOLUMNWIDTH = (0x1000 + 29)
Local Const $_ARRAYCONSTANT_LVM_SETCOLUMNWIDTH = (0x1000 + 30)
Local Const $_ARRAYCONSTANT_LVM_GETITEMSTATE = (0x1000 + 44)
Local Const $_ARRAYCONSTANT_LVM_GETSELECTEDCOUNT = (0x1000 + 50)
Local Const $_ARRAYCONSTANT_LVM_SETEXTENDEDLISTVIEWSTYLE = (0x1000 + 54)
Local Const $_ARRAYCONSTANT_LVS_EX_GRIDLINES = 0x1
Local Const $_ARRAYCONSTANT_LVIS_SELECTED = 0x0002
Local Const $_ARRAYCONSTANT_LVS_SHOWSELALWAYS = 0x8
Local Const $_ARRAYCONSTANT_LVS_EX_FULLROWSELECT = 0x20
Local Const $_ARRAYCONSTANT_WS_EX_CLIENTEDGE = 0x0200
Local Const $_ARRAYCONSTANT_WS_MAXIMIZEBOX = 0x00010000
Local Const $_ARRAYCONSTANT_WS_MINIMIZEBOX = 0x00020000
Local Const $_ARRAYCONSTANT_WS_SIZEBOX = 0x00040000
Local Const $_ARRAYCONSTANT_WM_SETREDRAW = 11
Local Const $_ARRAYCONSTANT_LVSCW_AUTOSIZE = -1
Local Const $_ARRAYCONSTANT_LVSCW_AUTOSIZE_USEHEADER = -2
Local $iCoordMode = Opt("GUICoordMode", 1)
Local $iOrgWidth = 210, $iHeight = 200, $iMinSize = 250
Local $hGUI = GUICreate($sTitle, $iOrgWidth, $iHeight, Default, Default, BitOR($_ARRAYCONSTANT_WS_SIZEBOX, $_ARRAYCONSTANT_WS_MINIMIZEBOX, $_ARRAYCONSTANT_WS_MAXIMIZEBOX))
Local $aiGUISize = WinGetClientSize($hGUI)
Local $iButtonWidth_1 = $aiGUISize[0] / 2
Local $iButtonWidth_2 = $aiGUISize[0] / 3
Local $idListView = GUICtrlCreateListView($sHeader, 0, 0, $aiGUISize[0], $aiGUISize[1] - $iButtonBorder, $_ARRAYCONSTANT_LVS_SHOWSELALWAYS)
GUICtrlSendMsg($idListView, $_ARRAYCONSTANT_LVM_SETEXTENDEDLISTVIEWSTYLE, $_ARRAYCONSTANT_LVS_EX_GRIDLINES, $_ARRAYCONSTANT_LVS_EX_GRIDLINES)
GUICtrlSendMsg($idListView, $_ARRAYCONSTANT_LVM_SETEXTENDEDLISTVIEWSTYLE, $_ARRAYCONSTANT_LVS_EX_FULLROWSELECT, $_ARRAYCONSTANT_LVS_EX_FULLROWSELECT)
GUICtrlSendMsg($idListView, $_ARRAYCONSTANT_LVM_SETEXTENDEDLISTVIEWSTYLE, $_ARRAYCONSTANT_WS_EX_CLIENTEDGE, $_ARRAYCONSTANT_WS_EX_CLIENTEDGE)
Local $idCopy_ID = 9999, $idCopy_Data = 99999, $idData_Label = 99999, $idUser_Func = 99999, $idExit_Script = 99999
If $bDebug Then
$idCopy_ID = GUICtrlCreateButton("Copy Data && Hdr/Row", 0, $aiGUISize[1] - $iButtonBorder, $iButtonWidth_1, 20)
$idCopy_Data = GUICtrlCreateButton("Copy Data Only", $iButtonWidth_1, $aiGUISize[1] - $iButtonBorder, $iButtonWidth_1, 20)
Local $iButtonWidth_Var = $iButtonWidth_1
Local $iOffset = $iButtonWidth_1
If IsFunc($hUser_Function) Then
$idUser_Func = GUICtrlCreateButton("Run User Func", $iButtonWidth_2, $aiGUISize[1] - 20, $iButtonWidth_2, 20)
$iButtonWidth_Var = $iButtonWidth_2
$iOffset = $iButtonWidth_2 * 2
EndIf
$idExit_Script = GUICtrlCreateButton("Exit Script", $iOffset, $aiGUISize[1] - 20, $iButtonWidth_Var, 20)
$idData_Label = GUICtrlCreateLabel($sDisplayData, 0, $aiGUISize[1] - 20, $iButtonWidth_Var, 18, BitOR($_ARRAYCONSTANT_SS_CENTER, $_ARRAYCONSTANT_SS_CENTERIMAGE))
Else
$idData_Label = GUICtrlCreateLabel($sDisplayData, 0, $aiGUISize[1] - 20, $aiGUISize[0], 18, BitOR($_ARRAYCONSTANT_SS_CENTER, $_ARRAYCONSTANT_SS_CENTERIMAGE))
EndIf
Select
Case $iTranspose Or $bRange_Flag
GUICtrlSetColor($idData_Label, 0xFF0000)
GUICtrlSetTip($idData_Label, $sTipData)
EndSelect
GUICtrlSetResizing($idListView, $_ARRAYCONSTANT_GUI_DOCKBORDERS)
GUICtrlSetResizing($idCopy_ID, $_ARRAYCONSTANT_GUI_DOCKLEFT + $_ARRAYCONSTANT_GUI_DOCKBOTTOM + $_ARRAYCONSTANT_GUI_DOCKHEIGHT)
GUICtrlSetResizing($idCopy_Data, $_ARRAYCONSTANT_GUI_DOCKRIGHT + $_ARRAYCONSTANT_GUI_DOCKBOTTOM + $_ARRAYCONSTANT_GUI_DOCKHEIGHT)
GUICtrlSetResizing($idData_Label, $_ARRAYCONSTANT_GUI_DOCKLEFT + $_ARRAYCONSTANT_GUI_DOCKBOTTOM + $_ARRAYCONSTANT_GUI_DOCKHEIGHT)
GUICtrlSetResizing($idUser_Func, $_ARRAYCONSTANT_GUI_DOCKHCENTER + $_ARRAYCONSTANT_GUI_DOCKBOTTOM + $_ARRAYCONSTANT_GUI_DOCKHEIGHT)
GUICtrlSetResizing($idExit_Script, $_ARRAYCONSTANT_GUI_DOCKRIGHT + $_ARRAYCONSTANT_GUI_DOCKBOTTOM + $_ARRAYCONSTANT_GUI_DOCKHEIGHT)
GUICtrlSendMsg($idListView, $_ARRAYCONSTANT_WM_SETREDRAW, 0, 0)
Local $iRowIndex, $iColFill
If $iTranspose Then
For $i = $iSubItem_Start To $iSubItem_End
$iRowIndex = __ArrayDisplay_AddItem($idListView, "NULL")
If $iNoRow Then
$iColFill = 0
Else
__ArrayDisplay_AddSubItem($idListView, $iRowIndex, "Col " & $i, 0)
$iColFill = 1
EndIf
For $j = $iItem_Start To $iItem_End
If $iDimension = 2 Then
$vTmp = $aArray[$j][$i]
Else
$vTmp = $aArray[$j]
EndIf
Switch VarGetType($vTmp)
Case "Array"
__ArrayDisplay_AddSubItem($idListView, $iRowIndex, "{Array}", $iColFill)
Case Else
__ArrayDisplay_AddSubItem($idListView, $iRowIndex, $vTmp, $iColFill)
EndSwitch
$iColFill += 1
Next
Next
Else
For $i = $iItem_Start To $iItem_End
$iRowIndex = __ArrayDisplay_AddItem($idListView, "NULL")
If $iNoRow Then
$iColFill = 0
Else
__ArrayDisplay_AddSubItem($idListView, $iRowIndex, "Row " & $i, 0)
$iColFill = 1
EndIf
For $j = $iSubItem_Start To $iSubItem_End
If $iDimension = 2 Then
$vTmp = $aArray[$i][$j]
Else
$vTmp = $aArray[$i]
EndIf
Switch VarGetType($vTmp)
Case "Array"
__ArrayDisplay_AddSubItem($idListView, $iRowIndex, "{Array}", $iColFill)
Case Else
__ArrayDisplay_AddSubItem($idListView, $iRowIndex, $vTmp, $iColFill)
EndSwitch
$iColFill += 1
Next
Next
EndIf
If $iColAlign Then
For $i = 0 To $iColFill - 1
__ArrayDisplay_JustifyColumn($idListView, $i, $iColAlign / 2)
Next
EndIf
GUICtrlSendMsg($idListView, $_ARRAYCONSTANT_WM_SETREDRAW, 1, 0)
Local $iBorder = (($iRowIndex > 19) ? (65) : (45))
Local $iWidth = $iBorder, $iColWidth = 0, $aiColWidth[$iColFill], $iMin_ColWidth = 55
For $i = 0 To UBound($aiColWidth) - 1
GUICtrlSendMsg($idListView, $_ARRAYCONSTANT_LVM_SETCOLUMNWIDTH, $i, $_ARRAYCONSTANT_LVSCW_AUTOSIZE)
$iColWidth = GUICtrlSendMsg($idListView, $_ARRAYCONSTANT_LVM_GETCOLUMNWIDTH, $i, 0)
If $sHeader <> "" Then
GUICtrlSendMsg($idListView, $_ARRAYCONSTANT_LVM_SETCOLUMNWIDTH, $i, $_ARRAYCONSTANT_LVSCW_AUTOSIZE_USEHEADER)
Local $iColWidthHeader = GUICtrlSendMsg($idListView, $_ARRAYCONSTANT_LVM_GETCOLUMNWIDTH, $i, 0)
If $iColWidth < $iMin_ColWidth And $iColWidthHeader < $iMin_ColWidth Then
GUICtrlSendMsg($idListView, $_ARRAYCONSTANT_LVM_SETCOLUMNWIDTH, $i, $iMin_ColWidth)
$iColWidth = $iMin_ColWidth
ElseIf $iColWidthHeader < $iColWidth Then
GUICtrlSendMsg($idListView, $_ARRAYCONSTANT_LVM_SETCOLUMNWIDTH, $i, $iColWidth)
Else
$iColWidth = $iColWidthHeader
EndIf
Else
If $iColWidth < $iMin_ColWidth Then
GUICtrlSendMsg($idListView, $_ARRAYCONSTANT_LVM_SETCOLUMNWIDTH, $i, $iMin_ColWidth)
$iColWidth = $iMin_ColWidth
EndIf
EndIf
$iWidth += $iColWidth
$aiColWidth[$i] = $iColWidth
Next
If $iWidth > @DesktopWidth - 100 Then
$iWidth = $iBorder
For $i = 0 To UBound($aiColWidth) - 1
If $aiColWidth[$i] > $iMax_ColWidth Then
GUICtrlSendMsg($idListView, $_ARRAYCONSTANT_LVM_SETCOLUMNWIDTH, $i, $iMax_ColWidth)
$iWidth += $iMax_ColWidth
Else
$iWidth += $aiColWidth[$i]
EndIf
Next
EndIf
If $iWidth > @DesktopWidth - 100 Then
$iWidth = @DesktopWidth - 100
ElseIf $iWidth < $iMinSize Then
$iWidth = $iMinSize
EndIf
Local $tRECT = DllStructCreate("struct; long Left;long Top;long Right;long Bottom; endstruct")
DllCall("user32.dll", "struct*", "SendMessageW", "hwnd", GUICtrlGetHandle($idListView), "uint", $_ARRAYCONSTANT_LVM_GETITEMRECT, "wparam", 0, "struct*", $tRECT)
Local $aiWin_Pos = WinGetPos($hGUI)
Local $aiLV_Pos = ControlGetPos($hGUI, "", $idListView)
$iHeight = (($iRowIndex + 4) * (DllStructGetData($tRECT, "Bottom") - DllStructGetData($tRECT, "Top"))) + $aiWin_Pos[3] - $aiLV_Pos[3]
If $iHeight > @DesktopHeight - 100 Then
$iHeight = @DesktopHeight - 100
ElseIf $iHeight < $iMinSize Then
$iHeight = $iMinSize
EndIf
If $iVerbose Then SplashOff()
GUISetState(@SW_HIDE, $hGUI)
WinMove($hGUI, "", (@DesktopWidth - $iWidth) / 2, (@DesktopHeight - $iHeight) / 2, $iWidth, $iHeight)
GUISetState(@SW_SHOW, $hGUI)
Local $iOnEventMode = Opt("GUIOnEventMode", 0), $iMsg
__ArrayDisplay_RegisterSortCallBack($idListView, 2, True, "__ArrayDisplay_SortCallBack")
While 1
$iMsg = GUIGetMsg()
Switch $iMsg
Case $_ARRAYCONSTANT_GUI_EVENT_CLOSE
ExitLoop
Case $idCopy_ID, $idCopy_Data
Local $iSel_Count = GUICtrlSendMsg($idListView, $_ARRAYCONSTANT_LVM_GETSELECTEDCOUNT, 0, 0)
If $iVerbose And (Not $iSel_Count) And ($iItem_End - $iItem_Start) * ($iSubItem_End - $iSubItem_Start) > 10000 Then
SplashTextOn($sMsgBoxTitle, "Copying data" & @CRLF & @CRLF & "Please be patient", 300, 100)
EndIf
Local $sClip = "", $sItem, $aSplit
For $i = 0 To GUICtrlSendMsg($idListView, $_ARRAYCONSTANT_LVM_GETITEMCOUNT, 0, 0) - 1
If $iSel_Count And Not (GUICtrlSendMsg($idListView, $_ARRAYCONSTANT_LVM_GETITEMSTATE, $i, $_ARRAYCONSTANT_LVIS_SELECTED) <> 0) Then
ContinueLoop
EndIf
$sItem = __ArrayDisplay_GetItemTextString($idListView, $i)
If $iMsg = $idCopy_ID And $iNoRow Then
$sItem = "Row " & ($i + (($iTranspose) ? ($iSubItem_Start) : ($iItem_Start))) & $sCurr_Separator & $sItem
EndIf
If $iMsg = $idCopy_Data And Not $iNoRow Then
$sItem = StringRegExpReplace($sItem, "^Row\s\d+\|(.*)$", "$1")
EndIf
If $iCW_ColWidth Then
$aSplit = StringSplit($sItem, $sCurr_Separator)
$sItem = ""
For $j = 1 To $aSplit[0]
$sItem &= StringFormat("%-" & $iCW_ColWidth + 1 & "s", StringLeft($aSplit[$j], $iCW_ColWidth))
Next
Else
$sItem = StringReplace($sItem, $sCurr_Separator, $vUser_Separator)
EndIf
$sClip &= $sItem & @CRLF
Next
$sItem = $sHeader
If $iMsg = $idCopy_ID Then
$sItem = $sHeader
If $iNoRow Then
$sItem = "Row|" & $sItem
EndIf
If $iCW_ColWidth Then
$aSplit = StringSplit($sItem, $sCurr_Separator)
$sItem = ""
For $j = 1 To $aSplit[0]
$sItem &= StringFormat("%-" & $iCW_ColWidth + 1 & "s", StringLeft($aSplit[$j], $iCW_ColWidth))
Next
Else
$sItem = StringReplace($sItem, $sCurr_Separator, $vUser_Separator)
EndIf
$sClip = $sItem & @CRLF & $sClip
EndIf
ClipPut($sClip)
SplashOff()
GUICtrlSetState($idListView, $_ARRAYCONSTANT_GUI_FOCUS)
Case $idListView
__ArrayDisplay_SortItems($idListView, GUICtrlGetState($idListView))
Case $idUser_Func
Local $aiSelItems[1] = [0]
For $i = 0 To GUICtrlSendMsg($idListView, $_ARRAYCONSTANT_LVM_GETITEMCOUNT, 0, 0) - 1
If (GUICtrlSendMsg($idListView, $_ARRAYCONSTANT_LVM_GETITEMSTATE, $i, $_ARRAYCONSTANT_LVIS_SELECTED) <> 0) Then
$aiSelItems[0] += 1
ReDim $aiSelItems[$aiSelItems[0] + 1]
$aiSelItems[$aiSelItems[0]] = $i + $iItem_Start
EndIf
Next
$hUser_Function($aArray, $aiSelItems)
GUICtrlSetState($idListView, $_ARRAYCONSTANT_GUI_FOCUS)
Case $idExit_Script
GUIDelete($hGUI)
Exit
EndSwitch
WEnd
GUIDelete($hGUI)
Opt("GUICoordMode", $iCoordMode)
Opt("GUIOnEventMode", $iOnEventMode)
Return 1
EndFunc
Func __ArrayDisplay_RegisterSortCallBack($hWnd, $vCompareType = 2, $bArrows = True, $sSort_Callback = "__ArrayDisplay_SortCallBack")
#Au3Stripper_Ignore_Funcs=$sSort_Callback
If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
Local $hHeader =  HWnd(GUICtrlSendMsg($hWnd, 0x101F, 0, 0))
$__g_aArrayDisplay_SortInfo[1] = $hWnd
$__g_aArrayDisplay_SortInfo[2] = DllCallbackRegister($sSort_Callback, "int", "int;int;hwnd")
$__g_aArrayDisplay_SortInfo[3] = -1
$__g_aArrayDisplay_SortInfo[4] = -1
$__g_aArrayDisplay_SortInfo[5] = 1
$__g_aArrayDisplay_SortInfo[6] = -1
$__g_aArrayDisplay_SortInfo[7] = 0
$__g_aArrayDisplay_SortInfo[8] = $vCompareType
$__g_aArrayDisplay_SortInfo[9] = $bArrows
$__g_aArrayDisplay_SortInfo[10] = $hHeader
Return $__g_aArrayDisplay_SortInfo[2] <> 0
EndFunc
#Au3Stripper_Ignore_Funcs=__ArrayDisplay_SortCallBack
Func __ArrayDisplay_SortCallBack($nItem1, $nItem2, $hWnd)
If $__g_aArrayDisplay_SortInfo[3] = $__g_aArrayDisplay_SortInfo[4] Then
If Not $__g_aArrayDisplay_SortInfo[7] Then
$__g_aArrayDisplay_SortInfo[5] *= -1
$__g_aArrayDisplay_SortInfo[7] = 1
EndIf
Else
$__g_aArrayDisplay_SortInfo[7] = 1
EndIf
$__g_aArrayDisplay_SortInfo[6] = $__g_aArrayDisplay_SortInfo[3]
Local $sVal1 = __ArrayDisplay_GetItemText($hWnd, $nItem1, $__g_aArrayDisplay_SortInfo[3])
Local $sVal2 = __ArrayDisplay_GetItemText($hWnd, $nItem2, $__g_aArrayDisplay_SortInfo[3])
If $__g_aArrayDisplay_SortInfo[8] = 1 Then
If (StringIsFloat($sVal1) Or StringIsInt($sVal1)) Then $sVal1 = Number($sVal1)
If (StringIsFloat($sVal2) Or StringIsInt($sVal2)) Then $sVal2 = Number($sVal2)
EndIf
Local $nResult
If $__g_aArrayDisplay_SortInfo[8] < 2 Then
$nResult = 0
If $sVal1 < $sVal2 Then
$nResult = -1
ElseIf $sVal1 > $sVal2 Then
$nResult = 1
EndIf
Else
$nResult = DllCall('shlwapi.dll', 'int', 'StrCmpLogicalW', 'wstr', $sVal1, 'wstr', $sVal2)[0]
EndIf
$nResult = $nResult * $__g_aArrayDisplay_SortInfo[5]
Return $nResult
EndFunc
Func __ArrayDisplay_SortItems($hWnd, $iCol)
If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
Local $pFunction = DllCallbackGetPtr($__g_aArrayDisplay_SortInfo[2])
$__g_aArrayDisplay_SortInfo[3] = $iCol
$__g_aArrayDisplay_SortInfo[7] = 0
$__g_aArrayDisplay_SortInfo[4] = $__g_aArrayDisplay_SortInfo[6]
Local $aResult = DllCall("user32.dll", "lresult", "SendMessageW", "hwnd", $hWnd, "uint", 0x1051, "hwnd", $hWnd, "ptr", $pFunction)
If $aResult[0] <> 0 Then
If $__g_aArrayDisplay_SortInfo[9] Then
Local $hHeader = $__g_aArrayDisplay_SortInfo[10], $iFormat
For $x = 0 To __ArrayDisplay_GetItemCount($hHeader) - 1
$iFormat = __ArrayDisplay_GetItemFormat($hHeader, $x)
If BitAND($iFormat, 0x00000200) Then
__ArrayDisplay_SetItemFormat($hHeader, $x, BitXOR($iFormat, 0x00000200))
ElseIf BitAND($iFormat, 0x00000400) Then
__ArrayDisplay_SetItemFormat($hHeader, $x, BitXOR($iFormat, 0x00000400))
EndIf
Next
$iFormat = __ArrayDisplay_GetItemFormat($hHeader, $iCol)
If $__g_aArrayDisplay_SortInfo[5] = 1 Then
__ArrayDisplay_SetItemFormat($hHeader, $iCol, BitOR($iFormat, 0x00000400))
Else
__ArrayDisplay_SetItemFormat($hHeader, $iCol, BitOR($iFormat, 0x00000200))
EndIf
EndIf
Return True
EndIf
Return False
EndFunc
Func __ArrayDisplay_AddItem($hWnd, $sText)
Local $tItem = DllStructCreate($_ARRAYCONSTANT_tagLVITEM)
DllStructSetData($tItem, "Param", 0)
Local $iBuffer = StringLen($sText) + 1
Local $tBuffer = DllStructCreate("wchar Text[" & $iBuffer & "]")
$iBuffer *= 2
DllStructSetData($tBuffer, "Text", $sText)
DllStructSetData($tItem, "Text", DllStructGetPtr($tBuffer))
DllStructSetData($tItem, "TextMax", $iBuffer)
Local $iMask = 0x00000005
DllStructSetData($tItem, "Mask", $iMask)
DllStructSetData($tItem, "Item", 999999999)
DllStructSetData($tItem, "Image", -1)
Local $pItem = DllStructGetPtr($tItem)
Local $iRet = GUICtrlSendMsg($hWnd, 0x104D, 0, $pItem)
Return $iRet
EndFunc
Func __ArrayDisplay_AddSubItem($hWnd, $iIndex, $sText, $iSubItem)
Local $iBuffer = StringLen($sText) + 1
Local $tBuffer = DllStructCreate("wchar Text[" & $iBuffer & "]")
$iBuffer *= 2
Local $pBuffer = DllStructGetPtr($tBuffer)
Local $tItem = DllStructCreate($_ARRAYCONSTANT_tagLVITEM)
Local $iMask = 0x00000001
DllStructSetData($tBuffer, "Text", $sText)
DllStructSetData($tItem, "Mask", $iMask)
DllStructSetData($tItem, "Item", $iIndex)
DllStructSetData($tItem, "SubItem", $iSubItem)
DllStructSetData($tItem, "Image", -1)
Local $pItem = DllStructGetPtr($tItem)
DllStructSetData($tItem, "Text", $pBuffer)
Local $iRet = GUICtrlSendMsg($hWnd, 0x104C, 0, $pItem)
Return $iRet <> 0
EndFunc
Func __ArrayDisplay_GetColumnCount($hWnd)
Local $hHeader = HWnd(GUICtrlSendMsg($hWnd, 0x101F, 0, 0))
Return __ArrayDisplay_GetItemCount($hHeader)
EndFunc
Func __ArrayDisplay_GetHeader($hWnd)
Return HWnd(GUICtrlSendMsg($hWnd, 0x101F, 0, 0))
EndFunc
Func __ArrayDisplay_GetItem($hWnd, $iIndex, ByRef $tItem)
Local $aResult = DllCall("user32.dll", "lresult", "SendMessageW", "hwnd", $hWnd, "uint", 0x120B, "wparam", $iIndex, "struct*", $tItem)
Return $aResult[0] <> 0
EndFunc
Func __ArrayDisplay_GetItemCount($hWnd)
Local $aResult = DllCall("user32.dll", "lresult", "SendMessageW", "hwnd", $hWnd, "uint", 0x1200, "wparam", 0, "lparam", 0)
Return $aResult[0]
EndFunc
Func __ArrayDisplay_GetItemFormat($hWnd, $iIndex)
Local $tItem = DllStructCreate($_ARRAYCONSTANT_tagHDITEM)
DllStructSetData($tItem, "Mask", 0x00000004)
__ArrayDisplay_GetItem($hWnd, $iIndex, $tItem)
Return DllStructGetData($tItem, "Fmt")
EndFunc
Func __ArrayDisplay_GetItemText($hWnd, $iIndex, $iSubItem = 0)
Local $tBuffer = DllStructCreate("wchar Text[4096]")
Local $pBuffer = DllStructGetPtr($tBuffer)
Local $tItem = DllStructCreate($_ARRAYCONSTANT_tagLVITEM)
DllStructSetData($tItem, "SubItem", $iSubItem)
DllStructSetData($tItem, "TextMax", 4096)
DllStructSetData($tItem, "Text", $pBuffer)
If IsHWnd($hWnd) Then
DllCall("user32.dll", "lresult", "SendMessageW", "hwnd", $hWnd, "uint", 0x1073, "wparam", $iIndex, "struct*", $tItem)
Else
Local $pItem = DllStructGetPtr($tItem)
GUICtrlSendMsg($hWnd, 0x1073, $iIndex, $pItem)
EndIf
Return DllStructGetData($tBuffer, "Text")
EndFunc
Func __ArrayDisplay_GetItemTextString($hWnd, $iItem)
Local $sRow = "", $sSeparatorChar = Opt('GUIDataSeparatorChar')
Local $iSelected = $iItem
For $x = 0 To __ArrayDisplay_GetColumnCount($hWnd) - 1
$sRow &= __ArrayDisplay_GetItemText($hWnd, $iSelected, $x) & $sSeparatorChar
Next
Return StringTrimRight($sRow, 1)
EndFunc
Func __ArrayDisplay_JustifyColumn($idListView, $iIndex, $iAlign = -1)
Local $tColumn = DllStructCreate("uint Mask;int Fmt;int CX;ptr Text;int TextMax;int SubItem;int Image;int Order;int cxMin;int cxDefault;int cxIdeal")
If $iAlign < 0 Or $iAlign > 2 Then $iAlign = 0
DllStructSetData($tColumn, "Mask", 0x01)
DllStructSetData($tColumn, "Fmt", $iAlign)
Local $pColumn = DllStructGetPtr($tColumn)
Local $iRet = GUICtrlSendMsg($idListView, 0x1060 , $iIndex, $pColumn)
Return $iRet <> 0
EndFunc
Func __ArrayDisplay_SetItemFormat($hWnd, $iIndex, $iFormat)
Local $tItem = DllStructCreate($_ARRAYCONSTANT_tagHDITEM)
DllStructSetData($tItem, "Mask", 0x00000004)
DllStructSetData($tItem, "Fmt", $iFormat)
Local $aResult = DllCall("user32.dll", "lresult", "SendMessageW", "hwnd", $hWnd, "uint", 0x120C, "wparam", $iIndex, "struct*", $tItem)
Return $aResult[0] <> 0
EndFunc
Global Enum $ARRAYFILL_FORCE_DEFAULT, $ARRAYFILL_FORCE_SINGLEITEM, $ARRAYFILL_FORCE_INT, $ARRAYFILL_FORCE_NUMBER,  $ARRAYFILL_FORCE_PTR, $ARRAYFILL_FORCE_HWND, $ARRAYFILL_FORCE_STRING, $ARRAYFILL_FORCE_BOOLEAN
Global Enum $ARRAYUNIQUE_NOCOUNT, $ARRAYUNIQUE_COUNT
Global Enum $ARRAYUNIQUE_AUTO, $ARRAYUNIQUE_FORCE32, $ARRAYUNIQUE_FORCE64, $ARRAYUNIQUE_MATCH, $ARRAYUNIQUE_DISTINCT
Func _ArrayAdd(ByRef $aArray, $vValue, $iStart = 0, $sDelim_Item = "|", $sDelim_Row = @CRLF, $iForce = $ARRAYFILL_FORCE_DEFAULT)
If $iStart = Default Then $iStart = 0
If $sDelim_Item = Default Then $sDelim_Item = "|"
If $sDelim_Row = Default Then $sDelim_Row = @CRLF
If $iForce = Default Then $iForce = $ARRAYFILL_FORCE_DEFAULT
If Not IsArray($aArray) Then Return SetError(1, 0, -1)
Local $iDim_1 = UBound($aArray, $UBOUND_ROWS)
Local $hDataType = 0
Switch $iForce
Case $ARRAYFILL_FORCE_INT
$hDataType = Int
Case $ARRAYFILL_FORCE_NUMBER
$hDataType = Number
Case $ARRAYFILL_FORCE_PTR
$hDataType = Ptr
Case $ARRAYFILL_FORCE_HWND
$hDataType = Hwnd
Case $ARRAYFILL_FORCE_STRING
$hDataType = String
Case $ARRAYFILL_FORCE_BOOLEAN
$hDataType = "Boolean"
EndSwitch
Switch UBound($aArray, $UBOUND_DIMENSIONS)
Case 1
If $iForce = $ARRAYFILL_FORCE_SINGLEITEM Then
ReDim $aArray[$iDim_1 + 1]
$aArray[$iDim_1] = $vValue
Return $iDim_1
EndIf
If IsArray($vValue) Then
If UBound($vValue, $UBOUND_DIMENSIONS) <> 1 Then Return SetError(5, 0, -1)
$hDataType = 0
Else
Local $aTmp = StringSplit($vValue, $sDelim_Item, $STR_NOCOUNT + $STR_ENTIRESPLIT)
If UBound($aTmp, $UBOUND_ROWS) = 1 Then
$aTmp[0] = $vValue
EndIf
$vValue = $aTmp
EndIf
Local $iAdd = UBound($vValue, $UBOUND_ROWS)
ReDim $aArray[$iDim_1 + $iAdd]
For $i = 0 To $iAdd - 1
If String($hDataType) = "Boolean" Then
Switch $vValue[$i]
Case "True", "1"
$aArray[$iDim_1 + $i] = True
Case "False", "0", ""
$aArray[$iDim_1 + $i] = False
EndSwitch
ElseIf IsFunc($hDataType) Then
$aArray[$iDim_1 + $i] = $hDataType($vValue[$i])
Else
$aArray[$iDim_1 + $i] = $vValue[$i]
EndIf
Next
Return $iDim_1 + $iAdd - 1
Case 2
Local $iDim_2 = UBound($aArray, $UBOUND_COLUMNS)
If $iStart < 0 Or $iStart > $iDim_2 - 1 Then Return SetError(4, 0, -1)
Local $iValDim_1, $iValDim_2 = 0, $iColCount
If IsArray($vValue) Then
If UBound($vValue, $UBOUND_DIMENSIONS) <> 2 Then Return SetError(5, 0, -1)
$iValDim_1 = UBound($vValue, $UBOUND_ROWS)
$iValDim_2 = UBound($vValue, $UBOUND_COLUMNS)
$hDataType = 0
Else
Local $aSplit_1 = StringSplit($vValue, $sDelim_Row, $STR_NOCOUNT + $STR_ENTIRESPLIT)
$iValDim_1 = UBound($aSplit_1, $UBOUND_ROWS)
Local $aTmp[$iValDim_1][0], $aSplit_2
For $i = 0 To $iValDim_1 - 1
$aSplit_2 = StringSplit($aSplit_1[$i], $sDelim_Item, $STR_NOCOUNT + $STR_ENTIRESPLIT)
$iColCount = UBound($aSplit_2)
If $iColCount > $iValDim_2 Then
$iValDim_2 = $iColCount
ReDim $aTmp[$iValDim_1][$iValDim_2]
EndIf
For $j = 0 To $iColCount - 1
$aTmp[$i][$j] = $aSplit_2[$j]
Next
Next
$vValue = $aTmp
EndIf
If UBound($vValue, $UBOUND_COLUMNS) + $iStart > UBound($aArray, $UBOUND_COLUMNS) Then Return SetError(3, 0, -1)
ReDim $aArray[$iDim_1 + $iValDim_1][$iDim_2]
For $iWriteTo_Index = 0 To $iValDim_1 - 1
For $j = 0 To $iDim_2 - 1
If $j < $iStart Then
$aArray[$iWriteTo_Index + $iDim_1][$j] = ""
ElseIf $j - $iStart > $iValDim_2 - 1 Then
$aArray[$iWriteTo_Index + $iDim_1][$j] = ""
Else
If String($hDataType) = "Boolean" Then
Switch $vValue[$iWriteTo_Index][$j - $iStart]
Case "True", "1"
$aArray[$iWriteTo_Index + $iDim_1][$j] = True
Case "False", "0", ""
$aArray[$iWriteTo_Index + $iDim_1][$j] = False
EndSwitch
ElseIf IsFunc($hDataType) Then
$aArray[$iWriteTo_Index + $iDim_1][$j] = $hDataType($vValue[$iWriteTo_Index][$j - $iStart])
Else
$aArray[$iWriteTo_Index + $iDim_1][$j] = $vValue[$iWriteTo_Index][$j - $iStart]
EndIf
EndIf
Next
Next
Case Else
Return SetError(2, 0, -1)
EndSwitch
Return UBound($aArray, $UBOUND_ROWS) - 1
EndFunc
Func _ArrayBinarySearch(Const ByRef $aArray, $vValue, $iStart = 0, $iEnd = 0, $iColumn = 0)
If $iStart = Default Then $iStart = 0
If $iEnd = Default Then $iEnd = 0
If $iColumn = Default Then $iColumn = 0
If Not IsArray($aArray) Then Return SetError(1, 0, -1)
Local $iDim_1 = UBound($aArray, $UBOUND_ROWS)
If $iDim_1 = 0 Then Return SetError(6, 0, -1)
If $iEnd < 1 Or $iEnd > $iDim_1 - 1 Then $iEnd = $iDim_1 - 1
If $iStart < 0 Then $iStart = 0
If $iStart > $iEnd Then Return SetError(4, 0, -1)
Local $iMid = Int(($iEnd + $iStart) / 2)
Switch UBound($aArray, $UBOUND_DIMENSIONS)
Case 1
If $aArray[$iStart] > $vValue Or $aArray[$iEnd] < $vValue Then Return SetError(2, 0, -1)
While $iStart <= $iMid And $vValue <> $aArray[$iMid]
If $vValue < $aArray[$iMid] Then
$iEnd = $iMid - 1
Else
$iStart = $iMid + 1
EndIf
$iMid = Int(($iEnd + $iStart) / 2)
WEnd
If $iStart > $iEnd Then Return SetError(3, 0, -1)
Case 2
Local $iDim_2 = UBound($aArray, $UBOUND_COLUMNS) - 1
If $iColumn < 0 Or $iColumn > $iDim_2 Then Return SetError(7, 0, -1)
If $aArray[$iStart][$iColumn] > $vValue Or $aArray[$iEnd][$iColumn] < $vValue Then Return SetError(2, 0, -1)
While $iStart <= $iMid And $vValue <> $aArray[$iMid][$iColumn]
If $vValue < $aArray[$iMid][$iColumn] Then
$iEnd = $iMid - 1
Else
$iStart = $iMid + 1
EndIf
$iMid = Int(($iEnd + $iStart) / 2)
WEnd
If $iStart > $iEnd Then Return SetError(3, 0, -1)
Case Else
Return SetError(5, 0, -1)
EndSwitch
Return $iMid
EndFunc
Func _ArrayColDelete(ByRef $aArray, $iColumn, $bConvert = False)
If $bConvert = Default Then $bConvert = False
If Not IsArray($aArray) Then Return SetError(1, 0, -1)
Local $iDim_1 = UBound($aArray, $UBOUND_ROWS)
If UBound($aArray, $UBOUND_DIMENSIONS) <> 2 Then Return SetError(2, 0, -1)
Local $iDim_2 = UBound($aArray, $UBOUND_COLUMNS)
Switch $iDim_2
Case 2
If $iColumn < 0 Or $iColumn > 1 Then Return SetError(3, 0, -1)
If $bConvert Then
Local $aTempArray[$iDim_1]
For $i = 0 To $iDim_1 - 1
$aTempArray[$i] = $aArray[$i][(Not $iColumn)]
Next
$aArray = $aTempArray
Else
ContinueCase
EndIf
Case Else
If $iColumn < 0 Or $iColumn > $iDim_2 - 1 Then Return SetError(3, 0, -1)
For $i = 0 To $iDim_1 - 1
For $j = $iColumn To $iDim_2 - 2
$aArray[$i][$j] = $aArray[$i][$j + 1]
Next
Next
ReDim $aArray[$iDim_1][$iDim_2 - 1]
EndSwitch
Return UBound($aArray, $UBOUND_COLUMNS)
EndFunc
Func _ArrayColInsert(ByRef $aArray, $iColumn)
If Not IsArray($aArray) Then Return SetError(1, 0, -1)
Local $iDim_1 = UBound($aArray, $UBOUND_ROWS)
Switch UBound($aArray, $UBOUND_DIMENSIONS)
Case 1
Local $aTempArray[$iDim_1][2]
Switch $iColumn
Case 0, 1
For $i = 0 To $iDim_1 - 1
$aTempArray[$i][(Not $iColumn)] = $aArray[$i]
Next
Case Else
Return SetError(3, 0, -1)
EndSwitch
$aArray = $aTempArray
Case 2
Local $iDim_2 = UBound($aArray, $UBOUND_COLUMNS)
If $iColumn < 0 Or $iColumn > $iDim_2 Then Return SetError(3, 0, -1)
ReDim $aArray[$iDim_1][$iDim_2 + 1]
For $i = 0 To $iDim_1 - 1
For $j = $iDim_2 To $iColumn + 1 Step -1
$aArray[$i][$j] = $aArray[$i][$j - 1]
Next
$aArray[$i][$iColumn] = ""
Next
Case Else
Return SetError(2, 0, -1)
EndSwitch
Return UBound($aArray, $UBOUND_COLUMNS)
EndFunc
Func _ArrayCombinations(Const ByRef $aArray, $iSet, $sDelimiter = "")
If $sDelimiter = Default Then $sDelimiter = ""
If Not IsArray($aArray) Then Return SetError(1, 0, 0)
If UBound($aArray, $UBOUND_DIMENSIONS) <> 1 Then Return SetError(2, 0, 0)
Local $iN = UBound($aArray)
Local $iR = $iSet
Local $aIdx[$iR]
For $i = 0 To $iR - 1
$aIdx[$i] = $i
Next
Local $iTotal = __Array_Combinations($iN, $iR)
Local $iLeft = $iTotal
Local $aResult[$iTotal + 1]
$aResult[0] = $iTotal
Local $iCount = 1
While $iLeft > 0
__Array_GetNext($iN, $iR, $iLeft, $iTotal, $aIdx)
For $i = 0 To $iSet - 1
$aResult[$iCount] &= $aArray[$aIdx[$i]] & $sDelimiter
Next
If $sDelimiter <> "" Then $aResult[$iCount] = StringTrimRight($aResult[$iCount], 1)
$iCount += 1
WEnd
Return $aResult
EndFunc
Func _ArrayConcatenate(ByRef $aArrayTarget, Const ByRef $aArraySource, $iStart = 0)
If $iStart = Default Then $iStart = 0
If Not IsArray($aArrayTarget) Then Return SetError(1, 0, -1)
If Not IsArray($aArraySource) Then Return SetError(2, 0, -1)
Local $iDim_Total_Tgt = UBound($aArrayTarget, $UBOUND_DIMENSIONS)
Local $iDim_Total_Src = UBound($aArraySource, $UBOUND_DIMENSIONS)
Local $iDim_1_Tgt = UBound($aArrayTarget, $UBOUND_ROWS)
Local $iDim_1_Src = UBound($aArraySource, $UBOUND_ROWS)
If $iStart < 0 Or $iStart > $iDim_1_Src - 1 Then Return SetError(6, 0, -1)
Switch $iDim_Total_Tgt
Case 1
If $iDim_Total_Src <> 1 Then Return SetError(4, 0, -1)
ReDim $aArrayTarget[$iDim_1_Tgt + $iDim_1_Src - $iStart]
For $i = $iStart To $iDim_1_Src - 1
$aArrayTarget[$iDim_1_Tgt + $i - $iStart] = $aArraySource[$i]
Next
Case 2
If $iDim_Total_Src <> 2 Then Return SetError(4, 0, -1)
Local $iDim_2_Tgt = UBound($aArrayTarget, $UBOUND_COLUMNS)
If UBound($aArraySource, $UBOUND_COLUMNS) <> $iDim_2_Tgt Then Return SetError(5, 0, -1)
ReDim $aArrayTarget[$iDim_1_Tgt + $iDim_1_Src - $iStart][$iDim_2_Tgt]
For $i = $iStart To $iDim_1_Src - 1
For $j = 0 To $iDim_2_Tgt - 1
$aArrayTarget[$iDim_1_Tgt + $i - $iStart][$j] = $aArraySource[$i][$j]
Next
Next
Case Else
Return SetError(3, 0, -1)
EndSwitch
Return UBound($aArrayTarget, $UBOUND_ROWS)
EndFunc
Func _ArrayDelete(ByRef $aArray, $vRange)
If Not IsArray($aArray) Then Return SetError(1, 0, -1)
Local $iDim_1 = UBound($aArray, $UBOUND_ROWS) - 1
If IsArray($vRange) Then
If UBound($vRange, $UBOUND_DIMENSIONS) <> 1 Or UBound($vRange, $UBOUND_ROWS) < 2 Then Return SetError(4, 0, -1)
Else
Local $iNumber, $aSplit_1, $aSplit_2
$vRange = StringStripWS($vRange, 8)
$aSplit_1 = StringSplit($vRange, ";")
$vRange = ""
For $i = 1 To $aSplit_1[0]
If Not StringRegExp($aSplit_1[$i], "^\d+(-\d+)?$") Then Return SetError(3, 0, -1)
$aSplit_2 = StringSplit($aSplit_1[$i], "-")
Switch $aSplit_2[0]
Case 1
$vRange &= $aSplit_2[1] & ";"
Case 2
If Number($aSplit_2[2]) >= Number($aSplit_2[1]) Then
$iNumber = $aSplit_2[1] - 1
Do
$iNumber += 1
$vRange &= $iNumber & ";"
Until $iNumber = $aSplit_2[2]
EndIf
EndSwitch
Next
$vRange = StringSplit(StringTrimRight($vRange, 1), ";")
EndIf
If $vRange[1] < 0 Or $vRange[$vRange[0]] > $iDim_1 Then Return SetError(5, 0, -1)
Local $iCopyTo_Index = 0
Switch UBound($aArray, $UBOUND_DIMENSIONS)
Case 1
For $i = 1 To $vRange[0]
$aArray[$vRange[$i]] = ChrW(0xFAB1)
Next
For $iReadFrom_Index = 0 To $iDim_1
If $aArray[$iReadFrom_Index] == ChrW(0xFAB1) Then
ContinueLoop
Else
If $iReadFrom_Index <> $iCopyTo_Index Then
$aArray[$iCopyTo_Index] = $aArray[$iReadFrom_Index]
EndIf
$iCopyTo_Index += 1
EndIf
Next
ReDim $aArray[$iDim_1 - $vRange[0] + 1]
Case 2
Local $iDim_2 = UBound($aArray, $UBOUND_COLUMNS) - 1
For $i = 1 To $vRange[0]
$aArray[$vRange[$i]][0] = ChrW(0xFAB1)
Next
For $iReadFrom_Index = 0 To $iDim_1
If $aArray[$iReadFrom_Index][0] == ChrW(0xFAB1) Then
ContinueLoop
Else
If $iReadFrom_Index <> $iCopyTo_Index Then
For $j = 0 To $iDim_2
$aArray[$iCopyTo_Index][$j] = $aArray[$iReadFrom_Index][$j]
Next
EndIf
$iCopyTo_Index += 1
EndIf
Next
ReDim $aArray[$iDim_1 - $vRange[0] + 1][$iDim_2 + 1]
Case Else
Return SetError(2, 0, False)
EndSwitch
Return UBound($aArray, $UBOUND_ROWS)
EndFunc
Func _ArrayDisplay(Const ByRef $aArray, $sTitle = Default, $sArrayRange = Default, $iFlags = Default, $vUser_Separator = Default, $sHeader = Default, $iMax_ColWidth = Default)
#forceref $vUser_Separator
Local $iRet = __ArrayDisplay_Share($aArray, $sTitle, $sArrayRange, $iFlags, Default, $sHeader, $iMax_ColWidth, 0, False)
Return SetError(@error, @extended, $iRet)
EndFunc
Func _ArrayExtract(Const ByRef $aArray, $iStart_Row = -1, $iEnd_Row = -1, $iStart_Col = -1, $iEnd_Col = -1)
If $iStart_Row = Default Then $iStart_Row = -1
If $iEnd_Row = Default Then $iEnd_Row = -1
If $iStart_Col = Default Then $iStart_Col = -1
If $iEnd_Col = Default Then $iEnd_Col = -1
If Not IsArray($aArray) Then Return SetError(1, 0, -1)
Local $iDim_1 = UBound($aArray, $UBOUND_ROWS) - 1
If $iEnd_Row = -1 Then $iEnd_Row = $iDim_1
If $iStart_Row = -1 Then $iStart_Row = 0
If $iStart_Row < -1 Or $iEnd_Row < -1 Then Return SetError(3, 0, -1)
If $iStart_Row > $iDim_1 Or $iEnd_Row > $iDim_1 Then Return SetError(3, 0, -1)
If $iStart_Row > $iEnd_Row Then Return SetError(4, 0, -1)
Switch UBound($aArray, $UBOUND_DIMENSIONS)
Case 1
Local $aRetArray[$iEnd_Row - $iStart_Row + 1]
For $i = 0 To $iEnd_Row - $iStart_Row
$aRetArray[$i] = $aArray[$i + $iStart_Row]
Next
Return $aRetArray
Case 2
Local $iDim_2 = UBound($aArray, $UBOUND_COLUMNS) - 1
If $iEnd_Col = -1 Then $iEnd_Col = $iDim_2
If $iStart_Col = -1 Then $iStart_Col = 0
If $iStart_Col < -1 Or $iEnd_Col < -1 Then Return SetError(5, 0, -1)
If $iStart_Col > $iDim_2 Or $iEnd_Col > $iDim_2 Then Return SetError(5, 0, -1)
If $iStart_Col > $iEnd_Col Then Return SetError(6, 0, -1)
If $iStart_Col = $iEnd_Col Then
Local $aRetArray[$iEnd_Row - $iStart_Row + 1]
Else
Local $aRetArray[$iEnd_Row - $iStart_Row + 1][$iEnd_Col - $iStart_Col + 1]
EndIf
For $i = 0 To $iEnd_Row - $iStart_Row
For $j = 0 To $iEnd_Col - $iStart_Col
If $iStart_Col = $iEnd_Col Then
$aRetArray[$i] = $aArray[$i + $iStart_Row][$j + $iStart_Col]
Else
$aRetArray[$i][$j] = $aArray[$i + $iStart_Row][$j + $iStart_Col]
EndIf
Next
Next
Return $aRetArray
Case Else
Return SetError(2, 0, -1)
EndSwitch
Return 1
EndFunc
Func _ArrayFindAll(Const ByRef $aArray, $vValue, $iStart = 0, $iEnd = 0, $iCase = 0, $iCompare = 0, $iSubItem = 0, $bRow = False)
If $iStart = Default Then $iStart = 0
If $iEnd = Default Then $iEnd = 0
If $iCase = Default Then $iCase = 0
If $iCompare = Default Then $iCompare = 0
If $iSubItem = Default Then $iSubItem = 0
If $bRow = Default Then $bRow = False
$iStart = _ArraySearch($aArray, $vValue, $iStart, $iEnd, $iCase, $iCompare, 1, $iSubItem, $bRow)
If @error Then Return SetError(@error, 0, -1)
Local $iIndex = 0, $avResult[UBound($aArray, ($bRow ? $UBOUND_COLUMNS : $UBOUND_ROWS))]
Do
$avResult[$iIndex] = $iStart
$iIndex += 1
$iStart = _ArraySearch($aArray, $vValue, $iStart + 1, $iEnd, $iCase, $iCompare, 1, $iSubItem, $bRow)
Until @error
ReDim $avResult[$iIndex]
Return $avResult
EndFunc
Func _ArrayInsert(ByRef $aArray, $vRange, $vValue = "", $iStart = 0, $sDelim_Item = "|", $sDelim_Row = @CRLF, $iForce = $ARRAYFILL_FORCE_DEFAULT)
If $vValue = Default Then $vValue = ""
If $iStart = Default Then $iStart = 0
If $sDelim_Item = Default Then $sDelim_Item = "|"
If $sDelim_Row = Default Then $sDelim_Row = @CRLF
If $iForce = Default Then $iForce = $ARRAYFILL_FORCE_DEFAULT
If Not IsArray($aArray) Then Return SetError(1, 0, -1)
Local $iDim_1 = UBound($aArray, $UBOUND_ROWS) - 1
Local $hDataType = 0
Switch $iForce
Case $ARRAYFILL_FORCE_INT
$hDataType = Int
Case $ARRAYFILL_FORCE_NUMBER
$hDataType = Number
Case $ARRAYFILL_FORCE_PTR
$hDataType = Ptr
Case $ARRAYFILL_FORCE_HWND
$hDataType = Hwnd
Case $ARRAYFILL_FORCE_STRING
$hDataType = String
EndSwitch
Local $aSplit_1, $aSplit_2
If IsArray($vRange) Then
If UBound($vRange, $UBOUND_DIMENSIONS) <> 1 Or UBound($vRange, $UBOUND_ROWS) < 2 Then Return SetError(4, 0, -1)
Else
Local $iNumber
$vRange = StringStripWS($vRange, 8)
$aSplit_1 = StringSplit($vRange, ";")
$vRange = ""
For $i = 1 To $aSplit_1[0]
If Not StringRegExp($aSplit_1[$i], "^\d+(-\d+)?$") Then Return SetError(3, 0, -1)
$aSplit_2 = StringSplit($aSplit_1[$i], "-")
Switch $aSplit_2[0]
Case 1
$vRange &= $aSplit_2[1] & ";"
Case 2
If Number($aSplit_2[2]) >= Number($aSplit_2[1]) Then
$iNumber = $aSplit_2[1] - 1
Do
$iNumber += 1
$vRange &= $iNumber & ";"
Until $iNumber = $aSplit_2[2]
EndIf
EndSwitch
Next
$vRange = StringSplit(StringTrimRight($vRange, 1), ";")
EndIf
If $vRange[1] < 0 Or $vRange[$vRange[0]] > $iDim_1 Then Return SetError(5, 0, -1)
For $i = 2 To $vRange[0]
If $vRange[$i] < $vRange[$i - 1] Then Return SetError(3, 0, -1)
Next
Local $iCopyTo_Index = $iDim_1 + $vRange[0]
Local $iInsertPoint_Index = $vRange[0]
Local $iInsert_Index = $vRange[$iInsertPoint_Index]
Switch UBound($aArray, $UBOUND_DIMENSIONS)
Case 1
If $iForce = $ARRAYFILL_FORCE_SINGLEITEM Then
ReDim $aArray[$iDim_1 + $vRange[0] + 1]
For $iReadFromIndex = $iDim_1 To 0 Step -1
$aArray[$iCopyTo_Index] = $aArray[$iReadFromIndex]
$iCopyTo_Index -= 1
$iInsert_Index = $vRange[$iInsertPoint_Index]
While $iReadFromIndex = $iInsert_Index
$aArray[$iCopyTo_Index] = $vValue
$iCopyTo_Index -= 1
$iInsertPoint_Index -= 1
If $iInsertPoint_Index < 1 Then ExitLoop 2
$iInsert_Index = $vRange[$iInsertPoint_Index]
WEnd
Next
Return $iDim_1 + $vRange[0] + 1
EndIf
ReDim $aArray[$iDim_1 + $vRange[0] + 1]
If IsArray($vValue) Then
If UBound($vValue, $UBOUND_DIMENSIONS) <> 1 Then Return SetError(5, 0, -1)
$hDataType = 0
Else
Local $aTmp = StringSplit($vValue, $sDelim_Item, $STR_NOCOUNT + $STR_ENTIRESPLIT)
If UBound($aTmp, $UBOUND_ROWS) = 1 Then
$aTmp[0] = $vValue
$hDataType = 0
EndIf
$vValue = $aTmp
EndIf
For $iReadFromIndex = $iDim_1 To 0 Step -1
$aArray[$iCopyTo_Index] = $aArray[$iReadFromIndex]
$iCopyTo_Index -= 1
$iInsert_Index = $vRange[$iInsertPoint_Index]
While $iReadFromIndex = $iInsert_Index
If $iInsertPoint_Index <= UBound($vValue, $UBOUND_ROWS) Then
If IsFunc($hDataType) Then
$aArray[$iCopyTo_Index] = $hDataType($vValue[$iInsertPoint_Index - 1])
Else
$aArray[$iCopyTo_Index] = $vValue[$iInsertPoint_Index - 1]
EndIf
Else
$aArray[$iCopyTo_Index] = ""
EndIf
$iCopyTo_Index -= 1
$iInsertPoint_Index -= 1
If $iInsertPoint_Index = 0 Then ExitLoop 2
$iInsert_Index = $vRange[$iInsertPoint_Index]
WEnd
Next
Case 2
Local $iDim_2 = UBound($aArray, $UBOUND_COLUMNS)
If $iStart < 0 Or $iStart > $iDim_2 - 1 Then Return SetError(6, 0, -1)
Local $iValDim_1, $iValDim_2
If IsArray($vValue) Then
If UBound($vValue, $UBOUND_DIMENSIONS) <> 2 Then Return SetError(7, 0, -1)
$iValDim_1 = UBound($vValue, $UBOUND_ROWS)
$iValDim_2 = UBound($vValue, $UBOUND_COLUMNS)
$hDataType = 0
Else
$aSplit_1 = StringSplit($vValue, $sDelim_Row, $STR_NOCOUNT + $STR_ENTIRESPLIT)
$iValDim_1 = UBound($aSplit_1, $UBOUND_ROWS)
StringReplace($aSplit_1[0], $sDelim_Item, "")
$iValDim_2 = @extended + 1
Local $aTmp[$iValDim_1][$iValDim_2]
For $i = 0 To $iValDim_1 - 1
$aSplit_2 = StringSplit($aSplit_1[$i], $sDelim_Item, $STR_NOCOUNT + $STR_ENTIRESPLIT)
For $j = 0 To $iValDim_2 - 1
$aTmp[$i][$j] = $aSplit_2[$j]
Next
Next
$vValue = $aTmp
EndIf
If UBound($vValue, $UBOUND_COLUMNS) + $iStart > UBound($aArray, $UBOUND_COLUMNS) Then Return SetError(8, 0, -1)
ReDim $aArray[$iDim_1 + $vRange[0] + 1][$iDim_2]
For $iReadFromIndex = $iDim_1 To 0 Step -1
For $j = 0 To $iDim_2 - 1
$aArray[$iCopyTo_Index][$j] = $aArray[$iReadFromIndex][$j]
Next
$iCopyTo_Index -= 1
$iInsert_Index = $vRange[$iInsertPoint_Index]
While $iReadFromIndex = $iInsert_Index
For $j = 0 To $iDim_2 - 1
If $j < $iStart Then
$aArray[$iCopyTo_Index][$j] = ""
ElseIf $j - $iStart > $iValDim_2 - 1 Then
$aArray[$iCopyTo_Index][$j] = ""
Else
If $iInsertPoint_Index - 1 < $iValDim_1 Then
If IsFunc($hDataType) Then
$aArray[$iCopyTo_Index][$j] = $hDataType($vValue[$iInsertPoint_Index - 1][$j - $iStart])
Else
$aArray[$iCopyTo_Index][$j] = $vValue[$iInsertPoint_Index - 1][$j - $iStart]
EndIf
Else
$aArray[$iCopyTo_Index][$j] = ""
EndIf
EndIf
Next
$iCopyTo_Index -= 1
$iInsertPoint_Index -= 1
If $iInsertPoint_Index = 0 Then ExitLoop 2
$iInsert_Index = $vRange[$iInsertPoint_Index]
WEnd
Next
Case Else
Return SetError(2, 0, -1)
EndSwitch
Return UBound($aArray, $UBOUND_ROWS)
EndFunc
Func _ArrayMax(Const ByRef $aArray, $iCompNumeric = 0, $iStart = -1, $iEnd = -1, $iSubItem = 0)
Local $iResult = _ArrayMaxIndex($aArray, $iCompNumeric, $iStart, $iEnd, $iSubItem)
If @error Then Return SetError(@error, 0, "")
If UBound($aArray, $UBOUND_DIMENSIONS) = 1 Then
Return $aArray[$iResult]
Else
Return $aArray[$iResult][$iSubItem]
EndIf
EndFunc
Func _ArrayMaxIndex(Const ByRef $aArray, $iCompNumeric = 0, $iStart = -1, $iEnd = -1, $iSubItem = 0)
If $iCompNumeric = Default Then $iCompNumeric = 0
If $iStart = Default Then $iStart = -1
If $iEnd = Default Then $iEnd = -1
If $iSubItem = Default Then $iSubItem = 0
Local $iRet = __Array_MinMaxIndex($aArray, $iCompNumeric, $iStart, $iEnd, $iSubItem, __Array_GreaterThan)
Return SetError(@error, 0, $iRet)
EndFunc
Func _ArrayMin(Const ByRef $aArray, $iCompNumeric = 0, $iStart = -1, $iEnd = -1, $iSubItem = 0)
Local $iResult = _ArrayMinIndex($aArray, $iCompNumeric, $iStart, $iEnd, $iSubItem)
If @error Then Return SetError(@error, 0, "")
If UBound($aArray, $UBOUND_DIMENSIONS) = 1 Then
Return $aArray[$iResult]
Else
Return $aArray[$iResult][$iSubItem]
EndIf
EndFunc
Func _ArrayMinIndex(Const ByRef $aArray, $iCompNumeric = 0, $iStart = -1, $iEnd = -1, $iSubItem = 0)
If $iCompNumeric = Default Then $iCompNumeric = 0
If $iStart = Default Then $iStart = -1
If $iEnd = Default Then $iEnd = -1
If $iSubItem = Default Then $iSubItem = 0
Local $iRet = __Array_MinMaxIndex($aArray, $iCompNumeric, $iStart, $iEnd, $iSubItem, __Array_LessThan)
Return SetError(@error, 0, $iRet)
EndFunc
Func _ArrayPermute(ByRef $aArray, $sDelimiter = "")
If $sDelimiter = Default Then $sDelimiter = ""
If Not IsArray($aArray) Then Return SetError(1, 0, 0)
If UBound($aArray, $UBOUND_DIMENSIONS) <> 1 Then Return SetError(2, 0, 0)
Local $iSize = UBound($aArray), $iFactorial = 1, $aIdx[$iSize], $aResult[1], $iCount = 1
If UBound($aArray) Then
For $i = 0 To $iSize - 1
$aIdx[$i] = $i
Next
For $i = $iSize To 1 Step -1
$iFactorial *= $i
Next
ReDim $aResult[$iFactorial + 1]
$aResult[0] = $iFactorial
__Array_ExeterInternal($aArray, 0, $iSize, $sDelimiter, $aIdx, $aResult, $iCount)
Else
$aResult[0] = 0
EndIf
Return $aResult
EndFunc
Func _ArrayPop(ByRef $aArray)
If (Not IsArray($aArray)) Then Return SetError(1, 0, "")
If UBound($aArray, $UBOUND_DIMENSIONS) <> 1 Then Return SetError(2, 0, "")
Local $iUBound = UBound($aArray) - 1
If $iUBound = -1 Then Return SetError(3, 0, "")
Local $sLastVal = $aArray[$iUBound]
If $iUBound > -1 Then
ReDim $aArray[$iUBound]
EndIf
Return $sLastVal
EndFunc
Func _ArrayPush(ByRef $aArray, $vValue, $iDirection = 0)
If $iDirection = Default Then $iDirection = 0
If (Not IsArray($aArray)) Then Return SetError(1, 0, 0)
If UBound($aArray, $UBOUND_DIMENSIONS) <> 1 Then Return SetError(3, 0, 0)
Local $iUBound = UBound($aArray) - 1
If IsArray($vValue) Then
Local $iUBoundS = UBound($vValue)
If ($iUBoundS - 1) > $iUBound Then Return SetError(2, 0, 0)
If $iDirection Then
For $i = $iUBound To $iUBoundS Step -1
$aArray[$i] = $aArray[$i - $iUBoundS]
Next
For $i = 0 To $iUBoundS - 1
$aArray[$i] = $vValue[$i]
Next
Else
For $i = 0 To $iUBound - $iUBoundS
$aArray[$i] = $aArray[$i + $iUBoundS]
Next
For $i = 0 To $iUBoundS - 1
$aArray[$i + $iUBound - $iUBoundS + 1] = $vValue[$i]
Next
EndIf
Else
If $iUBound > -1 Then
If $iDirection Then
For $i = $iUBound To 1 Step -1
$aArray[$i] = $aArray[$i - 1]
Next
$aArray[0] = $vValue
Else
For $i = 0 To $iUBound - 1
$aArray[$i] = $aArray[$i + 1]
Next
$aArray[$iUBound] = $vValue
EndIf
EndIf
EndIf
Return 1
EndFunc
Func _ArrayReverse(ByRef $aArray, $iStart = 0, $iEnd = 0)
If $iStart = Default Then $iStart = 0
If $iEnd = Default Then $iEnd = 0
If Not IsArray($aArray) Then Return SetError(1, 0, 0)
If UBound($aArray, $UBOUND_DIMENSIONS) <> 1 Then Return SetError(3, 0, 0)
If Not UBound($aArray) Then Return SetError(4, 0, 0)
Local $vTmp, $iUBound = UBound($aArray) - 1
If $iEnd < 1 Or $iEnd > $iUBound Then $iEnd = $iUBound
If $iStart < 0 Then $iStart = 0
If $iStart > $iEnd Then Return SetError(2, 0, 0)
For $i = $iStart To Int(($iStart + $iEnd - 1) / 2)
$vTmp = $aArray[$i]
$aArray[$i] = $aArray[$iEnd]
$aArray[$iEnd] = $vTmp
$iEnd -= 1
Next
Return 1
EndFunc
Func _ArraySearch(Const ByRef $aArray, $vValue, $iStart = 0, $iEnd = 0, $iCase = 0, $iCompare = 0, $iForward = 1, $iSubItem = -1, $bRow = False)
If $iStart = Default Then $iStart = 0
If $iEnd = Default Then $iEnd = 0
If $iCase = Default Then $iCase = 0
If $iCompare = Default Then $iCompare = 0
If $iForward = Default Then $iForward = 1
If $iSubItem = Default Then $iSubItem = -1
If $bRow = Default Then $bRow = False
If Not IsArray($aArray) Then Return SetError(1, 0, -1)
Local $iDim_1 = UBound($aArray) - 1
If $iDim_1 = -1 Then Return SetError(3, 0, -1)
Local $iDim_2 = UBound($aArray, $UBOUND_COLUMNS) - 1
Local $bCompType = False
If $iCompare = 2 Then
$iCompare = 0
$bCompType = True
EndIf
If $bRow Then
If UBound($aArray, $UBOUND_DIMENSIONS) = 1 Then Return SetError(5, 0, -1)
If $iEnd < 1 Or $iEnd > $iDim_2 Then $iEnd = $iDim_2
If $iStart < 0 Then $iStart = 0
If $iStart > $iEnd Then Return SetError(4, 0, -1)
Else
If $iEnd < 1 Or $iEnd > $iDim_1 Then $iEnd = $iDim_1
If $iStart < 0 Then $iStart = 0
If $iStart > $iEnd Then Return SetError(4, 0, -1)
EndIf
Local $iStep = 1
If Not $iForward Then
Local $iTmp = $iStart
$iStart = $iEnd
$iEnd = $iTmp
$iStep = -1
EndIf
Switch UBound($aArray, $UBOUND_DIMENSIONS)
Case 1
If Not $iCompare Then
If Not $iCase Then
For $i = $iStart To $iEnd Step $iStep
If $bCompType And VarGetType($aArray[$i]) <> VarGetType($vValue) Then ContinueLoop
If $aArray[$i] = $vValue Then Return $i
Next
Else
For $i = $iStart To $iEnd Step $iStep
If $bCompType And VarGetType($aArray[$i]) <> VarGetType($vValue) Then ContinueLoop
If $aArray[$i] == $vValue Then Return $i
Next
EndIf
Else
For $i = $iStart To $iEnd Step $iStep
If $iCompare = 3 Then
If StringRegExp($aArray[$i], $vValue) Then Return $i
Else
If StringInStr($aArray[$i], $vValue, $iCase) > 0 Then Return $i
EndIf
Next
EndIf
Case 2
Local $iDim_Sub
If $bRow Then
$iDim_Sub = $iDim_1
If $iSubItem > $iDim_Sub Then $iSubItem = $iDim_Sub
If $iSubItem < 0 Then
$iSubItem = 0
Else
$iDim_Sub = $iSubItem
EndIf
Else
$iDim_Sub = $iDim_2
If $iSubItem > $iDim_Sub Then $iSubItem = $iDim_Sub
If $iSubItem < 0 Then
$iSubItem = 0
Else
$iDim_Sub = $iSubItem
EndIf
EndIf
For $j = $iSubItem To $iDim_Sub
If Not $iCompare Then
If Not $iCase Then
For $i = $iStart To $iEnd Step $iStep
If $bRow Then
If $bCompType And VarGetType($aArray[$j][$i]) <> VarGetType($vValue) Then ContinueLoop
If $aArray[$j][$i] = $vValue Then Return $i
Else
If $bCompType And VarGetType($aArray[$i][$j]) <> VarGetType($vValue) Then ContinueLoop
If $aArray[$i][$j] = $vValue Then Return $i
EndIf
Next
Else
For $i = $iStart To $iEnd Step $iStep
If $bRow Then
If $bCompType And VarGetType($aArray[$j][$i]) <> VarGetType($vValue) Then ContinueLoop
If $aArray[$j][$i] == $vValue Then Return $i
Else
If $bCompType And VarGetType($aArray[$i][$j]) <> VarGetType($vValue) Then ContinueLoop
If $aArray[$i][$j] == $vValue Then Return $i
EndIf
Next
EndIf
Else
For $i = $iStart To $iEnd Step $iStep
If $iCompare = 3 Then
If $bRow Then
If StringRegExp($aArray[$j][$i], $vValue) Then Return $i
Else
If StringRegExp($aArray[$i][$j], $vValue) Then Return $i
EndIf
Else
If $bRow Then
If StringInStr($aArray[$j][$i], $vValue, $iCase) > 0 Then Return $i
Else
If StringInStr($aArray[$i][$j], $vValue, $iCase) > 0 Then Return $i
EndIf
EndIf
Next
EndIf
Next
Case Else
Return SetError(2, 0, -1)
EndSwitch
Return SetError(6, 0, -1)
EndFunc
Func _ArrayShuffle(ByRef $aArray, $iStart_Row = 0, $iEnd_Row = 0, $iCol = -1)
If $iStart_Row = Default Then $iStart_Row = 0
If $iEnd_Row = Default Then $iEnd_Row = 0
If $iCol = Default Then $iCol = -1
If Not IsArray($aArray) Then Return SetError(1, 0, -1)
Local $iDim_1 = UBound($aArray, $UBOUND_ROWS)
If $iEnd_Row = 0 Then $iEnd_Row = $iDim_1 - 1
If $iStart_Row < 0 Or $iStart_Row > $iDim_1 - 1 Then Return SetError(3, 0, -1)
If $iEnd_Row < 1 Or $iEnd_Row > $iDim_1 - 1 Then Return SetError(3, 0, -1)
If $iStart_Row > $iEnd_Row Then Return SetError(4, 0, -1)
Local $vTmp, $iRand
Switch UBound($aArray, $UBOUND_DIMENSIONS)
Case 1
For $i = $iEnd_Row To $iStart_Row + 1 Step -1
$iRand = Random($iStart_Row, $i, 1)
$vTmp = $aArray[$i]
$aArray[$i] = $aArray[$iRand]
$aArray[$iRand] = $vTmp
Next
Return 1
Case 2
Local $iDim_2 = UBound($aArray, $UBOUND_COLUMNS)
If $iCol < -1 Or $iCol > $iDim_2 - 1 Then Return SetError(5, 0, -1)
Local $iCol_Start, $iCol_End
If $iCol = -1 Then
$iCol_Start = 0
$iCol_End = $iDim_2 - 1
Else
$iCol_Start = $iCol
$iCol_End = $iCol
EndIf
For $i = $iEnd_Row To $iStart_Row + 1 Step -1
$iRand = Random($iStart_Row, $i, 1)
For $j = $iCol_Start To $iCol_End
$vTmp = $aArray[$i][$j]
$aArray[$i][$j] = $aArray[$iRand][$j]
$aArray[$iRand][$j] = $vTmp
Next
Next
Return 1
Case Else
Return SetError(2, 0, -1)
EndSwitch
EndFunc
Func _ArraySort(ByRef $aArray, $iDescending = 0, $iStart = 0, $iEnd = 0, $iSubItem = 0, $iPivot = 0)
If $iDescending = Default Then $iDescending = 0
If $iStart = Default Then $iStart = 0
If $iEnd = Default Then $iEnd = 0
If $iSubItem = Default Then $iSubItem = 0
If $iPivot = Default Then $iPivot = 0
If Not IsArray($aArray) Then Return SetError(1, 0, 0)
Local $iUBound = UBound($aArray) - 1
If $iUBound = -1 Then Return SetError(5, 0, 0)
If $iEnd = Default Then $iEnd = 0
If $iEnd < 1 Or $iEnd > $iUBound Or $iEnd = Default Then $iEnd = $iUBound
If $iStart < 0 Or $iStart = Default Then $iStart = 0
If $iStart > $iEnd Then Return SetError(2, 0, 0)
Switch UBound($aArray, $UBOUND_DIMENSIONS)
Case 1
If $iPivot Then
__ArrayDualPivotSort($aArray, $iStart, $iEnd)
Else
__ArrayQuickSort1D($aArray, $iStart, $iEnd)
EndIf
If $iDescending Then _ArrayReverse($aArray, $iStart, $iEnd)
Case 2
If $iPivot Then Return SetError(6, 0, 0)
Local $iSubMax = UBound($aArray, $UBOUND_COLUMNS) - 1
If $iSubItem > $iSubMax Then Return SetError(3, 0, 0)
If $iDescending Then
$iDescending = -1
Else
$iDescending = 1
EndIf
__ArrayQuickSort2D($aArray, $iDescending, $iStart, $iEnd, $iSubItem, $iSubMax)
Case Else
Return SetError(4, 0, 0)
EndSwitch
Return 1
EndFunc
Func __ArrayQuickSort1D(ByRef $aArray, Const ByRef $iStart, Const ByRef $iEnd)
If $iEnd <= $iStart Then Return
Local $vTmp
If ($iEnd - $iStart) < 15 Then
Local $vCur
For $i = $iStart + 1 To $iEnd
$vTmp = $aArray[$i]
If IsNumber($vTmp) Then
For $j = $i - 1 To $iStart Step -1
$vCur = $aArray[$j]
If ($vTmp >= $vCur And IsNumber($vCur)) Or (Not IsNumber($vCur) And StringCompare($vTmp, $vCur) >= 0) Then ExitLoop
$aArray[$j + 1] = $vCur
Next
Else
For $j = $i - 1 To $iStart Step -1
If (StringCompare($vTmp, $aArray[$j]) >= 0) Then ExitLoop
$aArray[$j + 1] = $aArray[$j]
Next
EndIf
$aArray[$j + 1] = $vTmp
Next
Return
EndIf
Local $L = $iStart, $R = $iEnd, $vPivot = $aArray[Int(($iStart + $iEnd) / 2)], $bNum = IsNumber($vPivot)
Do
If $bNum Then
While ($aArray[$L] < $vPivot And IsNumber($aArray[$L])) Or (Not IsNumber($aArray[$L]) And StringCompare($aArray[$L], $vPivot) < 0)
$L += 1
WEnd
While ($aArray[$R] > $vPivot And IsNumber($aArray[$R])) Or (Not IsNumber($aArray[$R]) And StringCompare($aArray[$R], $vPivot) > 0)
$R -= 1
WEnd
Else
While (StringCompare($aArray[$L], $vPivot) < 0)
$L += 1
WEnd
While (StringCompare($aArray[$R], $vPivot) > 0)
$R -= 1
WEnd
EndIf
If $L <= $R Then
$vTmp = $aArray[$L]
$aArray[$L] = $aArray[$R]
$aArray[$R] = $vTmp
$L += 1
$R -= 1
EndIf
Until $L > $R
__ArrayQuickSort1D($aArray, $iStart, $R)
__ArrayQuickSort1D($aArray, $L, $iEnd)
EndFunc
Func __ArrayQuickSort2D(ByRef $aArray, Const ByRef $iStep, Const ByRef $iStart, Const ByRef $iEnd, Const ByRef $iSubItem, Const ByRef $iSubMax)
If $iEnd <= $iStart Then Return
Local $vTmp, $L = $iStart, $R = $iEnd, $vPivot = $aArray[Int(($iStart + $iEnd) / 2)][$iSubItem], $bNum = IsNumber($vPivot)
Do
If $bNum Then
While ($iStep * ($aArray[$L][$iSubItem] - $vPivot) < 0 And IsNumber($aArray[$L][$iSubItem])) Or (Not IsNumber($aArray[$L][$iSubItem]) And $iStep * StringCompare($aArray[$L][$iSubItem], $vPivot) < 0)
$L += 1
WEnd
While ($iStep * ($aArray[$R][$iSubItem] - $vPivot) > 0 And IsNumber($aArray[$R][$iSubItem])) Or (Not IsNumber($aArray[$R][$iSubItem]) And $iStep * StringCompare($aArray[$R][$iSubItem], $vPivot) > 0)
$R -= 1
WEnd
Else
While ($iStep * StringCompare($aArray[$L][$iSubItem], $vPivot) < 0)
$L += 1
WEnd
While ($iStep * StringCompare($aArray[$R][$iSubItem], $vPivot) > 0)
$R -= 1
WEnd
EndIf
If $L <= $R Then
For $i = 0 To $iSubMax
$vTmp = $aArray[$L][$i]
$aArray[$L][$i] = $aArray[$R][$i]
$aArray[$R][$i] = $vTmp
Next
$L += 1
$R -= 1
EndIf
Until $L > $R
__ArrayQuickSort2D($aArray, $iStep, $iStart, $R, $iSubItem, $iSubMax)
__ArrayQuickSort2D($aArray, $iStep, $L, $iEnd, $iSubItem, $iSubMax)
EndFunc
Func __ArrayDualPivotSort(ByRef $aArray, $iPivot_Left, $iPivot_Right, $bLeftMost = True)
If $iPivot_Left > $iPivot_Right Then Return
Local $iLength = $iPivot_Right - $iPivot_Left + 1
Local $i, $j, $k, $iAi, $iAk, $iA1, $iA2, $iLast
If $iLength < 45 Then
If $bLeftMost Then
$i = $iPivot_Left
While $i < $iPivot_Right
$j = $i
$iAi = $aArray[$i + 1]
While $iAi < $aArray[$j]
$aArray[$j + 1] = $aArray[$j]
$j -= 1
If $j + 1 = $iPivot_Left Then ExitLoop
WEnd
$aArray[$j + 1] = $iAi
$i += 1
WEnd
Else
While 1
If $iPivot_Left >= $iPivot_Right Then Return 1
$iPivot_Left += 1
If $aArray[$iPivot_Left] < $aArray[$iPivot_Left - 1] Then ExitLoop
WEnd
While 1
$k = $iPivot_Left
$iPivot_Left += 1
If $iPivot_Left > $iPivot_Right Then ExitLoop
$iA1 = $aArray[$k]
$iA2 = $aArray[$iPivot_Left]
If $iA1 < $iA2 Then
$iA2 = $iA1
$iA1 = $aArray[$iPivot_Left]
EndIf
$k -= 1
While $iA1 < $aArray[$k]
$aArray[$k + 2] = $aArray[$k]
$k -= 1
WEnd
$aArray[$k + 2] = $iA1
While $iA2 < $aArray[$k]
$aArray[$k + 1] = $aArray[$k]
$k -= 1
WEnd
$aArray[$k + 1] = $iA2
$iPivot_Left += 1
WEnd
$iLast = $aArray[$iPivot_Right]
$iPivot_Right -= 1
While $iLast < $aArray[$iPivot_Right]
$aArray[$iPivot_Right + 1] = $aArray[$iPivot_Right]
$iPivot_Right -= 1
WEnd
$aArray[$iPivot_Right + 1] = $iLast
EndIf
Return 1
EndIf
Local $iSeventh = BitShift($iLength, 3) + BitShift($iLength, 6) + 1
Local $iE1, $iE2, $iE3, $iE4, $iE5, $t
$iE3 = Ceiling(($iPivot_Left + $iPivot_Right) / 2)
$iE2 = $iE3 - $iSeventh
$iE1 = $iE2 - $iSeventh
$iE4 = $iE3 + $iSeventh
$iE5 = $iE4 + $iSeventh
If $aArray[$iE2] < $aArray[$iE1] Then
$t = $aArray[$iE2]
$aArray[$iE2] = $aArray[$iE1]
$aArray[$iE1] = $t
EndIf
If $aArray[$iE3] < $aArray[$iE2] Then
$t = $aArray[$iE3]
$aArray[$iE3] = $aArray[$iE2]
$aArray[$iE2] = $t
If $t < $aArray[$iE1] Then
$aArray[$iE2] = $aArray[$iE1]
$aArray[$iE1] = $t
EndIf
EndIf
If $aArray[$iE4] < $aArray[$iE3] Then
$t = $aArray[$iE4]
$aArray[$iE4] = $aArray[$iE3]
$aArray[$iE3] = $t
If $t < $aArray[$iE2] Then
$aArray[$iE3] = $aArray[$iE2]
$aArray[$iE2] = $t
If $t < $aArray[$iE1] Then
$aArray[$iE2] = $aArray[$iE1]
$aArray[$iE1] = $t
EndIf
EndIf
EndIf
If $aArray[$iE5] < $aArray[$iE4] Then
$t = $aArray[$iE5]
$aArray[$iE5] = $aArray[$iE4]
$aArray[$iE4] = $t
If $t < $aArray[$iE3] Then
$aArray[$iE4] = $aArray[$iE3]
$aArray[$iE3] = $t
If $t < $aArray[$iE2] Then
$aArray[$iE3] = $aArray[$iE2]
$aArray[$iE2] = $t
If $t < $aArray[$iE1] Then
$aArray[$iE2] = $aArray[$iE1]
$aArray[$iE1] = $t
EndIf
EndIf
EndIf
EndIf
Local $iLess = $iPivot_Left
Local $iGreater = $iPivot_Right
If (($aArray[$iE1] <> $aArray[$iE2]) And ($aArray[$iE2] <> $aArray[$iE3]) And ($aArray[$iE3] <> $aArray[$iE4]) And ($aArray[$iE4] <> $aArray[$iE5])) Then
Local $iPivot_1 = $aArray[$iE2]
Local $iPivot_2 = $aArray[$iE4]
$aArray[$iE2] = $aArray[$iPivot_Left]
$aArray[$iE4] = $aArray[$iPivot_Right]
Do
$iLess += 1
Until $aArray[$iLess] >= $iPivot_1
Do
$iGreater -= 1
Until $aArray[$iGreater] <= $iPivot_2
$k = $iLess
While $k <= $iGreater
$iAk = $aArray[$k]
If $iAk < $iPivot_1 Then
$aArray[$k] = $aArray[$iLess]
$aArray[$iLess] = $iAk
$iLess += 1
ElseIf $iAk > $iPivot_2 Then
While $aArray[$iGreater] > $iPivot_2
$iGreater -= 1
If $iGreater + 1 = $k Then ExitLoop 2
WEnd
If $aArray[$iGreater] < $iPivot_1 Then
$aArray[$k] = $aArray[$iLess]
$aArray[$iLess] = $aArray[$iGreater]
$iLess += 1
Else
$aArray[$k] = $aArray[$iGreater]
EndIf
$aArray[$iGreater] = $iAk
$iGreater -= 1
EndIf
$k += 1
WEnd
$aArray[$iPivot_Left] = $aArray[$iLess - 1]
$aArray[$iLess - 1] = $iPivot_1
$aArray[$iPivot_Right] = $aArray[$iGreater + 1]
$aArray[$iGreater + 1] = $iPivot_2
__ArrayDualPivotSort($aArray, $iPivot_Left, $iLess - 2, True)
__ArrayDualPivotSort($aArray, $iGreater + 2, $iPivot_Right, False)
If ($iLess < $iE1) And ($iE5 < $iGreater) Then
While $aArray[$iLess] = $iPivot_1
$iLess += 1
WEnd
While $aArray[$iGreater] = $iPivot_2
$iGreater -= 1
WEnd
$k = $iLess
While $k <= $iGreater
$iAk = $aArray[$k]
If $iAk = $iPivot_1 Then
$aArray[$k] = $aArray[$iLess]
$aArray[$iLess] = $iAk
$iLess += 1
ElseIf $iAk = $iPivot_2 Then
While $aArray[$iGreater] = $iPivot_2
$iGreater -= 1
If $iGreater + 1 = $k Then ExitLoop 2
WEnd
If $aArray[$iGreater] = $iPivot_1 Then
$aArray[$k] = $aArray[$iLess]
$aArray[$iLess] = $iPivot_1
$iLess += 1
Else
$aArray[$k] = $aArray[$iGreater]
EndIf
$aArray[$iGreater] = $iAk
$iGreater -= 1
EndIf
$k += 1
WEnd
EndIf
__ArrayDualPivotSort($aArray, $iLess, $iGreater, False)
Else
Local $iPivot = $aArray[$iE3]
$k = $iLess
While $k <= $iGreater
If $aArray[$k] = $iPivot Then
$k += 1
ContinueLoop
EndIf
$iAk = $aArray[$k]
If $iAk < $iPivot Then
$aArray[$k] = $aArray[$iLess]
$aArray[$iLess] = $iAk
$iLess += 1
Else
While $aArray[$iGreater] > $iPivot
$iGreater -= 1
WEnd
If $aArray[$iGreater] < $iPivot Then
$aArray[$k] = $aArray[$iLess]
$aArray[$iLess] = $aArray[$iGreater]
$iLess += 1
Else
$aArray[$k] = $iPivot
EndIf
$aArray[$iGreater] = $iAk
$iGreater -= 1
EndIf
$k += 1
WEnd
__ArrayDualPivotSort($aArray, $iPivot_Left, $iLess - 1, True)
__ArrayDualPivotSort($aArray, $iGreater + 1, $iPivot_Right, False)
EndIf
EndFunc
Func _ArraySwap(ByRef $aArray, $iIndex_1, $iIndex_2, $bCol = False, $iStart = -1, $iEnd = -1)
If $bCol = Default Then $bCol = False
If $iStart = Default Then $iStart = -1
If $iEnd = Default Then $iEnd = -1
If Not IsArray($aArray) Then Return SetError(1, 0, -1)
Local $iDim_1 = UBound($aArray, $UBOUND_ROWS) - 1
Local $iDim_2 = UBound($aArray, $UBOUND_COLUMNS) - 1
If $iDim_2 = -1 Then
$bCol = False
$iStart = -1
$iEnd = -1
EndIf
If $iStart > $iEnd Then Return SetError(5, 0, -1)
If $bCol Then
If $iIndex_1 < 0 Or $iIndex_2 > $iDim_2 Then Return SetError(3, 0, -1)
If $iStart = -1 Then $iStart = 0
If $iEnd = -1 Then $iEnd = $iDim_1
Else
If $iIndex_1 < 0 Or $iIndex_2 > $iDim_1 Then Return SetError(3, 0, -1)
If $iStart = -1 Then $iStart = 0
If $iEnd = -1 Then $iEnd = $iDim_2
EndIf
Local $vTmp
Switch UBound($aArray, $UBOUND_DIMENSIONS)
Case 1
$vTmp = $aArray[$iIndex_1]
$aArray[$iIndex_1] = $aArray[$iIndex_2]
$aArray[$iIndex_2] = $vTmp
Case 2
If $iStart < -1 Or $iEnd < -1 Then Return SetError(4, 0, -1)
If $bCol Then
If $iStart > $iDim_1 Or $iEnd > $iDim_1 Then Return SetError(4, 0, -1)
For $j = $iStart To $iEnd
$vTmp = $aArray[$j][$iIndex_1]
$aArray[$j][$iIndex_1] = $aArray[$j][$iIndex_2]
$aArray[$j][$iIndex_2] = $vTmp
Next
Else
If $iStart > $iDim_2 Or $iEnd > $iDim_2 Then Return SetError(4, 0, -1)
For $j = $iStart To $iEnd
$vTmp = $aArray[$iIndex_1][$j]
$aArray[$iIndex_1][$j] = $aArray[$iIndex_2][$j]
$aArray[$iIndex_2][$j] = $vTmp
Next
EndIf
Case Else
Return SetError(2, 0, -1)
EndSwitch
Return 1
EndFunc
Func _ArrayToClip(Const ByRef $aArray, $sDelim_Col = "|", $iStart_Row = -1, $iEnd_Row = -1, $sDelim_Row = @CRLF, $iStart_Col = -1, $iEnd_Col = -1)
Local $sResult = _ArrayToString($aArray, $sDelim_Col, $iStart_Row, $iEnd_Row, $sDelim_Row, $iStart_Col, $iEnd_Col)
If @error Then Return SetError(@error, 0, 0)
If ClipPut($sResult) Then Return 1
Return SetError(-1, 0, 0)
EndFunc
Func _ArrayToString(Const ByRef $aArray, $sDelim_Col = "|", $iStart_Row = -1, $iEnd_Row = -1, $sDelim_Row = @CRLF, $iStart_Col = -1, $iEnd_Col = -1)
If $sDelim_Col = Default Then $sDelim_Col = "|"
If $sDelim_Row = Default Then $sDelim_Row = @CRLF
If $iStart_Row = Default Then $iStart_Row = -1
If $iEnd_Row = Default Then $iEnd_Row = -1
If $iStart_Col = Default Then $iStart_Col = -1
If $iEnd_Col = Default Then $iEnd_Col = -1
If Not IsArray($aArray) Then Return SetError(1, 0, -1)
Local $iDim_1 = UBound($aArray, $UBOUND_ROWS) - 1
If $iStart_Row = -1 Then $iStart_Row = 0
If $iEnd_Row = -1 Then $iEnd_Row = $iDim_1
If $iStart_Row < -1 Or $iEnd_Row < -1 Then Return SetError(3, 0, -1)
If $iStart_Row > $iDim_1 Or $iEnd_Row > $iDim_1 Then Return SetError(3, 0, "")
If $iStart_Row > $iEnd_Row Then Return SetError(4, 0, -1)
Local $sRet = ""
Switch UBound($aArray, $UBOUND_DIMENSIONS)
Case 1
For $i = $iStart_Row To $iEnd_Row
$sRet &= $aArray[$i] & $sDelim_Col
Next
Return StringTrimRight($sRet, StringLen($sDelim_Col))
Case 2
Local $iDim_2 = UBound($aArray, $UBOUND_COLUMNS) - 1
If $iStart_Col = -1 Then $iStart_Col = 0
If $iEnd_Col = -1 Then $iEnd_Col = $iDim_2
If $iStart_Col < -1 Or $iEnd_Col < -1 Then Return SetError(5, 0, -1)
If $iStart_Col > $iDim_2 Or $iEnd_Col > $iDim_2 Then Return SetError(5, 0, -1)
If $iStart_Col > $iEnd_Col Then Return SetError(6, 0, -1)
For $i = $iStart_Row To $iEnd_Row
For $j = $iStart_Col To $iEnd_Col
$sRet &= $aArray[$i][$j] & $sDelim_Col
Next
$sRet = StringTrimRight($sRet, StringLen($sDelim_Col)) & $sDelim_Row
Next
Return StringTrimRight($sRet, StringLen($sDelim_Row))
Case Else
Return SetError(2, 0, -1)
EndSwitch
Return 1
EndFunc
Func _ArrayTranspose(ByRef $aArray)
Switch UBound($aArray, 0)
Case 0
Return SetError(2, 0, 0)
Case 1
Local $aTemp[1][UBound($aArray)]
For $i = 0 To UBound($aArray) - 1
$aTemp[0][$i] = $aArray[$i]
Next
$aArray = $aTemp
Case 2
Local $iDim_1 = UBound($aArray, 1), $iDim_2 = UBound($aArray, 2)
If $iDim_1 <> $iDim_2 Then
Local $aTemp[$iDim_2][$iDim_1]
For $i = 0 To $iDim_1 - 1
For $j = 0 To $iDim_2 - 1
$aTemp[$j][$i] = $aArray[$i][$j]
Next
Next
$aArray = $aTemp
Else
Local $vElement
For $i = 0 To $iDim_1 - 1
For $j = $i + 1 To $iDim_2 - 1
$vElement = $aArray[$i][$j]
$aArray[$i][$j] = $aArray[$j][$i]
$aArray[$j][$i] = $vElement
Next
Next
EndIf
Case Else
Return SetError(1, 0, 0)
EndSwitch
Return 1
EndFunc
Func _ArrayTrim(ByRef $aArray, $iTrimNum, $iDirection = 0, $iStart = 0, $iEnd = 0, $iSubItem = 0)
If $iDirection = Default Then $iDirection = 0
If $iStart = Default Then $iStart = 0
If $iEnd = Default Then $iEnd = 0
If $iSubItem = Default Then $iSubItem = 0
If Not IsArray($aArray) Then Return SetError(1, 0, 0)
Local $iDim_1 = UBound($aArray, $UBOUND_ROWS) - 1
If $iEnd = 0 Then $iEnd = $iDim_1
If $iStart > $iEnd Then Return SetError(3, 0, -1)
If $iStart < 0 Or $iEnd < 0 Then Return SetError(3, 0, -1)
If $iStart > $iDim_1 Or $iEnd > $iDim_1 Then Return SetError(3, 0, -1)
If $iStart > $iEnd Then Return SetError(4, 0, -1)
Switch UBound($aArray, $UBOUND_DIMENSIONS)
Case 1
If $iDirection Then
For $i = $iStart To $iEnd
$aArray[$i] = StringTrimRight($aArray[$i], $iTrimNum)
Next
Else
For $i = $iStart To $iEnd
$aArray[$i] = StringTrimLeft($aArray[$i], $iTrimNum)
Next
EndIf
Case 2
Local $iDim_2 = UBound($aArray, $UBOUND_COLUMNS) - 1
If $iSubItem < 0 Or $iSubItem > $iDim_2 Then Return SetError(5, 0, -1)
If $iDirection Then
For $i = $iStart To $iEnd
$aArray[$i][$iSubItem] = StringTrimRight($aArray[$i][$iSubItem], $iTrimNum)
Next
Else
For $i = $iStart To $iEnd
$aArray[$i][$iSubItem] = StringTrimLeft($aArray[$i][$iSubItem], $iTrimNum)
Next
EndIf
Case Else
Return SetError(2, 0, 0)
EndSwitch
Return 1
EndFunc
Func _ArrayUnique(Const ByRef $aArray, $iColumn = 0, $iBase = 0, $iCase = 0, $iCount = $ARRAYUNIQUE_COUNT, $iIntType = $ARRAYUNIQUE_AUTO)
If $iColumn = Default Then $iColumn = 0
If $iBase = Default Then $iBase = 0
If $iCase = Default Then $iCase = 0
If $iCount = Default Then $iCount = $ARRAYUNIQUE_COUNT
If UBound($aArray, $UBOUND_ROWS) = 0 Then Return SetError(1, 0, 0)
Local $iDims = UBound($aArray, $UBOUND_DIMENSIONS), $iNumColumns = UBound($aArray, $UBOUND_COLUMNS)
If $iDims > 2 Then Return SetError(2, 0, 0)
If $iBase < 0 Or $iBase > 1 Or (Not IsInt($iBase)) Then Return SetError(3, 0, 0)
If $iCase < 0 Or $iCase > 1 Or (Not IsInt($iCase)) Then Return SetError(3, 0, 0)
If $iCount < 0 Or $iCount > 1 Or (Not IsInt($iCount)) Then Return SetError(4, 0, 0)
If $iIntType < 0 Or $iIntType > 4 Or (Not IsInt($iIntType)) Then Return SetError(5, 0, 0)
If $iColumn < 0 Or ($iNumColumns = 0 And $iColumn > 0) Or ($iNumColumns > 0 And $iColumn >= $iNumColumns) Then Return SetError(6, 0, 0)
If $iIntType = $ARRAYUNIQUE_AUTO Then
Local $bInt, $sVarType
If $iDims = 1 Then
$bInt = IsInt($aArray[$iBase])
$sVarType = VarGetType($aArray[$iBase])
Else
$bInt = IsInt($aArray[$iBase][$iColumn])
$sVarType = VarGetType($aArray[$iBase][$iColumn])
EndIf
If $bInt And $sVarType = "Int64" Then
$iIntType = $ARRAYUNIQUE_FORCE64
Else
$iIntType = $ARRAYUNIQUE_FORCE32
EndIf
EndIf
ObjEvent("AutoIt.Error", __ArrayUnique_AutoErrFunc)
Local $oDictionary = ObjCreate("Scripting.Dictionary")
$oDictionary.CompareMode = Number(Not $iCase)
Local $vElem, $sType, $vKey, $bCOMError = False
For $i = $iBase To UBound($aArray) - 1
If $iDims = 1 Then
$vElem = $aArray[$i]
Else
$vElem = $aArray[$i][$iColumn]
EndIf
Switch $iIntType
Case $ARRAYUNIQUE_FORCE32
$oDictionary.Item($vElem)
If @error Then
$bCOMError = True
ExitLoop
EndIf
Case $ARRAYUNIQUE_FORCE64
$sType = VarGetType($vElem)
If $sType = "Int32" Then
$bCOMError = True
ExitLoop
EndIf
$vKey = "#" & $sType & "#" & String($vElem)
If Not $oDictionary.Item($vKey) Then
$oDictionary($vKey) = $vElem
EndIf
Case $ARRAYUNIQUE_MATCH
$sType = VarGetType($vElem)
If StringLeft($sType, 3) = "Int" Then
$vKey = "#Int#" & String($vElem)
Else
$vKey = "#" & $sType & "#" & String($vElem)
EndIf
If Not $oDictionary.Item($vKey) Then
$oDictionary($vKey) = $vElem
EndIf
Case $ARRAYUNIQUE_DISTINCT
$vKey = "#" & VarGetType($vElem) & "#" & String($vElem)
If Not $oDictionary.Item($vKey) Then
$oDictionary($vKey) = $vElem
EndIf
EndSwitch
Next
Local $aValues, $j = 0
If $bCOMError Then
Return SetError(7, 0, 0)
ElseIf $iIntType <> $ARRAYUNIQUE_FORCE32 Then
Local $aValues[$oDictionary.Count]
For $vKey In $oDictionary.Keys()
$aValues[$j] = $oDictionary($vKey)
If StringLeft($vKey, 5) = "#Ptr#" Then
$aValues[$j] = Ptr($aValues[$j])
EndIf
$j += 1
Next
Else
$aValues = $oDictionary.Keys()
EndIf
If $iCount Then
_ArrayInsert($aValues, 0, $oDictionary.Count)
EndIf
Return $aValues
EndFunc
Func _Array1DToHistogram($aArray, $iSizing = 100)
If UBound($aArray, 0) > 1 Then Return SetError(1, 0, "")
$iSizing = $iSizing * 8
Local $t, $n, $iMin = 0, $iMax = 0, $iOffset = 0
For $i = 0 To UBound($aArray) - 1
$t = $aArray[$i]
$t = IsNumber($t) ? Round($t) : 0
If $t < $iMin Then $iMin = $t
If $t > $iMax Then $iMax = $t
Next
Local $iRange = Int(Round(($iMax - $iMin) / 8)) * 8
Local $iSpaceRatio = 4
For $i = 0 To UBound($aArray) - 1
$t = $aArray[$i]
If $t Then
$n = Abs(Round(($iSizing * $t) / $iRange) / 8)
$aArray[$i] = ""
If $t > 0 Then
If $iMin Then
$iOffset = Int(Abs(Round(($iSizing * $iMin) / $iRange) / 8) / 8 * $iSpaceRatio)
$aArray[$i] = __Array_StringRepeat(ChrW(0x20), $iOffset)
EndIf
Else
If $iMin <> $t Then
$iOffset = Int(Abs(Round(($iSizing * ($t - $iMin)) / $iRange) / 8) / 8 * $iSpaceRatio)
$aArray[$i] = __Array_StringRepeat(ChrW(0x20), $iOffset)
EndIf
EndIf
$aArray[$i] &= __Array_StringRepeat(ChrW(0x2588), Int($n / 8))
$n = Mod($n, 8)
If $n > 0 Then $aArray[$i] &= ChrW(0x2588 + 8 - $n)
$aArray[$i] &= ' ' & $t
Else
$aArray[$i] = ""
EndIf
Next
Return $aArray
EndFunc
Func __Array_StringRepeat($sString, $iRepeatCount)
$iRepeatCount = Int($iRepeatCount)
If StringLen($sString) < 1 Or $iRepeatCount <= 0 Then Return SetError(1, 0, "")
Local $sResult = ""
While $iRepeatCount > 1
If BitAND($iRepeatCount, 1) Then $sResult &= $sString
$sString &= $sString
$iRepeatCount = BitShift($iRepeatCount, 1)
WEnd
Return $sString & $sResult
EndFunc
Func __Array_ExeterInternal(ByRef $aArray, $iStart, $iSize, $sDelimiter, ByRef $aIdx, ByRef $aResult, ByRef $iCount)
If $iStart == $iSize - 1 Then
For $i = 0 To $iSize - 1
$aResult[$iCount] &= $aArray[$aIdx[$i]] & $sDelimiter
Next
If $sDelimiter <> "" Then $aResult[$iCount] = StringTrimRight($aResult[$iCount], StringLen($sDelimiter))
$iCount += 1
Else
Local $iTemp
For $i = $iStart To $iSize - 1
$iTemp = $aIdx[$i]
$aIdx[$i] = $aIdx[$iStart]
$aIdx[$iStart] = $iTemp
__Array_ExeterInternal($aArray, $iStart + 1, $iSize, $sDelimiter, $aIdx, $aResult, $iCount)
$aIdx[$iStart] = $aIdx[$i]
$aIdx[$i] = $iTemp
Next
EndIf
EndFunc
Func __Array_Combinations($iN, $iR)
Local $i_Total = 1
For $i = $iR To 1 Step -1
$i_Total *= ($iN / $i)
$iN -= 1
Next
Return Round($i_Total)
EndFunc
Func __Array_GetNext($iN, $iR, ByRef $iLeft, $iTotal, ByRef $aIdx)
If $iLeft == $iTotal Then
$iLeft -= 1
Return
EndIf
Local $i = $iR - 1
While $aIdx[$i] == $iN - $iR + $i
$i -= 1
WEnd
$aIdx[$i] += 1
For $j = $i + 1 To $iR - 1
$aIdx[$j] = $aIdx[$i] + $j - $i
Next
$iLeft -= 1
EndFunc
Func __Array_MinMaxIndex(Const ByRef $aArray, $iCompNumeric, $iStart, $iEnd, $iSubItem, $fuComparison)
If $iCompNumeric = Default Then $iCompNumeric = 0
If $iCompNumeric <> 1 Then $iCompNumeric = 0
If $iStart = Default Then $iStart = 0
If $iEnd = Default Then $iEnd = 0
If $iSubItem = Default Then $iSubItem = 0
If Not IsArray($aArray) Then Return SetError(1, 0, -1)
Local $iDim_1 = UBound($aArray, $UBOUND_ROWS) - 1
If $iDim_1 < 0 Then Return SetError(1, 0, -1)
If $iEnd = -1 Then $iEnd = $iDim_1
If $iStart = -1 Then $iStart = 0
If $iStart < -1 Or $iEnd < -1 Then Return SetError(3, 0, -1)
If $iStart > $iDim_1 Or $iEnd > $iDim_1 Then Return SetError(3, 0, -1)
If $iStart > $iEnd Then Return SetError(4, 0, -1)
If $iDim_1 < 0 Then Return SetError(5, 0, -1)
Local $iMaxMinIndex = $iStart
Switch UBound($aArray, $UBOUND_DIMENSIONS)
Case 1
If $iCompNumeric Then
For $i = $iStart To $iEnd
If $fuComparison(Number($aArray[$i]), Number($aArray[$iMaxMinIndex])) Then $iMaxMinIndex = $i
Next
Else
For $i = $iStart To $iEnd
If $fuComparison($aArray[$i], $aArray[$iMaxMinIndex]) Then $iMaxMinIndex = $i
Next
EndIf
Case 2
If $iSubItem < 0 Or $iSubItem > UBound($aArray, $UBOUND_COLUMNS) - 1 Then Return SetError(6, 0, -1)
If $iCompNumeric Then
For $i = $iStart To $iEnd
If $fuComparison(Number($aArray[$i][$iSubItem]), Number($aArray[$iMaxMinIndex][$iSubItem])) Then $iMaxMinIndex = $i
Next
Else
For $i = $iStart To $iEnd
If $fuComparison($aArray[$i][$iSubItem], $aArray[$iMaxMinIndex][$iSubItem]) Then $iMaxMinIndex = $i
Next
EndIf
Case Else
Return SetError(2, 0, -1)
EndSwitch
Return $iMaxMinIndex
EndFunc
Func __Array_GreaterThan($vValue1, $vValue2)
Return $vValue1 > $vValue2
EndFunc
Func __Array_LessThan($vValue1, $vValue2)
Return $vValue1 < $vValue2
EndFunc
Func __ArrayUnique_AutoErrFunc()
EndFunc
Global Const $FC_NOOVERWRITE = 0
Global Const $FC_OVERWRITE = 1
Global Const $FC_CREATEPATH = 8
Global Const $FT_MODIFIED = 0
Global Const $FT_CREATED = 1
Global Const $FT_ACCESSED = 2
Global Const $FT_ARRAY = 0
Global Const $FT_STRING = 1
Global Const $FSF_CREATEBUTTON = 1
Global Const $FSF_NEWDIALOG = 2
Global Const $FSF_EDITCONTROL = 4
Global Const $FT_NONRECURSIVE = 0
Global Const $FT_RECURSIVE = 1
Global Const $FO_READ = 0
Global Const $FO_APPEND = 1
Global Const $FO_OVERWRITE = 2
Global Const $FO_CREATEPATH = 8
Global Const $FO_BINARY = 16
Global Const $FO_UNICODE = 32
Global Const $FO_UTF16_LE = 32
Global Const $FO_UTF16_BE = 64
Global Const $FO_UTF8 = 128
Global Const $FO_UTF8_NOBOM = 256
Global Const $FO_ANSI = 512
Global Const $FO_UTF16_LE_NOBOM = 1024
Global Const $FO_UTF16_BE_NOBOM = 2048
Global Const $FO_UTF8_FULL = 16384
Global Const $FO_FULLFILE_DETECT = 16384
Global Const $EOF = -1
Global Const $FD_FILEMUSTEXIST = 1
Global Const $FD_PATHMUSTEXIST = 2
Global Const $FD_MULTISELECT = 4
Global Const $FD_PROMPTCREATENEW = 8
Global Const $FD_PROMPTOVERWRITE = 16
Global Const $CREATE_NEW = 1
Global Const $CREATE_ALWAYS = 2
Global Const $OPEN_EXISTING = 3
Global Const $OPEN_ALWAYS = 4
Global Const $TRUNCATE_EXISTING = 5
Global Const $INVALID_SET_FILE_POINTER = -1
Global Const $FILE_BEGIN = 0
Global Const $FILE_CURRENT = 1
Global Const $FILE_END = 2
Global Const $FILE_ATTRIBUTE_READONLY = 0x00000001
Global Const $FILE_ATTRIBUTE_HIDDEN = 0x00000002
Global Const $FILE_ATTRIBUTE_SYSTEM = 0x00000004
Global Const $FILE_ATTRIBUTE_DIRECTORY = 0x00000010
Global Const $FILE_ATTRIBUTE_ARCHIVE = 0x00000020
Global Const $FILE_ATTRIBUTE_DEVICE = 0x00000040
Global Const $FILE_ATTRIBUTE_NORMAL = 0x00000080
Global Const $FILE_ATTRIBUTE_TEMPORARY = 0x00000100
Global Const $FILE_ATTRIBUTE_SPARSE_FILE = 0x00000200
Global Const $FILE_ATTRIBUTE_REPARSE_POINT = 0x00000400
Global Const $FILE_ATTRIBUTE_COMPRESSED = 0x00000800
Global Const $FILE_ATTRIBUTE_OFFLINE = 0x00001000
Global Const $FILE_ATTRIBUTE_NOT_CONTENT_INDEXED = 0x00002000
Global Const $FILE_ATTRIBUTE_ENCRYPTED = 0x00004000
Global Const $FILE_SHARE_READ = 0x00000001
Global Const $FILE_SHARE_WRITE = 0x00000002
Global Const $FILE_SHARE_DELETE = 0x00000004
Global Const $FILE_SHARE_READWRITE = BitOR($FILE_SHARE_READ, $FILE_SHARE_WRITE)
Global Const $FILE_SHARE_ANY = BitOR($FILE_SHARE_READ, $FILE_SHARE_WRITE, $FILE_SHARE_DELETE)
Global Const $GENERIC_ALL = 0x10000000
Global Const $GENERIC_EXECUTE = 0x20000000
Global Const $GENERIC_WRITE = 0x40000000
Global Const $GENERIC_READ = 0x80000000
Global Const $GENERIC_READWRITE = BitOR($GENERIC_READ, $GENERIC_WRITE)
Global Const $FILE_ENCODING_UTF16LE = 32
Global Const $FE_ENTIRE_UTF8 = 1
Global Const $FE_PARTIALFIRST_UTF8 = 2
Global Const $FN_FULLPATH = 0
Global Const $FN_RELATIVEPATH = 1
Global Const $FV_COMMENTS = "Comments"
Global Const $FV_COMPANYNAME = "CompanyName"
Global Const $FV_FILEDESCRIPTION = "FileDescription"
Global Const $FV_FILEVERSION = "FileVersion"
Global Const $FV_INTERNALNAME = "InternalName"
Global Const $FV_LEGALCOPYRIGHT = "LegalCopyright"
Global Const $FV_LEGALTRADEMARKS = "LegalTrademarks"
Global Const $FV_ORIGINALFILENAME = "OriginalFilename"
Global Const $FV_PRODUCTNAME = "ProductName"
Global Const $FV_PRODUCTVERSION = "ProductVersion"
Global Const $FV_PRIVATEBUILD = "PrivateBuild"
Global Const $FV_SPECIALBUILD = "SpecialBuild"
Global Const $FRTA_NOCOUNT = 0
Global Const $FRTA_COUNT = 1
Global Const $FRTA_INTARRAYS = 2
Global Const $FRTA_ENTIRESPLIT = 4
Global Const $FLTA_FILESFOLDERS = 0
Global Const $FLTA_FILES = 1
Global Const $FLTA_FOLDERS = 2
Global Const $FLTAR_FILESFOLDERS = 0
Global Const $FLTAR_FILES = 1
Global Const $FLTAR_FOLDERS = 2
Global Const $FLTAR_NOHIDDEN = 4
Global Const $FLTAR_NOSYSTEM = 8
Global Const $FLTAR_NOLINK = 16
Global Const $FLTAR_NORECUR = 0
Global Const $FLTAR_RECUR = 1
Global Const $FLTAR_NOSORT = 0
Global Const $FLTAR_SORT = 1
Global Const $FLTAR_FASTSORT = 2
Global Const $FLTAR_NOPATH = 0
Global Const $FLTAR_RELPATH = 1
Global Const $FLTAR_FULLPATH = 2
Global Const $PATH_ORIGINAL = 0
Global Const $PATH_DRIVE = 1
Global Const $PATH_DIRECTORY = 2
Global Const $PATH_FILENAME = 3
Global Const $PATH_EXTENSION = 4
Func _FileCountLines($sFilePath)
FileReadToArray($sFilePath)
If @error Then Return SetError(@error, @extended, 0)
Return @extended
EndFunc
Func _FileCreate($sFilePath)
Local $hFileOpen = FileOpen($sFilePath, BitOR($FO_OVERWRITE, $FO_CREATEPATH))
If $hFileOpen = -1 Then Return SetError(1, 0, 0)
Local $iFileWrite = FileWrite($hFileOpen, "")
FileClose($hFileOpen)
If Not $iFileWrite Then Return SetError(2, 0, 0)
Return 1
EndFunc
Func _FileListToArray($sFilePath, $sFilter = "*", $iFlag = $FLTA_FILESFOLDERS, $bReturnPath = False)
Local $sDelimiter = "|", $sFileList = "", $sFileName = "", $sFullPath = ""
$sFilePath = StringRegExpReplace($sFilePath, "[\\/]+$", "") & "\"
If $iFlag = Default Then $iFlag = $FLTA_FILESFOLDERS
If $bReturnPath Then $sFullPath = $sFilePath
If $sFilter = Default Then $sFilter = "*"
If Not FileExists($sFilePath) Then Return SetError(1, 0, 0)
If StringRegExp($sFilter, "[\\/:><\|]|(?s)^\s*$") Then Return SetError(2, 0, 0)
If Not ($iFlag = 0 Or $iFlag = 1 Or $iFlag = 2) Then Return SetError(3, 0, 0)
Local $hSearch = FileFindFirstFile($sFilePath & $sFilter)
If @error Then Return SetError(4, 0, 0)
While 1
$sFileName = FileFindNextFile($hSearch)
If @error Then ExitLoop
If ($iFlag + @extended = 2) Then ContinueLoop
$sFileList &= $sDelimiter & $sFullPath & $sFileName
WEnd
FileClose($hSearch)
If $sFileList = "" Then Return SetError(4, 0, 0)
Return StringSplit(StringTrimLeft($sFileList, 1), $sDelimiter)
EndFunc
Func _FileListToArrayRec($sFilePath, $sMask = "*", $iReturn = $FLTAR_FILESFOLDERS, $iRecur = $FLTAR_NORECUR, $iSort = $FLTAR_NOSORT, $iReturnPath = $FLTAR_RELPATH)
If Not FileExists($sFilePath) Then Return SetError(1, 1, "")
If $sMask = Default Then $sMask = "*"
If $iReturn = Default Then $iReturn = $FLTAR_FILESFOLDERS
If $iRecur = Default Then $iRecur = $FLTAR_NORECUR
If $iSort = Default Then $iSort = $FLTAR_NOSORT
If $iReturnPath = Default Then $iReturnPath = $FLTAR_RELPATH
If $iRecur > 1 Or Not IsInt($iRecur) Then Return SetError(1, 6, "")
Local $bLongPath = False
If StringLeft($sFilePath, 4) == "\\?\" Then
$bLongPath = True
EndIf
Local $sFolderSlash = ""
If StringRight($sFilePath, 1) = "\" Then
$sFolderSlash = "\"
Else
$sFilePath = $sFilePath & "\"
EndIf
Local $asFolderSearchList[100] = [1]
$asFolderSearchList[1] = $sFilePath
Local $iHide_HS = 0,  $sHide_HS = ""
If BitAND($iReturn, $FLTAR_NOHIDDEN) Then
$iHide_HS += 2
$sHide_HS &= "H"
$iReturn -= $FLTAR_NOHIDDEN
EndIf
If BitAND($iReturn, $FLTAR_NOSYSTEM) Then
$iHide_HS += 4
$sHide_HS &= "S"
$iReturn -= $FLTAR_NOSYSTEM
EndIf
Local $iHide_Link = 0
If BitAND($iReturn, $FLTAR_NOLINK) Then
$iHide_Link = 0x400
$iReturn -= $FLTAR_NOLINK
EndIf
Local $iMaxLevel = 0
If $iRecur < 0 Then
StringReplace($sFilePath, "\", "", 0, $STR_NOCASESENSEBASIC)
$iMaxLevel = @extended - $iRecur
EndIf
Local $sExclude_List = "", $sExclude_List_Folder = "", $sInclude_List = "*"
Local $aMaskSplit = StringSplit($sMask, "|")
Switch $aMaskSplit[0]
Case 3
$sExclude_List_Folder = $aMaskSplit[3]
ContinueCase
Case 2
$sExclude_List = $aMaskSplit[2]
ContinueCase
Case 1
$sInclude_List = $aMaskSplit[1]
EndSwitch
Local $sInclude_File_Mask = ".+"
If $sInclude_List <> "*" Then
If Not __FLTAR_ListToMask($sInclude_File_Mask, $sInclude_List) Then Return SetError(1, 2, "")
EndIf
Local $sInclude_Folder_Mask = ".+"
Switch $iReturn
Case 0
Switch $iRecur
Case 0
$sInclude_Folder_Mask = $sInclude_File_Mask
EndSwitch
Case 2
$sInclude_Folder_Mask = $sInclude_File_Mask
EndSwitch
Local $sExclude_File_Mask = ":"
If $sExclude_List <> "" Then
If Not __FLTAR_ListToMask($sExclude_File_Mask, $sExclude_List) Then Return SetError(1, 3, "")
EndIf
Local $sExclude_Folder_Mask = ":"
If $iRecur Then
If $sExclude_List_Folder Then
If Not __FLTAR_ListToMask($sExclude_Folder_Mask, $sExclude_List_Folder) Then Return SetError(1, 4, "")
EndIf
If $iReturn = 2 Then
$sExclude_Folder_Mask = $sExclude_File_Mask
EndIf
Else
$sExclude_Folder_Mask = $sExclude_File_Mask
EndIf
If Not ($iReturn = 0 Or $iReturn = 1 Or $iReturn = 2) Then Return SetError(1, 5, "")
If Not ($iSort = 0 Or $iSort = 1 Or $iSort = 2) Then Return SetError(1, 7, "")
If Not ($iReturnPath = 0 Or $iReturnPath = 1 Or $iReturnPath = 2) Then Return SetError(1, 8, "")
If $iHide_Link Then
Local $tFile_Data = DllStructCreate("struct;align 4;dword FileAttributes;uint64 CreationTime;uint64 LastAccessTime;uint64 LastWriteTime;" &  "dword FileSizeHigh;dword FileSizeLow;dword Reserved0;dword Reserved1;wchar FileName[260];wchar AlternateFileName[14];endstruct")
Local $hDLL = DllOpen('kernel32.dll'), $aDLL_Ret
EndIf
Local $asReturnList[100] = [0]
Local $asFileMatchList = $asReturnList, $asRootFileMatchList = $asReturnList, $asFolderMatchList = $asReturnList
Local $bFolder = False,  $hSearch = 0,  $sCurrentPath = "", $sName = "", $sRetPath = ""
Local $iAttribs = 0,  $sAttribs = ''
Local $asFolderFileSectionList[100][2] = [[0, 0]]
While $asFolderSearchList[0] > 0
$sCurrentPath = $asFolderSearchList[$asFolderSearchList[0]]
$asFolderSearchList[0] -= 1
Switch $iReturnPath
Case 1
$sRetPath = StringReplace($sCurrentPath, $sFilePath, "")
Case 2
If $bLongPath Then
$sRetPath = StringTrimLeft($sCurrentPath, 4)
Else
$sRetPath = $sCurrentPath
EndIf
EndSwitch
If $iHide_Link Then
$aDLL_Ret = DllCall($hDLL, 'handle', 'FindFirstFileW', 'wstr', $sCurrentPath & "*", 'struct*', $tFile_Data)
If @error Or Not $aDLL_Ret[0] Then
ContinueLoop
EndIf
$hSearch = $aDLL_Ret[0]
Else
$hSearch = FileFindFirstFile($sCurrentPath & "*")
If $hSearch = -1 Then
ContinueLoop
EndIf
EndIf
If $iReturn = 0 And $iSort And $iReturnPath Then
__FLTAR_AddToList($asFolderFileSectionList, $sRetPath, $asFileMatchList[0] + 1)
EndIf
$sAttribs = ''
While 1
If $iHide_Link Then
$aDLL_Ret = DllCall($hDLL, 'int', 'FindNextFileW', 'handle', $hSearch, 'struct*', $tFile_Data)
If @error Or Not $aDLL_Ret[0] Then
ExitLoop
EndIf
$sName = DllStructGetData($tFile_Data, "FileName")
If $sName = ".." Then
ContinueLoop
EndIf
$iAttribs = DllStructGetData($tFile_Data, "FileAttributes")
If $iHide_HS And BitAND($iAttribs, $iHide_HS) Then
ContinueLoop
EndIf
If BitAND($iAttribs, $iHide_Link) Then
ContinueLoop
EndIf
$bFolder = False
If BitAND($iAttribs, 16) Then
$bFolder = True
EndIf
Else
$bFolder = False
$sName = FileFindNextFile($hSearch, 1)
If @error Then
ExitLoop
EndIf
$sAttribs = @extended
If StringInStr($sAttribs, "D") Then
$bFolder = True
EndIf
If StringRegExp($sAttribs, "[" & $sHide_HS & "]") Then
ContinueLoop
EndIf
EndIf
If $bFolder Then
Select
Case $iRecur < 0
StringReplace($sCurrentPath, "\", "", 0, $STR_NOCASESENSEBASIC)
If @extended < $iMaxLevel Then
ContinueCase
EndIf
Case $iRecur = 1
If Not StringRegExp($sName, $sExclude_Folder_Mask) Then
__FLTAR_AddToList($asFolderSearchList, $sCurrentPath & $sName & "\")
EndIf
EndSelect
EndIf
If $iSort Then
If $bFolder Then
If StringRegExp($sName, $sInclude_Folder_Mask) And Not StringRegExp($sName, $sExclude_Folder_Mask) Then
__FLTAR_AddToList($asFolderMatchList, $sRetPath & $sName & $sFolderSlash)
EndIf
Else
If StringRegExp($sName, $sInclude_File_Mask) And Not StringRegExp($sName, $sExclude_File_Mask) Then
If $sCurrentPath = $sFilePath Then
__FLTAR_AddToList($asRootFileMatchList, $sRetPath & $sName)
Else
__FLTAR_AddToList($asFileMatchList, $sRetPath & $sName)
EndIf
EndIf
EndIf
Else
If $bFolder Then
If $iReturn <> 1 And StringRegExp($sName, $sInclude_Folder_Mask) And Not StringRegExp($sName, $sExclude_Folder_Mask) Then
__FLTAR_AddToList($asReturnList, $sRetPath & $sName & $sFolderSlash)
EndIf
Else
If $iReturn <> 2 And StringRegExp($sName, $sInclude_File_Mask) And Not StringRegExp($sName, $sExclude_File_Mask) Then
__FLTAR_AddToList($asReturnList, $sRetPath & $sName)
EndIf
EndIf
EndIf
WEnd
If $iHide_Link Then
DllCall($hDLL, 'int', 'FindClose', 'ptr', $hSearch)
Else
FileClose($hSearch)
EndIf
WEnd
If $iHide_Link Then
DllClose($hDLL)
EndIf
If $iSort Then
Switch $iReturn
Case 2
If $asFolderMatchList[0] = 0 Then Return SetError(1, 9, "")
ReDim $asFolderMatchList[$asFolderMatchList[0] + 1]
$asReturnList = $asFolderMatchList
__ArrayDualPivotSort($asReturnList, 1, $asReturnList[0])
Case 1
If $asRootFileMatchList[0] = 0 And $asFileMatchList[0] = 0 Then Return SetError(1, 9, "")
If $iReturnPath = 0 Then
__FLTAR_AddFileLists($asReturnList, $asRootFileMatchList, $asFileMatchList)
__ArrayDualPivotSort($asReturnList, 1, $asReturnList[0])
Else
__FLTAR_AddFileLists($asReturnList, $asRootFileMatchList, $asFileMatchList, 1)
EndIf
Case 0
If $asRootFileMatchList[0] = 0 And $asFolderMatchList[0] = 0 Then Return SetError(1, 9, "")
If $iReturnPath = 0 Then
__FLTAR_AddFileLists($asReturnList, $asRootFileMatchList, $asFileMatchList)
$asReturnList[0] += $asFolderMatchList[0]
ReDim $asFolderMatchList[$asFolderMatchList[0] + 1]
_ArrayConcatenate($asReturnList, $asFolderMatchList, 1)
__ArrayDualPivotSort($asReturnList, 1, $asReturnList[0])
Else
Local $asReturnList[$asFileMatchList[0] + $asRootFileMatchList[0] + $asFolderMatchList[0] + 1]
$asReturnList[0] = $asFileMatchList[0] + $asRootFileMatchList[0] + $asFolderMatchList[0]
__ArrayDualPivotSort($asRootFileMatchList, 1, $asRootFileMatchList[0])
For $i = 1 To $asRootFileMatchList[0]
$asReturnList[$i] = $asRootFileMatchList[$i]
Next
Local $iNextInsertionIndex = $asRootFileMatchList[0] + 1
__ArrayDualPivotSort($asFolderMatchList, 1, $asFolderMatchList[0])
Local $sFolderToFind = ""
For $i = 1 To $asFolderMatchList[0]
$asReturnList[$iNextInsertionIndex] = $asFolderMatchList[$i]
$iNextInsertionIndex += 1
If $sFolderSlash Then
$sFolderToFind = $asFolderMatchList[$i]
Else
$sFolderToFind = $asFolderMatchList[$i] & "\"
EndIf
Local $iFileSectionEndIndex = 0, $iFileSectionStartIndex = 0
For $j = 1 To $asFolderFileSectionList[0][0]
If $sFolderToFind = $asFolderFileSectionList[$j][0] Then
$iFileSectionStartIndex = $asFolderFileSectionList[$j][1]
If $j = $asFolderFileSectionList[0][0] Then
$iFileSectionEndIndex = $asFileMatchList[0]
Else
$iFileSectionEndIndex = $asFolderFileSectionList[$j + 1][1] - 1
EndIf
If $iSort = 1 Then
__ArrayDualPivotSort($asFileMatchList, $iFileSectionStartIndex, $iFileSectionEndIndex)
EndIf
For $k = $iFileSectionStartIndex To $iFileSectionEndIndex
$asReturnList[$iNextInsertionIndex] = $asFileMatchList[$k]
$iNextInsertionIndex += 1
Next
ExitLoop
EndIf
Next
Next
EndIf
EndSwitch
Else
If $asReturnList[0] = 0 Then Return SetError(1, 9, "")
ReDim $asReturnList[$asReturnList[0] + 1]
EndIf
Return $asReturnList
EndFunc
Func __FLTAR_AddFileLists(ByRef $asTarget, $asSource_1, $asSource_2, $iSort = 0)
ReDim $asSource_1[$asSource_1[0] + 1]
If $iSort = 1 Then __ArrayDualPivotSort($asSource_1, 1, $asSource_1[0])
$asTarget = $asSource_1
$asTarget[0] += $asSource_2[0]
ReDim $asSource_2[$asSource_2[0] + 1]
If $iSort = 1 Then __ArrayDualPivotSort($asSource_2, 1, $asSource_2[0])
_ArrayConcatenate($asTarget, $asSource_2, 1)
EndFunc
Func __FLTAR_AddToList(ByRef $aList, $vValue_0, $vValue_1 = -1)
If $vValue_1 = -1 Then
$aList[0] += 1
If UBound($aList) <= $aList[0] Then ReDim $aList[UBound($aList) * 2]
$aList[$aList[0]] = $vValue_0
Else
$aList[0][0] += 1
If UBound($aList) <= $aList[0][0] Then ReDim $aList[UBound($aList) * 2][2]
$aList[$aList[0][0]][0] = $vValue_0
$aList[$aList[0][0]][1] = $vValue_1
EndIf
EndFunc
Func __FLTAR_ListToMask(ByRef $sMask, $sList)
If StringRegExp($sList, "\\|/|:|\<|\>|\|") Then Return 0
$sList = StringReplace(StringStripWS(StringRegExpReplace($sList, "\s*;\s*", ";"), BitOR($STR_STRIPLEADING, $STR_STRIPTRAILING)), ";", "|")
$sList = StringReplace(StringReplace(StringRegExpReplace($sList, "[][$^.{}()+\-]", "\\$0"), "?", "."), "*", ".*?")
$sMask = "(?i)^(" & $sList & ")\z"
Return 1
EndFunc
Func _FilePrint($sFilePath, $iShow = @SW_HIDE)
Return ShellExecute($sFilePath, "", @WorkingDir, "print", $iShow = Default ? @SW_HIDE : $iShow)
EndFunc
Func _FileReadToArray($sFilePath, ByRef $vReturn, $iFlags = $FRTA_COUNT, $sDelimiter = "")
$vReturn = 0
If $iFlags = Default Then $iFlags = $FRTA_COUNT
If $sDelimiter = Default Then $sDelimiter = ""
Local $bExpand = True
If BitAND($iFlags, $FRTA_INTARRAYS) Then
$bExpand = False
$iFlags -= $FRTA_INTARRAYS
EndIf
Local $iEntire = $STR_CHRSPLIT
If BitAND($iFlags, $FRTA_ENTIRESPLIT) Then
$iEntire = $STR_ENTIRESPLIT
$iFlags -= $FRTA_ENTIRESPLIT
EndIf
Local $iNoCount = 0
If $iFlags <> $FRTA_COUNT Then
$iFlags = $FRTA_NOCOUNT
$iNoCount = $STR_NOCOUNT
EndIf
If $sDelimiter Then
Local $aLines = FileReadToArray($sFilePath)
If @error Then Return SetError(@error, 0, 0)
Local $iDim_1 = UBound($aLines) + $iFlags
If $bExpand Then
Local $iDim_2 = UBound(StringSplit($aLines[0], $sDelimiter, $iEntire + $STR_NOCOUNT))
Local $aTemp_Array[$iDim_1][$iDim_2]
Local $iFields,  $aSplit
For $i = 0 To $iDim_1 - $iFlags - 1
$aSplit = StringSplit($aLines[$i], $sDelimiter, $iEntire + $STR_NOCOUNT)
$iFields = UBound($aSplit)
If $iFields <> $iDim_2 Then
Return SetError(3, 0, 0)
EndIf
For $j = 0 To $iFields - 1
$aTemp_Array[$i + $iFlags][$j] = $aSplit[$j]
Next
Next
If $iDim_2 < 2 Then Return SetError(4, 0, 0)
If $iFlags Then
$aTemp_Array[0][0] = $iDim_1 - $iFlags
$aTemp_Array[0][1] = $iDim_2
EndIf
Else
Local $aTemp_Array[$iDim_1]
For $i = 0 To $iDim_1 - $iFlags - 1
$aTemp_Array[$i + $iFlags] = StringSplit($aLines[$i], $sDelimiter, $iEntire + $iNoCount)
Next
If $iFlags Then
$aTemp_Array[0] = $iDim_1 - $iFlags
EndIf
EndIf
$vReturn = $aTemp_Array
Else
If $iFlags Then
Local $hFileOpen = FileOpen($sFilePath, $FO_READ)
If $hFileOpen = -1 Then Return SetError(1, 0, 0)
Local $sFileRead = FileRead($hFileOpen)
FileClose($hFileOpen)
If StringLen($sFileRead) Then
$vReturn = StringRegExp(@LF & $sFileRead, "(?|(\N+)\z|(\N*)(?:\R))", $STR_REGEXPARRAYGLOBALMATCH)
$vReturn[0] = UBound($vReturn) - 1
Else
Return SetError(2, 0, 0)
EndIf
Else
$vReturn = FileReadToArray($sFilePath)
If @error Then
$vReturn = 0
Return SetError(@error, 0, 0)
EndIf
EndIf
EndIf
Return 1
EndFunc
Func _FileWriteFromArray($sFilePath, Const ByRef $aArray, $iBase = Default, $iUBound = Default, $sDelimiter = "|")
Local $iReturn = 0
If Not IsArray($aArray) Then Return SetError(2, 0, $iReturn)
Local $iDims = UBound($aArray, $UBOUND_DIMENSIONS)
If $iDims > 2 Then Return SetError(4, 0, 0)
Local $iLast = UBound($aArray) - 1
If $iUBound = Default Or $iUBound > $iLast Then $iUBound = $iLast
If $iBase < 0 Or $iBase = Default Then $iBase = 0
If $iBase > $iUBound Then Return SetError(5, 0, $iReturn)
If $sDelimiter = Default Then $sDelimiter = "|"
Local $hFileOpen = $sFilePath
If IsString($sFilePath) Then
$hFileOpen = FileOpen($sFilePath, $FO_OVERWRITE)
If $hFileOpen = -1 Then Return SetError(1, 0, $iReturn)
EndIf
Local $iError = 0
$iReturn = 1
Switch $iDims
Case 1
For $i = $iBase To $iUBound
If Not FileWrite($hFileOpen, $aArray[$i] & @CRLF) Then
$iError = 3
$iReturn = 0
ExitLoop
EndIf
Next
Case 2
Local $sTemp = ""
For $i = $iBase To $iUBound
$sTemp = $aArray[$i][0]
For $j = 1 To UBound($aArray, $UBOUND_COLUMNS) - 1
$sTemp &= $sDelimiter & $aArray[$i][$j]
Next
If Not FileWrite($hFileOpen, $sTemp & @CRLF) Then
$iError = 3
$iReturn = 0
ExitLoop
EndIf
Next
EndSwitch
If IsString($sFilePath) Then FileClose($hFileOpen)
Return SetError($iError, 0, $iReturn)
EndFunc
Func _FileWriteLog($sLogPath, $sLogMsg, $iFlag = -1)
Local $iOpenMode = $FO_APPEND
Local $sMsg = @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & " : " & $sLogMsg
If $iFlag = Default Then $iFlag = -1
If $iFlag <> -1 Then
$iOpenMode = $FO_OVERWRITE
$sMsg &= @CRLF & FileRead($sLogPath)
EndIf
Local $hFileOpen = $sLogPath
If IsString($sLogPath) Then $hFileOpen = FileOpen($sLogPath, $iOpenMode)
If $hFileOpen = -1 Then Return SetError(1, 0, 0)
Local $iReturn = FileWriteLine($hFileOpen, $sMsg)
If IsString($sLogPath) Then $iReturn = FileClose($hFileOpen)
If $iReturn <= 0 Then Return SetError(2, $iReturn, 0)
Return $iReturn
EndFunc
Func _FileWriteToLine($sFilePath, $iLine, $sText, $bOverWrite = False, $bFill = False)
If $bOverWrite = Default Then $bOverWrite = False
If $bFill = Default Then $bFill = False
If Not FileExists($sFilePath) Then Return SetError(2, 0, 0)
If $iLine <= 0 Then Return SetError(4, 0, 0)
If Not (IsBool($bOverWrite) Or $bOverWrite = 0 Or $bOverWrite = 1) Then Return SetError(5, 0, 0)
If Not IsString($sText) Then
$sText = String($sText)
If $sText = "" Then Return SetError(6, 0, 0)
EndIf
If Not IsBool($bFill) Then Return SetError(7, 0, 0)
Local $aArray = FileReadToArray($sFilePath)
If @error Then Local $aArray[0]
Local $iUBound = UBound($aArray) - 1
If $bFill Then
If $iUBound < $iLine Then
ReDim $aArray[$iLine]
$iUBound = $iLine - 1
EndIf
Else
If ($iUBound + 1) < $iLine Then Return SetError(1, 0, 0)
EndIf
$aArray[$iLine - 1] = ($bOverWrite ? $sText : $sText & @CRLF & $aArray[$iLine - 1])
Local $sData = ""
For $i = 0 To $iUBound
$sData &= $aArray[$i] & @CRLF
Next
$sData = StringTrimRight($sData, StringLen(@CRLF))
Local $hFileOpen = FileOpen($sFilePath, FileGetEncoding($sFilePath) + $FO_OVERWRITE)
If $hFileOpen = -1 Then Return SetError(3, 0, 0)
FileWrite($hFileOpen, $sData)
FileClose($hFileOpen)
Return 1
EndFunc
Func _PathFull($sRelativePath, $sBasePath = @WorkingDir)
If Not $sRelativePath Or $sRelativePath = "." Then Return $sBasePath
Local $sFullPath = StringReplace($sRelativePath, "/", "\")
Local Const $sFullPathConst = $sFullPath
Local $sPath
Local $bRootOnly = StringLeft($sFullPath, 1) = "\" And StringMid($sFullPath, 2, 1) <> "\"
If $sBasePath = Default Then $sBasePath = @WorkingDir
For $i = 1 To 2
$sPath = StringLeft($sFullPath, 2)
If $sPath = "\\" Then
$sFullPath = StringTrimLeft($sFullPath, 2)
Local $nServerLen = StringInStr($sFullPath, "\") - 1
$sPath = "\\" & StringLeft($sFullPath, $nServerLen)
$sFullPath = StringTrimLeft($sFullPath, $nServerLen)
ExitLoop
ElseIf StringRight($sPath, 1) = ":" Then
$sFullPath = StringTrimLeft($sFullPath, 2)
ExitLoop
Else
$sFullPath = $sBasePath & "\" & $sFullPath
EndIf
Next
If StringLeft($sFullPath, 1) <> "\" Then
If StringLeft($sFullPathConst, 2) = StringLeft($sBasePath, 2) Then
$sFullPath = $sBasePath & "\" & $sFullPath
Else
$sFullPath = "\" & $sFullPath
EndIf
EndIf
Local $aTemp = StringSplit($sFullPath, "\")
Local $aPathParts[$aTemp[0]], $j = 0
For $i = 2 To $aTemp[0]
If $aTemp[$i] = ".." Then
If $j Then $j -= 1
ElseIf Not ($aTemp[$i] = "" And $i <> $aTemp[0]) And $aTemp[$i] <> "." Then
$aPathParts[$j] = $aTemp[$i]
$j += 1
EndIf
Next
$sFullPath = $sPath
If Not $bRootOnly Then
For $i = 0 To $j - 1
$sFullPath &= "\" & $aPathParts[$i]
Next
Else
$sFullPath &= $sFullPathConst
If StringInStr($sFullPath, "..") Then $sFullPath = _PathFull($sFullPath)
EndIf
Do
$sFullPath = StringReplace($sFullPath, ".\", "\")
Until @extended = 0
Return $sFullPath
EndFunc
Func _PathGetRelative($sFrom, $sTo)
If StringRight($sFrom, 1) <> "\" Then $sFrom &= "\"
If StringRight($sTo, 1) <> "\" Then $sTo &= "\"
If $sFrom = $sTo Then Return SetError(1, 0, StringTrimRight($sTo, 1))
Local $asFrom = StringSplit($sFrom, "\")
Local $asTo = StringSplit($sTo, "\")
If $asFrom[1] <> $asTo[1] Then Return SetError(2, 0, StringTrimRight($sTo, 1))
Local $i = 2
Local $iDiff = 1
While 1
If $asFrom[$i] <> $asTo[$i] Then
$iDiff = $i
ExitLoop
EndIf
$i += 1
WEnd
$i = 1
Local $sRelPath = ""
For $j = 1 To $asTo[0]
If $i >= $iDiff Then
$sRelPath &= "\" & $asTo[$i]
EndIf
$i += 1
Next
$sRelPath = StringTrimLeft($sRelPath, 1)
$i = 1
For $j = 1 To $asFrom[0]
If $i > $iDiff Then
$sRelPath = "..\" & $sRelPath
EndIf
$i += 1
Next
If StringRight($sRelPath, 1) == "\" Then $sRelPath = StringTrimRight($sRelPath, 1)
Return $sRelPath
EndFunc
Func _PathMake($sDrive, $sDir, $sFileName, $sExtension)
If StringLen($sDrive) Then
If Not (StringLeft($sDrive, 2) = "\\") Then $sDrive = StringLeft($sDrive, 1) & ":"
EndIf
If StringLen($sDir) Then
If Not (StringRight($sDir, 1) = "\") And Not (StringRight($sDir, 1) = "/") Then $sDir = $sDir & "\"
Else
$sDir = "\"
EndIf
If StringLen($sDir) Then
If Not (StringLeft($sDir, 1) = "\") And Not (StringLeft($sDir, 1) = "/") Then $sDir = "\" & $sDir
EndIf
If StringLen($sExtension) Then
If Not (StringLeft($sExtension, 1) = ".") Then $sExtension = "." & $sExtension
EndIf
Return $sDrive & $sDir & $sFileName & $sExtension
EndFunc
Func _PathSplit($sFilePath, ByRef $sDrive, ByRef $sDir, ByRef $sFileName, ByRef $sExtension)
Local $aArray = StringRegExp($sFilePath, "^\h*((?:\\\\\?\\)*(\\\\[^\?\/\\]+|[A-Za-z]:)?(.*[\/\\]\h*)?((?:[^\.\/\\]|(?(?=\.[^\/\\]*\.)\.))*)?([^\/\\]*))$", $STR_REGEXPARRAYMATCH)
If @error Then
ReDim $aArray[5]
$aArray[$PATH_ORIGINAL] = $sFilePath
EndIf
$sDrive = $aArray[$PATH_DRIVE]
If StringLeft($aArray[$PATH_DIRECTORY], 1) == "/" Then
$sDir = StringRegExpReplace($aArray[$PATH_DIRECTORY], "\h*[\/\\]+\h*", "\/")
Else
$sDir = StringRegExpReplace($aArray[$PATH_DIRECTORY], "\h*[\/\\]+\h*", "\\")
EndIf
$aArray[$PATH_DIRECTORY] = $sDir
$sFileName = $aArray[$PATH_FILENAME]
$sExtension = $aArray[$PATH_EXTENSION]
Return $aArray
EndFunc
Func _ReplaceStringInFile($sFilePath, $sSearchString, $sReplaceString, $iCaseSensitive = 0, $iOccurance = 1)
If StringInStr(FileGetAttrib($sFilePath), "R") Then Return SetError(1, 0, -1)
Local $hFileOpen = FileOpen($sFilePath, $FO_READ)
If $hFileOpen = -1 Then Return SetError(2, 0, -1)
Local $sFileRead = FileRead($hFileOpen)
FileClose($hFileOpen)
If $iCaseSensitive = Default Then $iCaseSensitive = 0
If $iOccurance = Default Then $iOccurance = 1
$sFileRead = StringReplace($sFileRead, $sSearchString, $sReplaceString, 1 - $iOccurance, $iCaseSensitive)
Local $iReturn = @extended
If $iReturn Then
Local $iFileEncoding = FileGetEncoding($sFilePath)
$hFileOpen = FileOpen($sFilePath, $iFileEncoding + $FO_OVERWRITE)
If $hFileOpen = -1 Then Return SetError(3, 0, -1)
FileWrite($hFileOpen, $sFileRead)
FileClose($hFileOpen)
EndIf
Return $iReturn
EndFunc
Func _TempFile($sDirectoryName = @TempDir, $sFilePrefix = "~", $sFileExtension = ".tmp", $iRandomLength = 7)
If $iRandomLength = Default Or $iRandomLength <= 0 Then $iRandomLength = 7
If $sDirectoryName = Default Or (Not FileExists($sDirectoryName)) Then $sDirectoryName = @TempDir
If $sFileExtension = Default Then $sFileExtension = ".tmp"
If $sFilePrefix = Default Then $sFilePrefix = "~"
If Not FileExists($sDirectoryName) Then $sDirectoryName = @ScriptDir
$sDirectoryName = StringRegExpReplace($sDirectoryName, "[\\/]+$", "")
$sFileExtension = StringRegExpReplace($sFileExtension, "^\.+", "")
$sFilePrefix = StringRegExpReplace($sFilePrefix, '[\\/:*?"<>|]', "")
Local $sTempName = ""
Do
$sTempName = ""
While StringLen($sTempName) < $iRandomLength
$sTempName &= Chr(Random(97, 122, 1))
WEnd
$sTempName = $sDirectoryName & "\" & $sFilePrefix & $sTempName & "." & $sFileExtension
Until Not FileExists($sTempName)
Return $sTempName
EndFunc
Global Const $GUI_EVENT_SINGLE = 0
Global Const $GUI_EVENT_ARRAY = 1
Global Const $GUI_EVENT_NONE = 0
Global Const $GUI_EVENT_CLOSE = -3
Global Const $GUI_EVENT_MINIMIZE = -4
Global Const $GUI_EVENT_RESTORE = -5
Global Const $GUI_EVENT_MAXIMIZE = -6
Global Const $GUI_EVENT_PRIMARYDOWN = -7
Global Const $GUI_EVENT_PRIMARYUP = -8
Global Const $GUI_EVENT_SECONDARYDOWN = -9
Global Const $GUI_EVENT_SECONDARYUP = -10
Global Const $GUI_EVENT_MOUSEMOVE = -11
Global Const $GUI_EVENT_RESIZED = -12
Global Const $GUI_EVENT_DROPPED = -13
Global Const $GUI_RUNDEFMSG = 'GUI_RUNDEFMSG'
Global Const $GUI_AVISTOP = 0
Global Const $GUI_AVISTART = 1
Global Const $GUI_AVICLOSE = 2
Global Const $GUI_CHECKED = 1
Global Const $GUI_INDETERMINATE = 2
Global Const $GUI_UNCHECKED = 4
Global Const $GUI_DROPACCEPTED = 8
Global Const $GUI_NODROPACCEPTED = 4096
Global Const $GUI_ACCEPTFILES = $GUI_DROPACCEPTED
Global Const $GUI_SHOW = 16
Global Const $GUI_HIDE = 32
Global Const $GUI_ENABLE = 64
Global Const $GUI_DISABLE = 128
Global Const $GUI_FOCUS = 256
Global Const $GUI_NOFOCUS = 8192
Global Const $GUI_DEFBUTTON = 512
Global Const $GUI_EXPAND = 1024
Global Const $GUI_ONTOP = 2048
Global Const $GUI_FONTNORMAL = 0
Global Const $GUI_FONTITALIC = 2
Global Const $GUI_FONTUNDER = 4
Global Const $GUI_FONTSTRIKE = 8
Global Const $GUI_DOCKAUTO = 0x0001
Global Const $GUI_DOCKLEFT = 0x0002
Global Const $GUI_DOCKRIGHT = 0x0004
Global Const $GUI_DOCKHCENTER = 0x0008
Global Const $GUI_DOCKTOP = 0x0020
Global Const $GUI_DOCKBOTTOM = 0x0040
Global Const $GUI_DOCKVCENTER = 0x0080
Global Const $GUI_DOCKWIDTH = 0x0100
Global Const $GUI_DOCKHEIGHT = 0x0200
Global Const $GUI_DOCKSIZE = 0x0300
Global Const $GUI_DOCKMENUBAR = 0x0220
Global Const $GUI_DOCKSTATEBAR = 0x0240
Global Const $GUI_DOCKALL = 0x0322
Global Const $GUI_DOCKBORDERS = 0x0066
Global Const $GUI_GR_CLOSE = 1
Global Const $GUI_GR_LINE = 2
Global Const $GUI_GR_BEZIER = 4
Global Const $GUI_GR_MOVE = 6
Global Const $GUI_GR_COLOR = 8
Global Const $GUI_GR_RECT = 10
Global Const $GUI_GR_ELLIPSE = 12
Global Const $GUI_GR_PIE = 14
Global Const $GUI_GR_DOT = 16
Global Const $GUI_GR_PIXEL = 18
Global Const $GUI_GR_HINT = 20
Global Const $GUI_GR_REFRESH = 22
Global Const $GUI_GR_PENSIZE = 24
Global Const $GUI_GR_NOBKCOLOR = -2
Global Const $GUI_BKCOLOR_DEFAULT = -1
Global Const $GUI_BKCOLOR_TRANSPARENT = -2
Global Const $GUI_BKCOLOR_LV_ALTERNATE = 0xFE000000
Global Const $GUI_READ_DEFAULT = 0
Global Const $GUI_READ_EXTENDED = 1
Global Const $GUI_CURSOR_NOOVERRIDE = 0
Global Const $GUI_CURSOR_OVERRIDE = 1
Global Const $GUI_WS_EX_PARENTDRAG = 0x00100000
Global Const $DTS_SHORTDATEFORMAT = 0
Global Const $DTS_UPDOWN = 1
Global Const $DTS_SHOWNONE = 2
Global Const $DTS_LONGDATEFORMAT = 4
Global Const $DTS_TIMEFORMAT = 9
Global Const $DTS_RIGHTALIGN = 32
Global Const $DTS_SHORTDATECENTURYFORMAT = 0x0000000C
Global Const $DTS_APPCANPARSE = 0x00000010
Global Const $DMW_LONGNAME = 0
Global Const $DMW_SHORTNAME = 1
Global Const $DMW_LOCALE_LONGNAME = 2
Global Const $DMW_LOCALE_SHORTNAME = 3
Global Const $GDT_ERROR = -1
Global Const $GDT_VALID = 0
Global Const $GDT_NONE = 1
Global Const $GDTR_MIN = 0x0001
Global Const $GDTR_MAX = 0x0002
Global Const $MCHT_NOWHERE = 0x00000000
Global Const $MCHT_TITLE = 0x00010000
Global Const $MCHT_CALENDAR = 0x00020000
Global Const $MCHT_TODAYLINK = 0x00030000
Global Const $MCHT_NEXT = 0x01000000
Global Const $MCHT_PREV = 0x02000000
Global Const $MCHT_TITLEBK = 0x00010000
Global Const $MCHT_TITLEMONTH = 0x00010001
Global Const $MCHT_TITLEYEAR = 0x00010002
Global Const $MCHT_TITLEBTNNEXT = 0x01010003
Global Const $MCHT_TITLEBTNPREV = 0x02010003
Global Const $MCHT_CALENDARBK = 0x00020000
Global Const $MCHT_CALENDARDATE = 0x00020001
Global Const $MCHT_CALENDARDAY = 0x00020002
Global Const $MCHT_CALENDARWEEKNUM = 0x00020003
Global Const $MCHT_CALENDARDATENEXT = 0x01020000
Global Const $MCHT_CALENDARDATEPREV = 0x02020000
Global Const $MCS_DAYSTATE = 0x0001
Global Const $MCS_MULTISELECT = 0x0002
Global Const $MCS_WEEKNUMBERS = 0x0004
Global Const $MCS_NOTODAYCIRCLE = 0x0008
Global Const $MCS_NOTODAY = 0x0010
Global Const $MCS_NOTRAILINGDATES = 0x0040
Global Const $MCS_SHORTDAYSOFWEEK = 0x0080
Global Const $MCS_NOSELCHANGEONNAV = 0x0100
Global Const $MCM_FIRST = 0x1000
Global Const $MCM_GETCALENDARBORDER = ($MCM_FIRST + 31)
Global Const $MCM_GETCALENDARCOUNT = ($MCM_FIRST + 23)
Global Const $MCM_GETCALENDARGRIDINFO = ($MCM_FIRST + 24)
Global Const $MCM_GETCALID = ($MCM_FIRST + 27)
Global Const $MCM_GETCOLOR = ($MCM_FIRST + 11)
Global Const $MCM_GETCURRENTVIEW = ($MCM_FIRST + 22)
Global Const $MCM_GETCURSEL = ($MCM_FIRST + 1)
Global Const $MCM_GETFIRSTDAYOFWEEK = ($MCM_FIRST + 16)
Global Const $MCM_GETMAXSELCOUNT = ($MCM_FIRST + 3)
Global Const $MCM_GETMAXTODAYWIDTH = ($MCM_FIRST + 21)
Global Const $MCM_GETMINREQRECT = ($MCM_FIRST + 9)
Global Const $MCM_GETMONTHDELTA = ($MCM_FIRST + 19)
Global Const $MCM_GETMONTHRANGE = ($MCM_FIRST + 7)
Global Const $MCM_GETRANGE = ($MCM_FIRST + 17)
Global Const $MCM_GETSELRANGE = ($MCM_FIRST + 5)
Global Const $MCM_GETTODAY = ($MCM_FIRST + 13)
Global Const $MCM_GETUNICODEFORMAT = 0x2000 + 6
Global Const $MCM_HITTEST = ($MCM_FIRST + 14)
Global Const $MCM_SETCALENDARBORDER = ($MCM_FIRST + 30)
Global Const $MCM_SETCALID = ($MCM_FIRST + 28)
Global Const $MCM_SETCOLOR = ($MCM_FIRST + 10)
Global Const $MCM_SETCURRENTVIEW = ($MCM_FIRST + 32)
Global Const $MCM_SETCURSEL = ($MCM_FIRST + 2)
Global Const $MCM_SETDAYSTATE = ($MCM_FIRST + 8)
Global Const $MCM_SETFIRSTDAYOFWEEK = ($MCM_FIRST + 15)
Global Const $MCM_SETMAXSELCOUNT = ($MCM_FIRST + 4)
Global Const $MCM_SETMONTHDELTA = ($MCM_FIRST + 20)
Global Const $MCM_SETRANGE = ($MCM_FIRST + 18)
Global Const $MCM_SETSELRANGE = ($MCM_FIRST + 6)
Global Const $MCM_SETTODAY = ($MCM_FIRST + 12)
Global Const $MCM_SETUNICODEFORMAT = 0x2000 + 5
Global Const $MCM_SIZERECTTOMIN = ($MCM_FIRST + 29)
Global Const $MCN_FIRST = -746
Global Const $MCN_SELCHANGE = ($MCN_FIRST - 3)
Global Const $MCN_GETDAYSTATE = ($MCN_FIRST - 1)
Global Const $MCN_SELECT = ($MCN_FIRST)
Global Const $MCN_VIEWCHANGE = ($MCN_FIRST - 4)
Global Const $MCSC_BACKGROUND = 0
Global Const $MCSC_MONTHBK = 4
Global Const $MCSC_TEXT = 1
Global Const $MCSC_TITLEBK = 2
Global Const $MCSC_TITLETEXT = 3
Global Const $MCSC_TRAILINGTEXT = 5
Global Const $DTM_FIRST = 0x1000
Global Const $DTM_GETSYSTEMTIME = $DTM_FIRST + 1
Global Const $DTM_SETSYSTEMTIME = $DTM_FIRST + 2
Global Const $DTM_GETRANGE = $DTM_FIRST + 3
Global Const $DTM_SETRANGE = $DTM_FIRST + 4
Global Const $DTM_SETFORMAT = $DTM_FIRST + 5
Global Const $DTM_SETMCCOLOR = $DTM_FIRST + 6
Global Const $DTM_GETMCCOLOR = $DTM_FIRST + 7
Global Const $DTM_GETMONTHCAL = $DTM_FIRST + 8
Global Const $DTM_SETMCFONT = $DTM_FIRST + 9
Global Const $DTM_GETMCFONT = $DTM_FIRST + 10
Global Const $DTM_SETFORMATW = $DTM_FIRST + 50
Global Const $DTN_FIRST = -740
Global Const $DTN_FIRST2 = -753
Global Const $DTN_DATETIMECHANGE = $DTN_FIRST2 - 6
Global Const $DTN_USERSTRING = $DTN_FIRST2 - 5
Global Const $DTN_WMKEYDOWN = $DTN_FIRST2 - 4
Global Const $DTN_FORMAT = $DTN_FIRST2 - 3
Global Const $DTN_FORMATQUERY = $DTN_FIRST2 - 2
Global Const $DTN_DROPDOWN = $DTN_FIRST2 - 1
Global Const $DTN_CLOSEUP = $DTN_FIRST2 - 0
Global Const $DTN_USERSTRINGW = $DTN_FIRST - 5
Global Const $DTN_WMKEYDOWNW = $DTN_FIRST - 4
Global Const $DTN_FORMATW = $DTN_FIRST - 3
Global Const $DTN_FORMATQUERYW = $DTN_FIRST - 2
Global Const $GUI_SS_DEFAULT_DATE = $DTS_LONGDATEFORMAT
Global Const $GUI_SS_DEFAULT_MONTHCAL = 0
Global Const $GMEM_FIXED = 0x0000
Global Const $GMEM_MOVEABLE = 0x0002
Global Const $GMEM_NOCOMPACT = 0x0010
Global Const $GMEM_NODISCARD = 0x0020
Global Const $GMEM_ZEROINIT = 0x0040
Global Const $GMEM_MODIFY = 0x0080
Global Const $GMEM_DISCARDABLE = 0x0100
Global Const $GMEM_NOT_BANKED = 0x1000
Global Const $GMEM_SHARE = 0x2000
Global Const $GMEM_DDESHARE = 0x2000
Global Const $GMEM_NOTIFY = 0x4000
Global Const $GMEM_LOWER = 0x1000
Global Const $GMEM_VALID_FLAGS = 0x7F72
Global Const $GMEM_INVALID_HANDLE = 0x8000
Global Const $GPTR = BitOR($GMEM_FIXED, $GMEM_ZEROINIT)
Global Const $GHND = BitOR($GMEM_MOVEABLE, $GMEM_ZEROINIT)
Global Const $MEM_COMMIT = 0x00001000
Global Const $MEM_RESERVE = 0x00002000
Global Const $MEM_TOP_DOWN = 0x00100000
Global Const $MEM_SHARED = 0x08000000
Global Const $PAGE_NOACCESS = 0x00000001
Global Const $PAGE_READONLY = 0x00000002
Global Const $PAGE_READWRITE = 0x00000004
Global Const $PAGE_EXECUTE = 0x00000010
Global Const $PAGE_EXECUTE_READ = 0x00000020
Global Const $PAGE_EXECUTE_READWRITE = 0x00000040
Global Const $PAGE_EXECUTE_WRITECOPY = 0x00000080
Global Const $PAGE_GUARD = 0x00000100
Global Const $PAGE_NOCACHE = 0x00000200
Global Const $PAGE_WRITECOMBINE = 0x00000400
Global Const $PAGE_WRITECOPY = 0x00000008
Global Const $MEM_DECOMMIT = 0x00004000
Global Const $MEM_RELEASE = 0x00008000
Global Enum $MEM_LOAD, $MEM_TOTALPHYSRAM, $MEM_AVAILPHYSRAM, $MEM_TOTALPAGEFILE, $MEM_AVAILPAGEFILE, $MEM_TOTALVIRTUAL, $MEM_AVAILVIRTUAL
Global Const $PROCESS_TERMINATE = 0x00000001
Global Const $PROCESS_CREATE_THREAD = 0x00000002
Global Const $PROCESS_SET_SESSIONID = 0x00000004
Global Const $PROCESS_VM_OPERATION = 0x00000008
Global Const $PROCESS_VM_READ = 0x00000010
Global Const $PROCESS_VM_WRITE = 0x00000020
Global Const $PROCESS_DUP_HANDLE = 0x00000040
Global Const $PROCESS_CREATE_PROCESS = 0x00000080
Global Const $PROCESS_SET_QUOTA = 0x00000100
Global Const $PROCESS_SET_INFORMATION = 0x00000200
Global Const $PROCESS_QUERY_INFORMATION = 0x00000400
Global Const $PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
Global Const $PROCESS_SUSPEND_RESUME = 0x00000800
Global Const $PROCESS_ALL_ACCESS = 0x001F0FFF
Global Const $SE_ASSIGNPRIMARYTOKEN_NAME = "SeAssignPrimaryTokenPrivilege"
Global Const $SE_AUDIT_NAME = "SeAuditPrivilege"
Global Const $SE_BACKUP_NAME = "SeBackupPrivilege"
Global Const $SE_CHANGE_NOTIFY_NAME = "SeChangeNotifyPrivilege"
Global Const $SE_CREATE_GLOBAL_NAME = "SeCreateGlobalPrivilege"
Global Const $SE_CREATE_PAGEFILE_NAME = "SeCreatePagefilePrivilege"
Global Const $SE_CREATE_PERMANENT_NAME = "SeCreatePermanentPrivilege"
Global Const $SE_CREATE_SYMBOLIC_LINK_NAME = 'SeCreateSymbolicLinkPrivilege'
Global Const $SE_CREATE_TOKEN_NAME = "SeCreateTokenPrivilege"
Global Const $SE_DEBUG_NAME = "SeDebugPrivilege"
Global Const $SE_ENABLE_DELEGATION_NAME = "SeEnableDelegationPrivilege"
Global Const $SE_IMPERSONATE_NAME = "SeImpersonatePrivilege"
Global Const $SE_INC_BASE_PRIORITY_NAME = "SeIncreaseBasePriorityPrivilege"
Global Const $SE_INC_WORKING_SET_NAME = 'SeIncreaseWorkingSetPrivilege'
Global Const $SE_INCREASE_QUOTA_NAME = "SeIncreaseQuotaPrivilege"
Global Const $SE_LOAD_DRIVER_NAME = "SeLoadDriverPrivilege"
Global Const $SE_LOCK_MEMORY_NAME = "SeLockMemoryPrivilege"
Global Const $SE_MACHINE_ACCOUNT_NAME = "SeMachineAccountPrivilege"
Global Const $SE_MANAGE_VOLUME_NAME = "SeManageVolumePrivilege"
Global Const $SE_PROF_SINGLE_PROCESS_NAME = "SeProfileSingleProcessPrivilege"
Global Const $SE_RELABEL_NAME = 'SeRelabelPrivilege'
Global Const $SE_REMOTE_SHUTDOWN_NAME = "SeRemoteShutdownPrivilege"
Global Const $SE_RESTORE_NAME = "SeRestorePrivilege"
Global Const $SE_SECURITY_NAME = "SeSecurityPrivilege"
Global Const $SE_SHUTDOWN_NAME = "SeShutdownPrivilege"
Global Const $SE_SYNC_AGENT_NAME = "SeSyncAgentPrivilege"
Global Const $SE_SYSTEM_ENVIRONMENT_NAME = "SeSystemEnvironmentPrivilege"
Global Const $SE_SYSTEM_PROFILE_NAME = "SeSystemProfilePrivilege"
Global Const $SE_SYSTEMTIME_NAME = "SeSystemtimePrivilege"
Global Const $SE_TAKE_OWNERSHIP_NAME = "SeTakeOwnershipPrivilege"
Global Const $SE_TCB_NAME = "SeTcbPrivilege"
Global Const $SE_TIME_ZONE_NAME = 'SeTimeZonePrivilege'
Global Const $SE_TRUSTED_CREDMAN_ACCESS_NAME = 'SeTrustedCredManAccessPrivilege'
Global Const $SE_UNSOLICITED_INPUT_NAME = "SeUnsolicitedInputPrivilege"
Global Const $SE_UNDOCK_NAME = "SeUndockPrivilege"
Global Const $SE_PRIVILEGE_ENABLED_BY_DEFAULT = 0x00000001
Global Const $SE_PRIVILEGE_ENABLED = 0x00000002
Global Const $SE_PRIVILEGE_REMOVED = 0x00000004
Global Const $SE_PRIVILEGE_USED_FOR_ACCESS = 0x80000000
Global Const $SE_GROUP_MANDATORY = 0x00000001
Global Const $SE_GROUP_ENABLED_BY_DEFAULT = 0x00000002
Global Const $SE_GROUP_ENABLED = 0x00000004
Global Const $SE_GROUP_OWNER = 0x00000008
Global Const $SE_GROUP_USE_FOR_DENY_ONLY = 0x00000010
Global Const $SE_GROUP_INTEGRITY = 0x00000020
Global Const $SE_GROUP_INTEGRITY_ENABLED = 0x00000040
Global Const $SE_GROUP_RESOURCE = 0x20000000
Global Const $SE_GROUP_LOGON_ID = 0xC0000000
Global Enum $TOKENPRIMARY = 1, $TOKENIMPERSONATION
Global Enum $SECURITYANONYMOUS = 0, $SECURITYIDENTIFICATION, $SECURITYIMPERSONATION, $SECURITYDELEGATION
Global Enum $TOKENUSER = 1, $TOKENGROUPS, $TOKENPRIVILEGES, $TOKENOWNER, $TOKENPRIMARYGROUP, $TOKENDEFAULTDACL, $TOKENSOURCE, $TOKENTYPE, $TOKENIMPERSONATIONLEVEL, $TOKENSTATISTICS, $TOKENRESTRICTEDSIDS, $TOKENSESSIONID, $TOKENGROUPSANDPRIVILEGES, $TOKENSESSIONREFERENCE, $TOKENSANDBOXINERT, $TOKENAUDITPOLICY, $TOKENORIGIN, $TOKENELEVATIONTYPE, $TOKENLINKEDTOKEN, $TOKENELEVATION, $TOKENHASRESTRICTIONS, $TOKENACCESSINFORMATION, $TOKENVIRTUALIZATIONALLOWED, $TOKENVIRTUALIZATIONENABLED, $TOKENINTEGRITYLEVEL, $TOKENUIACCESS, $TOKENMANDATORYPOLICY, $TOKENLOGONSID
Global Const $TOKEN_ASSIGN_PRIMARY = 0x00000001
Global Const $TOKEN_DUPLICATE = 0x00000002
Global Const $TOKEN_IMPERSONATE = 0x00000004
Global Const $TOKEN_QUERY = 0x00000008
Global Const $TOKEN_QUERY_SOURCE = 0x00000010
Global Const $TOKEN_ADJUST_PRIVILEGES = 0x00000020
Global Const $TOKEN_ADJUST_GROUPS = 0x00000040
Global Const $TOKEN_ADJUST_DEFAULT = 0x00000080
Global Const $TOKEN_ADJUST_SESSIONID = 0x00000100
Global Const $TOKEN_ALL_ACCESS = 0x000F01FF
Global Const $TOKEN_READ = 0x00020008
Global Const $TOKEN_WRITE = 0x000200E0
Global Const $TOKEN_EXECUTE = 0x00020000
Global Const $TOKEN_HAS_TRAVERSE_PRIVILEGE = 0x00000001
Global Const $TOKEN_HAS_BACKUP_PRIVILEGE = 0x00000002
Global Const $TOKEN_HAS_RESTORE_PRIVILEGE = 0x00000004
Global Const $TOKEN_HAS_ADMIN_GROUP = 0x00000008
Global Const $TOKEN_IS_RESTRICTED = 0x00000010
Global Const $TOKEN_SESSION_NOT_REFERENCED = 0x00000020
Global Const $TOKEN_SANDBOX_INERT = 0x00000040
Global Const $TOKEN_HAS_IMPERSONATE_PRIVILEGE = 0x00000080
Global Const $RIGHTS_DELETE = 0x00010000
Global Const $READ_CONTROL = 0x00020000
Global Const $WRITE_DAC = 0x00040000
Global Const $WRITE_OWNER = 0x00080000
Global Const $SYNCHRONIZE = 0x00100000
Global Const $ACCESS_SYSTEM_SECURITY = 0x01000000
Global Const $STANDARD_RIGHTS_REQUIRED = 0x000f0000
Global Const $STANDARD_RIGHTS_READ = $READ_CONTROL
Global Const $STANDARD_RIGHTS_WRITE = $READ_CONTROL
Global Const $STANDARD_RIGHTS_EXECUTE = $READ_CONTROL
Global Const $STANDARD_RIGHTS_ALL = 0x001F0000
Global Const $SPECIFIC_RIGHTS_ALL = 0x0000FFFF
Global Enum $NOT_USED_ACCESS = 0, $GRANT_ACCESS, $SET_ACCESS, $DENY_ACCESS, $REVOKE_ACCESS, $SET_AUDIT_SUCCESS, $SET_AUDIT_FAILURE
Global Enum $TRUSTEE_IS_UNKNOWN = 0, $TRUSTEE_IS_USER, $TRUSTEE_IS_GROUP, $TRUSTEE_IS_DOMAIN, $TRUSTEE_IS_ALIAS, $TRUSTEE_IS_WELL_KNOWN_GROUP, $TRUSTEE_IS_DELETED, $TRUSTEE_IS_INVALID, $TRUSTEE_IS_COMPUTER
Global Const $LOGON_WITH_PROFILE = 0x00000001
Global Const $LOGON_NETCREDENTIALS_ONLY = 0x00000002
Global Enum $SIDTYPEUSER = 1, $SIDTYPEGROUP, $SIDTYPEDOMAIN, $SIDTYPEALIAS, $SIDTYPEWELLKNOWNGROUP, $SIDTYPEDELETEDACCOUNT, $SIDTYPEINVALID, $SIDTYPEUNKNOWN, $SIDTYPECOMPUTER, $SIDTYPELABEL
Global Const $SID_ADMINISTRATORS = "S-1-5-32-544"
Global Const $SID_USERS = "S-1-5-32-545"
Global Const $SID_GUESTS = "S-1-5-32-546"
Global Const $SID_ACCOUNT_OPERATORS = "S-1-5-32-548"
Global Const $SID_SERVER_OPERATORS = "S-1-5-32-549"
Global Const $SID_PRINT_OPERATORS = "S-1-5-32-550"
Global Const $SID_BACKUP_OPERATORS = "S-1-5-32-551"
Global Const $SID_REPLICATOR = "S-1-5-32-552"
Global Const $SID_OWNER = "S-1-3-0"
Global Const $SID_EVERYONE = "S-1-1-0"
Global Const $SID_NETWORK = "S-1-5-2"
Global Const $SID_INTERACTIVE = "S-1-5-4"
Global Const $SID_SYSTEM = "S-1-5-18"
Global Const $SID_AUTHENTICATED_USERS = "S-1-5-11"
Global Const $SID_SCHANNEL_AUTHENTICATION = "S-1-5-64-14"
Global Const $SID_DIGEST_AUTHENTICATION = "S-1-5-64-21"
Global Const $SID_NT_SERVICE = "S-1-5-80"
Global Const $SID_UNTRUSTED_MANDATORY_LEVEL = "S-1-16-0"
Global Const $SID_LOW_MANDATORY_LEVEL = "S-1-16-4096"
Global Const $SID_MEDIUM_MANDATORY_LEVEL = "S-1-16-8192"
Global Const $SID_MEDIUM_PLUS_MANDATORY_LEVEL = "S-1-16-8448"
Global Const $SID_HIGH_MANDATORY_LEVEL = "S-1-16-12288"
Global Const $SID_SYSTEM_MANDATORY_LEVEL = "S-1-16-16384"
Global Const $SID_PROTECTED_PROCESS_MANDATORY_LEVEL = "S-1-16-20480"
Global Const $SID_SECURE_PROCESS_MANDATORY_LEVEL = "S-1-16-28672"
Global Const $SID_ALL_SERVICES = "S-1-5-80-0"
#Region Global Variables and Constants
Global Const $FORMAT_MESSAGE_ALLOCATE_BUFFER = 0x00000100
Global Const $FORMAT_MESSAGE_IGNORE_INSERTS = 0x00000200
Global Const $FORMAT_MESSAGE_FROM_STRING = 0x00000400
Global Const $FORMAT_MESSAGE_FROM_HMODULE = 0x00000800
Global Const $FORMAT_MESSAGE_FROM_SYSTEM = 0x00001000
Global Const $FORMAT_MESSAGE_ARGUMENT_ARRAY = 0x00002000
#EndRegion Global Variables and Constants
Func _WinAPI_Beep($iFreq = 500, $iDuration = 1000)
Local $aResult = DllCall("kernel32.dll", "bool", "Beep", "dword", $iFreq, "dword", $iDuration)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _WinAPI_FormatMessage($iFlags, $pSource, $iMessageID, $iLanguageID, ByRef $pBuffer, $iSize, $vArguments)
Local $sBufferType = "struct*"
If IsString($pBuffer) Then $sBufferType = "wstr"
Local $aResult = DllCall("kernel32.dll", "dword", "FormatMessageW", "dword", $iFlags, "struct*", $pSource, "dword", $iMessageID,  "dword", $iLanguageID, $sBufferType, $pBuffer, "dword", $iSize, "ptr", $vArguments)
If @error Or Not $aResult[0] Then Return SetError(@error + 10, @extended, 0)
If $sBufferType = "wstr" Then $pBuffer = $aResult[5]
Return $aResult[0]
EndFunc
Func _WinAPI_GetErrorMessage($iCode, $iLanguage = 0, Const $_iCurrentError = @error, Const $_iCurrentExtended = @extended)
Local $aRet = DllCall('kernel32.dll', 'dword', 'FormatMessageW', 'dword', 0x1000, 'ptr', 0, 'dword', $iCode,  'dword', $iLanguage, 'wstr', '', 'dword', 4096, 'ptr', 0)
If @error Or Not $aRet[0] Then Return SetError(@error, @extended, '')
Return SetError($_iCurrentError, $_iCurrentExtended, StringRegExpReplace($aRet[5], '[' & @LF & ',' & @CR & ']*\Z', ''))
EndFunc
Func _WinAPI_GetLastError(Const $_iCurrentError = @error, Const $_iCurrentExtended = @extended)
Local $aResult = DllCall("kernel32.dll", "dword", "GetLastError")
Return SetError($_iCurrentError, $_iCurrentExtended, $aResult[0])
EndFunc
Func _WinAPI_GetLastErrorMessage(Const $_iCurrentError = @error, Const $_iCurrentExtended = @extended)
Local $iLastError = _WinAPI_GetLastError()
Local $tBufferPtr = DllStructCreate("ptr")
Local $nCount = _WinAPI_FormatMessage(BitOR($FORMAT_MESSAGE_ALLOCATE_BUFFER, $FORMAT_MESSAGE_FROM_SYSTEM),  0, $iLastError, 0, $tBufferPtr, 0, 0)
If @error Then Return SetError(-@error, @extended, "")
Local $sText = ""
Local $pBuffer = DllStructGetData($tBufferPtr, 1)
If $pBuffer Then
If $nCount > 0 Then
Local $tBuffer = DllStructCreate("wchar[" & ($nCount + 1) & "]", $pBuffer)
$sText = DllStructGetData($tBuffer, 1)
If StringRight($sText, 2) = @CRLF Then $sText = StringTrimRight($sText, 2)
EndIf
DllCall("kernel32.dll", "handle", "LocalFree", "handle", $pBuffer)
EndIf
Return SetError($_iCurrentError, $_iCurrentExtended, $sText)
EndFunc
Func _WinAPI_MessageBeep($iType = 1)
Local $iSound
Switch $iType
Case 1
$iSound = 0
Case 2
$iSound = 16
Case 3
$iSound = 32
Case 4
$iSound = 48
Case 5
$iSound = 64
Case Else
$iSound = -1
EndSwitch
Local $aResult = DllCall("user32.dll", "bool", "MessageBeep", "uint", $iSound)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _WinAPI_MsgBox($iFlags, $sTitle, $sText)
BlockInput(0)
MsgBox($iFlags, $sTitle, $sText & "      ")
EndFunc
Func _WinAPI_SetLastError($iErrorCode, Const $_iCurrentError = @error, Const $_iCurrentExtended = @extended)
DllCall("kernel32.dll", "none", "SetLastError", "dword", $iErrorCode)
Return SetError($_iCurrentError, $_iCurrentExtended, Null)
EndFunc
Func _WinAPI_ShowError($sText, $bExit = True)
BlockInput(0)
MsgBox($MB_SYSTEMMODAL, "Error", $sText & "      ")
If $bExit Then Exit
EndFunc
Func _WinAPI_ShowLastError($sText = '', $bAbort = False, $iLanguage = 0, Const $_iCurrentError = @error, Const $_iCurrentExtended = @extended)
Local $sError
Local $iLastError = _WinAPI_GetLastError()
While 1
$sError = _WinAPI_GetErrorMessage($iLastError, $iLanguage)
If @error And $iLanguage Then
$iLanguage = 0
Else
ExitLoop
EndIf
WEnd
If StringStripWS($sText, $STR_STRIPLEADING + $STR_STRIPTRAILING) Then
$sText &= @CRLF & @CRLF
Else
$sText = ''
EndIf
_WinAPI_MsgBox(BitOR(0x00040000, BitShift(0x00000010, -2 * (Not $iLastError))), $iLastError, $sText & $sError)
If $iLastError Then
_WinAPI_SetLastError($iLastError)
If $bAbort Then
Exit $iLastError
EndIf
EndIf
Return SetError($_iCurrentError, $_iCurrentExtended, 1)
EndFunc
Func _WinAPI_ShowMsg($sText)
_WinAPI_MsgBox($MB_SYSTEMMODAL, "Information", $sText)
EndFunc
Func __COMErrorFormating(ByRef $oCOMError, $sPrefix = @TAB)
Local Const $STR_STRIPTRAILING = 2
Local $sError = "COM Error encountered in " & @ScriptName & " (" & $oCOMError.Scriptline & ") :" & @CRLF &  $sPrefix & "Number        " & @TAB & "= 0x" & Hex($oCOMError.Number, 8) & " (" & $oCOMError.Number & ")" & @CRLF &  $sPrefix & "WinDescription" & @TAB & "= " & StringStripWS($oCOMError.WinDescription, $STR_STRIPTRAILING) & @CRLF &  $sPrefix & "Description   " & @TAB & "= " & StringStripWS($oCOMError.Description, $STR_STRIPTRAILING) & @CRLF &  $sPrefix & "Source        " & @TAB & "= " & $oCOMError.Source & @CRLF &  $sPrefix & "HelpFile      " & @TAB & "= " & $oCOMError.HelpFile & @CRLF &  $sPrefix & "HelpContext   " & @TAB & "= " & $oCOMError.HelpContext & @CRLF &  $sPrefix & "LastDllError  " & @TAB & "= " & $oCOMError.LastDllError & @CRLF &  $sPrefix & "Retcode       " & @TAB & "= 0x" & Hex($oCOMError.retcode)
Return $sError
EndFunc
Func _Security__AdjustTokenPrivileges($hToken, $bDisableAll, $tNewState, $iBufferLen, $tPrevState = 0, $pRequired = 0)
Local $aCall = DllCall("advapi32.dll", "bool", "AdjustTokenPrivileges", "handle", $hToken, "bool", $bDisableAll, "struct*", $tNewState, "dword", $iBufferLen, "struct*", $tPrevState, "struct*", $pRequired)
If @error Then Return SetError(@error, @extended, False)
Return Not ($aCall[0] = 0)
EndFunc
Func _Security__CreateProcessWithToken($hToken, $iLogonFlags, $sCommandLine, $iCreationFlags, $sCurDir, $tSTARTUPINFO, $tPROCESS_INFORMATION)
Local $aCall = DllCall("advapi32.dll", "bool", "CreateProcessWithTokenW", "handle", $hToken, "dword", $iLogonFlags, "ptr", 0, "wstr", $sCommandLine, "dword", $iCreationFlags, "struct*", 0, "wstr", $sCurDir, "struct*", $tSTARTUPINFO, "struct*", $tPROCESS_INFORMATION)
If @error Or Not $aCall[0] Then Return SetError(@error, @extended, False)
Return True
EndFunc
Func _Security__DuplicateTokenEx($hExistingToken, $iDesiredAccess, $iImpersonationLevel, $iTokenType)
Local $aCall = DllCall("advapi32.dll", "bool", "DuplicateTokenEx", "handle", $hExistingToken, "dword", $iDesiredAccess, "struct*", 0, "int", $iImpersonationLevel, "int", $iTokenType, "handle*", 0)
If @error Or Not $aCall[0] Then Return SetError(@error, @extended, 0)
Return $aCall[6]
EndFunc
Func _Security__GetAccountSid($sAccount, $sSystem = "")
Local $aAcct = _Security__LookupAccountName($sAccount, $sSystem)
If @error Then Return SetError(@error, @extended, 0)
If IsArray($aAcct) Then Return _Security__StringSidToSid($aAcct[0])
Return ''
EndFunc
Func _Security__GetLengthSid($pSID)
If Not _Security__IsValidSid($pSID) Then Return SetError(@error + 10, @extended, 0)
Local $aCall = DllCall("advapi32.dll", "dword", "GetLengthSid", "struct*", $pSID)
If @error Then Return SetError(@error, @extended, 0)
Return $aCall[0]
EndFunc
Func _Security__GetTokenInformation($hToken, $iClass)
Local $aCall = DllCall("advapi32.dll", "bool", "GetTokenInformation", "handle", $hToken, "int", $iClass, "struct*", 0, "dword", 0, "dword*", 0)
If @error Or Not $aCall[5] Then Return SetError(@error + 10, @extended, 0)
Local $iLen = $aCall[5]
Local $tBuffer = DllStructCreate("byte[" & $iLen & "]")
$aCall = DllCall("advapi32.dll", "bool", "GetTokenInformation", "handle", $hToken, "int", $iClass, "struct*", $tBuffer, "dword", DllStructGetSize($tBuffer), "dword*", 0)
If @error Or Not $aCall[0] Then Return SetError(@error, @extended, 0)
Return $tBuffer
EndFunc
Func _Security__ImpersonateSelf($iLevel = $SECURITYIMPERSONATION)
Local $aCall = DllCall("advapi32.dll", "bool", "ImpersonateSelf", "int", $iLevel)
If @error Then Return SetError(@error, @extended, False)
Return Not ($aCall[0] = 0)
EndFunc
Func _Security__IsValidSid($pSID)
Local $aCall = DllCall("advapi32.dll", "bool", "IsValidSid", "struct*", $pSID)
If @error Then Return SetError(@error, @extended, False)
Return Not ($aCall[0] = 0)
EndFunc
Func _Security__LookupAccountName($sAccount, $sSystem = "")
Local $tData = DllStructCreate("byte SID[256]")
Local $aCall = DllCall("advapi32.dll", "bool", "LookupAccountNameW", "wstr", $sSystem, "wstr", $sAccount, "struct*", $tData, "dword*", DllStructGetSize($tData), "wstr", "", "dword*", DllStructGetSize($tData), "int*", 0)
If @error Or Not $aCall[0] Then Return SetError(@error, @extended, 0)
Local $aAcct[3]
$aAcct[0] = _Security__SidToStringSid(DllStructGetPtr($tData, "SID"))
$aAcct[1] = $aCall[5]
$aAcct[2] = $aCall[7]
Return $aAcct
EndFunc
Func _Security__LookupAccountSid($vSID, $sSystem = "")
Local $pSID, $aAcct[3]
If IsString($vSID) Then
$pSID = _Security__StringSidToSid($vSID)
Else
$pSID = $vSID
EndIf
If Not _Security__IsValidSid($pSID) Then Return SetError(@error + 10, @extended, 0)
Local $sTypeSystem = "ptr"
If $sSystem Then $sTypeSystem = "wstr"
Local $aCall = DllCall("advapi32.dll", "bool", "LookupAccountSidW", $sTypeSystem, $sSystem, "struct*", $pSID, "wstr", "", "dword*", 65536, "wstr", "", "dword*", 65536, "int*", 0)
If @error Or Not $aCall[0] Then Return SetError(@error, @extended, 0)
Local $aAcct[3]
$aAcct[0] = $aCall[3]
$aAcct[1] = $aCall[5]
$aAcct[2] = $aCall[7]
Return $aAcct
EndFunc
Func _Security__LookupPrivilegeValue($sSystem, $sName)
Local $aCall = DllCall("advapi32.dll", "bool", "LookupPrivilegeValueW", "wstr", $sSystem, "wstr", $sName, "int64*", 0)
If @error Or Not $aCall[0] Then Return SetError(@error, @extended, 0)
Return $aCall[3]
EndFunc
Func _Security__OpenProcessToken($hProcess, $iAccess)
Local $aCall = DllCall("advapi32.dll", "bool", "OpenProcessToken", "handle", $hProcess, "dword", $iAccess, "handle*", 0)
If @error Or Not $aCall[0] Then Return SetError(@error, @extended, 0)
Return $aCall[3]
EndFunc
Func _Security__OpenThreadToken($iAccess, $hThread = 0, $bOpenAsSelf = False)
If $hThread = 0 Then
Local $aResult = DllCall("kernel32.dll", "handle", "GetCurrentThread")
If @error Then Return SetError(@error + 10, @extended, 0)
$hThread = $aResult[0]
EndIf
Local $aCall = DllCall("advapi32.dll", "bool", "OpenThreadToken", "handle", $hThread, "dword", $iAccess, "bool", $bOpenAsSelf, "handle*", 0)
If @error Or Not $aCall[0] Then Return SetError(@error, @extended, 0)
Return $aCall[4]
EndFunc
Func _Security__OpenThreadTokenEx($iAccess, $hThread = 0, $bOpenAsSelf = False)
Local $hToken = _Security__OpenThreadToken($iAccess, $hThread, $bOpenAsSelf)
If $hToken = 0 Then
Local Const $ERROR_NO_TOKEN = 1008
If _WinAPI_GetLastError() <> $ERROR_NO_TOKEN Then Return SetError(20, _WinAPI_GetLastError(), 0)
If Not _Security__ImpersonateSelf() Then Return SetError(@error + 10, _WinAPI_GetLastError(), 0)
$hToken = _Security__OpenThreadToken($iAccess, $hThread, $bOpenAsSelf)
If $hToken = 0 Then Return SetError(@error, _WinAPI_GetLastError(), 0)
EndIf
Return $hToken
EndFunc
Func _Security__SetPrivilege($hToken, $sPrivilege, $bEnable)
Local $iLUID = _Security__LookupPrivilegeValue("", $sPrivilege)
If $iLUID = 0 Then Return SetError(@error + 10, @extended, False)
Local Const $tagTOKEN_PRIVILEGES = "dword Count;align 4;int64 LUID;dword Attributes"
Local $tCurrState = DllStructCreate($tagTOKEN_PRIVILEGES)
Local $iCurrState = DllStructGetSize($tCurrState)
Local $tPrevState = DllStructCreate($tagTOKEN_PRIVILEGES)
Local $iPrevState = DllStructGetSize($tPrevState)
Local $tRequired = DllStructCreate("int Data")
DllStructSetData($tCurrState, "Count", 1)
DllStructSetData($tCurrState, "LUID", $iLUID)
If Not _Security__AdjustTokenPrivileges($hToken, False, $tCurrState, $iCurrState, $tPrevState, $tRequired) Then Return SetError(2, @error, False)
DllStructSetData($tPrevState, "Count", 1)
DllStructSetData($tPrevState, "LUID", $iLUID)
Local $iAttributes = DllStructGetData($tPrevState, "Attributes")
If $bEnable Then
$iAttributes = BitOR($iAttributes, $SE_PRIVILEGE_ENABLED)
Else
$iAttributes = BitAND($iAttributes, BitNOT($SE_PRIVILEGE_ENABLED))
EndIf
DllStructSetData($tPrevState, "Attributes", $iAttributes)
If Not _Security__AdjustTokenPrivileges($hToken, False, $tPrevState, $iPrevState, $tCurrState, $tRequired) Then  Return SetError(3, @error, False)
Return True
EndFunc
Func _Security__SetTokenInformation($hToken, $iTokenInformation, $vTokenInformation, $iTokenInformationLength)
Local $aCall = DllCall("advapi32.dll", "bool", "SetTokenInformation", "handle", $hToken, "int", $iTokenInformation, "struct*", $vTokenInformation, "dword", $iTokenInformationLength)
If @error Or Not $aCall[0] Then Return SetError(@error, @extended, False)
Return True
EndFunc
Func _Security__SidToStringSid($pSID)
If Not _Security__IsValidSid($pSID) Then Return SetError(@error + 10, 0, "")
Local $aCall = DllCall("advapi32.dll", "bool", "ConvertSidToStringSidW", "struct*", $pSID, "ptr*", 0)
If @error Or Not $aCall[0] Then Return SetError(@error, @extended, "")
Local $pStringSid = $aCall[2]
Local $aLen = DllCall("kernel32.dll", "int", "lstrlenW", "struct*", $pStringSid)
Local $sSID = DllStructGetData(DllStructCreate("wchar Text[" & $aLen[0] + 1 & "]", $pStringSid), "Text")
DllCall("kernel32.dll", "handle", "LocalFree", "handle", $pStringSid)
Return $sSID
EndFunc
Func _Security__SidTypeStr($iType)
Switch $iType
Case $SIDTYPEUSER
Return "User"
Case $SIDTYPEGROUP
Return "Group"
Case $SIDTYPEDOMAIN
Return "Domain"
Case $SIDTYPEALIAS
Return "Alias"
Case $SIDTYPEWELLKNOWNGROUP
Return "Well Known Group"
Case $SIDTYPEDELETEDACCOUNT
Return "Deleted Account"
Case $SIDTYPEINVALID
Return "Invalid"
Case $SIDTYPEUNKNOWN
Return "Unknown Type"
Case $SIDTYPECOMPUTER
Return "Computer"
Case $SIDTYPELABEL
Return "A mandatory integrity label SID"
Case Else
Return "Unknown SID Type"
EndSwitch
EndFunc
Func _Security__StringSidToSid($sSID)
Local $aCall = DllCall("advapi32.dll", "bool", "ConvertStringSidToSidW", "wstr", $sSID, "ptr*", 0)
If @error Or Not $aCall[0] Then Return SetError(@error, @extended, 0)
Local $pSID = $aCall[2]
Local $tBuffer = DllStructCreate("byte Data[" & _Security__GetLengthSid($pSID) & "]", $pSID)
Local $tSID = DllStructCreate("byte Data[" & DllStructGetSize($tBuffer) & "]")
DllStructSetData($tSID, "Data", DllStructGetData($tBuffer, "Data"))
DllCall("kernel32.dll", "handle", "LocalFree", "handle", $pSID)
Return $tSID
EndFunc
Global Const $tagPOINT = "struct;long X;long Y;endstruct"
Global Const $tagRECT = "struct;long Left;long Top;long Right;long Bottom;endstruct"
Global Const $tagSIZE = "struct;long X;long Y;endstruct"
Global Const $tagMARGINS = "int cxLeftWidth;int cxRightWidth;int cyTopHeight;int cyBottomHeight"
Global Const $tagFILETIME = "struct;dword Lo;dword Hi;endstruct"
Global Const $tagSYSTEMTIME = "struct;word Year;word Month;word Dow;word Day;word Hour;word Minute;word Second;word MSeconds;endstruct"
Global Const $tagTIME_ZONE_INFORMATION = "struct;long Bias;wchar StdName[32];word StdDate[8];long StdBias;wchar DayName[32];word DayDate[8];long DayBias;endstruct"
Global Const $tagNMHDR = "struct;hwnd hWndFrom;uint_ptr IDFrom;INT Code;endstruct"
Global Const $tagCOMBOBOXEXITEM = "uint Mask;int_ptr Item;ptr Text;int TextMax;int Image;int SelectedImage;int OverlayImage;" &  "int Indent;lparam Param"
Global Const $tagNMCBEDRAGBEGIN = $tagNMHDR & ";int ItemID;wchar szText[260]"
Global Const $tagNMCBEENDEDIT = $tagNMHDR & ";bool fChanged;int NewSelection;wchar szText[260];int Why"
Global Const $tagNMCOMBOBOXEX = $tagNMHDR & ";uint Mask;int_ptr Item;ptr Text;int TextMax;int Image;" &  "int SelectedImage;int OverlayImage;int Indent;lparam Param"
Global Const $tagDTPRANGE = "word MinYear;word MinMonth;word MinDOW;word MinDay;word MinHour;word MinMinute;" &  "word MinSecond;word MinMSecond;word MaxYear;word MaxMonth;word MaxDOW;word MaxDay;word MaxHour;" &  "word MaxMinute;word MaxSecond;word MaxMSecond;bool MinValid;bool MaxValid"
Global Const $tagNMDATETIMECHANGE = $tagNMHDR & ";dword Flag;" & $tagSYSTEMTIME
Global Const $tagNMDATETIMEFORMAT = $tagNMHDR & ";ptr Format;" & $tagSYSTEMTIME & ";ptr pDisplay;wchar Display[64]"
Global Const $tagNMDATETIMEFORMATQUERY = $tagNMHDR & ";ptr Format;struct;long SizeX;long SizeY;endstruct"
Global Const $tagNMDATETIMEKEYDOWN = $tagNMHDR & ";int VirtKey;ptr Format;" & $tagSYSTEMTIME
Global Const $tagNMDATETIMESTRING = $tagNMHDR & ";ptr UserString;" & $tagSYSTEMTIME & ";dword Flags"
Global Const $tagEVENTLOGRECORD = "dword Length;dword Reserved;dword RecordNumber;dword TimeGenerated;dword TimeWritten;dword EventID;" &  "word EventType;word NumStrings;word EventCategory;word ReservedFlags;dword ClosingRecordNumber;dword StringOffset;" &  "dword UserSidLength;dword UserSidOffset;dword DataLength;dword DataOffset"
Global Const $tagGDIP_EFFECTPARAMS_Blur = "float Radius; bool ExpandEdge"
Global Const $tagGDIP_EFFECTPARAMS_BrightnessContrast = "int BrightnessLevel; int ContrastLevel"
Global Const $tagGDIP_EFFECTPARAMS_ColorBalance = "int CyanRed; int MagentaGreen; int YellowBlue"
Global Const $tagGDIP_EFFECTPARAMS_ColorCurve = "int Adjustment; int Channel; int AdjustValue"
Global Const $tagGDIP_EFFECTPARAMS_ColorLUT = "byte LutB[256]; byte LutG[256]; byte LutR[256]; byte LutA[256]"
Global Const $tagGDIP_EFFECTPARAMS_HueSaturationLightness = "int HueLevel; int SaturationLevel; int LightnessLevel"
Global Const $tagGDIP_EFFECTPARAMS_Levels = "int Highlight; int Midtone; int Shadow"
Global Const $tagGDIP_EFFECTPARAMS_RedEyeCorrection = "uint NumberOfAreas; ptr Areas"
Global Const $tagGDIP_EFFECTPARAMS_Sharpen = "float Radius; float Amount"
Global Const $tagGDIP_EFFECTPARAMS_Tint = "int Hue; int Amount"
Global Const $tagGDIPBITMAPDATA = "uint Width;uint Height;int Stride;int Format;ptr Scan0;uint_ptr Reserved"
Global Const $tagGDIPCOLORMATRIX = "float m[25]"
Global Const $tagGDIPENCODERPARAM = "struct;byte GUID[16];ulong NumberOfValues;ulong Type;ptr Values;endstruct"
Global Const $tagGDIPENCODERPARAMS = "uint Count;" & $tagGDIPENCODERPARAM
Global Const $tagGDIPRECTF = "struct;float X;float Y;float Width;float Height;endstruct"
Global Const $tagGDIPSTARTUPINPUT = "uint Version;ptr Callback;bool NoThread;bool NoCodecs"
Global Const $tagGDIPSTARTUPOUTPUT = "ptr HookProc;ptr UnhookProc"
Global Const $tagGDIPIMAGECODECINFO = "byte CLSID[16];byte FormatID[16];ptr CodecName;ptr DllName;ptr FormatDesc;ptr FileExt;" &  "ptr MimeType;dword Flags;dword Version;dword SigCount;dword SigSize;ptr SigPattern;ptr SigMask"
Global Const $tagGDIPPENCODERPARAMS = "uint Count;byte Params[1]"
Global Const $tagHDITEM = "uint Mask;int XY;ptr Text;handle hBMP;int TextMax;int Fmt;lparam Param;int Image;int Order;uint Type;ptr pFilter;uint State"
Global Const $tagNMHDDISPINFO = $tagNMHDR & ";int Item;uint Mask;ptr Text;int TextMax;int Image;lparam lParam"
Global Const $tagNMHDFILTERBTNCLICK = $tagNMHDR & ";int Item;" & $tagRECT
Global Const $tagNMHEADER = $tagNMHDR & ";int Item;int Button;ptr pItem"
Global Const $tagGETIPAddress = "byte Field4;byte Field3;byte Field2;byte Field1"
Global Const $tagNMIPADDRESS = $tagNMHDR & ";int Field;int Value"
Global Const $tagLVFINDINFO = "struct;uint Flags;ptr Text;lparam Param;" & $tagPOINT & ";uint Direction;endstruct"
Global Const $tagLVHITTESTINFO = $tagPOINT & ";uint Flags;int Item;int SubItem;int iGroup"
Global Const $tagLVITEM = "struct;uint Mask;int Item;int SubItem;uint State;uint StateMask;ptr Text;int TextMax;int Image;lparam Param;" &  "int Indent;int GroupID;uint Columns;ptr pColumns;ptr piColFmt;int iGroup;endstruct"
Global Const $tagNMLISTVIEW = $tagNMHDR & ";int Item;int SubItem;uint NewState;uint OldState;uint Changed;" &  "struct;long ActionX;long ActionY;endstruct;lparam Param"
Global Const $tagNMLVCUSTOMDRAW = "struct;" & $tagNMHDR & ";dword dwDrawStage;handle hdc;" & $tagRECT &  ";dword_ptr dwItemSpec;uint uItemState;lparam lItemlParam;endstruct" &  ";dword clrText;dword clrTextBk;int iSubItem;dword dwItemType;dword clrFace;int iIconEffect;" &  "int iIconPhase;int iPartID;int iStateID;struct;long TextLeft;long TextTop;long TextRight;long TextBottom;endstruct;uint uAlign"
Global Const $tagNMLVDISPINFO = $tagNMHDR & ";" & $tagLVITEM
Global Const $tagNMLVFINDITEM = $tagNMHDR & ";int Start;" & $tagLVFINDINFO
Global Const $tagNMLVGETINFOTIP = $tagNMHDR & ";dword Flags;ptr Text;int TextMax;int Item;int SubItem;lparam lParam"
Global Const $tagNMITEMACTIVATE = $tagNMHDR & ";int Index;int SubItem;uint NewState;uint OldState;uint Changed;" &  $tagPOINT & ";lparam lParam;uint KeyFlags"
Global Const $tagNMLVKEYDOWN = "align 1;" & $tagNMHDR & ";word VKey;uint Flags"
Global Const $tagNMLVSCROLL = $tagNMHDR & ";int DX;int DY"
Global Const $tagMCHITTESTINFO = "uint Size;" & $tagPOINT & ";uint Hit;" & $tagSYSTEMTIME &  ";" & $tagRECT & ";int iOffset;int iRow;int iCol"
Global Const $tagMCMONTHRANGE = "word MinYear;word MinMonth;word MinDOW;word MinDay;word MinHour;word MinMinute;word MinSecond;" &  "word MinMSeconds;word MaxYear;word MaxMonth;word MaxDOW;word MaxDay;word MaxHour;word MaxMinute;word MaxSecond;" &  "word MaxMSeconds;short Span"
Global Const $tagMCRANGE = "word MinYear;word MinMonth;word MinDOW;word MinDay;word MinHour;word MinMinute;word MinSecond;" &  "word MinMSeconds;word MaxYear;word MaxMonth;word MaxDOW;word MaxDay;word MaxHour;word MaxMinute;word MaxSecond;" &  "word MaxMSeconds;short MinSet;short MaxSet"
Global Const $tagMCSELRANGE = "word MinYear;word MinMonth;word MinDOW;word MinDay;word MinHour;word MinMinute;word MinSecond;" &  "word MinMSeconds;word MaxYear;word MaxMonth;word MaxDOW;word MaxDay;word MaxHour;word MaxMinute;word MaxSecond;" &  "word MaxMSeconds"
Global Const $tagNMDAYSTATE = $tagNMHDR & ";" & $tagSYSTEMTIME & ";int DayState;ptr pDayState"
Global Const $tagNMSELCHANGE = $tagNMHDR &  ";struct;word BegYear;word BegMonth;word BegDOW;word BegDay;word BegHour;word BegMinute;word BegSecond;word BegMSeconds;endstruct;" &  "struct;word EndYear;word EndMonth;word EndDOW;word EndDay;word EndHour;word EndMinute;word EndSecond;word EndMSeconds;endstruct"
Global Const $tagNMOBJECTNOTIFY = $tagNMHDR & ";int Item;ptr piid;ptr pObject;long Result;dword dwFlags"
Global Const $tagNMTCKEYDOWN = "align 1;" & $tagNMHDR & ";word VKey;uint Flags"
Global Const $tagTVITEM = "struct;uint Mask;handle hItem;uint State;uint StateMask;ptr Text;int TextMax;int Image;int SelectedImage;" &  "int Children;lparam Param;endstruct"
Global Const $tagTVITEMEX = "struct;" & $tagTVITEM & ";int Integral;uint uStateEx;hwnd hwnd;int iExpandedImage;int iReserved;endstruct"
Global Const $tagNMTREEVIEW = $tagNMHDR & ";uint Action;" &  "struct;uint OldMask;handle OldhItem;uint OldState;uint OldStateMask;" &  "ptr OldText;int OldTextMax;int OldImage;int OldSelectedImage;int OldChildren;lparam OldParam;endstruct;" &  "struct;uint NewMask;handle NewhItem;uint NewState;uint NewStateMask;" &  "ptr NewText;int NewTextMax;int NewImage;int NewSelectedImage;int NewChildren;lparam NewParam;endstruct;" &  "struct;long PointX;long PointY;endstruct"
Global Const $tagNMTVCUSTOMDRAW = "struct;" & $tagNMHDR & ";dword DrawStage;handle HDC;" & $tagRECT &  ";dword_ptr ItemSpec;uint ItemState;lparam ItemParam;endstruct" &  ";dword ClrText;dword ClrTextBk;int Level"
Global Const $tagNMTVDISPINFO = $tagNMHDR & ";" & $tagTVITEM
Global Const $tagNMTVGETINFOTIP = $tagNMHDR & ";ptr Text;int TextMax;handle hItem;lparam lParam"
Global Const $tagNMTVITEMCHANGE = $tagNMHDR & ";uint Changed;handle hItem;uint StateNew;uint StateOld;lparam lParam;"
Global Const $tagTVHITTESTINFO = $tagPOINT & ";uint Flags;handle Item"
Global Const $tagNMTVKEYDOWN = "align 1;" & $tagNMHDR & ";word VKey;uint Flags"
Global Const $tagNMMOUSE = $tagNMHDR & ";dword_ptr ItemSpec;dword_ptr ItemData;" & $tagPOINT & ";lparam HitInfo"
Global Const $tagTOKEN_PRIVILEGES = "dword Count;align 4;int64 LUID;dword Attributes"
Global Const $tagIMAGEINFO = "handle hBitmap;handle hMask;int Unused1;int Unused2;" & $tagRECT
Global Const $tagMENUINFO = "dword Size;INT Mask;dword Style;uint YMax;handle hBack;dword ContextHelpID;ulong_ptr MenuData"
Global Const $tagMENUITEMINFO = "uint Size;uint Mask;uint Type;uint State;uint ID;handle SubMenu;handle BmpChecked;handle BmpUnchecked;" &  "ulong_ptr ItemData;ptr TypeData;uint CCH;handle BmpItem"
Global Const $tagREBARBANDINFO = "uint cbSize;uint fMask;uint fStyle;dword clrFore;dword clrBack;ptr lpText;uint cch;" &  "int iImage;hwnd hwndChild;uint cxMinChild;uint cyMinChild;uint cx;handle hbmBack;uint wID;uint cyChild;uint cyMaxChild;" &  "uint cyIntegral;uint cxIdeal;lparam lParam;uint cxHeader" & ((@OSVersion = "WIN_XP") ? "" : ";" & $tagRECT & ";uint uChevronState")
Global Const $tagNMREBARAUTOBREAK = $tagNMHDR & ";uint uBand;uint wID;lparam lParam;uint uMsg;uint fStyleCurrent;bool fAutoBreak"
Global Const $tagNMRBAUTOSIZE = $tagNMHDR & ";bool fChanged;" &  "struct;long TargetLeft;long TargetTop;long TargetRight;long TargetBottom;endstruct;" &  "struct;long ActualLeft;long ActualTop;long ActualRight;long ActualBottom;endstruct"
Global Const $tagNMREBAR = $tagNMHDR & ";dword dwMask;uint uBand;uint fStyle;uint wID;lparam lParam"
Global Const $tagNMREBARCHEVRON = $tagNMHDR & ";uint uBand;uint wID;lparam lParam;" & $tagRECT & ";lparam lParamNM"
Global Const $tagNMREBARCHILDSIZE = $tagNMHDR & ";uint uBand;uint wID;" &  "struct;long CLeft;long CTop;long CRight;long CBottom;endstruct;" &  "struct;long BLeft;long BTop;long BRight;long BBottom;endstruct"
Global Const $tagCOLORSCHEME = "dword Size;dword BtnHighlight;dword BtnShadow"
Global Const $tagNMTOOLBAR = $tagNMHDR & ";int iItem;" &  "struct;int iBitmap;int idCommand;byte fsState;byte fsStyle;dword_ptr dwData;int_ptr iString;endstruct" &  ";int cchText;ptr pszText;" & $tagRECT
Global Const $tagNMTBHOTITEM = $tagNMHDR & ";int idOld;int idNew;dword dwFlags"
Global Const $tagTBBUTTON = "int Bitmap;int Command;byte State;byte Style;dword_ptr Param;int_ptr String"
Global Const $tagTBBUTTONINFO = "uint Size;dword Mask;int Command;int Image;byte State;byte Style;word CX;dword_ptr Param;ptr Text;int TextMax"
Global Const $tagNETRESOURCE = "dword Scope;dword Type;dword DisplayType;dword Usage;ptr LocalName;ptr RemoteName;ptr Comment;ptr Provider"
Global Const $tagOVERLAPPED = "ulong_ptr Internal;ulong_ptr InternalHigh;struct;dword Offset;dword OffsetHigh;endstruct;handle hEvent"
Global Const $tagOPENFILENAME = "dword StructSize;hwnd hwndOwner;handle hInstance;ptr lpstrFilter;ptr lpstrCustomFilter;" &  "dword nMaxCustFilter;dword nFilterIndex;ptr lpstrFile;dword nMaxFile;ptr lpstrFileTitle;dword nMaxFileTitle;" &  "ptr lpstrInitialDir;ptr lpstrTitle;dword Flags;word nFileOffset;word nFileExtension;ptr lpstrDefExt;lparam lCustData;" &  "ptr lpfnHook;ptr lpTemplateName;ptr pvReserved;dword dwReserved;dword FlagsEx"
Global Const $tagBITMAPINFOHEADER = "struct;dword biSize;long biWidth;long biHeight;word biPlanes;word biBitCount;" &  "dword biCompression;dword biSizeImage;long biXPelsPerMeter;long biYPelsPerMeter;dword biClrUsed;dword biClrImportant;endstruct"
Global Const $tagBITMAPINFO = $tagBITMAPINFOHEADER & ";dword biRGBQuad[1]"
Global Const $tagBLENDFUNCTION = "byte Op;byte Flags;byte Alpha;byte Format"
Global Const $tagGUID = "struct;ulong Data1;ushort Data2;ushort Data3;byte Data4[8];endstruct"
Global Const $tagWINDOWPLACEMENT = "uint length;uint flags;uint showCmd;long ptMinPosition[2];long ptMaxPosition[2];long rcNormalPosition[4]"
Global Const $tagWINDOWPOS = "hwnd hWnd;hwnd InsertAfter;int X;int Y;int CX;int CY;uint Flags"
Global Const $tagSCROLLINFO = "uint cbSize;uint fMask;int nMin;int nMax;uint nPage;int nPos;int nTrackPos"
Global Const $tagSCROLLBARINFO = "dword cbSize;" & $tagRECT & ";int dxyLineButton;int xyThumbTop;" &  "int xyThumbBottom;int reserved;dword rgstate[6]"
Global Const $tagLOGFONT = "struct;long Height;long Width;long Escapement;long Orientation;long Weight;byte Italic;byte Underline;" &  "byte Strikeout;byte CharSet;byte OutPrecision;byte ClipPrecision;byte Quality;byte PitchAndFamily;wchar FaceName[32];endstruct"
Global Const $tagKBDLLHOOKSTRUCT = "dword vkCode;dword scanCode;dword flags;dword time;ulong_ptr dwExtraInfo"
Global Const $tagPROCESS_INFORMATION = "handle hProcess;handle hThread;dword ProcessID;dword ThreadID"
Global Const $tagSTARTUPINFO = "dword Size;ptr Reserved1;ptr Desktop;ptr Title;dword X;dword Y;dword XSize;dword YSize;dword XCountChars;" &  "dword YCountChars;dword FillAttribute;dword Flags;word ShowWindow;word Reserved2;ptr Reserved3;handle StdInput;" &  "handle StdOutput;handle StdError"
Global Const $tagSECURITY_ATTRIBUTES = "dword Length;ptr Descriptor;bool InheritHandle"
Global Const $tagWIN32_FIND_DATA = "dword dwFileAttributes;dword ftCreationTime[2];dword ftLastAccessTime[2];dword ftLastWriteTime[2];dword nFileSizeHigh;dword nFileSizeLow;dword dwReserved0;dword dwReserved1;wchar cFileName[260];wchar cAlternateFileName[14]"
Global Const $tagTEXTMETRIC = "long tmHeight;long tmAscent;long tmDescent;long tmInternalLeading;long tmExternalLeading;" &  "long tmAveCharWidth;long tmMaxCharWidth;long tmWeight;long tmOverhang;long tmDigitizedAspectX;long tmDigitizedAspectY;" &  "wchar tmFirstChar;wchar tmLastChar;wchar tmDefaultChar;wchar tmBreakChar;byte tmItalic;byte tmUnderlined;byte tmStruckOut;" &  "byte tmPitchAndFamily;byte tmCharSet"
Global Const $tagMEMMAP = "handle hProc;ulong_ptr Size;ptr Mem"
Func _MemFree(ByRef $tMemMap)
Local $pMemory = DllStructGetData($tMemMap, "Mem")
Local $hProcess = DllStructGetData($tMemMap, "hProc")
Local $bResult = _MemVirtualFreeEx($hProcess, $pMemory, 0, $MEM_RELEASE)
DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $hProcess)
If @error Then Return SetError(@error, @extended, False)
Return $bResult
EndFunc
Func _MemGlobalAlloc($iBytes, $iFlags = 0)
Local $aResult = DllCall("kernel32.dll", "handle", "GlobalAlloc", "uint", $iFlags, "ulong_ptr", $iBytes)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _MemGlobalFree($hMemory)
Local $aResult = DllCall("kernel32.dll", "ptr", "GlobalFree", "handle", $hMemory)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _MemGlobalLock($hMemory)
Local $aResult = DllCall("kernel32.dll", "ptr", "GlobalLock", "handle", $hMemory)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _MemGlobalSize($hMemory)
Local $aResult = DllCall("kernel32.dll", "ulong_ptr", "GlobalSize", "handle", $hMemory)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _MemGlobalUnlock($hMemory)
Local $aResult = DllCall("kernel32.dll", "bool", "GlobalUnlock", "handle", $hMemory)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _MemInit($hWnd, $iSize, ByRef $tMemMap)
Local $aResult = DllCall("user32.dll", "dword", "GetWindowThreadProcessId", "hwnd", $hWnd, "dword*", 0)
If @error Then Return SetError(@error + 10, @extended, 0)
Local $iProcessID = $aResult[2]
If $iProcessID = 0 Then Return SetError(1, 0, 0)
Local $iAccess = BitOR($PROCESS_VM_OPERATION, $PROCESS_VM_READ, $PROCESS_VM_WRITE)
Local $hProcess = __Mem_OpenProcess($iAccess, False, $iProcessID, True)
Local $iAlloc = BitOR($MEM_RESERVE, $MEM_COMMIT)
Local $pMemory = _MemVirtualAllocEx($hProcess, 0, $iSize, $iAlloc, $PAGE_READWRITE)
If $pMemory = 0 Then Return SetError(2, 0, 0)
$tMemMap = DllStructCreate($tagMEMMAP)
DllStructSetData($tMemMap, "hProc", $hProcess)
DllStructSetData($tMemMap, "Size", $iSize)
DllStructSetData($tMemMap, "Mem", $pMemory)
Return $pMemory
EndFunc
Func _MemMoveMemory($pSource, $pDest, $iLength)
DllCall("kernel32.dll", "none", "RtlMoveMemory", "struct*", $pDest, "struct*", $pSource, "ulong_ptr", $iLength)
If @error Then Return SetError(@error, @extended)
EndFunc
Func _MemRead(ByRef $tMemMap, $pSrce, $pDest, $iSize)
Local $aResult = DllCall("kernel32.dll", "bool", "ReadProcessMemory", "handle", DllStructGetData($tMemMap, "hProc"),  "ptr", $pSrce, "struct*", $pDest, "ulong_ptr", $iSize, "ulong_ptr*", 0)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _MemWrite(ByRef $tMemMap, $pSrce, $pDest = 0, $iSize = 0, $sSrce = "struct*")
If $pDest = 0 Then $pDest = DllStructGetData($tMemMap, "Mem")
If $iSize = 0 Then $iSize = DllStructGetData($tMemMap, "Size")
Local $aResult = DllCall("kernel32.dll", "bool", "WriteProcessMemory", "handle", DllStructGetData($tMemMap, "hProc"),  "ptr", $pDest, $sSrce, $pSrce, "ulong_ptr", $iSize, "ulong_ptr*", 0)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _MemVirtualAlloc($pAddress, $iSize, $iAllocation, $iProtect)
Local $aResult = DllCall("kernel32.dll", "ptr", "VirtualAlloc", "ptr", $pAddress, "ulong_ptr", $iSize, "dword", $iAllocation, "dword", $iProtect)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _MemVirtualAllocEx($hProcess, $pAddress, $iSize, $iAllocation, $iProtect)
Local $aResult = DllCall("kernel32.dll", "ptr", "VirtualAllocEx", "handle", $hProcess, "ptr", $pAddress, "ulong_ptr", $iSize, "dword", $iAllocation, "dword", $iProtect)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _MemVirtualFree($pAddress, $iSize, $iFreeType)
Local $aResult = DllCall("kernel32.dll", "bool", "VirtualFree", "ptr", $pAddress, "ulong_ptr", $iSize, "dword", $iFreeType)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _MemVirtualFreeEx($hProcess, $pAddress, $iSize, $iFreeType)
Local $aResult = DllCall("kernel32.dll", "bool", "VirtualFreeEx", "handle", $hProcess, "ptr", $pAddress, "ulong_ptr", $iSize, "dword", $iFreeType)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func __Mem_OpenProcess($iAccess, $bInherit, $iPID, $bDebugPriv = False)
Local $aResult = DllCall("kernel32.dll", "handle", "OpenProcess", "dword", $iAccess, "bool", $bInherit, "dword", $iPID)
If @error Then Return SetError(@error, @extended, 0)
If $aResult[0] Then Return $aResult[0]
If Not $bDebugPriv Then Return SetError(100, 0, 0)
Local $hToken = _Security__OpenThreadTokenEx(BitOR($TOKEN_ADJUST_PRIVILEGES, $TOKEN_QUERY))
If @error Then Return SetError(@error + 10, @extended, 0)
_Security__SetPrivilege($hToken, "SeDebugPrivilege", True)
Local $iError = @error
Local $iExtended = @extended
Local $iRet = 0
If Not @error Then
$aResult = DllCall("kernel32.dll", "handle", "OpenProcess", "dword", $iAccess, "bool", $bInherit, "dword", $iPID)
$iError = @error
$iExtended = @extended
If $aResult[0] Then $iRet = $aResult[0]
_Security__SetPrivilege($hToken, "SeDebugPrivilege", False)
If @error Then
$iError = @error + 20
$iExtended = @extended
EndIf
Else
$iError = @error + 30
EndIf
DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $hToken)
Return SetError($iError, $iExtended, $iRet)
EndFunc
#Region Global Variables and Constants
Global $__g_vEnum, $__g_vExt = 0
Global $__g_iRGBMode = 1
Global Const $tagOSVERSIONINFO = 'struct;dword OSVersionInfoSize;dword MajorVersion;dword MinorVersion;dword BuildNumber;dword PlatformId;wchar CSDVersion[128];endstruct'
Global Const $IMAGE_BITMAP = 0
Global Const $IMAGE_ICON = 1
Global Const $IMAGE_CURSOR = 2
Global Const $IMAGE_ENHMETAFILE = 3
Global Const $LR_DEFAULTCOLOR = 0x0000
Global Const $LR_MONOCHROME = 0x0001
Global Const $LR_COLOR = 0x0002
Global Const $LR_COPYRETURNORG = 0x0004
Global Const $LR_COPYDELETEORG = 0x0008
Global Const $LR_LOADFROMFILE = 0x0010
Global Const $LR_LOADTRANSPARENT = 0x0020
Global Const $LR_DEFAULTSIZE = 0x0040
Global Const $LR_VGACOLOR = 0x0080
Global Const $LR_LOADMAP3DCOLORS = 0x1000
Global Const $LR_CREATEDIBSECTION = 0x2000
Global Const $LR_COPYFROMRESOURCE = 0x4000
Global Const $LR_SHARED = 0x8000
Global Const $__tagCURSORINFO = "dword Size;dword Flags;handle hCursor;" & "struct;long X;long Y;endstruct"
Global Const $__WINVER = __WINVER()
#EndRegion Global Variables and Constants
#Region Functions list
#EndRegion Functions list
#Region Public Functions
Func _WinAPI_CreateFile($sFileName, $iCreation, $iAccess = 4, $iShare = 0, $iAttributes = 0, $tSecurity = 0)
Local $iDA = 0, $iSM = 0, $iCD = 0, $iFA = 0
If BitAND($iAccess, 1) <> 0 Then $iDA = BitOR($iDA, $GENERIC_EXECUTE)
If BitAND($iAccess, 2) <> 0 Then $iDA = BitOR($iDA, $GENERIC_READ)
If BitAND($iAccess, 4) <> 0 Then $iDA = BitOR($iDA, $GENERIC_WRITE)
If BitAND($iShare, 1) <> 0 Then $iSM = BitOR($iSM, $FILE_SHARE_DELETE)
If BitAND($iShare, 2) <> 0 Then $iSM = BitOR($iSM, $FILE_SHARE_READ)
If BitAND($iShare, 4) <> 0 Then $iSM = BitOR($iSM, $FILE_SHARE_WRITE)
Switch $iCreation
Case 0
$iCD = $CREATE_NEW
Case 1
$iCD = $CREATE_ALWAYS
Case 2
$iCD = $OPEN_EXISTING
Case 3
$iCD = $OPEN_ALWAYS
Case 4
$iCD = $TRUNCATE_EXISTING
EndSwitch
If BitAND($iAttributes, 1) <> 0 Then $iFA = BitOR($iFA, $FILE_ATTRIBUTE_ARCHIVE)
If BitAND($iAttributes, 2) <> 0 Then $iFA = BitOR($iFA, $FILE_ATTRIBUTE_HIDDEN)
If BitAND($iAttributes, 4) <> 0 Then $iFA = BitOR($iFA, $FILE_ATTRIBUTE_READONLY)
If BitAND($iAttributes, 8) <> 0 Then $iFA = BitOR($iFA, $FILE_ATTRIBUTE_SYSTEM)
Local $aResult = DllCall("kernel32.dll", "handle", "CreateFileW", "wstr", $sFileName, "dword", $iDA, "dword", $iSM,  "struct*", $tSecurity, "dword", $iCD, "dword", $iFA, "ptr", 0)
If @error Or ($aResult[0] = Ptr(-1)) Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_FreeLibrary($hModule)
Local $aResult = DllCall("kernel32.dll", "bool", "FreeLibrary", "handle", $hModule)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _WinAPI_GetCursorInfo()
Local $tCursor = DllStructCreate($__tagCURSORINFO)
Local $iCursor = DllStructGetSize($tCursor)
DllStructSetData($tCursor, "Size", $iCursor)
Local $aRet = DllCall("user32.dll", "bool", "GetCursorInfo", "struct*", $tCursor)
If @error Or Not $aRet[0] Then Return SetError(@error + 10, @extended, 0)
Local $aCursor[5]
$aCursor[0] = True
$aCursor[1] = DllStructGetData($tCursor, "Flags") <> 0
$aCursor[2] = DllStructGetData($tCursor, "hCursor")
$aCursor[3] = DllStructGetData($tCursor, "X")
$aCursor[4] = DllStructGetData($tCursor, "Y")
Return $aCursor
EndFunc
Func _WinAPI_GetDlgCtrlID($hWnd)
Local $aResult = DllCall("user32.dll", "int", "GetDlgCtrlID", "hwnd", $hWnd)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_GetModuleHandle($sModuleName)
Local $sModuleNameType = "wstr"
If $sModuleName = "" Then
$sModuleName = 0
$sModuleNameType = "ptr"
EndIf
Local $aResult = DllCall("kernel32.dll", "handle", "GetModuleHandleW", $sModuleNameType, $sModuleName)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_GetString($pString, $bUnicode = True)
Local $iLength = _WinAPI_StrLen($pString, $bUnicode)
If @error Or Not $iLength Then Return SetError(@error + 10, @extended, '')
Local $tString = DllStructCreate(($bUnicode ? 'wchar' : 'char') & '[' & ($iLength + 1) & ']', $pString)
If @error Then Return SetError(@error, @extended, '')
Return SetExtended($iLength, DllStructGetData($tString, 1))
EndFunc
Func _WinAPI_IsWow64Process($iPID = 0)
If Not $iPID Then $iPID = @AutoItPID
Local $hProcess = DllCall('kernel32.dll', 'handle', 'OpenProcess', 'dword', ($__WINVER < 0x0600 ? 0x00000400 : 0x00001000),  'bool', 0, 'dword', $iPID)
If @error Or Not $hProcess[0] Then Return SetError(@error + 20, @extended, False)
Local $aRet = DllCall('kernel32.dll', 'bool', 'IsWow64Process', 'handle', $hProcess[0], 'bool*', 0)
If __CheckErrorCloseHandle($aRet, $hProcess[0]) Then Return SetError(@error, @extended, False)
Return $aRet[2]
EndFunc
Func _WinAPI_LoadImage($hInstance, $sImage, $iType, $iXDesired, $iYDesired, $iLoad)
Local $aResult, $sImageType = "int"
If IsString($sImage) Then $sImageType = "wstr"
$aResult = DllCall("user32.dll", "handle", "LoadImageW", "handle", $hInstance, $sImageType, $sImage, "uint", $iType,  "int", $iXDesired, "int", $iYDesired, "uint", $iLoad)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_LoadLibrary($sFileName)
Local $aResult = DllCall("kernel32.dll", "handle", "LoadLibraryW", "wstr", $sFileName)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_PathIsDirectory($sFilePath)
Local $aRet = DllCall('shlwapi.dll', 'bool', 'PathIsDirectoryW', 'wstr', $sFilePath)
If @error Then Return SetError(@error, @extended, False)
Return $aRet[0]
EndFunc
Func _WinAPI_ReadFile($hFile, $pBuffer, $iToRead, ByRef $iRead, $tOverlapped = 0)
Local $aResult = DllCall("kernel32.dll", "bool", "ReadFile", "handle", $hFile, "struct*", $pBuffer, "dword", $iToRead,  "dword*", 0, "struct*", $tOverlapped)
If @error Then Return SetError(@error, @extended, False)
$iRead = $aResult[4]
Return $aResult[0]
EndFunc
Func _WinAPI_StrLen($pString, $bUnicode = True)
Local $W = ''
If $bUnicode Then $W = 'W'
Local $aRet = DllCall('kernel32.dll', 'int', 'lstrlen' & $W, 'struct*', $pString)
If @error Then Return SetError(@error, @extended, 0)
Return $aRet[0]
EndFunc
Func _WinAPI_SwitchColor($iColor)
If $iColor = -1 Then Return $iColor
Return BitOR(BitAND($iColor, 0x00FF00), BitShift(BitAND($iColor, 0x0000FF), -16), BitShift(BitAND($iColor, 0xFF0000), 16))
EndFunc
Func _WinAPI_WriteFile($hFile, $pBuffer, $iToWrite, ByRef $iWritten, $tOverlapped = 0)
Local $aResult = DllCall("kernel32.dll", "bool", "WriteFile", "handle", $hFile, "struct*", $pBuffer, "dword", $iToWrite,  "dword*", 0, "struct*", $tOverlapped)
If @error Then Return SetError(@error, @extended, False)
$iWritten = $aResult[4]
Return $aResult[0]
EndFunc
#EndRegion Public Functions
#Region Internal Functions
Func __CheckErrorArrayBounds(Const ByRef $aData, ByRef $iStart, ByRef $iEnd, $nDim = 1, $iDim = $UBOUND_DIMENSIONS)
If Not IsArray($aData) Then Return SetError(1, 0, 1)
If UBound($aData, $iDim) <> $nDim Then Return SetError(2, 0, 1)
If $iStart < 0 Then $iStart = 0
Local $iUBound = UBound($aData) - 1
If $iEnd < 1 Or $iEnd > $iUBound Then $iEnd = $iUBound
If $iStart > $iEnd Then Return SetError(4, 0, 1)
Return 0
EndFunc
Func __CheckErrorCloseHandle($aRet, $hFile, $bLastError = False, $iCurErr = @error, $iCurExt = @extended)
If Not $iCurErr And Not $aRet[0] Then $iCurErr = 10
Local $aLastError = DllCall("kernel32.dll", "dword", "GetLastError")
DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $hFile)
If $iCurErr Then DllCall("kernel32.dll", "none", "SetLastError", "dword", $aLastError[0])
If $bLastError Then $iCurExt = $aLastError[0]
Return SetError($iCurErr, $iCurExt, $iCurErr)
EndFunc
Func __DLL($sPath, $bPin = False)
Local $aRet = DllCall('kernel32.dll', 'bool', 'GetModuleHandleExW', 'dword', ($bPin ? 0x0001 : 0x0002), "wstr", $sPath, 'ptr*', 0)
If Not $aRet[3] Then
Local $aResult = DllCall("kernel32.dll", "handle", "LoadLibraryW", "wstr", $sPath)
If Not $aResult[0] Then Return 0
EndIf
Return 1
EndFunc
Func __EnumWindowsProc($hWnd, $bVisible)
Local $aResult
If $bVisible Then
$aResult = DllCall("user32.dll", "bool", "IsWindowVisible", "hwnd", $hWnd)
If Not $aResult[0] Then
Return 1
EndIf
EndIf
__Inc($__g_vEnum)
$__g_vEnum[$__g_vEnum[0][0]][0] = $hWnd
$aResult = DllCall("user32.dll", "int", "GetClassNameW", "hwnd", $hWnd, "wstr", "", "int", 4096)
$__g_vEnum[$__g_vEnum[0][0]][1] = $aResult[2]
Return 1
EndFunc
Func __FatalExit($iCode, $sText = '')
If $sText Then MsgBox($MB_SYSTEMMODAL, 'AutoIt', $sText)
DllCall('kernel32.dll', 'none', 'FatalExit', 'int', $iCode)
EndFunc
Func __Inc(ByRef $aData, $iIncrement = 100)
Select
Case UBound($aData, $UBOUND_COLUMNS)
If $iIncrement < 0 Then
ReDim $aData[$aData[0][0] + 1][UBound($aData, $UBOUND_COLUMNS)]
Else
$aData[0][0] += 1
If $aData[0][0] > UBound($aData) - 1 Then
ReDim $aData[$aData[0][0] + $iIncrement][UBound($aData, $UBOUND_COLUMNS)]
EndIf
EndIf
Case UBound($aData, $UBOUND_ROWS)
If $iIncrement < 0 Then
ReDim $aData[$aData[0] + 1]
Else
$aData[0] += 1
If $aData[0] > UBound($aData) - 1 Then
ReDim $aData[$aData[0] + $iIncrement]
EndIf
EndIf
Case Else
Return 0
EndSelect
Return 1
EndFunc
Func __RGB($iColor)
If $__g_iRGBMode Then
$iColor = _WinAPI_SwitchColor($iColor)
EndIf
Return $iColor
EndFunc
Func __WINVER()
Local $tOSVI = DllStructCreate($tagOSVERSIONINFO)
DllStructSetData($tOSVI, 1, DllStructGetSize($tOSVI))
Local $aRet = DllCall('kernel32.dll', 'bool', 'GetVersionExW', 'struct*', $tOSVI)
If @error Or Not $aRet[0] Then Return SetError(@error, @extended, 0)
Return BitOR(BitShift(DllStructGetData($tOSVI, 2), -8), DllStructGetData($tOSVI, 3))
EndFunc
#EndRegion Internal Functions
#Region Global Variables and Constants
Global Const $DUPLICATE_CLOSE_SOURCE = 0x00000001
Global Const $DUPLICATE_SAME_ACCESS = 0x00000002
Global Const $OBJ_BITMAP = 7
Global Const $OBJ_BRUSH = 2
Global Const $OBJ_COLORSPACE = 14
Global Const $OBJ_DC = 3
Global Const $OBJ_ENHMETADC = 12
Global Const $OBJ_ENHMETAFILE = 13
Global Const $OBJ_EXTPEN = 11
Global Const $OBJ_FONT = 6
Global Const $OBJ_MEMDC = 10
Global Const $OBJ_METADC = 4
Global Const $OBJ_METAFILE = 9
Global Const $OBJ_PAL = 5
Global Const $OBJ_PEN = 1
Global Const $OBJ_REGION = 8
Global Const $NULL_BRUSH = 5
Global Const $NULL_PEN = 8
Global Const $BLACK_BRUSH = 4
Global Const $DKGRAY_BRUSH = 3
Global Const $DC_BRUSH = 18
Global Const $GRAY_BRUSH = 2
Global Const $HOLLOW_BRUSH = $NULL_BRUSH
Global Const $LTGRAY_BRUSH = 1
Global Const $WHITE_BRUSH = 0
Global Const $BLACK_PEN = 7
Global Const $DC_PEN = 19
Global Const $WHITE_PEN = 6
Global Const $ANSI_FIXED_FONT = 11
Global Const $ANSI_VAR_FONT = 12
Global Const $DEVICE_DEFAULT_FONT = 14
Global Const $DEFAULT_GUI_FONT = 17
Global Const $OEM_FIXED_FONT = 10
Global Const $SYSTEM_FONT = 13
Global Const $SYSTEM_FIXED_FONT = 16
Global Const $DEFAULT_PALETTE = 15
#EndRegion Global Variables and Constants
#Region Functions list
#EndRegion Functions list
#Region Public Functions
Func _WinAPI_CloseHandle($hObject)
Local $aResult = DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $hObject)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _WinAPI_DeleteObject($hObject)
Local $aResult = DllCall("gdi32.dll", "bool", "DeleteObject", "handle", $hObject)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _WinAPI_DuplicateHandle($hSourceProcessHandle, $hSourceHandle, $hTargetProcessHandle, $iDesiredAccess, $iInheritHandle, $iOptions)
Local $aResult = DllCall("kernel32.dll", "bool", "DuplicateHandle",  "handle", $hSourceProcessHandle,  "handle", $hSourceHandle,  "handle", $hTargetProcessHandle,  "handle*", 0,  "dword", $iDesiredAccess,  "bool", $iInheritHandle,  "dword", $iOptions)
If @error Or Not $aResult[0] Then Return SetError(@error, @extended, 0)
Return $aResult[4]
EndFunc
Func _WinAPI_GetCurrentObject($hDC, $iType)
Local $aRet = DllCall('gdi32.dll', 'handle', 'GetCurrentObject', 'handle', $hDC, 'uint', $iType)
If @error Or Not $aRet[0] Then Return SetError(@error, @extended, 0)
Return $aRet[0]
EndFunc
Func _WinAPI_GetCurrentProcess()
Local $aResult = DllCall("kernel32.dll", "handle", "GetCurrentProcess")
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_GetObject($hObject, $iSize, $pObject)
Local $aResult = DllCall("gdi32.dll", "int", "GetObjectW", "handle", $hObject, "int", $iSize, "struct*", $pObject)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_GetObjectInfoByHandle($hObject)
Local $tagPUBLIC_OBJECT_BASIC_INFORMATION = 'ulong Attributes;ulong GrantedAcess;ulong HandleCount;ulong PointerCount;ulong Reserved[10]'
Local $tPOBI = DllStructCreate($tagPUBLIC_OBJECT_BASIC_INFORMATION)
Local $aRet = DllCall('ntdll.dll', 'long', 'ZwQueryObject', 'handle', $hObject, 'uint', 0, 'struct*', $tPOBI,  'ulong', DllStructGetSize($tPOBI), 'ptr', 0)
If @error Then Return SetError(@error, @extended, 0)
If $aRet[0] Then Return SetError(10, $aRet[0], 0)
Local $aResult[4]
For $i = 0 To 3
$aResult[$i] = DllStructGetData($tPOBI, $i + 1)
Next
Return $aResult
EndFunc
Func _WinAPI_GetObjectNameByHandle($hObject)
Local $tagUNICODE_STRING = 'struct;ushort Length;ushort MaximumLength;ptr Buffer;endstruct'
Local $tagPUBLIC_OBJECT_TYPE_INFORMATION = 'struct;' & $tagUNICODE_STRING & ';ulong Reserved[22];endstruct'
Local $tPOTI = DllStructCreate($tagPUBLIC_OBJECT_TYPE_INFORMATION & ';byte[32]')
Local $aRet = DllCall('ntdll.dll', 'long', 'ZwQueryObject', 'handle', $hObject, 'uint', 2, 'struct*', $tPOTI,  'ulong', DllStructGetSize($tPOTI), 'ulong*', 0)
If @error Then Return SetError(@error, @extended, '')
If $aRet[0] Then Return SetError(10, $aRet[0], '')
Local $pData = DllStructGetData($tPOTI, 3)
If Not $pData Then Return SetError(11, 0, '')
Return _WinAPI_GetString($pData)
EndFunc
Func _WinAPI_GetObjectType($hObject)
Local $aRet = DllCall('gdi32.dll', 'dword', 'GetObjectType', 'handle', $hObject)
If @error Then Return SetError(@error, @extended, 0)
Return $aRet[0]
EndFunc
Func _WinAPI_GetStdHandle($iStdHandle)
If $iStdHandle < 0 Or $iStdHandle > 2 Then Return SetError(2, 0, -1)
Local Const $aHandle[3] = [-10, -11, -12]
Local $aResult = DllCall("kernel32.dll", "handle", "GetStdHandle", "dword", $aHandle[$iStdHandle])
If @error Then Return SetError(@error, @extended, -1)
Return $aResult[0]
EndFunc
Func _WinAPI_GetStockObject($iObject)
Local $aResult = DllCall("gdi32.dll", "handle", "GetStockObject", "int", $iObject)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_SelectObject($hDC, $hGDIObj)
Local $aResult = DllCall("gdi32.dll", "handle", "SelectObject", "handle", $hDC, "handle", $hGDIObj)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _WinAPI_SetHandleInformation($hObject, $iMask, $iFlags)
Local $aResult = DllCall("kernel32.dll", "bool", "SetHandleInformation", "handle", $hObject, "dword", $iMask, "dword", $iFlags)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
#EndRegion Public Functions
Global Const $LINGUISTIC_IGNORECASE = 0x00000010
Global Const $LINGUISTIC_IGNOREDIACRITIC = 0x00000020
Global Const $NORM_IGNORECASE = 0x00000001
Global Const $NORM_IGNOREKANATYPE = 0x00010000
Global Const $NORM_IGNORENONSPACE = 0x00000002
Global Const $NORM_IGNORESYMBOLS = 0x00000004
Global Const $NORM_IGNOREWIDTH = 0x00020000
Global Const $NORM_LINGUISTIC_CASING = 0x08000000
Global Const $SORT_DIGITSASNUMBERS = 0x00000008
Global Const $SORT_STRINGSORT = 0x00001000
Global Const $CSTR_LESS_THAN = 1
Global Const $CSTR_EQUAL = 2
Global Const $CSTR_GREATER_THAN = 3
Global Const $MUI_LANGUAGE_ID = 0x0004
Global Const $MUI_LANGUAGE_NAME = 0x0008
Global Const $DATE_AUTOLAYOUT = 0x40
Global Const $DATE_LONGDATE = 0x02
Global Const $DATE_LTRREADING = 0x10
Global Const $DATE_SHORTDATE = 0x01
Global Const $DATE_RTLREADING = 0x20
Global Const $DATE_USE_ALT_CALENDAR = 0x04
Global Const $DATE_YEARMONTH = 0x08
Global Const $GEO_NATION = 1
Global Const $GEO_LATITUDE = 2
Global Const $GEO_LONGITUDE = 3
Global Const $GEO_ISO2 = 4
Global Const $GEO_ISO3 = 5
Global Const $GEO_RFC1766 = 6
Global Const $GEO_LCID = 7
Global Const $GEO_FRIENDLYNAME = 8
Global Const $GEO_OFFICIALNAME = 9
Global Const $GEO_TIMEZONES = 10
Global Const $GEO_OFFICIALLANGUAGES = 11
Global Const $GEO_ISO_UN_NUMBER = 12
Global Const $GEO_PARENT = 13
Global Const $LOCALE_ILANGUAGE = 0x0001
Global Const $LOCALE_SLANGUAGE = 0x0002
Global Const $LOCALE_SENGLANGUAGE = 0x1001
Global Const $LOCALE_SABBREVLANGNAME = 0x0003
Global Const $LOCALE_SNATIVELANGNAME = 0x0004
Global Const $LOCALE_ICOUNTRY = 0x0005
Global Const $LOCALE_SCOUNTRY = 0x0006
Global Const $LOCALE_SENGCOUNTRY = 0x1002
Global Const $LOCALE_SABBREVCTRYNAME = 0x0007
Global Const $LOCALE_SNATIVECTRYNAME = 0x0008
Global Const $LOCALE_IDEFAULTLANGUAGE = 0x0009
Global Const $LOCALE_IDEFAULTCOUNTRY = 0x000A
Global Const $LOCALE_IDEFAULTCODEPAGE = 0x000B
Global Const $LOCALE_IDEFAULTANSICODEPAGE = 0x1004
Global Const $LOCALE_IDEFAULTMACCODEPAGE = 0x1011
Global Const $LOCALE_SLIST = 0x000C
Global Const $LOCALE_IMEASURE = 0x000D
Global Const $LOCALE_SDECIMAL = 0x000E
Global Const $LOCALE_STHOUSAND = 0x000F
Global Const $LOCALE_SGROUPING = 0x0010
Global Const $LOCALE_IDIGITS = 0x0011
Global Const $LOCALE_ILZERO = 0x0012
Global Const $LOCALE_INEGNUMBER = 0x1010
Global Const $LOCALE_SNATIVEDIGITS = 0x0013
Global Const $LOCALE_SCURRENCY = 0x0014
Global Const $LOCALE_SINTLSYMBOL = 0x0015
Global Const $LOCALE_SMONDECIMALSEP = 0x0016
Global Const $LOCALE_SMONTHOUSANDSEP = 0x0017
Global Const $LOCALE_SMONGROUPING = 0x0018
Global Const $LOCALE_ICURRDIGITS = 0x0019
Global Const $LOCALE_IINTLCURRDIGITS = 0x001A
Global Const $LOCALE_ICURRENCY = 0x001B
Global Const $LOCALE_INEGCURR = 0x001C
Global Const $LOCALE_SDATE = 0x001D
Global Const $LOCALE_STIME = 0x001E
Global Const $LOCALE_SSHORTDATE = 0x001F
Global Const $LOCALE_SLONGDATE = 0x0020
Global Const $LOCALE_STIMEFORMAT = 0x1003
Global Const $LOCALE_IDATE = 0x0021
Global Const $LOCALE_ILDATE = 0x0022
Global Const $LOCALE_ITIME = 0x0023
Global Const $LOCALE_ITIMEMARKPOSN = 0x1005
Global Const $LOCALE_ICENTURY = 0x0024
Global Const $LOCALE_ITLZERO = 0x0025
Global Const $LOCALE_IDAYLZERO = 0x0026
Global Const $LOCALE_IMONLZERO = 0x0027
Global Const $LOCALE_S1159 = 0x0028
Global Const $LOCALE_S2359 = 0x0029
Global Const $LOCALE_ICALENDARTYPE = 0x1009
Global Const $LOCALE_IOPTIONALCALENDAR = 0x100B
Global Const $LOCALE_IFIRSTDAYOFWEEK = 0x100C
Global Const $LOCALE_IFIRSTWEEKOFYEAR = 0x100D
Global Const $LOCALE_SDAYNAME1 = 0x002A
Global Const $LOCALE_SDAYNAME2 = 0x002B
Global Const $LOCALE_SDAYNAME3 = 0x002C
Global Const $LOCALE_SDAYNAME4 = 0x002D
Global Const $LOCALE_SDAYNAME5 = 0x002E
Global Const $LOCALE_SDAYNAME6 = 0x002F
Global Const $LOCALE_SDAYNAME7 = 0x0030
Global Const $LOCALE_SABBREVDAYNAME1 = 0x0031
Global Const $LOCALE_SABBREVDAYNAME2 = 0x0032
Global Const $LOCALE_SABBREVDAYNAME3 = 0x0033
Global Const $LOCALE_SABBREVDAYNAME4 = 0x0034
Global Const $LOCALE_SABBREVDAYNAME5 = 0x0035
Global Const $LOCALE_SABBREVDAYNAME6 = 0x0036
Global Const $LOCALE_SABBREVDAYNAME7 = 0x0037
Global Const $LOCALE_SMONTHNAME1 = 0x0038
Global Const $LOCALE_SMONTHNAME2 = 0x0039
Global Const $LOCALE_SMONTHNAME3 = 0x003A
Global Const $LOCALE_SMONTHNAME4 = 0x003B
Global Const $LOCALE_SMONTHNAME5 = 0x003C
Global Const $LOCALE_SMONTHNAME6 = 0x003D
Global Const $LOCALE_SMONTHNAME7 = 0x003E
Global Const $LOCALE_SMONTHNAME8 = 0x003F
Global Const $LOCALE_SMONTHNAME9 = 0x0040
Global Const $LOCALE_SMONTHNAME10 = 0x0041
Global Const $LOCALE_SMONTHNAME11 = 0x0042
Global Const $LOCALE_SMONTHNAME12 = 0x0043
Global Const $LOCALE_SMONTHNAME13 = 0x100E
Global Const $LOCALE_SABBREVMONTHNAME1 = 0x0044
Global Const $LOCALE_SABBREVMONTHNAME2 = 0x0045
Global Const $LOCALE_SABBREVMONTHNAME3 = 0x0046
Global Const $LOCALE_SABBREVMONTHNAME4 = 0x0047
Global Const $LOCALE_SABBREVMONTHNAME5 = 0x0048
Global Const $LOCALE_SABBREVMONTHNAME6 = 0x0049
Global Const $LOCALE_SABBREVMONTHNAME7 = 0x004A
Global Const $LOCALE_SABBREVMONTHNAME8 = 0x004B
Global Const $LOCALE_SABBREVMONTHNAME9 = 0x004C
Global Const $LOCALE_SABBREVMONTHNAME10 = 0x004D
Global Const $LOCALE_SABBREVMONTHNAME11 = 0x004E
Global Const $LOCALE_SABBREVMONTHNAME12 = 0x004F
Global Const $LOCALE_SABBREVMONTHNAME13 = 0x100F
Global Const $LOCALE_SPOSITIVESIGN = 0x0050
Global Const $LOCALE_SNEGATIVESIGN = 0x0051
Global Const $LOCALE_IPOSSIGNPOSN = 0x0052
Global Const $LOCALE_INEGSIGNPOSN = 0x0053
Global Const $LOCALE_IPOSSYMPRECEDES = 0x0054
Global Const $LOCALE_IPOSSEPBYSPACE = 0x0055
Global Const $LOCALE_INEGSYMPRECEDES = 0x0056
Global Const $LOCALE_INEGSEPBYSPACE = 0x0057
Global Const $LOCALE_FONTSIGNATURE = 0x0058
Global Const $LOCALE_SISO639LANGNAME = 0x0059
Global Const $LOCALE_SISO3166CTRYNAME = 0x005A
Global Const $LOCALE_IDEFAULTEBCDICCODEPAGE = 0x1012
Global Const $LOCALE_IPAPERSIZE = 0x100A
Global Const $LOCALE_SENGCURRNAME = 0x1007
Global Const $LOCALE_SNATIVECURRNAME = 0x1008
Global Const $LOCALE_SYEARMONTH = 0x1006
Global Const $LOCALE_SSORTNAME = 0x1013
Global Const $LOCALE_IDIGITSUBSTITUTION = 0x1014
Global Const $LOCALE_CUSTOM_DEFAULT = 0x0C00
Global Const $LOCALE_CUSTOM_UI_DEFAULT = 0x1400
Global Const $LOCALE_CUSTOM_UNSPECIFIED = 0x1000
Global Const $LOCALE_INVARIANT = 0x007F
Global Const $LOCALE_SYSTEM_DEFAULT = 0x0800
Global Const $LOCALE_USER_DEFAULT = 0x0400
Global Const $TIME_FORCE24HOURFORMAT = 0x08
Global Const $TIME_NOMINUTESORSECONDS = 0x01
Global Const $TIME_NOSECONDS = 0x02
Global Const $TIME_NOTIMEMARKER = 0x04
Global Const $LCID_INSTALLED = 1
Global Const $LCID_SUPPORTED = 2
#Region Global Variables and Constants
Global Const $tagNUMBERFMT = 'uint NumDigits;uint LeadingZero;uint Grouping;ptr DecimalSep;ptr ThousandSep;uint NegativeOrder'
#EndRegion Global Variables and Constants
#Region Functions list
#EndRegion Functions list
#Region Public Functions
Func _WinAPI_CompareString($iLCID, $sString1, $sString2, $iFlags = 0)
Local $aRet = DllCall('kernel32.dll', 'int', 'CompareStringW', 'dword', $iLCID, 'dword', $iFlags, 'wstr', $sString1,  'int', -1, 'wstr', $sString2, 'int', -1)
If @error Or Not $aRet[0] Then Return SetError(@error, @extended, 0)
Return $aRet[0]
EndFunc
Func _WinAPI_CreateNumberFormatInfo($iNumDigits, $iLeadingZero, $iGrouping, $sDecimalSep, $sThousandSep, $iNegativeOrder)
Local $tFMT = DllStructCreate($tagNUMBERFMT & ';wchar[' & (StringLen($sDecimalSep) + 1) & '];wchar[' & (StringLen($sThousandSep) + 1) & ']')
DllStructSetData($tFMT, 1, $iNumDigits)
DllStructSetData($tFMT, 2, $iLeadingZero)
DllStructSetData($tFMT, 3, $iGrouping)
DllStructSetData($tFMT, 4, DllStructGetPtr($tFMT, 7))
DllStructSetData($tFMT, 5, DllStructGetPtr($tFMT, 8))
DllStructSetData($tFMT, 6, $iNegativeOrder)
DllStructSetData($tFMT, 7, $sDecimalSep)
DllStructSetData($tFMT, 8, $sThousandSep)
Return $tFMT
EndFunc
Func _WinAPI_EnumSystemGeoID()
Local $hEnumProc = DllCallbackRegister('__EnumGeoIDProc', 'bool', 'long')
Dim $__g_vEnum[101] = [0]
Local $aRet = DllCall('kernel32.dll', 'bool', 'EnumSystemGeoID', 'dword', 16, 'long', 0, 'ptr', DllCallbackGetPtr($hEnumProc))
If @error Or Not $aRet[0] Or Not $__g_vEnum[0] Then
$__g_vEnum = @error + 10
EndIf
DllCallbackFree($hEnumProc)
If $__g_vEnum Then Return SetError($__g_vEnum, 0, 0)
__Inc($__g_vEnum, -1)
Return $__g_vEnum
EndFunc
Func _WinAPI_EnumSystemLocales($iFlag)
Local $hEnumProc = DllCallbackRegister('__EnumLocalesProc', 'bool', 'ptr')
Dim $__g_vEnum[101] = [0]
Local $aRet = DllCall('kernel32.dll', 'bool', 'EnumSystemLocalesW', 'ptr', DllCallbackGetPtr($hEnumProc), 'dword', $iFlag)
If @error Or Not $aRet[0] Or Not $__g_vEnum[0] Then
$__g_vEnum = @error + 10
EndIf
DllCallbackFree($hEnumProc)
If $__g_vEnum Then Return SetError($__g_vEnum, 0, 0)
__Inc($__g_vEnum, -1)
Return $__g_vEnum
EndFunc
Func _WinAPI_EnumUILanguages($iFlag = 0)
Local $hEnumProc = DllCallbackRegister('__EnumUILanguagesProc', 'bool', 'ptr;long_ptr')
Local $iID = 1
If $__WINVER >= 0x0600 Then
If BitAND($iFlag, 0x0008) Then
$iID = 0
EndIf
Else
$iFlag = 0
EndIf
Dim $__g_vEnum[101] = [0]
Local $aRet = DllCall('kernel32.dll', 'bool', 'EnumUILanguagesW', 'ptr', DllCallbackGetPtr($hEnumProc), 'dword', $iFlag,  'long_ptr', $iID)
If @error Or Not $aRet[0] Or Not $__g_vEnum[0] Then
$__g_vEnum = @error + 10
EndIf
DllCallbackFree($hEnumProc)
If $__g_vEnum Then Return SetError($__g_vEnum, 0, 0)
__Inc($__g_vEnum, -1)
Return $__g_vEnum
EndFunc
Func _WinAPI_GetDateFormat($iLCID = 0, $tSYSTEMTIME = 0, $iFlags = 0, $sFormat = '')
If Not $iLCID Then $iLCID = 0x0400
Local $sTypeOfFormat = 'wstr'
If Not StringStripWS($sFormat, $STR_STRIPLEADING + $STR_STRIPTRAILING) Then
$sTypeOfFormat = 'ptr'
$sFormat = 0
EndIf
Local $aRet = DllCall('kernel32.dll', 'int', 'GetDateFormatW', 'dword', $iLCID, 'dword', $iFlags, 'struct*', $tSYSTEMTIME,  $sTypeOfFormat, $sFormat, 'wstr', '', 'int', 2048)
If @error Or Not $aRet[0] Then Return SetError(@error, @extended, '')
Return $aRet[5]
EndFunc
Func _WinAPI_GetDurationFormat($iLCID, $iDuration, $sFormat = '')
If Not $iLCID Then $iLCID = 0x0400
Local $pST, $iVal
If IsDllStruct($iDuration) Then
$pST = DllStructGetPtr($iDuration)
$iVal = 0
Else
$pST = 0
$iVal = $iDuration
EndIf
Local $sTypeOfFormat = 'wstr'
If Not StringStripWS($sFormat, $STR_STRIPLEADING + $STR_STRIPTRAILING) Then
$sTypeOfFormat = 'ptr'
$sFormat = 0
EndIf
Local $aRet = DllCall('kernel32.dll', 'int', 'GetDurationFormat', 'dword', $iLCID, 'dword', 0, 'ptr', $pST, 'uint64', $iVal,  $sTypeOfFormat, $sFormat, 'wstr', '', 'int', 2048)
If @error Or Not $aRet[0] Then Return SetError(@error, @extended, '')
Return $aRet[6]
EndFunc
Func _WinAPI_GetGeoInfo($iGEOID, $iType, $iLanguage = 0)
Local $aRet = DllCall('kernel32.dll', 'int', 'GetGeoInfoW', 'long', $iGEOID, 'dword', $iType, 'wstr', '', 'int', 4096,  'word', $iLanguage)
If @error Or Not $aRet[0] Then Return SetError(@error, @extended, 0)
Return $aRet[3]
EndFunc
Func _WinAPI_GetLocaleInfo($iLCID, $iType)
Local $aRet = DllCall('kernel32.dll', 'int', 'GetLocaleInfoW', 'dword', $iLCID, 'dword', $iType, 'wstr', '', 'int', 2048)
If @error Or Not $aRet[0] Then Return SetError(@error + 10, @extended, '')
Return $aRet[3]
EndFunc
Func _WinAPI_GetNumberFormat($iLCID, $sNumber, $tNUMBERFMT = 0)
If Not $iLCID Then $iLCID = 0x0400
Local $aRet = DllCall('kernel32.dll', 'int', 'GetNumberFormatW', 'dword', $iLCID, 'dword', 0, 'wstr', $sNumber,  'struct*', $tNUMBERFMT, 'wstr', '', 'int', 2048)
If @error Or Not $aRet[0] Then Return SetError(@error, @extended, '')
Return $aRet[5]
EndFunc
Func _WinAPI_GetSystemDefaultLangID()
Local $aRet = DllCall('kernel32.dll', 'word', 'GetSystemDefaultLangID')
If @error Then Return SetError(@error, @extended, 0)
Return $aRet[0]
EndFunc
Func _WinAPI_GetSystemDefaultLCID()
Local $aRet = DllCall('kernel32.dll', 'dword', 'GetSystemDefaultLCID')
If @error Then Return SetError(@error, @extended, 0)
Return $aRet[0]
EndFunc
Func _WinAPI_GetSystemDefaultUILanguage()
Local $aRet = DllCall('kernel32.dll', 'word', 'GetSystemDefaultUILanguage')
If @error Then Return SetError(@error, @extended, 0)
Return $aRet[0]
EndFunc
Func _WinAPI_GetThreadLocale()
Local $aRet = DllCall('kernel32.dll', 'dword', 'GetThreadLocale')
If @error Then Return SetError(@error, @extended, 0)
Return $aRet[0]
EndFunc
Func _WinAPI_GetThreadUILanguage()
Local $aRet = DllCall('kernel32.dll', 'word', 'GetThreadUILanguage')
If @error Then Return SetError(@error, @extended, 0)
Return $aRet[0]
EndFunc
Func _WinAPI_GetTimeFormat($iLCID = 0, $tSYSTEMTIME = 0, $iFlags = 0, $sFormat = '')
If Not $iLCID Then $iLCID = 0x0400
Local $sTypeOfFormat = 'wstr'
If Not StringStripWS($sFormat, $STR_STRIPLEADING + $STR_STRIPTRAILING) Then
$sTypeOfFormat = 'ptr'
$sFormat = 0
EndIf
Local $aRet = DllCall('kernel32.dll', 'int', 'GetTimeFormatW', 'dword', $iLCID, 'dword', $iFlags, 'struct*', $tSYSTEMTIME,  $sTypeOfFormat, $sFormat, 'wstr', '', 'int', 2048)
If @error Or Not $aRet[0] Then Return SetError(@error + 10, @extended, '')
Return $aRet[5]
EndFunc
Func _WinAPI_GetUserDefaultLangID()
Local $aRet = DllCall('kernel32.dll', 'word', 'GetUserDefaultLangID')
If @error Then Return SetError(@error, @extended, 0)
Return $aRet[0]
EndFunc
Func _WinAPI_GetUserDefaultLCID()
Local $aRet = DllCall('kernel32.dll', 'dword', 'GetUserDefaultLCID')
If @error Then Return SetError(@error, @extended, 0)
Return $aRet[0]
EndFunc
Func _WinAPI_GetUserDefaultUILanguage()
Local $aRet = DllCall('kernel32.dll', 'word', 'GetUserDefaultUILanguage')
If @error Then Return SetError(@error, @extended, 0)
Return $aRet[0]
EndFunc
Func _WinAPI_GetUserGeoID()
Local $aRet = DllCall('kernel32.dll', 'long', 'GetUserGeoID', 'uint', 16)
If @error Then Return SetError(@error, @extended, -1)
Return $aRet[0]
EndFunc
Func _WinAPI_IsValidLocale($iLCID, $iFlag = 0)
Local $aRet = DllCall('kernel32.dll', 'bool', 'IsValidLocale', 'dword', $iLCID, 'dword', $iFlag)
If @error Then Return SetError(@error, @extended, False)
Return $aRet[0]
EndFunc
Func _WinAPI_SetLocaleInfo($iLCID, $iType, $sData)
Local $aRet = DllCall('kernel32.dll', 'bool', 'SetLocaleInfoW', 'dword', $iLCID, 'dword', $iType, 'wstr', $sData)
If @error Then Return SetError(@error, @extended, False)
Return $aRet[0]
EndFunc
Func _WinAPI_SetThreadLocale($iLCID)
Local $aRet = DllCall('kernel32.dll', 'bool', 'SetThreadLocale', 'dword', $iLCID)
If @error Then Return SetError(@error, @extended, False)
Return $aRet[0]
EndFunc
Func _WinAPI_SetThreadUILanguage($iLanguage)
Local $aRet = DllCall('kernel32.dll', 'word', 'SetThreadUILanguage', 'word', $iLanguage)
If @error Then Return SetError(@error, @extended, False)
Return ($aRet[0] = $aRet[1])
EndFunc
Func _WinAPI_SetUserGeoID($iGEOID)
Local $aRet = DllCall('kernel32.dll', 'bool', 'SetUserGeoID', 'long', $iGEOID)
If @error Then Return SetError(@error, @extended, False)
Return $aRet[0]
EndFunc
#EndRegion Public Functions
#Region Internal Functions
Func __EnumGeoIDProc($iID)
__Inc($__g_vEnum)
$__g_vEnum[$__g_vEnum[0]] = $iID
Return 1
EndFunc
Func __EnumLocalesProc($pLocale)
__Inc($__g_vEnum)
$__g_vEnum[$__g_vEnum[0]] = Dec(DllStructGetData(DllStructCreate('wchar[' & (_WinAPI_StrLen($pLocale) + 1) & ']', $pLocale), 1))
Return 1
EndFunc
Func __EnumUILanguagesProc($pLanguage, $iID)
__Inc($__g_vEnum)
$__g_vEnum[$__g_vEnum[0]] = DllStructGetData(DllStructCreate('wchar[' & (_WinAPI_StrLen($pLanguage) + 1) & ']', $pLanguage), 1)
If $iID Then
$__g_vEnum[$__g_vEnum[0]] = Dec($__g_vEnum[$__g_vEnum[0]])
EndIf
Return 1
EndFunc
#EndRegion Internal Functions
Func _DateAdd($sType, $iNumber, $sDate)
Local $asTimePart[4]
Local $asDatePart[4]
Local $iJulianDate
$sType = StringLeft($sType, 1)
If StringInStr("D,M,Y,w,h,n,s", $sType) = 0 Or $sType = "" Then
Return SetError(1, 0, 0)
EndIf
If Not StringIsInt($iNumber) Then
Return SetError(2, 0, 0)
EndIf
If Not _DateIsValid($sDate) Then
Return SetError(3, 0, 0)
EndIf
_DateTimeSplit($sDate, $asDatePart, $asTimePart)
If $sType = "d" Or $sType = "w" Then
If $sType = "w" Then $iNumber = $iNumber * 7
$iJulianDate = _DateToDayValue($asDatePart[1], $asDatePart[2], $asDatePart[3]) + $iNumber
_DayValueToDate($iJulianDate, $asDatePart[1], $asDatePart[2], $asDatePart[3])
EndIf
If $sType = "m" Then
$asDatePart[2] = $asDatePart[2] + $iNumber
While $asDatePart[2] > 12
$asDatePart[2] = $asDatePart[2] - 12
$asDatePart[1] = $asDatePart[1] + 1
WEnd
While $asDatePart[2] < 1
$asDatePart[2] = $asDatePart[2] + 12
$asDatePart[1] = $asDatePart[1] - 1
WEnd
EndIf
If $sType = "y" Then
$asDatePart[1] = $asDatePart[1] + $iNumber
EndIf
If $sType = "h" Or $sType = "n" Or $sType = "s" Then
Local $iTimeVal = _TimeToTicks($asTimePart[1], $asTimePart[2], $asTimePart[3]) / 1000
If $sType = "h" Then $iTimeVal = $iTimeVal + $iNumber * 3600
If $sType = "n" Then $iTimeVal = $iTimeVal + $iNumber * 60
If $sType = "s" Then $iTimeVal = $iTimeVal + $iNumber
Local $iDay2Add = Int($iTimeVal / (24 * 60 * 60))
$iTimeVal = $iTimeVal - $iDay2Add * 24 * 60 * 60
If $iTimeVal < 0 Then
$iDay2Add = $iDay2Add - 1
$iTimeVal = $iTimeVal + 24 * 60 * 60
EndIf
$iJulianDate = _DateToDayValue($asDatePart[1], $asDatePart[2], $asDatePart[3]) + $iDay2Add
_DayValueToDate($iJulianDate, $asDatePart[1], $asDatePart[2], $asDatePart[3])
_TicksToTime($iTimeVal * 1000, $asTimePart[1], $asTimePart[2], $asTimePart[3])
EndIf
Local $iNumDays = _DaysInMonth($asDatePart[1])
If $iNumDays[$asDatePart[2]] < $asDatePart[3] Then $asDatePart[3] = $iNumDays[$asDatePart[2]]
$sDate = $asDatePart[1] & '/' & StringRight("0" & $asDatePart[2], 2) & '/' & StringRight("0" & $asDatePart[3], 2)
If $asTimePart[0] > 0 Then
If $asTimePart[0] > 2 Then
$sDate = $sDate & " " & StringRight("0" & $asTimePart[1], 2) & ':' & StringRight("0" & $asTimePart[2], 2) & ':' & StringRight("0" & $asTimePart[3], 2)
Else
$sDate = $sDate & " " & StringRight("0" & $asTimePart[1], 2) & ':' & StringRight("0" & $asTimePart[2], 2)
EndIf
EndIf
Return $sDate
EndFunc
Func _DateDayOfWeek($iDayNum, $iFormat = Default)
Local Const $MONDAY_IS_NO1 = 128
If $iFormat = Default Then $iFormat = 0
$iDayNum = Int($iDayNum)
If $iDayNum < 1 Or $iDayNum > 7 Then Return SetError(1, 0, "")
Local $tSYSTEMTIME = DllStructCreate($tagSYSTEMTIME)
DllStructSetData($tSYSTEMTIME, "Year", BitAND($iFormat, $MONDAY_IS_NO1) ? 2007 : 2006)
DllStructSetData($tSYSTEMTIME, "Month", 1)
DllStructSetData($tSYSTEMTIME, "Day", $iDayNum)
Return _WinAPI_GetDateFormat(BitAND($iFormat, $DMW_LOCALE_LONGNAME) ? $LOCALE_USER_DEFAULT : $LOCALE_INVARIANT, $tSYSTEMTIME, 0, BitAND($iFormat, $DMW_SHORTNAME) ? "ddd" : "dddd")
EndFunc
Func _DateDaysInMonth($iYear, $iMonthNum)
$iMonthNum = Int($iMonthNum)
$iYear = Int($iYear)
Return __DateIsMonth($iMonthNum) And __DateIsYear($iYear) ? _DaysInMonth($iYear)[$iMonthNum] : SetError(1, 0, 0)
EndFunc
Func _DateDiff($sType, $sStartDate, $sEndDate)
$sType = StringLeft($sType, 1)
If StringInStr("d,m,y,w,h,n,s", $sType) = 0 Or $sType = "" Then
Return SetError(1, 0, 0)
EndIf
If Not _DateIsValid($sStartDate) Then
Return SetError(2, 0, 0)
EndIf
If Not _DateIsValid($sEndDate) Then
Return SetError(3, 0, 0)
EndIf
Local $asStartDatePart[4], $asStartTimePart[4], $asEndDatePart[4], $asEndTimePart[4]
_DateTimeSplit($sStartDate, $asStartDatePart, $asStartTimePart)
_DateTimeSplit($sEndDate, $asEndDatePart, $asEndTimePart)
Local $aDaysDiff = _DateToDayValue($asEndDatePart[1], $asEndDatePart[2], $asEndDatePart[3]) - _DateToDayValue($asStartDatePart[1], $asStartDatePart[2], $asStartDatePart[3])
Local $iTimeDiff, $iYearDiff, $iStartTimeInSecs, $iEndTimeInSecs
If $asStartTimePart[0] > 1 And $asEndTimePart[0] > 1 Then
$iStartTimeInSecs = $asStartTimePart[1] * 3600 + $asStartTimePart[2] * 60 + $asStartTimePart[3]
$iEndTimeInSecs = $asEndTimePart[1] * 3600 + $asEndTimePart[2] * 60 + $asEndTimePart[3]
$iTimeDiff = $iEndTimeInSecs - $iStartTimeInSecs
If $iTimeDiff < 0 Then
$aDaysDiff = $aDaysDiff - 1
$iTimeDiff = $iTimeDiff + 24 * 60 * 60
EndIf
Else
$iTimeDiff = 0
EndIf
Select
Case $sType = "d"
Return $aDaysDiff
Case $sType = "m"
$iYearDiff = $asEndDatePart[1] - $asStartDatePart[1]
Local $iMonthDiff = $asEndDatePart[2] - $asStartDatePart[2] + $iYearDiff * 12
If $asEndDatePart[3] < $asStartDatePart[3] Then $iMonthDiff = $iMonthDiff - 1
$iStartTimeInSecs = $asStartTimePart[1] * 3600 + $asStartTimePart[2] * 60 + $asStartTimePart[3]
$iEndTimeInSecs = $asEndTimePart[1] * 3600 + $asEndTimePart[2] * 60 + $asEndTimePart[3]
$iTimeDiff = $iEndTimeInSecs - $iStartTimeInSecs
If $asEndDatePart[3] = $asStartDatePart[3] And $iTimeDiff < 0 Then $iMonthDiff = $iMonthDiff - 1
Return $iMonthDiff
Case $sType = "y"
$iYearDiff = $asEndDatePart[1] - $asStartDatePart[1]
If $asEndDatePart[2] < $asStartDatePart[2] Then $iYearDiff = $iYearDiff - 1
If $asEndDatePart[2] = $asStartDatePart[2] And $asEndDatePart[3] < $asStartDatePart[3] Then $iYearDiff = $iYearDiff - 1
$iStartTimeInSecs = $asStartTimePart[1] * 3600 + $asStartTimePart[2] * 60 + $asStartTimePart[3]
$iEndTimeInSecs = $asEndTimePart[1] * 3600 + $asEndTimePart[2] * 60 + $asEndTimePart[3]
$iTimeDiff = $iEndTimeInSecs - $iStartTimeInSecs
If $asEndDatePart[2] = $asStartDatePart[2] And $asEndDatePart[3] = $asStartDatePart[3] And $iTimeDiff < 0 Then $iYearDiff = $iYearDiff - 1
Return $iYearDiff
Case $sType = "w"
Return Int($aDaysDiff / 7)
Case $sType = "h"
Return $aDaysDiff * 24 + Int($iTimeDiff / 3600)
Case $sType = "n"
Return $aDaysDiff * 24 * 60 + Int($iTimeDiff / 60)
Case $sType = "s"
Return $aDaysDiff * 24 * 60 * 60 + $iTimeDiff
EndSelect
EndFunc
Func _DateIsLeapYear($iYear)
If StringIsInt($iYear) Then
Select
Case Mod($iYear, 4) = 0 And Mod($iYear, 100) <> 0
Return 1
Case Mod($iYear, 400) = 0
Return 1
Case Else
Return 0
EndSelect
EndIf
Return SetError(1, 0, 0)
EndFunc
Func __DateIsMonth($iNumber)
$iNumber = Int($iNumber)
Return $iNumber >= 1 And $iNumber <= 12
EndFunc
Func _DateIsValid($sDate)
Local $asDatePart[4], $asTimePart[4]
_DateTimeSplit($sDate, $asDatePart, $asTimePart)
If Not StringIsInt($asDatePart[1]) Then Return 0
If Not StringIsInt($asDatePart[2]) Then Return 0
If Not StringIsInt($asDatePart[3]) Then Return 0
$asDatePart[1] = Int($asDatePart[1])
$asDatePart[2] = Int($asDatePart[2])
$asDatePart[3] = Int($asDatePart[3])
Local $iNumDays = _DaysInMonth($asDatePart[1])
If $asDatePart[1] < 1000 Or $asDatePart[1] > 2999 Then Return 0
If $asDatePart[2] < 1 Or $asDatePart[2] > 12 Then Return 0
If $asDatePart[3] < 1 Or $asDatePart[3] > $iNumDays[$asDatePart[2]] Then Return 0
If $asTimePart[0] < 1 Then Return 1
If $asTimePart[0] < 2 Then Return 0
If $asTimePart[0] = 2 Then $asTimePart[3] = "00"
If Not StringIsInt($asTimePart[1]) Then Return 0
If Not StringIsInt($asTimePart[2]) Then Return 0
If Not StringIsInt($asTimePart[3]) Then Return 0
$asTimePart[1] = Int($asTimePart[1])
$asTimePart[2] = Int($asTimePart[2])
$asTimePart[3] = Int($asTimePart[3])
If $asTimePart[1] < 0 Or $asTimePart[1] > 23 Then Return 0
If $asTimePart[2] < 0 Or $asTimePart[2] > 59 Then Return 0
If $asTimePart[3] < 0 Or $asTimePart[3] > 59 Then Return 0
Return 1
EndFunc
Func __DateIsYear($iNumber)
Return StringLen($iNumber) = 4
EndFunc
Func _DateLastWeekdayNum($iWeekdayNum)
Select
Case Not StringIsInt($iWeekdayNum)
Return SetError(1, 0, 0)
Case $iWeekdayNum < 1 Or $iWeekdayNum > 7
Return SetError(2, 0, 0)
Case Else
Local $iLastWeekdayNum
If $iWeekdayNum = 1 Then
$iLastWeekdayNum = 7
Else
$iLastWeekdayNum = $iWeekdayNum - 1
EndIf
Return $iLastWeekdayNum
EndSelect
EndFunc
Func _DateLastMonthNum($iMonthNum)
Select
Case Not StringIsInt($iMonthNum)
Return SetError(1, 0, 0)
Case Not __DateIsMonth($iMonthNum)
Return SetError(2, 0, 0)
Case Else
Local $iLastMonthNum
If $iMonthNum = 1 Then
$iLastMonthNum = 12
Else
$iLastMonthNum = $iMonthNum - 1
EndIf
$iLastMonthNum = StringFormat("%02d", $iLastMonthNum)
Return $iLastMonthNum
EndSelect
EndFunc
Func _DateLastMonthYear($iMonthNum, $iYear)
Select
Case Not StringIsInt($iMonthNum) Or Not StringIsInt($iYear)
Return SetError(1, 0, 0)
Case Not __DateIsMonth($iMonthNum)
Return SetError(2, 0, 0)
Case Else
Local $iLastYear
If $iMonthNum = 1 Then
$iLastYear = $iYear - 1
Else
$iLastYear = $iYear
EndIf
$iLastYear = StringFormat("%04d", $iLastYear)
Return $iLastYear
EndSelect
EndFunc
Func _DateNextWeekdayNum($iWeekdayNum)
Select
Case Not StringIsInt($iWeekdayNum)
Return SetError(1, 0, 0)
Case $iWeekdayNum < 1 Or $iWeekdayNum > 7
Return SetError(2, 0, 0)
Case Else
Local $iNextWeekdayNum
If $iWeekdayNum = 7 Then
$iNextWeekdayNum = 1
Else
$iNextWeekdayNum = $iWeekdayNum + 1
EndIf
Return $iNextWeekdayNum
EndSelect
EndFunc
Func _DateNextMonthNum($iMonthNum)
Select
Case Not StringIsInt($iMonthNum)
Return SetError(1, 0, 0)
Case Not __DateIsMonth($iMonthNum)
Return SetError(2, 0, 0)
Case Else
Local $iNextMonthNum
If $iMonthNum = 12 Then
$iNextMonthNum = 1
Else
$iNextMonthNum = $iMonthNum + 1
EndIf
$iNextMonthNum = StringFormat("%02d", $iNextMonthNum)
Return $iNextMonthNum
EndSelect
EndFunc
Func _DateNextMonthYear($iMonthNum, $iYear)
Select
Case Not StringIsInt($iMonthNum) Or Not StringIsInt($iYear)
Return SetError(1, 0, 0)
Case Not __DateIsMonth($iMonthNum)
Return SetError(2, 0, 0)
Case Else
Local $iNextYear
If $iMonthNum = 12 Then
$iNextYear = $iYear + 1
Else
$iNextYear = $iYear
EndIf
$iNextYear = StringFormat("%04d", $iNextYear)
Return $iNextYear
EndSelect
EndFunc
Func _DateTimeFormat($sDate, $sType)
Local $asDatePart[4], $asTimePart[4]
Local $sTempDate = "", $sTempTime = ""
Local $sAM, $sPM, $sTempString = ""
If Not _DateIsValid($sDate) Then
Return SetError(1, 0, "")
EndIf
If $sType < 0 Or $sType > 5 Or Not IsInt($sType) Then
Return SetError(2, 0, "")
EndIf
_DateTimeSplit($sDate, $asDatePart, $asTimePart)
Switch $sType
Case 0
$sTempString = _WinAPI_GetLocaleInfo($LOCALE_USER_DEFAULT, $LOCALE_SSHORTDATE)
If Not @error And Not ($sTempString = '') Then
$sTempDate = $sTempString
Else
$sTempDate = "M/d/yyyy"
EndIf
If $asTimePart[0] > 1 Then
$sTempString = _WinAPI_GetLocaleInfo($LOCALE_USER_DEFAULT, $LOCALE_STIMEFORMAT)
If Not @error And Not ($sTempString = '') Then
$sTempTime = $sTempString
Else
$sTempTime = "h:mm:ss tt"
EndIf
EndIf
Case 1
$sTempString = _WinAPI_GetLocaleInfo($LOCALE_USER_DEFAULT, $LOCALE_SLONGDATE)
If Not @error And Not ($sTempString = '') Then
$sTempDate = $sTempString
Else
$sTempDate = "dddd, MMMM dd, yyyy"
EndIf
Case 2
$sTempString = _WinAPI_GetLocaleInfo($LOCALE_USER_DEFAULT, $LOCALE_SSHORTDATE)
If Not @error And Not ($sTempString = '') Then
$sTempDate = $sTempString
Else
$sTempDate = "M/d/yyyy"
EndIf
Case 3
If $asTimePart[0] > 1 Then
$sTempString = _WinAPI_GetLocaleInfo($LOCALE_USER_DEFAULT, $LOCALE_STIMEFORMAT)
If Not @error And Not ($sTempString = '') Then
$sTempTime = $sTempString
Else
$sTempTime = "h:mm:ss tt"
EndIf
EndIf
Case 4
If $asTimePart[0] > 1 Then
$sTempTime = "hh:mm"
EndIf
Case 5
If $asTimePart[0] > 1 Then
$sTempTime = "hh:mm:ss"
EndIf
EndSwitch
If $sTempDate <> "" Then
$sTempString = _WinAPI_GetLocaleInfo($LOCALE_USER_DEFAULT, $LOCALE_SDATE)
If Not @error And Not ($sTempString = '') Then
$sTempDate = StringReplace($sTempDate, "/", $sTempString)
EndIf
Local $iWday = _DateToDayOfWeek($asDatePart[1], $asDatePart[2], $asDatePart[3])
$asDatePart[3] = StringRight("0" & $asDatePart[3], 2)
$asDatePart[2] = StringRight("0" & $asDatePart[2], 2)
$sTempDate = StringReplace($sTempDate, "d", "@")
$sTempDate = StringReplace($sTempDate, "m", "#")
$sTempDate = StringReplace($sTempDate, "y", "&")
$sTempDate = StringReplace($sTempDate, "@@@@", _DateDayOfWeek($iWday, 0))
$sTempDate = StringReplace($sTempDate, "@@@", _DateDayOfWeek($iWday, 1))
$sTempDate = StringReplace($sTempDate, "@@", $asDatePart[3])
$sTempDate = StringReplace($sTempDate, "@", StringReplace(StringLeft($asDatePart[3], 1), "0", "") & StringRight($asDatePart[3], 1))
$sTempDate = StringReplace($sTempDate, "####", _DateToMonth($asDatePart[2], 0))
$sTempDate = StringReplace($sTempDate, "###", _DateToMonth($asDatePart[2], 1))
$sTempDate = StringReplace($sTempDate, "##", $asDatePart[2])
$sTempDate = StringReplace($sTempDate, "#", StringReplace(StringLeft($asDatePart[2], 1), "0", "") & StringRight($asDatePart[2], 1))
$sTempDate = StringReplace($sTempDate, "&&&&", $asDatePart[1])
$sTempDate = StringReplace($sTempDate, "&&", StringRight($asDatePart[1], 2))
EndIf
If $sTempTime <> "" Then
$sTempString = _WinAPI_GetLocaleInfo($LOCALE_USER_DEFAULT, $LOCALE_S1159)
If Not @error And Not ($sTempString = '') Then
$sAM = $sTempString
Else
$sAM = "AM"
EndIf
$sTempString = _WinAPI_GetLocaleInfo($LOCALE_USER_DEFAULT, $LOCALE_S2359)
If Not @error And Not ($sTempString = '') Then
$sPM = $sTempString
Else
$sPM = "PM"
EndIf
$sTempString = _WinAPI_GetLocaleInfo($LOCALE_USER_DEFAULT, $LOCALE_STIME)
If Not @error And Not ($sTempString = '') Then
$sTempTime = StringReplace($sTempTime, ":", $sTempString)
EndIf
If StringInStr($sTempTime, "tt") Then
If $asTimePart[1] < 12 Then
$sTempTime = StringReplace($sTempTime, "tt", $sAM)
If $asTimePart[1] = 0 Then $asTimePart[1] = 12
Else
$sTempTime = StringReplace($sTempTime, "tt", $sPM)
If $asTimePart[1] > 12 Then $asTimePart[1] = $asTimePart[1] - 12
EndIf
EndIf
$asTimePart[1] = StringRight("0" & $asTimePart[1], 2)
$asTimePart[2] = StringRight("0" & $asTimePart[2], 2)
$asTimePart[3] = StringRight("0" & $asTimePart[3], 2)
$sTempTime = StringReplace($sTempTime, "hh", StringFormat("%02d", $asTimePart[1]))
$sTempTime = StringReplace($sTempTime, "h", StringReplace(StringLeft($asTimePart[1], 1), "0", "") & StringRight($asTimePart[1], 1))
$sTempTime = StringReplace($sTempTime, "mm", StringFormat("%02d", $asTimePart[2]))
$sTempTime = StringReplace($sTempTime, "ss", StringFormat("%02d", $asTimePart[3]))
$sTempDate = StringStripWS($sTempDate & " " & $sTempTime, $STR_STRIPLEADING + $STR_STRIPTRAILING)
EndIf
Return $sTempDate
EndFunc
Func _DateTimeSplit($sDate, ByRef $aDatePart, ByRef $iTimePart)
Local $sDateTime = StringSplit($sDate, " T")
If $sDateTime[0] > 0 Then $aDatePart = StringSplit($sDateTime[1], "/-.")
If $sDateTime[0] > 1 Then
$iTimePart = StringSplit($sDateTime[2], ":")
If UBound($iTimePart) < 4 Then ReDim $iTimePart[4]
Else
Dim $iTimePart[4]
EndIf
If UBound($aDatePart) < 4 Then ReDim $aDatePart[4]
For $x = 1 To 3
If StringIsInt($aDatePart[$x]) Then
$aDatePart[$x] = Int($aDatePart[$x])
Else
$aDatePart[$x] = -1
EndIf
If StringIsInt($iTimePart[$x]) Then
$iTimePart[$x] = Int($iTimePart[$x])
Else
$iTimePart[$x] = 0
EndIf
Next
Return 1
EndFunc
Func _DateToDayOfWeek($iYear, $iMonth, $iDay)
If Not _DateIsValid($iYear & "/" & $iMonth & "/" & $iDay) Then
Return SetError(1, 0, "")
EndIf
Local $i_FactorA = Int((14 - $iMonth) / 12)
Local $i_FactorY = $iYear - $i_FactorA
Local $i_FactorM = $iMonth + (12 * $i_FactorA) - 2
Local $i_FactorD = Mod($iDay + $i_FactorY + Int($i_FactorY / 4) - Int($i_FactorY / 100) + Int($i_FactorY / 400) + Int((31 * $i_FactorM) / 12), 7)
Return $i_FactorD + 1
EndFunc
Func _DateToDayOfWeekISO($iYear, $iMonth, $iDay)
Local $iDow = _DateToDayOfWeek($iYear, $iMonth, $iDay)
If @error Then
Return SetError(1, 0, "")
EndIf
If $iDow >= 2 Then Return $iDow - 1
Return 7
EndFunc
Func _DateToDayValue($iYear, $iMonth, $iDay)
If Not _DateIsValid(StringFormat("%04d/%02d/%02d", $iYear, $iMonth, $iDay)) Then
Return SetError(1, 0, "")
EndIf
If $iMonth < 3 Then
$iMonth = $iMonth + 12
$iYear = $iYear - 1
EndIf
Local $i_FactorA = Int($iYear / 100)
Local $i_FactorB = Int($i_FactorA / 4)
Local $i_FactorC = 2 - $i_FactorA + $i_FactorB
Local $i_FactorE = Int(1461 * ($iYear + 4716) / 4)
Local $i_FactorF = Int(153 * ($iMonth + 1) / 5)
Local $iJulianDate = $i_FactorC + $iDay + $i_FactorE + $i_FactorF - 1524.5
Return $iJulianDate
EndFunc
Func _DateToMonth($iMonNum, $iFormat = Default)
If $iFormat = Default Then $iFormat = 0
$iMonNum = Int($iMonNum)
If Not __DateIsMonth($iMonNum) Then Return SetError(1, 0, "")
Local $tSYSTEMTIME = DllStructCreate($tagSYSTEMTIME)
DllStructSetData($tSYSTEMTIME, "Year", @YEAR)
DllStructSetData($tSYSTEMTIME, "Month", $iMonNum)
DllStructSetData($tSYSTEMTIME, "Day", 1)
Return _WinAPI_GetDateFormat(BitAND($iFormat, $DMW_LOCALE_LONGNAME) ? $LOCALE_USER_DEFAULT : $LOCALE_INVARIANT, $tSYSTEMTIME, 0, BitAND($iFormat, $DMW_SHORTNAME) ? "MMM" : "MMMM")
EndFunc
Func _DayValueToDate($iJulianDate, ByRef $iYear, ByRef $iMonth, ByRef $iDay)
If $iJulianDate < 0 Or Not IsNumber($iJulianDate) Then
Return SetError(1, 0, 0)
EndIf
Local $i_FactorZ = Int($iJulianDate + 0.5)
Local $i_FactorW = Int(($i_FactorZ - 1867216.25) / 36524.25)
Local $i_FactorX = Int($i_FactorW / 4)
Local $i_FactorA = $i_FactorZ + 1 + $i_FactorW - $i_FactorX
Local $i_FactorB = $i_FactorA + 1524
Local $i_FactorC = Int(($i_FactorB - 122.1) / 365.25)
Local $i_FactorD = Int(365.25 * $i_FactorC)
Local $i_FactorE = Int(($i_FactorB - $i_FactorD) / 30.6001)
Local $i_FactorF = Int(30.6001 * $i_FactorE)
$iDay = $i_FactorB - $i_FactorD - $i_FactorF
If $i_FactorE - 1 < 13 Then
$iMonth = $i_FactorE - 1
Else
$iMonth = $i_FactorE - 13
EndIf
If $iMonth < 3 Then
$iYear = $i_FactorC - 4715
Else
$iYear = $i_FactorC - 4716
EndIf
$iYear = StringFormat("%04d", $iYear)
$iMonth = StringFormat("%02d", $iMonth)
$iDay = StringFormat("%02d", $iDay)
Return $iYear & "/" & $iMonth & "/" & $iDay
EndFunc
Func _Date_JulianDayNo($iYear, $iMonth, $iDay)
Local $sFullDate = StringFormat("%04d/%02d/%02d", $iYear, $iMonth, $iDay)
If Not _DateIsValid($sFullDate) Then
Return SetError(1, 0, "")
EndIf
Local $iJDay = 0
Local $aiDaysInMonth = _DaysInMonth($iYear)
For $iCntr = 1 To $iMonth - 1
$iJDay = $iJDay + $aiDaysInMonth[$iCntr]
Next
$iJDay = ($iYear * 1000) + ($iJDay + $iDay)
Return $iJDay
EndFunc
Func _JulianToDate($iJDay, $sSep = "/")
Local $iYear = Int($iJDay / 1000)
Local $iDays = Mod($iJDay, 1000)
Local $iMaxDays = 365
If _DateIsLeapYear($iYear) Then $iMaxDays = 366
If $iDays > $iMaxDays Then
Return SetError(1, 0, "")
EndIf
Local $aiDaysInMonth = _DaysInMonth($iYear)
Local $iMonth = 1
While $iDays > $aiDaysInMonth[$iMonth]
$iDays = $iDays - $aiDaysInMonth[$iMonth]
$iMonth = $iMonth + 1
WEnd
Return StringFormat("%04d%s%02d%s%02d", $iYear, $sSep, $iMonth, $sSep, $iDays)
EndFunc
Func _Now()
Return _DateTimeFormat(@YEAR & "/" & @MON & "/" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC, 0)
EndFunc
Func _NowCalc()
Return @YEAR & "/" & @MON & "/" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC
EndFunc
Func _NowCalcDate()
Return @YEAR & "/" & @MON & "/" & @MDAY
EndFunc
Func _NowDate()
Return _DateTimeFormat(@YEAR & "/" & @MON & "/" & @MDAY, 0)
EndFunc
Func _NowTime($sType = 3)
If $sType < 3 Or $sType > 5 Then $sType = 3
Return _DateTimeFormat(@YEAR & "/" & @MON & "/" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC, $sType)
EndFunc
Func _SetDate($iDay, $iMonth = 0, $iYear = 0)
If $iYear = 0 Then $iYear = @YEAR
If $iMonth = 0 Then $iMonth = @MON
If Not _DateIsValid($iYear & "/" & $iMonth & "/" & $iDay) Then Return 1
Local $tSYSTEMTIME = DllStructCreate($tagSYSTEMTIME)
DllCall("kernel32.dll", "none", "GetLocalTime", "struct*", $tSYSTEMTIME)
If @error Then Return SetError(@error, @extended, 0)
DllStructSetData($tSYSTEMTIME, "Day", $iDay)
If $iMonth > 0 Then DllStructSetData($tSYSTEMTIME, "Month", $iMonth)
If $iYear > 0 Then DllStructSetData($tSYSTEMTIME, "Year", $iYear)
Local $iReturn = _Date_Time_SetLocalTime($tSYSTEMTIME)
If @error Then Return SetError(@error + 10, @extended, 0)
Return Int($iReturn)
EndFunc
Func _SetTime($iHour, $iMinute, $iSecond = 0, $iMSeconds = 0)
If $iHour < 0 Or $iHour > 23 Then Return 1
If $iMinute < 0 Or $iMinute > 59 Then Return 1
If $iSecond < 0 Or $iSecond > 59 Then Return 1
If $iMSeconds < 0 Or $iMSeconds > 999 Then Return 1
Local $tSYSTEMTIME = DllStructCreate($tagSYSTEMTIME)
DllCall("kernel32.dll", "none", "GetLocalTime", "struct*", $tSYSTEMTIME)
If @error Then Return SetError(@error, @extended, 0)
DllStructSetData($tSYSTEMTIME, "Hour", $iHour)
DllStructSetData($tSYSTEMTIME, "Minute", $iMinute)
If $iSecond > 0 Then DllStructSetData($tSYSTEMTIME, "Second", $iSecond)
If $iMSeconds > 0 Then DllStructSetData($tSYSTEMTIME, "MSeconds", $iMSeconds)
Local $iReturn = _Date_Time_SetLocalTime($tSYSTEMTIME)
If @error Then Return SetError(@error + 10, @extended, 0)
Return Int($iReturn)
EndFunc
Func _TicksToTime($iTicks, ByRef $iHours, ByRef $iMins, ByRef $iSecs)
If Number($iTicks) > 0 Then
$iTicks = Int($iTicks / 1000)
$iHours = Int($iTicks / 3600)
$iTicks = Mod($iTicks, 3600)
$iMins = Int($iTicks / 60)
$iSecs = Mod($iTicks, 60)
Return 1
ElseIf Number($iTicks) = 0 Then
$iHours = 0
$iTicks = 0
$iMins = 0
$iSecs = 0
Return 1
Else
Return SetError(1, 0, 0)
EndIf
EndFunc
Func _TimeToTicks($iHours = @HOUR, $iMins = @MIN, $iSecs = @SEC)
If StringIsInt($iHours) And StringIsInt($iMins) And StringIsInt($iSecs) Then
Local $iTicks = 1000 * ((3600 * $iHours) + (60 * $iMins) + $iSecs)
Return $iTicks
Else
Return SetError(1, 0, 0)
EndIf
EndFunc
Func _WeekNumberISO($iYear = @YEAR, $iMonth = @MON, $iDay = @MDAY)
If $iDay > 31 Or $iDay < 1 Then
Return SetError(1, 0, -1)
ElseIf Not __DateIsMonth($iMonth) Then
Return SetError(2, 0, -1)
ElseIf $iYear < 1 Or $iYear > 2999 Then
Return SetError(3, 0, -1)
EndIf
Local $iDow = _DateToDayOfWeekISO($iYear, $iMonth, $iDay) - 1
Local $iDow0101 = _DateToDayOfWeekISO($iYear, 1, 1) - 1
If ($iMonth = 1 And 3 < $iDow0101 And $iDow0101 < 7 - ($iDay - 1)) Then
$iDow = $iDow0101 - 1
$iDow0101 = _DateToDayOfWeekISO($iYear - 1, 1, 1) - 1
$iMonth = 12
$iDay = 31
$iYear = $iYear - 1
ElseIf ($iMonth = 12 And 30 - ($iDay - 1) < _DateToDayOfWeekISO($iYear + 1, 1, 1) - 1 And _DateToDayOfWeekISO($iYear + 1, 1, 1) - 1 < 4) Then
Return 1
EndIf
Return Int((_DateToDayOfWeekISO($iYear, 1, 1) - 1 < 4) + 4 * ($iMonth - 1) + (2 * ($iMonth - 1) + ($iDay - 1) + $iDow0101 - $iDow + 6) * 36 / 256)
EndFunc
Func _WeekNumber($iYear = @YEAR, $iMonth = @MON, $iDay = @MDAY, $iWeekStart = 1)
If $iDay > 31 Or $iDay < 1 Then
Return SetError(1, 0, -1)
ElseIf Not __DateIsMonth($iMonth) Then
Return SetError(3, 0, -1)
ElseIf $iYear < 1 Or $iYear > 2999 Then
Return SetError(4, 0, -1)
ElseIf $iWeekStart < 1 Or $iWeekStart > 2 Then
Return SetError(2, 0, -1)
EndIf
Local $iStartWeek1, $iEndWeek1
Local $iDow0101 = _DateToDayOfWeekISO($iYear, 1, 1)
Local $iDate = $iYear & '/' & $iMonth & '/' & $iDay
If $iWeekStart = 1 Then
If $iDow0101 = 6 Then
$iStartWeek1 = 0
Else
$iStartWeek1 = -1 * $iDow0101 - 1
EndIf
$iEndWeek1 = $iStartWeek1 + 6
Else
$iStartWeek1 = $iDow0101 * -1
$iEndWeek1 = $iStartWeek1 + 6
EndIf
Local $iStartWeek1ny
Local $iEndWeek1Date = _DateAdd('d', $iEndWeek1, $iYear & '/01/01')
Local $iDow0101ny = _DateToDayOfWeekISO($iYear + 1, 1, 1)
If $iWeekStart = 1 Then
If $iDow0101ny = 6 Then
$iStartWeek1ny = 0
Else
$iStartWeek1ny = -1 * $iDow0101ny - 1
EndIf
Else
$iStartWeek1ny = $iDow0101ny * -1
EndIf
Local $iStartWeek1Dateny = _DateAdd('d', $iStartWeek1ny, $iYear + 1 & '/01/01')
Local $iCurrDateDiff = _DateDiff('d', $iEndWeek1Date, $iDate) - 1
Local $iCurrDateDiffny = _DateDiff('d', $iStartWeek1Dateny, $iDate)
If $iCurrDateDiff >= 0 And $iCurrDateDiffny < 0 Then Return 2 + Int($iCurrDateDiff / 7)
If $iCurrDateDiff < 0 Or $iCurrDateDiffny >= 0 Then Return 1
EndFunc
Func _DaysInMonth($iYear)
Local $aDays = [12, 31, (_DateIsLeapYear($iYear) ? 29 : 28), 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
Return $aDays
EndFunc
Func __Date_Time_CloneSystemTime($pSystemTime)
Local $tSystemTime1 = DllStructCreate($tagSYSTEMTIME, $pSystemTime)
Local $tSystemTime2 = DllStructCreate($tagSYSTEMTIME)
DllStructSetData($tSystemTime2, "Month", DllStructGetData($tSystemTime1, "Month"))
DllStructSetData($tSystemTime2, "Day", DllStructGetData($tSystemTime1, "Day"))
DllStructSetData($tSystemTime2, "Year", DllStructGetData($tSystemTime1, "Year"))
DllStructSetData($tSystemTime2, "Hour", DllStructGetData($tSystemTime1, "Hour"))
DllStructSetData($tSystemTime2, "Minute", DllStructGetData($tSystemTime1, "Minute"))
DllStructSetData($tSystemTime2, "Second", DllStructGetData($tSystemTime1, "Second"))
DllStructSetData($tSystemTime2, "MSeconds", DllStructGetData($tSystemTime1, "MSeconds"))
DllStructSetData($tSystemTime2, "DOW", DllStructGetData($tSystemTime1, "DOW"))
Return $tSystemTime2
EndFunc
Func _Date_Time_CompareFileTime($tFileTime1, $tFileTime2)
Local $aResult = DllCall("kernel32.dll", "long", "CompareFileTime", "struct*", $tFileTime1, "struct*", $tFileTime2)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _Date_Time_DOSDateTimeToFileTime($iFatDate, $iFatTime)
Local $tTime = DllStructCreate($tagFILETIME)
Local $aResult = DllCall("kernel32.dll", "bool", "DosDateTimeToFileTime", "word", $iFatDate, "word", $iFatTime, "struct*", $tTime)
If @error Then Return SetError(@error, @extended, 0)
Return SetExtended($aResult[0], $tTime)
EndFunc
Func _Date_Time_DOSDateToArray($iDosDate)
Local $aDate[3]
$aDate[0] = BitAND($iDosDate, 0x1F)
$aDate[1] = BitAND(BitShift($iDosDate, 5), 0x0F)
$aDate[2] = BitAND(BitShift($iDosDate, 9), 0x3F) + 1980
Return $aDate
EndFunc
Func _Date_Time_DOSDateTimeToArray($iDosDate, $iDosTime)
Local $aDate[6]
$aDate[0] = BitAND($iDosDate, 0x1F)
$aDate[1] = BitAND(BitShift($iDosDate, 5), 0x0F)
$aDate[2] = BitAND(BitShift($iDosDate, 9), 0x3F) + 1980
$aDate[5] = BitAND($iDosTime, 0x1F) * 2
$aDate[4] = BitAND(BitShift($iDosTime, 5), 0x3F)
$aDate[3] = BitAND(BitShift($iDosTime, 11), 0x1F)
Return $aDate
EndFunc
Func _Date_Time_DOSDateTimeToStr($iDosDate, $iDosTime)
Local $aDate = _Date_Time_DOSDateTimeToArray($iDosDate, $iDosTime)
Return StringFormat("%02d/%02d/%04d %02d:%02d:%02d", $aDate[0], $aDate[1], $aDate[2], $aDate[3], $aDate[4], $aDate[5])
EndFunc
Func _Date_Time_DOSDateToStr($iDosDate)
Local $aDate = _Date_Time_DOSDateToArray($iDosDate)
Return StringFormat("%02d/%02d/%04d", $aDate[0], $aDate[1], $aDate[2])
EndFunc
Func _Date_Time_DOSTimeToArray($iDosTime)
Local $aTime[3]
$aTime[2] = BitAND($iDosTime, 0x1F) * 2
$aTime[1] = BitAND(BitShift($iDosTime, 5), 0x3F)
$aTime[0] = BitAND(BitShift($iDosTime, 11), 0x1F)
Return $aTime
EndFunc
Func _Date_Time_DOSTimeToStr($iDosTime)
Local $aTime = _Date_Time_DOSTimeToArray($iDosTime)
Return StringFormat("%02d:%02d:%02d", $aTime[0], $aTime[1], $aTime[2])
EndFunc
Func _Date_Time_EncodeFileTime($iMonth, $iDay, $iYear, $iHour = 0, $iMinute = 0, $iSecond = 0, $iMSeconds = 0)
Local $tSYSTEMTIME = _Date_Time_EncodeSystemTime($iMonth, $iDay, $iYear, $iHour, $iMinute, $iSecond, $iMSeconds)
Return _Date_Time_SystemTimeToFileTime($tSYSTEMTIME)
EndFunc
Func _Date_Time_EncodeSystemTime($iMonth, $iDay, $iYear, $iHour = 0, $iMinute = 0, $iSecond = 0, $iMSeconds = 0)
Local $tSYSTEMTIME = DllStructCreate($tagSYSTEMTIME)
DllStructSetData($tSYSTEMTIME, "Month", $iMonth)
DllStructSetData($tSYSTEMTIME, "Day", $iDay)
DllStructSetData($tSYSTEMTIME, "Year", $iYear)
DllStructSetData($tSYSTEMTIME, "Hour", $iHour)
DllStructSetData($tSYSTEMTIME, "Minute", $iMinute)
DllStructSetData($tSYSTEMTIME, "Second", $iSecond)
DllStructSetData($tSYSTEMTIME, "MSeconds", $iMSeconds)
Return $tSYSTEMTIME
EndFunc
Func _Date_Time_FileTimeToArray(ByRef $tFileTime)
If ((DllStructGetData($tFileTime, 1) + DllStructGetData($tFileTime, 2)) = 0) Then Return SetError(10, 0, 0)
Local $tSYSTEMTIME = _Date_Time_FileTimeToSystemTime($tFileTime)
If @error Then Return SetError(@error, @extended, 0)
Return _Date_Time_SystemTimeToArray($tSYSTEMTIME)
EndFunc
Func _Date_Time_FileTimeToStr(ByRef $tFileTime, $iFmt = 0)
Local $aDate = _Date_Time_FileTimeToArray($tFileTime)
If @error Then Return SetError(@error, @extended, "")
If $iFmt Then
Return StringFormat("%04d/%02d/%02d %02d:%02d:%02d", $aDate[2], $aDate[0], $aDate[1], $aDate[3], $aDate[4], $aDate[5])
Else
Return StringFormat("%02d/%02d/%04d %02d:%02d:%02d", $aDate[0], $aDate[1], $aDate[2], $aDate[3], $aDate[4], $aDate[5])
EndIf
EndFunc
Func _Date_Time_FileTimeToDOSDateTime($tFileTime)
Local $aDate[2]
Local $aResult = DllCall("kernel32.dll", "bool", "FileTimeToDosDateTime", "struct*", $tFileTime, "word*", 0, "word*", 0)
If @error Then Return SetError(@error, @extended, $aDate)
$aDate[0] = $aResult[2]
$aDate[1] = $aResult[3]
Return SetExtended($aResult[0], $aDate)
EndFunc
Func _Date_Time_FileTimeToLocalFileTime($tFileTime)
Local $tLocal = DllStructCreate($tagFILETIME)
Local $aResult = DllCall("kernel32.dll", "bool", "FileTimeToLocalFileTime", "struct*", $tFileTime, "struct*", $tLocal)
If @error Then Return SetError(@error, @extended, 0)
Return SetExtended($aResult[0], $tLocal)
EndFunc
Func _Date_Time_FileTimeToSystemTime($tFileTime)
Local $tSystTime = DllStructCreate($tagSYSTEMTIME)
Local $aResult = DllCall("kernel32.dll", "bool", "FileTimeToSystemTime", "struct*", $tFileTime, "struct*", $tSystTime)
If @error Then Return SetError(@error, @extended, 0)
Return SetExtended($aResult[0], $tSystTime)
EndFunc
Func _Date_Time_GetFileTime($hFile)
Local $aDate[3]
$aDate[0] = DllStructCreate($tagFILETIME)
$aDate[1] = DllStructCreate($tagFILETIME)
$aDate[2] = DllStructCreate($tagFILETIME)
Local $aResult = DllCall("kernel32.dll", "bool", "GetFileTime", "handle", $hFile, "struct*", $aDate[0], "struct*", $aDate[1], "struct*", $aDate[2])
If @error Then Return SetError(@error, @extended, 0)
Return SetExtended($aResult[0], $aDate)
EndFunc
Func _Date_Time_GetLocalTime()
Local $tSystTime = DllStructCreate($tagSYSTEMTIME)
DllCall("kernel32.dll", "none", "GetLocalTime", "struct*", $tSystTime)
If @error Then Return SetError(@error, @extended, 0)
Return $tSystTime
EndFunc
Func _Date_Time_GetSystemTime()
Local $tSystTime = DllStructCreate($tagSYSTEMTIME)
DllCall("kernel32.dll", "none", "GetSystemTime", "struct*", $tSystTime)
If @error Then Return SetError(@error, @extended, 0)
Return $tSystTime
EndFunc
Func _Date_Time_GetSystemTimeAdjustment()
Local $aInfo[3]
Local $aResult = DllCall("kernel32.dll", "bool", "GetSystemTimeAdjustment", "dword*", 0, "dword*", 0, "bool*", 0)
If @error Then Return SetError(@error, @extended, 0)
$aInfo[0] = $aResult[1]
$aInfo[1] = $aResult[2]
$aInfo[2] = $aResult[3] <> 0
Return SetExtended($aResult[0], $aInfo)
EndFunc
Func _Date_Time_GetSystemTimeAsFileTime()
Local $tFileTime = DllStructCreate($tagFILETIME)
DllCall("kernel32.dll", "none", "GetSystemTimeAsFileTime", "struct*", $tFileTime)
If @error Then Return SetError(@error, @extended, 0)
Return $tFileTime
EndFunc
Func _Date_Time_GetSystemTimes()
Local $aInfo[3]
$aInfo[0] = DllStructCreate($tagFILETIME)
$aInfo[1] = DllStructCreate($tagFILETIME)
$aInfo[2] = DllStructCreate($tagFILETIME)
Local $aResult = DllCall("kernel32.dll", "bool", "GetSystemTimes", "struct*", $aInfo[0], "struct*", $aInfo[1], "struct*", $aInfo[2])
If @error Then Return SetError(@error, @extended, 0)
Return SetExtended($aResult[0], $aInfo)
EndFunc
Func _Date_Time_GetTickCount()
Local $aResult = DllCall("kernel32.dll", "dword", "GetTickCount")
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _Date_Time_GetTimeZoneInformation()
Local $tTimeZone = DllStructCreate($tagTIME_ZONE_INFORMATION)
Local $aResult = DllCall("kernel32.dll", "dword", "GetTimeZoneInformation", "struct*", $tTimeZone)
If @error Or $aResult[0] = -1 Then Return SetError(@error, @extended, 0)
Local $aInfo[8]
$aInfo[0] = $aResult[0]
$aInfo[1] = DllStructGetData($tTimeZone, "Bias")
$aInfo[2] = DllStructGetData($tTimeZone, "StdName")
$aInfo[3] = __Date_Time_CloneSystemTime(DllStructGetPtr($tTimeZone, "StdDate"))
$aInfo[4] = DllStructGetData($tTimeZone, "StdBias")
$aInfo[5] = DllStructGetData($tTimeZone, "DayName")
$aInfo[6] = __Date_Time_CloneSystemTime(DllStructGetPtr($tTimeZone, "DayDate"))
$aInfo[7] = DllStructGetData($tTimeZone, "DayBias")
Return $aInfo
EndFunc
Func _Date_Time_LocalFileTimeToFileTime($tLocalTime)
Local $tFileTime = DllStructCreate($tagFILETIME)
Local $aResult = DllCall("kernel32.dll", "bool", "LocalFileTimeToFileTime", "struct*", $tLocalTime, "struct*", $tFileTime)
If @error Then Return SetError(@error, @extended, 0)
Return SetExtended($aResult[0], $tFileTime)
EndFunc
Func _Date_Time_SetFileTime($hFile, $tCreateTime, $tLastAccess, $tLastWrite)
Local $aResult = DllCall("kernel32.dll", "bool", "SetFileTime", "handle", $hFile, "struct*", $tCreateTime, "struct*", $tLastAccess, "struct*", $tLastWrite)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _Date_Time_SetLocalTime($tSYSTEMTIME)
Local $aResult = DllCall("kernel32.dll", "bool", "SetLocalTime", "struct*", $tSYSTEMTIME)
If @error Or Not $aResult[0] Then Return SetError(@error + 10, @extended, False)
$aResult = DllCall("kernel32.dll", "bool", "SetLocalTime", "struct*", $tSYSTEMTIME)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _Date_Time_SetSystemTime($tSYSTEMTIME)
Local $aResult = DllCall("kernel32.dll", "bool", "SetSystemTime", "struct*", $tSYSTEMTIME)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _Date_Time_SetSystemTimeAdjustment($iAdjustment, $bDisabled)
Local $hToken = _Security__OpenThreadTokenEx(BitOR($TOKEN_ADJUST_PRIVILEGES, $TOKEN_QUERY))
If @error Then Return SetError(@error + 10, @extended, False)
_Security__SetPrivilege($hToken, "SeSystemtimePrivilege", True)
Local $iError = @error
Local $iLastError = @extended
Local $bRet = False
If Not @error Then
Local $aResult = DllCall("kernel32.dll", "bool", "SetSystemTimeAdjustment", "dword", $iAdjustment, "bool", $bDisabled)
If @error Then
$iError = @error
$iLastError = @extended
ElseIf $aResult[0] Then
$bRet = True
Else
$iError = 20
$iLastError = _WinAPI_GetLastError()
EndIf
_Security__SetPrivilege($hToken, "SeSystemtimePrivilege", False)
If Not $iError And @error Then $iError = 22
EndIf
_WinAPI_CloseHandle($hToken)
Return SetError($iError, $iLastError, $bRet)
EndFunc
Func _Date_Time_SetTimeZoneInformation($iBias, $sStdName, $tStdDate, $iStdBias, $sDayName, $tDayDate, $iDayBias)
Local $tZoneInfo = DllStructCreate($tagTIME_ZONE_INFORMATION)
DllStructSetData($tZoneInfo, "Bias", $iBias)
DllStructSetData($tZoneInfo, "StdName", $sStdName)
_MemMoveMemory($tStdDate, DllStructGetPtr($tZoneInfo, "StdDate"), DllStructGetSize($tStdDate))
DllStructSetData($tZoneInfo, "StdBias", $iStdBias)
DllStructSetData($tZoneInfo, "DayName", $sDayName)
_MemMoveMemory($tDayDate, DllStructGetPtr($tZoneInfo, "DayDate"), DllStructGetSize($tDayDate))
DllStructSetData($tZoneInfo, "DayBias", $iDayBias)
Local $hToken = _Security__OpenThreadTokenEx(BitOR($TOKEN_ADJUST_PRIVILEGES, $TOKEN_QUERY))
If @error Then Return SetError(@error + 10, @extended, False)
_Security__SetPrivilege($hToken, "SeTimeZonePrivilege", True)
Local $iError = @error
Local $iLastError = @extended
Local $bRet = False
If Not @error Then
Local $aResult = DllCall("kernel32.dll", "bool", "SetTimeZoneInformation", "struct*", $tZoneInfo)
If @error Then
$iError = @error
$iLastError = @extended
ElseIf $aResult[0] Then
$iLastError = 0
$bRet = True
Else
$iError = 20
$iLastError = _WinAPI_GetLastError()
EndIf
_Security__SetPrivilege($hToken, "SeTimeZonePrivilege", False)
If Not $iError And @error Then $iError = 22
EndIf
_WinAPI_CloseHandle($hToken)
Return SetError($iError, $iLastError, $bRet)
EndFunc
Func _Date_Time_SystemTimeToArray(ByRef $tSYSTEMTIME)
Local $aInfo[8]
$aInfo[0] = DllStructGetData($tSYSTEMTIME, "Month")
$aInfo[1] = DllStructGetData($tSYSTEMTIME, "Day")
$aInfo[2] = DllStructGetData($tSYSTEMTIME, "Year")
$aInfo[3] = DllStructGetData($tSYSTEMTIME, "Hour")
$aInfo[4] = DllStructGetData($tSYSTEMTIME, "Minute")
$aInfo[5] = DllStructGetData($tSYSTEMTIME, "Second")
$aInfo[6] = DllStructGetData($tSYSTEMTIME, "MSeconds")
$aInfo[7] = DllStructGetData($tSYSTEMTIME, "DOW")
Return $aInfo
EndFunc
Func _Date_Time_SystemTimeToDateStr(ByRef $tSYSTEMTIME, $iFmt = 0)
Local $aInfo = _Date_Time_SystemTimeToArray($tSYSTEMTIME)
If @error Then Return SetError(@error, @extended, "")
If $iFmt Then
Return StringFormat("%04d/%02d/%02d", $aInfo[2], $aInfo[0], $aInfo[1])
Else
Return StringFormat("%02d/%02d/%04d", $aInfo[0], $aInfo[1], $aInfo[2])
EndIf
EndFunc
Func _Date_Time_SystemTimeToDateTimeStr(ByRef $tSYSTEMTIME, $iFmt = 0)
Local $aInfo = _Date_Time_SystemTimeToArray($tSYSTEMTIME)
If @error Then Return SetError(@error, @extended, "")
If $iFmt Then
Return StringFormat("%04d/%02d/%02d %02d:%02d:%02d", $aInfo[2], $aInfo[0], $aInfo[1], $aInfo[3], $aInfo[4], $aInfo[5])
Else
Return StringFormat("%02d/%02d/%04d %02d:%02d:%02d", $aInfo[0], $aInfo[1], $aInfo[2], $aInfo[3], $aInfo[4], $aInfo[5])
EndIf
EndFunc
Func _Date_Time_SystemTimeToFileTime($tSYSTEMTIME)
Local $tFileTime = DllStructCreate($tagFILETIME)
Local $aResult = DllCall("kernel32.dll", "bool", "SystemTimeToFileTime", "struct*", $tSYSTEMTIME, "struct*", $tFileTime)
If @error Then Return SetError(@error, @extended, 0)
Return SetExtended($aResult[0], $tFileTime)
EndFunc
Func _Date_Time_SystemTimeToTimeStr(ByRef $tSYSTEMTIME)
Local $aInfo = _Date_Time_SystemTimeToArray($tSYSTEMTIME)
Return StringFormat("%02d:%02d:%02d", $aInfo[3], $aInfo[4], $aInfo[5])
EndFunc
Func _Date_Time_SystemTimeToTzSpecificLocalTime($tUTC, $tTimeZone = 0)
Local $tLocalTime = DllStructCreate($tagSYSTEMTIME)
Local $aResult = DllCall("kernel32.dll", "bool", "SystemTimeToTzSpecificLocalTime", "struct*", $tTimeZone, "struct*", $tUTC, "struct*", $tLocalTime)
If @error Then Return SetError(@error, @extended, 0)
Return SetExtended($aResult[0], $tLocalTime)
EndFunc
Func _Date_Time_TzSpecificLocalTimeToSystemTime($tLocalTime, $tTimeZone = 0)
Local $tUTC = DllStructCreate($tagSYSTEMTIME)
Local $aResult = DllCall("kernel32.dll", "bool", "TzSpecificLocalTimeToSystemTime", "struct*", $tTimeZone, "struct*", $tLocalTime, "struct*", $tUTC)
If @error Then Return SetError(@error, @extended, 0)
Return SetExtended($aResult[0], $tUTC)
EndFunc
Global Const $INET_LOCALCACHE = 0
Global Const $INET_FORCERELOAD = 1
Global Const $INET_IGNORESSL = 2
Global Const $INET_ASCIITRANSFER = 4
Global Const $INET_BINARYTRANSFER = 8
Global Const $INET_FORCEBYPASS = 16
Global Const $INET_DOWNLOADWAIT = 0
Global Const $INET_DOWNLOADBACKGROUND = 1
Global Const $INET_DOWNLOADREAD = 0
Global Const $INET_DOWNLOADSIZE = 1
Global Const $INET_DOWNLOADCOMPLETE = 2
Global Const $INET_DOWNLOADSUCCESS = 3
Global Const $INET_DOWNLOADERROR = 4
Global Const $INET_DOWNLOADEXTENDED = 5
Func _GetIP()
Local Const $GETIP_TIMER = 300000
Local Static $hTimer = 0
Local Static $sLastIP = 0
If TimerDiff($hTimer) < $GETIP_TIMER And Not $sLastIP Then
Return SetExtended(1, $sLastIP)
EndIf
Local $aGetIPURL = ["https://api.ipify.org", "http://checkip.dyndns.org", "http://www.myexternalip.com/raw", "http://bot.whatismyipaddress.com"],  $aReturn = 0,  $sReturn = ""
For $i = 0 To UBound($aGetIPURL) - 1
$sReturn = InetRead($aGetIPURL[$i])
If @error Or $sReturn == "" Then ContinueLoop
$aReturn = StringRegExp(BinaryToString($sReturn), "((?:\d{1,3}\.){3}\d{1,3})", $STR_REGEXPARRAYGLOBALMATCH)
If Not @error Then
$sReturn = $aReturn[0]
ExitLoop
EndIf
$sReturn = ""
Next
$hTimer = TimerInit()
$sLastIP = $sReturn
If $sReturn == "" Then Return SetError(1, 0, -1)
Return $sReturn
EndFunc
Func _INetExplorerCapable($sIEString)
If StringLen($sIEString) <= 0 Then Return SetError(1, 0, '')
Local $s_IEReturn
Local $n_IEChar
For $i_IECount = 1 To StringLen($sIEString)
$n_IEChar = '0x' & Hex(Asc(StringMid($sIEString, $i_IECount, 1)), 2)
If $n_IEChar < 0x21 Or $n_IEChar = 0x25 Or $n_IEChar = 0x2f Or $n_IEChar > 0x7f Then
$s_IEReturn = $s_IEReturn & '%' & StringRight($n_IEChar, 2)
Else
$s_IEReturn = $s_IEReturn & Chr($n_IEChar)
EndIf
Next
Return $s_IEReturn
EndFunc
Func _INetGetSource($sURL, $bString = True)
Local $sString = InetRead($sURL, $INET_FORCERELOAD)
Local $iError = @error, $iExtended = @extended
If $bString = Default Or $bString Then $sString = BinaryToString($sString)
Return SetError($iError, $iExtended, $sString)
EndFunc
Func _INetMail($sMailTo, $sMailSubject, $sMailBody)
Local $iPrev = Opt("ExpandEnvStrings", 1)
Local $sVar, $sDflt = RegRead('HKCU\Software\Clients\Mail', "")
If $sDflt = "Windows Live Mail" Then
$sVar = RegRead('HKCR\WLMail.Url.Mailto\Shell\open\command', "")
Else
$sVar = RegRead('HKCR\mailto\shell\open\command', "")
EndIf
Local $iRet = Run(StringReplace($sVar, '%1', _INetExplorerCapable('mailto:' & $sMailTo & '?subject=' & $sMailSubject & '&body=' & $sMailBody)))
Local $iError = @error, $iExtended = @extended
Opt("ExpandEnvStrings", $iPrev)
Return SetError($iError, $iExtended, $iRet)
EndFunc
Func _INetSmtpMail($sSMTPServer, $sFromName, $sFromAddress, $sToAddress, $sSubject = "", $aBody = "", $sEHLO = "", $sFirst = "", $bTrace = 0)
If $sSMTPServer = "" Or $sFromAddress = "" Or $sToAddress = "" Or $sFromName = "" Or StringLen($sFromName) > 256 Then Return SetError(1, 0, 0)
If $sEHLO = "" Then $sEHLO = @ComputerName
If TCPStartup() = 0 Then Return SetError(2, 0, 0)
Local $s_IPAddress, $i_Count
If StringRegExp($sSMTPServer, "^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$") Then
$s_IPAddress = $sSMTPServer
Else
$s_IPAddress = TCPNameToIP($sSMTPServer)
EndIf
If $s_IPAddress = "" Then
TCPShutdown()
Return SetError(3, 0, 0)
EndIf
Local $vSocket = TCPConnect($s_IPAddress, 25)
If $vSocket = -1 Then
TCPShutdown()
Return SetError(4, 0, 0)
EndIf
Local $aSend[6], $aReplyCode[6]
$aSend[0] = "HELO " & $sEHLO & @CRLF
If StringLeft($sEHLO, 5) = "EHLO " Then $aSend[0] = $sEHLO & @CRLF
$aReplyCode[0] = "250"
$aSend[1] = "MAIL FROM: <" & $sFromAddress & ">" & @CRLF
$aReplyCode[1] = "250"
$aSend[2] = "RCPT TO: <" & $sToAddress & ">" & @CRLF
$aReplyCode[2] = "250"
$aSend[3] = "DATA" & @CRLF
$aReplyCode[3] = "354"
Local $aResult = _Date_Time_GetTimeZoneInformation()
Local $iBias = -$aResult[1] / 60
Local $iBiasH = Int($iBias)
Local $iBiasM = 0
If $iBiasH <> $iBias Then $iBiasM = Abs($iBias - $iBiasH) * 60
$iBias = StringFormat(" (%+.2d%.2d)", $iBiasH, $iBiasM)
$aSend[4] = "From:" & $sFromName & "<" & $sFromAddress & ">" & @CRLF &  "To:" & "<" & $sToAddress & ">" & @CRLF &  "Subject:" & $sSubject & @CRLF &  "Mime-Version: 1.0" & @CRLF &  "Date: " & _DateDayOfWeek(@WDAY, 1) & ", " & @MDAY & " " & _DateToMonth(@MON, 1) & " " & @YEAR & " " & @HOUR & ":" & @MIN & ":" & @SEC & $iBias & @CRLF &  "Content-Type: text/plain; charset=US-ASCII" & @CRLF &  @CRLF
$aReplyCode[4] = ""
$aSend[5] = @CRLF & "." & @CRLF
$aReplyCode[5] = "250"
If __SmtpSend($vSocket, $aSend[0], $aReplyCode[0], $bTrace, "220", $sFirst) Then Return SetError(50, 0, 0)
For $i_Count = 1 To UBound($aSend) - 2
If __SmtpSend($vSocket, $aSend[$i_Count], $aReplyCode[$i_Count], $bTrace) Then Return SetError(50 + $i_Count, 0, 0)
Next
For $i_Count = 0 To UBound($aBody) - 1
If StringLeft($aBody[$i_Count], 1) = "." Then $aBody[$i_Count] = "." & $aBody[$i_Count]
If __SmtpSend($vSocket, $aBody[$i_Count] & @CRLF, "", $bTrace) Then Return SetError(500 + $i_Count, 0, 0)
Next
$i_Count = UBound($aSend) - 1
If __SmtpSend($vSocket, $aSend[$i_Count], $aReplyCode[$i_Count], $bTrace) Then Return SetError(5000, 0, 0)
TCPCloseSocket($vSocket)
TCPShutdown()
Return 1
EndFunc
Func __SmtpTrace($sStr, $iTimeout = 0)
Local $sW_TITLE = "SMTP trace"
Local $sSmtpTrace = ControlGetText($sW_TITLE, "", "Static1")
$sStr = StringLeft(StringReplace($sStr, @CRLF, ""), 70)
$sSmtpTrace &= @HOUR & ":" & @MIN & ":" & @SEC & " " & $sStr & @LF
If WinExists($sW_TITLE) Then
ControlSetText($sW_TITLE, "", "Static1", $sSmtpTrace)
Else
SplashTextOn($sW_TITLE, $sSmtpTrace, 400, 500, 500, 100, 4 + 16, "", 8)
EndIf
If $iTimeout Then Sleep($iTimeout * 1000)
EndFunc
Func __SmtpSend($vSocket, $sSend, $sReplyCode, $bTrace, $sIntReply = "", $sFirst = "")
Local $sReceive, $i, $hTimer
If $bTrace Then __SmtpTrace($sSend)
If $sIntReply <> "" Then
If $sFirst <> -1 Then
If TCPSend($vSocket, $sFirst) = 0 Then
TCPCloseSocket($vSocket)
TCPShutdown()
Return 1
EndIf
EndIf
$sReceive = ""
$hTimer = TimerInit()
While StringLeft($sReceive, StringLen($sIntReply)) <> $sIntReply And TimerDiff($hTimer) < 45000
$sReceive = TCPRecv($vSocket, 1000)
If $bTrace And $sReceive <> "" Then __SmtpTrace("intermediate->" & $sReceive)
WEnd
EndIf
If TCPSend($vSocket, $sSend) = 0 Then
TCPCloseSocket($vSocket)
TCPShutdown()
Return 1
EndIf
$hTimer = TimerInit()
$sReceive = ""
While $sReceive = "" And TimerDiff($hTimer) < 45000
$i += 1
$sReceive = TCPRecv($vSocket, 1000)
If $sReplyCode = "" Then ExitLoop
WEnd
If $sReplyCode <> "" Then
If $bTrace Then __SmtpTrace($i & " <- " & $sReceive)
If StringLeft($sReceive, StringLen($sReplyCode)) <> $sReplyCode Then
TCPCloseSocket($vSocket)
TCPShutdown()
If $bTrace Then __SmtpTrace("<-> " & $sReplyCode, 5)
Return 2
EndIf
EndIf
Return 0
EndFunc
Func _TCPIpToName($sIp, $iOption = Default, $hDll = Default)
Local $iINADDR_NONE = 0xffffffff, $iAF_INET = 2, $sSeparator = @CR
If $iOption = Default Then $iOption = 0
If $hDll = Default Then $hDll = "ws2_32.dll"
Local $avDllCall = DllCall($hDll, "ulong", "inet_addr", "STR", $sIp)
If @error Then Return SetError(1, 0, "")
Local $vBinIP = $avDllCall[0]
If $vBinIP = $iINADDR_NONE Then Return SetError(2, 0, "")
$avDllCall = DllCall($hDll, "ptr", "gethostbyaddr", "ptr*", $vBinIP, "int", 4, "int", $iAF_INET)
If @error Then Return SetError(3, 0, "")
Local $pvHostent = $avDllCall[0]
If $pvHostent = 0 Then
$avDllCall = DllCall($hDll, "int", "WSAGetLastError")
If @error Then Return SetError(5, 0, "")
Return SetError(4, $avDllCall[0], "")
EndIf
Local $tHostent = DllStructCreate("ptr;ptr;short;short;ptr", $pvHostent)
Local $sHostnames = __TCPIpToName_szStringRead(DllStructGetData($tHostent, 1))
If @error Then Return SetError(6, 0, $sHostnames)
If $iOption = 1 Then
Local $tAliases
$sHostnames &= $sSeparator
For $i = 0 To 63
$tAliases = DllStructCreate("ptr", DllStructGetData($tHostent, 2) + ($i * 4))
If DllStructGetData($tAliases, 1) = 0 Then ExitLoop
$sHostnames &= __TCPIpToName_szStringRead(DllStructGetData($tAliases, 1))
If @error Then
SetError(7)
ExitLoop
EndIf
Next
Return StringSplit(StringStripWS($sHostnames, $STR_STRIPTRAILING), @CR)
Else
Return $sHostnames
EndIf
EndFunc
Func __TCPIpToName_szStringRead($pStr, $iLen = -1)
Local $tString
If $pStr < 1 Then Return ""
If $iLen < 0 Then $iLen = _WinAPI_StrLen($pStr, False)
$tString = DllStructCreate("char[" & $iLen & "]", $pStr)
If @error Then Return SetError(2, 0, "")
Return SetExtended($iLen, DllStructGetData($tString, 1))
EndFunc
Global Const $SS_LEFT = 0x0
Global Const $SS_CENTER = 0x1
Global Const $SS_RIGHT = 0x2
Global Const $SS_ICON = 0x3
Global Const $SS_BLACKRECT = 0x4
Global Const $SS_GRAYRECT = 0x5
Global Const $SS_WHITERECT = 0x6
Global Const $SS_BLACKFRAME = 0x7
Global Const $SS_GRAYFRAME = 0x8
Global Const $SS_WHITEFRAME = 0x9
Global Const $SS_SIMPLE = 0xB
Global Const $SS_LEFTNOWORDWRAP = 0xC
Global Const $SS_BITMAP = 0xE
Global Const $SS_ENHMETAFILE = 0xF
Global Const $SS_ETCHEDHORZ = 0x10
Global Const $SS_ETCHEDVERT = 0x11
Global Const $SS_ETCHEDFRAME = 0x12
Global Const $SS_REALSIZECONTROL = 0x40
Global Const $SS_NOPREFIX = 0x0080
Global Const $SS_NOTIFY = 0x0100
Global Const $SS_CENTERIMAGE = 0x0200
Global Const $SS_RIGHTJUST = 0x0400
Global Const $SS_SUNKEN = 0x1000
Global Const $GUI_SS_DEFAULT_LABEL = 0
Global Const $GUI_SS_DEFAULT_GRAPHIC = 0
Global Const $GUI_SS_DEFAULT_ICON = $SS_NOTIFY
Global Const $GUI_SS_DEFAULT_PIC = $SS_NOTIFY
Global Const $STM_SETICON = 0x0170
Global Const $STM_GETICON = 0x0171
Global Const $STM_SETIMAGE = 0x0172
Global Const $STM_GETIMAGE = 0x0173
Func _HexToString($sHex)
If Not (StringLeft($sHex, 2) == "0x") Then $sHex = "0x" & $sHex
Return BinaryToString($sHex, $SB_UTF8)
EndFunc
Func _StringBetween($sString, $sStart, $sEnd, $iMode = $STR_ENDISSTART, $bCase = False)
$sStart = $sStart ? "\Q" & $sStart & "\E" : "\A"
If $iMode <> $STR_ENDNOTSTART Then $iMode = $STR_ENDISSTART
If $iMode = $STR_ENDISSTART Then
$sEnd = $sEnd ? "(?=\Q" & $sEnd & "\E)" : "\z"
Else
$sEnd = $sEnd ? "\Q" & $sEnd & "\E" : "\z"
EndIf
If $bCase = Default Then
$bCase = False
EndIf
Local $aReturn = StringRegExp($sString, "(?s" & (Not $bCase ? "i" : "") & ")" & $sStart & "(.*?)" & $sEnd, $STR_REGEXPARRAYGLOBALMATCH)
If @error Then Return SetError(1, 0, 0)
Return $aReturn
EndFunc
Func _StringExplode($sString, $sDelimiter, $iLimit = 0)
If $iLimit = Default Then $iLimit = 0
If $iLimit > 0 Then
Local Const $NULL = Chr(0)
$sString = StringReplace($sString, $sDelimiter, $NULL, $iLimit)
$sDelimiter = $NULL
ElseIf $iLimit < 0 Then
Local $iIndex = StringInStr($sString, $sDelimiter, $STR_NOCASESENSEBASIC, $iLimit)
If $iIndex Then
$sString = StringLeft($sString, $iIndex - 1)
EndIf
EndIf
Return StringSplit($sString, $sDelimiter, BItOR($STR_ENTIRESPLIT, $STR_NOCOUNT))
EndFunc
Func _StringInsert($sString, $sInsertion, $iPosition)
Local $iLength = StringLen($sString)
$iPosition = Int($iPosition)
If $iPosition < 0 Then $iPosition = $iLength + $iPosition
If $iLength < $iPosition Or $iPosition < 0 Then Return SetError(1, 0, $sString)
Return StringLeft($sString, $iPosition) & $sInsertion & StringRight($sString, $iLength - $iPosition)
EndFunc
Func _StringProper($sString)
Local $bCapNext = True, $sChr = "", $sReturn = ""
For $i = 1 To StringLen($sString)
$sChr = StringMid($sString, $i, 1)
Select
Case $bCapNext = True
If StringRegExp($sChr, '[a-zA-ZÀ-ÿšœžŸ]') Then
$sChr = StringUpper($sChr)
$bCapNext = False
EndIf
Case Not StringRegExp($sChr, '[a-zA-ZÀ-ÿšœžŸ]')
$bCapNext = True
Case Else
$sChr = StringLower($sChr)
EndSelect
$sReturn &= $sChr
Next
Return $sReturn
EndFunc
Func _StringRepeat($sString, $iRepeatCount)
$iRepeatCount = Int($iRepeatCount)
If $iRepeatCount = 0 Then Return ""
If StringLen($sString) < 1 Or $iRepeatCount < 0 Then Return SetError(1, 0, "")
Local $sResult = ""
While $iRepeatCount > 1
If BitAND($iRepeatCount, 1) Then $sResult &= $sString
$sString &= $sString
$iRepeatCount = BitShift($iRepeatCount, 1)
WEnd
Return $sString & $sResult
EndFunc
Func _StringTitleCase($sString)
Local $bCapNext = True, $sChr = "", $sReturn = ""
For $i = 1 To StringLen($sString)
$sChr = StringMid($sString, $i, 1)
Select
Case $bCapNext = True
If StringRegExp($sChr, "[a-zA-Z\xC0-\xFF0-9]") Then
$sChr = StringUpper($sChr)
$bCapNext = False
EndIf
Case Not StringRegExp($sChr, "[a-zA-Z\xC0-\xFF'0-9]")
$bCapNext = True
Case Else
$sChr = StringLower($sChr)
EndSelect
$sReturn &= $sChr
Next
Return $sReturn
EndFunc
Func _StringToHex($sString)
Return Hex(StringToBinary($sString, $SB_UTF8))
EndFunc
Global $__g_aTimers_aTimerIDs[1][3]
Func _Timer_Diff($iTimeStamp)
Return 1000 * (__Timer_QueryPerformanceCounter() - $iTimeStamp) / __Timer_QueryPerformanceFrequency()
EndFunc
Func _Timer_GetIdleTime()
Local $tStruct = DllStructCreate("uint;dword")
DllStructSetData($tStruct, 1, DllStructGetSize($tStruct))
Local $aResult = DllCall("user32.dll", "bool", "GetLastInputInfo", "struct*", $tStruct)
If @error Or $aResult[0] = 0 Then Return SetError(@error, @extended, 0)
Local $avTicks = DllCall("kernel32.dll", "dword", "GetTickCount")
If @error Or Not $aResult[0] Then Return SetError(@error, @extended, 0)
Local $iDiff = $avTicks[0] - DllStructGetData($tStruct, 2)
If $iDiff < 0 Then Return SetExtended(1, $avTicks[0])
Return $iDiff
EndFunc
Func _Timer_GetTimerID($wParam)
Local $_iTimerID = Dec(Hex($wParam, 8)), $iMax = UBound($__g_aTimers_aTimerIDs) - 1
For $x = 1 To $iMax
If $_iTimerID = $__g_aTimers_aTimerIDs[$x][1] Then Return $__g_aTimers_aTimerIDs[$x][0]
Next
Return 0
EndFunc
Func _Timer_Init()
Return __Timer_QueryPerformanceCounter()
EndFunc
Func _Timer_KillAllTimers($hWnd)
Local $iNumTimers = $__g_aTimers_aTimerIDs[0][0]
If $iNumTimers = 0 Then Return False
Local $aResult, $hCallBack = 0
For $x = $iNumTimers To 1 Step -1
If IsHWnd($hWnd) Then
$aResult = DllCall("user32.dll", "bool", "KillTimer", "hwnd", $hWnd, "uint_ptr", $__g_aTimers_aTimerIDs[$x][1])
Else
$aResult = DllCall("user32.dll", "bool", "KillTimer", "hwnd", $hWnd, "uint_ptr", $__g_aTimers_aTimerIDs[$x][0])
EndIf
If @error Or $aResult[0] = 0 Then Return SetError(@error, @extended, False)
$hCallBack = $__g_aTimers_aTimerIDs[$x][2]
If $hCallBack <> 0 Then DllCallbackFree($hCallBack)
$__g_aTimers_aTimerIDs[0][0] -= 1
Next
ReDim $__g_aTimers_aTimerIDs[1][3]
Return True
EndFunc
Func _Timer_KillTimer($hWnd, $iTimerID)
Local $aResult[1] = [0], $hCallBack = 0, $iUBound = UBound($__g_aTimers_aTimerIDs) - 1
For $x = 1 To $iUBound
If $__g_aTimers_aTimerIDs[$x][0] = $iTimerID Then
If IsHWnd($hWnd) Then
$aResult = DllCall("user32.dll", "bool", "KillTimer", "hwnd", $hWnd, "uint_ptr", $__g_aTimers_aTimerIDs[$x][1])
Else
$aResult = DllCall("user32.dll", "bool", "KillTimer", "hwnd", $hWnd, "uint_ptr", $__g_aTimers_aTimerIDs[$x][0])
EndIf
If @error Or $aResult[0] = 0 Then Return SetError(@error, @extended, False)
$hCallBack = $__g_aTimers_aTimerIDs[$x][2]
If $hCallBack <> 0 Then DllCallbackFree($hCallBack)
For $i = $x To $iUBound - 1
$__g_aTimers_aTimerIDs[$i][0] = $__g_aTimers_aTimerIDs[$i + 1][0]
$__g_aTimers_aTimerIDs[$i][1] = $__g_aTimers_aTimerIDs[$i + 1][1]
$__g_aTimers_aTimerIDs[$i][2] = $__g_aTimers_aTimerIDs[$i + 1][2]
Next
ReDim $__g_aTimers_aTimerIDs[UBound($__g_aTimers_aTimerIDs - 1)][3]
$__g_aTimers_aTimerIDs[0][0] -= 1
ExitLoop
EndIf
Next
Return $aResult[0] <> 0
EndFunc
Func __Timer_QueryPerformanceCounter()
Local $aResult = DllCall("kernel32.dll", "bool", "QueryPerformanceCounter", "int64*", 0)
If @error Then Return SetError(@error, @extended, -1)
Return SetExtended($aResult[0], $aResult[1])
EndFunc
Func __Timer_QueryPerformanceFrequency()
Local $aResult = DllCall("kernel32.dll", "bool", "QueryPerformanceFrequency", "int64*", 0)
If @error Then Return SetError(@error, @extended, 0)
Return SetExtended($aResult[0], $aResult[1])
EndFunc
Func _Timer_SetTimer($hWnd, $iElapse = 250, $sTimerFunc = "", $iTimerID = -1)
#Au3Stripper_Ignore_Funcs=$sTimerFunc
Local $aResult[1] = [0], $pTimerFunc = 0, $hCallBack = 0, $iIndex = $__g_aTimers_aTimerIDs[0][0] + 1
If $iTimerID = -1 Then
ReDim $__g_aTimers_aTimerIDs[$iIndex + 1][3]
$__g_aTimers_aTimerIDs[0][0] = $iIndex
$iTimerID = $iIndex + 1000
For $x = 1 To $iIndex
If $__g_aTimers_aTimerIDs[$x][0] = $iTimerID Then
$iTimerID = $iTimerID + 1
$x = 0
EndIf
Next
If $sTimerFunc <> "" Then
$hCallBack = DllCallbackRegister($sTimerFunc, "none", "hwnd;uint;uint_ptr;dword")
If $hCallBack = 0 Then Return SetError(-1, -1, 0)
$pTimerFunc = DllCallbackGetPtr($hCallBack)
If $pTimerFunc = 0 Then Return SetError(-1, -1, 0)
EndIf
$aResult = DllCall("user32.dll", "uint_ptr", "SetTimer", "hwnd", $hWnd, "uint_ptr", $iTimerID, "uint", $iElapse, "ptr", $pTimerFunc)
If @error Or $aResult[0] = 0 Then Return SetError(@error, @extended, 0)
$__g_aTimers_aTimerIDs[$iIndex][0] = $aResult[0]
$__g_aTimers_aTimerIDs[$iIndex][1] = $iTimerID
$__g_aTimers_aTimerIDs[$iIndex][2] = $hCallBack
Else
For $x = 1 To $iIndex - 1
If $__g_aTimers_aTimerIDs[$x][0] = $iTimerID Then
If IsHWnd($hWnd) Then $iTimerID = $__g_aTimers_aTimerIDs[$x][1]
$hCallBack = $__g_aTimers_aTimerIDs[$x][2]
If $hCallBack <> 0 Then
$pTimerFunc = DllCallbackGetPtr($hCallBack)
If $pTimerFunc = 0 Then Return SetError(-1, -1, 0)
EndIf
$aResult = DllCall("user32.dll", "uint_ptr", "SetTimer", "hwnd", $hWnd, "uint_ptr", $iTimerID, "uint", $iElapse, "ptr", $pTimerFunc)
If @error Or $aResult[0] = 0 Then Return SetError(@error, @extended, 0)
ExitLoop
EndIf
Next
EndIf
Return $aResult[0]
EndFunc
Global Const $WC_ANIMATE = 'SysAnimate32'
Global Const $WC_BUTTON = 'Button'
Global Const $WC_COMBOBOX = 'ComboBox'
Global Const $WC_COMBOBOXEX = 'ComboBoxEx32'
Global Const $WC_DATETIMEPICK = 'SysDateTimePick32'
Global Const $WC_EDIT = 'Edit'
Global Const $WC_HEADER = 'SysHeader32'
Global Const $WC_HOTKEY = 'msctls_hotkey32'
Global Const $WC_IPADDRESS = 'SysIPAddress32'
Global Const $WC_LINK = 'SysLink'
Global Const $WC_LISTBOX = 'ListBox'
Global Const $WC_LISTVIEW = 'SysListView32'
Global Const $WC_MONTHCAL = 'SysMonthCal32'
Global Const $WC_NATIVEFONTCTL = 'NativeFontCtl'
Global Const $WC_PAGESCROLLER = 'SysPager'
Global Const $WC_PROGRESS = 'msctls_progress32'
Global Const $WC_REBAR = 'ReBarWindow32'
Global Const $WC_SCROLLBAR = 'ScrollBar'
Global Const $WC_STATIC = 'Static'
Global Const $WC_STATUSBAR = 'msctls_statusbar32'
Global Const $WC_TABCONTROL = 'SysTabControl32'
Global Const $WC_TOOLBAR = 'ToolbarWindow32'
Global Const $WC_TOOLTIPS = 'tooltips_class32'
Global Const $WC_TRACKBAR = 'msctls_trackbar32'
Global Const $WC_TREEVIEW = 'SysTreeView32'
Global Const $WC_UPDOWN = 'msctls_updown32'
Global Const $WS_OVERLAPPED = 0
Global Const $WS_TILED = $WS_OVERLAPPED
Global Const $WS_MAXIMIZEBOX = 0x00010000
Global Const $WS_MINIMIZEBOX = 0x00020000
Global Const $WS_TABSTOP = 0x00010000
Global Const $WS_GROUP = 0x00020000
Global Const $WS_SIZEBOX = 0x00040000
Global Const $WS_THICKFRAME = $WS_SIZEBOX
Global Const $WS_SYSMENU = 0x00080000
Global Const $WS_HSCROLL = 0x00100000
Global Const $WS_VSCROLL = 0x00200000
Global Const $WS_DLGFRAME = 0x00400000
Global Const $WS_BORDER = 0x00800000
Global Const $WS_CAPTION = 0x00C00000
Global Const $WS_OVERLAPPEDWINDOW = BitOR($WS_CAPTION, $WS_MAXIMIZEBOX, $WS_MINIMIZEBOX, $WS_OVERLAPPED, $WS_SYSMENU, $WS_THICKFRAME)
Global Const $WS_TILEDWINDOW = $WS_OVERLAPPEDWINDOW
Global Const $WS_MAXIMIZE = 0x01000000
Global Const $WS_CLIPCHILDREN = 0x02000000
Global Const $WS_CLIPSIBLINGS = 0x04000000
Global Const $WS_DISABLED = 0x08000000
Global Const $WS_VISIBLE = 0x10000000
Global Const $WS_MINIMIZE = 0x20000000
Global Const $WS_ICONIC = $WS_MINIMIZE
Global Const $WS_CHILD = 0x40000000
Global Const $WS_CHILDWINDOW = $WS_CHILD
Global Const $WS_POPUP = 0x80000000
Global Const $WS_POPUPWINDOW = 0x80880000
Global Const $DS_3DLOOK = 0x0004
Global Const $DS_ABSALIGN = 0x0001
Global Const $DS_CENTER = 0x0800
Global Const $DS_CENTERMOUSE = 0x1000
Global Const $DS_CONTEXTHELP = 0x2000
Global Const $DS_CONTROL = 0x0400
Global Const $DS_FIXEDSYS = 0x0008
Global Const $DS_LOCALEDIT = 0x0020
Global Const $DS_MODALFRAME = 0x0080
Global Const $DS_NOFAILCREATE = 0x0010
Global Const $DS_NOIDLEMSG = 0x0100
Global Const $DS_SETFONT = 0x0040
Global Const $DS_SETFOREGROUND = 0x0200
Global Const $DS_SHELLFONT = BitOR($DS_FIXEDSYS, $DS_SETFONT)
Global Const $DS_SYSMODAL = 0x0002
Global Const $WS_EX_ACCEPTFILES = 0x00000010
Global Const $WS_EX_APPWINDOW = 0x00040000
Global Const $WS_EX_COMPOSITED = 0x02000000
Global Const $WS_EX_CONTROLPARENT = 0x10000
Global Const $WS_EX_CLIENTEDGE = 0x00000200
Global Const $WS_EX_CONTEXTHELP = 0x00000400
Global Const $WS_EX_DLGMODALFRAME = 0x00000001
Global Const $WS_EX_LAYERED = 0x00080000
Global Const $WS_EX_LAYOUTRTL = 0x400000
Global Const $WS_EX_LEFT = 0x00000000
Global Const $WS_EX_LEFTSCROLLBAR = 0x00004000
Global Const $WS_EX_LTRREADING = 0x00000000
Global Const $WS_EX_MDICHILD = 0x00000040
Global Const $WS_EX_NOACTIVATE = 0x08000000
Global Const $WS_EX_NOINHERITLAYOUT = 0x00100000
Global Const $WS_EX_NOPARENTNOTIFY = 0x00000004
Global Const $WS_EX_RIGHT = 0x00001000
Global Const $WS_EX_RIGHTSCROLLBAR = 0x00000000
Global Const $WS_EX_RTLREADING = 0x2000
Global Const $WS_EX_STATICEDGE = 0x00020000
Global Const $WS_EX_TOOLWINDOW = 0x00000080
Global Const $WS_EX_TOPMOST = 0x00000008
Global Const $WS_EX_TRANSPARENT = 0x00000020
Global Const $WS_EX_WINDOWEDGE = 0x00000100
Global Const $WS_EX_OVERLAPPEDWINDOW = BitOR($WS_EX_CLIENTEDGE, $WS_EX_WINDOWEDGE)
Global Const $WS_EX_PALETTEWINDOW = BitOR($WS_EX_TOOLWINDOW, $WS_EX_TOPMOST, $WS_EX_WINDOWEDGE)
Global Const $WM_NULL = 0x0000
Global Const $WM_CREATE = 0x0001
Global Const $WM_DESTROY = 0x0002
Global Const $WM_MOVE = 0x0003
Global Const $WM_SIZEWAIT = 0x0004
Global Const $WM_SIZE = 0x0005
Global Const $WM_ACTIVATE = 0x0006
Global Const $WM_SETFOCUS = 0x0007
Global Const $WM_KILLFOCUS = 0x0008
Global Const $WM_SETVISIBLE = 0x0009
Global Const $WM_ENABLE = 0x000A
Global Const $WM_SETREDRAW = 0x000B
Global Const $WM_SETTEXT = 0x000C
Global Const $WM_GETTEXT = 0x000D
Global Const $WM_GETTEXTLENGTH = 0x000E
Global Const $WM_PAINT = 0x000F
Global Const $WM_CLOSE = 0x0010
Global Const $WM_QUERYENDSESSION = 0x0011
Global Const $WM_QUIT = 0x0012
Global Const $WM_ERASEBKGND = 0x0014
Global Const $WM_QUERYOPEN = 0x0013
Global Const $WM_SYSCOLORCHANGE = 0x0015
Global Const $WM_ENDSESSION = 0x0016
Global Const $WM_SYSTEMERROR = 0x0017
Global Const $WM_SHOWWINDOW = 0x0018
Global Const $WM_CTLCOLOR = 0x0019
Global Const $WM_SETTINGCHANGE = 0x001A
Global Const $WM_WININICHANGE = 0x001A
Global Const $WM_DEVMODECHANGE = 0x001B
Global Const $WM_ACTIVATEAPP = 0x001C
Global Const $WM_FONTCHANGE = 0x001D
Global Const $WM_TIMECHANGE = 0x001E
Global Const $WM_CANCELMODE = 0x001F
Global Const $WM_SETCURSOR = 0x0020
Global Const $WM_MOUSEACTIVATE = 0x0021
Global Const $WM_CHILDACTIVATE = 0x0022
Global Const $WM_QUEUESYNC = 0x0023
Global Const $WM_GETMINMAXINFO = 0x0024
Global Const $WM_LOGOFF = 0x0025
Global Const $WM_PAINTICON = 0x0026
Global Const $WM_ICONERASEBKGND = 0x0027
Global Const $WM_NEXTDLGCTL = 0x0028
Global Const $WM_ALTTABACTIVE = 0x0029
Global Const $WM_SPOOLERSTATUS = 0x002A
Global Const $WM_DRAWITEM = 0x002B
Global Const $WM_MEASUREITEM = 0x002C
Global Const $WM_DELETEITEM = 0x002D
Global Const $WM_VKEYTOITEM = 0x002E
Global Const $WM_CHARTOITEM = 0x002F
Global Const $WM_SETFONT = 0x0030
Global Const $WM_GETFONT = 0x0031
Global Const $WM_SETHOTKEY = 0x0032
Global Const $WM_GETHOTKEY = 0x0033
Global Const $WM_FILESYSCHANGE = 0x0034
Global Const $WM_ISACTIVEICON = 0x0035
Global Const $WM_QUERYPARKICON = 0x0036
Global Const $WM_QUERYDRAGICON = 0x0037
Global Const $WM_WINHELP = 0x0038
Global Const $WM_COMPAREITEM = 0x0039
Global Const $WM_FULLSCREEN = 0x003A
Global Const $WM_CLIENTSHUTDOWN = 0x003B
Global Const $WM_DDEMLEVENT = 0x003C
Global Const $WM_GETOBJECT = 0x003D
Global Const $WM_CALCSCROLL = 0x003F
Global Const $WM_TESTING = 0x0040
Global Const $WM_COMPACTING = 0x0041
Global Const $WM_OTHERWINDOWCREATED = 0x0042
Global Const $WM_OTHERWINDOWDESTROYED = 0x0043
Global Const $WM_COMMNOTIFY = 0x0044
Global Const $WM_MEDIASTATUSCHANGE = 0x0045
Global Const $WM_WINDOWPOSCHANGING = 0x0046
Global Const $WM_WINDOWPOSCHANGED = 0x0047
Global Const $WM_POWER = 0x0048
Global Const $WM_COPYGLOBALDATA = 0x0049
Global Const $WM_COPYDATA = 0x004A
Global Const $WM_CANCELJOURNAL = 0x004B
Global Const $WM_LOGONNOTIFY = 0x004C
Global Const $WM_KEYF1 = 0x004D
Global Const $WM_NOTIFY = 0x004E
Global Const $WM_ACCESS_WINDOW = 0x004F
Global Const $WM_INPUTLANGCHANGEREQUEST = 0x0050
Global Const $WM_INPUTLANGCHANGE = 0x0051
Global Const $WM_TCARD = 0x0052
Global Const $WM_HELP = 0x0053
Global Const $WM_USERCHANGED = 0x0054
Global Const $WM_NOTIFYFORMAT = 0x0055
Global Const $WM_QM_ACTIVATE = 0x0060
Global Const $WM_HOOK_DO_CALLBACK = 0x0061
Global Const $WM_SYSCOPYDATA = 0x0062
Global Const $WM_FINALDESTROY = 0x0070
Global Const $WM_MEASUREITEM_CLIENTDATA = 0x0071
Global Const $WM_CONTEXTMENU = 0x007B
Global Const $WM_STYLECHANGING = 0x007C
Global Const $WM_STYLECHANGED = 0x007D
Global Const $WM_DISPLAYCHANGE = 0x007E
Global Const $WM_GETICON = 0x007F
Global Const $WM_SETICON = 0x0080
Global Const $WM_NCCREATE = 0x0081
Global Const $WM_NCDESTROY = 0x0082
Global Const $WM_NCCALCSIZE = 0x0083
Global Const $WM_NCHITTEST = 0x0084
Global Const $WM_NCPAINT = 0x0085
Global Const $WM_NCACTIVATE = 0x0086
Global Const $WM_GETDLGCODE = 0x0087
Global Const $WM_SYNCPAINT = 0x0088
Global Const $WM_SYNCTASK = 0x0089
Global Const $WM_KLUDGEMINRECT = 0x008B
Global Const $WM_LPKDRAWSWITCHWND = 0x008C
Global Const $WM_UAHDESTROYWINDOW = 0x0090
Global Const $WM_UAHDRAWMENU = 0x0091
Global Const $WM_UAHDRAWMENUITEM = 0x0092
Global Const $WM_UAHINITMENU = 0x0093
Global Const $WM_UAHMEASUREMENUITEM = 0x0094
Global Const $WM_UAHNCPAINTMENUPOPUP = 0x0095
Global Const $WM_NCMOUSEMOVE = 0x00A0
Global Const $WM_NCLBUTTONDOWN = 0x00A1
Global Const $WM_NCLBUTTONUP = 0x00A2
Global Const $WM_NCLBUTTONDBLCLK = 0x00A3
Global Const $WM_NCRBUTTONDOWN = 0x00A4
Global Const $WM_NCRBUTTONUP = 0x00A5
Global Const $WM_NCRBUTTONDBLCLK = 0x00A6
Global Const $WM_NCMBUTTONDOWN = 0x00A7
Global Const $WM_NCMBUTTONUP = 0x00A8
Global Const $WM_NCMBUTTONDBLCLK = 0x00A9
Global Const $WM_NCXBUTTONDOWN = 0x00AB
Global Const $WM_NCXBUTTONUP = 0x00AC
Global Const $WM_NCXBUTTONDBLCLK = 0x00AD
Global Const $WM_NCUAHDRAWCAPTION = 0x00AE
Global Const $WM_NCUAHDRAWFRAME = 0x00AF
Global Const $WM_INPUT_DEVICE_CHANGE = 0x00FE
Global Const $WM_INPUT = 0x00FF
Global Const $WM_KEYDOWN = 0x0100
Global Const $WM_KEYFIRST = 0x0100
Global Const $WM_KEYUP = 0x0101
Global Const $WM_CHAR = 0x0102
Global Const $WM_DEADCHAR = 0x0103
Global Const $WM_SYSKEYDOWN = 0x0104
Global Const $WM_SYSKEYUP = 0x0105
Global Const $WM_SYSCHAR = 0x0106
Global Const $WM_SYSDEADCHAR = 0x0107
Global Const $WM_YOMICHAR = 0x0108
Global Const $WM_KEYLAST = 0x0109
Global Const $WM_UNICHAR = 0x0109
Global Const $WM_CONVERTREQUEST = 0x010A
Global Const $WM_CONVERTRESULT = 0x010B
Global Const $WM_IM_INFO = 0x010C
Global Const $WM_IME_STARTCOMPOSITION = 0x010D
Global Const $WM_IME_ENDCOMPOSITION = 0x010E
Global Const $WM_IME_COMPOSITION = 0x010F
Global Const $WM_IME_KEYLAST = 0x010F
Global Const $WM_INITDIALOG = 0x0110
Global Const $WM_COMMAND = 0x0111
Global Const $WM_SYSCOMMAND = 0x0112
Global Const $WM_TIMER = 0x0113
Global Const $WM_HSCROLL = 0x0114
Global Const $WM_VSCROLL = 0x0115
Global Const $WM_INITMENU = 0x0116
Global Const $WM_INITMENUPOPUP = 0x0117
Global Const $WM_SYSTIMER = 0x0118
Global Const $WM_GESTURE = 0x0119
Global Const $WM_GESTURENOTIFY = 0x011A
Global Const $WM_GESTUREINPUT = 0x011B
Global Const $WM_GESTURENOTIFIED = 0x011C
Global Const $WM_MENUSELECT = 0x011F
Global Const $WM_MENUCHAR = 0x0120
Global Const $WM_ENTERIDLE = 0x0121
Global Const $WM_MENURBUTTONUP = 0x0122
Global Const $WM_MENUDRAG = 0x0123
Global Const $WM_MENUGETOBJECT = 0x0124
Global Const $WM_UNINITMENUPOPUP = 0x0125
Global Const $WM_MENUCOMMAND = 0x0126
Global Const $WM_CHANGEUISTATE = 0x0127
Global Const $WM_UPDATEUISTATE = 0x0128
Global Const $WM_QUERYUISTATE = 0x0129
Global Const $WM_LBTRACKPOINT = 0x0131
Global Const $WM_CTLCOLORMSGBOX = 0x0132
Global Const $WM_CTLCOLOREDIT = 0x0133
Global Const $WM_CTLCOLORLISTBOX = 0x0134
Global Const $WM_CTLCOLORBTN = 0x0135
Global Const $WM_CTLCOLORDLG = 0x0136
Global Const $WM_CTLCOLORSCROLLBAR = 0x0137
Global Const $WM_CTLCOLORSTATIC = 0x0138
Global Const $MN_GETHMENU = 0x01E1
Global Const $WM_PARENTNOTIFY = 0x0210
Global Const $WM_ENTERMENULOOP = 0x0211
Global Const $WM_EXITMENULOOP = 0x0212
Global Const $WM_NEXTMENU = 0x0213
Global Const $WM_SIZING = 0x0214
Global Const $WM_CAPTURECHANGED = 0x0215
Global Const $WM_MOVING = 0x0216
Global Const $WM_POWERBROADCAST = 0x0218
Global Const $WM_DEVICECHANGE = 0x0219
Global Const $WM_MDICREATE = 0x0220
Global Const $WM_MDIDESTROY = 0x0221
Global Const $WM_MDIACTIVATE = 0x0222
Global Const $WM_MDIRESTORE = 0x0223
Global Const $WM_MDINEXT = 0x0224
Global Const $WM_MDIMAXIMIZE = 0x0225
Global Const $WM_MDITILE = 0x0226
Global Const $WM_MDICASCADE = 0x0227
Global Const $WM_MDIICONARRANGE = 0x0228
Global Const $WM_MDIGETACTIVE = 0x0229
Global Const $WM_DROPOBJECT = 0x022A
Global Const $WM_QUERYDROPOBJECT = 0x022B
Global Const $WM_BEGINDRAG = 0x022C
Global Const $WM_DRAGLOOP = 0x022D
Global Const $WM_DRAGSELECT = 0x022E
Global Const $WM_DRAGMOVE = 0x022F
Global Const $WM_MDISETMENU = 0x0230
Global Const $WM_ENTERSIZEMOVE = 0x0231
Global Const $WM_EXITSIZEMOVE = 0x0232
Global Const $WM_DROPFILES = 0x0233
Global Const $WM_MDIREFRESHMENU = 0x0234
Global Const $WM_TOUCH = 0x0240
Global Const $WM_IME_SETCONTEXT = 0x0281
Global Const $WM_IME_NOTIFY = 0x0282
Global Const $WM_IME_CONTROL = 0x0283
Global Const $WM_IME_COMPOSITIONFULL = 0x0284
Global Const $WM_IME_SELECT = 0x0285
Global Const $WM_IME_CHAR = 0x0286
Global Const $WM_IME_SYSTEM = 0x0287
Global Const $WM_IME_REQUEST = 0x0288
Global Const $WM_IME_KEYDOWN = 0x0290
Global Const $WM_IME_KEYUP = 0x0291
Global Const $WM_NCMOUSEHOVER = 0x02A0
Global Const $WM_MOUSEHOVER = 0x02A1
Global Const $WM_NCMOUSELEAVE = 0x02A2
Global Const $WM_MOUSELEAVE = 0x02A3
Global Const $WM_WTSSESSION_CHANGE = 0x02B1
Global Const $WM_TABLET_FIRST = 0x02C0
Global Const $WM_TABLET_LAST = 0x02DF
Global Const $WM_CUT = 0x0300
Global Const $WM_COPY = 0x0301
Global Const $WM_PASTE = 0x0302
Global Const $WM_CLEAR = 0x0303
Global Const $WM_UNDO = 0x0304
Global Const $WM_PALETTEISCHANGING = 0x0310
Global Const $WM_HOTKEY = 0x0312
Global Const $WM_PALETTECHANGED = 0x0311
Global Const $WM_SYSMENU = 0x0313
Global Const $WM_HOOKMSG = 0x0314
Global Const $WM_EXITPROCESS = 0x0315
Global Const $WM_WAKETHREAD = 0x0316
Global Const $WM_PRINT = 0x0317
Global Const $WM_PRINTCLIENT = 0x0318
Global Const $WM_APPCOMMAND = 0x0319
Global Const $WM_QUERYNEWPALETTE = 0x030F
Global Const $WM_THEMECHANGED = 0x031A
Global Const $WM_UAHINIT = 0x031B
Global Const $WM_DESKTOPNOTIFY = 0x031C
Global Const $WM_CLIPBOARDUPDATE = 0x031D
Global Const $WM_DWMCOMPOSITIONCHANGED = 0x031E
Global Const $WM_DWMNCRENDERINGCHANGED = 0x031F
Global Const $WM_DWMCOLORIZATIONCOLORCHANGED = 0x0320
Global Const $WM_DWMWINDOWMAXIMIZEDCHANGE = 0x0321
Global Const $WM_DWMEXILEFRAME = 0x0322
Global Const $WM_DWMSENDICONICTHUMBNAIL = 0x0323
Global Const $WM_MAGNIFICATION_STARTED = 0x0324
Global Const $WM_MAGNIFICATION_ENDED = 0x0325
Global Const $WM_DWMSENDICONICLIVEPREVIEWBITMAP = 0x0326
Global Const $WM_DWMTHUMBNAILSIZECHANGED = 0x0327
Global Const $WM_MAGNIFICATION_OUTPUT = 0x0328
Global Const $WM_MEASURECONTROL = 0x0330
Global Const $WM_GETACTIONTEXT = 0x0331
Global Const $WM_FORWARDKEYDOWN = 0x0333
Global Const $WM_FORWARDKEYUP = 0x0334
Global Const $WM_GETTITLEBARINFOEX = 0x033F
Global Const $WM_NOTIFYWOW = 0x0340
Global Const $WM_HANDHELDFIRST = 0x0358
Global Const $WM_HANDHELDLAST = 0x035F
Global Const $WM_AFXFIRST = 0x0360
Global Const $WM_AFXLAST = 0x037F
Global Const $WM_PENWINFIRST = 0x0380
Global Const $WM_PENWINLAST = 0x038F
Global Const $WM_DDE_INITIATE = 0x03E0
Global Const $WM_DDE_TERMINATE = 0x03E1
Global Const $WM_DDE_ADVISE = 0x03E2
Global Const $WM_DDE_UNADVISE = 0x03E3
Global Const $WM_DDE_ACK = 0x03E4
Global Const $WM_DDE_DATA = 0x03E5
Global Const $WM_DDE_REQUEST = 0x03E6
Global Const $WM_DDE_POKE = 0x03E7
Global Const $WM_DDE_EXECUTE = 0x03E8
Global Const $WM_DBNOTIFICATION = 0x03FD
Global Const $WM_NETCONNECT = 0x03FE
Global Const $WM_HIBERNATE = 0x03FF
Global Const $WM_USER = 0x0400
Global Const $WM_APP = 0x8000
Global Const $NM_FIRST = 0
Global Const $NM_OUTOFMEMORY = $NM_FIRST - 1
Global Const $NM_CLICK = $NM_FIRST - 2
Global Const $NM_DBLCLK = $NM_FIRST - 3
Global Const $NM_RETURN = $NM_FIRST - 4
Global Const $NM_RCLICK = $NM_FIRST - 5
Global Const $NM_RDBLCLK = $NM_FIRST - 6
Global Const $NM_SETFOCUS = $NM_FIRST - 7
Global Const $NM_KILLFOCUS = $NM_FIRST - 8
Global Const $NM_CUSTOMDRAW = $NM_FIRST - 12
Global Const $NM_HOVER = $NM_FIRST - 13
Global Const $NM_NCHITTEST = $NM_FIRST - 14
Global Const $NM_KEYDOWN = $NM_FIRST - 15
Global Const $NM_RELEASEDCAPTURE = $NM_FIRST - 16
Global Const $NM_SETCURSOR = $NM_FIRST - 17
Global Const $NM_CHAR = $NM_FIRST - 18
Global Const $NM_TOOLTIPSCREATED = $NM_FIRST - 19
Global Const $NM_LDOWN = $NM_FIRST - 20
Global Const $NM_RDOWN = $NM_FIRST - 21
Global Const $NM_THEMECHANGED = $NM_FIRST - 22
Global Const $WM_MOUSEFIRST = 0x0200
Global Const $WM_MOUSEMOVE = 0x0200
Global Const $WM_LBUTTONDOWN = 0x0201
Global Const $WM_LBUTTONUP = 0x0202
Global Const $WM_LBUTTONDBLCLK = 0x0203
Global Const $WM_RBUTTONDOWN = 0x0204
Global Const $WM_RBUTTONUP = 0x0205
Global Const $WM_RBUTTONDBLCLK = 0x0206
Global Const $WM_MBUTTONDOWN = 0x0207
Global Const $WM_MBUTTONUP = 0x0208
Global Const $WM_MBUTTONDBLCLK = 0x0209
Global Const $WM_MOUSEWHEEL = 0x020A
Global Const $WM_XBUTTONDOWN = 0x020B
Global Const $WM_XBUTTONUP = 0x020C
Global Const $WM_XBUTTONDBLCLK = 0x020D
Global Const $WM_MOUSEHWHEEL = 0x020E
Global Const $PS_SOLID = 0
Global Const $PS_DASH = 1
Global Const $PS_DOT = 2
Global Const $PS_DASHDOT = 3
Global Const $PS_DASHDOTDOT = 4
Global Const $PS_NULL = 5
Global Const $PS_INSIDEFRAME = 6
Global Const $PS_USERSTYLE = 7
Global Const $PS_ALTERNATE = 8
Global Const $PS_ENDCAP_ROUND = 0x00000000
Global Const $PS_ENDCAP_SQUARE = 0x00000100
Global Const $PS_ENDCAP_FLAT = 0x00000200
Global Const $PS_JOIN_BEVEL = 0x00001000
Global Const $PS_JOIN_MITER = 0x00002000
Global Const $PS_JOIN_ROUND = 0x00000000
Global Const $PS_GEOMETRIC = 0x00010000
Global Const $PS_COSMETIC = 0x00000000
Global Const $LWA_ALPHA = 0x2
Global Const $LWA_COLORKEY = 0x1
Global Const $RGN_AND = 1
Global Const $RGN_OR = 2
Global Const $RGN_XOR = 3
Global Const $RGN_DIFF = 4
Global Const $RGN_COPY = 5
Global Const $ERRORREGION = 0
Global Const $NULLREGION = 1
Global Const $SIMPLEREGION = 2
Global Const $COMPLEXREGION = 3
Global Const $TRANSPARENT = 1
Global Const $OPAQUE = 2
Global Const $CCM_FIRST = 0x2000
Global Const $CCM_GETUNICODEFORMAT = ($CCM_FIRST + 6)
Global Const $CCM_SETUNICODEFORMAT = ($CCM_FIRST + 5)
Global Const $CCM_SETBKCOLOR = $CCM_FIRST + 1
Global Const $CCM_SETCOLORSCHEME = $CCM_FIRST + 2
Global Const $CCM_GETCOLORSCHEME = $CCM_FIRST + 3
Global Const $CCM_GETDROPTARGET = $CCM_FIRST + 4
Global Const $CCM_SETWINDOWTHEME = $CCM_FIRST + 11
Global Const $GA_PARENT = 1
Global Const $GA_ROOT = 2
Global Const $GA_ROOTOWNER = 3
Global Const $SM_CXSCREEN = 0
Global Const $SM_CYSCREEN = 1
Global Const $SM_CXVSCROLL = 2
Global Const $SM_CYHSCROLL = 3
Global Const $SM_CYCAPTION = 4
Global Const $SM_CXBORDER = 5
Global Const $SM_CYBORDER = 6
Global Const $SM_CXFIXEDFRAME = 7
Global Const $SM_CXDLGFRAME = $SM_CXFIXEDFRAME
Global Const $SM_CYFIXEDFRAME = 8
Global Const $SM_CYDLGFRAME = $SM_CYFIXEDFRAME
Global Const $SM_CYVTHUMB = 9
Global Const $SM_CXHTHUMB = 10
Global Const $SM_CXICON = 11
Global Const $SM_CYICON = 12
Global Const $SM_CXCURSOR = 13
Global Const $SM_CYCURSOR = 14
Global Const $SM_CYMENU = 15
Global Const $SM_CXFULLSCREEN = 16
Global Const $SM_CYFULLSCREEN = 17
Global Const $SM_CYKANJIWINDOW = 18
Global Const $SM_MOUSEPRESENT = 19
Global Const $SM_CYVSCROLL = 20
Global Const $SM_CXHSCROLL = 21
Global Const $SM_DEBUG = 22
Global Const $SM_SWAPBUTTON = 23
Global Const $SM_RESERVED1 = 24
Global Const $SM_RESERVED2 = 25
Global Const $SM_RESERVED3 = 26
Global Const $SM_RESERVED4 = 27
Global Const $SM_CXMIN = 28
Global Const $SM_CYMIN = 29
Global Const $SM_CXSIZE = 30
Global Const $SM_CYSIZE = 31
Global Const $SM_CXSIZEFRAME = 32
Global Const $SM_CXFRAME = $SM_CXSIZEFRAME
Global Const $SM_CYSIZEFRAME = 33
Global Const $SM_CYFRAME = $SM_CYSIZEFRAME
Global Const $SM_CXMINTRACK = 34
Global Const $SM_CYMINTRACK = 35
Global Const $SM_CXDOUBLECLK = 36
Global Const $SM_CYDOUBLECLK = 37
Global Const $SM_CXICONSPACING = 38
Global Const $SM_CYICONSPACING = 39
Global Const $SM_MENUDROPALIGNMENT = 40
Global Const $SM_PENWINDOWS = 41
Global Const $SM_DBCSENABLED = 42
Global Const $SM_CMOUSEBUTTONS = 43
Global Const $SM_SECURE = 44
Global Const $SM_CXEDGE = 45
Global Const $SM_CYEDGE = 46
Global Const $SM_CXMINSPACING = 47
Global Const $SM_CYMINSPACING = 48
Global Const $SM_CXSMICON = 49
Global Const $SM_CYSMICON = 50
Global Const $SM_CYSMCAPTION = 51
Global Const $SM_CXSMSIZE = 52
Global Const $SM_CYSMSIZE = 53
Global Const $SM_CXMENUSIZE = 54
Global Const $SM_CYMENUSIZE = 55
Global Const $SM_ARRANGE = 56
Global Const $SM_CXMINIMIZED = 57
Global Const $SM_CYMINIMIZED = 58
Global Const $SM_CXMAXTRACK = 59
Global Const $SM_CYMAXTRACK = 60
Global Const $SM_CXMAXIMIZED = 61
Global Const $SM_CYMAXIMIZED = 62
Global Const $SM_NETWORK = 63
Global Const $SM_CLEANBOOT = 67
Global Const $SM_CXDRAG = 68
Global Const $SM_CYDRAG = 69
Global Const $SM_SHOWSOUNDS = 70
Global Const $SM_CXMENUCHECK = 71
Global Const $SM_CYMENUCHECK = 72
Global Const $SM_SLOWMACHINE = 73
Global Const $SM_MIDEASTENABLED = 74
Global Const $SM_MOUSEWHEELPRESENT = 75
Global Const $SM_XVIRTUALSCREEN = 76
Global Const $SM_YVIRTUALSCREEN = 77
Global Const $SM_CXVIRTUALSCREEN = 78
Global Const $SM_CYVIRTUALSCREEN = 79
Global Const $SM_CMONITORS = 80
Global Const $SM_SAMEDISPLAYFORMAT = 81
Global Const $SM_IMMENABLED = 82
Global Const $SM_CXFOCUSBORDER = 83
Global Const $SM_CYFOCUSBORDER = 84
Global Const $SM_TABLETPC = 86
Global Const $SM_MEDIACENTER = 87
Global Const $SM_STARTER = 88
Global Const $SM_SERVERR2 = 89
Global Const $SM_CMETRICS = 90
Global Const $SM_REMOTESESSION = 0x1000
Global Const $SM_SHUTTINGDOWN = 0x2000
Global Const $SM_REMOTECONTROL = 0x2001
Global Const $SM_CARETBLINKINGENABLED = 0x2002
Global Const $BLACKNESS = 0x00000042
Global Const $CAPTUREBLT = 0X40000000
Global Const $DSTINVERT = 0x00550009
Global Const $MERGECOPY = 0x00C000CA
Global Const $MERGEPAINT = 0x00BB0226
Global Const $NOMIRRORBITMAP = 0X80000000
Global Const $NOTSRCCOPY = 0x00330008
Global Const $NOTSRCERASE = 0x001100A6
Global Const $PATCOPY = 0x00F00021
Global Const $PATINVERT = 0x005A0049
Global Const $PATPAINT = 0x00FB0A09
Global Const $SRCAND = 0x008800C6
Global Const $SRCCOPY = 0x00CC0020
Global Const $SRCERASE = 0x00440328
Global Const $SRCINVERT = 0x00660046
Global Const $SRCPAINT = 0x00EE0086
Global Const $WHITENESS = 0x00FF0062
Global Const $DT_BOTTOM = 0x8
Global Const $DT_CALCRECT = 0x400
Global Const $DT_CENTER = 0x1
Global Const $DT_EDITCONTROL = 0x2000
Global Const $DT_END_ELLIPSIS = 0x8000
Global Const $DT_EXPANDTABS = 0x40
Global Const $DT_EXTERNALLEADING = 0x200
Global Const $DT_HIDEPREFIX = 0x100000
Global Const $DT_INTERNAL = 0x1000
Global Const $DT_LEFT = 0x0
Global Const $DT_MODIFYSTRING = 0x10000
Global Const $DT_NOCLIP = 0x100
Global Const $DT_NOFULLWIDTHCHARBREAK = 0x80000
Global Const $DT_NOPREFIX = 0x800
Global Const $DT_PATH_ELLIPSIS = 0x4000
Global Const $DT_PREFIXONLY = 0x200000
Global Const $DT_RIGHT = 0x2
Global Const $DT_RTLREADING = 0x20000
Global Const $DT_SINGLELINE = 0x20
Global Const $DT_TABSTOP = 0x80
Global Const $DT_TOP = 0x0
Global Const $DT_VCENTER = 0x4
Global Const $DT_WORDBREAK = 0x10
Global Const $DT_WORD_ELLIPSIS = 0x40000
Global Const $RDW_ERASE = 0x0004
Global Const $RDW_FRAME = 0x0400
Global Const $RDW_INTERNALPAINT = 0x0002
Global Const $RDW_INVALIDATE = 0x0001
Global Const $RDW_NOERASE = 0x0020
Global Const $RDW_NOFRAME = 0x0800
Global Const $RDW_NOINTERNALPAINT = 0x0010
Global Const $RDW_VALIDATE = 0x0008
Global Const $RDW_ERASENOW = 0x0200
Global Const $RDW_UPDATENOW = 0x0100
Global Const $RDW_ALLCHILDREN = 0x0080
Global Const $RDW_NOCHILDREN = 0x0040
Global Const $WM_RENDERFORMAT = 0x0305
Global Const $WM_RENDERALLFORMATS = 0x0306
Global Const $WM_DESTROYCLIPBOARD = 0x0307
Global Const $WM_DRAWCLIPBOARD = 0x0308
Global Const $WM_PAINTCLIPBOARD = 0x0309
Global Const $WM_VSCROLLCLIPBOARD = 0x030A
Global Const $WM_SIZECLIPBOARD = 0x030B
Global Const $WM_ASKCBFORMATNAME = 0x030C
Global Const $WM_CHANGECBCHAIN = 0x030D
Global Const $WM_HSCROLLCLIPBOARD = 0x030E
Global Const $HTERROR = -2
Global Const $HTTRANSPARENT = -1
Global Const $HTNOWHERE = 0
Global Const $HTCLIENT = 1
Global Const $HTCAPTION = 2
Global Const $HTSYSMENU = 3
Global Const $HTGROWBOX = 4
Global Const $HTSIZE = $HTGROWBOX
Global Const $HTMENU = 5
Global Const $HTHSCROLL = 6
Global Const $HTVSCROLL = 7
Global Const $HTMINBUTTON = 8
Global Const $HTMAXBUTTON = 9
Global Const $HTLEFT = 10
Global Const $HTRIGHT = 11
Global Const $HTTOP = 12
Global Const $HTTOPLEFT = 13
Global Const $HTTOPRIGHT = 14
Global Const $HTBOTTOM = 15
Global Const $HTBOTTOMLEFT = 16
Global Const $HTBOTTOMRIGHT = 17
Global Const $HTBORDER = 18
Global Const $HTREDUCE = $HTMINBUTTON
Global Const $HTZOOM = $HTMAXBUTTON
Global Const $HTSIZEFIRST = $HTLEFT
Global Const $HTSIZELAST = $HTBOTTOMRIGHT
Global Const $HTOBJECT = 19
Global Const $HTCLOSE = 20
Global Const $HTHELP = 21
Global Const $COLOR_SCROLLBAR = 0
Global Const $COLOR_BACKGROUND = 1
Global Const $COLOR_ACTIVECAPTION = 2
Global Const $COLOR_INACTIVECAPTION = 3
Global Const $COLOR_MENU = 4
Global Const $COLOR_WINDOW = 5
Global Const $COLOR_WINDOWFRAME = 6
Global Const $COLOR_MENUTEXT = 7
Global Const $COLOR_WINDOWTEXT = 8
Global Const $COLOR_CAPTIONTEXT = 9
Global Const $COLOR_ACTIVEBORDER = 10
Global Const $COLOR_INACTIVEBORDER = 11
Global Const $COLOR_APPWORKSPACE = 12
Global Const $COLOR_HIGHLIGHT = 13
Global Const $COLOR_HIGHLIGHTTEXT = 14
Global Const $COLOR_BTNFACE = 15
Global Const $COLOR_BTNSHADOW = 16
Global Const $COLOR_GRAYTEXT = 17
Global Const $COLOR_BTNTEXT = 18
Global Const $COLOR_INACTIVECAPTIONTEXT = 19
Global Const $COLOR_BTNHIGHLIGHT = 20
Global Const $COLOR_3DDKSHADOW = 21
Global Const $COLOR_3DLIGHT = 22
Global Const $COLOR_INFOTEXT = 23
Global Const $COLOR_INFOBK = 24
Global Const $COLOR_HOTLIGHT = 26
Global Const $COLOR_GRADIENTACTIVECAPTION = 27
Global Const $COLOR_GRADIENTINACTIVECAPTION = 28
Global Const $COLOR_MENUHILIGHT = 29
Global Const $COLOR_MENUBAR = 30
Global Const $COLOR_DESKTOP = 1
Global Const $COLOR_3DFACE = 15
Global Const $COLOR_3DSHADOW = 16
Global Const $COLOR_3DHIGHLIGHT = 20
Global Const $COLOR_3DHILIGHT = 20
Global Const $COLOR_BTNHILIGHT = 20
Global Const $HINST_COMMCTRL = -1
Global Const $IDB_STD_SMALL_COLOR = 0
Global Const $IDB_STD_LARGE_COLOR = 1
Global Const $IDB_VIEW_SMALL_COLOR = 4
Global Const $IDB_VIEW_LARGE_COLOR = 5
Global Const $IDB_HIST_SMALL_COLOR = 8
Global Const $IDB_HIST_LARGE_COLOR = 9
Global Const $STARTF_FORCEOFFFEEDBACK = 0x80
Global Const $STARTF_FORCEONFEEDBACK = 0x40
Global Const $STARTF_PREVENTPINNING = 0x00002000
Global Const $STARTF_RUNFULLSCREEN = 0x20
Global Const $STARTF_TITLEISAPPID = 0x00001000
Global Const $STARTF_TITLEISLINKNAME = 0x00000800
Global Const $STARTF_USECOUNTCHARS = 0x8
Global Const $STARTF_USEFILLATTRIBUTE = 0x10
Global Const $STARTF_USEHOTKEY = 0x200
Global Const $STARTF_USEPOSITION = 0x4
Global Const $STARTF_USESHOWWINDOW = 0x1
Global Const $STARTF_USESIZE = 0x2
Global Const $STARTF_USESTDHANDLES = 0x100
Global Const $CDDS_PREPAINT = 0x00000001
Global Const $CDDS_POSTPAINT = 0x00000002
Global Const $CDDS_PREERASE = 0x00000003
Global Const $CDDS_POSTERASE = 0x00000004
Global Const $CDDS_ITEM = 0x00010000
Global Const $CDDS_ITEMPREPAINT = 0x00010001
Global Const $CDDS_ITEMPOSTPAINT = 0x00010002
Global Const $CDDS_ITEMPREERASE = 0x00010003
Global Const $CDDS_ITEMPOSTERASE = 0x00010004
Global Const $CDDS_SUBITEM = 0x00020000
Global Const $CDIS_SELECTED = 0x0001
Global Const $CDIS_GRAYED = 0x0002
Global Const $CDIS_DISABLED = 0x0004
Global Const $CDIS_CHECKED = 0x0008
Global Const $CDIS_FOCUS = 0x0010
Global Const $CDIS_DEFAULT = 0x0020
Global Const $CDIS_HOT = 0x0040
Global Const $CDIS_MARKED = 0x0080
Global Const $CDIS_INDETERMINATE = 0x0100
Global Const $CDIS_SHOWKEYBOARDCUES = 0x0200
Global Const $CDIS_NEARHOT = 0x0400
Global Const $CDIS_OTHERSIDEHOT = 0x0800
Global Const $CDIS_DROPHILITED = 0x1000
Global Const $CDRF_DODEFAULT = 0x00000000
Global Const $CDRF_NEWFONT = 0x00000002
Global Const $CDRF_SKIPDEFAULT = 0x00000004
Global Const $CDRF_NOTIFYPOSTPAINT = 0x00000010
Global Const $CDRF_NOTIFYITEMDRAW = 0x00000020
Global Const $CDRF_NOTIFYSUBITEMDRAW = 0x00000020
Global Const $CDRF_NOTIFYPOSTERASE = 0x00000040
Global Const $CDRF_DOERASE = 0x00000008
Global Const $CDRF_SKIPPOSTPAINT = 0x00000100
Global Const $GUI_SS_DEFAULT_GUI = BitOR($WS_MINIMIZEBOX, $WS_CAPTION, $WS_POPUP, $WS_SYSMENU)
Local $aUtilVer = "v1.0"
Global $aUtilVerNumber = 0
Opt("GUIOnEventMode", 1)
Global $aNumberOfContacts = 0
Global $aIniFile = @ScriptDir & "\" & $aUtilFileName & ".ini"
Global $aINIHeader = " --------------- " & StringUpper($aUtilName) & " --------------- "
Global $cSWButtonStartServer = "0xA5B89B"
Global $cSWButtonStopServer = "0xB89B9B"
Global $cSWButtonRestartUtil = "0xD3D17F"
Global $cButtonFadedBlue = "0xCBDDFF"
Global $cSWTitle = "0xDC143C"
Global $aServerUpdateLinkVer = "http://www.phoenix125.com/share/mailenable/latestver.txt"
Global $aServerUpdateLinkDL = "http://www.phoenix125.com/share/mailenable/GoogleContactsToMailEnable.zip"
Global $aUpdateAvailableTF = False
_IniRead()
_IniWrite()
_ShowGUI()
While True
Sleep(100)
WEnd
Func _SetGlobalVariablesGoogle($tMax)
Global $gFN[$tMax]
Global $gNameFirst[$tMax]
Global $gNameMiddle[$tMax]
Global $gNameLast[$tMax]
Global $gNamePre[$tMax]
Global $gNameSuffix[$tMax]
Global $gNickName[$tMax]
Global $gEmailMain[$tMax]
Global $gPhoneWork[$tMax]
Global $gPhoneHome[$tMax]
Global $gPhoneOther[$tMax]
Global $gPhoneMobile[$tMax]
Global $gPhoneWorkFax[$tMax]
Global $gAddressHomePOBox[$tMax]
Global $gAddressHomeStreet1[$tMax]
Global $gAddressHomeStreet2[$tMax]
Global $gAddressHomeCity[$tMax]
Global $gAddressHomeState[$tMax]
Global $gAddressHomeZIP[$tMax]
Global $gAddressHomeCountry[$tMax]
Global $gAddressWorkPOBox[$tMax]
Global $gAddressWorkStreet1[$tMax]
Global $gAddressWorkStreet2[$tMax]
Global $gAddressWorkCity[$tMax]
Global $gAddressWorkState[$tMax]
Global $gAddressWorkZIP[$tMax]
Global $gAddressWorkCountry[$tMax]
Global $gAddressOtherPOBox[$tMax]
Global $gAddressOtherStreet1[$tMax]
Global $gAddressOtherStreet2[$tMax]
Global $gAddressOtherCity[$tMax]
Global $gAddressOtherState[$tMax]
Global $gAddressOtherZIP[$tMax]
Global $gAddressOtherCountry[$tMax]
Global $gBirthday[$tMax]
Global $gCompany[$tMax]
Global $gDepartment[$tMax]
Global $gURLWork[$tMax]
Global $gJobTitle[$tMax]
Global $gHomePage[$tMax]
Global $gRelatedLabel[$tMax]
Global $gRelatedName[$tMax]
Global $gNote[$tMax]
Global $gCustom[$tMax]
Global $gURLBlog[$tMax]
Global $gURLProfile[$tMax]
Global $gCategories[$tMax]
Global $gPhotoURL[$tMax]
Global $gEmailWork[$tMax]
Global $gPhonePager[$tMax]
Global $gEmailOther[$tMax]
Global $gChatYAHOO[$tMax]
Global $gChatAIM[$tMax]
Global $gChatGTALK[$tMax]
Global $gChatICQ[$tMax]
Global $gChatQQ[$tMax]
Global $gChatSkype[$tMax]
Global $gChatMSN[$tMax]
Global $gChatJabber[$tMax]
Global $gCategNames[0]
Global $gCategContacts[0]
EndFunc
Func _SetGlobalVariablesMailEnable($tMax)
Global $mFN[$tMax]
Global $mN[$tMax]
Global $mNICKNAME[$tMax]
Global $mEMAIL_PREF_INTERNET[$tMax]
Global $mTEL_HOME_VOICE[$tMax]
Global $mTEL_CELL_VOICE[$tMax]
Global $mTEL_PAGER_VOICE[$tMax]
Global $mNOTE[$tMax]
Global $mTEL_WORK_VOICE[$tMax]
Global $mTEL_WORK_FAX[$tMax]
Global $mADR_HOME[$tMax]
Global $mADR_WORK[$tMax]
Global $mADR_OTHER[$tMax]
Global $mBDAY[$tMax]
Global $mORG[$tMax]
Global $mURL_WORK[$tMax]
Global $mTITLE[$tMax]
Global $mFileName[$tMax]
Global $mIndex[$tMax]
Global $mUID[$tMax]
Global $mEMAIL_WORK_INTERNET[$tMax]
Global $mEMAIL_OTHER_INTERNET[$tMax]
Global $mTYPE_MSMESSENGER[$tMax]
EndFunc
Func _IniRead()
Global $aExportFolder = _RemoveTrailingSlash(IniRead($aIniFile, $aINIHeader, "ExportFolder", "C:\Program Files (x86)\Mail Enable\Postoffices"))
Global $gGoogleSourceFile = IniRead($aIniFile, $aINIHeader, "GoogleSourceFile", @ScriptDir & "\contacts.vcf")
Global $aAddGroupsToNotesYN = IniRead($aIniFile, $aINIHeader, "AddGroupsToNotesYN", "no")
Global $aDownloadPhotosYN = IniRead($aIniFile, $aINIHeader, "DownloadPhotosYN", "yes")
Global $aCopyPhonesNosToMobileYN = IniRead($aIniFile, $aINIHeader, "CopyPhonesNosToMobileYN", "no")
Global $aCopyWorkorOtherEmailsToPrefYN = IniRead($aIniFile, $aINIHeader, "CopyWorkorOtherEmailsToPrefYN", "yes")
Global $aReformatPhoneNumbersYN = IniRead($aIniFile, $aINIHeader, "ReformatPhoneNumbersYN", "yes")
Global $aWorkAliases = IniRead($aIniFile, $aINIHeader, "WorkAliases", "Work,Office,Corporate")
Global $aMobileAliases = IniRead($aIniFile, $aINIHeader, "MobileAliases", "Mobile,Cell")
Global $aHomeAliases = IniRead($aIniFile, $aINIHeader, "HomeAliases", "Home")
Global $aCreateGroupsYN = IniRead($aIniFile, $aINIHeader, "CreateGroupsYN", "yes")
Global $aOpenFolderWhenDoneYN = IniRead($aIniFile, $aINIHeader, "OpenFolderWhenDoneYN", "yes")
Global $aCheckForUpdatesYN = IniRead($aIniFile, $aINIHeader, "CheckForUpdatesYN", "yes")
Global $aIndex_XML = $aExportFolder & "\_index.xml"
Global $aGroups_XML = $aExportFolder & "\_GROUPS.XML"
Global $aUtilLastVerNumber = IniRead($aIniFile, $aINIHeader, "UtilLastVerNumber", $aUtilVerNumber)
EndFunc
Func _IniWrite()
IniWrite($aIniFile, $aINIHeader, "ExportFolder", $aExportFolder)
IniWrite($aIniFile, $aINIHeader, "GoogleSourceFile", $gGoogleSourceFile)
IniWrite($aIniFile, $aINIHeader, "AddGroupsToNotesYN", $aAddGroupsToNotesYN)
IniWrite($aIniFile, $aINIHeader, "DownloadPhotosYN", $aDownloadPhotosYN)
IniWrite($aIniFile, $aINIHeader, "CopyPhonesNosToMobileYN", $aCopyPhonesNosToMobileYN)
IniWrite($aIniFile, $aINIHeader, "CopyWorkorOtherEmailsToPrefYN", $aCopyWorkorOtherEmailsToPrefYN)
IniWrite($aIniFile, $aINIHeader, "ReformatPhoneNumbersYN", $aReformatPhoneNumbersYN)
IniWrite($aIniFile, $aINIHeader, "WorkAliases", $aWorkAliases)
IniWrite($aIniFile, $aINIHeader, "MobileAliases", $aMobileAliases)
IniWrite($aIniFile, $aINIHeader, "HomeAliases", $aHomeAliases)
IniWrite($aIniFile, $aINIHeader, "CreateGroupsYN", $aCreateGroupsYN)
IniWrite($aIniFile, $aINIHeader, "OpenFolderWhenDoneYN", $aOpenFolderWhenDoneYN)
IniWrite($aIniFile, $aINIHeader, "CheckForUpdatesYN", $aCheckForUpdatesYN)
IniWrite($aIniFile, $aINIHeader, "UtilLastVerNumber", $aUtilVerNumber)
EndFunc
Func _ReadGoogleVCAR()
Global $gFileReadPre
Global $gFileRead[0]
ControlSetText($aSplash, "", "Static1", "Reading Google VCF File." & @CRLF & @CRLF & "Timer:" & Round(_Timer_Diff($aTimerStart) / 1000) & "s")
_FileReadToArray($gGoogleSourceFile, $gFileReadPre)
$gUbound = UBound($gFileReadPre) - 1
$tCnt = 0
For $t0 = 0 To $gUbound
$tCnt += 1
If $tCnt = 100 Then
ControlSetText($aSplash, "", "Static1", "Combining split fields." & @CRLF & "Processing line " & $t0 & " of " & $gUbound & @CRLF & @CRLF & "Timer:" & Round(_Timer_Diff($aTimerStart) / 1000) & "s")
$tCnt = 0
EndIf
Local $tTxt = $gFileReadPre[$t0]
If $t0 <= $gUbound - 1 Then
If StringLeft($gFileReadPre[$t0 + 1], 1) = " " Then
Do
If $t0 <= $gUbound + 1 Then
$t0 += 1
$tTxt &= StringTrimLeft($gFileReadPre[$t0], 1)
EndIf
Until StringLeft($gFileReadPre[$t0 + 1], 1) <> " "
EndIf
EndIf
_ArrayAdd($gFileRead, $tTxt)
Next
$gUbound = UBound($gFileRead) - 1
For $t1 = 0 To ($gUbound)
If $gFileRead[$t1] = "BEGIN:VCARD" Then
$aNumberOfContacts += 1
ControlSetText($aSplash, "", "Static1", "Counting number of contacts." & @CRLF & "Found " & $aNumberOfContacts & " contacts" & @CRLF & @CRLF & "Timer:" & Round(_Timer_Diff($aTimerStart) / 1000) & "s")
EndIf
Next
_SetGlobalVariablesGoogle($aNumberOfContacts)
Local $tNo = -1
Local $tCnt = 0
For $t1 = 0 To $gUbound
If $gFileRead[$t1] = "BEGIN:VCARD" Then
$tCnt += 1
If $tCnt = 10 Then
ControlSetText($aSplash, "", "Static1", "Reading Gmail VCF file." & @CRLF & "Processing contact " & $t1 & " of " & $aNumberOfContacts & @CRLF & @CRLF & "Timer:" & Round(_Timer_Diff($aTimerStart) / 1000) & "s")
$tCnt = 0
EndIf
Local $tNextTemp = ""
Local $tRelatedTemp = ""
Local $tExtendedPre = ""
Local $tExtendedPost = ""
$tNo += 1
If $tNo >= $aNumberOfContacts Then
MsgBox(0, "Kim", "Error. More More contacts than calculated total." & @CRLF & "Count:" & $tNo & " | Total:" & $aNumberOfContacts, 30)
ExitLoop
EndIf
For $t2 = $t1 To $gUbound
If $gFileRead[$t2] = "END:VCARD" Then ExitLoop
Local $tExtendedPre = $gFileRead[$t2]
Local $tPre = "!NoColon!"
Local $tPost = ""
For $t3 = 1 To StringLen($gFileRead[$t2])
If StringMid($gFileRead[$t2], $t3, 1) = ":" Then
$tPre = StringLeft($gFileRead[$t2], ($t3 - 1))
$tPost = StringRight($gFileRead[$t2], (StringLen($gFileRead[$t2]) - $t3))
ExitLoop
EndIf
Next
Local $tNoNumber = StringRegExpReplace($tPre, "\d", "")
If StringLeft($tNoNumber, 5) = "item." Then $tNoNumber = StringTrimLeft($tNoNumber, 5)
Local $tNumber = StringRegExpReplace($tPre, "[A-Za-z.,=;-]", "")
If $tPre = "FN" Then
$gFN[$tNo] = $tPost
ElseIf $tPre = "N" Then
Local $tSplit = StringSplit($tPost, ";")
If UBound($tSplit) > 0 Then $gNameLast[$tNo] = $tSplit[1]
If UBound($tSplit) > 1 Then $gNameFirst[$tNo] = $tSplit[2]
If UBound($tSplit) > 2 Then $gNameMiddle[$tNo] = $tSplit[3]
If UBound($tSplit) > 3 Then $gNamePre[$tNo] = $tSplit[4]
If UBound($tSplit) > 4 Then $gNameSuffix[$tNo] = $tSplit[5]
ElseIf $tPre = "NICKNAME" Then
$gNickName[$tNo] = $tPost
ElseIf $tPre = "TEL;TYPE=WORK" Then
If $gPhoneWork[$tNo] = "" Then
$gPhoneWork[$tNo] = $tPost
Else
$gCustom[$tNo] &= "Work:" & $tNextTemp & @CRLF
EndIf
ElseIf $tPre = "TEL;TYPE=HOME" Then
If $gPhoneHome[$tNo] = "" Then
$gPhoneHome[$tNo] = $tPost
Else
$gCustom[$tNo] &= "Home:" & $tNextTemp & @CRLF
EndIf
ElseIf $tPre = "TEL;TYPE=PAGER" Then
If $gPhonePager[$tNo] = "" Then
$gPhonePager[$tNo] = $tPost
Else
$gCustom[$tNo] &= "Pager:" & $tNextTemp & @CRLF
EndIf
ElseIf $tPre = "TEL" Then
$gPhoneOther[$tNo] &= "Phone Other:" & $tPost & @CRLF
ElseIf $tPre = "TEL;TYPE=CELL" Then
If $gPhoneMobile[$tNo] = "" Then
$gPhoneMobile[$tNo] = $tPost
Else
$gCustom[$tNo] &= "Cell:" & $tNextTemp & @CRLF
EndIf
EndIf
If $tRelatedTemp <> "" Then
$gRelatedName[$tNo] = $tRelatedTemp
Local $tTxt = StringReplace($tPost, "_$!<", "")
$tTxt = StringReplace($tTxt, ">!$_", "")
If $tNoNumber = "X-ABLabel" Then $gRelatedLabel[$tNo] = $tTxt
$tRelatedTemp = ""
$tNoNumber = ""
EndIf
If $tNoNumber = "EMAIL;TYPE=INTERNET;TYPE=HOME" Then
If $gEmailMain[$tNo] = "" Then
$gEmailMain[$tNo] = $tPost
Else
$gCustom[$tNo] &= "Home Email:" & $tPost & @CRLF
EndIf
ElseIf $tNoNumber = "EMAIL;TYPE=INTERNET;TYPE=WORK" Then
If $gEmailWork[$tNo] = "" Then
$gEmailWork[$tNo] = $tPost
Else
$gCustom[$tNo] &= "Work Email:" & $tPost & @CRLF
EndIf
ElseIf $tNoNumber = "EMAIL;TYPE=INTERNET" Then
If $gEmailMain[$tNo] = "" Then
$gEmailMain[$tNo] = $tPost
Else
If $t1 = $gUbound Then
$gCustom[$tNo] &= "Email:" & $tPost & @CRLF
Else
Local $tPre1 = ""
Local $tPost1 = ""
For $t3 = 1 To StringLen($gFileRead[$t2 + 1])
If StringMid($gFileRead[$t2 + 1], $t3, 1) = ":" Then
$tPre1 = StringLeft($gFileRead[$t2 + 1], ($t3 - 1))
$tPost1 = StringRight($gFileRead[$t2 + 1], (StringLen($gFileRead[$t2 + 1]) - $t3))
ExitLoop
EndIf
Next
Local $tNoNumber1 = StringRegExpReplace($tPre1, "\d", "")
If StringLeft($tNoNumber1, 5) = "item." Then $tNoNumber1 = StringTrimLeft($tNoNumber1, 5)
Local $tLabel1 = _ConvertCustomLabel($tPost1)
If $tNoNumber1 = "X-ABLabel" Then
If $gEmailOther[$tNo] = "" Then
$gEmailOther[$tNo] = $tPost
Else
$gCustom[$tNo] &= $tLabel1 & ":" & $tPost & @CRLF
EndIf
Else
If $gEmailOther[$tNo] = "" Then
$gEmailOther[$tNo] = $tPost
Else
$gCustom[$tNo] &= "Email:" & $tPost & @CRLF
EndIf
EndIf
EndIf
$tLabel = ""
EndIf
ElseIf $tNoNumber = "TITLE" Then
$gJobTitle[$tNo] = $tPost
ElseIf $tNoNumber = "TEL" Then
$tNextTemp = $tPost
ElseIf $tNoNumber = "URL" Then
$tNextTemp = $tPost
ElseIf $tNoNumber = "X-ABRELATEDNAMES" Then
$tRelatedTemp = $tPost
ElseIf $tNoNumber = "X-ABDATE" Then
$tNextTemp = $tPost
ElseIf $tNoNumber = "PHOTO" Then
$gPhotoURL[$tNo] = StringReplace($tPost, "\:", ":")
ElseIf $tNoNumber = "ORG" Then
Local $tSplit = StringSplit($tPost, ";")
If UBound($tSplit) > 1 Then $gCompany[$tNo] = $tSplit[1]
If UBound($tSplit) > 2 Then $gDepartment[$tNo] = $tSplit[2]
ElseIf $tNoNumber = "X-ABLabel" Then
If StringLeft($tPost, 4) = "_$!<" Then
Local $tString = StringTrimLeft($tPost, 4)
$tPost = StringTrimRight($tString, 4)
EndIf
Local $tLabel = _ConvertCustomLabel($tPost)
If $tLabel = "Home Fax" Then
If $gPhoneOther[$tNo] = "" Then
$gPhoneOther[$tNo] &= "Home Fax:" & $tNextTemp & @CRLF
Else
$gCustom[$tNo] &= "Home Fax:" & $tNextTemp & @CRLF
EndIf
$tLabel = ""
ElseIf $tLabel = "Work Fax" Then
If $gPhoneWorkFax[$tNo] = "" Then
$gPhoneWorkFax[$tNo] = $tNextTemp
Else
$gCustom[$tNo] &= "Work Fax:" & $tNextTemp & @CRLF
EndIf
$tLabel = ""
ElseIf $tLabel = "Home Page" Then
If $gHomePage[$tNo] = "" Then
$gHomePage[$tNo] = $tNextTemp
Else
$gCustom[$tNo] &= "Home Page:" & $tNextTemp & @CRLF
EndIf
$tLabel = ""
ElseIf $tLabel = "Blog" Then
If $gURLBlog[$tNo] = "" Then
$gURLBlog[$tNo] = $tNextTemp
Else
$gCustom[$tNo] &= "Blog:" & $tNextTemp & @CRLF
EndIf
$tLabel = ""
ElseIf $tLabel = "Profile" Then
$gURLProfile[$tNo] = $tNextTemp
$tLabel = ""
Else
If $tLabel <> "" And $tNextTemp <> "" Then
Local $tFound = False
Local $tHome = StringSplit($aHomeAliases, ",")
If UBound($tHome) > 1 Then
For $t = 1 To (UBound($tHome) - 1)
If $tHome[$t] = $tLabel Then
If $gPhoneHome[$tNo] = "" Then
$gPhoneHome[$tNo] = $tNextTemp
$tFound = True
ExitLoop
EndIf
EndIf
Next
EndIf
Local $tWork = StringSplit($aWorkAliases, ",")
If UBound($tWork) > 1 And $tFound = False Then
For $t = 1 To (UBound($tWork) - 1)
If $tWork[$t] = $tLabel Then
If $gPhoneWork[$tNo] = "" Then
$gPhoneWork[$tNo] = $tNextTemp
$tFound = True
ExitLoop
EndIf
EndIf
Next
EndIf
Local $tMobile = StringSplit($aMobileAliases, ",")
If UBound($tMobile) > 1 And $tFound = False Then
For $t = 1 To (UBound($tMobile) - 1)
If $tMobile[$t] = $tLabel Then
If $gPhoneMobile[$tNo] = "" Then
$gPhoneMobile[$tNo] = $tNextTemp
$tFound = True
ExitLoop
EndIf
EndIf
Next
EndIf
If $tFound = False Then $gCustom[$tNo] &= $tLabel & ":" & $tNextTemp & @CRLF
EndIf
EndIf
$tNextTemp = ""
EndIf
If $tPre = "ADR;TYPE=HOME" Then
Local $tSplit = StringSplit($tPost, ";")
If UBound($tSplit) > 1 Then $gAddressHomePOBox[$tNo] = $tSplit[1]
If UBound($tSplit) > 2 Then $gAddressHomeStreet2[$tNo] = $tSplit[2]
If UBound($tSplit) > 3 Then $gAddressHomeStreet1[$tNo] = $tSplit[3]
If UBound($tSplit) > 4 Then $gAddressHomeCity[$tNo] = $tSplit[4]
If UBound($tSplit) > 5 Then $gAddressHomeState[$tNo] = $tSplit[5]
If UBound($tSplit) > 6 Then $gAddressHomeZIP[$tNo] = $tSplit[6]
If UBound($tSplit) > 7 Then $gAddressHomeCountry[$tNo] = $tSplit[7]
ElseIf $tPre = "ADR;TYPE=WORK" Then
Local $tSplit = StringSplit($tPost, ";")
If UBound($tSplit) > 1 Then $gAddressWorkPOBox[$tNo] = $tSplit[1]
If UBound($tSplit) > 2 Then $gAddressWorkStreet2[$tNo] = $tSplit[2]
If UBound($tSplit) > 3 Then $gAddressWorkStreet1[$tNo] = $tSplit[3]
If UBound($tSplit) > 4 Then $gAddressWorkCity[$tNo] = $tSplit[4]
If UBound($tSplit) > 5 Then $gAddressWorkState[$tNo] = $tSplit[5]
If UBound($tSplit) > 6 Then $gAddressWorkZIP[$tNo] = $tSplit[6]
If UBound($tSplit) > 7 Then $gAddressWorkCountry[$tNo] = $tSplit[7]
ElseIf $tPre = "ADR" Then
Local $tSplit = StringSplit($tPost, ";")
If UBound($tSplit) > 1 Then $gAddressOtherPOBox[$tNo] = $tSplit[1]
If UBound($tSplit) > 2 Then $gAddressOtherStreet2[$tNo] = $tSplit[2]
If UBound($tSplit) > 3 Then $gAddressOtherStreet1[$tNo] = $tSplit[3]
If UBound($tSplit) > 4 Then $gAddressOtherCity[$tNo] = $tSplit[4]
If UBound($tSplit) > 5 Then $gAddressOtherState[$tNo] = $tSplit[5]
If UBound($tSplit) > 6 Then $gAddressOtherZIP[$tNo] = $tSplit[6]
If UBound($tSplit) > 7 Then $gAddressOtherCountry[$tNo] = $tSplit[7]
ElseIf $tPre = "URL;TYPE=WORK" Then
$gURLWork[$tNo] = $tPost
ElseIf $tPre = "BDAY" Then
$gBirthday[$tNo] = $tPost
ElseIf $tPre = "NOTE" Then
$gNote[$tNo] = $tPost
ElseIf $tPre = "YAHOO" Then
$gChatYAHOO[$tNo] = $tPost
ElseIf $tPre = "AIM" Then
$gChatAIM[$tNo] = $tPost
ElseIf $tPre = "GTALK" Then
$gChatGTALK[$tNo] = $tPost
ElseIf $tPre = "ICQ" Then
$gChatICQ[$tNo] = $tPost
ElseIf $tPre = "JABBER" Then
$gChatJabber[$tNo] = $tPost
ElseIf $tPre = "SKYPE" Then
$gChatSkype[$tNo] = $tPost
ElseIf $tPre = "QQ" Then
$gChatQQ[$tNo] = $tPost
ElseIf $tPre = "MSN" Then
$gChatMSN[$tNo] = $tPost
ElseIf $tPre = "CATEGORIES" Then
$gCategories[$tNo] = $tPost
EndIf
Next
$gPhoneOther[$tNo] = StringTrimRight($gPhoneOther[$tNo], 2)
EndIf
Next
EndFunc
Func _ConvertGoogleToMailEnable()
Local $tCnt = 0
For $tNo = 0 To ($aNumberOfContacts - 1)
$tCnt += 1
If $tCnt = 10 Then
ControlSetText($aSplash, "", "Static1", "Converting to Mail Enable files." & @CRLF & "Processing contact " & $tNo & " of " & $aNumberOfContacts & @CRLF & @CRLF & "Timer:" & Round(_Timer_Diff($aTimerStart) / 1000) & "s")
$tCnt = 0
EndIf
Local $tAdd2Home = ""
Local $tAddPOHome = ""
Local $tAdd2Other = ""
Local $tAddPOOther = ""
Local $tAdd2Work = ""
Local $tAddPOWork = ""
$mFN[$tNo] = _ReplaceChar($gFN[$tNo])
$mN[$tNo] = _ReplaceChar($gNameLast[$tNo] & ";" & $gNameFirst[$tNo] & ";" & $gNameMiddle[$tNo] & ";" & $gNameSuffix[$tNo] & ";" & $gNamePre[$tNo])
$mNICKNAME[$tNo] = _ReplaceChar($gNickName[$tNo])
$mEMAIL_PREF_INTERNET[$tNo] = $gEmailMain[$tNo]
$mEMAIL_WORK_INTERNET[$tNo] = $gEmailWork[$tNo]
$mEMAIL_OTHER_INTERNET[$tNo] = $gEmailOther[$tNo]
$mTEL_WORK_VOICE[$tNo] = _ReplaceChar($gPhoneWork[$tNo])
$mTEL_HOME_VOICE[$tNo] = _ReplaceChar($gPhoneHome[$tNo])
$mTEL_CELL_VOICE[$tNo] = _ReplaceChar($gPhoneMobile[$tNo])
$mTEL_PAGER_VOICE[$tNo] = _ReplaceChar($gPhonePager[$tNo])
$mNOTE[$tNo] = StringReplace($gNote[$tNo], "\n", @CRLF) & @CRLF & @CRLF
If $gChatGTALK[$tNo] <> "" Then
If $mTYPE_MSMESSENGER[$tNo] = "" Then
$mTYPE_MSMESSENGER[$tNo] = _ReplaceChar($gChatGTALK[$tNo])
Else
$mNOTE[$tNo] &= "Chat Google Talk:" & $gChatGTALK[$tNo] & @CRLF
EndIf
EndIf
If $gChatSkype[$tNo] <> "" Then
If $mTYPE_MSMESSENGER[$tNo] = "" Then
$mTYPE_MSMESSENGER[$tNo] = _ReplaceChar($gChatSkype[$tNo])
Else
$mNOTE[$tNo] &= "Chat Skype:" & $gChatSkype[$tNo] & @CRLF
EndIf
EndIf
If $gChatYAHOO[$tNo] <> "" Then
If $mTYPE_MSMESSENGER[$tNo] = "" Then
$mTYPE_MSMESSENGER[$tNo] = _ReplaceChar($gChatYAHOO[$tNo])
Else
$mNOTE[$tNo] &= "Chat Yahoo:" & $gChatYAHOO[$tNo] & @CRLF
EndIf
EndIf
If $gChatMSN[$tNo] <> "" Then
If $mTYPE_MSMESSENGER[$tNo] = "" Then
$mTYPE_MSMESSENGER[$tNo] = _ReplaceChar($gChatMSN[$tNo])
Else
$mNOTE[$tNo] &= "Chat MSN:" & $gChatMSN[$tNo] & @CRLF
EndIf
EndIf
If $gChatICQ[$tNo] <> "" Then
If $mTYPE_MSMESSENGER[$tNo] = "" Then
$mTYPE_MSMESSENGER[$tNo] = _ReplaceChar($gChatICQ[$tNo])
Else
$mNOTE[$tNo] &= "Chat ICQ:" & $gChatICQ[$tNo] & @CRLF
EndIf
EndIf
If $gChatAIM[$tNo] <> "" Then
If $mTYPE_MSMESSENGER[$tNo] = "" Then
$mTYPE_MSMESSENGER[$tNo] = _ReplaceChar($gChatAIM[$tNo])
Else
$mNOTE[$tNo] &= "Chat AIM:" & $gChatAIM[$tNo] & @CRLF
EndIf
EndIf
If $gChatQQ[$tNo] <> "" Then
If $mTYPE_MSMESSENGER[$tNo] = "" Then
$mTYPE_MSMESSENGER[$tNo] = _ReplaceChar($gChatQQ[$tNo])
Else
$mNOTE[$tNo] &= "Chat QQ:" & $gChatQQ[$tNo] & @CRLF
EndIf
EndIf
If $gChatJabber[$tNo] <> "" Then
If $mTYPE_MSMESSENGER[$tNo] = "" Then
$mTYPE_MSMESSENGER[$tNo] = _ReplaceChar($gChatJabber[$tNo])
Else
$mNOTE[$tNo] &= "Chat Jabber:" & $gChatJabber[$tNo] & @CRLF
EndIf
EndIf
If $gHomePage[$tNo] <> "" Then $mNOTE[$tNo] &= "Home Page:" & $gHomePage[$tNo] & @CRLF
If $gPhoneOther[$tNo] <> "" Then $mNOTE[$tNo] &= $gPhoneOther[$tNo] & @CRLF
If $gURLBlog[$tNo] <> "" Then $mNOTE[$tNo] &= "Blog:" & $gURLBlog[$tNo] & @CRLF
If $gURLProfile[$tNo] <> "" Then $mNOTE[$tNo] &= "Profile:" & $gURLProfile[$tNo] & @CRLF
If $gRelatedLabel[$tNo] <> "" Then $mNOTE[$tNo] &= $gRelatedLabel[$tNo] & ": " & $gRelatedName[$tNo] & @CRLF
If $gCategories[$tNo] <> "" Then
Local $tSplit = StringSplit($gCategories[$tNo], ",", 2)
For $t = 0 To (UBound($tSplit) - 1)
$tSplit[$t] = _ConvertCustomLabel($tSplit[$t])
$tSplit[$t] = _RemoveSpecialChars($tSplit[$t])
Next
$gCategories[$tNo] = _ArrayToString($tSplit, ",")
If $aAddGroupsToNotesYN = "yes" Then $mNOTE[$tNo] &= "Categories:" & $gCategories[$tNo] & @CRLF
EndIf
$mTEL_WORK_FAX[$tNo] = $gPhoneWorkFax[$tNo]
If $gAddressHomeStreet2[$tNo] <> "" Then _ReplaceChar($tAdd2Home = ", " & $gAddressHomeStreet2[$tNo])
If $gAddressHomePOBox[$tNo] <> "" Then _ReplaceChar($tAddPOHome = $gAddressHomePOBox[$tNo] & ", ")
$mADR_HOME[$tNo] = _ReplaceChar(";;" & $tAddPOHome & $gAddressHomeStreet1[$tNo] & $tAdd2Home & ";" & $gAddressHomeCity[$tNo] & ";" & $gAddressHomeState[$tNo] & ";" & $gAddressHomeZIP[$tNo] & ";" & $gAddressHomeCountry[$tNo])
If $gAddressWorkStreet2[$tNo] <> "" Then _ReplaceChar($tAdd2Work = ", " & $gAddressWorkStreet2[$tNo])
If $gAddressWorkPOBox[$tNo] <> "" Then $tAddPOWork = $gAddressWorkPOBox[$tNo] & ", "
$mADR_WORK[$tNo] = _ReplaceChar(";;" & $tAddPOWork & $gAddressWorkStreet1[$tNo] & $tAdd2Work & ";" & $gAddressWorkCity[$tNo] & ";" & $gAddressWorkState[$tNo] & ";" & $gAddressWorkZIP[$tNo] & ";" & $gAddressWorkCountry[$tNo])
If $gAddressOtherStreet2[$tNo] <> "" Then $tAdd2Other = ", " & $gAddressOtherStreet2[$tNo]
If $gAddressOtherPOBox[$tNo] <> "" Then $tAddPOOther = $gAddressOtherPOBox[$tNo] & ", "
$mADR_OTHER[$tNo] = _ReplaceChar(";;" & $tAddPOOther & $gAddressOtherStreet1[$tNo] & $tAdd2Other & ";" & $gAddressOtherCity[$tNo] & ";" & $gAddressOtherState[$tNo] & ";" & $gAddressOtherZIP[$tNo] & ";" & $gAddressOtherCountry[$tNo])
If $gBirthday[$tNo] <> "" Then $mBDAY[$tNo] = StringLeft($gBirthday[$tNo], 4) & "-" & StringMid($gBirthday[$tNo], 5, 2) & "-" & StringMid($gBirthday[$tNo], 7, 2) & "T00:00:00"
$mORG[$tNo] = _ReplaceChar($gCompany[$tNo] & ";" & $gDepartment[$tNo])
$mURL_WORK[$tNo] = StringReplace($gURLWork[$tNo], "\:", ":")
$mTITLE[$tNo] = _ReplaceChar($gJobTitle[$tNo])
$mNOTE[$tNo] &= $gCustom[$tNo]
$mNOTE[$tNo] = StringReplace($mNOTE[$tNo], "http\:", "http:")
$mNOTE[$tNo] = _ReplaceChar($mNOTE[$tNo])
$mFileName[$tNo] = _GenerateFileName() & ".VCF"
If StringLen(_RemoveSpecialChars($mADR_HOME[$tNo])) < 5 Then $mADR_HOME[$tNo] = ""
If StringLen(_RemoveSpecialChars($mADR_OTHER[$tNo])) < 5 Then $mADR_OTHER[$tNo] = ""
If StringLen(_RemoveSpecialChars($mADR_WORK[$tNo])) < 5 Then $mADR_WORK[$tNo] = ""
Next
EndFunc
Func _GenerateFileName()
Local $tFileName = ""
For $t = 1 To 32
Local $tRandom = Random(0, 16)
If $tRandom < 10 Then
$tFileName &= Chr(48 + $tRandom)
Else
$tFileName &= Chr(55 + $tRandom)
EndIf
Next
Return $tFileName
EndFunc
Func _ReplaceChar($tTxt)
Local $tReturn = StringReplace($tTxt, "&", "&amp;")
Return $tReturn
EndFunc
Func _ConvertPhoneNumbers($tNum)
If $mTEL_CELL_VOICE[$tNum] <> "" Then $mTEL_CELL_VOICE[$tNum] = _FormatPhoneNumber($mTEL_CELL_VOICE[$tNum])
If $mTEL_HOME_VOICE[$tNum] <> "" Then $mTEL_HOME_VOICE[$tNum] = _FormatPhoneNumber($mTEL_HOME_VOICE[$tNum])
If $mTEL_PAGER_VOICE[$tNum] <> "" Then $mTEL_PAGER_VOICE[$tNum] = _FormatPhoneNumber($mTEL_PAGER_VOICE[$tNum])
If $mTEL_WORK_FAX[$tNum] <> "" Then $mTEL_WORK_FAX[$tNum] = _FormatPhoneNumber($mTEL_WORK_FAX[$tNum])
If $mTEL_WORK_VOICE[$tNum] <> "" Then $mTEL_WORK_VOICE[$tNum] = _FormatPhoneNumber($mTEL_WORK_VOICE[$tNum])
EndFunc
Func _WriteMailEnable()
For $tNo = 0 To ($aNumberOfContacts - 1)
Local $tArray[0]
ControlSetText($aSplash, "", "Static1", "Writing Mail Enable files." & @CRLF & "Processing contact " & $tNo & " of " & $aNumberOfContacts & @CRLF & @CRLF & "Timer:" & Round(_Timer_Diff($aTimerStart) / 1000) & "s")
_ArrayAdd($tArray, "BEGIN:VCARD")
_ArrayAdd($tArray, "VERSION:2.1")
If StringReplace($mADR_HOME[$tNo], ";", "") <> "" Then
_ArrayAdd($tArray, "ADR;HOME:" & $mADR_HOME[$tNo])
Else
_ArrayAdd($tArray, "ADR;HOME:;;;;;;")
EndIf
If StringReplace($mADR_OTHER[$tNo], ";", "") <> "" Then
_ArrayAdd($tArray, "ADR;OTHER:" & $mADR_OTHER[$tNo])
Else
_ArrayAdd($tArray, "ADR;OTHER:;;;;;;")
EndIf
If StringReplace($mADR_WORK[$tNo], ";", "") <> "" Then
_ArrayAdd($tArray, "ADR;WORK:" & $mADR_WORK[$tNo])
Else
_ArrayAdd($tArray, "ADR;WORK:;;;;;;")
EndIf
If $mBDAY[$tNo] <> "" Then _ArrayAdd($tArray, "BDAY:" & $mBDAY[$tNo])
If $mEMAIL_OTHER_INTERNET[$tNo] <> "" Then _ArrayAdd($tArray, "EMAIL;OTHER;INTERNET:" & $mEMAIL_OTHER_INTERNET[$tNo])
If $mEMAIL_PREF_INTERNET[$tNo] <> "" Then _ArrayAdd($tArray, "EMAIL;PREF;INTERNET:" & $mEMAIL_PREF_INTERNET[$tNo])
If $mEMAIL_WORK_INTERNET[$tNo] <> "" Then _ArrayAdd($tArray, "EMAIL;WORK;INTERNET:" & $mEMAIL_WORK_INTERNET[$tNo])
If $mTYPE_MSMESSENGER[$tNo] <> "" Then _ArrayAdd($tArray, "IM;TYPE-MSMESSENGER:" & $mTYPE_MSMESSENGER[$tNo])
If $mFN[$tNo] <> "" Then _ArrayAdd($tArray, "FN:" & $mFN[$tNo])
If $mN[$tNo] <> "" Then
_ArrayAdd($tArray, "N:" & $mN[$tNo])
Else
_ArrayAdd($tArray, "N:;;;;")
EndIf
If $mNICKNAME[$tNo] <> "" Then _ArrayAdd($tArray, "NICKNAME:" & $mNICKNAME[$tNo])
If $mORG[$tNo] <> "" Then
_ArrayAdd($tArray, "ORG:" & $mORG[$tNo])
Else
_ArrayAdd($tArray, "ORG:;")
EndIf
If $aReformatPhoneNumbersYN = "yes" Then _ConvertPhoneNumbers($tNo)
If $mTEL_CELL_VOICE[$tNo] <> "" Then _ArrayAdd($tArray, "TEL;CELL;VOICE:" & $mTEL_CELL_VOICE[$tNo])
If $mTEL_HOME_VOICE[$tNo] <> "" Then _ArrayAdd($tArray, "TEL;HOME;VOICE:" & $mTEL_HOME_VOICE[$tNo])
If $mTEL_PAGER_VOICE[$tNo] <> "" Then _ArrayAdd($tArray, "TEL;PAGER;VOICE:" & $mTEL_PAGER_VOICE[$tNo])
If $mTEL_WORK_FAX[$tNo] <> "" Then _ArrayAdd($tArray, "TEL;WORK;FAX:" & $mTEL_WORK_FAX[$tNo])
If $mTEL_WORK_VOICE[$tNo] <> "" Then _ArrayAdd($tArray, "TEL;WORK;VOICE:" & $mTEL_WORK_VOICE[$tNo])
If $mTITLE[$tNo] <> "" Then _ArrayAdd($tArray, "TITLE:" & $mTITLE[$tNo])
If $mURL_WORK[$tNo] <> "" Then _ArrayAdd($tArray, "URL;WORK:" & $mURL_WORK[$tNo])
Local $tNote = StringReplace($mNOTE[$tNo], @CRLF, "")
$tNote = _RemoveExtraSpaces($tNote)
If $tNote <> "" Then
$mNOTE[$tNo] = StringReplace($mNOTE[$tNo], @CRLF, "=0A")
$mNOTE[$tNo] = StringReplace($mNOTE[$tNo], "\,", ",")
$mNOTE[$tNo] = StringReplace($mNOTE[$tNo], "\:", ":")
_ArrayAdd($tArray, "NOTE;ENCODING=QUOTED-PRINTABLE;CHARSET=UTF-8:" & $mNOTE[$tNo])
EndIf
If $aDownloadPhotosYN = "yes" Then
If $gPhotoURL[$tNo] <> "" Then
Local $tPhoto = StringRegExpReplace(_ConvertToBase64(_INetGetSource($gPhotoURL[$tNo])), @LF, "")
If StringLen($tPhoto) > 40 Then _ArrayAdd($tArray, "PHOTO;TYPE=JPEG;ENCODING=BASE64:")
Local $tStop = False
Local $tTxt = " "
Do
$tTxt = StringLeft($tPhoto, 75)
_ArrayAdd($tArray, " " & $tTxt)
If StringLen($tTxt) < 75 Then
$tStop = True
Else
$tPhoto = StringTrimLeft($tPhoto, 75)
EndIf
Until $tStop
_ArrayAdd($tArray, "")
EndIf
EndIf
_ArrayAdd($tArray, "REV:20210201T024607")
_ArrayAdd($tArray, "END:VCARD")
_FileWriteFromArray($aExportFolder & "\" & $mFileName[$tNo], $tArray)
Next
EndFunc
Func _UpdateIndex()
_GetLastUID()
For $tNo = 0 To ($aNumberOfContacts - 1)
Local $tELEMNT_ID = ""
Local $tADR_HOME = ""
Local $tADR_OTHER = ""
Local $tADR_WORK = ""
Local $tAGENT = ""
Local $tBDAY = ""
Local $tATTACHMENTS = ""
Local $tEMAIL_OTHER_INTERNET = ""
Local $tEMAIL_PREF_INTERNET = ""
Local $tEMAIL_WORK_INTERNET = ""
Local $tFN = ""
Local $tIMPORTANCE = ""
Local $tIM_TYPE_MSMESSENGER = ""
Local $tN = ""
Local $tNICKNAME = ""
Local $tORG = ""
Local $tREAD = ""
Local $tSIZE = ""
Local $tSTATE = ""
Local $tTEL_CELL_VOICE = ""
Local $tTEL_HOME_VOICE = ""
Local $tTEL_PAGER_VOICE = ""
Local $tTEL_WORK_VOICE = ""
Local $tTITLE = ""
Local $tELEMENT_END = ""
$aLastUIDDec += 1
$mUID[$tNo] = $aLastUIDDec
$tELEMNT_ID = '<ELEMENT ID="' & $mFileName[$tNo] & '" UID="' & $mUID[$tNo] & '">'
If $mADR_HOME[$tNo] <> "" Then
$tADR_HOME = '<ADR_HOME>' & $mADR_HOME[$tNo] & '</ADR_HOME>'
Else
$tADR_HOME = '<ADR_HOME>;;;;;;</ADR_HOME>'
EndIf
If $mADR_OTHER[$tNo] <> "" Then
$tADR_OTHER = '<ADR_OTHER>' & $mADR_OTHER[$tNo] & '</ADR_OTHER>'
Else
$tADR_OTHER = '<ADR_OTHER>;;;;;;</ADR_OTHER>'
EndIf
If $mADR_WORK[$tNo] <> "" Then
$tADR_WORK = '<ADR_WORK>' & $mADR_WORK[$tNo] & '</ADR_WORK>'
Else
$tADR_WORK = '<ADR_WORK>;;;;;;</ADR_WORK>'
EndIf
$tATTACHMENTS = '<ATTACHMENTS>0</ATTACHMENTS>'
If $aCopyWorkorOtherEmailsToPrefYN = "yes" Then
If $mEMAIL_PREF_INTERNET[$tNo] = "" Then
If $mEMAIL_WORK_INTERNET[$tNo] <> "" Then
$mEMAIL_PREF_INTERNET[$tNo] = $mEMAIL_WORK_INTERNET[$tNo]
ElseIf $mEMAIL_OTHER_INTERNET[$tNo] <> "" Then
$mEMAIL_PREF_INTERNET[$tNo] = $mEMAIL_OTHER_INTERNET[$tNo]
EndIf
EndIf
EndIf
If $aCopyPhonesNosToMobileYN = "yes" Then
If $mTEL_CELL_VOICE[$tNo] = "" Then
If $mTEL_HOME_VOICE[$tNo] <> "" Then
$mTEL_CELL_VOICE[$tNo] = $mTEL_HOME_VOICE[$tNo]
ElseIf $mTEL_WORK_VOICE[$tNo] <> "" Then
$mTEL_CELL_VOICE[$tNo] = $mTEL_WORK_VOICE[$tNo]
ElseIf $mTEL_PAGER_VOICE[$tNo] <> "" Then
$mTEL_CELL_VOICE[$tNo] = $mTEL_PAGER_VOICE[$tNo]
EndIf
EndIf
EndIf
If $mBDAY[$tNo] <> "" Then $tBDAY = '<BDAY>' & $mBDAY[$tNo] & '</BDAY>'
If $mEMAIL_OTHER_INTERNET[$tNo] <> "" Then $tEMAIL_OTHER_INTERNET = '<EMAIL_OTHER_INTERNET>' & $mEMAIL_OTHER_INTERNET[$tNo] & '</EMAIL_OTHER_INTERNET>'
If $mEMAIL_PREF_INTERNET[$tNo] <> "" Then $tEMAIL_PREF_INTERNET = '<EMAIL_PREF_INTERNET>' & $mEMAIL_PREF_INTERNET[$tNo] & '</EMAIL_PREF_INTERNET>'
If $mEMAIL_WORK_INTERNET[$tNo] <> "" Then $tEMAIL_WORK_INTERNET = '<EMAIL_WORK_INTERNET>' & $mEMAIL_WORK_INTERNET[$tNo] & '</EMAIL_WORK_INTERNET>'
If $mFN[$tNo] <> "" Then $tFN = '<FN>' & $mFN[$tNo] & '</FN>'
$tIMPORTANCE = '<IMPORTANCE>0</IMPORTANCE>'
If $mTYPE_MSMESSENGER[$tNo] <> "" Then $tIM_TYPE_MSMESSENGER = '<IM_TYPE-MSMESSENGER>' & $mTYPE_MSMESSENGER[$tNo] & '</IM_TYPE-MSMESSENGER>'
Local $tDate = _DateDiff('s', "1977/05/14 00:00:00", _NowCalc())
$tMODIFIEDDATE = '<MODIFIEDDATE>' & $tDate & '</MODIFIEDDATE>'
If $mN[$tNo] <> "" Then
$tN = '<N>' & $mN[$tNo] & '</N>'
Else
$tN = '<N>;;;;</N>'
EndIf
If $mNICKNAME[$tNo] <> "" Then $tNICKNAME = '<NICKNAME>' & $mNICKNAME[$tNo] & '</NICKNAME>'
If $mORG[$tNo] <> "" Then
$tORG = '<ORG>' & $mORG[$tNo] & '</ORG>'
Else
$tORG = '<ORG>;</ORG>'
EndIf
$tREAD = '<READ>0</READ>'
$tSIZE = '<SIZE>0</SIZE>'
$tSTATE = '<STATE>0</STATE>'
If $mTEL_CELL_VOICE[$tNo] <> "" Then $tTEL_CELL_VOICE = '<TEL_CELL_VOICE>' & $mTEL_CELL_VOICE[$tNo] & '</TEL_CELL_VOICE>'
If $mTEL_HOME_VOICE[$tNo] <> "" Then $tTEL_HOME_VOICE = '<TEL_HOME_VOICE>' & $mTEL_HOME_VOICE[$tNo] & '</TEL_HOME_VOICE>'
If $mTEL_PAGER_VOICE[$tNo] <> "" Then $tTEL_PAGER_VOICE = '<TEL_PAGER_VOICE>' & $mTEL_PAGER_VOICE[$tNo] & '</TEL_PAGER_VOICE>'
If $mTEL_WORK_VOICE[$tNo] <> "" Then $tTEL_WORK_VOICE = '<TEL_WORK_VOICE>' & $mTEL_WORK_VOICE[$tNo] & '</TEL_WORK_VOICE>'
If $mTITLE[$tNo] <> "" Then $tTITLE = '<TITLE>' & $mTITLE[$tNo] & '</TITLE>'
$tELEMENT_END = '</ELEMENT>'
$mIndex[$tNo] = $tELEMNT_ID & $tADR_HOME & $tADR_OTHER & $tADR_WORK & $tAGENT & $tATTACHMENTS & $tBDAY & $tEMAIL_OTHER_INTERNET & $tEMAIL_PREF_INTERNET & $tEMAIL_WORK_INTERNET & $tFN & $tIMPORTANCE & $tMODIFIEDDATE & $tIM_TYPE_MSMESSENGER & $tN & $tNICKNAME & $tORG & $tREAD & $tSIZE & $tSTATE & $tTEL_CELL_VOICE & $tTEL_HOME_VOICE & $tTEL_PAGER_VOICE & $tTEL_WORK_VOICE & $tTITLE & $tELEMENT_END
Next
Local $tFileRead = FileRead($aIndex_XML, 100000000)
If StringInStr($tFileRead, "<BASEELEMENT LASTUID=") Then
Else
$tFileRead = '<BASEELEMENT LASTUID="x00000000"></BASEELEMENT>'
EndIf
Local $tAdd = ""
For $tNo = 0 To ($aNumberOfContacts - 1)
$tAdd &= $mIndex[$tNo]
Next
$tFileRead = StringReplace($tFileRead, '</BASEELEMENT>', $tAdd & '</BASEELEMENT>')
Local $tUpdatedLastUID = StringLower(Hex($aLastUIDDec, 8))
$tFileRead = StringRegExpReplace($tFileRead, '<BASEELEMENT LASTUID="x........">', '<BASEELEMENT LASTUID="x' & $tUpdatedLastUID & '">')
If FileExists($aIndex_XML) Then FileMove($aIndex_XML, $aExportFolder & "\" & _FileNameRemoveExt($aIndex_XML) & "_Backup_" & @YEAR & "-" & @MON & "-" & @MDAY & ".xml", 1)
FileWrite($aIndex_XML, $tFileRead)
EndFunc
Func _WriteCategories()
For $tNo = 0 To ($aNumberOfContacts - 1)
ControlSetText($aSplash, "", "Static1", "Prosessing categories." & @CRLF & "Processing contact " & $tNo & " of " & $aNumberOfContacts & @CRLF & @CRLF & "Timer:" & Round(_Timer_Diff($aTimerStart) / 1000) & "s")
If $gCategories[$tNo] <> "" Then
Local $tSplit = StringSplit($gCategories[$tNo], ",", 2)
For $t = 0 To (UBound($tSplit) - 1)
$tSplit[$t] = _ConvertCustomLabel($tSplit[$t])
$tSplit[$t] = _RemoveSpecialChars($tSplit[$t])
Local $tFound = False
For $i = 0 To (UBound($gCategNames) - 1)
If $tSplit[$t] = $gCategNames[$i] Then
$gCategContacts[$i] &= $mFileName[$tNo] & @CRLF
$tFound = True
ExitLoop
EndIf
Next
If $tFound = False Then
_ArrayAdd($gCategNames, $tSplit[$t])
_ArrayAdd($gCategContacts, $mFileName[$tNo] & @CRLF)
EndIf
Next
EndIf
Next
Global $mGroupsFileName[0]
Local $tAdd = ""
Local $tFileRead = FileRead($aGroups_XML, 100000000)
If StringInStr($tFileRead, "<BASEELEMENT>") Then
Else
Local $tFileRead = '<BASEELEMENT>' & @CRLF & '</BASEELEMENT>'
FileWrite($aGroups_XML, $tFileRead)
EndIf
For $tNo = 0 To (UBound($gCategNames) - 1)
_ArrayAdd($mGroupsFileName, _GenerateFileName() & ".GRP")
FileWrite($aExportFolder & "\" & $mGroupsFileName[$tNo], $gCategContacts[$tNo])
$tAdd &= '  <ELEMENT ID="' & $mGroupsFileName[$tNo] & '">' & @CRLF & '    <NAME>' & $gCategNames[$tNo] & '</NAME>' & @CRLF & '  </ELEMENT>' & @CRLF
Next
$tFileRead = StringReplace($tFileRead, '</BASEELEMENT>', $tAdd & '</BASEELEMENT>')
If FileExists($aGroups_XML) Then FileMove($aGroups_XML, $aExportFolder & "\" & _FileNameRemoveExt($aGroups_XML) & "_Backup_" & @YEAR & "-" & @MON & "-" & @MDAY & ".XML")
FileWrite($aGroups_XML, $tFileRead)
EndFunc
Func _FileNameRemoveExt($tPath)
Local $xFileName = _PathSplit($tPath, "", "", "", "")
If UBound($xFileName) > 4 Then
Return $xFileName[3]
Else
Return "[ERROR]"
EndIf
EndFunc
Func _RemoveExtraSpaces($tText)
Return StringRegExpReplace($tText, "(?m)\h{2,}", " ")
EndFunc
Func _GetLastUID()
Local $tFileRead = FileRead($aIndex_XML, 50)
Global $aLastUIDHex = _ArrayToString(_StringBetween($tFileRead, '<BASEELEMENT LASTUID="x', '">'))
Global $aLastUIDDec = Dec($aLastUIDHex)
EndFunc
Func _RemoveComma($tTxt)
Return StringReplace($tTxt, ",", " ")
EndFunc
Func _FormatPhoneNumber($tNumber)
If StringLen(StringRegExpReplace($tNumber, "[^0-9]", "")) > 7 Then
Local $tReturn = ""
For $tX1 = 1 To StringLen($tNumber)
Local $tDig1 = StringMid($tNumber, $tX1, 1)
Local $tDig2 = StringRegExpReplace($tDig1, "[^0-9]", "")
If $tDig2 <> "" Then
$tReturn &= $tDig2
If StringLeft($tReturn, 1) = "1" Then $tReturn = StringTrimLeft($tReturn, 1)
If StringLen($tReturn) >= 10 Then
$tReturn &= StringRight($tNumber, (StringLen($tNumber) - $tX1))
ExitLoop
EndIf
EndIf
Next
Else
Return $tNumber
EndIf
If StringLen($tReturn) >= 10 Then
$tReturn = "(" & StringMid($tReturn, 1, 3) & ") " & StringMid($tReturn, 4, 3) & "-" & StringMid($tReturn, 7, 4) & " " & StringRight($tReturn, (StringLen($tNumber) - 15))
$tReturn = StringReplace($tReturn, "  ", " ")
ElseIf StringLen($tReturn) = 7 Then
$tReturn = StringMid($tReturn, 1, 3) & "-" & StringMid($tReturn, 4, 4)
Else
EndIf
Return $tReturn
EndFunc
Func _Splash($tTxt, $tTime = 0, $tWidth = 400, $tHeight = 125)
Local $tPID = SplashTextOn($aUtilName, $tTxt, $tWidth, $tHeight, -1, -1, $DLG_MOVEABLE, "")
If $tTime > 0 Then
Sleep($tTime)
SplashOff()
EndIf
Return $tPID
EndFunc
Func _RemoveTrailingSlash($aString)
Local $bString = StringRight($aString, 1)
If $bString = "\" Then $aString = StringTrimRight($aString, 1)
Return $aString
EndFunc
Func _ConvertCustomLabel($tTxt)
If StringIsUpper(StringLeft($tTxt, 1)) Then
Return $tTxt
Else
Local $tLen = StringLen($tTxt)
Local $tChar = StringMid($tTxt, 1, 1)
Local $tNew = StringUpper($tChar)
For $t = 2 To $tLen
Local $tChar = StringMid($tTxt, $t, 1)
If StringIsUpper($tChar) Then
$tNew &= " " & StringUpper($tChar)
Else
$tNew &= $tChar
EndIf
Next
Return StringReplace($tNew, "  ", " ")
EndIf
EndFunc
Func _RemoveSpecialChars($aString)
Return StringRegExpReplace($aString, "(?i)([^a-z0-9-_ .])", "")
EndFunc
Func _ConvertToBase64($tDat1)
$tDat = StringToBinary($tDat1)
$objXML = ObjCreate("MSXML2.DOMDocument")
$objNode = $objXML.createElement("b64")
$objNode.dataType = "bin.base64"
$objNode.nodeTypedValue = $tDat
Return $objNode.Text
EndFunc
Func UtilUpdate($tLink, $tDL, $tUtil, $tUtilName, $tSplash = 0, $tUpdate = "show")
Local $tUtilUpdateAvailableTF = False
$aUpdateAutoUtil = False
If $tUpdate = "show" Then _Splash("Checking for update", 0, 350, 75)
Local $tVer[2]
Local $sFilePath = $aFolderTemp & $aUtilName & "_latest_ver.tmp"
$iGet = _InetGetMulti(20, $sFilePath, $tLink)
If $iGet = "Error" Then
If $tUpdate = "show" Then _Splash("Update check failed." & @CRLF & "Please try again later.", 2000, 350, 75)
Else
$tVer = StringSplit($iGet, "^", 2)
If UBound($tVer) < 2 Then Return False
Local $tTxt1 = ReplaceCRLF(ReplaceCRwithCRLF($tVer[1]))
If $tVer[0] = $tUtil Then
If $tUpdate = "show" Then _Splash($aUtilName & " up to date.", 2000, 350, 75)
Else
$tUtilUpdateAvailableTF = True
If ($tUpdate = "show") Or ($tUpdate = "auto") Or ($tUpdate = "UpdateOnly") Then
SplashOff()
If ($tUpdate = "Auto") And ($aUpdateAutoUtil = "yes") Then
Local $tMB = 6
Else
Local $tMB = MsgBox($MB_YESNOCANCEL, $aUtilVer, "New " & $aUtilName & " update available. " & @CRLF & "Installed version: " & $tUtil & @CRLF & "Latest version: " & $tVer[0] & @CRLF & @CRLF & "Notes: " & @CRLF & $tVer[1] & @CRLF & @CRLF & "Click (YES) to download update to " & @CRLF & @ScriptDir & @CRLF & "Click (NO) to stop checking for updates." & @CRLF & "Click (CANCEL) to skip this update check.", 15)
EndIf
If $tMB = 6 Then
_Splash(" Downloading latest version of " & @CRLF & $tUtilName)
Local $tZIP = @ScriptDir & "\" & $tUtilName & "_" & $tVer[0] & ".zip"
If FileExists($tZIP) Then FileDelete($tZIP)
If FileExists($tUtilName & "_" & $tVer[0] & ".exe") Then FileDelete($tUtilName & "_" & $tVer[0] & ".exe")
If FileExists(@ScriptDir & "\readme.txt") Then FileDelete(@ScriptDir & "\readme.txt")
InetGet($tDL, $tZIP, 1)
_ExtractZipAll($tZIP, @ScriptDir)
If Not FileExists(@ScriptDir & "\" & $tUtilName & "_" & $tVer[0] & ".exe") Then
SplashOff()
$tMB = MsgBox($MB_OKCANCEL, $aUtilVer, "Utility update download failed . . . " & @CRLF & "Go to """ & $tLink & """ to download latest version." & @CRLF & @CRLF & "Click (OK), (CANCEL), or wait 60 seconds, to resume current version.", 60)
Else
SplashOff()
If ($tUpdate = "Auto") And ($aUpdateAutoUtil = "yes") Then
$tMB = MsgBox($MB_OKCANCEL, $aUtilVer, "Auto utility update download complete. . . " & @CRLF & @CRLF & "Click (OK) to run new version or wait 60 seconds (servers will remain running) OR" & @CRLF & "Click (CANCEL) to resume current version.", 60)
If $tMB = 1 Then
ElseIf $tMB = -1 Then
$tMB = 1
ElseIf $tMB = 2 Then
EndIf
Else
$tMB = MsgBox($MB_OKCANCEL, $aUtilVer, "Utility update download complete. . . " & @CRLF & @CRLF & "Click (OK) to run new version (servers will remain running) OR" & @CRLF & "Click (CANCEL), or wait 15 seconds, to resume current version.", 15)
EndIf
If $tMB = 1 Then
Local $xArray[13]
$xArray[0] = '@echo off'
$xArray[1] = 'echo --------------------------------------------'
$xArray[2] = 'echo  Waiting 5 seconds for shutdown to complete'
$xArray[3] = 'echo --------------------------------------------'
$xArray[4] = 'timeout 5'
$xArray[5] = 'start "Starting ' & $aUtilName & '" "' & @ScriptDir & "\" & $tUtilName & "_" & $tVer[0] & ".exe" & '"'
$xArray[6] = 'echo --------------------------------------------'
$xArray[7] = 'echo  ' & $aUtilName & ' started . . .'
$xArray[8] = 'echo --------------------------------------------'
$xArray[9] = 'timeout 3'
$xArray[10] = 'exit'
Local $tBatFile = $aFolderTemp & $aUtilName & "_Delay_Restart.bat"
FileDelete($tBatFile)
_FileWriteFromArray($tBatFile, $xArray)
If FileExists($tBatFile) Then
Run($tBatFile)
Else
Run(@ScriptDir & "\" & $tUtilName & "_" & $tVer[0] & ".exe")
EndIf
Exit
Else
_Splash("Utility update check canceled by user." & @CRLF & "Resuming utility . . .", 2000)
EndIf
EndIf
ElseIf $tMB = 7 Then
$aCheckForUpdatesYN = "no"
IniWrite($aIniFile, $aINIHeader, "CheckForUpdatesYN", $aCheckForUpdatesYN)
_Splash("Utility update check disabled." & @CRLF & "To enable update check, change [CheckForUpdatesYN=yes] in the .ini.", 5000, 500)
ElseIf $tMB = 2 Then
EndIf
EndIf
EndIf
EndIf
Return $tUtilUpdateAvailableTF
EndFunc
Func _InetGetMulti($tCnt, $tFile, $tLink1, $tLink2 = "0")
FileDelete($tFile)
Local $i = 0
Local $tTmp1 = InetGet($tLink1, $tFile, $INET_FORCERELOAD, $INET_DOWNLOADBACKGROUND)
Do
Sleep(100)
$i += 1
Until InetGetInfo($tTmp1, $INET_DOWNLOADCOMPLETE) Or $i = $tCnt
InetClose($tTmp1)
If $i = $tCnt And $tLink2 <> "0" Then
$tTmp2 = InetGet($tLink2, $tFile, $INET_FORCERELOAD, $INET_DOWNLOADBACKGROUND)
Do
Sleep(100)
$i += 1
Until InetGetInfo($tTmp2, $INET_DOWNLOADCOMPLETE) Or $i = $tCnt
InetClose($tTmp2)
EndIf
Local $hFileOpen = FileOpen($tFile, 0)
Local $hFileRead = FileRead($hFileOpen, 100000000)
If $hFileOpen = -1 Then
InetClose($tTmp1)
Sleep(200)
FileClose($hFileOpen)
Local $hFileRead = _INetGetSource($tLink1)
If @error Then
If $tLink2 <> "0" Then
$hFileRead = _INetGetSource($tLink2)
If @error Then
Return "Error"
Else
FileClose($hFileOpen)
FileDelete($tFile)
FileWrite($tFile, $hFileRead)
EndIf
Else
Return True
EndIf
Else
FileClose($hFileOpen)
FileDelete($tFile)
FileWrite($tFile, $hFileRead)
EndIf
Else
FileClose($hFileOpen)
EndIf
Return $hFileRead
EndFunc
Func ReplaceCRwithCRLF($sString)
Return StringRegExpReplace($sString, '(*BSR_ANYCRLF)\R', @CRLF)
EndFunc
Func ReplaceCRLF($tMsg0)
Return StringReplace($tMsg0, @CRLF, "|")
EndFunc
Func _ExtractZipAll($sZipFile, $sDestinationFolder, $sFolderStructure = "")
Local $i
Do
$i += 1
$sTempZipFolder = @TempDir & "\Temporary Directory " & $i & " for " & StringRegExpReplace($sZipFile, ".*\\", "")
Until Not FileExists($sTempZipFolder)
Local $oShell = ObjCreate("Shell.Application")
If Not IsObj($oShell) Then
Return SetError(1, 0, 0)
EndIf
Local $oDestinationFolder = $oShell.NameSpace($sDestinationFolder)
If Not IsObj($oDestinationFolder) Then
DirCreate($sDestinationFolder)
EndIf
Local $oOriginFolder = $oShell.NameSpace($sZipFile & "\" & $sFolderStructure)
If Not IsObj($oOriginFolder) Then
Return SetError(3, 0, 0)
EndIf
Local $oOriginFile = $oOriginFolder.Items()
If Not IsObj($oOriginFile) Then
Return SetError(4, 0, 0)
EndIf
$oDestinationFolder.CopyHere($oOriginFile, 20)
DirRemove($sTempZipFolder, 1)
Return 1
EndFunc
Func _ShowGUI()
#Region ### START Koda GUI section ### Form=K:\AutoIT\_MyProgs\GoogleToMailEnable\Koda\googletome(b7).kxf
Global $W_MainWindow = GUICreate("Google Contacts to Mail Enable", 902, 578, -1, -1, BitOR($GUI_SS_DEFAULT_GUI, $WS_SIZEBOX, $WS_THICKFRAME))
GUISetIcon($aIconFile, 99)
GUISetBkColor(0x808080)
GUISetOnEvent($GUI_EVENT_CLOSE, "W_MainWindowClose")
GUISetOnEvent($GUI_EVENT_MINIMIZE, "W_MainWindowMinimize")
GUISetOnEvent($GUI_EVENT_MAXIMIZE, "W_MainWindowMaximize")
GUISetOnEvent($GUI_EVENT_RESTORE, "W_MainWindowRestore")
Global $G_MainWindow = GUICtrlCreateGroup("", 4, -3, 891, 571)
GUICtrlSetFont(-1, 12, 800, 0, "MS Sans Serif")
Global $G_SourceGroup = GUICtrlCreateGroup("Source: Google vCard", 23, 85, 855, 267)
GUICtrlSetFont(-1, 9, 400, 0, "MS Sans Serif")
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
Global $I_SourceFile = GUICtrlCreateInput("contacts.vcf", 31, 320, 471, 21)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
GUICtrlSetOnEvent(-1, "I_SourceFileChange")
GUICtrlSetTip(-1, "Click to select the source .vcf file")
Global $L_SourceInstructions = GUICtrlCreateLabel("Export your contacts from Google as vCard (for IOS Contacts)", 34, 104, 443, 23)
GUICtrlSetFont(-1, 12, 400, 2, "Tahoma")
GUICtrlSetColor(-1, 0xFFFF00)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
GUICtrlSetOnEvent(-1, "L_SourceInstructionsClick")
Global $P_GoogleContactsExportImage = GUICtrlCreatePic($aFolderTemp & "GoogleExport1.jpg", 505, 96, 368, 218)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
GUICtrlSetTip(-1, "Click to enlarge image")
GUICtrlSetOnEvent(-1, "P_GoogleContactsExportImageClick")
Global $L_SelectSource = GUICtrlCreateLabel("Select Source VCF File:", 31, 294, 190, 23)
GUICtrlSetFont(-1, 12, 800, 2, "Tahoma")
GUICtrlSetColor(-1, 0xFFFF00)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
GUICtrlSetOnEvent(-1, "L_SelectSourceClick")
Global $L_SourceDefault = GUICtrlCreateLabel("(Default: contacts.vcf)", 590, 322, 130, 20)
GUICtrlSetFont(-1, 10, 400, 0, "Tahoma")
GUICtrlSetColor(-1, 0x00FF00)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
GUICtrlSetOnEvent(-1, "L_SourceDefaultClick")
Global $B_SelectSourceFile = GUICtrlCreateButton("Select File", 503, 318, 83, 25)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
GUICtrlSetBkColor(-1, $cButtonFadedBlue)
GUICtrlSetTip(-1, "Click to select the source .vcf file")
GUICtrlSetOnEvent(-1, "B_SelectSourceFileClick")
Global $C_ImportPhotos = GUICtrlCreateCheckbox("Import photos", 36, 136, 117, 17)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
GUICtrlSetTip(-1, "Import Google Contact photos")
GUICtrlSetOnEvent(-1, "C_ImportPhotosClick")
Global $C_ReformatPhoneNos = GUICtrlCreateCheckbox("Reformat phone numbers", 36, 159, 145, 17)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
GUICtrlSetTip(-1, "Convert phone numbers to (111) 222-3333 format. ie. +1 231 453.6754  --> (231) 453-6754")
GUICtrlSetOnEvent(-1, "C_ReformatPhoneNosClick")
Global $C_ImportGroups = GUICtrlCreateCheckbox("Import Groups", 36, 182, 117, 17)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
GUICtrlSetTip(-1, "Import and create Groups")
GUICtrlSetOnEvent(-1, "C_ImportGroupsClick")
Global $C_CopyPhone = GUICtrlCreateCheckbox("Copy Phone #s to Preferred (M)", 219, 134, 261, 17)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
GUICtrlSetTip(-1, "If contact does not have a mobile number, copy Home/Work/Pager to Mobile so that it shows on main screen")
GUICtrlSetOnEvent(-1, "C_CopyPhoneClick")
Global $C_CopyEmail = GUICtrlCreateCheckbox("Copy Emails to Preferred", 219, 158, 261, 17)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
GUICtrlSetTip(-1, "If contact does not have a preferred (main) email, copy Work/Other email to preferred so that it shows on contact page.")
GUICtrlSetOnEvent(-1, "C_CopyEmailClick")
Global $L_WorkPhoneAliases = GUICtrlCreateLabel("Work Phone Aliases", 36, 238, 100, 17)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
GUICtrlSetTip(-1, "(Comma separated) If no work phone, import these custom labels as work phone. ")
GUICtrlSetOnEvent(-1, "L_WorkPhoneAliasesClick")
Global $L_HomePhoneAliases = GUICtrlCreateLabel("Home Phone Aliases", 36, 262, 102, 17)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
GUICtrlSetTip(-1, "(Comma separated) If no home phone, import these custom labels as home phone. ie. If you used custom label, Personal, then any email labelled as Personal will import as Home")
GUICtrlSetOnEvent(-1, "L_HomePhoneAliasesClick")
Global $I_WorkAliases = GUICtrlCreateInput("I_WorkAliases", 144, 235, 359, 21)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
GUICtrlSetOnEvent(-1, "I_WorkAliasesChange")
GUICtrlSetTip(-1, "(Comma separated) If no work phone, import these custom labels as work phone. ")
Global $I_HomeAliases = GUICtrlCreateInput("I_HomeAliases", 144, 259, 359, 21)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
GUICtrlSetOnEvent(-1, "I_HomeAliasesChange")
GUICtrlSetTip(-1, "(Comma separated) If no home phone, import these custom labels as home phone. ie. If you used custom label, Personal, then any email labelled as Personal will import as Home")
Global $L_MobilePhoneAlias = GUICtrlCreateLabel("Mobile Phone Aliases", 36, 213, 105, 17)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
GUICtrlSetTip(-1, "(Comma separated) If no mobile phone, import these custom labels as mobile phone. ie. If you used custom label, Cell, then any phone labelled as Cell will import as Mobile.")
GUICtrlSetOnEvent(-1, "L_MobilePhoneAliasClick")
Global $I_MobileAliases = GUICtrlCreateInput("I_MobileAliases", 144, 211, 359, 21)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
GUICtrlSetOnEvent(-1, "I_MobileAliasesChange")
GUICtrlSetTip(-1, "(Comma separated) If no mobile phone, import these custom labels as mobile phone. ie. If you used custom label, Cell, then any phone labelled as Cell will import as Mobile.")
GUICtrlCreateGroup("", -99, -99, 1, 1)
Global $P_MailEnableLogo = GUICtrlCreatePic($aFolderTemp & "MailEnableLogo1.jpg", 23, 21, 173, 58, BitOR($GUI_SS_DEFAULT_PIC, $WS_BORDER))
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
GUICtrlSetTip(-1, "Click to visit https://www.mailenable.com")
GUICtrlSetOnEvent(-1, "P_MailEnableLogoClick")
Global $L_Title = GUICtrlCreateLabel("Google Contacts to Mail Enable", 228, 31, 484, 41)
GUICtrlSetFont(-1, 24, 800, 0, "MS Sans Serif")
GUICtrlSetColor(-1, $cSWTitle)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
GUICtrlSetTip(-1, "Click to open the GitHub source website.")
GUICtrlSetOnEvent(-1, "L_TitleClick")
Global $G_ImportGroup = GUICtrlCreateGroup("Import", 22, 491, 253, 61)
GUICtrlSetFont(-1, 9, 400, 0, "MS Sans Serif")
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
Global $B_Exit = GUICtrlCreateButton("Exit", 203, 512, 58, 29)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
GUICtrlSetTip(-1, "Click to exit")
GUICtrlSetBkColor(-1, $cSWButtonStopServer)
GUICtrlSetOnEvent(-1, "B_ExitClick")
Global $B_ImportContacts = GUICtrlCreateButton("START (Import Contacts)", 37, 512, 158, 29)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
GUICtrlSetBkColor(-1, $cSWButtonStartServer)
GUICtrlSetTip(-1, "Click to start the importation process.")
GUICtrlSetOnEvent(-1, "B_ImportContactsClick")
GUICtrlCreateGroup("", -99, -99, 1, 1)
Global $P_GoogleContactsLogo = GUICtrlCreatePic($aFolderTemp & "GoogleLogo.jpg", 819, 21, 58, 58)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
GUICtrlSetTip(-1, "Click to open https://contacts.google.com")
GUICtrlSetOnEvent(-1, "P_GoogleContactsLogoClick")
Global $G_OutputGroup = GUICtrlCreateGroup("Output: Mail Enable Mailbox Contact Folder", 22, 363, 857, 115)
GUICtrlSetFont(-1, 9, 400, 0, "MS Sans Serif")
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
Global $I_OutputContactsFolder = GUICtrlCreateInput("Contacts", 31, 417, 475, 21)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
GUICtrlSetOnEvent(-1, "I_OutputContactsFolderChange")
GUICtrlSetTip(-1, "Click to select the output folder: typically the {MailBox}/Contacts folder")
Global $L_SelectContactsFolder = GUICtrlCreateLabel("Select Mail Enable Mailbox Contact Folder:", 33, 392, 352, 23)
GUICtrlSetFont(-1, 12, 800, 2, "Tahoma")
GUICtrlSetColor(-1, 0xFFFF00)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
GUICtrlSetOnEvent(-1, "L_SelectContactsFolderClick")
Global $L_SelectContactsDefault = GUICtrlCreateLabel("(Default: C:\Program Files (x86)\Mail Enable\Postoffices\example.com\MAILROOT\joe\Contacts)", 31, 442, 555, 20)
GUICtrlSetFont(-1, 10, 400, 0, "Tahoma")
GUICtrlSetColor(-1, 0x00FF00)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
GUICtrlSetOnEvent(-1, "L_SelectContactsDefaultClick")
Global $B_OutputContactsSelectFolder = GUICtrlCreateButton("Select Folder", 507, 415, 83, 25)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
GUICtrlSetBkColor(-1, $cButtonFadedBlue)
GUICtrlSetTip(-1, "Click to select the output folder: typically the {MailBox}/Contacts folder")
GUICtrlSetOnEvent(-1, "B_OutputContactsSelectFolderClick")
Global $C_OpenContactsFolder = GUICtrlCreateCheckbox("Open Contacts Folder When Done", 664, 448, 201, 21)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
GUICtrlSetTip(-1, "Open the Contacts folder in File Explorer when complete.")
GUICtrlSetOnEvent(-1, "C_OpenContactsFolderClick")
GUICtrlCreateGroup("", -99, -99, 1, 1)
Global $L_FooterPhoenixURL = GUICtrlCreateLabel("http://www.Phoenix125.com", 644, 535, 168, 20)
GUICtrlSetFont(-1, 10, 400, 0, "Tahoma")
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
GUICtrlSetOnEvent(-1, "L_FooterPhoenixURLClick")
Global $L_FooterProgName = GUICtrlCreateLabel($aUtilName & " " & $aUtilVer, 581, 507, 231, 20, $SS_RIGHT)
GUICtrlSetFont(-1, 10, 400, 0, "Tahoma")
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
GUICtrlSetOnEvent(-1, "L_FooterProgNameClick")
Global $P_PhoenixLogo = GUICtrlCreatePic($aFolderTemp & "phoenixlogo.jpg", 818, 495, 60, 60)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
GUICtrlSetTip(-1, "Click to visit http://www.phoenix125.com")
GUICtrlSetOnEvent(-1, "P_PhoenixLogoClick")
Global $P_About = GUICtrlCreateIcon($aIconFile, 202, 534, 511, 24, 24)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
GUICtrlSetTip(-1, "About")
GUICtrlSetOnEvent(-1, "P_AboutClick")
Global $P_Discord = GUICtrlCreateIcon($aIconFile, 201, 499, 511, 24, 24)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
GUICtrlSetTip(-1, "Discord Page")
GUICtrlSetOnEvent(-1, "P_DiscordClick")
Global $P_Discussion = GUICtrlCreateIcon($aIconFile, 203, 463, 511, 24, 24)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
GUICtrlSetTip(-1, "Discussion Forum")
GUICtrlSetOnEvent(-1, "P_DiscussionClick")
Global $P_Help = GUICtrlCreateIcon($aIconFile, 204, 425, 511, 24, 24)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
GUICtrlSetTip(-1, "Online Help")
GUICtrlSetOnEvent(-1, "P_HelpClick")
Global $P_Update = GUICtrlCreateIcon($aIconFile, 205, 359, 511, 24, 24)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
GUICtrlSetTip(-1, "Check for updates")
GUICtrlSetOnEvent(-1, "P_UpdateClick")
GUICtrlCreateGroup("", -99, -99, 1, 1)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###
_UpdateGUI()
If $aCheckForUpdatesYN = "yes" Then
$aUpdateAvailableTF = UtilUpdate($aServerUpdateLinkVer, $aServerUpdateLinkDL, $aUtilVer, "GoogleContactsToMailEnable", 0, "UpdateOnly")
If $aUpdateAvailableTF Then
GUICtrlSetImage($P_Update, $aIconFile, 206)
Else
GUICtrlSetImage($P_Update, $aIconFile, 205)
EndIf
EndIf
EndFunc
Func _UpdateGUI()
GUICtrlSetData($I_OutputContactsFolder, $aExportFolder)
GUICtrlSetData($I_SourceFile, $gGoogleSourceFile)
GUICtrlSetData($I_MobileAliases, $aMobileAliases)
GUICtrlSetData($I_WorkAliases, $aWorkAliases)
GUICtrlSetData($I_HomeAliases, $aHomeAliases)
If $aDownloadPhotosYN = "yes" Then
GUICtrlSetState($C_ImportPhotos, $GUI_CHECKED)
Else
GUICtrlSetState($C_ImportPhotos, $GUI_UNCHECKED)
EndIf
If $aCopyPhonesNosToMobileYN = "yes" Then
GUICtrlSetState($C_CopyPhone, $GUI_CHECKED)
Else
GUICtrlSetState($C_CopyPhone, $GUI_UNCHECKED)
EndIf
If $aCopyWorkorOtherEmailsToPrefYN = "yes" Then
GUICtrlSetState($C_CopyEmail, $GUI_CHECKED)
Else
GUICtrlSetState($C_CopyEmail, $GUI_UNCHECKED)
EndIf
If $aReformatPhoneNumbersYN = "yes" Then
GUICtrlSetState($C_ReformatPhoneNos, $GUI_CHECKED)
Else
GUICtrlSetState($C_ReformatPhoneNos, $GUI_UNCHECKED)
EndIf
If $aCreateGroupsYN = "yes" Then
GUICtrlSetState($C_ImportGroups, $GUI_CHECKED)
Else
GUICtrlSetState($C_ImportGroups, $GUI_UNCHECKED)
EndIf
If $aOpenFolderWhenDoneYN = "yes" Then
GUICtrlSetState($C_OpenContactsFolder, $GUI_CHECKED)
Else
GUICtrlSetState($C_OpenContactsFolder, $GUI_UNCHECKED)
EndIf
If $aUpdateAvailableTF Then
GUICtrlSetImage($P_Update, $aIconFile, 206)
Else
GUICtrlSetImage($P_Update, $aIconFile, 205)
EndIf
EndFunc
Func _ReadGUI()
$aExportFolder = _RemoveTrailingSlash(GUICtrlRead($I_OutputContactsFolder))
$gGoogleSourceFile = GUICtrlRead($I_SourceFile)
EndFunc
Func B_ExitClick()
_Exit()
EndFunc
Func B_ImportContactsClick()
Global $aTimerStart = TimerInit()
_ReadGUI()
_IniWrite()
Global $aStartText = $aUtilName & @CRLF & @CRLF
Global $aSplash = _Splash($aStartText)
_ReadGoogleVCAR()
_SetGlobalVariablesMailEnable($aNumberOfContacts)
_ConvertGoogleToMailEnable()
_WriteMailEnable()
_UpdateIndex()
If $aCreateGroupsYN = "yes" Then _WriteCategories()
If $aOpenFolderWhenDoneYN = "yes" Then ShellExecute($aExportFolder)
SplashOff()
MsgBox(0, $aUtilName, $aStartText & "Complete. " & $aNumberOfContacts & " contacts imported." & @CRLF & @CRLF & "Timer:" & Round(_Timer_Diff($aTimerStart) / 1000) & "s")
EndFunc
Func B_OutputContactsSelectFolderClick()
Local $tExportFolder = FileSelectFolder("Please select Mail Enable Contact (Output) Folder.", $aExportFolder)
If @error Then
GUICtrlSetData($I_OutputContactsFolder, $aExportFolder)
Else
$aExportFolder = _RemoveTrailingSlash($tExportFolder)
EndIf
GUICtrlSetData($I_OutputContactsFolder, $aExportFolder)
IniWrite($aIniFile, $aINIHeader, "ExportFolder", $aExportFolder)
$aIndex_XML = $aExportFolder & "\_index.xml"
$aGroups_XML = $aExportFolder & "\_GROUPS.XML"
EndFunc
Func B_SelectSourceFileClick()
Local $tGoogleSourceFile = FileOpenDialog("Please select source file", $gGoogleSourceFile, "VCF File (*.vcf)", 3, "contacts.vcf")
If @error Then
GUICtrlSetData($I_SourceFile, $gGoogleSourceFile)
Else
$gGoogleSourceFile = $tGoogleSourceFile
EndIf
GUICtrlSetData($I_SourceFile, $gGoogleSourceFile)
IniWrite($aIniFile, $aINIHeader, "GoogleSourceFile", $gGoogleSourceFile)
EndFunc
Func C_CopyEmailClick()
If GUICtrlRead($C_CopyEmail) = $GUI_CHECKED Then
$aCopyWorkorOtherEmailsToPrefYN = "yes"
Else
$aCopyWorkorOtherEmailsToPrefYN = "no"
EndIf
IniWrite($aIniFile, $aINIHeader, "CopyWorkorOtherEmailsToPrefYN", $aCopyWorkorOtherEmailsToPrefYN)
_UpdateGUI()
EndFunc
Func C_CopyPhoneClick()
If GUICtrlRead($C_CopyPhone) = $GUI_CHECKED Then
$aCopyPhonesNosToMobileYN = "yes"
Else
$aCopyPhonesNosToMobileYN = "no"
EndIf
IniWrite($aIniFile, $aINIHeader, "CopyPhonesNosToMobileYN", $aCopyPhonesNosToMobileYN)
_UpdateGUI()
EndFunc
Func C_ImportGroupsClick()
If GUICtrlRead($C_ImportGroups) = $GUI_CHECKED Then
$aCreateGroupsYN = "yes"
Else
$aCreateGroupsYN = "no"
EndIf
IniWrite($aIniFile, $aINIHeader, "CreateGroupsYN", $aCreateGroupsYN)
_UpdateGUI()
EndFunc
Func C_ImportPhotosClick()
If GUICtrlRead($C_ImportPhotos) = $GUI_CHECKED Then
$aDownloadPhotosYN = "yes"
Else
$aDownloadPhotosYN = "no"
EndIf
IniWrite($aIniFile, $aINIHeader, "DownloadPhotosYN", $aDownloadPhotosYN)
_UpdateGUI()
EndFunc
Func C_OpenContactsFolderClick()
If GUICtrlRead($C_OpenContactsFolder) = $GUI_CHECKED Then
$aOpenFolderWhenDoneYN = "yes"
Else
$aOpenFolderWhenDoneYN = "no"
EndIf
IniWrite($aIniFile, $aINIHeader, "OpenFolderWhenDoneYN", $aOpenFolderWhenDoneYN)
_UpdateGUI()
EndFunc
Func C_ReformatPhoneNosClick()
If GUICtrlRead($C_ReformatPhoneNos) = $GUI_CHECKED Then
$aReformatPhoneNumbersYN = "yes"
Else
$aReformatPhoneNumbersYN = "no"
EndIf
IniWrite($aIniFile, $aINIHeader, "ReformatPhoneNumbersYN", $aReformatPhoneNumbersYN)
_UpdateGUI()
EndFunc
Func I_HomeAliasesChange()
EndFunc
Func I_MobileAliasesChange()
EndFunc
Func P_GoogleContactsExportImageClick()
ShellExecute($aFolderTemp & "GoogleExport1.jpg")
EndFunc
Func W_MainWindowClose()
_Exit()
EndFunc
Func W_MainWindowMaximize()
EndFunc
Func W_MainWindowMinimize()
EndFunc
Func W_MainWindowRestore()
EndFunc
Func I_OutputContactsFolderChange()
$aExportFolder = _RemoveTrailingSlash(GUICtrlRead($I_OutputContactsFolder))
IniWrite($aIniFile, $aINIHeader, "ExportFolder", $aExportFolder)
$aIndex_XML = $aExportFolder & "\_index.xml"
$aGroups_XML = $aExportFolder & "\_GROUPS.XML"
EndFunc
Func I_SourceFileChange()
$gGoogleSourceFile = GUICtrlRead($I_SourceFile)
IniWrite($aIniFile, $aINIHeader, "GoogleSourceFile", $gGoogleSourceFile)
EndFunc
Func I_WorkAliasesChange()
EndFunc
Func L_FooterPhoenixURLClick()
ShellExecute("http://www.phoenix125.com")
EndFunc
Func L_FooterProgNameClick()
ShellExecute("https://github.com/phoenix125/GoogleContactsToMailEnable")
EndFunc
Func L_HomePhoneAliasesClick()
EndFunc
Func L_MobilePhoneAliasClick()
EndFunc
Func L_SelectContactsDefaultClick()
EndFunc
Func L_SelectContactsFolderClick()
EndFunc
Func L_SelectSourceClick()
EndFunc
Func L_SourceDefaultClick()
EndFunc
Func L_SourceInstructionsClick()
EndFunc
Func L_TitleClick()
ShellExecute("https://github.com/phoenix125/GoogleContactsToMailEnable")
EndFunc
Func L_WorkPhoneAliasesClick()
EndFunc
Func P_AboutClick()
MsgBox($MB_SYSTEMMODAL, $aUtilName, $aUtilName & @CRLF & "Version: " & $aUtilVer & @CRLF & @CRLF & "Install Path: " & @ScriptDir & @CRLF & @CRLF & "Discord: http://discord.gg/EU7pzPs" & @CRLF & "Website: http://www.phoenix125.com", 15)
EndFunc
Func P_DiscordClick()
ShellExecute("http://discord.gg/EU7pzPs")
EndFunc
Func P_DiscussionClick()
ShellExecute("https://phoenix125.createaforum.com/google-contacts-to-mail-enable-discussion/")
EndFunc
Func P_GoogleContactsLogoClick()
ShellExecute("https://contacts.google.com")
EndFunc
Func P_HelpClick()
ShellExecute("http://www.phoenix125.com/share/mailenable/ReadMe.pdf")
EndFunc
Func P_MailEnableLogoClick()
ShellExecute("https://www.mailenable.com")
EndFunc
Func P_PhoenixLogoClick()
ShellExecute("http://www.phoenix125.com")
EndFunc
Func P_UpdateClick()
$aUpdateAvailableTF = UtilUpdate($aServerUpdateLinkVer, $aServerUpdateLinkDL, $aUtilVer, "GoogleContactsToMailEnable", 0, "Show")
If $aUpdateAvailableTF Then
GUICtrlSetImage($P_Update, $aIconFile, 206)
Else
GUICtrlSetImage($P_Update, $aIconFile, 205)
EndIf
EndFunc
Func _Exit()
GUIDelete($W_MainWindow)
MsgBox(0, $aUtilName, "Thank you for using " & $aUtilName & "." & @CRLF & @CRLF & "Please report any problems or comments to: " & @CRLF & "Discord: http://discord.gg/EU7pzPs or " & @CRLF & "email: kim@phoenix125.com" & @CRLF & @CRLF & "Visit http://www.Phoenix125.com", 20)
Exit
EndFunc
