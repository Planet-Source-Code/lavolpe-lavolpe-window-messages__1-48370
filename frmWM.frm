VERSION 5.00
Begin VB.Form frmWM 
   Caption         =   "Windows Messages - Cheat Sheet"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Search for Window Message"
      Height          =   1170
      Left            =   4110
      TabIndex        =   17
      Top             =   2145
      Width           =   3420
      Begin VB.TextBox txtCriteria 
         Height          =   285
         Left            =   1290
         TabIndex        =   6
         Text            =   "Enter criteria here"
         Top             =   270
         Width           =   2025
      End
      Begin VB.OptionButton optCriteria 
         Caption         =   "In Description or message name"
         Height          =   210
         Index           =   1
         Left            =   105
         TabIndex        =   5
         Top             =   555
         Width           =   2655
      End
      Begin VB.OptionButton optCriteria 
         Caption         =   "By value"
         Height          =   210
         Index           =   0
         Left            =   105
         TabIndex        =   4
         Top             =   285
         Value           =   -1  'True
         Width           =   1365
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Previous"
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   2205
         TabIndex        =   9
         Top             =   810
         Width           =   1140
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Next"
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   1245
         TabIndex        =   8
         Top             =   810
         Width           =   960
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Find First"
         Height          =   315
         Index           =   0
         Left            =   105
         TabIndex        =   7
         Top             =   810
         Width           =   1140
      End
   End
   Begin VB.TextBox txtDeclare 
      Height          =   285
      Left            =   60
      TabIndex        =   10
      Top             =   4875
      Width           =   7560
   End
   Begin VB.TextBox txtValue 
      Height          =   315
      Index           =   1
      Left            =   5940
      TabIndex        =   3
      Top             =   1710
      Width           =   1365
   End
   Begin VB.TextBox txtValue 
      Height          =   315
      Index           =   0
      Left            =   4260
      TabIndex        =   2
      Top             =   1710
      Width           =   1365
   End
   Begin VB.TextBox txtMessage 
      Height          =   960
      Left            =   4065
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   480
      Width           =   3435
   End
   Begin VB.ListBox lstMessages 
      Height          =   4350
      Left            =   90
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   3855
   End
   Begin VB.Label Labelurl 
      Alignment       =   2  'Center
      Caption         =   "Click to copy MSDN URL into memory."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   4110
      TabIndex        =   16
      Tag             =   "http://search.microsoft.com/default.asp"
      ToolTipText     =   "Copies website address to clipboard"
      Top             =   4605
      Width           =   3360
   End
   Begin VB.Label Label1 
      Caption         =   "Hex Value"
      Height          =   225
      Index           =   3
      Left            =   5970
      TabIndex        =   14
      Top             =   1485
      Width           =   1380
   End
   Begin VB.Label Label1 
      Caption         =   "Long Value"
      Height          =   225
      Index           =   2
      Left            =   4275
      TabIndex        =   13
      Top             =   1485
      Width           =   1380
   End
   Begin VB.Label Label1 
      Caption         =   "Brief Description or Explanation"
      Height          =   225
      Index           =   1
      Left            =   4095
      TabIndex        =   12
      Top             =   255
      Width           =   3360
   End
   Begin VB.Label Label1 
      Caption         =   "Windows Messages"
      Height          =   225
      Index           =   0
      Left            =   135
      TabIndex        =   11
      Top             =   225
      Width           =   2355
   End
   Begin VB.Label Label2 
      Caption         =   $"frmWM.frx":0000
      ForeColor       =   &H00C00000&
      Height          =   1185
      Left            =   4110
      TabIndex        =   15
      Top             =   3330
      Width           =   3450
   End
End
Attribute VB_Name = "frmWM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
' =================================================================
' These are provided here for copying and pasting to other projects.

' Note the 1st three are unknown values.  I truly didn't spend too much time trying to
' find their values. But if it matters to you, I guess you'll need to find their true values

' =================================================================


Private Const WM_THEMECHANGED = 0   ' unknown XP only, gotta exist somewhere
Private Const WM_UNICHAR = 0                 ' unknown
Private Const MN_GETHMENU = 0               ' unknown W2K only, gotta exist somewhere

Private Const WM_USER = &H400
Private Const WM_APP = &H8000
Private Const WM_APPCOMMAND = &H319
Private Const WM_ASKCBFORMATNAME = &H30C
Private Const WM_CANCELJOURNAL = &H4B
Private Const WM_CANCELMODE = &H1F
Private Const WM_CAPTURECHANGED = &H215
Private Const WM_CHANGECBCHAIN = &H30D
Private Const WM_CHANGEUISTATE = &H127
Private Const WM_CHAR = &H102
Private Const WM_CHARTOITEM = &H2F
Private Const WM_CHILDACTIVATE = &H22
Private Const WM_CHOOSEFONT_GETLOGFONT = (WM_USER + 1)
Private Const WM_CHOOSEFONT_SETFLAGS = (WM_USER + 102)
Private Const WM_CHOOSEFONT_SETLOGFONT = (WM_USER + 101)
Private Const WM_CLEAR = &H303
Private Const WM_CLOSE = &H10
Private Const WM_COMMAND = &H111
Private Const WM_COMMNOTIFY = &H44
Private Const WM_COMPACTING = &H41
Private Const WM_COMPAREITEM = &H39
Private Const WM_CONTEXTMENU = &H7B
Private Const WM_CONVERTREQUEST = &H10A
Private Const WM_CONVERTREQUESTEX = &H108
Private Const WM_CONVERTRESULT = &H10B
Private Const WM_COPY = &H301
Private Const WM_COPYDATA = &H4A
Private Const WM_CPL_LAUNCH = (WM_USER + 1000)
Private Const WM_CPL_LAUNCHED = (WM_USER + 1001)
Private Const WM_CREATE = &H1
Private Const WM_CTLCOLOR = &H19
Private Const WM_CTLCOLORBTN = &H135
Private Const WM_CTLCOLORDLG = &H136
Private Const WM_CTLCOLOREDIT = &H133
Private Const WM_CTLCOLORLISTBOX = &H134
Private Const WM_CTLCOLORMSGBOX = &H132
Private Const WM_CTLCOLORSCROLLBAR = &H137
Private Const WM_CTLCOLORSTATIC = &H138
Private Const WM_CUT = &H300
Private Const WM_DDE_FIRST = &H3E0
Private Const WM_DDE_ACK = (WM_DDE_FIRST + 4)
Private Const WM_DDE_ADVISE = (WM_DDE_FIRST + 2)
Private Const WM_DDE_DATA = (WM_DDE_FIRST + 5)
Private Const WM_DDE_EXECUTE = (WM_DDE_FIRST + 8)
Private Const WM_DDE_INITIATE = (WM_DDE_FIRST)
Private Const WM_DDE_LAST = (WM_DDE_FIRST + 8)
Private Const WM_DDE_POKE = (WM_DDE_FIRST + 7)
Private Const WM_DDE_REQUEST = (WM_DDE_FIRST + 6)
Private Const WM_DDE_TERMINATE = (WM_DDE_FIRST + 1)
Private Const WM_DDE_UNADVISE = (WM_DDE_FIRST + 3)
Private Const WM_DEADCHAR = &H103
Private Const WM_DELETEITEM = &H2D
Private Const WM_DESTROY = &H2
Private Const WM_DESTROYCLIPBOARD = &H307
Private Const WM_DEVICECHANGE = &H219
Private Const WM_DEVMODECHANGE = &H1B
Private Const WM_DISPLAYCHANGE = &H7E
Private Const WM_DRAWCLIPBOARD = &H308
Private Const WM_DRAWITEM = &H2B
Private Const WM_DROPFILES = &H233
Private Const WM_ENABLE = &HA
Private Const WM_ENDSESSION = &H16
Private Const WM_ENTERIDLE = &H121
Private Const WM_ENTERMENULOOP = &H211
Private Const WM_ENTERSIZEMOVE = &H231
Private Const WM_ERASEBKGND = &H14
Private Const WM_EXITMENULOOP = &H212
Private Const WM_EXITSIZEMOVE = &H232
Private Const WM_FONTCHANGE = &H1D
Private Const WM_FORWARDMSG = &H37F
Private Const WM_GETDLGCODE = &H87
Private Const WM_GETFONT = &H31
Private Const WM_GETHOTKEY = &H33
Private Const WM_GETICON = &H7F
Private Const WM_GETMINMAXINFO = &H24
Private Const WM_GETOBJECT = &H3D
Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE
Private Const WM_HANDHELDFIRST = &H358
Private Const WM_HANDHELDLAST = &H35F
Private Const WM_HELP = &H53
Private Const WM_HOTKEY = &H312
Private Const WM_HSCROLL = &H114
Private Const WM_HSCROLLCLIPBOARD = &H30E
Private Const WM_ICONERASEBKGND = &H27
Private Const WM_IME_CHAR = &H286
Private Const WM_IME_COMPOSITION = &H10F
Private Const WM_IME_COMPOSITIONFULL = &H284
Private Const WM_IME_CONTROL = &H283
Private Const WM_IME_ENDCOMPOSITION = &H10E
Private Const WM_IME_KEYDOWN = &H290
Private Const WM_IME_KEYLAST = &H10F
Private Const WM_IME_KEYUP = &H291
Private Const WM_IME_NOTIFY = &H282
Private Const WM_IME_REPORT = &H280
Private Const WM_IME_REQUEST = &H288
Private Const WM_IME_SELECT = &H285
Private Const WM_IME_SETCONTEXT = &H281
Private Const WM_IME_STARTCOMPOSITION = &H10D
Private Const WM_IMEKEYDOWN = &H290
Private Const WM_IMEKEYUP = &H291
Private Const WM_INITDIALOG = &H110
Private Const WM_INITMENU = &H116
Private Const WM_INITMENUPOPUP = &H117
Private Const WM_INPUTLANGCHANGE = &H51
Private Const WM_INPUTLANGCHANGEREQUEST = &H50
Private Const WM_INTERIM = &H10C
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYFIRST = &H100
Private Const WM_KEYLAST = &H108
Private Const WM_KEYUP = &H101
Private Const WM_KILLFOCUS = &H8
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_MBUTTONDBLCLK = &H209
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208
Private Const WM_MDIACTIVATE = &H222
Private Const WM_MDICASCADE = &H227
Private Const WM_MDICREATE = &H220
Private Const WM_MDIDESTROY = &H221
Private Const WM_MDIGETACTIVE = &H229
Private Const WM_MDIICONARRANGE = &H228
Private Const WM_MDIMAXIMIZE = &H225
Private Const WM_MDINEXT = &H224
Private Const WM_MDIREFRESHMENU = &H234
Private Const WM_MDIRESTORE = &H223
Private Const WM_MDISETMENU = &H230
Private Const WM_MDITILE = &H226
Private Const WM_MEASUREITEM = &H2C
Private Const WM_MENUCHAR = &H120
Private Const WM_MENUCOMMAND = &H126
Private Const WM_MENUDRAG = &H123
Private Const WM_MENUGETOBJECT = &H124
Private Const WM_MENURBUTTONUP = &H122
Private Const WM_MENUSELECT = &H11F
Private Const WM_MOUSEACTIVATE = &H21
Private Const WM_MOUSEFIRST = &H200
Private Const WM_MOUSEHOVER = &H2A1
Private Const WM_MOUSELAST = &H209
Private Const WM_MOUSELEAVE = &H2A3
Private Const WM_MOUSEMOVE = &H200
Private Const WM_MOUSEWHEEL = &H20A
Private Const WM_MOVE = &H3
Private Const WM_MOVING = &H216
Private Const WM_NCACTIVATE = &H86
Private Const WM_NCCALCSIZE = &H83
Private Const WM_NCCREATE = &H81
Private Const WM_NCDESTROY = &H82
Private Const WM_NCHITTEST = &H84
Private Const WM_NCLBUTTONDBLCLK = &HA3
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const WM_NCLBUTTONUP = &HA2
Private Const WM_NCMBUTTONDBLCLK = &HA9
Private Const WM_NCMBUTTONDOWN = &HA7
Private Const WM_NCMBUTTONUP = &HA8
Private Const WM_NCMOUSEHOVER = &H2A0
Private Const WM_NCMOUSELEAVE = &H2A2
Private Const WM_NCMOUSEMOVE = &HA0
Private Const WM_NCPAINT = &H85
Private Const WM_NCRBUTTONDBLCLK = &HA6
Private Const WM_NCRBUTTONDOWN = &HA4
Private Const WM_NCRBUTTONUP = &HA5
Private Const WM_NCXBUTTONDBLCLK = &HAD
Private Const WM_NCXBUTTONDOWN = &HAB
Private Const WM_NCXBUTTONUP = &HAC
Private Const WM_NEXTDLGCTL = &H28
Private Const WM_NEXTMENU = &H213
Private Const WM_NOTIFY = &H4E
Private Const WM_NOTIFYFORMAT = &H55
Private Const WM_NULL = &H0
Private Const WM_OTHERWINDOWCREATED = &H42
Private Const WM_OTHERWINDOWDESTROYED = &H43
Private Const WM_PAINT = &HF
Private Const WM_PAINTCLIPBOARD = &H309
Private Const WM_PAINTICON = &H26
Private Const WM_PALETTECHANGED = &H311
Private Const WM_PALETTEISCHANGING = &H310
Private Const WM_PARENTNOTIFY = &H210
Private Const WM_PASTE = &H302
Private Const WM_PENWINFIRST = &H380
Private Const WM_PENWINLAST = &H38F
Private Const WM_POWER = &H48
Private Const WM_POWERBROADCAST = &H218
Private Const WM_PRINT = &H317
Private Const WM_PRINTCLIENT = &H318
Private Const WM_PSD_ENVSTAMPRECT = (WM_USER + 5)
Private Const WM_PSD_FULLPAGERECT = (WM_USER + 1)
Private Const WM_PSD_GREEKTEXTRECT = (WM_USER + 4)
Private Const WM_PSD_MARGINRECT = (WM_USER + 3)
Private Const WM_PSD_MINMARGINRECT = (WM_USER + 2)
Private Const WM_PSD_PAGESETUPDLG = 0   '  unknown; sould be something like (WM_USER + ?)
Private Const WM_PSD_YAFULLPAGERECT = (WM_USER + 6)
Private Const WM_QUERYDRAGICON = &H37
Private Const WM_QUERYENDSESSION = &H11
Private Const WM_QUERYNEWPALETTE = &H30F
Private Const WM_QUERYOPEN = &H13
Private Const WM_QUERYUISTATE = &H129
Private Const WM_QUEUESYNC = &H23
Private Const WM_QUIT = &H12
Private Const WM_RASDIALEVENT = &HCCCD
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RENDERALLFORMATS = &H306
Private Const WM_RENDERFORMAT = &H305
Private Const WM_SETCURSOR = &H20
Private Const WM_SETFOCUS = &H7
Private Const WM_SETFONT = &H30
Private Const WM_SETHOTKEY = &H32
Private Const WM_SETICON = &H80
Private Const WM_SETREDRAW = &HB
Private Const WM_SETTEXT = &HC
Private Const WM_WININICHANGE = &H1A
Private Const WM_SETTINGCHANGE = WM_WININICHANGE
Private Const WM_SHOWWINDOW = &H18
Private Const WM_SIZE = &H5
Private Const WM_SIZECLIPBOARD = &H30B
Private Const WM_SIZING = &H214
Private Const WM_SPOOLERSTATUS = &H2A
Private Const WM_STYLECHANGED = &H7D
Private Const WM_STYLECHANGING = &H7C
Private Const WM_SYNCPAINT = &H88
Private Const WM_SYSCHAR = &H106
Private Const WM_SYSCOLORCHANGE = &H15
Private Const WM_SYSCOMMAND = &H112
Private Const WM_SYSDEADCHAR = &H107
Private Const WM_SYSKEYDOWN = &H104
Private Const WM_SYSKEYUP = &H105
Private Const WM_TCARD = &H52
Private Const WM_TIMECHANGE = &H1E
Private Const WM_TIMER = &H113
Private Const WM_UNDO = &H304
Private Const WM_UNINITMENUPOPUP = &H125
Private Const WM_UPDATEUISTATE = &H128
Private Const WM_USERCHANGED = &H54
Private Const WM_VKEYTOITEM = &H2E
Private Const WM_VSCROLL = &H115
Private Const WM_VSCROLLCLIPBOARD = &H30A
Private Const WM_WINDOWPOSCHANGED = &H47
Private Const WM_WINDOWPOSCHANGING = &H46
Private Const WM_WNT_CONVERTREQUESTEX = &H109
Private Const WM_XBUTTONDBLCLK = &H20D
Private Const WM_XBUTTONDOWN = &H20B
Private Const WM_XBUTTONUP = &H20C

' =================================================================
' If you want to add more, simply add a PRIVATE CONST statement above,
'  add the description below, and finally add a line entry to the LoadListBox function below

' P.S. after some of the descriptions are additional nice-to-know remarks
' =================================================================

Private Function GetDescription(Index As Long, bSearching As Boolean) As String
On Error Resume Next
Dim sDescrip As String, lValue As Long
sDescrip = "Description not provided. May no longer be supported."
Select Case Index
Case WM_APPCOMMAND: sDescrip = "Application started" ' wparam is the window
Case WM_ASKCBFORMATNAME: sDescrip = "Requesting name of clipboard format"
Case WM_CANCELJOURNAL: sDescrip = "journaling activities cancelled"
Case WM_CANCELMODE: sDescrip = "cancel event received"
Case WM_CAPTURECHANGED: sDescrip = "lost mouse capture" ' lparam is window taking capture
Case WM_CHANGECBCHAIN: sDescrip = "window removed from clipboard viewer chaing"  ' wParam is window removed
Case WM_CHANGEUISTATE: sDescrip = "user interface changed"
Case WM_CHAR: sDescrip = "key pressed" ' wparam is ASCII value of key
Case WM_CHARTOITEM: sDescrip = "listbox key stroke"
Case WM_CHILDACTIVATE: sDescrip = "child window activated"
Case WM_CHOOSEFONT_GETLOGFONT: sDescrip = "font selected from dialog" ' lparam is LOGFONT
Case WM_CHOOSEFONT_SETFLAGS: sDescrip = "font dialog options seclected" ' lparam is CHOOSEFONT
Case WM_CHOOSEFONT_SETLOGFONT: sDescrip = "font dialog properites set" ' lparam is LOGFONT
Case WM_CLEAR: sDescrip = "text, combo or list box cleared"
Case WM_CLOSE: sDescrip = "window close requested"
Case WM_COMMAND: sDescrip = "menu or hotkey selected; or control message recieved" ' wparam HiWord (0=menu, 1=hotkey), lparam (hWnd of control if control sent message, else Null)
Case WM_COMPACTING: sDescrip = "memory compacting, sys resources low"
Case WM_COMPAREITEM: sDescrip = "location requested for new list/combo box item (sorted lists only)"
Case WM_CONTEXTMENU: sDescrip = "mouse right button clicked" ' wparam is hWnd, lParam mouse coords
Case WM_COPY: sDescrip = "text or combo box contents copied" ' note: copied to clipboard in CF_TEXT format
Case WM_COPYDATA: sDescrip = "copy data requested" ' wParam window passing data, lParam is COPYDATASTRUCT
Case WM_CREATE: sDescrip = "Window about to be created"  'lParam is CREATESTRUCT
Case WM_CTLCOLOR: sDescrip = "default color scheme changing" ' note: 16-bit Win versions only
' the following are same as above but for 32-bit Win versions
Case WM_CTLCOLORBTN: sDescrip = "custom button about to be redrawn, options requestd (only for 32-bit Win versions)"
Case WM_CTLCOLORDLG: sDescrip = "dialog box about to be redrawn, options requested (only for 32-bit Win versions)"""
Case WM_CTLCOLOREDIT: sDescrip = "edit control about to be redrawn, options requested (only for 32-bit Win versions)"""
Case WM_CTLCOLORLISTBOX: sDescrip = "listbox about to be repainted (only for 32-bit Win versions)""  ' wparm is listbox hdc, lparam is listbox handle"
Case WM_CTLCOLORSCROLLBAR: sDescrip = "scrollbar about to be repainted, options requested (only for 32-bit Win versions)"""
Case WM_CTLCOLORSTATIC: sDescrip = "static control about to be redrawn, options requested (only for 32-bit Win versions)"""
'
Case WM_CUT: sDescrip = "text or combo box text cut/deleted"  ' contents sent to clipboard in CF_TEXT format
Case WM_DDE_ACK: sDescrip = "DDE acknowledged"  ' note see MSDN for proper use of these DDE messages
Case WM_DDE_ADVISE: sDescrip = "request to DDE server for automatic updates"
Case WM_DDE_DATA: sDescrip = "data-ready notification to DDE client"
Case WM_DDE_EXECUTE: sDescrip = "DDE execution sent"
Case WM_DDE_INITIATE: sDescrip = "DDE conversation initiated"
Case WM_DDE_POKE: sDescrip = "Client requests DDE server to accept data"
Case WM_DDE_REQUEST: sDescrip = "Client requesting data from DDE server"
Case WM_DDE_TERMINATE: sDescrip = "DDE conversation terminated"
Case WM_DDE_UNADVISE: sDescrip = "stop DDE automatic notification of updates"
Case WM_DEADCHAR: sDescrip = "keyboard dead letter posted"
Case WM_DELETEITEM: sDescrip = "listbox is deleting item"
Case WM_DESTROY: sDescrip = "Window is being closed/destroyed" ' unload all memory objects now!
Case WM_DESTROYCLIPBOARD: sDescrip = "clipboard emptied"
Case WM_DEVICECHANGE: sDescrip = "hardware configuartion changed"
Case WM_DEVMODECHANGE: sDescrip = "system device context changed"
Case WM_DISPLAYCHANGE: sDescrip = "system display resolution changed"
Case WM_DRAWCLIPBOARD: sDescrip = "clipboard contents changed"
Case WM_DRAWITEM: sDescrip = "redraw owner-drawn button, menu, combo/list box"  ' lParam is DRAWITEMSTRUCT
Case WM_DROPFILES: sDescrip = "files dropped on window"
Case WM_ENABLE: sDescrip = "Window is being enabled/disabled" ' wPram is True if Enabled
Case WM_ENDSESSION: sDescrip = "system is logging off or shutting down"
Case WM_ENTERIDLE: sDescrip = "menu/dialog box in idle state"
Case WM_ENTERMENULOOP: sDescrip = "modal menu loop started" ' wParam True if TrackPopupMenu is being used
Case WM_ENTERSIZEMOVE: sDescrip = "window is beginning resizing"
Case WM_ERASEBKGND: sDescrip = "window background erased"
Case WM_EXITMENULOOP: sDescrip = "modal menu loop ended" ' wParam True if a shortcut menu
Case WM_EXITSIZEMOVE: sDescrip = "window is done resizing"
Case WM_FONTCHANGE: sDescrip = "font added/removed from system"
Case WM_GETDLGCODE: sDescrip = "control input redirection offered" ' lparam is MSG
Case WM_GETFONT: sDescrip = "font requested"
Case WM_GETHOTKEY: sDescrip = "window hotkey requested"
Case WM_GETICON: sDescrip = "application icon requested"  ' wParam is icon size requested
Case WM_GETMINMAXINFO: sDescrip = "min/max windows size requested" ' lparam is MINMAXINFO
Case WM_GETOBJECT: sDescrip = "object information requested"  ' see MSDN, good size topic
Case WM_GETTEXT: sDescrip = "window text requested"
Case WM_GETTEXTLENGTH: sDescrip = "window text length requested"
Case WM_HELP: sDescrip = "help requested"  ' lparam is HELPINFO
Case WM_HOTKEY: sDescrip = "window hotkey pressed"
Case WM_HSCROLL: sDescrip = "horizontal scroll bar clicked"  ' lParam hWnd of scrollbar if bar sent message, else NULL
Case WM_HSCROLLCLIPBOARD: sDescrip = "clipboard viewer being scrolled"  ' only when clipboard contains CF_OWNERDISPLAY format
Case WM_ICONERASEBKGND: sDescrip = "minimized window icon erased"  ' WinNT3.51 & earlier
' WM_IME... messages are international in nature and can be viewed at MSDN for there
Case WM_IME_CHAR, WM_IME_COMPOSITION, WM_IME_COMPOSITIONFULL, WM_IME_CONTROL, WM_IME_ENDCOMPOSITION, _
         WM_IME_KEYDOWN, WM_IME_KEYLAST, WM_IME_KEYUP, WM_IME_NOTIFY, WM_IME_REPORT, WM_IME_REQUEST, _
         WM_IME_SELECT, WM_IME_SETCONTEXT, WM_IME_STARTCOMPOSITION, WM_IMEKEYDOWN, WM_IMEKEYUP
    sDescrip = "messages are international in nature and can be viewed at MSDN for there"
Case WM_INITDIALOG: sDescrip = "dialog box about to be displayed"  ' wParam hWnd, lParam dialog box structure(s)
Case WM_INITMENU: sDescrip = "menu becoming active"  ' wparam is handle to menu
Case WM_INITMENUPOPUP: sDescrip = "submenu/dropdown menu becoming active"  ' wParam is menu handle
Case WM_INPUTLANGCHANGE: sDescrip = "input language changed"
Case WM_INPUTLANGCHANGEREQUEST: sDescrip = "input language change is requested"
Case WM_KEYDOWN: sDescrip = "non-ALT key pressed"
Case WM_KEYUP: sDescrip = "non-ALT key released"
Case WM_KILLFOCUS: sDescrip = "window losing keyboard focus"
Case WM_LBUTTONDBLCLK: sDescrip = "left mouse button double clicked"  ' lparam is cursor location
Case WM_LBUTTONDOWN: sDescrip = "left mouse button down"  ' lparam is cursor location
Case WM_LBUTTONUP: sDescrip = "left mouse button released"  ' lparam is cursor location
Case WM_MBUTTONDBLCLK: sDescrip = "middle mouse button double clicked"  ' lparam is cursor location
Case WM_MBUTTONDOWN: sDescrip = "middle mouse button down"  ' lparam is cursor location
Case WM_MBUTTONUP: sDescrip = "middle mouse button released" ' ' lparam is cursor location
Case WM_MDIACTIVATE: sDescrip = "child window activated" ' wparam MDI hWnd
Case WM_MDICASCADE: sDescrip = "child windows cascaded"
Case WM_MDICREATE: sDescrip = "child window created" ' lparam is MDICREATESTRUCT
Case WM_MDIDESTROY: sDescrip = "child window closed" ' wparam MDI hWnd
Case WM_MDIGETACTIVE: sDescrip = "active child window requested"
Case WM_MDIICONARRANGE: sDescrip = "child minimized windows arranged"
Case WM_MDIMAXIMIZE: sDescrip = "child window maximized" ' wparam MDI hWnd
Case WM_MDINEXT: sDescrip = "next/previous child window activated" ' wparam MDI hWnd
Case WM_MDIREFRESHMENU: sDescrip = "MDI frame window's menu refreshed"
Case WM_MDIRESTORE: sDescrip = "child window size restored" ' wparam MDI hWnd
Case WM_MDISETMENU: sDescrip = "child window's menu replaced"
Case WM_MDITILE: sDescrip = "child windows tiled"
Case WM_MEASUREITEM: sDescrip = "New owner-drawn menu, button, combo/list box size requested"  ' lparam is MEASUREITEMSTRUCT
Case WM_MENUCHAR: sDescrip = "non-menu hotkey pressed"  ' lparam is menu handle
Case WM_MENUCOMMAND: sDescrip = "menu item selected" ' wParam (98/ME menu index & ID, else Index only), lParam is menu item handle
Case WM_MENUDRAG: sDescrip = "menu being dragged/dropped" ' wParam is start of drag, lParam is menu handle
Case WM_MENUGETOBJECT: sDescrip = "get drag/drop menu item" ' lParam is a MENUGETOBJECTINFO
Case WM_MENURBUTTONUP: sDescrip = "right mouse button up on menu" ' wParam is menu item position, lParam is menu handle
Case WM_MENUSELECT: sDescrip = "menu item selected"  ' lparam is menu handle
Case WM_MOUSEACTIVATE: sDescrip = "mouse activating a window" ' wparam is parent window being activated
Case WM_MOUSEHOVER: sDescrip = "mouse is hovering"  ' lparam is cursor location
Case WM_MOUSELEAVE: sDescrip = "mouse is leaving client area"
Case WM_MOUSEMOVE: sDescrip = "mouse is moving"  ' lparam is cursor location
Case WM_MOUSEWHEEL: sDescrip = "mouse wheel rotated"   ' lparam is cursor location
Case WM_MOVE: sDescrip = "window was moved"  ' lParam is x,y coords of new position
Case WM_MOVING: sDescrip = "window is moving"  ' lparam is RECT of current position
Case WM_NCACTIVATE: sDescrip = "redrawing title bar to active/inactive"  ' lparam True if Active
Case WM_NCCALCSIZE: sDescrip = "client are measurement requested"
Case WM_NCCREATE: sDescrip = "window about to be created" ' lParam is CREATESTRUCT (return False to prevent new window from being created at all)
Case WM_NCDESTROY: sDescrip = "window's nonclient area being destroyed"
Case WM_NCHITTEST: sDescrip = "cursor/mouse moved"  ' lparam is cursor location
Case WM_NCLBUTTONDBLCLK: sDescrip = "left mouse button double clicked (nonclient area)"  ' lparam is cursor location
Case WM_NCLBUTTONDOWN: sDescrip = "left mouse button down (nonclient area)"  ' lparam is cursor location
Case WM_NCLBUTTONUP: sDescrip = "left mouse button released (nonclient area)"  ' lparam is cursor location
Case WM_NCMBUTTONDBLCLK: sDescrip = "middle mouse button double clicked (nonclient area)"  ' lparam is cursor location
Case WM_NCMBUTTONDOWN: sDescrip = "middle mouse button down (nonclient area)"  ' lparam is cursor location
Case WM_NCMBUTTONUP: sDescrip = "middle mouse button released (nonclient area)"  ' lparam is cursor location
Case WM_NCMOUSEHOVER: sDescrip = "mouse hovering (nonclient area)"  ' lparam is cursor location
Case WM_NCMOUSELEAVE: sDescrip = "mouse leaving area (nonclient area)"
Case WM_NCMOUSEMOVE: sDescrip = "mouse moving (nonclient area)"  ' lparam is cursor location
Case WM_NCPAINT: sDescrip = "windows nonclient area must be repainted"
Case WM_NCRBUTTONDBLCLK: sDescrip = "mouse right button double clicked (nonclient area)"  ' lparam is cursor location
Case WM_NCRBUTTONDOWN: sDescrip = "mouse right button down (nonclient area)"  ' lparam is cursor location
Case WM_NCRBUTTONUP: sDescrip = "mouse right button released (nonclient area)"  ' lparam is cursor location
Case WM_NCXBUTTONDBLCLK: sDescrip = "mouse X-button double clicked (nonclient area)"  ' lparam is cursor location
Case WM_NCXBUTTONDOWN: sDescrip = "mouse X-button down (nonclient area)"  ' lparam is cursor location
Case WM_NCXBUTTONUP: sDescrip = "mouse X-button released (nonclient area)"  ' lparam is cursor location
Case WM_NEXTDLGCTL: sDescrip = "dialog box keyboard focus changed"  ' used to tab in dialog boxes
Case WM_NEXTMENU: sDescrip = "arrow selecting next menu item" ' wParam is ASCII key code, lpMdiNextMenu is MDINEXTMENU
Case WM_NOTIFY: sDescrip = "control notifying parent"  ' see MSDN on handling this message
Case WM_NOTIFYFORMAT: sDescrip = "ANSI/Unicode structures allowed when notifying" ' See MSDN
Case WM_NULL: sDescrip = "no message"
Case WM_PAINT: sDescrip = "window repaint is requested"
Case WM_PAINTCLIPBOARD: sDescrip = "clipboard viewer needs repainting"    ' only when clipboard contains CF_OWNERDISPLAY format
Case WM_PAINTICON: sDescrip = "minimized window icon being repainted"  ' WinNT3.51 & earlier
Case WM_PALETTECHANGED: sDescrip = "system palette changed"  ' wParam hWnd of window causing change
Case WM_PALETTEISCHANGING: sDescrip = "system palette is changing" ' wParam hWnd of window causing change
Case WM_PARENTNOTIFY: sDescrip = "child window has mouse click or is being created/destroyed"  '
Case WM_PASTE: sDescrip = "text or combo box text pasted"  ' note: only in CF_TEXT format
Case WM_POWER: sDescrip = "system about to enter suspended mode"
Case WM_POWERBROADCAST: sDescrip = "power management event posted"
Case WM_PRINT: sDescrip = "window requested to draw in device context"  ' wparam is hDC
Case WM_PRINTCLIENT: sDescrip = "window requested to draw client area in device context"  ' wParam is hDC
Case WM_PSD_ENVSTAMPRECT: sDescrip = "page-setup dialog box about to display envelope stamp"
Case WM_PSD_FULLPAGERECT: sDescrip = "page-setup dialog box sending sample page coordinates"
Case WM_PSD_GREEKTEXTRECT: sDescrip = "page-setup dialog box about to draw sample page Greek text"
Case WM_PSD_MARGINRECT: sDescrip = "page-setup dialog box about to draw sample page margin rectangle"
Case WM_PSD_MINMARGINRECT: sDescrip = "page-setup dialog box sending margin coordinates of sample page"
Case WM_PSD_PAGESETUPDLG: sDescrip = "page-setup dialog box about to draw contents of sample page"
Case WM_PSD_YAFULLPAGERECT: sDescrip = "page-setup dialog box about to draw envelope return address portion"
Case WM_QUERYDRAGICON: sDescrip = "minimized window being dragged, drag icon requested"
Case WM_QUERYENDSESSION: sDescrip = "log off or shutdown requested"  ' any app can reject request
Case WM_QUERYNEWPALETTE: sDescrip = "window about to receive focus"
Case WM_QUERYOPEN: sDescrip = "window is being restored to size"
Case WM_QUERYUISTATE: sDescrip = "window user interface state requested"
Case WM_QUEUESYNC: sDescrip = "CBT application syncing messages"
Case WM_QUIT: sDescrip = "application termination requested"  ' Note: message only sent w/PostQuitMessage and recieved via Get/PeekMessage APIs not CallWindowProc
Case WM_RASDIALEVENT: sDescrip = "Remote Access Service event posted"
Case WM_RBUTTONDBLCLK: sDescrip = "mouse right button double clicked" ' lparam is location
Case WM_RBUTTONDOWN: sDescrip = "mouse right button down" ' lparam is location
Case WM_RBUTTONUP: sDescrip = "mouse right button released"  ' lparam is location
Case WM_RENDERALLFORMATS: sDescrip = "clipboard about to be destroyed"
Case WM_RENDERFORMAT: sDescrip = "clipboard requesting data rendering"
Case WM_SETCURSOR: sDescrip = "mouse moving in window unknowingly" ' wparam is hWnd & lParam is x,y coords
Case WM_SETFOCUS: sDescrip = "window gained keyboard focus"
Case WM_SETFONT: sDescrip = "font for control set"
Case WM_SETHOTKEY: sDescrip = "window hotkey set"
Case WM_SETICON: sDescrip = "window icon set/removed" ' wParam is size of icon, lParam is icon handle
Case WM_SETREDRAW: sDescrip = "window changes can/cannot be drawn" ' wParam True if redrawing allowed
Case WM_SETTEXT: sDescrip = "window text set"
Case WM_SETTINGCHANGE: sDescrip = WM_WININICHANGE
Case WM_SHOWWINDOW: sDescrip = "window about to be hidden/shown"  ' wParam True if Shown
Case WM_SIZE: sDescrip = "Window size changed"
Case WM_SIZECLIPBOARD: sDescrip = "clipboard resized"  ' only when clipboard contains CF_OWNERDISPLAY format
Case WM_SIZING: sDescrip = "window is resizing"
Case WM_SPOOLERSTATUS: sDescrip = "Print Manager job added/removed"
Case WM_STYLECHANGED: sDescrip = "Window style changed" ' wparam is style type, lParam is STYLESTRUCT
Case WM_STYLECHANGING: sDescrip = "window style about to change"  ' wparam is style type, lParam is STYLESTRUCT
Case WM_SYNCPAINT: sDescrip = "window synchronizing painting"
Case WM_SYSCHAR: sDescrip = "ALT-key character pressed"  ' wparam is ASCII key
Case WM_SYSCOLORCHANGE: sDescrip = "system color setting changed"
Case WM_SYSCOMMAND: sDescrip = "system menu item selected"
Case WM_SYSDEADCHAR: sDescrip = "system dead letter from keyboard"
Case WM_SYSKEYDOWN: sDescrip = "system key pressed (ALT, F10, etc)"
Case WM_SYSKEYUP: sDescrip = "system key released"
Case WM_TCARD: sDescrip = "training card action (help)"
Case WM_THEMECHANGED: sDescrip = "desktop them changed (WinXP only)"
Case WM_TIMECHANGE: sDescrip = "system time is being changed"
Case WM_TIMER: sDescrip = "timer has expired"  ' must have already made a SetTimer call
Case WM_UNDO: sDescrip = "edit control last action is undone"
Case WM_UNICHAR: sDescrip = "key pressed" ' same as WM_CHAR but unicode
Case WM_UNINITMENUPOPUP: sDescrip = "submenu closed"  ' wparam is menu handle
Case WM_UPDATEUISTATE: sDescrip = "window's user interface state changed"
Case WM_USERCHANGED: sDescrip = "new user logging on/off"
Case WM_VKEYTOITEM: sDescrip = "listbox keystroke"  ' wparam is ASCII value of key
Case WM_VSCROLL: sDescrip = "vertical scroll bar clicked"  ' lParam hWnd of scrollbar if bar sent message, else NULL
Case WM_VSCROLLCLIPBOARD: sDescrip = "clipboard vertical scroll clicked"   ' only when clipboard contains CF_OWNERDISPLAY format
Case WM_WINDOWPOSCHANGED: sDescrip = "window size, position or z-order changed"  ' lParam is WINDOWPOS
Case WM_WINDOWPOSCHANGING: sDescrip = "window size, position or z-order is changing"   ' lParam is WINDOWPOS
Case WM_WININICHANGE: sDescrip = "WIN.INI file changed"  ' lparam name of sys parameter that changed
Case WM_XBUTTONDBLCLK: sDescrip = "mouse X-button double clicked" ' lparam is location
Case WM_XBUTTONDOWN: sDescrip = "mouse X-button down"  ' lparam is location
Case WM_XBUTTONUP: sDescrip = "mouse X-button released" ' lparam is location
Case MN_GETHMENU: sDescrip = "window's hMenu is being requested (Win2000 only)"
Case WM_APP, WM_USER: sDescrip = "used in conjunction with another number to uniquely identify a message"
End Select
If bSearching Then
    GetDescription = sDescrip
Else
    lValue = Index
    txtMessage = sDescrip
    If lValue <> 0 Or lstMessages.Text = "WM_NULL" Then
        txtValue(0) = lValue
        txtValue(1) = "&H" & Hex(lValue)
        txtDeclare = "Private Const " & lstMessages.Text & " = "
        If (lValue And WM_USER) = WM_USER Then
            ' here we filter out any messages that are formatted out as WM_USER + ## when they shouldn't be
            Select Case lValue
            Case WM_USER, WM_RASDIALEVENT   ' only value above that formats incorrectly as (WM_USER + ###)
                ' but if you add your own above & they format wrong, add their values with WM_RASDIALEVENT
                txtDeclare = txtDeclare & txtValue(1)
            Case Else
                txtDeclare = txtDeclare & "( WM_USER "
                If lValue - WM_USER >= 0 Then txtDeclare = txtDeclare & "+ "
                txtDeclare = txtDeclare & (lValue - WM_USER) & ")"
            End Select
        Else
            txtDeclare = txtDeclare & txtValue(1)
        End If
    Else
        txtValue(0) = "Unknown."
        txtValue(1) = "Unknown?"
        txtDeclare = ""
    End If
End If
End Function

Private Sub LoadListBox()

' =================================================================
' Know this is not the best formatting.  But I copied all of this stuff from a text file I created
' and didn't feel like separating them into individual lines.
' =================================================================

With lstMessages
    .AddItem "WM_APP": .ItemData(.NewIndex) = &H8000
    .AddItem "WM_APPCOMMAND": .ItemData(.NewIndex) = &H319
    .AddItem "WM_ASKCBFORMATNAME": .ItemData(.NewIndex) = &H30C
    .AddItem "WM_CANCELJOURNAL": .ItemData(.NewIndex) = &H4B
    .AddItem "WM_CANCELMODE": .ItemData(.NewIndex) = &H1F
    .AddItem "WM_CAPTURECHANGED": .ItemData(.NewIndex) = &H215
    .AddItem "WM_CHANGECBCHAIN": .ItemData(.NewIndex) = &H30D
    .AddItem "WM_CHANGEUISTATE": .ItemData(.NewIndex) = &H127
    .AddItem "WM_CHAR": .ItemData(.NewIndex) = &H102
    .AddItem "WM_CHARTOITEM": .ItemData(.NewIndex) = &H2F
    .AddItem "WM_CHILDACTIVATE": .ItemData(.NewIndex) = &H22
    .AddItem "WM_CHOOSEFONT_GETLOGFONT": .ItemData(.NewIndex) = (WM_USER + 1)
    .AddItem "WM_CHOOSEFONT_SETFLAGS": .ItemData(.NewIndex) = (WM_USER + 102)
    .AddItem "WM_CHOOSEFONT_SETLOGFONT": .ItemData(.NewIndex) = (WM_USER + 101)
    .AddItem "WM_CLEAR": .ItemData(.NewIndex) = &H303
    .AddItem "WM_CLOSE": .ItemData(.NewIndex) = &H10
    .AddItem "WM_COMMAND": .ItemData(.NewIndex) = &H111
    .AddItem "WM_COMMNOTIFY": .ItemData(.NewIndex) = &H44
    .AddItem "WM_COMPACTING": .ItemData(.NewIndex) = &H41
    .AddItem "WM_COMPAREITEM": .ItemData(.NewIndex) = &H39
    .AddItem "WM_CONTEXTMENU": .ItemData(.NewIndex) = &H7B
    .AddItem "WM_CONVERTREQUEST": .ItemData(.NewIndex) = &H10A
    .AddItem "WM_CONVERTREQUESTEX": .ItemData(.NewIndex) = &H108
    .AddItem "WM_CONVERTRESULT": .ItemData(.NewIndex) = &H10B
    .AddItem "WM_COPY": .ItemData(.NewIndex) = &H301
    .AddItem "WM_COPYDATA": .ItemData(.NewIndex) = &H4A
    .AddItem "WM_CPL_LAUNCH": .ItemData(.NewIndex) = (WM_USER + 1000)
    .AddItem "WM_CPL_LAUNCHED": .ItemData(.NewIndex) = (WM_USER + 1001)
    .AddItem "WM_CREATE": .ItemData(.NewIndex) = &H1
    .AddItem "WM_CTLCOLOR": .ItemData(.NewIndex) = &H19
    .AddItem "WM_CTLCOLORBTN": .ItemData(.NewIndex) = &H135
    .AddItem "WM_CTLCOLORDLG": .ItemData(.NewIndex) = &H136
    .AddItem "WM_CTLCOLOREDIT": .ItemData(.NewIndex) = &H133
    .AddItem "WM_CTLCOLORLISTBOX": .ItemData(.NewIndex) = &H134
    .AddItem "WM_CTLCOLORMSGBOX": .ItemData(.NewIndex) = &H132
    .AddItem "WM_CTLCOLORSCROLLBAR": .ItemData(.NewIndex) = &H137
    .AddItem "WM_CTLCOLORSTATIC": .ItemData(.NewIndex) = &H138
    .AddItem "WM_CUT": .ItemData(.NewIndex) = &H300
    .AddItem "WM_DDE_ACK": .ItemData(.NewIndex) = (WM_DDE_FIRST + 4)
    .AddItem "WM_DDE_ADVISE": .ItemData(.NewIndex) = (WM_DDE_FIRST + 2)
    .AddItem "WM_DDE_DATA": .ItemData(.NewIndex) = (WM_DDE_FIRST + 5)
    .AddItem "WM_DDE_EXECUTE": .ItemData(.NewIndex) = (WM_DDE_FIRST + 8)
    .AddItem "WM_DDE_FIRST": .ItemData(.NewIndex) = &H3E0
    .AddItem "WM_DDE_INITIATE": .ItemData(.NewIndex) = (WM_DDE_FIRST)
    .AddItem "WM_DDE_LAST": .ItemData(.NewIndex) = (WM_DDE_FIRST + 8)
    .AddItem "WM_DDE_POKE": .ItemData(.NewIndex) = (WM_DDE_FIRST + 7)
    .AddItem "WM_DDE_REQUEST": .ItemData(.NewIndex) = (WM_DDE_FIRST + 6)
    .AddItem "WM_DDE_TERMINATE": .ItemData(.NewIndex) = (WM_DDE_FIRST + 1)
    .AddItem "WM_DDE_UNADVISE": .ItemData(.NewIndex) = (WM_DDE_FIRST + 3)
    .AddItem "WM_DEADCHAR": .ItemData(.NewIndex) = &H103
    .AddItem "WM_DELETEITEM": .ItemData(.NewIndex) = &H2D
    .AddItem "WM_DESTROY": .ItemData(.NewIndex) = &H2
    .AddItem "WM_DESTROYCLIPBOARD": .ItemData(.NewIndex) = &H307
    .AddItem "WM_DEVICECHANGE": .ItemData(.NewIndex) = &H219
    .AddItem "WM_DEVMODECHANGE": .ItemData(.NewIndex) = &H1B
    .AddItem "WM_DISPLAYCHANGE": .ItemData(.NewIndex) = &H7E
    .AddItem "WM_DRAWCLIPBOARD": .ItemData(.NewIndex) = &H308
    .AddItem "WM_DRAWITEM": .ItemData(.NewIndex) = &H2B
    .AddItem "WM_DROPFILES": .ItemData(.NewIndex) = &H233
    .AddItem "WM_ENABLE": .ItemData(.NewIndex) = &HA
    .AddItem "WM_ENDSESSION": .ItemData(.NewIndex) = &H16
    .AddItem "WM_ENTERIDLE": .ItemData(.NewIndex) = &H121
    .AddItem "WM_ENTERMENULOOP": .ItemData(.NewIndex) = &H211
    .AddItem "WM_ENTERSIZEMOVE": .ItemData(.NewIndex) = &H231
    .AddItem "WM_ERASEBKGND": .ItemData(.NewIndex) = &H14
    .AddItem "WM_EXITMENULOOP": .ItemData(.NewIndex) = &H212
    .AddItem "WM_EXITSIZEMOVE": .ItemData(.NewIndex) = &H232
    .AddItem "WM_FONTCHANGE": .ItemData(.NewIndex) = &H1D
    .AddItem "WM_FORWARDMSG": .ItemData(.NewIndex) = &H37F
    .AddItem "WM_GETDLGCODE": .ItemData(.NewIndex) = &H87
    .AddItem "WM_GETFONT": .ItemData(.NewIndex) = &H31
    .AddItem "WM_GETHOTKEY": .ItemData(.NewIndex) = &H33
    .AddItem "WM_GETICON": .ItemData(.NewIndex) = &H7F
    .AddItem "WM_GETMINMAXINFO": .ItemData(.NewIndex) = &H24
    .AddItem "WM_GETOBJECT": .ItemData(.NewIndex) = &H3D
    .AddItem "WM_GETTEXT": .ItemData(.NewIndex) = &HD
    .AddItem "WM_GETTEXTLENGTH": .ItemData(.NewIndex) = &HE
    .AddItem "WM_HANDHELDFIRST": .ItemData(.NewIndex) = &H358
    .AddItem "WM_HANDHELDLAST": .ItemData(.NewIndex) = &H35F
    .AddItem "WM_HELP": .ItemData(.NewIndex) = &H53
    .AddItem "WM_HOTKEY": .ItemData(.NewIndex) = &H312
    .AddItem "WM_HSCROLL": .ItemData(.NewIndex) = &H114
    .AddItem "WM_HSCROLLCLIPBOARD": .ItemData(.NewIndex) = &H30E
    .AddItem "WM_ICONERASEBKGND": .ItemData(.NewIndex) = &H27
    .AddItem "WM_IME_CHAR": .ItemData(.NewIndex) = &H286
    .AddItem "WM_IME_COMPOSITION": .ItemData(.NewIndex) = &H10F
    .AddItem "WM_IME_COMPOSITIONFULL": .ItemData(.NewIndex) = &H284
    .AddItem "WM_IME_CONTROL": .ItemData(.NewIndex) = &H283
    .AddItem "WM_IME_ENDCOMPOSITION": .ItemData(.NewIndex) = &H10E
    .AddItem "WM_IME_KEYDOWN": .ItemData(.NewIndex) = &H290
    .AddItem "WM_IME_KEYLAST": .ItemData(.NewIndex) = &H10F
    .AddItem "WM_IME_KEYUP": .ItemData(.NewIndex) = &H291
    .AddItem "WM_IME_NOTIFY": .ItemData(.NewIndex) = &H282
    .AddItem "WM_IME_REPORT": .ItemData(.NewIndex) = &H280
    .AddItem "WM_IME_REQUEST": .ItemData(.NewIndex) = &H288
    .AddItem "WM_IME_SELECT": .ItemData(.NewIndex) = &H285
    .AddItem "WM_IME_SETCONTEXT": .ItemData(.NewIndex) = &H281
    .AddItem "WM_IME_STARTCOMPOSITION": .ItemData(.NewIndex) = &H10D
    .AddItem "WM_IMEKEYDOWN": .ItemData(.NewIndex) = &H290
    .AddItem "WM_IMEKEYUP": .ItemData(.NewIndex) = &H291
    .AddItem "WM_INITDIALOG": .ItemData(.NewIndex) = &H110
    .AddItem "WM_INITMENU": .ItemData(.NewIndex) = &H116
    .AddItem "WM_INITMENUPOPUP": .ItemData(.NewIndex) = &H117
    .AddItem "WM_INPUTLANGCHANGE": .ItemData(.NewIndex) = &H51
    .AddItem "WM_INPUTLANGCHANGEREQUEST": .ItemData(.NewIndex) = &H50
    .AddItem "WM_INTERIM": .ItemData(.NewIndex) = &H10C
    .AddItem "WM_KEYDOWN": .ItemData(.NewIndex) = &H100
    .AddItem "WM_KEYFIRST": .ItemData(.NewIndex) = &H100
    .AddItem "WM_KEYLAST": .ItemData(.NewIndex) = &H108
    .AddItem "WM_KEYUP": .ItemData(.NewIndex) = &H101
    .AddItem "WM_KILLFOCUS": .ItemData(.NewIndex) = &H8
    .AddItem "WM_LBUTTONDBLCLK": .ItemData(.NewIndex) = &H203
    .AddItem "WM_LBUTTONDOWN": .ItemData(.NewIndex) = &H201
    .AddItem "WM_LBUTTONUP": .ItemData(.NewIndex) = &H202
    .AddItem "WM_MBUTTONDBLCLK": .ItemData(.NewIndex) = &H209
    .AddItem "WM_MBUTTONDOWN": .ItemData(.NewIndex) = &H207
    .AddItem "WM_MBUTTONUP": .ItemData(.NewIndex) = &H208
    .AddItem "WM_MDIACTIVATE": .ItemData(.NewIndex) = &H222
    .AddItem "WM_MDICASCADE": .ItemData(.NewIndex) = &H227
    .AddItem "WM_MDICREATE": .ItemData(.NewIndex) = &H220
    .AddItem "WM_MDIDESTROY": .ItemData(.NewIndex) = &H221
    .AddItem "WM_MDIGETACTIVE": .ItemData(.NewIndex) = &H229
    .AddItem "WM_MDIICONARRANGE": .ItemData(.NewIndex) = &H228
    .AddItem "WM_MDIMAXIMIZE": .ItemData(.NewIndex) = &H225
    .AddItem "WM_MDINEXT": .ItemData(.NewIndex) = &H224
    .AddItem "WM_MDIREFRESHMENU": .ItemData(.NewIndex) = &H234
    .AddItem "WM_MDIRESTORE": .ItemData(.NewIndex) = &H223
    .AddItem "WM_MDISETMENU": .ItemData(.NewIndex) = &H230
    .AddItem "WM_MDITILE": .ItemData(.NewIndex) = &H226
    .AddItem "WM_MEASUREITEM": .ItemData(.NewIndex) = &H2C
    .AddItem "WM_MENUCHAR": .ItemData(.NewIndex) = &H120
    .AddItem "WM_MENUCOMMAND": .ItemData(.NewIndex) = &H126
    .AddItem "WM_MENUDRAG": .ItemData(.NewIndex) = &H123
    .AddItem "WM_MENUGETOBJECT": .ItemData(.NewIndex) = &H124
    .AddItem "WM_MENURBUTTONUP": .ItemData(.NewIndex) = &H122
    .AddItem "WM_MENUSELECT": .ItemData(.NewIndex) = &H11F
    .AddItem "WM_MOUSEACTIVATE": .ItemData(.NewIndex) = &H21
    .AddItem "WM_MOUSEFIRST": .ItemData(.NewIndex) = &H200
    .AddItem "WM_MOUSEHOVER": .ItemData(.NewIndex) = &H2A1
    .AddItem "WM_MOUSELAST": .ItemData(.NewIndex) = &H209
    .AddItem "WM_MOUSELEAVE": .ItemData(.NewIndex) = &H2A3
    .AddItem "WM_MOUSEMOVE": .ItemData(.NewIndex) = &H200
    .AddItem "WM_MOUSEWHEEL": .ItemData(.NewIndex) = &H20A
    .AddItem "WM_MOVE": .ItemData(.NewIndex) = &H3
    .AddItem "WM_MOVING": .ItemData(.NewIndex) = &H216
    .AddItem "WM_NCACTIVATE": .ItemData(.NewIndex) = &H86
    .AddItem "WM_NCCALCSIZE": .ItemData(.NewIndex) = &H83
    .AddItem "WM_NCCREATE": .ItemData(.NewIndex) = &H81
    .AddItem "WM_NCDESTROY": .ItemData(.NewIndex) = &H82
    .AddItem "WM_NCHITTEST": .ItemData(.NewIndex) = &H84
    .AddItem "WM_NCLBUTTONDBLCLK": .ItemData(.NewIndex) = &HA3
    .AddItem "WM_NCLBUTTONDOWN": .ItemData(.NewIndex) = &HA1
    .AddItem "WM_NCLBUTTONUP": .ItemData(.NewIndex) = &HA2
    .AddItem "WM_NCMBUTTONDBLCLK": .ItemData(.NewIndex) = &HA9
    .AddItem "WM_NCMBUTTONDOWN": .ItemData(.NewIndex) = &HA7
    .AddItem "WM_NCMBUTTONUP": .ItemData(.NewIndex) = &HA8
    .AddItem "WM_NCMOUSEHOVER": .ItemData(.NewIndex) = &H2A0
    .AddItem "WM_NCMOUSELEAVE": .ItemData(.NewIndex) = &H2A2
    .AddItem "WM_NCMOUSEMOVE": .ItemData(.NewIndex) = &HA0
    .AddItem "WM_NCPAINT": .ItemData(.NewIndex) = &H85
    .AddItem "WM_NCRBUTTONDBLCLK": .ItemData(.NewIndex) = &HA6
    .AddItem "WM_NCRBUTTONDOWN": .ItemData(.NewIndex) = &HA4
    .AddItem "WM_NCRBUTTONUP": .ItemData(.NewIndex) = &HA5
    .AddItem "WM_NCXBUTTONDBLCLK": .ItemData(.NewIndex) = &HAD
    .AddItem "WM_NCXBUTTONDOWN": .ItemData(.NewIndex) = &HAB
    .AddItem "WM_NCXBUTTONUP": .ItemData(.NewIndex) = &HAC
    .AddItem "WM_NEXTDLGCTL": .ItemData(.NewIndex) = &H28
    .AddItem "WM_NEXTMENU": .ItemData(.NewIndex) = &H213
    .AddItem "WM_NOTIFY": .ItemData(.NewIndex) = &H4E
    .AddItem "WM_NOTIFYFORMAT": .ItemData(.NewIndex) = &H55
    .AddItem "WM_NULL": .ItemData(.NewIndex) = &H0
    .AddItem "WM_OTHERWINDOWCREATED": .ItemData(.NewIndex) = &H42
    .AddItem "WM_OTHERWINDOWDESTROYED": .ItemData(.NewIndex) = &H43
    .AddItem "WM_PAINT": .ItemData(.NewIndex) = &HF
    .AddItem "WM_PAINTCLIPBOARD": .ItemData(.NewIndex) = &H309
    .AddItem "WM_PAINTICON": .ItemData(.NewIndex) = &H26
    .AddItem "WM_PALETTECHANGED": .ItemData(.NewIndex) = &H311
    .AddItem "WM_PALETTEISCHANGING": .ItemData(.NewIndex) = &H310
    .AddItem "WM_PARENTNOTIFY": .ItemData(.NewIndex) = &H210
    .AddItem "WM_PASTE": .ItemData(.NewIndex) = &H302
    .AddItem "WM_PENWINFIRST": .ItemData(.NewIndex) = &H380
    .AddItem "WM_PENWINLAST": .ItemData(.NewIndex) = &H38F
    .AddItem "WM_POWER": .ItemData(.NewIndex) = &H48
    .AddItem "WM_POWERBROADCAST": .ItemData(.NewIndex) = &H218
    .AddItem "WM_PRINT": .ItemData(.NewIndex) = &H317
    .AddItem "WM_PRINTCLIENT": .ItemData(.NewIndex) = &H318
    .AddItem "WM_PSD_ENVSTAMPRECT": .ItemData(.NewIndex) = (WM_USER + 5)
    .AddItem "WM_PSD_FULLPAGERECT": .ItemData(.NewIndex) = (WM_USER + 1)
    .AddItem "WM_PSD_GREEKTEXTRECT": .ItemData(.NewIndex) = (WM_USER + 4)
    .AddItem "WM_PSD_MARGINRECT": .ItemData(.NewIndex) = (WM_USER + 3)
    .AddItem "WM_PSD_MINMARGINRECT": .ItemData(.NewIndex) = (WM_USER + 2)
    .AddItem "WM_PSD_PAGESETUPDLG": .ItemData(.NewIndex) = (WM_USER)
    .AddItem "WM_PSD_YAFULLPAGERECT": .ItemData(.NewIndex) = (WM_USER + 6)
    .AddItem "WM_QUERYDRAGICON": .ItemData(.NewIndex) = &H37
    .AddItem "WM_QUERYENDSESSION": .ItemData(.NewIndex) = &H11
    .AddItem "WM_QUERYNEWPALETTE": .ItemData(.NewIndex) = &H30F
    .AddItem "WM_QUERYOPEN": .ItemData(.NewIndex) = &H13
    .AddItem "WM_QUERYUISTATE": .ItemData(.NewIndex) = &H129
    .AddItem "WM_QUEUESYNC": .ItemData(.NewIndex) = &H23
    .AddItem "WM_QUIT": .ItemData(.NewIndex) = &H12
    .AddItem "WM_RASDIALEVENT": .ItemData(.NewIndex) = &HCCCD
    .AddItem "WM_RBUTTONDBLCLK": .ItemData(.NewIndex) = &H206
    .AddItem "WM_RBUTTONDOWN": .ItemData(.NewIndex) = &H204
    .AddItem "WM_RBUTTONUP": .ItemData(.NewIndex) = &H205
    .AddItem "WM_RENDERALLFORMATS": .ItemData(.NewIndex) = &H306
    .AddItem "WM_RENDERFORMAT": .ItemData(.NewIndex) = &H305
    .AddItem "WM_SETCURSOR": .ItemData(.NewIndex) = &H20
    .AddItem "WM_SETFOCUS": .ItemData(.NewIndex) = &H7
    .AddItem "WM_SETFONT": .ItemData(.NewIndex) = &H30
    .AddItem "WM_SETHOTKEY": .ItemData(.NewIndex) = &H32
    .AddItem "WM_SETICON": .ItemData(.NewIndex) = &H80
    .AddItem "WM_SETREDRAW": .ItemData(.NewIndex) = &HB
    .AddItem "WM_SETTEXT": .ItemData(.NewIndex) = &HC
    .AddItem "WM_SETTINGCHANGE": .ItemData(.NewIndex) = WM_WININICHANGE
    .AddItem "WM_SHOWWINDOW": .ItemData(.NewIndex) = &H18
    .AddItem "WM_SIZE": .ItemData(.NewIndex) = &H5
    .AddItem "WM_SIZECLIPBOARD": .ItemData(.NewIndex) = &H30B
    .AddItem "WM_SIZING": .ItemData(.NewIndex) = &H214
    .AddItem "WM_SPOOLERSTATUS": .ItemData(.NewIndex) = &H2A
    .AddItem "WM_STYLECHANGED": .ItemData(.NewIndex) = &H7D
    .AddItem "WM_STYLECHANGING": .ItemData(.NewIndex) = &H7C
    .AddItem "WM_SYNCPAINT": .ItemData(.NewIndex) = &H88
    .AddItem "WM_SYSCHAR": .ItemData(.NewIndex) = &H106
    .AddItem "WM_SYSCOLORCHANGE": .ItemData(.NewIndex) = &H15
    .AddItem "WM_SYSCOMMAND": .ItemData(.NewIndex) = &H112
    .AddItem "WM_SYSDEADCHAR": .ItemData(.NewIndex) = &H107
    .AddItem "WM_SYSKEYDOWN": .ItemData(.NewIndex) = &H104
    .AddItem "WM_SYSKEYUP": .ItemData(.NewIndex) = &H105
    .AddItem "WM_TCARD": .ItemData(.NewIndex) = &H52
    .AddItem "WM_TIMECHANGE": .ItemData(.NewIndex) = &H1E
    .AddItem "WM_TIMER": .ItemData(.NewIndex) = &H113
    .AddItem "WM_UNDO": .ItemData(.NewIndex) = &H304
    .AddItem "WM_UNINITMENUPOPUP": .ItemData(.NewIndex) = &H125
    .AddItem "WM_UPDATEUISTATE": .ItemData(.NewIndex) = &H128
    .AddItem "WM_USER": .ItemData(.NewIndex) = &H400
    .AddItem "WM_USERCHANGED": .ItemData(.NewIndex) = &H54
    .AddItem "WM_VKEYTOITEM": .ItemData(.NewIndex) = &H2E
    .AddItem "WM_VSCROLL": .ItemData(.NewIndex) = &H115
    .AddItem "WM_VSCROLLCLIPBOARD": .ItemData(.NewIndex) = &H30A
    .AddItem "WM_WINDOWPOSCHANGED": .ItemData(.NewIndex) = &H47
    .AddItem "WM_WINDOWPOSCHANGING": .ItemData(.NewIndex) = &H46
    .AddItem "WM_WININICHANGE": .ItemData(.NewIndex) = &H1A
    .AddItem "WM_WNT_CONVERTREQUESTEX": .ItemData(.NewIndex) = &H109
    .AddItem "WM_XBUTTONDBLCLK": .ItemData(.NewIndex) = &H20D
    .AddItem "WM_XBUTTONDOWN": .ItemData(.NewIndex) = &H20B
    .AddItem "WM_XBUTTONUP": .ItemData(.NewIndex) = &H20C
End With
Label1(0).Caption = "Total of " & lstMessages.ListCount & " Windows Messages"
End Sub

Private Sub cmdFind_Click(Index As Integer)
If Len(Trim$(txtCriteria)) = 0 Then
    MsgBox "First supply a search criteria", vbInformation + vbOKOnly
    Exit Sub
End If
Dim xStart As Long, xStop As Long, xStep As Integer, Looper As Long
Select Case Index
Case 0: ' Find first
    xStart = 0
    xStop = lstMessages.ListCount - 1
    xStep = 1
Case 1: ' Find next
    xStart = lstMessages.ListIndex + 1
    xStop = lstMessages.ListCount - 1
    xStep = 1
Case 2: ' Find previous
    xStart = lstMessages.ListIndex - 1
    xStop = 0
    xStep = -1
End Select
For Looper = xStart To xStop Step xStep
    If optCriteria(1) = True Then
        If InStr(GetDescription(lstMessages.ItemData(Looper), True), txtCriteria) > 0 _
            Or InStr(lstMessages.List(Looper), txtCriteria) > 0 Then
            lstMessages.ListIndex = Looper
            Exit For
        End If
    Else
        If lstMessages.ItemData(Looper) = Val(txtCriteria) Then
            lstMessages.ListIndex = Looper
            Exit For
        End If
    End If
Next
If Looper < 0 Or Looper = lstMessages.ListCount Then
    MsgBox "Search criteria not found", vbInformation + vbOKOnly
Else
        cmdFind(1).Enabled = True
        cmdFind(2).Enabled = True
End If
End Sub

Private Sub Form_Load()
    LoadListBox
    Show
    MsgBox "Values for messages retrieved from allapi.net, descriptions of messages retrieved from MSDN", vbInformation + vbOKOnly, "Disclaimer"
End Sub

Private Sub Label2_Click()
Call Labelurl_Click
End Sub

Private Sub Labelurl_Click()
Clipboard.SetText Labelurl.Tag
MsgBox "URL Copied to Clipboard", vbInformation + vbOKOnly, "Windows Messages"
End Sub

Private Sub lstMessages_Click()
    GetDescription lstMessages.ItemData(lstMessages.ListIndex), False
End Sub

Private Sub optCriteria_Click(Index As Integer)
cmdFind(1).Enabled = False
cmdFind(2).Enabled = False
End Sub

Private Sub txtCriteria_Change()
cmdFind(1).Enabled = False
cmdFind(2).Enabled = False
End Sub

Private Sub txtCriteria_GotFocus()
With txtCriteria
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtDeclare_GotFocus()
With txtDeclare
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtValue_GotFocus(Index As Integer)
With txtValue(Index)
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub
