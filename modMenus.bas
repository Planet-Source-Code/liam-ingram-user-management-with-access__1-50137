Attribute VB_Name = "modMenus"
Option Explicit
Option Compare Text
'
' PROJECT NOT COMPATIBLE WITH WinNT 3.x
'
' Read the HowTo files provided. This is a quick summary.
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'  HIGHLIGHTS / IMPROVEMENTS
' - Icons/bitmaps can be displayed on each menu item, even submenus of submenus
' - Sidebars can display text or images and are now clickable if enabled
' - Sidebars can be hidden if a menu scrolls (see how scrolling menus affect sidebars in the HowTo files
' - Sidebar images that are bitmaps can be made transparent
' - Sidebar gradient backgrounds now more flexible and can be applied to image sidebars also
' - 3 properties added to modMenus:
'   -- 1. Highlight menu items with a gradient back color
'   -- 2. Always highlight disabled items
'   -- 3. Change highlighted menu item's font to italics
' - Contents of listboxes/comboboxes can be dynamically included in menus
'   -- selecting one of these menu items will update the listbox/combobox control
'   -- these can be referenced by handle or control name
'   -- owner-drawn listbox/combobox controls not supported at this time
' - Separator bars can have text and can be displayed with a sunken/raised effect
' - All images can be referenced by control name, image list index, or image handle
' - Menu help tips now imitate tooltips
' - All known memory leaks tracked down and put to death
' - 6 ready-to-use custom menus provided: Fonts,Days of Week,Days of Month,Months of Year,Colors,U.S. States
' - OPTIONAL user control provided to permit viewing of graphical menus in designed mode (IDE)
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' To use the graphical menus, two types of forms need to be subclassed.
' 1. SDI Forms (non-MDI forms)
' 2. MDI Parent forms only; their children are automatically subclassed when parent is
'  One function call does it all:  SetMenu form.hwnd, [ImageList], [Tips Options]  -- see that function
'  Popups should always call the SetPopupParentForm routine before calling any popup commands
'  see that routine for more information or read the notes below.
'  Forms automatically un-subclassed when form closes
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' How my mind works....
' I wanted to keep it kinda simple while making it relatively resource efficient, and at the same time,
' giving a user lots of options. Some of those options included referring to a menu items picture by a
' picture handle, image list index, or control (Image1, Picture1, etc).
' To keep it simple, each form has a class created to store menu item information. This seems simple in
' theory, until you play with MDI forms. Each MDI child passes/receives its menu commands from the
' parent therefore referring to MDI child form's menu data in it's own form class was difficult. Didn't want a
' menu item trying to load the parent's control vs the child's control if a control name was encoded in the
' caption. I got around the problem by positively identifying which form currently was active and
' redirecting menu processing to that form's class
' Ok, that problem solved, now came the problem with popups. Since any form can call any other form's
' menu as a popup, tracking the owner of the popup proved difficult. To get around that, the user should
' identify which form owns the popup before calling the popup. To do this, simply call the
' SetPopupParentForm routine and pass the owner's hWnd immediatley before calling the poup.

' That's basically the floor plan for this project.  Store each form's menus in a class for that form.
' I like this approach for two reasons.
' 1. Once processed, menus don't need to get fully processed ever again while the form is open. So we get
'     graphical menus displayed pretty darn quick after they have been displayed once.
' 2. Clean-up is a snap. When a form closes, clean out its class which removes all memory objects.
'     The only true downside is that sidebar images are maintained in memory until the form closes; as also
'     are the few arrays and collections.
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

' Now this routine has several Public subs & functions. Don't use just 'cause
' they're public--only use the ones described in the Readme file which are the
' same ones listed below. The other public routines are either for the optional
' usercontrol or the cMenuItems class. Feel free to preview remarks in these
' routines

' CreateImageSidebar, ChangeImageSidebar
' CreateTextSidebar, ChangeTextSidebar
' CreateMenuCaption, ChangeMenuCaption
' CreateSepartorBar, ChangeSepartorBar
' SetMenu
' RerouteTips
' SetPopupParentForm
' PopupMenuCustom

' The following routines Public or not could be referenced in your programs
' if you choose to. If so, some you may need to make public. But changing any
' of the routines' contents will have undesired effects

' ConvertColor , LoadFontMenu, ExchangeVBcolor, HiWord, LoWord, ShellSort
' and any of the drawing routines as long as you pass the correct parameters
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

' =====================================================================
' Only known issue: If a VB PopupMenu command is activated with a right
' click and the form right-clicked on does not have the focus, the menus
' may not be drawn -- all tags will be seen. The fix is easy.
' Prior to any VB PopupMenu command, add a SetFocus command. i.e...
'               SetFocus
'               PopupMenu mnuMain
' =====================================================================
' You were provided with a ReadMe file which is a detailed help
' file for using this project with all of its options. If you lost the
' file, you can email me at the_foxes@hotmail.com for a replacement.
' =====================================================================

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Following are broken down into several sections.
' Those that are PUBLIC are also referenced within the class cMenuItems
' 1. Section for each DLL referenced. cMenuItems class refs 3 additional DLLs
' 2. Section of standard and custom Type declarations
' 3. Section of standard and custom Constants
' 4. Last section contains private/public variables used throughout application
' =====================================================================
' GDI32 Function Calls
' =====================================================================
' Blt functions
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
' DC manipulation
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Integer
Private Declare Function GetMapMode Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetMapMode Lib "gdi32" (ByVal hDC As Long, ByVal nMapMode As Long) As Long
Private Declare Function SetGraphicsMode Lib "gdi32" (ByVal hDC As Long, ByVal iMode As Long) As Long
' Other drawing functions
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function EnumFontFamilies Lib "gdi32" Alias "EnumFontFamiliesA" (ByVal hDC As Long, ByVal lpszFamily As String, ByVal lpEnumFontFamProc As Long, lParam As Any) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
' =====================================================================
' KERNEL32 Function Calls
' =====================================================================
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (LpVersionInformation As OSVERSIONINFO) As Long
' =====================================================================
' SHELL32 Function Calls
' =====================================================================
Private Declare Function ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociatedIconA" ( _
    ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
' =====================================================================
' USER32 Function Calls
' =====================================================================
' General Windows related functions
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function CopyImage Lib "user32" (ByVal handle As Long, ByVal imageType As Long, ByVal newWidth As Long, ByVal newHeight As Long, ByVal lFlags As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, piconinfo As ICONINFO) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
' Menu related functions
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal Flags As Long) As Long
Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal byPosition As Long, lpMenuItemInfo As MENUITEMINFO) As Boolean
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Public Declare Function IsMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As Any) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As Any) As Long
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hWnd As Long, lprc As RECT) As Long
' =====================================================================
' Standard TYPE Declarations used
' =====================================================================
Public Type POINTAPI                ' general use. Typically used for cursor location
    X As Long
    Y As Long
End Type
Public Type RECT                    ' used to set/ref boundaries of a rectangle
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
Public Type BITMAP                  ' used to determine if an image is a bitmap
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Type DRAWITEMSTRUCT         ' used when owner drawn items are painted
     CtlType As Long
     CtlID As Long
     ItemID As Long
     itemAction As Long
     itemState As Long
     hWndItem As Long
     hDC As Long
     rcItem As RECT
     ItemData As Long
End Type
Public Type ICONINFO                ' used to determine if image is an icon
    fIcon As Long
    xHotSpot As Long
    yHotSpot As Long
    hbmMask As Long
    hbmColor As Long
End Type
Public Type LOGFONT               ' used to create fonts
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName As String * 32
End Type
Private Type NEWTEXTMETRIC      ' used by Font Enumerator routines
   tmHeight As Long
   tmAscent As Long
   tmDescent As Long
   tmInternalLeading As Long
   tmExternalLeading As Long
   tmAveCharWidth As Long
   tmMaxCharWidth As Long
   tmWeight As Long
   tmOverhang As Long
   tmDigitizedAspectX As Long
   tmDigitizedAspectY As Long
   tmFirstChar As Byte
   tmLastChar As Byte
   tmDefaultChar As Byte
   tmBreakChar As Byte
   tmItalic As Byte
   tmUnderlined As Byte
   tmStruckOut As Byte
   tmPitchAndFamily As Byte
   tmCharSet As Byte
   ntmFlags As Long
   ntmSizeEM As Long
   ntmCellHeight As Long
   ntmAveWidth As Long
End Type
Private Type MEASUREITEMSTRUCT  ' used when owner drawn items are first measured
     CtlType As Long
     CtlID As Long
     ItemID As Long
     ItemWidth As Long
     ItemHeight As Long
     ItemData As Long
End Type
Public Type MENUITEMINFO        ' used to retrieve/store menu items
     cbSize As Long
     fMask As Long
     fType As Long
     fState As Long
     wID As Long
     hSubMenu As Long
     hbmpChecked As Long
     hbmpUnchecked As Long
     dwItemData As Long
     dwTypeData As Long 'String
     cch As Long
End Type
Private Type NONCLIENTMETRICS     ' used to retrieve/set system settings
    cbSize As Long
    iBorderWidth As Long
    iScrollWidth As Long
    iScrollHeight As Long
    iCaptionWidth As Long
    iCaptionHeight As Long
    lfCaptionFont As LOGFONT
    iSMCaptionWidth As Long
    iSMCaptionHeight As Long
    lfSMCaptionFont As LOGFONT
    iMenuWidth As Long
    iMenuHeight As Long
    lfMenuFont As LOGFONT
    lfStatusFont As LOGFONT
    lfMessageFont As LOGFONT
End Type
Private Type OSVERSIONINFO          ' used to help identify operating system
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Private Type SHFILEINFO                 ' used to extract icon for files list menus
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * 255
    szTypeName As String * 80
End Type
Public Type MINMAXINFO
  ptReserved As POINTAPI
  ptMaxSize As POINTAPI
  ptMaxPosition As POINTAPI
  ptMinTrackSize As POINTAPI
  ptMaxTrackSize As POINTAPI
End Type
' =====================================================================
' Custom TYPE Declarations used
' =====================================================================
Public Type MenuComponentData
    Caption As String           ' original caption of a owner-drawn converted menu item
    Display As String           ' the caption to display (without hotkeys and codes)
    Cached As String            ' optional listbox type menu caption
    HotKey As String            ' the hotkey to display
    Tip As String               ' the tip to display
    Dimension As POINTAPI       ' height and width of the menu item
    OffsetCx As Integer         ' compensation of menu item width to force to appear standard across O/S's
    ID As Long                  ' the menu item unique ID
    Index As Integer            ' the menu item zero-based position on its submenu
    Icon As String              ' handle to an image to display for a menu item
    ShowBKG As Boolean          ' flag indicating to not make a bitmap menu image transparent (bitmaps only)
    ControlType As Byte         ' 0=combo box;1=list box;2=multiselect listbox
    hControl As Long            ' hWnd reference to a list/combo box if needed
    gControl As Long            ' same as above, but only referenced in child classes (see cMenuItems)
    Status As Integer           ' various attributes of a menu item. Currently, ...
                                '2=Separator,4=Disabled,8=Checked,16=Default,32=Raised Sep Bar,
                                '64=Sidebar,128=Sidebar Hidden,512=hasSubmenus
                                '1024=CustomMenu,2048=ColorMenu,4096=FontMenu,8192=Reserved
End Type
Public Type PanelData          ' each submenu gets a PanelData structure filled
    HasIcons As Boolean        ' indication panel has icons. Needed for checkmark styles
    IsSystem As Boolean        ' indication panel is for a system menu. Used to draw system menu icons
    PanelIcon As Long          ' handle to image to display as a sidebar
    SidebarMenuItem As Long    ' reference to menu item ID that is the sidebar
    SubmenuID As Long          ' reference to which submenu panel belongs to
    Hourglass As Boolean
    Accelerators As String
End Type
' =====================================================================
' Standard CONSTANTS as Constants or Enumerators
' =====================================================================
' //////////// Color constants. \\\\\\\\\\\\\\
Private Const COLOR_ACTIVEBORDER = 10
Private Const COLOR_ACTIVECAPTION = 2
Private Const COLOR_ADJ_MAX = 100
Private Const COLOR_ADJ_MIN = -100
Private Const COLOR_APPWORKSPACE = 12
Private Const COLOR_BACKGROUND = 1
Private Const COLOR_BTNFACE = 15
Private Const COLOR_BTNHIGHLIGHT = 20
Private Const COLOR_BTNLIGHT = 22
Private Const COLOR_BTNSHADOW = 16
Private Const COLOR_BTNTEXT = 18
Private Const COLOR_CAPTIONTEXT = 9
Private Const COLOR_GRAYTEXT = 17
Private Const COLOR_HIGHLIGHT = 13
Private Const COLOR_HIGHLIGHTTEXT = 14
Private Const COLOR_INACTIVEBORDER = 11
Private Const COLOR_INACTIVECAPTION = 3
Private Const COLOR_INACTIVECAPTIONTEXT = 19
Public Const COLOR_MENU = 4
Private Const COLOR_MENUTEXT = 7
Private Const COLOR_SCROLLBAR = 0
Private Const COLOR_WINDOW = 5
Private Const COLOR_WINDOWFRAME = 6
Private Const COLOR_WINDOWTEXT = 8
Public Const NEWTRANSPARENT = 3 'use with SetBkMode()
Private Const WHITENESS = &HFF0062
' //////////// Custom Colors \\\\\\\\\\\\\\\\\
Public Const vbMaroon = 128
Public Const vbOlive = 32896
Public Const vbNavy = 8388608
Public Const vbPurple = 8388736
Public Const vbTeal = 8421376
Public Const vbGray = 8421504
Public Const vbSilver = 12632256
Public Const vbViolet = 9445584
Public Const vbOrange = 42495
Public Const vbGold = 43724 '55295
Public Const vbIvory = 15794175
Public Const vbPeach = 12180223
Public Const vbTurquoise = 13749760
Public Const vbTan = 9221330
Public Const vbBrown = 17510
' //////////// DrawText API Constants \\\\\\\\\\\\\\
Public Const DT_CALCRECT = &H400
Public Const DT_CENTER = &H1
Public Const DT_HIDEPREFIX As Long = &H100000
Public Const DT_LEFT = &H0
Public Const DT_MULTILINE = (&H1)
Public Const DT_NOCLIP = &H100
Public Const DT_NOPREFIX = &H800
Public Const DT_RIGHT = &H2
Public Const DT_SINGLELINE = &H20
Public Const DT_VCENTER As Long = &H4
Public Const DT_WORDBREAK = &H10
' //////////// General Window messages or styles \\\\\\\\\\\\\\
Private Const WM_USER As Long = &H400
Private Const GW_CHILD = 5
Private Const GWL_WNDPROC = (-4)        ' current window procedure for hWnd
Public Const GWL_STYLE = -16            ' current window style for hWnd
Private Const GWL_EXSTYLE = -20         ' current extended window style for hWnd
Public Const GWL_ID = -12               ' current control ID for child window
Private Const TBN_FIRST = (-700&)
Private Const TBN_DROPDOWN = (TBN_FIRST - 10)
Private Const TB_SETSTYLE = (WM_USER + 56)
Private Const TBSTYLE_CUSTOMERASE = &H2000
Private Const TB_GETRECT = (WM_USER + 51)
Private Const WM_ACTIVATE As Long = &H6
Public Const WM_COMMAND As Long = &H111
Private Const WM_DESTROY = &H2
Private Const WM_DRAWITEM = &H2B
Private Const WM_ENTERIDLE = &H121
Private Const WM_ENTERMENULOOP = &H211
Private Const WM_EXITMENULOOP = &H212
Private Const WM_GETMINMAXINFO As Long = &H24&
Private Const WA_INACTIVE As Long = 0
Private Const WM_INITMENU = &H116
Private Const WM_INITMENUPOPUP = &H117
Private Const WM_MDIACTIVATE = &H222
Private Const WM_MDIDESTROY As Long = &H221
Private Const WM_MDIGETACTIVE As Long = &H229
Private Const WM_MDIMAXIMIZE As Long = &H225
Private Const WM_MEASUREITEM = &H2C
Private Const WM_MENUCHAR = &H120
Private Const WM_MENUCOMMAND As Long = &H126
Private Const WM_MENUSELECT As Long = &H11F
Private Const WM_SETFOCUS As Long = &H7
Private Const WS_EX_MDICHILD = &H40&
Private Const WM_KEYUP As Long = &H101
Private Const WM_KEYDOWN = &H100
Private Const WM_SYSKEYDOWN As Long = &H104
Private Const WM_SYSKEYUP As Long = &H105
' //////////// Menu-Related Constants \\\\\\\\\\\\\\
Public Const MF_CHANGE As Long = &H80&
Public Const MF_CHECKED As Long = &H8&
Public Const MF_DEFAULT As Long = &H1000&
Public Const MF_DISABLED As Long = &H2&
Public Const MF_GRAYED As Long = &H1&
Private Const MF_HILITE = &H80&
Public Const MF_MENUBARBREAK As Long = &H20&
Public Const MF_MENUBREAK As Long = &H40&
Private Const MF_MOUSESELECT = &H8000&
Public Const MF_OWNERDRAW = &H100
Public Const MF_POPUP As Long = &H10&
Public Const MF_SEPARATOR = &H800
Public Const MIIM_DATA = &H20
Public Const MIIM_ID As Long = &H2
Public Const MIIM_STATE As Long = &H1
Public Const MIIM_SUBMENU As Long = &H4
Public Const MIIM_TYPE = &H10
Private Const MNC_EXECUTE = 2
Private Const MNC_IGNORE = 0
Private Const MNC_SELECT = 3
Private Const ODA_DRAWENTIRE As Long = &H1
Private Const ODT_MENU = 1
Private Const ODS_SELECTED = &H1
Private Const TPM_RETURNCMD As Long = &H100&
Private Const TPM_NONOTIFY As Long = &H80&
Private Const TPM_LEFTALIGN = &H0&
Private Const TPM_VERTICAL = &H40&
Private Const TPM_LEFTBUTTON = &H0&
' System Menu items that will have their icons manually redrawn
Public Const SC_CLOSE = &HF060
Public Const SC_MINIMIZE = &HF020
Public Const SC_MAXIMIZE = &HF030
Public Const SC_RESTORE = &HF120
' Miscellaneous
Public Const RASTER_FONTTYPE As Long = &H1
Public Const TRUETYPE_FONTTYPE As Long = &H4
Private Const SPI_GETWORKAREA = 48
Private Const SHGFI_ICON = &H100
Private Const SHGFI_SMALLICON = &H1
Private Const SHGFI_USEFILEATTRIBUTES = &H10
Private Const VK_SHIFT As Long = &H10
Private Const VK_CONTROL As Long = &H11
Private Const VK_MENU As Long = &H12
' =====================================================================
' Custom CONSTANTS as Constants or Enumerators
' =====================================================================
' ////////////// Used to keep track of owner drawn menu item attributes \\\\\\\\\\\\\\
' Ref custom Type MenuComponentData above
Public Const lv_mSep As Integer = 2
Public Const lv_mDisabled As Integer = 4
Public Const lv_mChk As Integer = 8
Public Const lv_mDefault As Integer = 16
Public Const lv_mSepRaised As Integer = 32
Public Const lv_mSBar As Integer = 64
Public Const lv_mSBarHidden As Integer = 128
Public Const lv_mSubmenu As Integer = 512
Public Const lv_mCustom As Integer = 1024
Public Const lv_mColor As Integer = 2048
Public Const lv_mFont As Integer = 4096
' ////////////// Used for functions to format menu captions \\\\\\\\\\\\\\
Public Enum MenuImageType
    lv_ImgListIndex = 0
    lv_ImgHandle = 1
    lv_ImgControl = 2
End Enum
Public Enum MenuCtrlType
    lv_ListBox = 1
    lv_ComboBox = 2
End Enum
Public Enum MenuCaptionProps
    lv_Caption = 0
    lv_ImgID = 1
    lv_Bold = 2
    lv_Tip = 3
    lv_ListBoxID = 4
    lv_ComboxID = 5
    lv_ShowIconBkg = 6
    lv_HotKey = 7
    lv_FilesPath = 8
End Enum
Public Enum SidebarTextProps
    lv_txtText = 1
    lv_txtForeColor = 2
    lv_txtBackColor = 3
    lv_txtGradientColor = 4
    lv_txtFontName = 5
    lv_txtFontSize = 6
    lv_txtMinFontSize = 7
    lv_txtWidth = 8
    lv_txtAlignment = 9
    lv_txtTip = 10
    lv_txtBold = 11
    lv_txtItalic = 12
    lv_txtUnderline = 13
    lv_txtNoScroll = 14
    lv_txtDisabled = 15
End Enum
Public Enum SidebarImgProps
    lv_imgImgID = 1
    lv_imgBackColor = 2
    lv_imgGradientColor = 3
    lv_imgWidth = 4
    lv_imgAlignment = 5
    lv_imgTip = 6
    lv_imgNoScroll = 7
    lv_imgTransparent = 8
    lv_imgDisabled = 9
End Enum
Public Enum MenuSepProps
    lv_sCaption = 0
    lv_sRaisedEffect = 1
End Enum
Public Enum FontTypeEnum
    lv_fAllFonts = 0
    lv_fTrueType = 1
    lv_fNonTrueType = 2
End Enum
Public Enum AlignmentEnum
    lv_TopOfMenu = 1
    lv_BottomOfMenu = 2
    lv_CenterOfMenu = 0
End Enum
Public Enum CstmMonth
    lv_cDefault = 0
    lv_cCalendarQuarter = 1
    lv_cFiscalQuarter = 2
End Enum
' ////////////// Used to set/reset HDC objects \\\\\\\\\\\\\\
Public Enum ColorObjects
    cObj_Brush = 0
    cObj_Pen = 1
    cObj_Text = 2
End Enum
' ////////////// Used when calling SetMenu \\\\\\\\\\\\\\
Public Enum SubClassContainers
    lv_NonMDIform = 0                               ' typical SDI form, no children
    lv_MDIform_ChildrenHaveMenus = 0       ' typical MDI form if children have their own menus (Default for MDI forms)
    lv_MDIform_ChildrenMenuless = 5          ' typical MDI form when no child forms have menus...
    ' note: all child forms subclassed automatically will have their property set to lv_MDIchildForm_NoMenus
    lv_VB_Toolbar = 1                                 ' typical standard toolbar
    lv_MDIchildForm_NoMenus = 4              ' MDI child form has no menus see above note
    lv_MDIchildForm_WithMenus = 3           ' MDI child form has menus (default for MDI child forms)
End Enum
    
' ///////////////// PROJECT-WIDE VARIABLES \\\\\\\\\\\\\\
Public DefaultIcon As Long ' used by associated user control
'======================================================================
' following variable will restore menu items back to their original status
' after subclassing has been terminated. Don't set this during runtime
' as it will unnecessarily reset menus when a form closes
Public AmInIDE As Boolean  ' used by associated user control
'======================================================================
'                              IMPORTANT
' Somewhere in your primary form, declare following variable to
' True or False
' Set the following constant to TRUE if you need to debug your code
' When set to true, forms are not subclassed
' When set to False, stopping your code will crash VB
'======================================================================
Public bAmDebugging As Boolean

' Types used to retrieve current menu item information from the
' cMenuItems class and returned to DoMeausreItem & DoDrawItem
Public XferMenuData As MenuComponentData
Public XferPanelData As PanelData
' for Win98/ME--they seem to add extra pixels to menus & we account for the difference
Private ExtraOffsetX As Integer
' storage of the 7 fonts used for menus. See: CreateDestroyMenuFont
Private m_Font(0 To 7)
' collection of subclassed forms
Private colMenuItems As Collection
' collection of displayed menu panels. Prevents reprocessing of items
Private OpenMenus As Collection
' used only for the custom font menu in cMenuItems class to retrieve font names
Private vFonts() As String
' handle to form which owns menu being displayed
Private hWndRedirect As String
' used for popups to positively identify owner of the popup menu
Private tempRedirect As Long    ' see: SetPopupParentForm
' used to determine highlighting of disabled items if selected by keyboard vs mouse
Private bKeyBoardSelect As Boolean
'////////////////////// Public Properties \\\\\\\\\\\\\\\\\
Private bHiLiteDisabled As Boolean  'see HighlightDisabledMenuItems
Private bItalicSelected As Boolean  'see ItalicizeSelectedItems
Private bGradientSelect As Boolean  'see HighlightGradient
Private mFontName As String         'see MenuFontName
Private mFontSize As Single         'see MenuFontSize
Private vMenuListBox As Long        'see MenuCaptionListBox
Private bReturnMDIkeystrokes As Boolean ' see ReturnMDIkeystrokes
Private bRaisedIcons As Boolean     'see RaisedIconOnSelect
Private bXPcheckmarks As Boolean  ' XP/Win2K style checks
    '/////////////////////// color options \\\\\\\\\\\\\\\\\\\\\\\\\
    Private bModuleInitialized As Boolean       '
    Private lSelectBColor As Long       'see SelectedItemBackColor
    Private TextColorNormal As Long
    Private TextColorSelected As Long
    Private TextColorDisabledDark As Long
    Private TextColorDisabledLight As Long
    Private TextColorSeparatorBar As Long
    Private SeparatorBarColorDark As Long
    Private SeparatorBarColorLight As Long
    Private CheckedIconBColor As Long
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Private FloppyIcon As Long          'see GetFloppyIcon
Private tbarClass() As String       'see AddToolbarClass
' some menus could take time to display--the font custom menu, the drives
' custom menu if network drives exist, and file list menus where icon is
' retrieved from file's associated executable. So the user doesn't think
' their computer is slow or somethin' we will add an hourglass to these
' menu types and any menu that has 50 or more items. The hourglass will
' be reset once the menu has been displayed or the menu loop terminates
Private bUseHourglass As Boolean

Public Property Let MenuFontName(sFontName As String)
' =====================================================================
' by default, menu items have the system menu font name
' This property can change the font name to anything. Suggest setting this property in the first form that
' is displayed as each call to change the font name or font size will force the program to restore
' all menu items to non-owner drawn status which forces the program to remeasure each menu item again
' =====================================================================
    If Len(sFontName) > 0 And sFontName <> mFontName Then
        mFontName = sFontName
        ' here we destroy the previous memory fonts & recreate them using new font name
        CreateDestroyMenuFont False, False
        CreateDestroyMenuFont True, False
        If bItalicSelected Then CreateDestroyMenuFont True, True
        ' now we need to destroy our collection of processed menu items so they
        ' can be reprocessed and menu panels remeasured. Easier to delete the
        ' class but then we lose a bunch of variables we want to keep
        If Not colMenuItems Is Nothing Then
            Dim I As Integer
            For I = colMenuItems.Count To 1 Step -1
                colMenuItems(I).RestoreMenus
            Next
        End If
    End If
End Property

Public Property Get MenuFontName() As String
If mFontName = "" Then
    ' in order to set the font, we must first determine what it is
    Dim ncm As NONCLIENTMETRICS, I As Integer
    ncm.cbSize = Len(ncm)
    ' this will return the system menu font info
    SystemParametersInfo 41, 0, ncm, 0
    I = InStr(ncm.lfMenuFont.lfFaceName, Chr$(0))
    If I = 0 Then I = Len(ncm.lfMenuFont.lfFaceName) + 1
    mFontName = Left$(ncm.lfMenuFont.lfFaceName, I - 1)
    If mFontSize = 0 Then mFontSize = Abs(ncm.lfMenuFont.lfHeight) * 0.72
End If
MenuFontName = mFontName
End Property

Public Property Let MenuFontSize(NewSize As Single)
' =====================================================================
' by default, menu items have the system menu font size
' This property can change the font size to anything. Suggest setting this property in the first form that
' is displayed as each call to change the font name or font size will force the program to restore
' all menu items to non-owner drawn status which forces the program to remeasure each menu item again
' =====================================================================
    If NewSize <> mFontSize And NewSize > 0 Then
        mFontSize = NewSize
        ' here we destroy the previous memory fonts & recreate them using new font size
        CreateDestroyMenuFont False, False
        CreateDestroyMenuFont True, False
        If bItalicSelected Then CreateDestroyMenuFont True, True
        ' now we need to destroy our collection of processed menu items so they
        ' can be reprocessed and menu panels remeasured. Easier to destroy the
        ' class but then we lose a bunch of values we want to keep
        If Not colMenuItems Is Nothing Then
            Dim I As Integer
            For I = colMenuItems.Count To 1 Step -1
                colMenuItems(I).RestoreMenus
            Next
        End If
    End If
End Property
Public Property Get MenuFontSize() As Single
    MenuFontSize = mFontSize
End Property

Public Property Let MenuCaptionListBox(hWnd As Long)
' =====================================================================
' If menu captions are stored in a list box, that list box must be made
' available at all times. This property can be set ONLY ONCE. This is
' done to prevent using multiple listboxes which would inevitably lead
' to the wrong listbox being set at the wrong time and then the wrong
' captions being placed on a menu. If you tweak this, be very careful!
'
' One exception. If the original menu caption list box that was set
' previously is now closed 'cause its form is closed, it can be reset
' =====================================================================
If vMenuListBox Then
    If IsWindow(vMenuListBox) = 0 Then vMenuListBox = 0
End If
If Not vMenuListBox Then vMenuListBox = hWnd
End Property
Public Property Get MenuCaptionListBox() As Long
            MenuCaptionListBox = vMenuListBox
End Property

Public Property Let HighlightGradient(bGradient As Boolean)
' =====================================================================
' by default, menu items being highlighted are done so with a solid back color (system defined)
' This property can change it to highight with a gradient effect.
' The gradient is from "system back highlight" color to "menu backcolor" color (both are system defined)
' =====================================================================
    bGradientSelect = bGradient
End Property
Public Property Get HighlightGradient() As Boolean
    HighlightGradient = bGradientSelect
End Property

' =====================================================================
' All properties between these ======= symbos are menu item properties and are explained
' in more detail in the readme.html file

Public Property Get RaisedIconOnSelect() As Boolean
    RaisedIconOnSelect = bRaisedIcons
End Property
Public Property Let RaisedIconOnSelect(bYesNo As Boolean)
    bRaisedIcons = bYesNo
End Property
Public Property Get CheckMarksXPstyle() As Boolean
    CheckMarksXPstyle = bXPcheckmarks
End Property
Public Property Let CheckMarksXPstyle(bYesNo As Boolean)
    bXPcheckmarks = bYesNo
End Property

Public Property Get SelectedItemBackColor() As Long
    If Not bModuleInitialized Then LoadDefaultColors
    SelectedItemBackColor = lSelectBColor
End Property
Public Property Let SelectedItemBackColor(lColor As Long)
    If Not bModuleInitialized Then LoadDefaultColors
    lSelectBColor = ConvertColor(lColor)
End Property
Public Property Get SelectedItemTextColor() As Long
    If Not bModuleInitialized Then LoadDefaultColors
    SelectedItemTextColor = TextColorSelected
End Property
Public Property Let SelectedItemTextColor(lColor As Long)
    If Not bModuleInitialized Then LoadDefaultColors
    TextColorSelected = ConvertColor(lColor)
End Property
Public Property Get MenuItemTextColor() As Long
    If Not bModuleInitialized Then LoadDefaultColors
    MenuItemTextColor = TextColorNormal
End Property
Public Property Let MenuItemTextColor(lColor As Long)
    If Not bModuleInitialized Then LoadDefaultColors
    TextColorNormal = ConvertColor(lColor)
End Property
Public Property Get DisabledTextColor_Dark() As Long
    If Not bModuleInitialized Then LoadDefaultColors
    DisabledTextColor_Dark = TextColorDisabledDark
End Property
Public Property Let DisabledTextColor_Dark(lColor As Long)
    If Not bModuleInitialized Then LoadDefaultColors
    TextColorDisabledDark = ConvertColor(lColor)
End Property
Public Property Get DisabledTextColor_Light() As Long
    If Not bModuleInitialized Then LoadDefaultColors
    DisabledTextColor_Light = TextColorDisabledLight
End Property
Public Property Let DisabledTextColor_Light(lColor As Long)
    If Not bModuleInitialized Then LoadDefaultColors
    TextColorDisabledLight = ConvertColor(lColor)
End Property
Public Property Get SeparatorBarTextColor() As Long
    If Not bModuleInitialized Then LoadDefaultColors
    SeparatorBarTextColor = TextColorSeparatorBar
End Property
Public Property Let SeparatorBarTextColor(lColor As Long)
    If Not bModuleInitialized Then LoadDefaultColors
    TextColorSeparatorBar = ConvertColor(lColor)
End Property
Public Property Get SeparatorBarColor_Dark() As Long
    If Not bModuleInitialized Then LoadDefaultColors
    SeparatorBarColor_Dark = SeparatorBarColorDark
End Property
Public Property Let SeparatorBarColor_Dark(lColor As Long)
    If Not bModuleInitialized Then LoadDefaultColors
    SeparatorBarColorDark = ConvertColor(lColor)
End Property
Public Property Get SeparatorBarColor_Light() As Long
    If Not bModuleInitialized Then LoadDefaultColors
    SeparatorBarColor_Light = SeparatorBarColorLight
End Property
Public Property Let SeparatorBarColor_Light(lColor As Long)
    If Not bModuleInitialized Then LoadDefaultColors
    SeparatorBarColorLight = ConvertColor(lColor)
End Property
Public Property Get CheckedIconBackColor() As Long
    If Not bModuleInitialized Then LoadDefaultColors
    CheckedIconBackColor = CheckedIconBColor
End Property
Public Property Let CheckedIconBackColor(lColor As Long)
    If Not bModuleInitialized Then LoadDefaultColors
    CheckedIconBColor = ConvertColor(lColor)
End Property
' =====================================================================

Public Property Let ReturnMDIkeystrokes(bYesNo As Boolean)
' =====================================================================
' This property allows MD Parents to receive Key_Up & Key_Down events
' MDI parent when there are no MDI children opened. Otherwise, you
' should use the Key_Down events within the MDI child to trap keystrokes

' IMPORTANT SIDE NOTE: WinME users will not be able to use this property.
' completely. Per MSDN: the GetKeyState function has been disabled on ME only
' This will force the Shift parameter retunred in cTips MDIKeyDown & MDIKeyUp
' events to return the value of zero for each keystroke pressed
' =====================================================================
    bReturnMDIkeystrokes = bYesNo
End Property
Public Property Get ReturnMDIkeystrokes() As Boolean
    ReturnMDIkeystrokes = bReturnMDIkeystrokes
End Property

Public Property Let HighlightDisabledMenuItems(bHiLite As Boolean)
' =====================================================================
' By default, disabled items are highlighted in the following 2 cases.
' This flag will highlight disabled items in every case."
' 1. System Menu items
' 2. Items navigated via the keyboard
' =====================================================================
    bHiLiteDisabled = bHiLite
End Property
Public Property Get HighlightDisabledMenuItems() As Boolean
    HighlightDisabledMenuItems = bHiLiteDisabled
End Property

Public Property Let ItalicizeSelectedItems(bItalics As Boolean)
' =====================================================================
' This option will italicize items when they are highlighted
' =====================================================================
If bItalics = bItalicSelected Then Exit Property
    bItalicSelected = bItalics
    ' create italic fonts if needed (2 fonts created: 1:normal, italic font, 2:bold, italic font
    If bItalics Then CreateDestroyMenuFont True, True
End Property
Public Property Get ItalicizeSelectedItems() As Boolean
ItalicizeSelectedItems = bItalicSelected
End Property

Public Property Get Win98MEoffset() As Integer
' =====================================================================
' This read-only property passes back the extra pixels if system is Win98/ME. See DetermineOS
' =====================================================================
    Win98MEoffset = ExtraOffsetX
End Property
Public Function CreateTextSidebar(Caption As String, FontName As String, FontSize As Single, _
    Optional MinFontSize As Single = 9, Optional Bold As Boolean = False, Optional Underline As Boolean = False, _
    Optional Italic As Boolean = False, Optional ForeColor As Long, Optional Backcolor As Long = -1, Optional Gradient2ndColor As Long = vbNull, _
    Optional Width As Integer = 32, Optional Alignment As AlignmentEnum = lv_BottomOfMenu, _
    Optional NoShowIfScrolls As Boolean = False, Optional Tip As String, Optional AlwaysDisabled As Boolean) As String
' =====================================================================
' Function will create a text sidebar item with every option made available
' =====================================================================
Dim wCaption As String, sValue As String
If Caption = "" Then Caption = " "
If FontName = "" Then FontName = "Arial"
Caption = "{Sidebar|Text:" & Caption & "|Font:" & FontName
If FontSize < 9 Then wCaption = "|FSize:9" Else wCaption = "|FSize:" & FontSize
If MinFontSize Then wCaption = wCaption & "|MinFSize:" & MinFontSize
If Bold Then wCaption = wCaption & "|Bold"
If Underline Then wCaption = wCaption & "|Underline"
If Italic Then wCaption = wCaption & "|Italic"
If AlwaysDisabled Then wCaption = wCaption & "|SBDisabled"
wCaption = wCaption & "|FColor:" & ForeColor
wCaption = wCaption & "|BColor:" & Backcolor
If Gradient2ndColor <> vbNull Then wCaption = wCaption & "|GColor:" & Gradient2ndColor
If Width < 16 Then Width = 16
wCaption = wCaption & "|Width:" & Width
Select Case Alignment
    Case lv_TopOfMenu: wCaption = wCaption & "|Align:Top"
    Case lv_BottomOfMenu: wCaption = wCaption & "|Align:Bot"
    Case lv_CenterOfMenu: wCaption = wCaption & "|Align:Ctr"
End Select
If NoShowIfScrolls Then wCaption = wCaption & "|NoScroll"
If Len(Tip) Then Caption = Caption & "|Tip:" & Tip
CreateTextSidebar = Caption & wCaption & "}"
End Function

Public Function ChangeTextSidebar(CaptionNow As String, Property As SidebarTextProps, Optional newValue As Variant) As String
Dim sNewCaption As String, sValue As String, wCaption As String, sProp As String
' =====================================================================
' This function will add, modify or delete a specific flag/value in the text sidebar caption
' =====================================================================
ChangeTextSidebar = CaptionNow
If IsMissing(newValue) Then sProp = "" Else sProp = CStr(newValue)
Dim sParts(1 To 15) As String, sTarget As String
Dim I As Integer
For I = 1 To UBound(sParts)
    sTarget = Choose(I, "Text:", "FColor:", "BColor:", "GColor:", "Font:", "FSize:", "MinFSize:", "Width:", "Align:", "Tip", "Bold", "Italic", "Underline", "NoScroll", "SBDisabled")
    ReturnComponentValue CaptionNow, sTarget, sValue
    If Len(sValue) Then
        sParts(I) = sTarget
        If I < 11 Then sParts(I) = sParts(I) & sValue
        sParts(I) = sParts(I) & "|"
    Else
        sParts(I) = ""
    End If
Next
Select Case Property
Case lv_txtText
    sTarget = "Text:"
    If Len(sProp) = 0 Then sProp = " "
Case lv_txtForeColor
    sTarget = "FColor:"
    sProp = CStr(Val(sProp))
Case lv_txtBackColor
    If Len(sProp) = 0 Then sProp = vbNull Else sProp = CStr(Val(sProp))
    sTarget = "BColor:"
Case lv_txtGradientColor
    If Len(sProp) = 0 Then sProp = vbNull Else sProp = CStr(Val(sProp))
    sTarget = "GColor:"
Case lv_txtFontName
    If Len(sProp) = 0 Then sProp = "Tahoma"
    sTarget = "Font:"
Case lv_txtFontSize
    If Len(sProp) = 0 Then sProp = 9 Else sProp = CStr(Val(sProp))
    sTarget = "FSize:"
Case lv_txtMinFontSize
    If Len(sProp) = 0 Then
        sTarget = ""
        sParts(Property) = ""
    Else
        sProp = CStr(Val(sProp))
        sTarget = "MinFSize:"
    End If
Case lv_txtBold, lv_txtItalic, lv_txtUnderline, lv_txtNoScroll, lv_txtDisabled
    If Len(sProp) = 0 Then sProp = "False"
    If CBool(sProp) = False Then
        sParts(Property) = ""
    Else
        sParts(Property) = Choose(Property - 10, "Bold|", "Italic|", "Underline|", "NoScroll|", "SBDisabled|")
    End If
    sTarget = ""
Case lv_txtTip
    sTarget = "Tip:"
    If Len(sProp) = 0 Then sTarget = "": sParts(Property) = ""
Case lv_txtWidth
    If Val(sProp) < 35 Then sProp = 35
    sTarget = "Width:"
Case lv_txtAlignment
    sTarget = "Align:"
    Select Case sProp
    Case "1", "Top": sProp = "Top"
    Case "2", "Bot": sProp = "Bot"
    Case Else:
        sTarget = ""
        sParts(Property) = ""
    End Select
Case Else
    ChangeTextSidebar = CaptionNow
    Exit Function
End Select
If Len(sTarget) Then
    If Len(sProp) Then sParts(Property) = sTarget & sProp & "|" Else sParts(Property) = ""
End If
wCaption = ""
For I = 1 To UBound(sParts)
    wCaption = wCaption & sParts(I)
Next
If Len(wCaption) Then wCaption = Left$(wCaption, Len(wCaption) - 1) & "}"
ChangeTextSidebar = "{Sidebar|" & wCaption
Erase sParts
'Debug.Print "Passed "; CaptionNow
'Debug.Print "Change "; ChangeTextSidebar
End Function

Public Function CreateImageSidebar(ImgType As MenuImageType, ImgID As String, _
    Optional Transparent As Boolean = False, Optional Backcolor As Long = -1, Optional Gradient2ndColor As Long = vbNull, _
    Optional Width As Integer = 32, Optional Alignment As AlignmentEnum, Optional NoShowIfScrolls As Boolean = False, _
    Optional Tip As String) As String
' =====================================================================
' This function will create a image sidebar and make every option available
' =====================================================================
Dim Caption As String, sValue As String
' we'll add the image.  The order is not important
If Len(ImgID) = 0 Then Exit Function
If (ImgType > lv_ImgListIndex - 1 And ImgType < lv_ImgControl + 1) Then
    If ImgType = lv_ImgListIndex Then sValue = "IMG:i" & ImgID Else sValue = "IMG:" & ImgID
    Caption = sValue
Else
    Exit Function
End If
If Transparent Then Caption = Caption & "|Transparent"
Caption = Caption & "|BColor:" & Backcolor
If Gradient2ndColor <> vbNull Then Caption = Caption & "|GColor:" & Gradient2ndColor
If Width < 32 Then Width = 32
Caption = Caption & "|Width:" & Width
Select Case Alignment
    Case 1: Caption = Caption & "|Align:Top"
    Case 2: Caption = Caption & "|Align:Bot"
    Case Else
End Select
If NoShowIfScrolls Then Caption = Caption & "|NoScroll"
If Len(Tip) Then Caption = Caption & "|Tip:" & Tip
CreateImageSidebar = "{Sidebar|" & Caption & "}"
End Function

Public Function ChangeImageSidebar(CaptionNow As String, Property As SidebarImgProps, Optional newValue As Variant) As String
Dim sNewCaption As String, sValue As String, wCaption As String, sProp As String
' =====================================================================
' This option will add, remove or modify a specific flag/value
' =====================================================================
ChangeImageSidebar = CaptionNow
If IsMissing(newValue) Then sProp = "" Else sProp = CStr(newValue)

Dim sParts(0 To 9) As String, sTarget As String
Dim I As Integer
For I = 1 To UBound(sParts)
    sTarget = Choose(I, "IMG:", "BColor", "GColor", "Width:", "Align:", "Tip:", "NoScroll", "Transparent", "SBDisabled")
    ReturnComponentValue CaptionNow, sTarget, sValue
    If Len(sValue) Then
        sParts(I) = sTarget
        If I < 7 Then sParts(I) = sParts(I) & sValue
        sParts(I) = sParts(I) & "|"
    Else
        sParts(I) = ""
    End If
Next
Select Case Property
Case lv_imgImgID
    sTarget = "IMG:"
    If Len(sProp) = 0 Then
        ChangeImageSidebar = ""
        Exit Function
    End If
Case lv_imgBackColor
    If Len(sProp) = 0 Then sProp = vbButtonFace Else sProp = CStr(Val(sProp))
    sTarget = "BColor:"
Case lv_imgGradientColor
    If Len(sProp) = 0 Then sProp = vbNull Else sProp = CStr(Val(sProp))
    sTarget = "GColor:"
Case lv_imgAlignment
    sTarget = "Align:"
    Select Case sProp
    Case "1", "Top": sProp = "Top"
    Case "2", "Bot": sProp = "Bot"
    Case Else:
        sTarget = ""
        sParts(Property) = ""
    End Select
Case lv_imgTip
    sTarget = "Tip:"
    If Len(sProp) = 0 Then sParts(Property) = "": sTarget = ""
Case lv_imgWidth
    If Val(sProp) < 32 Then sProp = 32 Else sProp = CStr(Val(sProp))
    sTarget = "Width:"
Case lv_imgNoScroll, lv_imgTransparent, lv_imgDisabled
    If Len(sProp) = 0 Then sProp = "False"
    If CBool(sProp) Then
        sParts(Property) = Choose(Property - 6, "NoScroll|", "Transparent|")
    Else
        sParts(Property) = ""
    End If
    sTarget = ""
End Select
If Len(sTarget) Then
    If Len(sProp) Then sParts(Property) = sTarget & sProp & "|" Else sParts(Property) = ""
End If
wCaption = ""
For I = 1 To UBound(sParts)
    wCaption = wCaption & sParts(I)
Next
If Len(wCaption) Then wCaption = Left$(wCaption, Len(wCaption) - 1) & "}"
ChangeImageSidebar = "{Sidebar|" & sNewCaption & wCaption
Erase sParts
'Debug.Print "Passed (image) "; CaptionNow
'Debug.Print "Change (image) "; ChangeImageSidebar
End Function

Public Function CreateSepartorBar(Optional Caption As String, Optional RaisedEffect As Boolean) As String
' =====================================================================
' This function will create a separtor bar and make every option available
' =====================================================================
Dim newCaption As String
If Len(Caption) = 0 Then newCaption = "-" Else newCaption = Caption
If Left(newCaption, 1) <> "-" Then newCaption = "-" & newCaption
If RaisedEffect Then newCaption = newCaption & "{Raised}"
CreateSepartorBar = newCaption
End Function

Public Function ChangeSepartorBar(CaptionNow As String, Property As MenuSepProps, Optional newValue As Variant)
Dim sNewCaption As String, sValue As String, wCaption As String, sProp As String
' =====================================================================
' This function will add, modify or remove a specific flag/value
' =====================================================================
SeparateCaption CaptionNow, sNewCaption, wCaption
If IsMissing(newValue) Then sProp = "" Else sProp = CStr(newValue)
Select Case Property
Case lv_sCaption
    If Len(sProp) = 0 Then sProp = "-"
    If Left(sProp, 1) <> "-" Then sProp = "-" & sProp
    ReturnComponentValue wCaption, "Raised", sValue
    If Len(sValue) Then sValue = "{Raised}"
    sNewCaption = sProp & sValue
Case lv_sRaisedEffect
    If sProp = "" Then sProp = "False"
    If CBool(sProp) Then sProp = "{Raised}" Else sProp = ""
    sNewCaption = sNewCaption & sProp
End Select
ChangeSepartorBar = sNewCaption
End Function

Public Function CreateMenuCaption(Caption As String, Optional ImgType As MenuImageType, _
    Optional ImgID As String, Optional NoTransparency As Boolean, Optional HotKey As String, _
    Optional BoldText As Boolean = False, Optional Tip As String, _
    Optional ListComboType As MenuCtrlType, Optional ListComboID As String, _
    Optional ListsFiles As Boolean, Optional FilesPath As String) As String
' =====================================================================
' This function will create a typical, non-sidebar, non-separator bar caption and make all options available
' =====================================================================
Dim wCaption As String, sValue As String
' we'll add the image if any.  The order is not important
If (ImgType > lv_ImgListIndex - 1 And ImgType < lv_ImgControl + 1) And Len(ImgID) > 0 Then
    If ImgType = lv_ImgListIndex Then sValue = "IMG:i" & ImgID Else sValue = "IMG:" & ImgID
    wCaption = sValue
End If
If NoTransparency Then
    If Len(wCaption) Then sValue = "|" Else sValue = ""
    wCaption = wCaption & sValue & "ImgBkg"
End If
If Len(HotKey) Then ' add the hot key if any
    If Len(wCaption) Then sValue = "|" Else sValue = ""
    wCaption = wCaption & sValue & "HotKey:" & HotKey
End If
If BoldText = True Then ' add the Default type text (bold)
    If Len(wCaption) Then sValue = "|" Else sValue = ""
    wCaption = wCaption & sValue & "Default"
End If
If ListsFiles Then
    If Len(wCaption) Then sValue = "|" Else sValue = ""
    wCaption = wCaption & "|Files:"
    If Len(FilesPath) Then wCaption = wCaption & FilesPath Else wCaption = wCaption & "-1"
End If
If Len(Tip) Then    ' add the Tip
    If Len(wCaption) Then sValue = "|" Else sValue = ""
    wCaption = wCaption & sValue & "Tip:" & Tip
End If                  ' add the listbox/combo box control reference if any
If (ListComboType = lv_ComboBox Or ListComboType = lv_ListBox) And Len(ListComboID) > 0 Then
    If Len(wCaption) Then sValue = "|" Else sValue = ""
    If ListComboType = lv_ComboBox Then sValue = sValue & "CB:" Else sValue = sValue & "LB:"
    wCaption = wCaption & sValue & ListComboID
End If
If Len(wCaption) Then wCaption = "{" & wCaption & "}"
CreateMenuCaption = Caption & wCaption
End Function

Public Function ChangeMenuCaption(CaptionNow As String, Property As MenuCaptionProps, Optional newValue As Variant) As String
' =====================================================================
' This function will add, remove or modify a specific flag/value from a typical menu caption
' =====================================================================
Dim sNewCaption As String, sValue As String, wCaption As String, sProp As String
Dim sParts(0 To 8) As String, sTarget As String
Dim I As Integer

If IsMissing(newValue) Then sProp = "" Else sProp = CStr(newValue)
ChangeMenuCaption = CaptionNow
sNewCaption = CaptionNow
SeparateCaption CaptionNow, sNewCaption, wCaption
For I = 1 To UBound(sParts)
    sTarget = Choose(I, "IMG:", "Default", "Tip:", "LB:", "CB:", "IMGBKG", "HotKey:", "Files:")
    ReturnComponentValue wCaption, sTarget, sParts(I)
    If Len(sParts(I)) Then sParts(I) = sTarget & sParts(I) & "|"
Next
Select Case Property
Case lv_Caption:
    sTarget = ""
    sNewCaption = newValue
Case lv_ImgID
    sTarget = "IMG:"
Case lv_Bold
    sTarget = ""
    If Len(sProp) = 0 Then sProp = "False"
    If CBool(sProp) Then sParts(Property) = "Default|" Else sParts(Property) = ""
Case lv_Tip:
    sTarget = "Tip:"
Case lv_ListBoxID
    sParts(Property + 1) = ""
    sTarget = "LB:"
Case lv_ComboxID
    sParts(Property - 1) = ""
    sTarget = "CB:"
Case lv_ShowIconBkg
    sTarget = ""
    If Len(sProp) = 0 Then sProp = "False"
    If CBool(sProp) Then sParts(Property) = "ImgBkg|" Else sParts(Property) = ""
Case lv_HotKey
    sTarget = "HotKey:"
Case lv_FilesPath
    If sProp = "" Then
        sParts(Property) = ""
        sTarget = ""
    Else
        sParts(Property) = sProp
        sTarget = "Files:"
    End If
Case Else
    If Len(wCaption) Then wCaption = Left$(wCaption, Len(wCaption) - 1) & "}"
    ChangeMenuCaption = sNewCaption & wCaption
    Exit Function
End Select
If Len(sTarget) Then
    If Len(sProp) Then sParts(Property) = sTarget & sProp & "|" Else sParts(Property) = ""
End If
wCaption = ""
For I = 1 To UBound(sParts)
    wCaption = wCaption & sParts(I)
Next
If Len(wCaption) Then wCaption = "{" & Left$(wCaption, Len(wCaption) - 1) & "}"
ChangeMenuCaption = sNewCaption & wCaption
Erase sParts
End Function

Public Function CreateLvColors(CaptionNow As String, Optional UserID As String, _
    Optional CheckedColor As Long = -1) As String
' =====================================================================
' This function will create the custom menu lvColors
' =====================================================================
CreateLvColors = BuildSimpleCustomMenu(CaptionNow, "lvColors:", UserID, CStr(CheckedColor))

End Function

Public Function CreateLvDrives(CaptionNow As String, Optional UserID As String, _
    Optional CheckedDrive As String = "-1") As String
' =====================================================================
' This function will create the custom menu lvDays
' =====================================================================
CreateLvDrives = BuildSimpleCustomMenu(CaptionNow, "lvDrives:", UserID, CStr(CheckedDrive))

End Function

Public Function CreateLvDaysOfWeek(CaptionNow As String, Optional UserID As String, _
    Optional CheckedDay As Integer = -1) As String
' =====================================================================
' This function will create the custom menu lvDays
' =====================================================================
CreateLvDaysOfWeek = BuildSimpleCustomMenu(CaptionNow, "lvDays:", UserID, CStr(CheckedDay))

End Function

Public Function CreateLvStates(CaptionNow As String, Optional UserID As String, _
    Optional CheckedState As String = "-1") As String
' =====================================================================
' This function will create the custom menu lvStates
' =====================================================================
CreateLvStates = BuildSimpleCustomMenu(CaptionNow, "lvStates:", UserID, CheckedState)

End Function

Public Function CreateLvMonths(CaptionNow As String, Optional UserID As String, _
     Optional CheckedMonth As Integer = -1, Optional Grouping As CstmMonth) As String
' =====================================================================
' This function will create the custom menu lvMonths
' =====================================================================
Dim newCaption As String, wCaption As String
Dim sValue As String, sCode As String
Dim ArrayIndex As Integer, I As Integer

sCode = "lvMonths:" & CheckedMonth
Select Case Grouping
Case lv_cCalendarQuarter
    sCode = sCode & ":Group:CYQtr"
Case lv_cFiscalQuarter
    sCode = sCode & ":Group:FYQtr"
Case Else
    sCode = sCode & ":Group:Default"
End Select
If Len(UserID) Then sCode = sCode & ":ID:" & UserID
SeparateCaption CaptionNow, newCaption, wCaption
ReturnComponentValue wCaption, "lvMonths:", sValue
If Len(sValue) Then
    wCaption = Trim(Replace$(wCaption, "lvMonths:" & sValue & "|", ""))
    wCaption = Trim(Replace$(wCaption, "lvMonths:" & sValue, ""))
End If
If Len(wCaption) Then
    wCaption = Replace$(wCaption, "{", "")
    wCaption = Replace$(wCaption, "}", "")
    If Left$(wCaption, 1) <> "|" Then wCaption = "|" & wCaption
    wCaption = "{" & sCode & wCaption & "}"
Else
    wCaption = "{" & sCode & "}"
End If
CreateLvMonths = newCaption & Replace$(wCaption, "||", "|")

End Function

Public Function CreateLvDaysOfMonth(CaptionNow As String, Optional UserID As String, _
    Optional Year As Integer = 0, Optional Month As Integer = 0, _
    Optional CheckedDate As Integer = -1) As String
' =====================================================================
' This function will create the custom menu lvMonth
' =====================================================================
Dim newCaption As String, wCaption As String, sValue As String, sCode As String
sCode = "lvMonth:" & Month
sCode = sCode & ":Year:" & Year
sCode = sCode & ":Day:" & CheckedDate
If Len(UserID) Then sCode = sCode & ":ID:" & UserID
SeparateCaption CaptionNow, newCaption, wCaption
ReturnComponentValue wCaption, "lvMonth:", sValue
If Len(sValue) Then
    wCaption = Trim(Replace$(wCaption, "lvMonth:" & sValue & "|", ""))
    wCaption = Trim(Replace$(wCaption, "lvMonth:" & sValue, ""))
End If
If Len(wCaption) Then
    wCaption = Replace$(wCaption, "{", "")
    wCaption = Replace$(wCaption, "}", "")
    If Left$(wCaption, 1) <> "|" Then wCaption = "|" & wCaption
    wCaption = "{" & sCode & wCaption & "}"
Else
    wCaption = "{" & sCode & "}"
End If
CreateLvDaysOfMonth = newCaption & Replace$(wCaption, "||", "|")
End Function

Public Function CreateLvFonts(CaptionNow As String, Optional UserID As String, _
    Optional CheckedFont As String = "-1", Optional FontType As FontTypeEnum, _
    Optional FilterLetterA As String, Optional FilterLetterZ As String) As String
' =====================================================================
' This function will create the custom menu lvFonts
' =====================================================================
Dim newCaption As String, wCaption As String, sValue As String, sCode As String
sCode = "lvFonts:" & CheckedFont
sCode = sCode & ":Type:" & Choose(FontType + 1, "All", "TrueType", "System")
If FilterLetterA <> "" And FilterLetterZ <> "" Then
    sCode = sCode & ":Group:" & FilterLetterA & "-" & FilterLetterZ
End If
If Len(UserID) Then sCode = sCode & ":ID:" & UserID
SeparateCaption CaptionNow, newCaption, wCaption
ReturnComponentValue wCaption, "lvFonts:", sValue
If Len(sValue) Then
    wCaption = Trim(Replace$(wCaption, "lvFonts:" & sValue & "|", ""))
    wCaption = Trim(Replace$(wCaption, "lvFonts:" & sValue, ""))
End If
If Len(wCaption) Then
    wCaption = Replace$(wCaption, "{", "")
    wCaption = Replace$(wCaption, "}", "")
    If Left$(wCaption, 1) <> "|" Then wCaption = "|" & wCaption
    wCaption = "{" & sCode & wCaption & "}"
Else
    wCaption = "{" & sCode & "}"
End If
CreateLvFonts = newCaption & Replace$(wCaption, "||", "|")
End Function

Public Function ChangeCustomMenu(CaptionNow As String, _
    Optional NewUserID As String, Optional NewCheckedItem As Variant, _
    Optional NewGrouping As String, Optional NewFontType As FontTypeEnum = -1) As String

Dim sValue As String, wCaption As String, bRecognized As Boolean
Dim oldCaption As String, oldWcaption As String, chkType As String

SeparateCaption CaptionNow, oldCaption, wCaption
Dim sType As String, I As Integer, J As Integer
For I = 1 To 7
    sType = Choose(I, "lvFonts:", "lvColors:", "lvMonths:", "lvMonth:", "lvDrives:", "lvDays:", "lvStates:")
    If InStr(wCaption, sType) Then
        I = InStr(wCaption, sType)
        J = InStr(I, wCaption, "|")
        oldWcaption = Replace$(wCaption, Mid$(wCaption, I, J - I), "")
        wCaption = Mid$(wCaption, I, J - 1)
        oldWcaption = Replace$(oldWcaption, "{", "")
        If Right$(oldWcaption, 1) = "|" Then oldWcaption = Left$(oldWcaption, Len(oldWcaption) - 1)
        If Right$(wCaption, 1) <> "|" Then wCaption = wCaption & "|"
        wCaption = "{" & wCaption
        bRecognized = True
        Exit For
    End If
Next
If Not bRecognized Then Exit Function
If Not IsMissing(NewCheckedItem) Then
    If Len(CStr(NewCheckedItem)) Then
        If sType = "lvMonth:" Then chkType = "Day:" Else chkType = sType
        ReturnComponentValue wCaption, chkType, sValue
        J = InStr(sValue, ":")
        If J Then
            If Mid$(sValue, J + 1, 1) = "\" And sType = "lvDrives:" Then
                J = InStr(J + 1, sValue, ":")
            End If
        End If
        If J = 0 Then
            J = InStr(sValue, "|")
            If J = 0 Then J = Len(sValue) + 1
        End If
        wCaption = Replace$(wCaption, chkType & Left$(sValue, J - 1), chkType & CStr(NewCheckedItem))
    End If
End If
I = InStr(wCaption, sType)
I = InStr(I, wCaption, "|")
If Len(NewUserID) Then
    ReturnComponentValue wCaption, "ID:", sValue
    If sValue = "" Then
        wCaption = Left$(wCaption, Len(wCaption) - 1) & ":ID:" & NewUserID & "|"
    Else
        wCaption = Replace$(wCaption, ":ID:" & sValue, ":ID:" & NewUserID)
    End If
End If
If Len(NewGrouping) Then
    ReturnComponentValue wCaption, "Group:", sValue
    If sValue = "" Then
        wCaption = Left$(wCaption, Len(wCaption) - 1) & ":Group:" & NewGrouping & "|"
    Else
        wCaption = Replace$(wCaption, ":Group:" & sValue, ":Group:" & NewGrouping)
    End If
End If
If NewFontType > lv_fAllFonts - 1 And NewFontType < lv_fNonTrueType + 1 Then
    ReturnComponentValue wCaption, "Type:", sValue
    If sValue = "" Then
        wCaption = Left$(wCaption, Len(wCaption) - 1) & ":Type:" & Choose(NewFontType + 1, "ALL", "TrueType", "System") & "|"
    Else
        wCaption = Replace$(wCaption, ":Type:" & sValue, ":Type:" & Choose(NewFontType + 1, "ALL", "TrueType", "System"))
    End If
End If
ChangeCustomMenu = Replace$(oldCaption & wCaption & oldWcaption, "||", "|") & "}"
'Debug.Print CaptionNow: Debug.Print ChangeCustomMenu
End Function

Private Function BuildSimpleCustomMenu(CaptionNow As String, lvType As String, _
    Optional UserID As String, Optional CheckedItem As String) As String
' =====================================================================
' This function creates a few of the basic custom menus
' =====================================================================
Dim newCaption As String, wCaption As String, sValue As String, sCode As String
sCode = lvType & CheckedItem
If Len(UserID) Then sCode = sCode & ":ID:" & UserID
SeparateCaption CaptionNow, newCaption, wCaption
ReturnComponentValue wCaption, lvType, sValue
If Len(sValue) Then
    wCaption = Trim(Replace$(wCaption, lvType & sValue & "|", ""))
    wCaption = Trim(Replace$(wCaption, lvType & sValue, ""))
End If
If Len(wCaption) Then
    wCaption = Replace$(wCaption, "{", "")
    wCaption = Replace$(wCaption, "}", "")
    If Left$(wCaption, 1) <> "|" Then wCaption = "|" & wCaption
    wCaption = "{" & sCode & wCaption & "}"
Else
    wCaption = "{" & sCode & "}"
End If
BuildSimpleCustomMenu = newCaption & Replace$(wCaption, "||", "|")
End Function

Public Sub SetMenu(hWnd As Long, Optional ImageList As Control = Nothing, _
                    Optional TipClass As cTips = Nothing, Optional ContainerType As SubClassContainers = 0)
' =====================================================================
' Primary function to start subclassing forms and drawing their menus
' MDI Forms, when calling this function, flag this program to automatically subclass any of their children
' Forms are automatically unsubclassed when they are closed

' <parameters>
'   hWnd = handle to the form calling this function. MDI children should not call this function
'   ImageList = the image list object to use for menu icons. Can be an image list of any form
'   TipClass = the initialized cTips class within the form calling this function.
'       -- although optional. Tips will not be forwarded to the form without this parameter being called
' =====================================================================

' if user set the following flag, bypass any subclassing
If bAmDebugging Then Exit Sub

If colMenuItems Is Nothing Then
    If Not bModuleInitialized Then LoadDefaultColors
    ' primary collection used to reference subclassed form's cMenuItems class
    Set colMenuItems = New Collection
    ' for Win98/ME, systems add extra pixels to menu widths. We account for that.
    DetermineOS
    ' create the 3 primary fonts to use for menus 1:system menu font, 2:same font bolded, 3:mini font for separator bars
    CreateDestroyMenuFont True, False
End If
bUseHourglass = False

On Error Resume Next
Dim cMenu As cMenuItems, cHwnd As Long, sHwnd As Long, pTips As Long, targetHwnd As Long
Dim lFlags As Long
If Abs(CLng(ContainerType)) = lv_VB_Toolbar Then
    targetHwnd = GetToolTipWindow(hWnd)
Else
    targetHwnd = hWnd
End If
If targetHwnd = 0 Then Exit Sub
' simple test to see if we have already subclassed this form
Set cMenu = colMenuItems("h" & targetHwnd)
If Err = 0 Then
    Set cMenu = Nothing
    Select Case Abs(ContainerType)
    Case lv_MDIchildForm_NoMenus
        colMenuItems("h" & targetHwnd).IsMenuLess = True
    Case lv_MDIform_ChildrenHaveMenus, lv_MDIform_ChildrenMenuless
        If FindWindowEx(targetHwnd, 0, "MDIClient", "") Then
            colMenuItems("h" & targetHwnd).IsMenuLess = (ContainerType = lv_MDIform_ChildrenMenuless)
        End If
    Case lv_MDIchildForm_WithMenus
        colMenuItems("h" & targetHwnd).IsMenuLess = False
    End Select
    hWndRedirect = "h0"
    Exit Sub
End If
On Error GoTo 0
' if the user passed a Tips class, we reference its location now
If Not TipClass Is Nothing Then pTips = ObjPtr(TipClass)
' subclass the hWnd and add it to our collection
Set cMenu = New cMenuItems
cMenu.IsMenuLess = (ContainerType = lv_MDIchildForm_NoMenus Or ContainerType = lv_MDIform_ChildrenMenuless)
colMenuItems.Add cMenu, "h" & targetHwnd
Set cMenu = Nothing
' intialize the class with handle, image handle,  tips option and Win98/ME offsets if applicable
colMenuItems(colMenuItems.Count).InitializeSubMenu targetHwnd, ImageList, pTips, False
' start the subclassing now, saving the previous windows procedure at same time
colMenuItems(colMenuItems.Count).hPrevProc = SetWindowLong(targetHwnd, GWL_WNDPROC, AddressOf MenuMessages)
' here we are checking for MDI windows
cHwnd = FindWindowEx(targetHwnd, 0, "MDIClient", "")
If cHwnd Then
    ' this form is a MDI form, therefore, we need to subclass the MDI Client
    ' which contains the MDI children. By subclassing that window,
    ' we automatically get notified of children being created/destroyed
    Debug.Print "Found mdi client window of "; cHwnd; " parent is "; hWnd
    ' we don't need to test for a previous subclass on these -- it can only happen once
    Set cMenu = New cMenuItems
    colMenuItems.Add cMenu, "h" & cHwnd
    Set cMenu = Nothing
    ' for MDI clients, we won't get menu items, so we just set the image list and tips option to NULL
    colMenuItems(colMenuItems.Count).InitializeSubMenu cHwnd, , , True
    ' subclass the MDI Client now
    colMenuItems(colMenuItems.Count).hPrevProc = SetWindowLong(cHwnd, GWL_WNDPROC, AddressOf MenuMessages)
    ' identify window to MDI Parent, this way we cache whether or not the current form is a MDI or not.
    colMenuItems("h" & hWnd).MDIClient = cHwnd
    colMenuItems("h" & cHwnd).IsMDIclient = True
End If
If hWndRedirect = "MDIchildToSubclass" And cHwnd = 0 Then
    ' The hWndRedirect value is set in the MenuMessages routine when a  MDI Child is created
    Debug.Print "Auto-subclassed MDI child "; hWnd
    cHwnd = GetParent(GetParent(hWnd))
    ' here we subclass the MDI child, passing along the same ImageList object and TipsClass that its parent uses
    colMenuItems("h" & hWnd).InitializeSubMenu hWnd, colMenuItems("h" & cHwnd).ImageListObject, colMenuItems("h" & cHwnd).ShowTips
    If ContainerType < lv_MDIchildForm_WithMenus Then
        colMenuItems("h" & hWnd).IsMenuLess = colMenuItems("h" & cHwnd).IsMenuLess
    End If
End If
hWndRedirect = "h0"
End Sub

Public Sub SetPopupParentForm(hWnd As Long)
' =====================================================================
' Since a form can call any other form's menu as a popup, this function should be called prior to
' calling VB's PopupMenu command. By passing the handle of the form that owns the menu about
' to be popped up, the program can correctly identify any controls that may exist in those menu
' captions. This call is only used for the current popup.  Once the popup closes, this value is erased.

' Typically you should precede all popup commands with the function unless the form owning
' the menu is calling its own menu/submenu.  Example:
'       SetPopupParentForm MenuOwnerForm.hWnd
'       MenuPopup MenuOwnerForm.mnuName, , , , mnuDefaultItem

' Also see PopupMenuCustom for displaying a custom menu as its own popup
' =====================================================================
tempRedirect = hWnd
End Sub

Public Sub PopupMenuCustom(MenuFormsHwnd As Long, CustomCaption As String, _
    Optional Flags As Long, Optional X As Long = -1, Optional Y As Long = -1, _
    Optional TipsReRoute As cTips = Nothing)
' ========================================================================
' The same parameters are used as VB's PopupMenu with the following exceptions
' MenuFormsHwnd: VB would know this for its own menus, I need to be told
' CustomCaption: This caption is the lvColors, lvMonths, etc
'                This caption could also display LB: & CB: type menus
' Flags are same as VB's PopupMenu
' X & Y are same as VB's PopupMenu
' ========================================================================
' Custom menus are menus on the fly that are added to a non-submenu item.
' The program converts that non-submenu item to a submenu in order to display
' the custom menu. Well, VB don't like that. It will not recognize that the
' menu item is now a submenu. Therefore, you cannot call PopupMenu on that
' menu item even though it now has a submenu. You will consistently get the
' error that the menu needs at least one submenu item

' The downside with custom generated menus is that they cannot be popups
' unless they reside on an existing panel. In other words, if you wanted to
' display the lvColors custom menu as a popup, you would have to create a
' separate menu item and then create a submenu item under it formatted with
' the custom flag. Then you can call the parent menu and what do you get?
' a single submenu item that says something like "Colors" with the lvColors
' menu items as yet another submenu item to that "Colors" submenu. Crappy!

' Well, here's where we fling a custom menu on the fly without the need for
' it to be attached to any menu items. It is created just off a caption you
' pass to this routine.
' ========================================================================

Dim newMenu As Long, lReturn As Long, lPopup As Long
Dim menuPT As POINTAPI, pRect As RECT
Dim sPopupCaption As String
' we figure out where the menu will appear
If X < 0 Or Y < 0 Then  ' if either value is less < 0, we use cursor position
    GetCursorPos menuPT
Else
    menuPT.X = X: menuPT.Y = Y
End If
' create a new, blank popupmenu
newMenu = CreatePopupMenu()
If newMenu = 0 Then Exit Sub
If Not TipsReRoute Is Nothing Then RerouteTips MenuFormsHwnd, TipsReRoute
' now we ensure the program knows which form the custom menu belongs to
' Note: only truly necessary if the custom menu references a control by name
' (i.e., LB:list1, CB:combo1, IMG:picture1, etc)
SetPopupParentForm MenuFormsHwnd
' here we ensure the popup passed has some type of displayable caption even though it won't be
' displayed--it's submenu will be displayed. If we allow the passed caption to go through without this
' check, then if the displayed caption is blank, it is converted to a separator bar and separator bars
' cannot have submenus -- Windows prevents it.
sPopupCaption = "PopMenu" & CustomCaption
' now we add only one item to this new menu item which is the custom menu caption
AppendMenu newMenu, MF_DISABLED, 32500, sPopupCaption
' we process the above item so it will create the custom menus on the fly
' and those custom menus become a submenu of our newMenu
colMenuItems("h" & MenuFormsHwnd).IsWindowList newMenu, False
' we want to ensure user didn't accidentally provide this flag. If provided,
' we wouldn't get the WM_MEASUREITEM & WM_DRAWITEM messages -- ouch!
If ((Flags And TPM_NONOTIFY) = TPM_NONOTIFY) Then Flags = Flags And Not TPM_NONOTIFY
' we don't want the popup menu to send the result to the form's Message Processor
' we want it to report back to us.
Flags = Flags Or TPM_RETURNCMD
' call the API to display the submenu of our newMenu
lPopup = GetSubMenu(newMenu, 0)
lReturn = TrackPopupMenu(lPopup, Flags, menuPT.X, menuPT.Y, 0&, MenuFormsHwnd, pRect)
If lReturn Then
    ' user selected an item from the menu, so we let the class handle it &
    ' it will send the result to the cTips class associated with MenuFormsHwnd
    ' or update the list/combobox if it was that type of custom menu
    colMenuItems("h" & MenuFormsHwnd).MenuSelected lReturn, lPopup, 0
End If
' now regardless, we want to destroy the 2 menus we created
' the first function below will clear the class arrays for the custom submenu
colMenuItems("h" & MenuFormsHwnd).DestroyPopup newMenu, 32500
' and now we clear the newMenu 'cause it isn't attached to any form
' menu & it will stay in memory forever
DestroyMenu newMenu
' note the actual custom submenu of newMenu is destroyed automatically when
' newMenu is destroyed, otherwise we would have to kill that one too
' To prove it unrem this next line
'Debug.Print "popup menu destroyed successfully? "; CBool(IsMenu(lPopup) = 0)
DoTips 0, 0, 0
If Not TipsReRoute Is Nothing Then RerouteTips MenuFormsHwnd
End Sub

Public Sub RerouteTips(hWnd As Long, Optional TipsClass As cTips = Nothing)
' =====================================================================
' By default, menu tips of MDI children are passed to their parent.
' This is more efficient for the user.
' However, a user may want the menu item tips of MDI children to be sent
' to the child instead of the parent.
' By calling this procedure, the menu tips are sent to which ever class
' is provided in above parameters. If no class is provided, then the
' cTips reference for the hWnd parameter is reset to its previous value

' NOTE: By redirecting the Tips, you are also redirecting any custom menus
' that may exist. See cTips for more information on custom menus
' =====================================================================
Dim lStyle As Long

' do we have this item subclassed yet? A simple test
On Error Resume Next
lStyle = colMenuItems("h" & hWnd).ShowTips
If Err Then   ' this hasn't been subclassed yet, let's do it now
              ' first we'll test to see if this is a MDI child
    lStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
    If ((lStyle And WS_EX_MDICHILD) = WS_EX_MDICHILD) Then
        ' we have a MDI child that hasn't been subclassed yet. Not a normal
        ' situation, but can happen when the RerouteTips command is placed
        ' in the child's form load event and the child is opened by a call
        ' to FrmName.Show vs calling it's parent form. Got it?
        ' Set below variable so child can be subclassed using the parent's
        ' settings as a starting point
        On Error Resume Next
        hWndRedirect = "MDIchildToSubclass"
        SetMenu hWnd
        ' now if tips are being rerouted, reroute them here
        If Not TipsClass Is Nothing Then colMenuItems("h" & hWnd).InitializeSubMenu hWnd, , ObjPtr(TipsClass)
    Else
        ' non-MDI child being subclassed
        SetMenu hWnd, , TipsClass
    End If
    Err.Clear
Else  ' already subclassed, we'll just toggle the routing now
      ' passing the pointer to the cTips class toggles the value within the cMenuItems class
      ' if the pointer is zero, then the main TipsClass is used. That is the one
      ' when the form was first subclassed. If non-zero, the class stores the
      ' previous TipsClass and uses the passed parameter
    If TipsClass Is Nothing Then lStyle = 0 Else lStyle = ObjPtr(TipsClass)
    colMenuItems("h" & hWnd).InitializeSubMenu hWnd, , lStyle, True
End If
End Sub

Public Sub CleanClass(hWnd As Long)
' =====================================================================
' DO NOT CALL THIS ROUTINE FROM YOUR APPLICATION.
' Although it is public, it is only made public to be used with the optional ctrlMenuSample included
' with the original package you downloaded from PSC.

' Otherwise, this routine is called each time a subclassed form is closed.
' Its main purpose is simple:  release subclassing and release memory objects
' =====================================================================
Dim sysHwnd As Long
' clear collection of currently visible menus. Should already be cleared, just a sanity check
Set OpenMenus = Nothing
' restore the windows message processor to its original address
SetWindowLong hWnd, GWL_WNDPROC, colMenuItems("h" & hWnd).hPrevProc
' get the handle to the form's system menu (if any) before we destroy the class
sysHwnd = colMenuItems("h" & hWnd).SystemMenu
' for use with the optional user control. Put menus back to their original state
If AmInIDE Then colMenuItems("h" & hWnd).RestoreMenus
If colMenuItems.Count = 1 Then
    ' if this is the last form to be un-subclassed, we also....
    CreateDestroyMenuFont False, False          ' destroy all memory fonts
    Set colMenuItems = Nothing                  ' clear the main collection
    If FloppyIcon Then DestroyIcon FloppyIcon
    FloppyIcon = 0
    Erase tbarClass
    Debug.Print "All classes completely cleaned, no more subclassing"
Else    ' not the last form, so we simply clear only this form's class from the collection
    colMenuItems.Remove "h" & hWnd
End If
' here we restore the system menu if needed
On Error Resume Next
If sysHwnd Then GetSystemMenu sysHwnd, 1
End Sub

Public Sub CreateDestroyMenuFont(bCreate As Boolean, ItalicFonts As Boolean, Optional FontSample As String)
' DO NOT CALL THIS ROUTINE FROM YOUR APPLICATION.
' =====================================================================
' Simply create or destroy memory fonts. Up to 7 fonts created
'(1)=Normal menu caption font
'(2)=Bolded menu caption font
'(3)=Mini-font used on separator bars (hardcoded here as Tahoma 7.5
'(4)=Italicized normal menu font used for optional ItalicizeSelectedItems property
'(5)=Italicized bolded menu font used for optional ItalicizeSelectedItems property
'(6)=Normal sample font used for the custom Fonts menu
'(7)=Italicized sample font used for the custom Fonts menu
'(0)=original hDC's font before one of the above were selected into the hDC
' =====================================================================
If bCreate = True Then
    ' in order to set the font, we must first determine what it is
    Dim ncm As NONCLIENTMETRICS, newFont As LOGFONT, oldWT As Long
    ncm.cbSize = Len(ncm)
    ' this will return the system menu font info
    SystemParametersInfo 41, 0, ncm, 0
    newFont = ncm.lfMenuFont
    newFont.lfCharSet = 1
    newFont.lfFaceName = MenuFontName & Chr$(0)
    newFont.lfHeight = (mFontSize * -20) / Screen.TwipsPerPixelY
    ' here we create memory fonts based off of  the system menu font or user-defined font
    If ItalicFonts Then
        ' optional italicize on highlight. But font only needs to be created once
        If m_Font(4) = 0 Then
            newFont.lfItalic = 1                                        ' italics attribute
            m_Font(4) = CreateFontIndirect(newFont)        ' create the font
            newFont.lfWeight = 800                                 ' bold version
            m_Font(5) = CreateFontIndirect(newFont)        ' create the font
        End If
    Else
        If Len(FontSample) Then
            ' here the custom fonts menu is preparing to draw/measure the Font passed
            ' first delete any pre-existing sample fonts if needed
            If m_Font(6) Then DeleteObject m_Font(6)
            If m_Font(7) Then DeleteObject m_Font(7)
            ' now we create a normal and italicized version of the sample font requested
            newFont.lfWeight = 400                                      ' normal font
            newFont.lfFaceName = FontSample & Chr$(0)
            m_Font(6) = CreateFontIndirect(newFont)
            newFont.lfItalic = 1                                             ' italicized version
            m_Font(7) = CreateFontIndirect(newFont)
        Else
            ' this is where the 3 basic fonts are created for majority of the menus
            m_Font(1) = CreateFontIndirect(newFont)     ' normal font
            oldWT = newFont.lfWeight
            newFont.lfWeight = 800                              ' bold version
            m_Font(2) = CreateFontIndirect(newFont)
            newFont.lfWeight = oldWT                          ' restore the original boldness/weight
            ' now we are going to try to create a scalable font for
            ' separator bar text based off of the user's selected font
            ' We are also trying to make separator bars with text look good
            ' By default we'll use a fontsize 80% the size of the menu font & if that size is less than 7.5 pts,
            ' we'll default to 1/2 pt size less than the menu font
            If (mFontSize * 0.8) < 7.5 Then
                newFont.lfHeight = ((mFontSize - 0.5) * -20) / Screen.TwipsPerPixelY
            Else
                newFont.lfHeight = ((mFontSize * 0.8) * -20) / Screen.TwipsPerPixelY
            End If
            m_Font(3) = CreateFontIndirect(newFont)     ' mini-font
        End If
    End If
Else
    ' deleting all fonts
    On Error Resume Next
    DeleteObject m_Font(1)
    DeleteObject m_Font(2)
    DeleteObject m_Font(3)
    If m_Font(4) Then DeleteObject m_Font(4)
    If m_Font(5) Then DeleteObject m_Font(5)
    If m_Font(6) Then DeleteObject m_Font(6)
    If m_Font(7) Then DeleteObject m_Font(7)
    Erase m_Font
End If
End Sub

Public Sub ApplyMenuFont(FontID As Integer, hDC As Long)
' DO NOT CALL THIS ROUTINE FROM YOUR APPLICATION.
' =====================================================================
'  This either sets or replaces fonts in a DC
' Calling routine calls this function twice, once to set a font temporarily & again to restore original
' =====================================================================
If hDC = 0 Then Exit Sub
If FontID Then
    m_Font(0) = SelectObject(hDC, m_Font(FontID))
Else
    SelectObject hDC, m_Font(0)
End If
End Sub

Private Sub DoTips(wParam As Long, lParam As Long, hMenu As Long)
Dim hWord As Integer, tipProc As Long
Dim lWord As Long, sTip As String
Dim bMenuClosed As Boolean
' =====================================================================
' Routine will send the menu tip, if any, to the cTips class for the
' active form. Routine called from the Message Processor
' =====================================================================
' get the reference to the form's cTips class
tipProc = colMenuItems(hWndRedirect).ShowTips

On Error Resume Next
If hMenu Then
    If wParam <> 0 Or lParam <> 0 Then
        ' submenus, what a pain!
        lWord = CLng(LoWord(wParam)) ' the menu ID hidden in the LoWord of wParam,
        hWord = HiWord(wParam)       ' the menu status hidden in HiWord
        ' however,if the menu has submenus, then LoWord is the Index of the menu
        ' item and not the handle which we need to retrieve the tips; therefore,
        ' we need to get the handle by calling GetSubMenu
        If ((hWord And MF_POPUP) = MF_POPUP) Then lWord = GetSubMenu(lParam, lWord)
        sTip = colMenuItems(hWndRedirect).Tips(lWord, hMenu)
    End If
Else
    If wParam = 0 And lParam = 0 Then bMenuClosed = True
End If
If tipProc = 0 Then Exit Sub

' here we create another instance of the user's cTips class
' call the function we need & then kill the extra instance
Dim oTipClass As cTips
CopyMemory oTipClass, tipProc, 4&
oTipClass.SendTip sTip
If bMenuClosed Then oTipClass.SendCustomSelection "", "MenusClosed", 0&
CopyMemory oTipClass, 0&, 4&
Set oTipClass = Nothing
End Sub

Private Function DoMeasureItem(lParam As Long) As Boolean
' =====================================================================
' Called whenever a menu item is about to be initially displayed
' Only called once, so several routines in the cMenuItems class
' had to work around & trick windows to remeasure on an as needed basis
' =====================================================================
    
Dim MeasureInfo As MEASUREITEMSTRUCT
'Get the MEASUREITEM info, basically submenu item height/width
Call CopyMemory(MeasureInfo, ByVal lParam, Len(MeasureInfo))
' only process menu items, controls can send above message
' and we don't want to interfere with those.
If MeasureInfo.CtlType <> ODT_MENU Then Exit Function
colMenuItems(hWndRedirect).GetMenuItem MeasureInfo.ItemID, MeasureInfo.ItemData
'Tell Windows how big our items are.
MeasureInfo.ItemHeight = XferMenuData.Dimension.Y
MeasureInfo.ItemWidth = XferMenuData.Dimension.X + XferMenuData.OffsetCx
'Return the information back to Windows
Call CopyMemory(ByVal lParam, MeasureInfo, Len(MeasureInfo))
'Debug.Print "measured "; XferMenuData.Display
DoMeasureItem = True
End Function

Private Function DoDrawItem(lParam As Long) As Boolean
Dim DrawInfo As DRAWITEMSTRUCT
' TODO: continue updating remarks in this routine as it continues
' to get more complicated with the inclusion of more options
' =====================================================================
' Note: Any the "ConvertColor" functions below are unnecessary at this
' point in time. I've coded them in for the next version which would
' allow user selected colors, which then may need to be converted
' =====================================================================
 
'On Error Resume Next
On Error GoTo ShowErrors

Dim IsSep As Boolean, IsDisabled As Boolean
Dim sysIcon As Boolean, bSelectDisabled As Boolean
Dim Xoffset As Integer, yIconOffset As Integer
Dim IsSelected As Boolean, IsChecked As Boolean
Dim bGradientFill As Boolean, sFont As String, iFont As Integer
Dim lTextColor As Long, cBack As Long
Dim tRect As RECT, mItem As RECT

Dim sMsg As String
'Get DRAWINFOSTRUCT which gives us sizes & indexes
Call CopyMemory(DrawInfo, ByVal lParam, LenB(DrawInfo))
' only process menu items, controls can send above message
' and we don't want to interfere with those.
If DrawInfo.CtlType <> ODT_MENU Then Exit Function
DoDrawItem = True
sMsg = "retrieving menu info"
' return the menu item information and panel information
colMenuItems(hWndRedirect).GetMenuItem DrawInfo.ItemID, DrawInfo.ItemData
sMsg = "retrieving menu/panel info"
'If XferMenuData.ID <> DrawInfo.ItemData Then
colMenuItems(hWndRedirect).GetPanelItem DrawInfo.hWndItem
With DrawInfo
    ' get some basic attributes first
    sMsg = "getting attributes"
    IsSep = ((XferMenuData.Status And lv_mSep) = lv_mSep)
    IsDisabled = ((XferMenuData.Status And lv_mDisabled) = lv_mDisabled)
    If (.itemAction And Not ODA_DRAWENTIRE) And IsSep = True Then Exit Function
    IsChecked = ((XferMenuData.Status And lv_mChk) = lv_mChk)
    IsSelected = ((DrawInfo.itemState And ODS_SELECTED) = ODS_SELECTED)
    bSelectDisabled = (bHiLiteDisabled = True) Or (XferPanelData.IsSystem = True) Or (bKeyBoardSelect = True)
    If ((XferMenuData.Status And lv_mSBar) = lv_mSBar) Then
        If XferPanelData.PanelIcon <> 0 Then
            If (.itemAction And ODA_DRAWENTIRE) Or IsDisabled = False Then
                Dim tDC As Long, sDC As Long, oldBMP As Long
                sDC = GetDC(CLng(Mid$(hWndRedirect, 2)))
                tDC = CreateCompatibleDC(sDC)
                'Debug.Print "sidebar draw size", .rcItem.Right - .rcItem.Left; .rcItem.Left; .rcItem.Right
                oldBMP = SelectObject(tDC, XferPanelData.PanelIcon)
                If IsSelected = True And IsDisabled = False Then
                    StretchBlt .hDC, 2, 2, .rcItem.Right - 4, .rcItem.Bottom - 4, _
                                tDC, 0, 0, XferMenuData.Dimension.X, XferMenuData.Dimension.Y, vbSrcCopy
                    ThreeDbox .hDC, .rcItem.Left + 1, .rcItem.Top + 1, .rcItem.Right - 1, .rcItem.Bottom - 1, False, False
                Else
                    BitBlt .hDC, 0, 0, _
                        .rcItem.Right, _
                        .rcItem.Bottom, tDC, 0, 0, vbSrcCopy
                End If
                SelectObject tDC, oldBMP
                DeleteDC tDC
                ReleaseDC CLng(Mid$(hWndRedirect, 2)), sDC
                DeleteObject oldBMP
            End If
        End If
        If Not IsDisabled Then
            ThreeDbox .hDC, .rcItem.Left, .rcItem.Top, .rcItem.Right, .rcItem.Bottom - 1, (IsSelected = True), False
        End If
       Exit Function
    End If
    
    sMsg = "converting colors"
    If IsSep = True Or IsSelected = False Then
        cBack = GetSysColor(COLOR_MENU)
    Else
        cBack = lSelectBColor
        'cBack = GetSysColor(COLOR_HIGHLIGHT)
    End If
    Select Case LoWord(.ItemID)     ' system menu
    Case SC_CLOSE, SC_MAXIMIZE, SC_MINIMIZE, SC_RESTORE
        sysIcon = True      ' we will manually draw these icons unless they are provided if a user
    Case Else                ' already modified the copy of the system menu
        sysIcon = False
    End Select
    mItem = .rcItem                                       ' working copy of the rectangle to draw in
    SetBkMode .hDC, NEWTRANSPARENT   ' make text print with transparent background
        
        mItem.Left = mItem.Left + 21    ' start highlighting background with this offset
        ' in the following cases we will highlight complete rectangle from left to right
        If ((Len(XferMenuData.Icon) = 0 And IsChecked = False) Or _
            XferPanelData.HasIcons = False) And sysIcon = False Then mItem.Left = .rcItem.Left
        If ((XferMenuData.Status And lv_mColor) = lv_mColor) Then
            If InStr(XferMenuData.Caption, "LColor:-1") = 0 Then mItem.Right = mItem.Right - 38
        End If
    If ((IsSep = False) And (IsDisabled = False)) Or (IsDisabled = True And bSelectDisabled = True) Then
    
        'Draw the highlighting rectangle
        sMsg = "drawing back rect"
        mItem.Bottom = mItem.Bottom - 1
        If bGradientSelect = False Or (IsSep = True Or IsSelected = False) Then
            DrawRect .hDC, mItem.Left, mItem.Top, mItem.Right, mItem.Bottom, cBack
        Else
            DrawGradient lSelectBColor, GetSysColor(COLOR_MENU), True, .hDC, mItem
            bGradientFill = True
        End If
    End If
    If IsSep Then
        sMsg = "drawing separator bars"
        ' separators... 2 types -- those with text & those without
        ' regardless, we want full width of the menu panel
        mItem = .rcItem                        '  refresh working copy of the rectangle to draw in
        If Len(XferMenuData.Display) Then     ' separator bars with text
            tRect = .rcItem   ' tRect will hold coords for horizontal-centered text in panel
            OffsetRect tRect, 0, 1
            ApplyMenuFont 3, .hDC  ' load mini font & set font color
            SetMenuColor True, .hDC, cObj_Text, TextColorSeparatorBar
            ' we calculate the rectangle size needed to print the text
            DrawText .hDC, XferMenuData.Display, Len(XferMenuData.Display), tRect, DT_CALCRECT Or DT_SINGLELINE Or DT_NOCLIP Or DT_CENTER Or DT_VCENTER
            OffsetRect tRect, (mItem.Right - tRect.Right) \ 2, 0     ' now we center the text in the panel & draw it
            DrawText .hDC, XferMenuData.Display, Len(XferMenuData.Display), tRect, DT_SINGLELINE Or DT_NOCLIP Or DT_VCENTER
            ' here we add the separator lines on both sides of the separator caption
            ThreeDbox .hDC, mItem.Left, _
                (mItem.Bottom - mItem.Top) \ 2 + mItem.Top + 1, _
                tRect.Left - 3, (mItem.Bottom - mItem.Top) \ 2 + mItem.Top + 1, _
                True, ((XferMenuData.Status And lv_mSepRaised) = lv_mSepRaised), True
            ThreeDbox .hDC, tRect.Right + 3, _
                (mItem.Bottom - mItem.Top) \ 2 + mItem.Top + 1, _
                mItem.Right, (mItem.Bottom - mItem.Top) \ 2 + mItem.Top + 1, _
                True, ((XferMenuData.Status And lv_mSepRaised) = lv_mSepRaised), True
        Else    ' standard separator bar
            ApplyMenuFont 1, .hDC     ' simply draw a separator bar from left to right
             ThreeDbox .hDC, mItem.Left + 2, (mItem.Bottom - mItem.Top) \ 2 + mItem.Top + 1, _
                mItem.Right - 2, (mItem.Bottom - mItem.Top) \ 2 + mItem.Top + 1, _
                True, ((XferMenuData.Status And lv_mSepRaised) = lv_mSepRaised), True
        End If
        ApplyMenuFont 0, .hDC
    Else    ' normal caption, not a separator bar
        mItem = .rcItem    '  refresh working copy of the rectangle to draw in
        If Len(XferMenuData.Icon) > 0 Or (IsChecked And XferPanelData.HasIcons = True) Then
            ' in above cases, we draw a lighter background rectangle to place icon/checkmark on
            If IsChecked = True And IsDisabled = False And (.itemAction And ODA_DRAWENTIRE Or bRaisedIcons = True) Then
                sMsg = "drawing icon back button                "
                DrawRect .hDC, mItem.Left + 1, mItem.Top + 1, 19 + mItem.Left, mItem.Bottom - 2, CheckedIconBColor
                yIconOffset = -1
            Else
                If bRaisedIcons Then
                    DrawRect .hDC, mItem.Left + 1, mItem.Top + 1, 19 + mItem.Left, mItem.Bottom - 2, GetSysColor(COLOR_MENU)
                End If
            End If
            If (IsDisabled = True And IsSelected = False) Or IsDisabled = False Or _
                (IsDisabled = True And (bHiLiteDisabled = True Or bKeyBoardSelect = True)) Then
            ' in above cases we draw a raised/sunken box around the icon/checkmark
                sMsg = "drawing icon border "
                ThreeDbox .hDC, mItem.Left, mItem.Top, 19 + mItem.Left, mItem.Bottom - 2, _
                        (IsSelected = True Or IsChecked = True), (IsChecked = True And IsSelected = False)
            End If
        End If
        If IsChecked = True And (Len(XferMenuData.Icon) = 0 And sysIcon = False) Then
            mItem = .rcItem    '  refresh working copy of the rectangle to draw in
            ' Draw checkmarks. Because of all the options, checkmarks can be displayed
            ' in a total of 4 different color variations
            ' we'll take the case of when the panel has icons first
            If XferPanelData.HasIcons Or XferPanelData.IsSystem Then
                If IsDisabled Then
                    OffsetRect mItem, 1, 1
                    DrawCheckMark .hDC, mItem, .rcItem.Left, bXPcheckmarks, TextColorDisabledLight, TextColorDisabledDark
                Else
                    DrawCheckMark .hDC, mItem, .rcItem.Left, bXPcheckmarks, TextColorNormal
                End If
            Else    ' no icons on the panel, we can have 3 different combinations
                If IsDisabled Then
                    If bSelectDisabled And IsSelected Then
                        DrawCheckMark .hDC, mItem, .rcItem.Left, bXPcheckmarks, TextColorDisabledDark
                    Else
                        OffsetRect mItem, 1, 1
                        DrawCheckMark .hDC, mItem, .rcItem.Left, bXPcheckmarks, TextColorDisabledLight, TextColorDisabledDark
                    End If
                Else
                    If IsSelected Then
                        DrawCheckMark .hDC, mItem, .rcItem.Left, bXPcheckmarks, TextColorSelected
                    Else
                        DrawCheckMark .hDC, mItem, .rcItem.Left, bXPcheckmarks, TextColorNormal
                    End If
                End If
            End If
        End If
        '  refresh working copy of the rectangle to draw in and offset it 3 pixels left of position where icon would end
        mItem = .rcItem                  '  refresh working copy of the rectangle to draw in
        mItem.Left = .rcItem.Left + 23
        iFont = 1 + Abs(CInt((XferMenuData.Status And lv_mDefault) = lv_mDefault))
        ApplyMenuFont iFont, .hDC  ' standard or bold font
        ' draw the item once & again later if it is disabled
        If IsDisabled Then
            If bSelectDisabled And IsSelected Then
                lTextColor = TextColorDisabledDark
            Else
                lTextColor = TextColorDisabledLight
                OffsetRect mItem, 1, 1
            End If
        Else
            If IsSelected = True Then
                lTextColor = TextColorSelected
                ApplyMenuFont 0, .hDC
                iFont = 1 + Abs(CInt((XferMenuData.Status And lv_mDefault) = lv_mDefault)) + (Abs(CInt(bItalicSelected)) * 3)
                ApplyMenuFont iFont, .hDC ' standard or bold font & italcized
            Else
                lTextColor = TextColorNormal
            End If
        End If
        SetMenuColor True, .hDC, cObj_Text, lTextColor
        DrawText .hDC, XferMenuData.Display, Len(XferMenuData.Display), mItem, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER
        If Len(XferMenuData.HotKey) Then
            mItem.Right = .rcItem.Right - 15
            If ((XferMenuData.Status And lv_mFont) = lv_mFont) Then
                ReturnComponentValue XferMenuData.Caption, "LFont:", sFont
                ApplyMenuFont 0, .hDC
                CreateDestroyMenuFont True, False, sFont
                ApplyMenuFont (Abs(CInt(bItalicSelected = True And IsSelected = True))) + 6, .hDC
            End If
            DrawText .hDC, XferMenuData.HotKey, Len(XferMenuData.HotKey), mItem, DT_SINGLELINE Or DT_RIGHT Or DT_NOCLIP Or DT_VCENTER
        End If
        If (IsDisabled = True And bSelectDisabled = False) = True Or _
            (IsDisabled = True And IsSelected = False) Then
            sMsg = "Drawing disabled text"
            mItem.Right = .rcItem.Right
            OffsetRect mItem, -1, -1  ' reset for next hot key drawing
            SetMenuColor True, .hDC, cObj_Text, TextColorDisabledDark
            If ((XferMenuData.Status And lv_mFont) = lv_mFont) Then
                ApplyMenuFont 0, .hDC
                ApplyMenuFont iFont, .hDC
            End If
            DrawText .hDC, XferMenuData.Display, Len(XferMenuData.Display), mItem, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER
            If Len(XferMenuData.HotKey) Then
                mItem.Right = .rcItem.Right - 16
                If ((XferMenuData.Status And lv_mFont) = lv_mFont) Then
                    ApplyMenuFont 0, .hDC
                    ApplyMenuFont (Abs(CInt(bItalicSelected = True And IsSelected = True))) + 6, .hDC
                End If
                DrawText .hDC, XferMenuData.HotKey, Len(XferMenuData.HotKey), mItem, DT_SINGLELINE Or DT_RIGHT Or DT_NOCLIP Or DT_VCENTER
            End If
        End If
        
        ApplyMenuFont 0, .hDC  ' replace menu default font
        ' last step - draw the icons!
        If Len(XferMenuData.Icon) > 0 Or sysIcon = True Then
            sMsg = "drawing icons"
            mItem = .rcItem                  '  refresh working copy of the rectangle to draw in
            If Len(XferMenuData.Icon) Then             ' if an icon was provided we use that one
                mItem.Left = mItem.Left + 2
                mItem.Top = mItem.Top + ((mItem.Bottom - mItem.Top) - 16) \ 2
                If IsDisabled = False And IsSelected = True And bRaisedIcons = True Then
                    DrawMenuIcon .hDC, XferMenuData.Icon, 0, mItem, True, False, Abs(CInt(XferMenuData.ShowBKG)), 0, 0
                    DrawMenuIcon .hDC, XferMenuData.Icon, 0, mItem, False, , Abs(CInt(XferMenuData.ShowBKG)), -1, -1
                Else
                    If (.itemAction = ODA_DRAWENTIRE) Or bRaisedIcons = True Then
                        DrawMenuIcon .hDC, XferMenuData.Icon, 0, mItem, IsDisabled, , Abs(CInt(XferMenuData.ShowBKG)), , yIconOffset
                    End If
                End If
            Else    ' no icon, which means this is system menu by default, call function to manually draw those icons
                DrawSystemIcon .hDC, mItem, .ItemID, IsDisabled, IsSelected
            End If
        End If
        If (.itemAction And ODA_DRAWENTIRE) Then
            If ((XferMenuData.Status And lv_mColor) = lv_mColor) Then
                With XferMenuData
                    Dim sColor As String
                    Xoffset = InStr(.Caption, "{")
                    ReturnComponentValue Mid$(.Caption, Xoffset), "LColor:", sColor
                End With
                If Val(sColor) <> -1 Then DrawRect .hDC, .rcItem.Right - 35, .rcItem.Top + 2, .rcItem.Right - 5, .rcItem.Bottom - 2, Val(sColor), vbBlack
            End If
        End If
    End If
End With
ShowErrors:
If Err Then Debug.Print Err.Description; " <<DoDrawItem "; sMsg, DrawInfo.hWndItem
End Function

Public Function LoWord(DWord As Long) As Integer
' DO NOT CALL THIS ROUTINE FROM YOUR APPLICATION.
' =====================================================================
' function to return the LoWord of a Long value
' =====================================================================
     If DWord And &H8000& Then
        LoWord = DWord Or &HFFFF0000
     Else
        LoWord = DWord And &HFFFF&
     End If
End Function
Public Function HiWord(DWord As Long) As Integer
' DO NOT CALL THIS ROUTINE FROM YOUR APPLICATION.
' =====================================================================
' function to return the HiWord of a Long value
' =====================================================================
     HiWord = (DWord And &HFFFF0000) \ &H10000
End Function
Public Function MakeLong(ByVal LoWord As Integer, ByVal HiWord As Integer) As Long
  MakeLong = CLng(LoWord)
  Call CopyMemory(ByVal VarPtr(MakeLong) + 2, HiWord, 2)
End Function

Private Function MenuMessages(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
' TODO: Remove any instances of sMsg from project -- troubleshooting purposes only
' =====================================================================
' The main Message Processor for all forms. It is pretty straight forward
' The only question you may have is why is the message being sent to the
' form before it is processed in some cases and in some cases it sent
' aftewards or not at all.  Well, good question. The answer: Trial and
' error. Where the message is sent before processing seems to be required.
' Sending it afterwards or not at all produced bad effects & errors.
' =====================================================================
Dim sMsg As String
On Error GoTo AllowMsgThru
Static LastHmenu As Long       ' used to keep track of which submenu is active (Not all WM_messages pass this)
Static CursorType As Integer

Select Case uMsg
Case WM_KEYDOWN, WM_KEYUP, WM_SYSKEYDOWN, WM_SYSKEYUP
    ' trapped to return keystrokes to a MDI Parent when no children are open
    MenuMessages = CallWindowProc(colMenuItems("h" & hWnd).hPrevProc, hWnd, uMsg, wParam, lParam)
    If OpenMenus Is Nothing And bReturnMDIkeystrokes = True Then
        ProcessKeyStroke hWnd, wParam, ((uMsg = WM_KEYUP) Or (uMsg = WM_SYSKEYUP))
        Exit Function
    End If
Case WM_GETMINMAXINFO
    ' used when window is resizing if user set this via the SetMinMaxInfo function
    If colMenuItems("h" & hWnd).RestrictSize(lParam, False) = True Then Exit Function
Case WM_ENTERMENULOOP           ' starting a new menu session
    If Screen.MousePointer <> vbHourglass Then CursorType = Screen.MousePointer + 1
    sMsg = "WM_Entermenuloop"
    Set OpenMenus = New Collection    ' create new index of opened menus
Case WM_MDIACTIVATE
    ' depending on whether the hWnd is the MDIclient or the MDIchild...
    sMsg = "WM_MdiActivate"
    If hWnd = GetParent(wParam) Then
        ' new MDI child been created, send message thru first, then hook the child
        MenuMessages = CallWindowProc(colMenuItems("h" & hWnd).hPrevProc, hWnd, uMsg, wParam, lParam)
        hWndRedirect = "MDIchildToSubclass"
        SetMenu wParam
        Exit Function
    End If
Case WM_MDIMAXIMIZE
    ' strange. when a MDI child is loaded and maximized, it does not fire
    ' a WM_MDIACTIVATE message. To absolutely catch all of them as they
    ' appear so we can subclass them, we will trap this message too
    MenuMessages = CallWindowProc(colMenuItems("h" & hWnd).hPrevProc, hWnd, uMsg, wParam, lParam)
    hWndRedirect = "MDIchildToSubclass"
    SetMenu wParam
    Exit Function
Case WM_MENUCHAR
    ' WinNT4 won't process menu accelerator keys when menus are owner-drawn. Win98 & Win2K will process
    ' them and will not even come thru this routine. Don't know about Win95 or WinXP.
    MenuMessages = IdentifyAccelerator(LoWord(wParam), lParam)
    Exit Function
Case WM_INITMENUPOPUP       ' a menu is about to appear
    sMsg = "WM_INitmenupopup"
    ' gotta forward this along first, otherwise bad things could happen with MDI forms and system menus
    MenuMessages = CallWindowProc(colMenuItems("h" & hWnd).hPrevProc, hWnd, uMsg, wParam, lParam)
    ' if the HiWord of lParam is not zero, we have a system menu
    IDCurrentWindow hWnd, wParam, (HiWord(lParam) <> 0)
    'now let's see if the menu about to be displayed could be a little
    'slow. If so, we set the screen cursor to an hourglass
    colMenuItems(hWndRedirect).GetPanelItem wParam
    If XferPanelData.Hourglass = True And CursorType - 1 <> vbHourglass Then
         Screen.MousePointer = vbHourglass
         bUseHourglass = True
    Else
        bUseHourglass = False
    End If
    'Debug.Print "Opening menu "; wParam; " for window "; hWndRedirect; " ActiveHwnd h " & hWnd; " h " & tempRedirect; ""
    On Error Resume Next
    wParam = OpenMenus("m" & wParam)    ' if menu is already open, no need to mess with it
    If Err = 0 Then Exit Function
    On Error GoTo AllowMsgThru
    ' the only time we want to process a submenu each & every time is when it is a window list submenu
    ' because windows will take ownership back each & every time. In this case, the function below returns True
    If colMenuItems(hWndRedirect).IsWindowList(wParam, False) = False Then OpenMenus.Add wParam, "m" & wParam
    Exit Function
Case WM_MEASUREITEM         ' windows requests measurement of menu item
    If DoMeasureItem(lParam) = True Then Exit Function
Case WM_DRAWITEM            ' windows requests drawing of menu item
    If DoDrawItem(lParam) = True Then Exit Function
Case WM_MENUSELECT          ' user has highlighted/unhighlighted a menu
    sMsg = "WM_Menuselect"
    If lParam Then LastHmenu = lParam    ' keep track of which menu panel is in use
    DoTips wParam, lParam, LastHmenu
    bKeyBoardSelect = ((HiWord(wParam) And MF_MOUSESELECT) <> MF_MOUSESELECT)
Case WM_COMMAND             ' user has clicked a menu item
    sMsg = "WM_Command"
    'Debug.Print "received a wm_command: "; hWnd; wParam; lParam
    If HiWord(wParam) = 0 And lParam = 0 Then
       ' If lParam=0 and HiWord(wParam)=0, then from a menu
       If colMenuItems(hWndRedirect).MenuSelected(CLng(LoWord(wParam)), LastHmenu, 0) Then Exit Function  ' LoWord is menu ID
    End If
Case WM_MENUCOMMAND         ' Win98/ME, Win2K -- not on all O/S
    sMsg = "WM_menucommand"
    ' Debug.Print "received a "; sMsg; " from: "; hWnd; wParam; lParam
    ' for Win98/ME the HiWord is what we need, but to keep things kinda
    ' simple by processing it just like a WM_COMMAND windows message
    ' To do this we look at the LoWord for the menu zero-based index &
    ' lParam is the menu handle. That & the zero-based index get's the ID
    colMenuItems(hWndRedirect).MenuSelected GetMenuItemID(lParam, LoWord(wParam)), lParam, 0
Case WM_EXITMENULOOP        ' menu session has ended
    sMsg = "WM_exitmenuloop"
    ' if we forced an hourglass pointer, reset it now
    If bUseHourglass Then
        Screen.MousePointer = CursorType - 1
        bUseHourglass = False
    End If
    DoTips 0, 0, 0                ' clear any menu tips
    tempRedirect = 0
    Dim lMenus As Integer
    For lMenus = 1 To OpenMenus.Count
        colMenuItems(hWndRedirect).UpdateMenuItems OpenMenus.Item(lMenus)
    Next
    Set OpenMenus = Nothing       ' clear collection of opened menus
Case WM_DESTROY
    sMsg = "WM_Destroy"
    ' see http://support.microsoft.com/default.aspx?scid=kb;en-us;89738
    ' above has a different way of trying to ensure no sys crash when this
    ' flag is sent to windows. I have had good success here, but if your
    ' system crashes when you close a form or project, please let me know
    ' so I can look at alternatives. Of course if you use the END command
    ' in your form, it will always crash while form is subclassed
    On Error Resume Next
    MenuMessages = CallWindowProc(colMenuItems("h" & hWnd).hPrevProc, hWnd, uMsg, wParam, lParam)
    'Debug.Print "destroying "; hWnd
    CleanClass hWnd
    Exit Function
Case WM_MDIDESTROY
    sMsg = "WM_MDIdestroy"
    ' A MDI child form is closing. Same as WM_DESTROY above, but the
    ' hWnd passed here is the MDI Parent's client window, not the MDI child
    ' However, the wParam is the the MDI child and we remove the associated
    ' class of that handle. We use the hWnd to get back to the parent MDI
    ' form as the new focus object for our classes.
    On Error Resume Next
    MenuMessages = CallWindowProc(colMenuItems("h" & hWnd).hPrevProc, hWnd, uMsg, wParam, lParam)
    CleanClass wParam
    'Debug.Print "destroying MDI "; wParam
    Exit Function
Case WM_ENTERIDLE
    ' reset screen icon if we forced an hourglass
    If bUseHourglass Then
        Screen.MousePointer = CursorType - 1
        bUseHourglass = False
    End If
End Select

AllowMsgThru:
'If Err Then Debug.Print Err.Description; " <<MenuMessages in "; sMsg
On Error Resume Next
MenuMessages = CallWindowProc(colMenuItems("h" & hWnd).hPrevProc, hWnd, uMsg, wParam, lParam)
End Function

Private Function DetermineOS(Optional SetGraphicsModeDC As Long = 0) As Integer
' =====================================================================
' Reliable rotine used to determine the operating system

' With Win98 & ME, extra pixels (35 to be exact) can be added to a
' menu which has the effect of this project showing extra wide menus
' We identify the system & account for those extra pixels
' =====================================================================
Const os_Win95 = "1.4.0"
Const os_Win98 = "1.4.10"
Const os_WinNT4 = "2.4.0"
Const os_WinNT351 = "2.3.51"
Const os_Win2K = "2.5.0"
Const os_WinME = "1.4.90"
Const os_WinXP = "2.5.1"

  Dim verinfo As OSVERSIONINFO, sVersion As String
  verinfo.dwOSVersionInfoSize = Len(verinfo)
  If (GetVersionEx(verinfo)) = 0 Then Exit Function
  With verinfo
    sVersion = .dwPlatformId & "." & .dwMajorVersion & "." & .dwMinorVersion
  End With
  Select Case sVersion
  Case os_Win98: ExtraOffsetX = 35
  Case os_Win2K: ExtraOffsetX = 0
  Case os_WinNT4: ExtraOffsetX = 0
  Case os_WinNT351  ' unsure if this will work with NT v3.51
    ' per MSDN this setting is required to print vertical text
    ' if any of you try this on NT 3.51, let me know. Thanx.
    SetGraphicsMode SetGraphicsModeDC, 2
  Case os_Win95
  Case os_WinXP: ExtraOffsetX = 0
  Case os_WinME: ExtraOffsetX = 35
  End Select
End Function

Public Sub ReturnComponentValue(sSource As String, sTarget As String, sRtnVal As String)
' DO NOT CALL THIS ROUTINE FROM YOUR APPLICATION.
' =====================================================================
' Returns the value of the menu item option
' =====================================================================
sRtnVal = ""
If Len(sSource) < 3 Then Exit Sub

Dim cI As Integer, cJ As Integer, sSourceNoTip As String
' The coded part of the caption can appear in a couple different areas of caption

sSourceNoTip = sSource
' here we remove the tip from the coded part of the caption before parsing it
' This prevents inadvertent parsing of coded words that exist in the tip
cI = InStr(sSource, "|Tip:")                     ' most follow a pipe symbol i.e., ( |Bold|Italic etc)
If cI = 0 Then cI = InStr(sSource, "{Tip:")      ' others follow the left parent i.e., {IMG: )
If cI Then
    cI = cI + 1
    cJ = InStr(cI + 1, sSource, "|")                                'end location of option, or...
    If cJ = 0 Then cJ = InStr(cI + 1, sSource, "}")
    If cJ Then
        If sTarget <> "Tip:" Then sSourceNoTip = Replace$(sSource, Mid$(sSource, cI, cJ - cI), "")
    End If
End If
If sTarget <> "Tip:" Then
    cI = InStr(sSourceNoTip, "|" & sTarget)                     ' most follow a pipe symbol i.e., ( |Bold|Italic etc)
    If cI = 0 Then cI = InStr(sSourceNoTip, "{" & sTarget)      ' others follow the left parent i.e., {IMG: )
    If cI = 0 Then cI = InStr(sSourceNoTip, ":" & sTarget)      ' others follow the left parent i.e., {LvColors:-1:ID:abc )
    If cI = 0 Then Exit Sub
    cI = cI + 1
    cJ = InStr(cI + 1, sSourceNoTip, "|")                                'end location of option, or...
    If cJ = 0 Then cJ = InStr(cI + 1, sSourceNoTip, "}")
End If
' if we got this far & cJ>0, then the property exists in the caption
If cJ Then
    sRtnVal = Mid$(sSourceNoTip, cI, cJ - cI)                   ' complete option with Tag included
    sRtnVal = Mid$(sRtnVal, Len(sTarget) + 1) ' Trim$(Replace$(sRtnVal, sTarget, ""))   ' strip out the tag
    If Len(sRtnVal) = 0 Then
        ' some tags don't require a value, in this case we supply a non Null value if needed
        Select Case sRtnVal
            Case "vbAlignBottom": sRtnVal = "Bot"       ' convert for alignment options
            Case "vbAlignTop": sRtnVal = "Top"            ' convert for alignment options
            Case "vbNull": sRtnVal = "1"                       ' convert for non-color option
            Case Else
                ' When a boolean type property is parsed out, there won't be a value associated with it
                ' so we return TRUE as the default for the following properties
                Select Case sTarget
                Case "Raised", "Bold", "Italic", "Underline", "NoScroll", "Transparent", "ImgBkg", "Default", "Sidebar", "SBDisabled"
                    sRtnVal = "True"
                Case Else: sRtnVal = "0"
            End Select
        End Select
    End If
End If
End Sub

Public Sub SeparateCaption(sSource, sNewCaption, sCodeCaption, Optional sHotKey)
' DO NOT CALL THIS ROUTINE FROM YOUR APPLICATION.
Dim IdX As Integer, tSource As String
' =====================================================================
' since the coded part of the caption can be before or after the actual
' menu caption, when hotkeys also exist via Menu Editor, the coded part
' can actually be between the actual caption & hotkey, therefore we
' process the hotkey first so we can get rid of it and make parsing easier
' =====================================================================
tSource = sSource
sHotKey = ""
sNewCaption = ""
sCodeCaption = ""
IdX = InStr(tSource, vbTab)
If IdX Then     ' we have a Menu Editor shortcut, let's process it now
    sHotKey = Mid$(tSource, IdX + 1)    ' and update the caption
    If IdX > 1 Then tSource = Left$(tSource, IdX - 1) Else tSource = ""
End If
sNewCaption = Trim$(tSource) ' initial caption to display on menu
If Len(tSource) Then   ' either a text separator bar or normal menu item
    ' let's look for component flags & markers
    IdX = InStr(tSource, "{") ' these are surrounded by brackets { }
    If IdX Then
        ' got a left bracket, look for right one & if not, send thru as is
        If InStr(IdX, tSource, "}") = 0 Then Exit Sub
        ' remove the working caption and the display caption is left over
        sCodeCaption = Mid$(tSource, IdX, InStr(IdX, tSource, "}") - IdX + 1)
        sNewCaption = Trim$(Replace$(tSource, sCodeCaption, ""))
        ' we end working caption with a pipe to make parsing a bit easier
        sCodeCaption = Left$(sCodeCaption, Len(sCodeCaption) - 1) & "|"
        If AmInIDE Then
            Dim Idy As Integer
            IdX = InStr(sCodeCaption, "IMG:")
            If IdX Then Idy = InStr(IdX, sCodeCaption, "|")
            If Idy Then
                sCodeCaption = Replace$(sCodeCaption, Mid$(sCodeCaption, IdX, Idy - IdX), "IMG:" & DefaultIcon)
            End If
        End If
    End If
End If
End Sub

Private Function DrawMenuIcon(m_HDC As Long, sImgID As String, imageType As Long, _
    rt As RECT, bDisabled As Boolean, Optional bDisableColored As Boolean = True, _
    Optional bNoTransparency As Long = 0, Optional iOffset As Integer = 0, _
    Optional Yoffset As Integer, Optional ImgWidth As Integer = 16, _
    Optional ImgHeight As Integer = 16, Optional lMask As Long = -1) As Boolean
' DO NOT CALL THIS ROUTINE FROM YOUR APPLICATION.
' =====================================================================
' Routine will draw the menu image onto the menu
' =====================================================================

Dim tDC As Long, lPrevImage As Long, lImageType As Long
Dim lImgCopy As Long, lImageHdl As Long, sImgHandle As String
Dim bmpInfo As BITMAP, icoInfo As ICONINFO
Dim rcImage As RECT, dRect As RECT
Dim shInfo As SHFILEINFO
Dim exeIcon As Long

Const CI_BITMAP = &H0
Const CI_ICON = &H1

sImgHandle = sImgID
If IsNumeric(sImgHandle) Or Len(sImgHandle) = 0 Then
    lImageHdl = Val(sImgHandle)
Else
    ' the icon passed is not numeric, this means it is a file name and path where the icon is to be
    ' extracted for display on the menu. We use same icons as Explorer.
    ' Even shows shortcuts (.lnk files) with the "shortcut arrow" icon
    SHGetFileInfo sImgHandle, 0, shInfo, Len(shInfo), SHGFI_ICON Or SHGFI_SMALLICON Or SHGFI_USEFILEATTRIBUTES
    lImageHdl = shInfo.hIcon    ' must be deleted later
End If
If lImageHdl = 0 Then Exit Function
' see if the image to be displayed is an icon or bitmap
' not an icon, is it a bitmap-like image?  If not, we can't draw the picture
' we check for bitmaps first
GetObject lImageHdl, Len(bmpInfo), bmpInfo
If bmpInfo.bmBits Then
    lImageType = CI_BITMAP      ' flag indicating image is a bitmap
    lImgCopy = CopyImage(lImageHdl, CI_BITMAP, ImgWidth, ImgHeight, 0)
Else
    GetIconInfo lImageHdl, icoInfo
    If icoInfo.hbmColor <> 0 Then
        ' downside... API creates 2 bitmaps that we need to destroy since they aren't used in this
        ' routine & are not destroyed automatically. To prevent memory leak, we destroy them here
        lImageType = CI_ICON        ' flag indicating image is an icon
        DeleteObject icoInfo.hbmColor
        If icoInfo.hbmMask <> 0 Then DeleteObject icoInfo.hbmMask
    Else
        Exit Function
    End If
End If
' destination rectangle for drawing on the DC
dRect = rt
OffsetRect dRect, iOffset, Yoffset      ' move the rectangle if needed
DrawMenuIcon = True
If Not bDisabled Then
    ' if the image is an icon or a bitmap with a background color that won't be transparent, then...
    If lImageType = CI_ICON Then
        ' for icons, we can draw directly on the destination DC
        exeIcon = CopyImage(lImageHdl, CI_ICON, ImgWidth, ImgHeight, 0)
        DrawIconEx m_HDC, dRect.Left, dRect.Top, exeIcon, 0, 0, 0, 0, &H3
        DestroyIcon exeIcon
        If shInfo.hIcon Then DestroyIcon shInfo.hIcon
    Else
        If bNoTransparency = 1 Then
            ' for bitmaps, we will load it into a temp DC and blit it over to the destination DC
            ' while resizing it on the way over
            tDC = CreateCompatibleDC(m_HDC)             ' will be destroyed at end of routine
            lPrevImage = SelectObject(tDC, lImgCopy)   ' will be destroyed at end of routine
            StretchBlt m_HDC, dRect.Left, dRect.Top, ImgWidth, ImgHeight, tDC, 0, 0, bmpInfo.bmWidth, bmpInfo.bmHeight, vbSrcCopy
        Else
            ' a bitmap where background is to be transparent, then we call function to do that, passing
            ' the destination DC, destination rectangle and new image size
            DrawTransparentBitmap m_HDC, dRect, lImgCopy, rcImage, , CLng(ImgWidth), CLng(ImgHeight)
        End If
        DeleteObject lImgCopy
    End If
    If IsNumeric(sImgID) = False Then DestroyIcon lImageHdl
Else    ' image is to be disabled

    Const MAGICROP = &HB8074A
    Dim hBitmap As Long, hOldBitmap As Long
    Dim hMemDC As Long
    Dim hOldBrush As Long
    Dim hOldBackColor As Long, hbrShadow As Long, hbrHilite As Long
    
    ' Create a temporary DC and bitmap to hold the image
    hMemDC = CreateCompatibleDC(m_HDC)
    hBitmap = CreateCompatibleBitmap(m_HDC, ImgWidth, ImgHeight)
    hOldBitmap = SelectObject(hMemDC, hBitmap)
    PatBlt hMemDC, 0, 0, ImgWidth, ImgHeight, WHITENESS
    
    dRect = rt
    If lImageType = CI_ICON Then
        ' draw icon directly onto the temporary DC
        ' for icons, we can draw directly on the destination DC
        exeIcon = CopyImage(lImageHdl, CI_ICON, ImgWidth, ImgHeight, 0)
        DrawIconEx hMemDC, 0, 0, exeIcon, 0, 0, 0, 0, &H3
        DestroyIcon exeIcon
        If shInfo.hIcon Then DestroyIcon shInfo.hIcon
    Else
        If bNoTransparency = 1 Then
            ' blit bitmap onto the temporary DC
            tDC = CreateCompatibleDC(m_HDC)
            lPrevImage = SelectObject(tDC, lImgCopy)
            StretchBlt hMemDC, 0, 0, ImgWidth, ImgHeight, tDC, 0, 0, bmpInfo.bmWidth, bmpInfo.bmHeight, vbSrcCopy
        Else
            OffsetRect dRect, rt.Left * -1, rt.Top * -1
            ' draw transparent bitmap onto the tempoary DC
            DrawTransparentBitmap hMemDC, dRect, lImgCopy, rcImage, , CLng(ImgWidth), CLng(ImgHeight)
            dRect = rt
        End If
        DeleteObject lImgCopy
    End If
    If IsNumeric(sImgID) = False Then DestroyIcon lImageHdl
  'OffsetRect dRect, iOffset, Yoffset
  hOldBackColor = SetBkColor(m_HDC, vbWhite)
  hbrShadow = CreateSolidBrush(ConvertColor(GetSysColor(COLOR_BTNSHADOW)))

  If bDisableColored Then
    hbrHilite = CreateSolidBrush(ConvertColor(GetSysColor(COLOR_BTNHIGHLIGHT)))
    hOldBrush = SelectObject(m_HDC, hbrHilite)
    BitBlt m_HDC, dRect.Left + 1, dRect.Top + 1, ImgWidth, ImgHeight, hMemDC, 0, 0, MAGICROP
    SelectObject m_HDC, hbrShadow
    BitBlt m_HDC, dRect.Left, dRect.Top, ImgWidth, ImgHeight, hMemDC, 0, 0, MAGICROP
  Else
    SelectObject m_HDC, hbrShadow
    BitBlt m_HDC, dRect.Left + 1, dRect.Top + 1, ImgWidth, ImgHeight, hMemDC, 0, 0, MAGICROP
    hbrHilite = CreateSolidBrush(hOldBackColor)
    SelectObject m_HDC, hbrHilite
    BitBlt m_HDC, dRect.Left, dRect.Top, ImgWidth, ImgHeight, hMemDC, 0, 0, MAGICROP
  End If
  
  SelectObject m_HDC, hOldBrush
  SetBkColor m_HDC, hOldBackColor
  SelectObject hMemDC, hOldBitmap
  DeleteObject hbrHilite
  DeleteObject hbrShadow
  DeleteObject hBitmap
  DeleteDC hMemDC
End If

If lPrevImage Then SelectObject tDC, lPrevImage
If tDC Then DeleteDC tDC
End Function

Public Sub DrawTransparentBitmap(lHDCdest As Long, destRect As RECT, _
                                                    lBMPsource As Long, bmpRect As RECT, _
                                                    Optional lMaskColor As Long = -1, _
                                                    Optional lNewBmpCx As Long, _
                                                    Optional lNewBmpCy As Long, _
                                                    Optional lBkgHDC As Long, _
                                                    Optional bkgX As Long, _
                                                    Optional bkgY As Long, _
                                                    Optional FlipHorz As Boolean = False, _
                                                    Optional FlipVert As Boolean = False, _
                                                    Optional srcDC As Long)
Const DSna = &H220326 '0x00220326
' DO NOT CALL THIS ROUTINE FROM YOUR APPLICATION.
' =====================================================================
' A pretty good transparent bitmap maker I use in several projects
' =====================================================================

' Above parameters are described...
' lHDCdest is the DC where the drawing will take place
' destRect is a RECT type indicating the left, top, right & bottom coords where drawing will be done
' lBMPsource is the handle to the bitmap to be made transparent and be re-drawn on lHDCdest
' bmpRect is a Rect type indicating the bitmaps coords to use for drawing
'   -- Note: If not provided, the entire bitmap is used.
' lMaskColor is the bitmap color to be made transparent. The value of -1 picks the top left corner pixel
' lNewBmpCx is the destination width of the source bitmap
'  -- Note: If not provided, the bitmap width is drawn with a 1:1 ratio
' lNewBmpCy is the destination height of the source bitmap
' -- Note: If not provided, the bitmap height is drawn with a 1:1 ratio
' ************ Following parameters are used if a separate HDC is used as a background or mask
'                 to be used for drawing. This option is used primarily as a background for animation
' lBkgHDC is the DC of the background image container
' bkgX, bkgYare the upper left/top coords to use on the background/mask DC for drawing on the
'   the destination DC. The width and height are determined by destRect's overall width/height

'-----------------------------------------------------------------
    Dim udtBitMap As BITMAP
    Dim lMask2Use As Long 'COLORREF
    Dim lBmMask As Long, lBmAndMem As Long, lBmColor As Long
    Dim lBmObjectOld As Long, lBmMemOld As Long, lBmColorOld As Long
    Dim lHDCMem As Long, lHDCscreen As Long, lHDCsrc As Long, lHDCMask As Long, lHDCcolor As Long
    Dim OrientX As Long, OrientY As Long
    Dim X As Long, Y As Long, srcX As Long, srcY As Long
    Dim lRatio(0 To 1) As Single
'-----------------------------------------------------------------
    Dim hPalOld As Long, hPalMem As Long
    
    lHDCscreen = GetDC(0&)
    lHDCsrc = CreateCompatibleDC(lHDCscreen)     'Create a temporary HDC compatible to the Destination HDC
    SelectObject lHDCsrc, lBMPsource             'Select the bitmap
    GetObject lBMPsource, Len(udtBitMap), udtBitMap

    ' Bmp size needed for original source
        srcX = udtBitMap.bmWidth                  'Get width of bitmap
        srcY = udtBitMap.bmHeight                 'Get height of bitmap
        If lNewBmpCx = 0 Then
            If bmpRect.Right > 0 Then lNewBmpCx = bmpRect.Right - bmpRect.Left Else lNewBmpCx = srcX
        End If
        'Use passed width and height parameters if provided
        If lNewBmpCy = 0 Then
            If bmpRect.Bottom > 0 Then lNewBmpCy = bmpRect.Bottom - bmpRect.Top Else lNewBmpCy = srcY
        End If
        
        If bmpRect.Right = 0 Then bmpRect.Right = srcX Else srcX = bmpRect.Right - bmpRect.Left
        If bmpRect.Bottom = 0 Then bmpRect.Bottom = srcY Else srcY = bmpRect.Bottom - bmpRect.Top
    ' Calculate size needed for drawing
        If (destRect.Right) = 0 Then X = lNewBmpCx Else X = (destRect.Right - destRect.Left)
        If (destRect.Bottom) = 0 Then Y = lNewBmpCy Else Y = (destRect.Bottom - destRect.Top)
'=========================================================================
' This routine will fail to draw properly if you try to draw a  larger dimension (lNewBmpCX or lNewBmpCy
' that is larger than the destination dimensions. Therefore, if the source dimensions are larger, then
' the routine will attempt to automatically scale the source image as needed.
'=========================================================================
        If lNewBmpCx > X Or lNewBmpCy > Y Then
            lRatio(0) = (X / lNewBmpCx)
            lRatio(1) = (Y / lNewBmpCy)
            If lRatio(1) < lRatio(0) Then lRatio(0) = lRatio(1)
            lNewBmpCx = lRatio(0) * lNewBmpCx
            lNewBmpCy = lRatio(0) * lNewBmpCy
            Erase lRatio
        End If
    
    lMask2Use = lMaskColor
    If lMask2Use < 0 Then lMask2Use = GetPixel(lHDCsrc, 0, 0)
    lMask2Use = ConvertColor(lMask2Use)
    'OleTranslateColor lMask2Use, 0, lMask2Use
    
    'Create some DCs to hold temporary data
    lHDCMask = CreateCompatibleDC(lHDCscreen)
    lHDCMem = CreateCompatibleDC(lHDCscreen)
    lHDCcolor = CreateCompatibleDC(lHDCscreen)
    'Create a bitmap for each DC.  DCs are required for a number of GDI functions
    'Compatible DC's
    lBmColor = CreateCompatibleBitmap(lHDCscreen, srcX, srcY)
    lBmAndMem = CreateCompatibleBitmap(lHDCscreen, X, Y)
    lBmMask = CreateBitmap(srcX, srcY, 1&, 1&, ByVal 0&)
    
    'Each DC must select a bitmap object to store pixel data.
    lBmColorOld = SelectObject(lHDCcolor, lBmColor)
    lBmMemOld = SelectObject(lHDCMem, lBmAndMem)
    lBmObjectOld = SelectObject(lHDCMask, lBmMask)
    
    ReleaseDC 0&, lHDCscreen
    
' ====================== Start working here ======================
    
    SetMapMode lHDCMem, GetMapMode(lHDCdest)
    hPalMem = SelectPalette(lHDCMem, 0, True)
    RealizePalette lHDCMem
    'Copy the background of the main DC to the destination
    If (lBkgHDC <> 0) Then
            BitBlt lHDCMem, 0, 0, X, Y, lBkgHDC, bkgX, bkgY, vbSrcCopy
    Else
            BitBlt lHDCMem, 0&, 0&, X, Y, lHDCdest, destRect.Left, destRect.Top, vbSrcCopy
    End If
    
    'Set proper mapping mode.
    hPalOld = SelectPalette(lHDCcolor, 0, True)
    RealizePalette lHDCcolor
    SetBkColor lHDCcolor, GetBkColor(lHDCsrc)
    SetTextColor lHDCcolor, GetTextColor(lHDCsrc)
    
    ' Get working copy of source bitmap
    'StretchBlt lHDCcolor, srcX, 0, -srcX, srcY, lHDCsrc, bmpRect.Left, bmpRect.Top, srcX, srcY, vbSrcCopy
    BitBlt lHDCcolor, 0&, 0&, srcX, srcY, lHDCsrc, bmpRect.Left, bmpRect.Top, vbSrcCopy
    If FlipHorz Then StretchBlt lHDCcolor, srcX, 0, -srcX, srcY, lHDCcolor, 0, 0, srcX, srcY, vbSrcCopy
    If FlipVert Then StretchBlt lHDCcolor, 0, srcY, srcX, -srcY, lHDCcolor, 0, 0, srcX, srcY, vbSrcCopy
    ' set working color back/fore colors. These colors will help create the mask
    SetBkColor lHDCcolor, lMask2Use
    SetTextColor lHDCcolor, vbWhite
    
    'Create the object mask for the bitmap by performaing a BitBlt
    BitBlt lHDCMask, 0&, 0&, srcX, srcY, lHDCcolor, 0&, 0&, vbSrcCopy
    
    ' This will create a mask of the source color
    SetTextColor lHDCcolor, vbBlack
    SetBkColor lHDCcolor, vbWhite
    BitBlt lHDCcolor, 0, 0, srcX, srcY, lHDCMask, 0, 0, DSna

    'Mask out the places where the bitmap will be placed while resizing as needed
    StretchBlt lHDCMem, 0, 0, lNewBmpCx, lNewBmpCy, lHDCMask, 0&, 0&, srcX, srcY, vbSrcAnd
    
    'XOR the bitmap with the background on the destination DC while resizing as needed
    StretchBlt lHDCMem, 0&, 0&, lNewBmpCx, lNewBmpCy, lHDCcolor, 0, 0, srcX, srcY, vbSrcPaint
    
    'Copy to the destination
    BitBlt lHDCdest, destRect.Left, destRect.Top, X, Y, lHDCMem, 0&, 0&, vbSrcCopy
    'StretchBlt lHDCdest, destRect.Left, destRect.Top, X, Y, lHDCMem, 0, 0, X, Y, vbSrcCopy
    
    
    'Delete memory bitmaps
    DeleteObject SelectObject(lHDCcolor, lBmColorOld)
    DeleteObject SelectObject(lHDCMask, lBmObjectOld)
    DeleteObject SelectObject(lHDCMem, lBmMemOld)
    
    'Delete memory DC's
    DeleteDC lHDCMem
    DeleteDC lHDCMask
    DeleteDC lHDCcolor
    If srcDC = 0 Then DeleteDC lHDCsrc
    
    'If hWndToRefresh Then InvalidateRect hWndToRefresh, destRect, 0
'-----------------------------------------------------------------
End Sub
'-----------------------------------------------------------------

Public Sub DrawRect(m_HDC As Long, ByVal X1 As Long, ByVal Y1 As Long, _
                                   ByVal X2 As Long, ByVal Y2 As Long, _
                                   tColor As Long, Optional pColor As Long = -1)
' DO NOT CALL THIS ROUTINE FROM YOUR APPLICATION.
' =====================================================================
' Simple routine to draw a rectangle & replace HDC objects when done
' =====================================================================
If pColor <> -1 Then SetMenuColor True, m_HDC, cObj_Pen, pColor
SetMenuColor True, m_HDC, cObj_Brush, tColor, (pColor = -1)
Call Rectangle(m_HDC, X1, Y1, X2, Y2)
SetMenuColor False, m_HDC, cObj_Brush, 0
End Sub

Private Sub ThreeDbox(tHdc As Long, ByVal X1 As Long, ByVal Y1 As Long, _
ByVal X2 As Long, ByVal Y2 As Long, bSelected As Boolean, _
Optional Sunken As Boolean = False, Optional bSeparatorBar As Boolean, _
Optional PenWidthLt As Long = 1, Optional PenWidthDk As Long = 1, _
Optional ColorLt As Long = -1, Optional ColorDk As Long = -1)
' =====================================================================
'   Draw/erase a raised/sunken box around the specified coordinates.
' =====================================================================
     
 If tHdc = 0 Then Exit Sub

 Dim dm As POINTAPI, iOffset As Integer
 
If ColorLt = -1 Then ColorLt = SeparatorBarColorLight
If ColorDk = -1 Then ColorDk = SeparatorBarColorDark
 ' select colors, offset when set indicates erasing
 iOffset = Abs(CInt(bSelected)) + 1

 If Sunken = False Then
    SetMenuColor True, tHdc, cObj_Pen, Choose(iOffset, GetSysColor(COLOR_MENU), ColorLt), , PenWidthLt
 Else
    SetMenuColor True, tHdc, cObj_Pen, Choose(iOffset, GetSysColor(COLOR_MENU), ColorDk), , PenWidthDk
 End If
 
 'First - Light Line
If bSeparatorBar Then
    MoveToEx tHdc, X1, Y1 - (CInt(Sunken) + 1), dm
    LineTo tHdc, X2 - 1, Y1 - (CInt(Sunken) + 1)
Else
    MoveToEx tHdc, X1, Y2, dm
    LineTo tHdc, X1, Y1
    LineTo tHdc, X2, Y1
End If
 SetMenuColor False, tHdc, cObj_Pen, 0
 ' Next - Dark line
 If Sunken = False Then
    SetMenuColor True, tHdc, cObj_Pen, Choose(iOffset, GetSysColor(COLOR_MENU), ColorDk), , PenWidthDk
 Else
    SetMenuColor True, tHdc, cObj_Pen, Choose(iOffset, GetSysColor(COLOR_MENU), ColorLt), , PenWidthLt
 End If
 
If bSeparatorBar = True Then
 MoveToEx tHdc, X1, Y2 - 1 - (CInt(Sunken) + 1), dm
 LineTo tHdc, X2 - 1, Y2 - 1 - (CInt(Sunken) + 1)
Else
 LineTo tHdc, X2, Y2
 LineTo tHdc, X1, Y2
End If
SetMenuColor False, tHdc, cObj_Pen, 0
End Sub

Private Sub DrawSystemIcon(m_HDC As Long, mItem As RECT, ItemID As Long, IsDisabled As Boolean, IsSelected As Boolean)
Dim X As Long, Y As Long, Xoffset As Long, Yoffset As Long, tPt As POINTAPI
Dim Looper As Integer, PenType As Integer
' =====================================================================
' Since this version of the menus program can successfully add color &
' subclass the system menu, we don't have any icons we it is subclassed
' This routine draws the icons in either enabled or disabled form
' =====================================================================

Xoffset = mItem.Left + 5
Yoffset = mItem.Top + ((mItem.Bottom - mItem.Top) - 8) \ 2 - 1
Select Case LoWord(ItemID)
Case SC_CLOSE       ' the Close command
    If IsDisabled Then
        PenType = 1: GoSub GetPen
        MoveToEx m_HDC, Xoffset + 2, Yoffset + 0, tPt
        LineTo m_HDC, Xoffset + 5, Yoffset + 3
        MoveToEx m_HDC, Xoffset + 9, Yoffset + 0, tPt
        For Looper = 1 To 4
            LineTo m_HDC, Xoffset + Choose(Looper, 9, 6, 9, 9), Yoffset + Choose(Looper, 1, 4, 7, 8)
        Next
        MoveToEx m_HDC, Xoffset + 1, Yoffset + 9, tPt
        LineTo m_HDC, Xoffset + 5, Yoffset + 5
        PenType = 0: GoSub GetPen
    End If
    PenType = 2: GoSub GetPen
    MoveToEx m_HDC, Xoffset + 0, Yoffset + 0, tPt
    LineTo m_HDC, Xoffset + 9, Yoffset + 9
    MoveToEx m_HDC, Xoffset + 8, Yoffset + 0, tPt
    LineTo m_HDC, Xoffset + -1, Yoffset + 9
    MoveToEx m_HDC, Xoffset + 0, Yoffset + 1, tPt
    LineTo m_HDC, Xoffset + 8, Yoffset + 9
    MoveToEx m_HDC, Xoffset + 1, Yoffset + 0, tPt
    LineTo m_HDC, Xoffset + 9, Yoffset + 8
    MoveToEx m_HDC, Xoffset + 7, Yoffset + 0, tPt
    LineTo m_HDC, Xoffset + -1, Yoffset + 8
    MoveToEx m_HDC, Xoffset + 8, Yoffset + 1, tPt
    LineTo m_HDC, Xoffset + 0, Yoffset + 9
    PenType = 0: GoSub GetPen
Case SC_MAXIMIZE    ' the Maximize command
    If IsDisabled Then
        PenType = 1: GoSub GetPen
        MoveToEx m_HDC, Xoffset + 10, Yoffset + 1, tPt
        For Looper = 1 To 4
            LineTo m_HDC, Xoffset + Choose(Looper, 10, 1, 1, 10), Yoffset + Choose(Looper, 10, 10, 2, 2)
        Next
        PenType = 0: GoSub GetPen
    End If
    PenType = 2: GoSub GetPen
    MoveToEx m_HDC, Xoffset + 0, Yoffset + 0, tPt
    For Looper = 1 To 5
        LineTo m_HDC, Xoffset + Choose(Looper, 9, 9, 0, 0, 9), Yoffset + Choose(Looper, 0, 9, 9, 1, 1)
    Next
    PenType = 0: GoSub GetPen
Case SC_MINIMIZE    ' the minimize command
    If IsDisabled Then
        PenType = 1: GoSub GetPen
        MoveToEx m_HDC, Xoffset + 9, Yoffset + 8, tPt
        LineTo m_HDC, Xoffset + 9, Yoffset + 9
        LineTo m_HDC, Xoffset + 1, Yoffset + 9
        PenType = 0: GoSub GetPen
    End If
    PenType = 2: GoSub GetPen
    MoveToEx m_HDC, Xoffset + 1, Yoffset + 8, tPt
    LineTo m_HDC, Xoffset + 9, Yoffset + 8
    MoveToEx m_HDC, Xoffset + 1, Yoffset + 7, tPt
    LineTo m_HDC, Xoffset + 9, Yoffset + 7
    PenType = 0: GoSub GetPen
Case SC_RESTORE     ' the restore command
    If IsDisabled Then
        PenType = 1: GoSub GetPen
        MoveToEx m_HDC, Xoffset + 3, Yoffset + 2, tPt
        LineTo m_HDC, Xoffset + 8, Yoffset + 2
        MoveToEx m_HDC, Xoffset + 8, Yoffset + 4, tPt
        For Looper = 1 To 7
            LineTo m_HDC, Xoffset + Choose(Looper, 8, 1, 1, 8, 8, 10, 10), Yoffset + Choose(Looper, 5, 5, 9, 9, 6, 6, 0)
        Next
        PenType = 0: GoSub GetPen
    End If
    PenType = 2: GoSub GetPen
    MoveToEx m_HDC, Xoffset + 2, Yoffset + 2, tPt
    For Looper = 1 To 10
        LineTo m_HDC, Xoffset + Choose(Looper, 2, 9, 9, 7, 7, 0, 0, 7, 7, 0), Yoffset + Choose(Looper, 0, 0, 5, 5, 3, 3, 8, 8, 4, 4)
    Next
    MoveToEx m_HDC, Xoffset + 3, Yoffset + 1, tPt
    LineTo m_HDC, Xoffset + 9, Yoffset + 1
    PenType = 0: GoSub GetPen
Case Else       ' maybe another system menu command, but it hasn't any icons
    Exit Sub
End Select
ThreeDbox m_HDC, mItem.Left, _
    mItem.Top, _
    mItem.Left + 19, _
    mItem.Top + (mItem.Bottom - mItem.Top) - 2, _
    (IsSelected = True), False
Exit Sub

GetPen:
Select Case PenType
Case 0
    SetMenuColor False, m_HDC, cObj_Pen, 0
Case 1
    SetMenuColor True, m_HDC, cObj_Pen, TextColorDisabledLight
Case 2
    If Not IsDisabled Then
        SetMenuColor True, m_HDC, cObj_Pen, lSelectBColor
    Else
        SetMenuColor True, m_HDC, cObj_Pen, TextColorDisabledDark
    End If
End Select
Return
End Sub

Private Sub DrawCheckMark(tDC As Long, tRect As RECT, CXoffset As Long, bXPstyle As Boolean, _
                        Color1 As Long, Optional Color2 As Long = -1)
' =====================================================================
' Simple little check mark drawing, looks good 'nuf I think
' =====================================================================

Dim dm As POINTAPI
Dim Yoffset As Integer, Xoffset As Integer
Dim X1 As Integer, X2 As Integer
Dim Y1 As Integer, Y2 As Integer
Dim Looper As Integer, Loops As Integer

Xoffset = CXoffset
Xoffset = 5 + Xoffset
Yoffset = ((tRect.Bottom - tRect.Top) - 8) \ 2 + tRect.Top

If Color2 <> -1 Then Loops = 1

For Looper = 0 To Loops
    If Looper Then
        SetMenuColor False, tDC, cObj_Pen, 0
        SetMenuColor True, tDC, cObj_Pen, Color2
        Yoffset = Yoffset - 1: Xoffset = Xoffset - 1
    Else
        SetMenuColor True, tDC, cObj_Pen, Color1
    End If
    If bXPstyle Then
        Yoffset = Yoffset + 1
        MoveToEx tDC, Xoffset + 1, Yoffset + 2, dm
        LineTo tDC, Xoffset + 3, Yoffset + 4
        LineTo tDC, Xoffset + 8, Yoffset - 1
        MoveToEx tDC, Xoffset + 1, Yoffset + 3, dm
        LineTo tDC, Xoffset + 3, Yoffset + 5
        LineTo tDC, Xoffset + 8, Yoffset
        Yoffset = Yoffset - 1
    Else
        ' Here we are simply tracing the outline of a check mark
        ' reated by opening a 8x8 bitmap editor and drawing a
        ' simple checkmark from left to right, bottom to top
        MoveToEx tDC, 1 + Xoffset, 4 + Yoffset, dm
        LineTo tDC, 2 + Xoffset, 4 + Yoffset
        LineTo tDC, 2 + Xoffset, 5 + Yoffset
        LineTo tDC, 3 + Xoffset, 5 + Yoffset
        LineTo tDC, 3 + Xoffset, 6 + Yoffset
        LineTo tDC, 4 + Xoffset, 6 + Yoffset
        LineTo tDC, 4 + Xoffset, 4 + Yoffset
        LineTo tDC, 5 + Xoffset, 4 + Yoffset
        LineTo tDC, 5 + Xoffset, 2 + Yoffset
        LineTo tDC, 6 + Xoffset, 2 + Yoffset
        LineTo tDC, 6 + Xoffset, 1 + Yoffset
        LineTo tDC, 7 + Xoffset, 1 + Yoffset
        LineTo tDC, 7 + Xoffset, 0 + Yoffset
    End If
Next
' replace original pen
SetMenuColor False, tDC, cObj_Pen, 0
End Sub

Private Sub SetMenuColor(bSet As Boolean, m_HDC As Long, TypeObject As ColorObjects, lColor As Long, _
    Optional bSamePenColor As Boolean = True, Optional PenWidth As Long = 1)
' =====================================================================
' This is the basic routine that sets a DC's pen, brush or font color
' =====================================================================

' here we store the most recent "sets" so we can reset when needed
Static bObject As Long, pObject As Long, bBMP As Long
Dim tBrush As Long, tPen As Long
If bSet Then    ' changing a DC's setting
    Select Case TypeObject
    Case cObj_Brush         ' brush is being changed
        tBrush = CreateSolidBrush(ConvertColor(lColor))
        bObject = SelectObject(m_HDC, tBrush)
        If bSamePenColor Then   ' if the pen color will be the same
            ' set it to. If not, you will have shapes filled with 1 color and outlined in another
            tPen = CreatePen(0, PenWidth, ConvertColor(lColor))
            pObject = SelectObject(m_HDC, tPen)
        End If
    Case cObj_Pen   ' pen is being changed (mostly for drawing lines)
            tPen = CreatePen(0, PenWidth, ConvertColor(lColor))
            pObject = SelectObject(m_HDC, tPen)
    Case cObj_Text  ' text color is changing
        SetTextColor m_HDC, ConvertColor(lColor)
    End Select
Else            ' resetting the DC back to the way it was
    Select Case TypeObject
    Case cObj_Brush     ' return original brush & delete one created
        tBrush = SelectObject(m_HDC, bObject)
        DeleteObject tBrush
        bBMP = 0
        If pObject Then  ' return original pen & delete one created
            tPen = SelectObject(m_HDC, pObject)
            DeleteObject tPen
            pObject = 0
        End If
    Case cObj_Pen        ' return original pen & delete one created
        tPen = SelectObject(m_HDC, pObject)
        DeleteObject tPen
        pObject = 0
    End Select
End If
End Sub

Public Function ConvertColor(tColor As Long) As Long
' DO NOT CALL THIS ROUTINE FROM YOUR APPLICATION.
' =====================================================================
' System colors can be referenced a couple of ways, for example using the "Button Face" color
' Button Face =  13160660 (long),   &H8000000F& (VB sys color),   15 (Windows system constant)
' GetSysColor(15) = 13160660        ' 15 is the system constant for button face
' GetSysColor(&H8000000F& And &HFF&) = 13160660     (&H8000000F& = -2147483633) is VB's system color

' Don't use function on an already valid color...
' Note: Converting 13160660 (valid color, long value)  will return Black
' So this function should only truly be used to convert negative numbers (VB system colors) to windows system colors
' =====================================================================
If tColor < 0 Then
    ConvertColor = GetSysColor(tColor And &HFF&)
Else
    ConvertColor = tColor
End If
End Function

Public Function DrawGradient(ByVal Color1 As Long, ByVal Color2 As Long, HorizontalGrade As Boolean, destDC As Long, dRect As RECT) As Long
' DO NOT CALL THIS ROUTINE FROM YOUR APPLICATION.
' =====================================================================
' probably should be revisited sometime soon.
' The gist is to draw 1 pixel rectangles of various colors to create
' the gradient effect. If the size of the rectangle is less than a
' fifth of the screen size, we'll step it up to 2 pixel rectangles
' to speed things up a bit
' =====================================================================
Dim mRect As RECT
Dim I As Long, rctOffset As Integer
Dim DestWidth As Long, DestHeight As Long
Dim PixelStep As Long, XBorder As Long
Dim Colors() As Long

On Error Resume Next
DestWidth = dRect.Right - dRect.Left
DestHeight = dRect.Bottom - dRect.Top
mRect = dRect
rctOffset = 1
If HorizontalGrade Then
    If (Screen.Width \ Screen.TwipsPerPixelX) \ DestWidth < 5 Then
        PixelStep = DestWidth \ 2
        rctOffset = 2
    Else
        PixelStep = DestWidth
    End If
    ReDim Colors(PixelStep)
    LoadColors Colors(), Color1, Color2
    mRect.Right = rctOffset + dRect.Left
    For I = 0 To PixelStep - 1
        DrawRect destDC, mRect.Left, mRect.Top, mRect.Right, mRect.Bottom, Colors(I)
        OffsetRect mRect, rctOffset, 0
    Next
Else
    If (Screen.Height \ Screen.TwipsPerPixelY) \ DestHeight < 5 Then
        PixelStep = DestHeight \ 2
        rctOffset = 2
    Else
        PixelStep = DestHeight
    End If
    ReDim Colors(PixelStep) As Long
    LoadColors Colors(), Color2, Color1
    mRect.Bottom = rctOffset + dRect.Top
    For I = 0 To PixelStep - 1
        DrawRect destDC, mRect.Left, mRect.Top, mRect.Right, mRect.Bottom, Colors(I)
        OffsetRect mRect, 0, rctOffset
    Next
End If
End Function

Private Sub LoadColors(Colors() As Long, ByVal Color1 As Long, ByVal Color2 As Long)
' =====================================================================
' routine adds/removes colors between a range of two colors
' Used by the DrawGradient routine
' =====================================================================
Dim I As Long
Dim BaseR As Single, BaseG As Single, BaseB As Single
Dim PlusR As Single, PlusG As Single, PlusB As Single
Dim MinusR As Single, MinusG As Single, MinusB As Single
BaseR = CSng(Color1 And &HFF)
BaseG = CSng(Color1 And &HFF00&) / 255
BaseB = CSng(Color1 And &HFF0000) / &HFF00&
MinusR = CSng(Color2 And &HFF&)
MinusG = CSng(Color2 And &HFF00&) / 255
MinusB = CSng(Color2 And &HFF0000) / &HFF00&
PlusR = (MinusR - BaseR) / UBound(Colors)
PlusG = (MinusG - BaseG) / UBound(Colors)
PlusB = (MinusB - BaseB) / UBound(Colors)
For I = 0 To UBound(Colors)
    BaseR = BaseR + PlusR
    BaseG = BaseG + PlusG
    BaseB = BaseB + PlusB
    If BaseR > 255 Then BaseR = 255
    If BaseG > 255 Then BaseG = 255
    If BaseB > 255 Then BaseB = 255
    If BaseR < 0 Then BaseR = 0
    If BaseG < 0 Then BaseG = 0
    If BaseG < 0 Then BaseB = 0
    Colors(I) = RGB(BaseR, BaseG, BaseB)
Next
End Sub

Public Function ExchangeVBcolor(vValue As Variant, DefaultColor As Long) As Long
' DO NOT CALL THIS ROUTINE FROM YOUR APPLICATION.
' =====================================================================
' used to convert VB colors and other colors to a long value
' =====================================================================
If IsNumeric(vValue) Then
    ' if variable passed is numeric, use it
    ExchangeVBcolor = ConvertColor(CLng(vValue))
Else
    ' otherwise the vValue is probably as string
    Select Case CStr(vValue)
    Case "vbWhite": ExchangeVBcolor = vbWhite
    Case "vbBlack": ExchangeVBcolor = vbBlack
    Case "vbBlue": ExchangeVBcolor = vbBlue
    Case "vbGreen": ExchangeVBcolor = vbGreen
    Case "vbRed": ExchangeVBcolor = vbRed
    Case "vbMagenta": ExchangeVBcolor = vbMagenta
    Case "vbYellow": ExchangeVBcolor = vbYellow
    Case "vbCyan": ExchangeVBcolor = vbCyan
    Case "vbMaroon": ExchangeVBcolor = vbMaroon
    Case "vbOlive": ExchangeVBcolor = vbOlive
    Case "vbNavy": ExchangeVBcolor = vbNavy
    Case "vbPurple": ExchangeVBcolor = vbPurple
    Case "vbTeal": ExchangeVBcolor = vbTeal
    Case "vbGray": ExchangeVBcolor = vbGray
    Case "vbSilver": ExchangeVBcolor = vbSilver
    Case "vbViolet": ExchangeVBcolor = vbViolet
    Case "vbOrange": ExchangeVBcolor = vbOrange
    Case "vbGold": ExchangeVBcolor = vbGold
    Case "vbIvory": ExchangeVBcolor = vbIvory
    Case "vbPeach": ExchangeVBcolor = vbPeach
    Case "vbTurquoise": ExchangeVBcolor = vbTurquoise
    Case "vbTan": ExchangeVBcolor = vbTan
    Case "vbBrown": ExchangeVBcolor = vbBrown
    Case "vbScrollBars": ExchangeVBcolor = ConvertColor(vbScrollBars)
    Case "vbDesktop": ExchangeVBcolor = ConvertColor(vbDesktop)
    Case "vbActiveTitleBar": ExchangeVBcolor = ConvertColor(vbActiveTitleBar)
    Case "vbInactiveTitleBar": ExchangeVBcolor = ConvertColor(vbInactiveTitleBar)
    Case "vbMenuBar": ExchangeVBcolor = ConvertColor(vbMenuBar)
    Case "vbWindowBackground": ExchangeVBcolor = ConvertColor(vbWindowBackground)
    Case "vbWindowFrame": ExchangeVBcolor = ConvertColor(vbWindowFrame)
    Case "vbMenuText": ExchangeVBcolor = ConvertColor(vbMenuText)
    Case "vbWindowText": ExchangeVBcolor = ConvertColor(vbWindowText)
    Case "vbTitleBarText": ExchangeVBcolor = ConvertColor(vbTitleBarText)
    Case "vbActiveBorder": ExchangeVBcolor = ConvertColor(vbActiveBorder)
    Case "vbInactiveBorder": ExchangeVBcolor = ConvertColor(vbInactiveBorder)
    Case "vbApplicationWorkspace": ExchangeVBcolor = ConvertColor(vbApplicationWorkspace)
    Case "vbHighlight": ExchangeVBcolor = ConvertColor(vbHighlight)
    Case "vbHighlightText": ExchangeVBcolor = ConvertColor(vbHighlightText)
    Case "vbButtonFace": ExchangeVBcolor = ConvertColor(vbButtonFace)
    Case "vbButtonShadow": ExchangeVBcolor = ConvertColor(vbButtonShadow)
    Case "vbGrayText": ExchangeVBcolor = ConvertColor(vbGrayText)
    Case "vbButtonText": ExchangeVBcolor = ConvertColor(vbButtonText)
    Case "vbInactiveCaptionText": ExchangeVBcolor = ConvertColor(vbInactiveCaptionText)
    Case "vb3DHighlight": ExchangeVBcolor = ConvertColor(vb3DHighlight)
    Case "vb3DDKShadow": ExchangeVBcolor = ConvertColor(vb3DDKShadow)
    Case "vb3DLight": ExchangeVBcolor = ConvertColor(vb3DLight)
    Case "vb3DFace": ExchangeVBcolor = ConvertColor(vb3DFace)
    Case "vb3Dshadow": ExchangeVBcolor = ConvertColor(vb3DShadow)
    Case "vbInfoText": ExchangeVBcolor = ConvertColor(vbInfoText)
    Case "vbInfoBackground": ExchangeVBcolor = ConvertColor(vbInfoBackground)
    Case Else
        ' not the expected string above, maybe its a Hex value
        If Left$(vValue, 2) = "&H" Then
            ExchangeVBcolor = ConvertColor(Val(vValue))
        Else ' nope, don't know what it could be, so return the default value
            ExchangeVBcolor = DefaultColor
        End If
    End Select
End If
End Function

Public Sub LoadFontMenu(vFontArray As Variant, Optional FontType As Long)
' DO NOT CALL THIS ROUTINE FROM YOUR APPLICATION.
' =====================================================================
' Routine called by the cMenuItems class to fill an array with all the
' fonts on the system
' =====================================================================
ReDim vFonts(0 To 0)    ' this is a module wide array
Dim hDC As Long
' need a DC for the EnumFontFamProc
hDC = GetDC(CLng(Mid(hWndRedirect, 2)))
' call the enumerator function
EnumFontFamilies hDC, vbNullString, AddressOf EnumFontFamProc, FontType
' we can release the DC
ReleaseDC CLng(Mid(hWndRedirect, 2)), hDC
' the EnumFontFamilies doesn't sort them alphabetically, so we do that here
If UBound(vFonts) Then ShellSort vFonts
' return the array
vFontArray = vFonts
Erase vFonts
End Sub

Private Function EnumFontFamProc(lpNLF As LOGFONT, lpNTM As NEWTEXTMETRIC, ByVal FontType As Long, lParam As Long) As Long
' =====================================================================
' This is the enumerator callback function
' For the cMenuItems class the 3 font type options are
' 1. All fonts
' 2. TrueType fonts
' 3. Non-TrueType fonts (basically system fonts)
 Dim FaceName As String, bInclude As Boolean
'continue enumeration
 EnumFontFamProc = 1
' if user opted for a specific type of font, test it
Select Case lParam      ' this was passed by calling function
  Case RASTER_FONTTYPE
      bInclude = ((FontType = RASTER_FONTTYPE) Or (FontType = 0)) ' = RASTER_FONTTYPE Or FontType = 0)
  Case TRUETYPE_FONTTYPE
      bInclude = (FontType = TRUETYPE_FONTTYPE) ' = TRUETYPE_FONTTYPE)
  Case Else
      bInclude = True
End Select
If bInclude Then
  'convert the returned string & add it to the array
  FaceName = StrConv(lpNLF.lfFaceName, vbUnicode)
  ReDim Preserve vFonts(0 To UBound(vFonts) + 1)
  vFonts(UBound(vFonts)) = StringFromBuffer(FaceName)
End If
End Function

Private Sub ShellSort(vArray As Variant)
' =====================================================================
' a pretty fast sorting function. Forgot where I got it from originally
' but modified to handle 1-dimensional arrays of any type
' =====================================================================
  Dim lLoop1 As Long
  Dim lHold As Long
  Dim lHValue As Long
  Dim lTemp As Variant

  lHValue = LBound(vArray)
  Do
    lHValue = 3 * lHValue + 1
  Loop Until lHValue > UBound(vArray)
  Do
    lHValue = lHValue / 3
    For lLoop1 = lHValue + LBound(vArray) To UBound(vArray)
      lTemp = vArray(lLoop1)
      lHold = lLoop1
      Do While vArray(lHold - lHValue) > lTemp
        vArray(lHold) = vArray(lHold - lHValue)
        lHold = lHold - lHValue
        If lHold < lHValue Then Exit Do
      Loop
      vArray(lHold) = lTemp
    Next lLoop1
  Loop Until lHValue = LBound(vArray)
End Sub

Public Function StripFile(Pathname As String, DPNEm As String) As String
Dim ChrsIn As String, Chrs As Integer
' =====================================================================
' Routine simply strips out either a drive, file name, folder or
' extension from the passed PathName. If empty string return value, function failed
' Function always returns a back slash after drive & path return values
' =====================================================================

' Inserted by LaVolpe OnError Insertion Program.
On Error GoTo StripFile_General_ErrTrap
If Pathname = "" Then Exit Function
ChrsIn = Pathname
Select Case InStr("DPNEm", DPNEm)
Case 1:     ' Return the Drive Letter
    Chrs = InStr(ChrsIn, ":")
    If Chrs Then
        StripFile = Left(ChrsIn, Chrs) & "\"         'get the drive
    Else    ' test for a network/shared drive
        Chrs = InStr(ChrsIn, "\\")
        If Chrs = 1 Then
            Chrs = InStr(Chrs + 2, ChrsIn, "\")
            If Chrs Then StripFile = Left$(ChrsIn, Chrs) Else StripFile = ChrsIn & "\"
        End If
    End If
Case 2:     ' Return the full Path
    Chrs = InStrRev(ChrsIn, "\")
    If Chrs = 0 Then Chrs = InStr(ChrsIn, ":") Else Chrs = Chrs - 1
    If Chrs Then StripFile = Left$(ChrsIn, Chrs) & "\"
Case 3:     ' Return the full File Name
    Chrs = InStrRev(ChrsIn, "\") 'check to see if a forward slash exists
    If Chrs = 0 Then Chrs = InStr(ChrsIn, ":")
    If Chrs Then StripFile = Mid$(ChrsIn, Chrs + 1)
Case 4:     ' Return the File Extension
    Chrs = InStrRev(ChrsIn, ".") 'check to see if a full stop exists
    If Chrs Then 'if a full stop is found in the passed string
        If InStr(Chrs, "\") = 0 Then StripFile = Mid$(ChrsIn, Chrs + 1)
    End If
Case 5:     ' Return filename less the extension
    Chrs = InStrRev(ChrsIn, "\") 'check to see if a forward slash exists
    If Chrs = 0 Then Chrs = InStr(ChrsIn, ":")
    If Chrs Then ChrsIn = Mid$(ChrsIn, Chrs + 1)
    Chrs = InStrRev(ChrsIn, ".")
    If Chrs > 1 Then ChrsIn = Left$(ChrsIn, Chrs - 1) Else Chrs = 0
    If Chrs Then StripFile = ChrsIn
End Select
Exit Function

' Inserted by LaVolpe OnError Insertion Program.
StripFile_General_ErrTrap:
MsgBox "Err: " & Err.Number & " - Procedure: StripFile" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Function

Public Function GetFloppyIcon(sDrive As String) As Long
' =====================================================================
' Only used by the cMenuItems class when displaying the lvDrives menu
' Since we query the drive for the icon it uses, we don't want to
' constantly query a floppy drive, especially if nothing is in it
' So, here we query once and store the icon to return time after time
' =====================================================================

' Note: The icon is destroyed when the last form is unsubclassed
If FloppyIcon = 0 Then
    ' we don't have a copy of it, lets get one now
    FloppyIcon = ExtractAssociatedIcon(App.hInstance, sDrive, 0)
End If
' pass the handle along
GetFloppyIcon = FloppyIcon
End Function

Public Function StringFromBuffer(Buffer As String) As String
' =====================================================================
' Used by several routines to retrieve a the string value from a
' string filled with Null Characters
' =====================================================================
Dim nPos As Long

nPos = InStr(Buffer, vbNullChar)
If nPos > 0 Then
    StringFromBuffer = Left$(Buffer, nPos - 1)
Else
    StringFromBuffer = Buffer
End If
End Function

Private Function GetToolTipWindow(pHwnd As Long) As Long
' =====================================================================
' Function returns the first toolbar child window of the pHwnd it finds
' =====================================================================
On Error GoTo ReDimensionArray
Dim I As Integer, tHwnd As Long
I = UBound(tbarClass)
For I = 0 To UBound(tbarClass)
    tHwnd = FindWindowEx(pHwnd, 0, tbarClass(I), vbNullString)
    If tHwnd Then Exit For
Next
GetToolTipWindow = tHwnd
Exit Function

ReDimensionArray:
ReDim tbarClass(0)
tbarClass(0) = "msvb_lib_toolbar"
Resume
End Function

Public Sub AddToolbarClass(sClass As String)
' =====================================================================
' This is an attempt to allow users to add custom toolbar classes to
' this program. If your custom toolbar class handles popups that will be
' modified by this program, call this function to add that class name.
' Now, if the class doesn't handle the menus, but it created a subclass
' to do them add the subclass name here.

' For Example: VB's Toolbar20WndClass creates the subclass msvb_lib_toolbar
' which processes the toolbar button menus"
' This VB subclass is automatically added to the array whenever you start
' this program
' =====================================================================
On Error GoTo ReDimensionArray
Dim I As Integer
I = UBound(tbarClass)
For I = 0 To UBound(tbarClass)
    If tbarClass(I) = sClass Then Exit Sub
Next
ReDim Preserve tbarClass(0 To UBound(tbarClass))
tbarClass(UBound(tbarClass)) = sClass
Exit Sub

ReDimensionArray:
ReDim tbarClass(0)
tbarClass(0) = "msvb_lib_toolbar"
Resume
End Sub

Private Sub IDCurrentWindow(hWnd As Long, MenuID As Long, bSystemMenu As Boolean)
' =====================================================================
' Primary purpose is to identify the window that has the focus with exceptions for MDI children
' This was removed from the MenuMessages function which ineffectively tried to track active
' forms real-time. This is more effective
' We need to positively identify the class for the active window's menus. Unfortunately, the
' MDI Client Window handles all MDI child forms; so each child's menus are handled via the
' MDI Client's hWnd messages. However, since each child can have its own menus, we create
' a class for each child; therefore, we can't use the MDI Client's hWnd as the key
' =====================================================================
' Following are all the exceptions we need to address. Note: Almost all are MDI related

' the override exceptions. If this is the system menu, the hWnd passed is always the parent of the system menu
If bSystemMenu Then
    hWndRedirect = "h" & hWnd
    MenuID = colMenuItems("h" & hWnd).SystemMenu ' return its handle
    Exit Sub
End If

' default active window is the window that received the windows message
hWndRedirect = "h" & hWnd
'Debug.Print "hwnd to reroute is "; hWnd
Dim hMDI As Long, hParent As Long

' With MDI's we need to reroute to the active child if it exists and it has menus
If colMenuItems("h" & hWnd).IsMDIclient Then  ' check if window is MDI Client
    ' it is, see if it has any active children
    hMDI = SendMessage(hWnd, WM_MDIGETACTIVE, 0&, ByVal 0&)
    hParent = GetParent(hWnd)   ' may need to reference back to the parent MDI form
Else
    If colMenuItems("h" & hWnd).MDIClient Then      ' see if window is the parent MDI form
        ' it is, see if it has any active children by querying it's MDI Client window
        hParent = hWnd  ' may need to reference back to the parent MDI form
        hMDI = SendMessage(colMenuItems("h" & hWnd).MDIClient, WM_MDIGETACTIVE, 0&, ByVal 0&)
    End If
End If
If hMDI Then ' initial assignment will the the active MDI child
    If tempRedirect Then hWndRedirect = "h" & tempRedirect Else hWndRedirect = "h" & hMDI
    ' we have an active MDI Child, but if it doesn't have any visible menus, we set the active form to the
    ' parent MDI form
    If colMenuItems("h" & hMDI).IsMenuLess Then    ' returns -1 or 0 depending if no menus
        'Debug.Print "child "; hMDI; " has no menus"; GetMenu(hMDI)
        ' child has no menus, but MDI parents always place a system menu on their menu bar for
        ' maximized children. We could check to  the window state or simply compare the cached copy of
        ' the system menu handle that we draw. If not the child's system menu, reroute back to the parent
        If colMenuItems("h" & hMDI).SystemMenu <> MenuID Then
           If tempRedirect = 0 Then hWndRedirect = "h" & hParent
        End If
    End If
Else
    If tempRedirect Then hWndRedirect = "h" & tempRedirect
End If
'Debug.Print "Opening menu "; MenuID; " for window "; hWndRedirect; " ActiveHwnd h " & hWnd; " h " & tempRedirect; ""
End Sub

Private Function IdentifyAccelerator(KeyCode As Long, hMenu As Long) As Long
' =====================================================================
' Function will correctly select the menu item assoicated with a pressed accelerator key
' The requirements for the return value are simple: Long value
' - High Word is the action code MNC_EXECUTE, MNC_IGNORE, MNC_SELECT or MNC_CLOSE (not used here)
' - Low Word only required if action is execute or select & will be the zero-based index of the menu item

' The XferPanelData.Accelerators is a 1-based list of accelerator keys used for the passed menu.
' This makes is very easy to correlate the accelerator key to the menu item's zero-based index
' =====================================================================

Dim Index As Integer, I As Integer
' ensure we have the correct panel data which contains the accelerator keys
colMenuItems(hWndRedirect).GetPanelItem hMenu
With XferPanelData
    ' find the accelerator key pressed
    Index = InStr(.Accelerators, UCase(Chr$(KeyCode)))
    If Index = 0 Then   ' doesn't exist so we send a beep
            IdentifyAccelerator = MakeLong(0, MNC_IGNORE)
    Else    ' it exists, but is it used more than once?
        ' we'll check this by looking for the same accelerator key backwards in the string and
        ' if the result is identical, the key is only used once!
        If InStrRev(.Accelerators, UCase(Chr$(KeyCode))) = Index Then
            ' used only once, we'll make windows select that item
            IdentifyAccelerator = MakeLong(Index - 1, MNC_EXECUTE)
        Else    ' geez!
            ' here's where it gets tricky: there is more than one active menu item with same accelerator key
            ' The good news is that we don't need to be concerned with hidden/non-visible menu items ' cause
            ' the cMenuItems' GetPanelMetrix routine only logs current menu accelerator keys each time a
            ' menu panel is displayed
            Dim vIndex() As Integer, MI() As Byte, MII As MENUITEMINFO
            ReDim vIndex(0)
            ' we'll build an array of menu items containing this accelerator
            Do While Index > 0
                ReDim Preserve vIndex(0 To UBound(vIndex) + 1)
                vIndex(UBound(vIndex)) = Index
                Index = InStr(Index + 1, .Accelerators, UCase(Chr$(KeyCode)))
            Loop
            ' now we need to loop thru the menu items having this accelerator to see if one of them is selected
            For I = 1 To UBound(vIndex)
                ReDim MI(0 To 1023)
                MII.cbSize = Len(MII)
                MII.fType = 0
                MII.fMask = MIIM_STATE
                MII.cch = UBound(MI)
                GetMenuItemInfo hMenu, vIndex(I) - 1, True, MII       ' get the submenu item information
                If ((MII.fState And MF_HILITE) = MF_HILITE) Then
                    Index = I   ' found the one that is highlighted, so abort loop
                    Exit For
                End If
            Next
            Erase MI
            ' good now we know if one of them is already selected. Regardless, we don't want to execute
            ' the menu item since we don't know which one the user really wants. We'll just select it.
            If Index Then
                ' yep, one currently selected, so select the next menu item using the same accelerator key
                If Index = UBound(vIndex) Then Index = 1 Else Index = Index + 1
            Else    ' nope, none currently selected, so select the first one
                Index = 1
            End If
            IdentifyAccelerator = MakeLong(vIndex(Index) - 1, MNC_SELECT)
            Erase vIndex
        End If
    End If
End With
End Function

Private Sub ProcessKeyStroke(hWnd As Long, KeyStroke As Long, bKeyUp As Boolean)
' =====================================================================
' Will not work on WinME since the GetKeyState function was disabled in that
' version of Windows. The function will still return the key pressed.

' this function will return keystrokes to a MDI Parent form when no children
' are loaded. This could be useful, since MDI parents have no KeyPreview
' property, nor KeyDown nor KeyUp events. The modMenus.ReturnMDIKeystrokes
' property must be set to true and the keystrokes will be returned via the
' MDI Parent's cTips class
' =====================================================================

On Error GoTo FailedKeyRepeater
Dim hMDI As Long, hParent As Long, pClass As Long, ShiftStatus As Long
If colMenuItems("h" & hWnd).IsMDIclient Then  ' check if window is MDI Client
    ' it is, see if it has any active children
    hMDI = SendMessage(hWnd, WM_MDIGETACTIVE, 0&, ByVal 0&)
    hParent = GetParent(hWnd)   ' may need to reference back to the parent MDI form
Else
    If colMenuItems("h" & hWnd).MDIClient Then      ' see if window is the parent MDI form
        ' it is, see if it has any active children by querying it's MDI Client window
        hParent = hWnd  ' may need to reference back to the parent MDI form
        hMDI = SendMessage(colMenuItems("h" & hWnd).MDIClient, WM_MDIGETACTIVE, 0&, ByVal 0&)
    End If
End If
If hParent <> 0 And hMDI = 0 Then
    pClass = colMenuItems("h" & hParent).ShowTips
    If pClass = 0 Then Exit Sub
    If (GetKeyState(VK_SHIFT) And &HF0000000) Then ShiftStatus = ShiftStatus Or 1
    If (GetKeyState(VK_CONTROL) And &HF0000000) Then ShiftStatus = ShiftStatus Or 2
    If (GetKeyState(VK_MENU) And &HF0000000) Then ShiftStatus = ShiftStatus Or 4
    Dim oTipClass As cTips
    CopyMemory oTipClass, pClass, 4&
    oTipClass.SendMDIKeyPress KeyStroke, ShiftStatus, bKeyUp
    CopyMemory oTipClass, 0&, 4&
    Set oTipClass = Nothing
End If
FailedKeyRepeater:
End Sub

Public Function SetMinMaxInfo(hWnd As Long, _
    MaximizedW As Long, MaximizedH As Long, _
    MaximizedLeft As Long, MaximizedTop As Long, _
    MaxDragSizeW As Long, MaxDragSizeH As Long, _
    MinDragSizeW As Long, MinDragSizeH As Long)

' =====================================================================
' Optional function that will restrict the size of a window when user
' attempts to resize it.
' =====================================================================

On Error Resume Next
If colMenuItems("h" & hWnd).hPrevProc = 0 Then Exit Function

Dim uMinMax As MINMAXINFO, sRect As RECT
' function to return the working area of the desktop as a rectangle
' function always returns pixels so no need to divide by screen.TwipsPerPixelX/Y
' For width/height, this function is same as screen.Width/ScreenTwipsPerPixelx, etc.
' However, this function will also return the correct Top/Left of the desktop so
' a left aligned toolbar or top aligned toolbar has no negative effects.
SystemParametersInfo SPI_GETWORKAREA, 0, sRect, 0

' width/height of the form when it's maximized
' Note: by passing -1, routine will get the working area of the desktop
' Passing a value > the working area for bordered forms has the same effect
' as using the working area size -- Windows sees to that
' In addition the Max H/W will never exceed the values passed for
' MaxDragSizeW & MaxDragSizeH; again Windows sees to this
' However, MaxDragSizeW/H can be larger than the maximized window size!
If MaximizedW < 0 Then MaximizedW = sRect.Right - sRect.Left
If MaximizedH < 0 Then MaximizedH = sRect.Bottom - sRect.Top
uMinMax.ptMaxSize.X = MaximizedW
uMinMax.ptMaxSize.Y = MaximizedH

' Left/Top position of the form when it's maximized
If MaximizedLeft < 0 Then MaximizedLeft = sRect.Left
If MaximizedTop < 0 Then MaximizedTop = sRect.Top
uMinMax.ptMaxPosition.X = MaximizedLeft
uMinMax.ptMaxPosition.Y = MaximizedTop

' max width that user can drag a form. This has an affect on the
' maximized height & width of the window also. See notes above
If MaxDragSizeW < 0 Then MaxDragSizeW = sRect.Right - sRect.Left
If MaxDragSizeH < 0 Then MaxDragSizeH = sRect.Bottom - sRect.Top
uMinMax.ptMaxTrackSize.X = MaxDragSizeW
uMinMax.ptMaxTrackSize.Y = MaxDragSizeH

' min width that user can drag a form
If MinDragSizeW < 0 Then MinDragSizeW = 0
If MinDragSizeH < 0 Then MinDragSizeH = 0
uMinMax.ptMinTrackSize.X = MinDragSizeW
uMinMax.ptMinTrackSize.Y = MinDragSizeH

colMenuItems("h" & hWnd).RestrictSize VarPtr(uMinMax), True
End Function

Private Sub LoadDefaultColors()
bModuleInitialized = True
lSelectBColor = GetSysColor(COLOR_HIGHLIGHT)
TextColorNormal = GetSysColor(COLOR_MENUTEXT)
TextColorSelected = GetSysColor(COLOR_HIGHLIGHTTEXT)
TextColorDisabledDark = GetSysColor(COLOR_GRAYTEXT)
TextColorDisabledLight = TextColorSelected ' GetSysColor(COLOR_HIGHLIGHTTEXT)
TextColorSeparatorBar = lSelectBColor
SeparatorBarColorDark = GetSysColor(COLOR_BTNSHADOW)
SeparatorBarColorLight = GetSysColor(COLOR_BTNHIGHLIGHT)
CheckedIconBColor = GetSysColor(COLOR_BTNLIGHT)
End Sub
