VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMenuItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ========================================================================
'  CLASS WHICH CREATES OWNER-DRAWN MENUS  AND STORES ALL ATTRIBUTES
' ========================================================================
'
' PROJECT NOT COMPATIBLE WITH WinNT 3.x
'___________________________________________________________________________________
' /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
'
' Author:  Keith Fox, aka LaVolpe   the_foxes@hotmail.com
' Written the hard way -- line by line !
' Couldn't have done it without 100's of visits to MSDN and lots of crashes :)
' ------------- history of current version
' 10 Feb. Complete rewrite of project
' 15 Feb. Added 7 ready-to-use custom menus (CreateSubMenuCustom)
' 17 Feb. Added properties to modMenus:
'       -- MenuFontName to change submenu font
'       -- MenuFontSize to change submenu font size
'       -- HighlightGradient to offer a gradient highlighting scheme for selected menu items
'       -- HighlightDisabledMenuItems to always highlight selected disabled menu items
'       -- ItalicizeSelectedItems to italicize highlighted menu items
' 19 Feb. Added call to color dialog for custom menus lvColors
' 23 Feb. Added 7 custom menu creation functions
'       -- Added PopupMenuCustom function to allow custom menus as popups
'       -- Added CACHE flag
' 2 Mar. Bug found by me. GetMenuMetrix does not constantly test for the image
'       handle ref'd by the menu item caption. This causes problems because
'       the ImageList ExtractIcon function replaces existing handle with a new
'       one & program fails to draw menu icon 'cause it no longer exists.
'       Additionally, if user using a picBox, Image control, etc and changes
'       the contents of that control, this program will never know. The
'       fix is to test the menu icon every single time.
' 3 Mar. Bug found by me. ExtractIconEx(filename,0,smallimg,1) was returning a 32x32 icon
'       if found in filename. Documentation I read said if 0 was passed above, no icon
'       would be returned. Fix, replace 0 with variable and destroy the icon if returned
' 6 Mar. Bug found by Zhu Jinyong. Menu items, if only one on a multi-column menu
'       didn't display properly. This was fixed in the GetPanelMetrix routine
' 7 Mar. Bug found by George Aslanis. When MDI children had no menus, the menus on the
'       parent form would disappear if they were previously displayed in a popup. After further testing,
'       found that system menus were affected also. The bottom line is that the program could not
'       identify if child forms have menus. Per MSDN the GetMenuItemCount function is meaningless
'       for child forms; therefore I needed to force user to identify whether or not MDI children had menus
'       or not. This is now added as an optional parameter in the SetMenu function. See the Enumerator
'       declaration for ContainerType in modMenus for a description of the flags to use
' 10 Mar. Ongoing battle with owner-drawn separator bars not being remeasured by Windows
'       A final solution of sorts: kill all menu items & rebuild them when a user changes
'       menu font/size during runtime. Function RestoreMenus rewritten
'       - Related problem with Sidebars. Was storing widest menu item so separator bars would
'         be measured indirectly. Screwed up menus when toggling sidebar visibility on
'         scrolling menus. This related error also corrected: GetPanelMetrix
'       - hWndRedirect was not being reset in SetMenu if form was already subclassed.
'         This could cause errors when subclassing a mix-match of MDI child & non-MDI child forms
'       - Changed the way icons are extracted for custom menus containing files list.
'         Icons now displayed are exactly what Explorer would display: DrawMenuIcon
'         -- RaisedIconOnSelect gives icons a 3D look
'         -- SelectedItemBackColor allows custom back color selection for menu items
'         -- ReturnMDIkeystrokes gives MDI parents without loaded child forms a KeyUp & KeyDown event
'       - Added following function
'         -- SetMinMaxInfo to allow forms using SetMenu to restrict sizing of the form (bonus routine)
' 11 Mar. Tweaked routines to allow sidebars to be the 1st item in a list/combo box if desired. These
'           can also retain the check value if needed
'           -- Added another flag for sidebars: SBDisabled to force a sidebar to be disabled as disabling a
'              sidebar if in the list/combo box is not possible
'       - Exposed the following menu properties as new modMenus properties
'           -- CheckMarksXPstyle replaces Win95 checkmarks with WinXP/2K style
'           -- MenuItemTextColor is menu item forecolor when not selected
'           -- SelectedItemTextColor is menu item forecolor when selected
'           -- DisabledTextColor_Dark is disabled menu item shadow color
'           -- DisabledTextColor_Light is disabled menu item highlight color
'           -- SeparatorBarTextColor is text forecolor for sidebars with text
'           -- SeparatorBarColor_Dark is separator bar shadow color
'           -- SeparatorBarColor_Light is separator bar highlight color
'           -- CheckedIconBackColor is the rectangle background color for icons used as checkmarks
'12 Mar. Bug identified by Robert Greig. WinNT won't process accelerator keys. Fix is to process the
'           WM_MenuChar message. MenuMessages updated to process this message, and
'           GetPanelMetrix modified to track accelerator keys and new functions created to assist:
'           modMenus.IdentifyAccelerator & this class's ReturnAcceleratorKey
' Last Remarked/Updated: 11 Mar 03. Readme.html last updated 11 Mar 03
'___________________________________________________________________________________
' \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/

Option Explicit
Option Compare Text  ' we don't want to mess with case sensitivity

Private imgList As Control             ' handle to image list. Cached to pass to child classes, if any
Private PrevProc As Long               ' previous window procedure
Private TipCallBackProc As Long        ' whether or not tips will be displayed (set to false if not needed)
Private PrevTipCallBack As Long        ' allows user to toggle between two form's Tips class
Private hSysMenu As Long               ' System menu for this form
Private hMDIclient As Long             ' MDI Client window if parent is MDI
Private bMDI As Boolean                 ' class belongs to a MDI client window
Private bNoChildMenus As Boolean   ' child menus have no menus/MDI parent handles all menus
Private uMinMax As MINMAXINFO

' see Class Initialize routine for brief description
Private vItems() As MenuComponentData
Private vPanels() As PanelData
Private cItems As Collection
Private cPanels As Collection
Private gMenus As Collection

' //////// Constants & Types used only in this Class \\\\\\\\\\\
Private Const LB_GETTEXT         As Long = &H189
Private Const LB_GETTEXTLEN      As Long = &H18A
Private Type CHOOSECOLORSTRUCT      ' for the color dialog box
   lStructSize     As Long
   hwndOwner       As Long
   hInstance       As Long
   rgbResult       As Long
   lpCustColors    As Long
   Flags           As Long
   lCustData       As Long
   lpfnHook        As Long
   lpTemplateName  As String
End Type
Private Type MENUITEMINFO_string    ' save menu items as string vs byte arrays
     cbSize As Long
     fMask As Long
     fType As Long
     fState As Long
     wID As Long
     hSubMenu As Long
     hbmpChecked As Long
     hbmpUnchecked As Long
     dwItemData As Long
     dwTypeData As String
     cch As Long
End Type
' =====================================================================
' COMDLG32.DLL Function Calls << used for the lvColors custom menu
' =====================================================================
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (lpcc As CHOOSECOLORSTRUCT) As Long
' =====================================================================
' MPR.DLL Function Calls << used for the lvDrives custom menu
' =====================================================================
Private Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long
' =====================================================================
' KERNEL32.DLL Function Calls << used for the lvDrives custom menu
' =====================================================================
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetLocaleInfoA Lib "kernel32" (ByVal lLCID As Long, ByVal lLCTYPE As Long, ByVal strLCData As String, ByVal lDataLen As Long) As Long
Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Sub lstrcpyn Lib "kernel32" (ByVal strDest As String, ByVal strSrc As Any, ByVal lBytes As Long)
' =====================================================================
' VERSION.DLL Function Calls << used for the FILES: flag
' =====================================================================
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal sFile As String, ByVal lpIgnored As Long, ByVal lpSize As Long, ByVal lpBuf As Long) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (ByVal lpBuf As Long, ByVal szReceive As String, lpBufPtr As Long, lLen As Long) As Long


Private FormHwnd As Long      ' form's handle. Cached for FormIndex routine
Private tipDefault As String  ' gMenu default tip from parent submenu

Public Sub InitializeSubMenu(hForm As Long, Optional oImgList As Control, _
                                            Optional CallBacks As Long, _
                                            Optional bNoSysMenu As Boolean, _
                                            Optional TimeConsumingMenu As Long = 0, _
                                            Optional DefaultTip As String, Optional Flags As Long = 0)
' ========================================================================
' Called when a form is being subclassed or a gMenu child class is being initialized
' ========================================================================
Dim sysMenu As Long
FormHwnd = hForm                                ' cache
' rerouting of cTips classes
If CallBacks Then
    ' cTips is being rerouted or intialized
    If CallBacks <> PrevTipCallBack Then
        ' we set the previous cTips reference, if any
        PrevTipCallBack = TipCallBackProc
        ' we set the current cTips reference
        TipCallBackProc = CallBacks                 ' cache
    End If
Else
    ' cTips is being routed back from a previous reRoute call
    TipCallBackProc = PrevTipCallBack
End If
If Not oImgList Is Nothing Then Set imgList = oImgList  ' imagelist to use, if any
' process the system menu if bNoSysMenu = False & not already processed
If hSysMenu = 0 And bNoSysMenu = False Then
    hSysMenu = GetSystemMenu(hForm, 0)
    If hSysMenu Then IsWindowList hSysMenu, True
End If
If TimeConsumingMenu Then
    ' some menus may take extra time to display like those that display the sample font or those that
    ' retrieve an icon from the file name or its associated application. If so, we'll add an hourglass
    ' icon when those submenus are about to be displayed.
    Dim mPanel As PanelData
    mPanel.Hourglass = True
    ReDim vPanels(0 To 1)
    vPanels(1) = mPanel
    cPanels.Add 1, "p" & TimeConsumingMenu
End If
' Except for the lvDrives custom menu, custom menus do not have tips. Therefore
' we supply the tip from the menu item displaying the custom menu
tipDefault = DefaultTip
' Note: Although changes incurred in this routine seem like it should effect any chlid classes, we don't have
' to be concerned because exisiting child classes are destroyed and re-created as needed.
'Debug.Print hForm; " initialzed with tips callback as "; Format(CallBacks, "True/False"); " toolbar="; IsToolbar
End Sub

Public Property Get MDIClient() As Long
' ========================================================================
' Returns/Sets of whether this form is a MDI Parent and hWnd of its Client
' ========================================================================
    MDIClient = hMDIclient
End Property
Public Property Let MDIClient(hWnd As Long)
    hMDIclient = hWnd
End Property

Public Property Get IsMDIclient() As Boolean
' ========================================================================
' Returns/Sets of whether this form is the MDI Parent's Client window
' ========================================================================
    IsMDIclient = bMDI
End Property
Public Property Let IsMDIclient(bYesNo As Boolean)
    bMDI = bYesNo
End Property

Public Property Get IsMenuLess() As Boolean
' ========================================================================
' Returns/Sets of whether this form is the MDI Parent's Client window
' ========================================================================
    IsMenuLess = bNoChildMenus
End Property
Public Property Let IsMenuLess(bYesNo As Boolean)
    bNoChildMenus = bYesNo
End Property

Public Property Get SystemMenu() As Long
' ========================================================================
' Returns the system menu handle
' ========================================================================
SystemMenu = hSysMenu
End Property

Public Property Get ShowTips() As Long
' ========================================================================
' Returns status of whether tips should be returned or not
' ========================================================================
ShowTips = TipCallBackProc
End Property

Public Function Tips(MenuID As Long, hMenu As Long) As String
' ========================================================================
' Function returns the menu items Tip as a string
' See GetMenuItem for more detailed information on referencing the gMenu collection
' System menu will provide the hMenu parameter of zero
' ========================================================================
On Error Resume Next
Dim mIndex As Integer
' see if we have this menu item cached
mIndex = cItems(MenuID & "." & hMenu)
If Err = 0 Then
    Tips = vItems(mIndex).Tip
    Exit Function
End If
' if not cached, then most likely an item in a child class
Tips = gMenus("g" & hMenu).Tips(MenuID, hMenu)
Err.Clear
'Debug.Print "Tips = "; Tips; "<"
End Function

Public Property Get ImageListObject() As Control
' ========================================================================
' Only used when automatically subclassing a MDI child in the SetMenu function
' This allows all MDI children to use same ImageList as MDI Parent
' ========================================================================
' return MDI Parent's imagelist control
     Set ImageListObject = imgList
End Property

Public Property Let hPrevProc(hProc As Long)
' ========================================================================
' Returns/sets the previous window procedure for the form
' ========================================================================
    PrevProc = hProc
End Property
Public Property Get hPrevProc() As Long
    hPrevProc = PrevProc
End Property

Public Function RestrictSize(lParam As Long, bSet As Boolean) As Boolean
' ========================================================================
' Function restricts the size of a window to that defined by the user
' ========================================================================

' lparam is pointer to a MinMax Structure
If bSet Then    ' setting the size restrictions
    CopyMemory uMinMax, ByVal lParam, Len(uMinMax)
Else            ' applying the size restrictions
    ' if the size wasn't set, bug out & let windows handle it
    ' note that uMinMax.ptMaxSize.X was set to -1 during class initialization
    If uMinMax.ptMaxSize.X < 0 Then Exit Function
    ' otherwise return the sizes the user defined
    CopyMemory ByVal lParam, uMinMax, Len(uMinMax)
    RestrictSize = True
End If
End Function

Public Function MenuSelected(Index As Long, hMenu As Long, lParam As Long) As Boolean
' ========================================================================
' See CreateSubMenu routine for more information....
' A little tricky. Remember that a menu class can create another menu class (child class) for
' generated menus created from list/combo boxes. If we didn't create the submenu ID passed here,
' then we don't process the menu command. If we did create a gMenu, then its data only exists
' in the gMenu collection (child class) not in this normal collection within this class.
' So we need to pass these parameters along to the child class. The child class menu items
' were created with 2 special flags to help update the list/combo box: gControl which is the
' handle to the list/combo box and gType which is the control's style.

' So what do we have?
' the gMenu special flag LIndex is the ListIndex value (this was added to the menu item caption)
' the gMenu item's gControl reference= hWnd to the list/combobox
' the gMenu item's gType reference= simple list/combo or multiselect listbox
' That's all we need to update the list/combobox

' This routine also processes the 7 custom menus
' ========================================================================

Dim mIndex As Long, pParam1 As Long, pParam2 As Long
Dim msgParam As String, sUserID As String, sDefValue As String
Dim I As Integer, sTarget As String, sValue As String, wCaption As String

' windows messages used to fire a click event on a list/combo box
Const CB_SETCURSEL As Long = &H14E
Const LB_SETCURSEL As Long = &H186
Const LB_SETSEL As Long = &H185
' the color dialog box constants used
Const CC_RGBINIT         As Long = &H1
Const CC_FULLOPEN        As Long = &H2
Const CC_ANYCOLOR        As Long = &H100

On Error Resume Next
If lParam Then    ' sent from parent class, otherwise it is zero when sent as a result of clicking on a menu
    mIndex = cItems(Index & "." & hMenu)        ' see if menu item in our collection
    If Err = 0 Then
        ' the menu exists in this child class
        MenuSelected = True
        'Debug.Print "Testing click on custom menu"; hMenu; Index
        ' the gMenu could be a list/combo box or a custom menu.
        ' we'll check for custom menu first
        SeparateCaption vItems(mIndex).Caption, "", wCaption, ""
        For I = 1 To 7  ' custom menus will have these undocumented tags
            sTarget = Choose(I, "LColor:", "LMonth:", "LState:", "LDay:", "LDate:", "LFont:", "LDrv:")
            ReturnComponentValue wCaption, sTarget, sValue
            If Len(sValue) Then Exit For
        Next
        If Len(sValue) Then ' we have a custom menu
            ' extract the user supplied ID
            ReturnComponentValue wCaption, "lvID:", sUserID
            ' build a reference to the user's cTip class
            Dim oTipClass As cTips
            CopyMemory oTipClass, TipCallBackProc, 4&
            ' simply send the menu selection to the class
            Select Case sTarget
                Case "LColor:"
                    ' if value is -1, then user wants a custom color
                    If Val(sValue) = -1 Then
                        Dim CC As CHOOSECOLORSTRUCT, cCusGray As Long, lGrays(0 To 15) As Long
                        ' API requires a value for the .lpCustColors property, so
                        ' instead of supplying all black values, we'll give some grays
                        For cCusGray = 240 To 15 Step -15
                           lGrays((cCusGray \ 15) - 1) = RGB(cCusGray, cCusGray, cCusGray)
                        Next
                        ' see if there was a checked color & if so, show that in color dialog
                        ReturnComponentValue wCaption, "LDefClr:", sDefValue
                        With CC         'set the flags
                          .hInstance = App.hInstance
                          If Val(sDefValue) > -1 Then   ' set checked color
                            .Flags = CC_RGBINIT
                            .rgbResult = Val(sDefValue)
                          Else                          ' not color set
                            .Flags = CC_ANYCOLOR
                          End If
                          ' finish up the structure & send to API
                          .Flags = .Flags Or CC_FULLOPEN
                          .lStructSize = Len(CC)
                          .hwndOwner = FormHwnd
                          .lpCustColors = VarPtr(lGrays(0))
                        End With
                        If ChooseColor(CC) = 0 Then
                            MenuSelected = False        ' if user cancelled out
                        Else
                            oTipClass.SendCustomSelection sUserID, "Color", CC.rgbResult
                        End If
                    Else
                        oTipClass.SendCustomSelection sUserID, "Color", Val(sValue)
                    End If
                Case "LDay:": oTipClass.SendCustomSelection sUserID, "WeekDay", Val(sValue)
                Case "LMonth:": oTipClass.SendCustomSelection sUserID, "Month", Val(sValue)
                Case "LState:": oTipClass.SendCustomSelection sUserID, "State", sValue
                Case "LDate:": oTipClass.SendCustomSelection sUserID, "DayOfMonth", CDate(sValue)
                Case "LFont:": oTipClass.SendCustomSelection sUserID, "Font", sValue
                Case "LDrv:": oTipClass.SendCustomSelection sUserID, "Drive", Left$(sValue, 3)
            End Select
            CopyMemory oTipClass, 0&, 4&
            Set oTipClass = Nothing
        Else    ' no custom menu, should be a list/combo box control then
            If IsNumeric(vItems(mIndex).gControl) = False Then Exit Function
            If vItems(mIndex).gControl = 0 Then Exit Function ' it is not a control
            If IsWindow(vItems(mIndex).gControl) = 0 Then Exit Function
            ReturnComponentValue wCaption, "LIndex:", sValue
            If vItems(mIndex).ControlType = 2 Then          ' style of control
                ' with multiselect listboxes, the listindex is passed as the
                ' 2nd parameter, otherwise listindex is passed as 1st one
                pParam2 = Val(sValue)
                ' if item was not selected, we will mutliselect it now
                ' otherwise it will no longer be selected (pParam1 = 0)
                If ((vItems(mIndex).Status And lv_mChk) <> lv_mChk) Then pParam1 = 1
            Else
                pParam1 = Val(sValue)
            End If
            If Err Then Exit Function      ' couldn't find it, bye-bye
            ' determine the proper message to send to the list/combobox & update the control
            Select Case vItems(mIndex).ControlType
            Case 0: msgParam = CB_SETCURSEL
            Case 1: msgParam = LB_SETCURSEL
            Case 2: msgParam = LB_SETSEL
            End Select
            With vItems(mIndex)
                SendMessage .gControl, msgParam, pParam1, ByVal pParam2
                ' the above will select/unselect the list/combobox item, but it does not fire a "Clicked" event for the item
                ' the next call fires the "Click" event. WM_Command wparam is LoWord of control ID, HiWord of message (1=list change)
                SendMessage GetParent(.gControl), WM_COMMAND, MakeLong(CInt(GetWindowLong(.gControl, GWL_ID)), 1), ByVal .gControl
            End With
        End If
    End If
Else   ' sent from the message processor
    ' pass the menu ID, submenu ID and ListIndex to the gMenu.
    ' If we don't have a gMenu, then error will just pass the command back to Windows
    MenuSelected = gMenus("g" & hMenu).MenuSelected(Index, hMenu, 1)
End If
Err.Clear       ' clear any errors
End Function

Public Sub GetMenuItem(ID As Long, hMenu As Long)
Dim MenuData As MenuComponentData, Index As Long
' ========================================================================
' Function pouplates a general use MenuComponentData structure with the requested menu item
' This is called each time a menu item is measured for display, selected, or unselected
' System menu will return a hMenu of zero
' ========================================================================

On Error Resume Next
XferMenuData = MenuData                        ' blank out the general use -- just in case
'Debug.Print "retrievining item for "; hMenu; ID
Index = cItems(ID & "." & hMenu)               ' find our copy in our array
If Err = 0 Then
    XferMenuData = vItems(Index)               ' and if so, then we set the general use structure
Else
    gMenus("g" & hMenu).GetMenuItem ID, hMenu   ' see if a child class has the item
End If
Err.Clear
End Sub

Public Sub UpdateMenuItems(hMenu As Long)
' ========================================================================
' General use. Update any menu items, do clean up, whatever.
' This routine called each time the user exits the menu loop.

' For now, the only needed action is to force a remeasure of the Sidebar, if it exists
' ========================================================================

Dim MIS As MENUITEMINFO_string, Looper As Integer
Dim pIndex As Long, mMode As Integer
Dim bUpdated As Boolean, nrMenus As Integer
Dim Index As Long
' first see if we have this panel in this form -- should never error
' cause one is always created; but just in case
On Error Resume Next
pIndex = cPanels("p" & hMenu)
If Err Then
    Err.Clear
    Exit Sub
End If
' now see if the panel has a sidebar. If it does, the SidebarMenuItem will not be zero
If vPanels(pIndex).SidebarMenuItem = 0 Then Exit Sub

' now let's retrieve the array index to the menu item which is the sidebar
On Error GoTo AbortMenuUpdate
Index = cItems(vPanels(pIndex).SidebarMenuItem & "." & hMenu)
MIS.cbSize = Len(MIS)
' we want to update the menu item's ID, Enabled state and caption (type)
MIS.fMask = MIIM_ID Or MIIM_STATE Or MIIM_TYPE
MIS.fType = 0       ' equivalent to MFT_STRING
If Len(vItems(Index).Cached) Then
    MIS.dwTypeData = vItems(Index).Cached
Else
    MIS.dwTypeData = vItems(Index).Caption
End If
MIS.cch = Len(MIS.dwItemData)
MIS.dwItemData = hMenu          ' reset the submenu crossreference
MIS.wID = vItems(Index).ID      ' reapply the old menu ID
' if the caption was disabled before it was removed, ensure it stays disabled
If ((vItems(Index).Status And lv_mDisabled) = lv_mDisabled) Then MIS.fState = MF_DISABLED
If ((vItems(Index).Status And lv_mSBarHidden) = lv_mSBarHidden) Then
    ' if it was removed from the menu, its status flag would include lv_mSBarHidden
    vItems(Index).Status = vItems(Index).Status And Not lv_mSBarHidden
Else
    ' This is done vs simply updating the menu to force windows to remeasure the sidebar. If we don't do
    ' this the sidebar may never be remeasured.  Simply changing the type of non-OwnerDrawn doesn't
    ' do the trick under Win98/ME
    RemoveMenu hMenu, MIS.wID, 0
End If
' now simply add it to the top of the menu list
If InsertMenuItem(hMenu, 0, True, MIS) Then bUpdated = True

If bUpdated Then DrawMenuBar FormHwnd
AbortMenuUpdate:
End Sub

Public Sub GetPanelItem(hMenu As Long)
' ========================================================================
' Function pouplates a general use MenuComponentData structure with the requested submenu ID
' This is called each time a menu item is selected or unselected
' ========================================================================

On Error Resume Next
Dim PanelInfo As PanelData, Index As Long
XferPanelData = PanelInfo                       ' blank out the general use -- just in case
Index = cPanels("p" & hMenu)                ' find our copy in the our array
If Err = 0 Then
    XferPanelData = vPanels(Index)          ' if found, then we set the general use structure
Else
    gMenus("g" & hMenu).GetPanelItem hMenu  ' see if a child class has the item
End If
Err.Clear
End Sub

Public Function IsWindowList(hSubMenu As Long, bIsSysMenu As Boolean) As Boolean
' ========================================================================
' Primary function to add or update owner drawn menus to our class
' The main message processor prevents calling this routine while the user
' is in a menu loop and the menu has been displayed already;
' otherwise it is called each time a menu is initially displayed and again
' each time the menu loop is closed and the menu is then displayed again
' Exception: Windows Lists are processed completely every single time
' ========================================================================

Dim lMenus As Long, lState As Long, Looper As Integer
On Error Resume Next
' since our gMenus (generated menus from list/comboboxes) are processed in a
' separate child class, we test to see if the menu being processed belongs to
' a child class, if any

' Never create a menu item in a list/combobox that would create another
' submenu item referring to a list/combo box or custom menu. Doing so
' would produce undesirable effects. These routines are not designed to
' nest gMenu items.

' simple test to see if we have a gMenu item for the submenu being processed
lMenus = gMenus("g" & hSubMenu).ShowTips
If Err = 0 Then
    ' no error? then child class exists, send menu to that class
    gMenus("g" & hSubMenu).IsWindowList hSubMenu, False
    Exit Function
End If
Err.Clear

Dim OverallHeight As Long, Height As Long, maxHeight As Long
Dim MI() As Byte, Items2Check() As Long
Dim MII As MENUITEMINFO
Dim lType As Long, lStatus As Long, CachedMenus As Long
Dim sCaption As String, sNoScroll As String
Dim bBarBreak As Boolean, NewSubmenuID As Long
Dim bSideBarItem As Boolean, bHasSideBar As Boolean
Dim PageIDs() As Long, bNoOwnerDrawn As Boolean

On Error Resume Next
'Debug.Print "Processing submenu panel "; hSubMenu
' test this once at this routine & pass to GetMenuMetrix.
' More efficient then testing it for each submenu item
CachedMenus = modMenus.MenuCaptionListBox
If CachedMenus Then
    ' user is using a listbox to store caption flags
    If IsWindow(CachedMenus) = 0 Then CachedMenus = 0
End If
' get the number of submenu menu items & loop thru each subitem
' we have to do this at a minimum, 'cause VB will take ownership
' back of any menu that has changed (i.e., visibility, etc)
lMenus = GetMenuItemCount(hSubMenu)
ReDim Items2Check(0 To lMenus)  ' array of items being processed
ReDim PageIDs(0 To 0)           ' array of items forcing a new menu column
maxHeight = Screen.Height / Screen.TwipsPerPixelY - 7

For Looper = 0 To lMenus - 1
    ' by referencing the dwTypeData as a byte array vs long or string,
    ' we bypass the VB crash that happens on Win98 & XP & probably others
    bNoOwnerDrawn = False   ' flag to include/exclude MF_OWNERDRAWN
    ReDim MI(0 To 1023)
    MII.cbSize = Len(MII)
    MII.fType = 0
    MII.fMask = MIIM_ID Or MIIM_STATE Or MIIM_TYPE Or MIIM_SUBMENU Or MIIM_DATA
    MII.dwTypeData = VarPtr(MI(0))
    MII.cch = UBound(MI)
    ' get the submenu item information
    GetMenuItemInfo hSubMenu, Looper, True, MII
    ' whether or not this is a system menu will be passed to this routine, via the bIsSysMenu parameter
    ' but we will check anyway. System menu items have IDs > &HF000&
    If MII.wID >= &HF000& And MII.hSubMenu = 0 Then
        bIsSysMenu = True
    Else
        ' the following most likely is a Windows List & we want to always process these since
        ' windows will take ownership back every single time it is displayed.  By setting the
        ' function return value to True, the core Message Processor function will never bypass
        ' this windows list submenu any time it appears.
        If MII.wID > &H7FFF And MII.hSubMenu = 0 Then IsWindowList = True
    End If
    Items2Check(Looper + 1) = MII.wID   ' add to array of items to check
    lStatus = 0                         ' reset the status
    ' now set/remove some flags & stuff to return the caption, checked & enabled status, etc
    ' first we look for the current status of the menu item
    If ((MII.fState And MF_DISABLED) = MF_DISABLED) Or ((MII.fState And MF_GRAYED) = MF_GRAYED) Then lStatus = lStatus Or lv_mDisabled
    If ((MII.fState And MF_CHECKED) = MF_CHECKED) Then lStatus = lStatus Or lv_mChk
    If ((MII.fState And MF_DEFAULT) = MF_DEFAULT) Then lStatus = lStatus Or lv_mDefault
    If MII.hSubMenu Then lStatus = lStatus Or lv_mSubmenu
    
    ' now we remove any previous column breaks. These would have been created if we have a scrolling menu
    ' and the user is showing a sidebar. Since menu items can be removed without me knowing (i.e., setting
    ' the menu visibility to False), we have no way to determine if these column breaks are still valid -- so we remove all of them
    If ((MII.fType And MF_MENUBARBREAK) = MF_MENUBARBREAK) Then
        MII.fType = MII.fType And Not MF_MENUBARBREAK
        bNoOwnerDrawn = True    ' we don't want MF_OWNERDRAWN for now
    End If
    If ((MII.fType And MF_MENUBREAK) = MF_MENUBREAK) Then
        MII.fType = MII.fType And Not MF_MENUBREAK
        bNoOwnerDrawn = True    ' we don't want MF_OWNERDRAWN for now
    End If
    If bIsSysMenu Then bNoOwnerDrawn = True ' we don't want MF_OWNERDRAWN for now
    lType = MII.fType
    lState = MII.fState
    sCaption = Trim$(Replace$(StrConv(MI, vbUnicode), Chr$(0), ""))    ' get the menu item caption
    'Debug.Print "caption = "; sCaption; "<"; hSubMenu; MII.wID; MII.fState
    
    ' lState will return some of the new fState attributes & lType returns new fType attributes
    ' If no major changes to a processed menu item, the next function returns within a few lines of code
    GetMenuMetrix sCaption, lStatus, lType, lState, hSubMenu, MII.wID, Looper, bSideBarItem, Height, NewSubmenuID, CachedMenus
    If bBarBreak = True Then    ' used in conjunction with the "If bSideBarItem" statement below
        ' if a sidebar is used, it gets placed in its own panel, this allows us to easily
        ' place it and we can even make it  clickable. Therefore item following sidebar
        ' starts a new column. We tag the menu id as starting a new column
        ReDim Preserve PageIDs(0 To UBound(PageIDs) + 1)
        PageIDs(UBound(PageIDs)) = MII.wID  ' add to array of new columns
        bBarBreak = False                   ' reset some flags
        OverallHeight = 0                   ' reset height for next panel
        bNoOwnerDrawn = True
    End If
    If bSideBarItem = True And bHasSideBar = False Then
        ' flag indicating that the next menu item after the sidebar
        ' must be a menu break to start a new column
        bBarBreak = True            ' flag indicating next item is a column break
        PageIDs(0) = MII.wID        ' identify which ID is the sidebar item
        bHasSideBar = True          ' flag indicating sidebar in use
        bNoOwnerDrawn = True
        Height = 0
    End If
    ' set for all non-sidebar/sys menu items
    If Not bNoOwnerDrawn Then lType = lType Or MF_OWNERDRAW
    ' calculate the overall height of the menu panel. This is only needed/used
    ' when a sidebar is being used for this menu panel
    OverallHeight = OverallHeight + Height
    If OverallHeight > maxHeight And bHasSideBar = True Then
        ' Curses on windows inflexibility. Menus scroll if there are too many
        ' menu items to display on screen. Well, if you insert a menubreak
        ' guest what? You loose the scroll functionality completely, even if the next column
        ' of menu items goes off the screen
        ' To get past that, we need to keep track of the estimated panel height
        ' and add any additional menubreaks ourselves
        ReDim Preserve PageIDs(0 To UBound(PageIDs) + 1)
        PageIDs(UBound(PageIDs)) = MII.wID  ' keep track of which IDs will start a new column/page
        If UBound(PageIDs) = 2 Then maxHeight = OverallHeight - Height
        OverallHeight = Height                   ' reset height for next panel
    End If
    MII.fMask = 0
    ' save updates to allow us to draw the menu item
    If NewSubmenuID <> 0 Or MII.hSubMenu <> 0 Then
        MII.fMask = MIIM_SUBMENU
        ' this menu item forced creation of a gMenu if NewSubmenuID<>0
        ' we need to change it to a menu having a submenu
        If NewSubmenuID <> 0 Then MII.hSubMenu = NewSubmenuID
    End If
        ' this helps us coordinate classes with menu items, without it? #&@!@
        ' only needed for the DoDrawItem & DoMeasureItem since Windows does not
        ' pass the submenu ID for ownerdrawn menus
        MII.dwItemData = hSubMenu   ' identifies which submenu menu item is on
        MII.fState = lState         ' menu item state
        MII.fType = lType           ' menu item type
        MII.fMask = MII.fMask Or MIIM_ID Or MIIM_STATE Or MIIM_DATA Or MIIM_TYPE
        SetMenuItemInfo hSubMenu, MII.wID, False, MII      ' update the menu item
        ' here's where we force windows to remeasure system menu items, by saving it again!
        If bNoOwnerDrawn Then
            MII.fType = MII.fType Or MF_OWNERDRAW
            SetMenuItemInfo hSubMenu, MII.wID, False, MII
        End If
Next
If bHasSideBar Then
' now if a sidebar is being shown, the user has an option to not show it if the menu scrolls since the
' sidebar forces columns on the menu. Well we see if that option has been selected
    With vItems(cItems(PageIDs(0) & "." & hSubMenu))
        ' get the coded part of the sidebar caption & test for the NoScroll option
        SeparateCaption .Caption, "", sCaption
        ReturnComponentValue sCaption, "NOSCROLL", sNoScroll
        If UBound(PageIDs) > 1 And Len(sNoScroll) > 0 Then
        ' scrolling menu & user doesn't want sidebar to show
            ' flag indicating we are deleting menu item but want it back
            .Status = .Status Or lv_mSBarHidden
            ' removing sidebar for scrolling menu
            RemoveMenu hSubMenu, PageIDs(0), 0
            ' we want no column breaks
            ReDim PageIDs(0 To 0)
        End If
    End With
    ' if menu will scroll we need to add menu breaks
    ' update the 1st PageID to be a menubreak
    ' then update remaining to be menubarbreaks
    For Looper = 1 To UBound(PageIDs)
        ReDim MI(0 To 1023)
        MII.cch = UBound(MI)
        MII.fMask = MIIM_TYPE
        MII.fType = 0
        GetMenuItemInfo hSubMenu, PageIDs(Looper), False, MII
        If Looper = 1 Then
            MII.fType = MII.fType Or MF_OWNERDRAW Or MF_MENUBREAK
        Else
            MII.fType = MII.fType Or MF_OWNERDRAW Or MF_MENUBARBREAK
        End If
        SetMenuItemInfo hSubMenu, PageIDs(Looper), False, MII
    Next
End If
' now add any cx offsets to each of these menu items & calculate sidebar height
GetPanelMetrix hSubMenu, bIsSysMenu, Items2Check(), PageIDs()
Erase Items2Check
Erase PageIDs
DrawMenuBar FormHwnd       ' not needed yet, next version will take control of the menu bar
End Function

Private Sub GetMenuMetrix(Caption As String, Status As Long, NewType As Long, NewState As Long, hMenu As Long, _
    MenuID As Long, menuPos As Integer, IsSideBar As Boolean, _
    ItemHeight As Long, NewSubmenuID As Long, AltMenuSource As Long)
' ========================================================================
' Function to parse caption components, determine menu attributes,
' calculate menu item widths/heights, build gMenus etc
' ========================================================================
Dim bChildClass As Boolean
Dim IdX As Integer, NewIndex As Integer
Dim wCaption As String, sTarget As String
Dim sValue As String, CachedCaption As String
Dim mComponent As MenuComponentData, frmObject As Control

On Error Resume Next
' we'll see if the menu item is already in our local array
NewIndex = cItems(MenuID & "." & hMenu)
IsSideBar = False       ' reset
NewSubmenuID = 0        ' reset
If AltMenuSource And Len(Caption) > 0 Then
' user is using a listbox to store caption flags
    'these need to be rechecked each time to see if the flags in the
    ' listbox has changed
    Dim sHotKey As String
    SeparateCaption Caption, CachedCaption, wCaption, sHotKey
    ReturnComponentValue wCaption, "Cache:", sValue
    If Len(sValue) Then
        mComponent.Cached = Caption
        IdX = SendMessage(AltMenuSource, LB_GETTEXTLEN, Val(sValue), ByVal 0&)
        sTarget = String$(IdX, 0)
        SendMessage AltMenuSource, LB_GETTEXT, Val(sValue), ByVal sTarget
        Caption = CachedCaption & StringFromBuffer(sTarget)
        If Len(sHotKey) Then Caption = Caption & vbTab & sHotKey
    End If
End If
If NewIndex > 0 Then
    ' we have a copy of it, is it the same? Otherwise, we measure everything again
    mComponent = vItems(NewIndex)
    ' captions with OwnerDrawn menus in NT disappear after we take ownership
    ' So our adjusted logic: if it's blank and we have in our array, then it hasn't changed.
    If StrComp(Caption, mComponent.Caption, vbBinaryCompare) = 0 Or Caption = "" Then
        ' it's the same, but if the following menu item options occur
        ' we remeasure because these could affect height/width of an item
        ' We can check for instances when we don't need to remeasure based on default status
        ' 1. If item is not default and it was before but isn't now, menu size is ok
        ' 2. If the default menu item status didn't change
        Caption = mComponent.Caption
        SeparateCaption Caption, "", wCaption, ""
        If (((Status And lv_mDefault) = lv_mDefault) Or ((mComponent.Status And _
            lv_mDefault) = (Status And lv_mDefault))) Then
            With mComponent
                ' here we remove the previous status and add them back if needed
                If ((.Status And lv_mDisabled) = lv_mDisabled) Then .Status = .Status And Not lv_mDisabled     ' remove the disabled property
                If ((.Status And lv_mChk) = lv_mChk) Then .Status = .Status And Not lv_mChk    ' remove the checked property
                If ((.Status And lv_mDefault) = lv_mDefault) Then .Status = .Status And Not lv_mDefault    ' remove the default item property
                .Status = .Status Or Status     ' reset to passed properties
            End With
            ' Still not out of the woods. If the menu item references a
            ' list/combobox, then we need to redo its submenu since the
            ' list/combobox contents/listindex could have changed
            ' Also if this was a custom menu, we redo it since custom menus
            ' have no permanent properties to refer back to.
            If mComponent.hControl = 0 And ((mComponent.Status And lv_mCustom) <> lv_mCustom) Then
                If ((mComponent.Status And lv_mSBar) = lv_mSBar) Then
                    IsSideBar = True
                    mComponent.HotKey = ""
                Else
                    ' No major change. Finalize some properties/attributes and return
                     If ((mComponent.Status And lv_mSep) = lv_mSep) Then NewType = NewType Or MF_SEPARATOR
                End If
                ' we got all the stuff we need and don't need to process much further
                ItemHeight = mComponent.Dimension.Y ' return HT of item
                vItems(NewIndex) = mComponent       ' refresh array item
                GoTo DoIconReference                ' ensure image handle didn't change
            Else    ' menu refs a list/combobox, so we redo it
                GoTo DoCustomMenus   ' skip majority of the process
            End If
        End If
    End If
End If
With mComponent
    .Caption = Caption      ' store original caption
    .Index = menuPos        ' zero-based position on submenu
    .Status = Status        ' all status, mine & windows
    SeparateCaption Caption, .Display, wCaption, sValue
    .Display = Replace$(.Display, "& ", "&& ") ' hack to display "&" in captions
    If Len(sValue) Then .HotKey = sValue
    ' since anything can have Tips, let's look for Tips first
    ' note: tips aren't normally included in Separators, but are allowed if
    '       manually entered in the menu caption
    ReturnComponentValue wCaption, "TIP:", sValue
    If Len(sValue) Then
        .Tip = sValue
    Else
        ' child classes may have this property set & if so, use it
        .Tip = tipDefault
    End If
    ReturnComponentValue wCaption, "SIDEBAR", sValue
    If Len(sValue) Then
        ' for sidebars we only accept the first one and it must be
        ' the first item in the menu panel, otherwise we force it to
        ' look and act like a separator bar
        If menuPos = 0 Then
            .Status = 0
            If ((NewType And MF_SEPARATOR) = MF_SEPARATOR) Then NewType = NewType And Not MF_SEPARATOR
            If ((NewState And MF_DISABLED) = MF_DISABLED) Then .Status = .Status Or lv_mDisabled
            If ((NewState And MF_CHECKED) = MF_CHECKED) Then .Status = .Status Or lv_mChk
            .Status = .Status Or lv_mSBar   ' flag indicating sidebar item
            ReturnComponentValue wCaption, "SBDisabled", sValue
            If Len(sValue) Then
                .Status = .Status Or lv_mDisabled
                NewState = NewState Or MF_DISABLED
            End If
            .HotKey = "Remeasure"       ' not displayed on sidebars
            IsSideBar = True            ' flag to return
            GoTo Check4ControlReference ' process sidebar information
        Else    ' another sidebar on same menu, we can't let it go
            .Display = "" ' reset caption which forces it to be a separator
        End If
    End If
    If Len(.Display) Then   ' see if the caption is a separator bar
        If Left$(.Display, 1) = "-" Then
            ' text separator bar -- update return status
            .Display = Mid$(.Display, 2)
            NewType = NewType Or MF_SEPARATOR
        End If
    Else
        ' no caption, then assume a separator bar
        NewType = NewType Or MF_SEPARATOR
    End If
    ' finally, we finalize the separator bar
    If ((NewType And MF_SEPARATOR) = MF_SEPARATOR) Then
        ' finish processing separator bar & then skip the rest
        .Status = 0
        NewState = 0
        ReturnComponentValue wCaption, "RAISED", sValue
        If Len(sValue) Then .Status = .Status Or lv_mSepRaised
        .Status = .Status Or lv_mSep Or lv_mDisabled   ' separator bar flag
        .HotKey = ""                    ' these can't have hotkeys
        .Icon = ""                       ' these can't have icons
        .hControl = 0                   ' these can't ref list/comboboxes
        NewState = NewState Or MF_DISABLED
        GoTo GetMeasurements
    Else    ' see if the menu item will be bolded
        ReturnComponentValue wCaption, "DEFAULT", sValue
        If Len(sValue) Then .Status = .Status Or lv_mDefault
    End If
    If Len(wCaption) = 0 Then GoTo GetMeasurements
    
    ' let's finish parsing out the working caption
    ' if a hotkey was provided in the coded caption, it overwrites
    ' the actual Menu Editor supplied hotkey, if any
    ReturnComponentValue wCaption, "HotKey:", sValue
    If Len(sValue) Then .HotKey = sValue
    ' now see if a combo/listbox will be used to create another submenu (gMenus)
    ' but if the menu item already has a submenu that isn't a gMenu,
    ' we don't allow it. We will not replace existing submenus.
    If ((.Status And lv_mSubmenu) = lv_mSubmenu) And .hControl <> 0 _
       Or ((.Status And lv_mSubmenu) <> lv_mSubmenu) Then
        ReturnComponentValue wCaption, "LB:", sValue
        If Len(sValue) Then     ' a combo box was chosen
            .ControlType = 1    ' flag indicating listbox, otherwise combobox
        Else    ' if not, test for a listbox
            ReturnComponentValue wCaption, "CB:", sValue
        End If
        If Len(sValue) Then ' if either was chosen, return handle to that control
            If IsNumeric(sValue) Then   ' ensure it's a valid window!
                If IsWindow(CLng(sValue)) Then .hControl = CLng(sValue)
            Else    ' control name passed vs its handle, validate it
                Set frmObject = SplitControlIndex(sValue)
                If Not frmObject Is Nothing Then
                    .hControl = frmObject.hWnd
                    Set frmObject = Nothing
                End If
            End If
        End If
    End If
    
Check4ControlReference:
    ' now we process gMenu items since this class can also be a child class
    ' containing gMenu items. gMenu items will have the undocumented tag gControl
    ReturnComponentValue wCaption, "gControl:", sValue
    If Len(sValue) Then
        bChildClass = True
        .gControl = Val(sValue)
        ' for gMenus that refer to list/combo boxes, this will be the type
        ReturnComponentValue wCaption, "gType:", sValue ' list or combo box
        If Len(sValue) Then .ControlType = Val(sValue)
        ReturnComponentValue wCaption, "LColor:", sValue 'color menu
        If Len(sValue) Then .Status = .Status Or lv_mColor
        ReturnComponentValue wCaption, "LFont:", sValue ' font menu
        If Len(sValue) Then .Status = .Status Or lv_mFont
        ReturnComponentValue wCaption, "LState:", sValue ' state menu
        If Len(sValue) Then .HotKey = sValue
        ReturnComponentValue wCaption, "gEXE:", sValue ' filename item
        If Len(sValue) Then
            ' the Icon becomes the filename so DrawMenuIcon knows to
            ' look for the icon in the file
            .Icon = GetShortFileName(sValue & .Display)
            ' here we just standardize what's displayed on the menu
            sValue = StripFile(.Display, "E")
            .Display = StripFile(.Display, "m")
            ' change all LCase filenames to UCase. But if mixed case, leave as is
            If StrComp(LCase(.Display), .Display, vbBinaryCompare) = 0 Then .Display = UCase(.Display)
            .Display = .Display & "." & LCase(sValue)
            ' a nice touch. We add the file description as a tip
            .Tip = GetFileDescription(.Icon)
            If .Tip = "" Then .Tip = tipDefault
        End If
        ReturnComponentValue wCaption, "LDrv:", sValue  ' drives menu
        If Len(sValue) Then
            IdX = InStr(sValue, "\")
            If IdX Then .Icon = Left$(sValue, IdX) Else .Icon = sValue  ' the value will be A:\, C:\, etc
            If .Tip = "Removable Media" Then
                ' for floppies, we get the cached icon. Otherwise program will
                ' check the floppy each time it draws this icon--unnecessary
                ' repeated reads to the floppy & time-consuming, especially if it's empty
                .Icon = GetFloppyIcon(.Icon)
                .Display = .Display & " Removable"
            Else    ' non floppies. We'll include the volume name with the drive
                If Len(.Icon) > IdX Then IdX = 1 Else IdX = 0
                sTarget = String$(255, 0)
                sValue = String$(55, 0)
                GetVolumeInformation .Icon, sTarget, 255, 0, 0, 0, sValue, 55
                sTarget = StringFromBuffer(sTarget)
                If Len(sTarget) Then
                    If .Tip = "Network/Mapped Drive" Then
                        .Display = .Display & "  (" & sTarget & ")"
                    Else
                        .Display = .Display & " " & sTarget
                    End If
                End If
            End If
        End If
    End If
    ' Remove any other conflicting status codes for submenu items
    If ((.Status And lv_mSubmenu) = lv_mSubmenu) Then
        .HotKey = ""       ' no hotkeys displayed for submenus
        ' but allow everything else, including check marks if user wants them
    End If
End With
GetMeasurements:
If ((mComponent.Status And lv_mSBar) = lv_mSBar) = False Then   ' skip for sidebars (for now)
    ' =====================================================================
    ' Routine gets the meaurements of the submenu items
    ' =====================================================================
    Dim tDC As Long, tRect As RECT
    tDC = GetDC(FormHwnd)
    With mComponent
        'Debug.Print "measuring "; .Display
        ' ========================================================================
        ' Calculate heights needed
        ' ========================================================================
        ' supply the proper font for measuring
        If ((.Status And lv_mSep) = lv_mSep) Then   ' bar with text
            ApplyMenuFont 3, tDC                      ' use the smaller font for these
        Else         ' otherwise, use normal font or bold if needed
            ApplyMenuFont 1 + Abs(CInt((.Status And lv_mDefault) = lv_mDefault)), tDC
        End If
        ' now we calculate the height & width of the text needed
        If ((.Status And lv_mFont) = lv_mFont) Then ' an lvFonts custom menu
            sTarget = .Display
            ' with the DT_CALCRECT flag, no text is actually drawn, it only calculates the rectangle passed to it
            DrawText tDC, sTarget, Len(sTarget), tRect, DT_CALCRECT Or DT_LEFT Or DT_SINGLELINE Or DT_NOCLIP
            ' store height/width for now, we'll use the largest when compared against the sample font
            .Dimension.X = tRect.Right
            .Dimension.Y = tRect.Bottom
            ApplyMenuFont 0, tDC    ' replace DC font
            ReturnComponentValue wCaption, "LFont:", sValue
            ' now create and load a "Sample" using the actual font name
            CreateDestroyMenuFont True, False, sValue
            ApplyMenuFont 6, tDC
            ' the word Sample is the hotkey
            sTarget = .HotKey
            DrawText tDC, sTarget, Len(sTarget), tRect, DT_CALCRECT Or DT_LEFT Or DT_SINGLELINE Or DT_NOCLIP
            ' here we ue the largest setting between the non-sample font & sample font
            tRect.Right = tRect.Right + .Dimension.X
            If tRect.Bottom > .Dimension.Y Then tRect.Bottom = .Dimension.Y
        Else
            sTarget = .Display & .HotKey    ' want total width first
            ' with the DT_CALCRECT flag, no text is actually drawn, it only calculates the rectangle passed to it
            DrawText tDC, sTarget, Len(sTarget), tRect, DT_CALCRECT Or DT_LEFT Or DT_SINGLELINE Or DT_NOCLIP
        End If
        If ((.Status And lv_mSep) = lv_mSep) = True And Len(.Display) = 0 Then
            tRect.Bottom = 4        ' standard separators will be hardcoded as a rect 4 pixels high
            .Dimension.X = tRect.Right
        Else
            ' give a 1 pixel separation top/bottom from next menu items
            ' this is more for icons then the text
            If ((.Status And lv_mSep) <> lv_mSep) Then
                ' ensure height is at least 16 + 4 pixel edge for icons
                ' -- 16 for icon, 1 pix for frame, 1 pix separation btwn frame and icon
               If tRect.Bottom < 20 Then tRect.Bottom = 20
            End If
            ' 20 pixel buffer on both sides of caption for checks/icons & submenu arrows
            .Dimension.X = tRect.Right + 40
        End If
        .Dimension.Y = tRect.Bottom + 2
        ApplyMenuFont 0, tDC ' return to default menu font
    End With
    ReleaseDC FormHwnd, tDC
End If

DoCustomMenus:
' ========================================================================
' Calls function to create submenus if this item is ref'ing list/combobox
' or one of the several custom menus (lvColors, lvDrives, etc)
' ========================================================================
With mComponent
    ' for custom menus, we look for these flags in the caption
    For IdX = 1 To 7
        sTarget = Choose(IdX, "lvColors:", "lvMonths:", "lvDays:", "lvStates:", "lvMonth:", "lvFonts:", "lvDrives:")
        ReturnComponentValue wCaption, sTarget, sValue
        If Len(sValue) Then Exit For
    Next
    If Len(sValue) Then
        ' one of the custom tags was found and we will process that
        CreateSubMenuCustom MenuID, menuPos, hMenu, NewSubmenuID, sTarget, sValue, .Tip
        .Status = .Status Or lv_mCustom ' flag to force reprocess if shown again
    Else    ' otherwise we test to see if this will ref a combo/listbox
        If .hControl Then   ' call function to create a new submenu on the fly
            CreateSubMenu MenuID, menuPos, hMenu, .ControlType, .hControl, NewSubmenuID, wCaption
        End If
    End If
    If NewSubmenuID Then ' only update following if a submenu was created
        .Status = .Status Or lv_mSubmenu
        If Len(.HotKey) Then
            ' if the menu item had a displayed hotkey, we need to remove
            ' the width it was taking up 'cause submenu menus don't have 'em
            sTarget = .HotKey
            tDC = GetDC(FormHwnd)
            ApplyMenuFont 1, tDC
            ' with the DT_CALCRECT flag, no text is actually drawn, it only calculates the rectangle passed to it
            DrawText tDC, sTarget, Len(sTarget), tRect, DT_CALCRECT Or DT_LEFT Or DT_SINGLELINE Or DT_NOCLIP
            .Dimension.X = .Dimension.X - tRect.Right
            ApplyMenuFont 0, tDC
            ReleaseDC FormHwnd, tDC
        End If
        .HotKey = ""
    End If
End With
DoIconReference:
' ref to menu image. This needs to be rechecked every single time to ensure the handle for the
' source image didn't change. FYI: ImageList's ExtractIcon function will change this handle
  With mComponent
    ' excpetions are separator bars (don't have icons) and custom menus where icons already determined
    .ShowBKG = False
    If bChildClass = False And ((.Status And lv_mSep) <> lv_mSep) Then
        .Icon = CStr(ValidateImage(wCaption, IdX))
        If .Icon = "0" Then
            .Icon = ""
        Else
            ' see if image is forced to not be transparent
            ' << n/a for Icons, only valid for Bitmaps (IDX=1) >>
            ReturnComponentValue wCaption, "IMGBKG", sValue
            If Len(sValue) And IdX = 1 Then .ShowBKG = True
        End If
    End If
  End With
' ========================================================================
' The End.  Simply update the array and index
' ========================================================================
If NewIndex = 0 Then                        ' new menu item, we need an index to the array entry
    NewIndex = UBound(vItems) + 1
    ReDim Preserve vItems(0 To NewIndex)
    cItems.Add NewIndex, MenuID & "." & hMenu
End If
mComponent.ID = MenuID
vItems(NewIndex) = mComponent
' if a gMenu or custom menu was created, we need to cross-reference it
' by the new submenu handle in order for the routines herein to correctly
' identify the menu item
If NewSubmenuID Then cItems.Add NewIndex, NewSubmenuID & "." & hMenu
ItemHeight = mComponent.Dimension.Y     ' return item height
'Debug.Print FormHwnd; " current class vArray count="; UBound(vItems)
End Sub

Private Sub GetPanelMetrix(hMenu As Long, bIsSysMenu As Boolean, _
    Items2Check() As Long, PageBreaks() As Long)
' ========================================================================
' Function primarily tallies up the overall height of menu items and bumps
' that up against the existing height of a sidebar (if used). This will
' determine if we need to reprocess a sidebar.
' Function also calculates menu item widths and offsets to display a more
' uniform menu across all Window systems.
' ========================================================================

Dim Looper As Integer, lItem As Long
Dim NewIndex As Long, NextPage As Long
Dim IdxStart As Long, IdxStop As Long
Dim tIndex As Integer, lSideBar As Long
Dim Xoffset As Integer, Win98MEoffset As Integer
Dim pData As PanelData
Dim maxHeight As Integer, MaxWidth As Integer
Dim bHasTabs As Boolean, bHasSubs As Boolean
Dim bHasHotKey As Boolean, bHideSidebarCheck As Boolean
Dim bNonSbar As Boolean, bReprocessSideBar As Boolean

On Error Resume Next
' The Win98MEoffset is used currently for Win98 & WinME. These systems seem
' to tack on some extra pixels to menu items. We will account for those pixels
Win98MEoffset = modMenus.Win98MEoffset
' See if the menu panel already exists
NewIndex = cPanels("p" & hMenu)
pData.SubmenuID = hMenu
If Err Then
    NewIndex = 0
    Err.Clear
Else
    ' we copy the existing sidebar reference to the new pData structure
    ' otherwise we lose it & can't delete it when needed
    pData.PanelIcon = vPanels(NewIndex).PanelIcon
End If
IdxStart = 1                    ' range of items to apply offsets
IdxStop = UBound(Items2Check)
Items2Check(0) = 0              ' reset sidebar ID if any
' here we set a measuring point for the sidebar if we have multiple
' columns to display. The total height at this point will be the sidebar height
If UBound(PageBreaks) > 1 Then NextPage = 2
' now loop thru the subset, extracting key values

If UBound(PageBreaks) > 1 Then NextPage = 2
' now loop thru the subset, extracting key values

For Looper = 1 To UBound(Items2Check)
    tIndex = cItems(Items2Check(Looper) & "." & hMenu)
    With vItems(tIndex)
        ' we don't add or include sidebars in any of our calculations
      If ((.Status And lv_mSBar) <> lv_mSBar) Then
        If (Items2Check(Looper) = PageBreaks(NextPage)) And NextPage > 1 Then
            ' we hit a column break
            IdxStop = Looper - 1     ' reset last item to apply offsets
            GoSub UpdateItemOffsets  ' apply the offsets
            bHasSubs = False         ' reset flag
            bHasHotKey = False       ' reset flag
            MaxWidth = 0             ' reset
            ' here we store the MaxHeight value to use for the sidebar
            If NextPage = 2 Then Items2Check(0) = maxHeight
            ' now we set the next page to stop and apply offsets
            If NextPage > UBound(PageBreaks) Then NextPage = 0 Else NextPage = NextPage + 1
            IdxStart = Looper        ' reset start & stop indexes
            IdxStop = UBound(Items2Check)
            maxHeight = 0
        End If
            ' keep track of whether or not the panel has any icons at all,
            ' whether or not it is a system menu panel and the widest item on the panel
            If Len(.Icon) Then pData.HasIcons = True
            If ((.Status And lv_mSubmenu) = lv_mSubmenu) Then bHasSubs = True
            If .ID >= &HF000& And ((.Status And lv_mSubmenu) <> lv_mSubmenu) Then pData.IsSystem = True
            If Len(.HotKey) Then bHasHotKey = True
            ' running sum of estimated menu panel width & height
            If .Dimension.X > MaxWidth Then MaxWidth = .Dimension.X
            maxHeight = maxHeight + .Dimension.Y
        'End If
      End If
    End With
Next

If Items2Check(0) > 0 Then maxHeight = Items2Check(0)
If NextPage = 0 Then GoSub UpdateItemOffsets            ' apply offsets
' update flag to indicate this submenu is the system menu
If bIsSysMenu Then pData.IsSystem = True
' Now we append/update the panel array and collection
If NewIndex = 0 Then
    NewIndex = UBound(vPanels) + 1
    ReDim Preserve vPanels(0 To NewIndex)
    cPanels.Add NewIndex, "p" & hMenu
End If
If UBound(Items2Check) > 50 Then pData.Hourglass = True
vPanels(NewIndex) = pData

If bReprocessSideBar = True Then    ' we need to build another sidebar
    lItem = cItems(PageBreaks(0) & "." & hMenu)
    Dim wCaption As String, sValue As String
    ' first separate the working caption
    SeparateCaption vItems(lItem).Caption, "", wCaption
    ReturnComponentValue wCaption, "IMG:", sValue
    If Len(sValue) Then
        CreateImageSidebar lItem, NewIndex, wCaption
    Else
        CreateTextSidebar lItem, NewIndex, wCaption
    End If
    With vItems(lItem)
        .OffsetCx = -12 ' windows adds 12 pixels to columns
        ' in addition it Win98 adds another 35 pixels if the sidebar has a hotkey
        If InStr(.Caption, vbTab) Then .OffsetCx = .OffsetCx - Win98MEoffset
        'Debug.Print "sidebar width sent "; .Dimension.X
    End With
End If
Exit Sub

UpdateItemOffsets:
    ' here we adjust widths to add or remove pixels
    Xoffset = 0
    If bHasHotKey Then
        If Win98MEoffset = 0 Then Xoffset = 15 ' add 15 for non-Win98/ME
        ' add 20 for non-Win98/ME
        If bHasSubs = True And Win98MEoffset = 0 Then Xoffset = Xoffset + 20
    End If
    ' for system menus, we need to add another 15 regardless for non-Win98/ME
    If pData.IsSystem = True And Win98MEoffset = 0 Then Xoffset = Xoffset + 15
    
    ' loop back thru the collection, updating the width offset for each
    ' menu item being checked. This is only truly required when using
    ' sidebars and a menu is too tall for the screen. Because columns are
    ' created to negate scrolling, each panel can have different widths.
    ' Each item has an offset to be used by the DoMeasureItem routine
    For lItem = IdxStart To IdxStop
        tIndex = cItems(Items2Check(lItem) & "." & hMenu)
        With vItems(tIndex)
            If ((.Status And lv_mSBar) = lv_mSBar) Then ' sidebar item
                pData.SidebarMenuItem = .ID     ' ref sidebar in panelstruct
                ' if the estimated height changed, we need to build
                ' another sidebar bitmap to display
                If .Dimension.Y <> maxHeight Then bReprocessSideBar = True
                .Dimension.Y = maxHeight
                ' if the hotkey is coded this way, the caption changed somehow and
                ' we need to build another bitmap to reflect the change
                If .HotKey = "Remeasure" Then bReprocessSideBar = True
                .HotKey = ""
                pData.Accelerators = pData.Accelerators & Chr$(246)  'N/A for sidebars
            Else    ' non sidebar items
                ' if we are getting icons from filenames or have more than 50 menu items
                ' then we want to display the hourglass when the submenu opens these items for display
                If IsNumeric(.Icon) = False And Len(.Icon) > 0 Then pData.Hourglass = True
                If InStr(.Caption, "lFont:") > 0 And UBound(Items2Check) > 20 Then pData.Hourglass = True
                .OffsetCx = Xoffset
                ' each menu item carries the width of the widest menu item
                ' for the lvColors menu we have no Xoffset so we supply one
                If ((.Status And lv_mColor) = lv_mColor) Then .OffsetCx = .OffsetCx + 10
                ' here's where we build a list of accelerator keys to be used on systems like WinNT4 that
                ' will not process them for owner-drawn menus
                If ((.Status And lv_mSep) = lv_mSep) Then
                    pData.Accelerators = pData.Accelerators & Chr$(246) ' N/A for separator bars
                Else
                    pData.Accelerators = pData.Accelerators & ReturnAcceleratorKey(.Display)
                End If
            End If
        End With
    Next
Return
End Sub

Private Function ReturnAcceleratorKey(sInKey As String) As String
Dim I As Integer, Index As Integer
' find the 1st occurrence of an ampersand
I = InStr(Index + 1, sInKey, "&")
Do While I
    ' loop thru looking for double && which are not accelerator keys
    If Mid$(sInKey, I, 2) = "&&" Then
        Index = I + 1
        I = InStr(Index + 1, sInKey, "&")
    Else
        Exit Do
    End If
Loop
' now return the result
If I Then
    ReturnAcceleratorKey = UCase(Mid$(sInKey, I + 1, 1))
Else
    ' default place holder for a non-accelerator key
    ReturnAcceleratorKey = Chr$(246)
End If
End Function
Private Sub CreateSubMenu(MenuID As Long, mnuPos As Integer, ParentMenu As Long, _
    ControlType As Byte, hControl As Long, NewSubmenuID As Long, sCustomFlag As String)
' ========================================================================
' This is the function to bring a combo/listbox onto the menus
' NOTE: Not completely finalized yet as I would like to play with the idea
' for displaying custom drawn list/combos (i.e., those that display images)
' As of this writing, this function will not work for owner-drawn list/comboboxes
' ========================================================================

If hControl = 0 Then Exit Sub

Dim msgListCount As Long, msgListIndex As Long, msgTextLen As Long
Dim msgText As Long, msgItemData As Long, ItemData As Long
Dim ListCount As Long, ListIndex As Long, lLen As Long, lStyle As Long
Dim Looper As Integer, SelCount As Long
Dim bMultiSelect As Boolean, bShowHourglass As Boolean
Dim sBuffer As String, sParent As String, sIcon As String

' constants to get the information out of the list/comboboxes
Const CB_GETCOUNT As Long = &H146
Const CB_GETLBTEXT As Long = &H148
Const CB_GETLBTEXTLEN As Long = &H149
Const CB_GETCURSEL As Long = &H147
Const CB_GETITEMDATA As Long = &H150
Const CBS_OWNERDRAWFIXED As Long = &H10&
Const CBS_OWNERDRAWVARIABLE As Long = &H20&
Const LBS_MULTIPLESEL As Long = &H8&
Const LBS_EXTENDEDSEL As Long = &H800&
Const LB_GETSEL As Long = &H187
Const LB_GETCURSEL As Long = &H188
Const LB_GETCOUNT As Long = &H18B
Const LB_GETITEMDATA As Long = &H199
Const LBS_MULTICOLUMN As Long = &H200&

lStyle = GetWindowLong(hControl, GWL_STYLE)
' this version will not display ownerdrawn list/combo boxes, so we disregard them
If ControlType Then    ' test for multi-column listboxes -- these won't be handled
    If ((lStyle And LBS_MULTICOLUMN) = LBS_MULTICOLUMN) Then Exit Sub
    bMultiSelect = (((lStyle And LBS_MULTIPLESEL) = LBS_MULTIPLESEL) Or _
        ((lStyle And LBS_EXTENDEDSEL) = LBS_EXTENDEDSEL))
Else                   ' test for ownerdrawn comboboxes -- these won't be handled
    If ((lStyle And CBS_OWNERDRAWVARIABLE) = CBS_OWNERDRAWVARIABLE) Then Exit Sub
    If ((lStyle And CBS_OWNERDRAWFIXED) = CBS_OWNERDRAWFIXED) Then Exit Sub
End If

' if this menu item has a submenu already, we will destroy it.
' It is destroyed vs deleting individual items and resusing the menu
' This is done for speed. What if the list/combobox had 500 items?
NewSubmenuID = GetSubMenu(ParentMenu, mnuPos)  ' see if the current item is a submenu
On Error Resume Next
If NewSubmenuID Then   ' if so we get rid of it and any pre-existing child class, if needed
    gMenus.Remove "g" & NewSubmenuID
    DestroyMenu NewSubmenuID
    cItems.Remove NewSubmenuID & "." & ParentMenu
    Err.Clear
End If
' here we create a new submenu. The handle will be returned to calling routines
' in order to update this menu item as now having a submenu
NewSubmenuID = CreatePopupMenu()
If NewSubmenuID = 0 Then Exit Sub
    ' we'll store the listindex for these items should a sidebar caption with
    ' the NoScroll flag set be provided. Otherwise we could just use the
    ' zero-based menu item position as the listindex. But if the menu scrolls
    ' and sidebar is hidden, the zero-based indexes will be off by one
    sParent = "gControl:" & hControl & "|LIndex:x"
    ' if the listbox contains file names we need to process the Files flag
    ReturnComponentValue sCustomFlag, "Files:", sBuffer
    If Len(sBuffer) > 0 And ControlType <> 0 Then
        ' not applicable to combo boxes
        ' if -1, then the listbox contents includes the full path
        If sBuffer = "-1" Then
            sIcon = "Paths Included"
         Else   ' otherwise, we process the path
            If Right$(sBuffer, 1) <> "\" Then sBuffer = sBuffer & "\"
            ' here we check to ensure the path exists. If not, VB seems to
            ' crash with an invalid page fault error
            If Len(Dir(sBuffer, vbDirectory)) Then
                sParent = sParent & "|gEXE:" & sBuffer
                bShowHourglass = True
            Else    ' path doesn't exist, hope user provided it in the file names
                sIcon = "Paths Included"
            End If
        End If
    End If
    sParent = sParent & "|gType:" & ControlType + Abs(CInt(bMultiSelect))
    ' set the appropriate window messages for the type of control we have
    msgListCount = Choose(ControlType + 1, CB_GETCOUNT, LB_GETCOUNT)
    msgListIndex = Choose(ControlType + 1, CB_GETCURSEL, LB_GETCURSEL)
    msgTextLen = Choose(ControlType + 1, CB_GETLBTEXTLEN, LB_GETTEXTLEN)
    msgText = Choose(ControlType + 1, CB_GETLBTEXT, LB_GETTEXT)
    msgItemData = Choose(ControlType + 1, CB_GETITEMDATA, LB_GETITEMDATA)
    
    ' now we simply query the list/text box to get nr of items & current ListIndex, if any
    ListCount = SendMessage(hControl, msgListCount, 0&, ByVal 0&) - 1
    ListIndex = SendMessage(hControl, msgListIndex, 0, ByVal 0&)
    For Looper = 0 To ListCount
        ' loop thru list/combobox item to get that item's text
        lLen = SendMessage(hControl, msgTextLen, Looper, ByVal 0&)
        sBuffer = Space$(lLen) & Chr$(0)
        SendMessage hControl, msgText, Looper, ByVal sBuffer
        ' add the item to our new menu. We will use 32500 and below for our IDs
        ' Note: No need to worry about duplicate IDs since each custom menu will be processed
        ' in their own class and references are made by submenu ID AND menu ID
        lStyle = 0
        If bMultiSelect Then
            If SendMessage(hControl, LB_GETSEL, Looper, ByVal 0&) > 0 Then lStyle = MF_CHECKED
        Else
            If Looper = ListIndex Then lStyle = MF_CHECKED
        End If
        sBuffer = Trim(Replace$(sBuffer, Chr$(0), ""))
        If InStr(sBuffer, "{") Then
            ' carry over embedded flags within the list item
            sBuffer = Left$(sBuffer, InStr(sBuffer, "{")) & sParent & "|" & Mid(sBuffer, InStr(sBuffer, "{") + 1)
        Else
            'here we check to see if verifying the file exists is needed
            ' again only needed 'cause VB seems to crash if file doesn't exist
            If Len(sIcon) Then  ' yep, we validate it here
                ' extract the path and name from the list item
                sIcon = StripFile(GetShortFileName(sBuffer), "P")
                sBuffer = StripFile(sBuffer, "N")
                ' if the file doesn't exist, we insert a bogus flag
                If Len(Dir(sIcon & sBuffer, vbDirectory)) = 0 Then
                    sIcon = "|LaVolpe2"
                Else ' otherwise, we insert the flag to get the icon from the file
                    sIcon = "|gEXE:" & sIcon
                    bShowHourglass = True
                End If
            End If
            sBuffer = sBuffer & "{" & sParent & sIcon & "}"
        End If
        ' by putting the listindex in the caption, we won't worry about offsets should
        ' as sidebar be supplied with the NoScroll option & the menu scrolls,
        sBuffer = Replace$(sBuffer, "LIndex:x", "LIndex:" & Looper)
        AppendMenu NewSubmenuID, lStyle, 32500 - Looper, Replace$(sBuffer, "&", "&&")
    Next
    If ListCount = -1 Then  ' no list items to display, create a default
        lStyle = MF_DISABLED Or MF_GRAYED
        AppendMenu NewSubmenuID, lStyle, 32500, "-No Selections Available{" & sParent & "|Raised}"
    End If
    ' here's where we create a class to process these gMenus. They are not processed
    ' ahead of time to prevent delay in menu display. If they aren't shown now, no need
    ' to take the time to process them now?
    Dim myMenus As cMenuItems
    Set myMenus = New cMenuItems
    ' set the child class's parent form, imagelist and tip callback status same as this class
    If bShowHourglass Then lStyle = NewSubmenuID Else lStyle = 0
    ReturnComponentValue sCustomFlag, "Tip:", sBuffer
    myMenus.InitializeSubMenu FormHwnd, imgList, TipCallBackProc, True, lStyle, sBuffer
    myMenus.hPrevProc = PrevProc
    gMenus.Add myMenus, "g" & NewSubmenuID     ' now add the class to the menu & destroy previous instance
    Set myMenus = Nothing
End Sub

Private Sub CreateSubMenuCustom(MenuID As Long, mnuPos As Integer, ParentMenu As Long, _
    NewSubmenuID As Long, CustomTag As String, CurCustomVal As String, ParentTip As String)
' ========================================================================
' This is the function to create general use menus on the fly. Menus created
' here are only single select. If multi-selection is desired, you should add
' your menu items to a list box or combo box and link menu item to that control
' with the LB: or CB: flags
' ========================================================================

Dim lStyle As Long, lValue As Long, lOffset As Long
Dim Looper As Integer, CheckValue As Long, sCheckValue As String
Dim sText As String, sParent As String, CustomID As String
Dim bShowHourglass As Boolean

' if this menu item has a submenu already, we will destroy it
NewSubmenuID = GetSubMenu(ParentMenu, mnuPos)  ' see if the current item is a submenu
On Error Resume Next
If NewSubmenuID Then   ' if so we get rid of it and an pre-existing child class if needed
    gMenus.Remove "g" & NewSubmenuID
    DestroyMenu NewSubmenuID
    cItems.Remove NewSubmenuID & "." & ParentMenu
    Err.Clear
End If
' here we create a new submenu. The handle will be returned to calling
' routines in order to update this menu item as now having a submenu
NewSubmenuID = CreatePopupMenu()
If NewSubmenuID = 0 Then Exit Sub
' return the UserID value supplied after the CustomTag
ReturnComponentValue CurCustomVal & "|", "ID:", CustomID
lOffset = InStr(CustomID, ":")
If lOffset Then CustomID = Left$(CustomID, lOffset - 1)
If Len(CustomID) Then CurCustomVal = Replace$(CurCustomVal, ":ID:" & CustomID, "")
' we start building a custom tag to be displayed in the new menu items
sParent = "{gControl:0|lvID:" & CustomID
Select Case CustomTag
    Case "lvDrives:"    ' the drives custom menu
        Dim vDrives() As String, sTip As String
        ' strip out the drive to display as checked
        lOffset = InStr(CurCustomVal, ":")
        If lOffset > 1 Then sCheckValue = Left(CurCustomVal, lOffset) Else sCheckValue = CurCustomVal
        ListAllDrives vDrives   ' function to return all drives
        If UBound(vDrives) Then bShowHourglass = True
        ' we loop thru each item in the array and format the menu item captions
        For Looper = 1 To UBound(vDrives)
            ' the array is in the format of "drive|type of drive"
            CustomID = Left(vDrives(Looper), InStr(vDrives(Looper), "|") - 1)
            ' ensure we hvae a backslash after drive
            lOffset = InStr(CustomID, "\")
            If lOffset = 0 Then
                lOffset = InStr(CustomID, ":")
                If lOffset = 0 Then lOffset = Len(CustomID)
            End If
            sText = Left(CustomID, lOffset)
            If Right$(sText, 1) <> "\" Then sText = sText & "\"
            ' now we format the drive letter for display like "[C:]"
            If InStr(sText, ":") Then sText = "[" & UCase(Left$(sText, InStr(sText, ":") - 1)) & "]"
            sText = sText & Mid$(CustomID, lOffset + 1)
            If Left$(sText, Len(sCheckValue)) = sCheckValue Then lStyle = MF_CHECKED Else lStyle = 0
            ' final formatting with undocumented tags
            sText = sText & sParent & "|LDrv:" & CustomID _
                          & "|Tip:" & Mid$(vDrives(Looper), InStr(vDrives(Looper), "|") + 1) & "}"
            AppendMenu NewSubmenuID, lStyle, 32500 - Looper, sText
        Next
    Case "lvFonts:"
        Dim vFonts() As String, sLtrA As String, sLtrZ As String, bInclude As Boolean
        ReturnComponentValue CurCustomVal & "|", "Type:", sText
        If Len(sText) Then lOffset = InStr(sText, ":")
        If lOffset Then sText = Left$(sText, lOffset - 1)
        Select Case sText
        Case "TrueType": lStyle = TRUETYPE_FONTTYPE
        Case "System": lStyle = RASTER_FONTTYPE ' System fonts
        Case Else: lStyle = 0                   ' all fonts
        End Select
        lOffset = InStr(CurCustomVal, ":")
        If lOffset Then sCheckValue = Left(CurCustomVal, lOffset - 1)
        ReturnComponentValue CurCustomVal & "|", "Group:", CustomTag
        lOffset = InStr(CustomTag, ":")
        If lOffset Then CustomTag = Left(CustomTag, lOffset - 1)
        If Len(CustomTag) = 3 Then
            sLtrA = Left(CustomTag, 1)
            sLtrZ = Right(CustomTag, 1)
        End If
        LoadFontMenu vFonts(), lStyle
        lOffset = 1
        For Looper = 1 To UBound(vFonts)
            sText = StrConv(vFonts(Looper), vbProperCase)
            bInclude = False
            If Len(CustomTag) Then
                If Left$(vFonts(Looper), 1) >= sLtrA And Left$(vFonts(Looper), 1) <= sLtrZ Then bInclude = True
            Else
                bInclude = True
            End If
            If bInclude = True Then
                sText = sText & sParent & "|LFont:" & vFonts(Looper) & "}" & vbTab & "Sample"
                If sCheckValue = vFonts(Looper) Then lStyle = MF_CHECKED Else lStyle = 0
                AppendMenu NewSubmenuID, lStyle, 32500 - lOffset, sText
                lOffset = lOffset + 1
            End If
        Next
        If UBound(vFonts) = 0 Or lOffset = 1 Then AppendMenu NewSubmenuID, MF_DISABLED, 32500, "-No Fonts Available{Raised}"
        If lOffset > 20 Then bShowHourglass = True
        Erase vFonts
    Case "lvColors:"    ' the 23 colors menu
        ' we set a value equal to the color provided in the caption to be checked
        If Len(CurCustomVal) > 0 And Val(CurCustomVal) <> -1 Then
            CheckValue = ExchangeVBcolor(CurCustomVal, -1)
        Else
            CheckValue = -1
        End If
        ' loop thru each color and add it to the menu
        For Looper = 1 To 23
            sText = Choose(Looper, "Black", "Blue", "Brown", "Cyan / Aqua", "Gold", "Gray", "Green / Lime", "Green Dark / Olive", "Ivory", "Magenta / Fuchsia", "Maroon", "Navy", "Orange", "Peach", "Purple", "Red", "Silver / Light Gray", "Tan", "Teal", "Turquoise", "Violet", "White", "Yellow")
            lValue = Choose(Looper, vbBlack, vbBlue, vbBrown, vbCyan, vbGold, vbGray, vbGreen, vbOlive, vbIvory, vbMagenta, vbMaroon, vbNavy, vbOrange, vbPeach, vbPurple, vbRed, vbSilver, vbTan, vbTeal, vbTurquoise, vbViolet, vbWhite, vbYellow)
            ' add this undocumented tag to the menu caption
            sText = sText & sParent & "|LColor:" & lValue & "}"
            If lValue = CheckValue Then lStyle = MF_CHECKED Else lStyle = 0
            AppendMenu NewSubmenuID, lStyle, 32500 - Looper, sText
        Next
        ' now we will add a generic select another color. The user will be
        ' responsible for providing a color popup dialog if user selects this
        AppendMenu NewSubmenuID, MF_DISABLED, 32500 - Looper, "-{Raised}"
        ' the value in this case would be returned as -1
        sText = "Select a custom color" & sParent & "|LDefClr:" & CheckValue & "|LColor:-1}"
        AppendMenu NewSubmenuID, 0, 32500 - Looper - 1, sText
    Case "lvMonths:"
        ' the 3 types of display for months. Look for the checked value
        CheckValue = Val(CurCustomVal)
        If CheckValue = 0 Then CheckValue = Month(Date)
        ReturnComponentValue CurCustomVal & "|", "Group:", CustomTag
        lOffset = InStr(CustomTag, ":")
        If lOffset Then CustomTag = Left$(CustomTag, lOffset - 1)
        ' add the grouping separator bar
        Select Case CustomTag
        Case "CYQtr"
            AppendMenu NewSubmenuID, MF_DISABLED, 32500, "-First Quarter{Raised}"
        Case "FYQtr":
            AppendMenu NewSubmenuID, MF_DISABLED, 32500, "-1st Fiscal Quarter{Raised}"
        End Select
        lValue = 1  ' we need to keep counter of menu items added
        For Looper = 1 To 12
            ' for fiscal years, the year starts in October
            If CustomTag = "FYQtr" Then
                lOffset = Choose(Looper, 10, 11, 12, 1, 2, 3, 4, 5, 6, 7, 8, 9)
            Else    ' for calendar years the year starts in January
                lOffset = Looper
            End If
            If (Looper - 1) Mod 3 = 0 And Looper > 1 Then
                ' every third month, we add a grouping separator bar
                Select Case CustomTag
                Case "CYQtr"
                    sText = Choose((Looper - 1) / 3, "Second", "Third", "Fourth") & " Quarter{Raised}"
                    AppendMenu NewSubmenuID, MF_DISABLED, 32500 - lValue, "-" & sText
                    lValue = lValue + 1
                Case "FYQtr"
                    sText = Choose((Looper - 1) / 3, "2nd", "3rd", "4th") & " Fiscal Quarter{Raised}"
                    AppendMenu NewSubmenuID, MF_DISABLED, 32500 - lValue, "-" & sText
                    lValue = lValue + 1
                End Select
            End If
            If lOffset = CheckValue Then lStyle = MF_CHECKED Else lStyle = 0
            ' add this undocumented tag to the caption
            sText = Format(DateSerial(1962, lOffset, 7), "mmmm") & sParent _
                    & "|LMonth:" & lOffset & "}"
            AppendMenu NewSubmenuID, lStyle, 32500 - lValue, sText
            lValue = lValue + 1
        Next
    Case "lvDays:"  ' week days
        CheckValue = Val(CurCustomVal)
        If CheckValue = 0 Then CheckValue = Weekday(Date)
        ' here we are just trying to find the first day that falls on the
        ' system setting for the beginning of the week. We do this to keep
        ' the program pretty much international aware vs hardcoding the
        ' weekdays in a specific language
        lOffset = Weekday(Date) - 1
        If lOffset Then lOffset = 7 - lOffset
        ' loop thru each day of the week
        For Looper = 1 To 7
            sText = Format(DateAdd("d", lOffset, Date), "dddd")
            If Looper = CheckValue Then lStyle = MF_CHECKED Else lStyle = 0
            ' add this undocumented tag
            sText = sText & sParent & "|LDay:" & Looper & "}"
            lOffset = lOffset + 1   ' used to add another day to start date
            AppendMenu NewSubmenuID, lStyle, 32500 - Looper, sText
        Next
    Case "lvStates:"    ' the 50 U.S. states & DC
        Dim sAbbr As String
        If Len(CurCustomVal) > 2 Then
            sCheckValue = Left$(CurCustomVal, 2)
        Else
            If Len(CurCustomVal) < 2 Then sCheckValue = "LaVolpe" Else sCheckValue = CurCustomVal
        End If
        bShowHourglass = True
        For Looper = 1 To 51
            ' loop thru each of the 51 items. Following broken down for readibility
            Select Case Looper
            Case Is < 11
                sText = Choose(Looper, "Alabama", "Alaska", "Arizona", "Arkansas", "California", "Colorado", "Connecticut", "Delaware", "District of Columbia", "Florida")
                sAbbr = Choose(Looper, "AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "DC", "FL")
            Case Is < 21
                sText = Choose(Looper - 10, "Georgia", "Hawaii", "Idaho", "Illinois", "Indiana", "Iowa", "Kansas", "Kentucky", "Louisiana", "Maine")
                sAbbr = Choose(Looper - 10, "GA", "HI", "ID", "IL", "IN", "IA", "KS", "KY", "LA", "ME")
            Case Is < 31
                sText = Choose(Looper - 20, "Maryland", "Massachusettes", "Michigan", "Minnesota", "Mississippi", "Missouri", "Montana", "Nebraska", "Nevada", "New Hampshire")
                sAbbr = Choose(Looper - 20, "MD", "MA", "MI", "MN", "MS", "MO", "MT", "NE", "NV", "NH")
            Case Is < 41
                sText = Choose(Looper - 30, "New Jersey", "New Mexico", "New York", "North Carolina", "North Dakota", "Ohio", "Oklahoma", "Oregon", "Pennsylvania", "Rhode Island")
                sAbbr = Choose(Looper - 30, "NJ", "NM", "NY", "NC", "ND", "OH", "OK", "OR", "PA", "RI")
            Case Is < 51
                sText = Choose(Looper - 40, "South Carolina", "South Dakota", "Tennessee", "Texas", "Utah", "Vermont", "Virginia", "Washington", "West Virginia", "Wisconsin")
                sAbbr = Choose(Looper - 40, "SC", "SD", "TN", "TX", "UT", "VT", "VA", "WA", "WV", "WI")
            Case Is < 61
                sText = Choose(Looper - 50, "Wyoming")
                sAbbr = Choose(Looper - 50, "WY")
            End Select
            ' add the following undocumented tag
            sText = sText & sParent & "|LState:" & sAbbr & "}"
            If sAbbr = sCheckValue Then lStyle = MF_CHECKED Else lStyle = 0
            AppendMenu NewSubmenuID, lStyle, 32500 - Looper, sText
        Next
    Case "lvMonth:" ' the days of the month
        Dim lvYear As Long, lDate As Long
        lValue = Val(CurCustomVal)  ' get the month to display
        If lValue = 0 Then lValue = Month(Date)
        If lValue > 0 And lValue < 13 Then
            ' now we get the year to display with current year as default
            ReturnComponentValue CurCustomVal & "|", "Year:", sText
            If Len(sText) Then lvYear = Val(sText) Else lvYear = Year(Date)
            If lvYear < 1 Then lvYear = Year(Date)
            ' now we get the day of the month to check, if any
            ReturnComponentValue CurCustomVal & "|", "Day:", sText
            CheckValue = Val(sText)
            If CheckValue = 0 Then CheckValue = Day(Date)
            ' loop thru each day of the month and add it to the menu
            For Looper = 1 To Day(DateSerial(lvYear, lValue + 1, 0))
                ' build a date to be used to format the menu caption
                lDate = CDate(DateSerial(lvYear, lValue, Looper))
                sText = Format(lDate, "mmm d, dddd")
                ' add undocumented tag
                sText = sText & sParent & "|LDate:" & lDate & "}"
                If Looper = CheckValue Then lStyle = MF_CHECKED Else lStyle = 0
                AppendMenu NewSubmenuID, lStyle, 32500 - Looper, sText
            Next
        End If
    Case Else
End Select
' here's where we create a class to process these gMenus. They are not processed
' ahead of time to prevent delay in menus. If they aren't shown, no need to take the
' time to process them now?
Dim myMenus As cMenuItems
Set myMenus = New cMenuItems
If bShowHourglass Then lValue = NewSubmenuID Else lValue = 0
' set the child class's parent form, imagelist and tip callback status same as this class
myMenus.InitializeSubMenu FormHwnd, imgList, TipCallBackProc, True, lValue, ParentTip
gMenus.Add myMenus, "g" & NewSubmenuID     ' now add the class to the menu & destroy previous instance
Set myMenus = Nothing
End Sub

Private Function FormIndex() As Integer
Dim I As Integer
' ========================================================================
' function returns the Forms(x) index of this class's parent form
' can't store this value 'cause index can change when lower indexed forms are closed
' ========================================================================
' locating the form based off of the form's hWnd
' simply loop thru (top to bottom) to find the right form
For I = Forms.Count - 1 To 0 Step -1
    If Forms(I).hWnd = FormHwnd Then
        FormIndex = I
        Exit For
    End If
Next
End Function

Private Function SplitControlIndex(sControlName As String) As Control
' ========================================================================
' Function converts a string ref to a control to an actual control/object for referencing
' ========================================================================
If Len(sControlName) = 0 Then Exit Function
Dim Index As Long, sName As String
' test to see if control is indexed, if so it will always end in a parenthesis
sControlName = Trim$(sControlName)
If Right$(sControlName, 1) = ")" Then
    ' extract the index
    Index = Val(Mid$(sControlName, InStrRev(sControlName, "(") + 1))
    ' extract the name
    sName = Left$(sControlName, InStrRev(sControlName, "(") - 1)
    ' locate the control
    Set SplitControlIndex = Forms(FormIndex).Controls(sName)(Index)
Else
    ' no index, so locate the control based on the name
    Set SplitControlIndex = Forms(FormIndex).Controls(sControlName)
End If
End Function

Private Sub Class_Terminate()
' ========================================================================
' clean up memory resources
' close all child classes, erase all arrays & reset all collections
' ========================================================================
Dim mIndex As Long
On Error Resume Next
For mIndex = 1 To UBound(vPanels)
    If vPanels(mIndex).PanelIcon Then
        Debug.Print "deleting panel image "; vPanels(mIndex).PanelIcon
        DeleteObject vPanels(mIndex).PanelIcon
    End If
Next
Erase vItems
Erase vPanels
Set imgList = Nothing
Set gMenus = Nothing
Set cItems = Nothing
Set cPanels = Nothing
End Sub

Private Sub Class_Initialize()
' ========================================================================
Set gMenus = New Collection   ' collection of child classes used for combo/listbox menus
Set cItems = New Collection     ' index to MenuComponentData structure array
Set cPanels = New Collection   ' index to PanelData structure array
ReDim vItems(0 To 0)               ' array of MenuComponentData structures
ReDim vPanels(0 To 0)             ' array of PanelData structures
uMinMax.ptMaxSize.X = -1
End Sub

Public Sub RestoreMenus()
' ========================================================================
' This routine is used by the optional UserControl provided with the project
' or when the menu font name or font size has been changed by the user
' Routine will revert menus back to their original captions/states and as non-ownerdrawn
' ========================================================================
Dim mIndex As Integer, Looper As Integer
Dim hMenu As Long, typeRemoval As Long
Dim nrItems As Long, mItem As Long

Dim MII As MENUITEMINFO, MIS As MENUITEMINFO_string
Dim MI() As Byte, mRestore() As Long, mSubMenus() As Long
On Error Resume Next
' need to loop thru each submenu we have processed
For Looper = 1 To UBound(vPanels)
    ' each menu item is stored by menu item ID and submenu it belongs too
    ' this prevents overwritting menu items that may have the same ID but belong to different submenus
    hMenu = vPanels(Looper).SubmenuID
    If IsMenu(hMenu) Then
        ReDim mRestore(0)                           ' array to replace menu items converted to submenus (gMenus)
        ReDim mSubMenus(0)
        nrItems = GetMenuItemCount(hMenu)   ' number of menu items in the submenu
        For mIndex = 0 To nrItems - 1
            ' by getting the menu items from the submenu directly, we avoid making items where the Visibile
            ' property is set to false from reappearing.  If we went from the stored array, those non-visible items
            ' are in our array and we would insert them back into the menu thereby making them visible
            ReDim MI(0 To 1023)
            MII.cbSize = Len(MII)
            MII.fMask = MIIM_TYPE Or MIIM_ID Or MIIM_SUBMENU
            MII.fType = 0       ' equivalent to MFT_STRING
            MII.dwTypeData = VarPtr(MI(0))
            MII.cch = UBound(MI)
            ' retrieve the menu items current settings
            GetMenuItemInfo hMenu, mIndex, True, MII
            If ((MII.fType And MF_OWNERDRAW) = MF_OWNERDRAW) Then
                ' build a reference to the item stored in our local array
                mItem = cItems(MII.wID & "." & hMenu)
                ReDim Preserve mRestore(0 To UBound(mRestore) + 1)
                mRestore(UBound(mRestore)) = mItem
                ReDim Preserve mSubMenus(0 To UBound(mSubMenus) + 1)
                mSubMenus(UBound(mSubMenus)) = MII.hSubMenu
            End If
        Next
        ' now for those gMenu items and custom menu items
        For mIndex = UBound(mRestore) To 1 Step -1
            ' we'll delete them first
            RemoveMenu hMenu, vItems(mRestore(mIndex)).ID, 0
        Next
        ' and now we re-insert them using the original menu caption and zero-based index on the submenu
        For mIndex = 1 To UBound(mRestore)
            MIS.fMask = MIIM_TYPE Or MIIM_STATE Or MIIM_ID Or MIIM_SUBMENU
            MIS.fState = 0
            MIS.fType = 0
            With vItems(mRestore(mIndex))
                If Len(.Cached) Then
                    MIS.dwTypeData = .Cached
                Else
                    MIS.dwTypeData = .Caption
                End If
                If ((.Status And lv_mSep) = lv_mSep) Then
                    If MIS.dwTypeData = "" Then
                        MIS.dwItemData = "-"
                    Else
                        .Caption = "LaVolpe" & Chr$(1) & Chr$(2) & Chr$(3) & mIndex
                    End If
                Else
                    If ((.Status And lv_mChk) = lv_mChk) Then MIS.fState = MIS.fState Or MF_CHECKED
                    .Caption = "LaVolpe" & Chr$(1) & Chr$(2) & Chr$(3) & mIndex
                End If
                MIS.cch = Len(MIS.dwTypeData)
                ' ensure we carry over the checked & enabled status
                If ((.Status And lv_mDisabled) = lv_mDisabled) Then
                    MIS.fState = MIS.fState Or MF_DISABLED
                    MIS.fState = MIS.fState Or MF_GRAYED
                End If
                MIS.hSubMenu = mSubMenus(mIndex)
                MIS.cbSize = Len(MIS)
                MIS.wID = .ID
                'Debug.Print "restoring "; MIS.dwTypeData; MIS.wID
                InsertMenuItem hMenu, .Index, True, MIS
            End With
        Next
    End If
Next
Erase mRestore
Erase mSubMenus
End Sub

Public Sub DestroyPopup(hMenu As Long, hMenuID As Long)
' ========================================================================
' This is called immediately after a call to to PopupMenuCustom and the
' user closed the popup menu. Since the menu displayed will always consist
' of 2 new submenus and numerous menu items, we want to delete them vs
' storing them indefinitely. It's truly easy since the parent to the menu
' that was displayed is always the last item in our local array. Likewise,
' the panel array's last entry will also be deleted. Well, that only leaves
' the gMenu that was created & guess what: it's the last gMenu in our collection
' ========================================================================
On Error Resume Next
Dim mIndex As Long
mIndex = cItems(hMenuID & "." & hMenu)
ReDim Preserve vItems(0 To UBound(vItems) - 1)
mIndex = cPanels("p" & hMenu)
' this line can error if a user passed bogus information to the PopupMenuCustom function
gMenus.Remove gMenus.Count
End Sub

Private Sub CreateImageSidebar(mIndex As Long, pIndex As Long, wCaption As String)
' ========================================================================
' This function will create a memory bitmap of the sidebar to use
' We want memory copy to speed up repainting menus as needed
' NOTE: this copy is not destroyed until the form is closed or it is
'       replaced by an updated bitmap if menu height/width changes
' ========================================================================

'Debug.Print "Processing sidebar"

Dim bmpInfo As BITMAP, icoInfo As ICONINFO, hRect As RECT
Dim imgSize As POINTAPI, imgAdj As POINTAPI, xyOffset As POINTAPI
Dim tDC As Long, fDC As Long, wDC As Long
Dim ImgType As Integer, lImgCopy As Long
Dim BColor As Long, GColor As Long

Dim sValue As String, iAlign As Integer
Dim imgRatioX As Single, imgRatioY As Single
Dim oldBMP(0 To 1) As Long, newBMP As Long, srcBMP As Long

On Error Resume Next
srcBMP = ValidateImage(wCaption, ImgType, imgSize.X, imgSize.Y)
If srcBMP = 0 Then
        Debug.Print "Sidebar failed to find image as "; sValue
        GoTo CreateGenericSidebar
        Exit Sub
End If
' temporary DC to draw in
fDC = GetDC(FormHwnd)
tDC = CreateCompatibleDC(fDC)   ' DC to hold final copy of bitmap
wDC = CreateCompatibleDC(fDC)   ' DC to work in
With vItems(mIndex)
    ' get the width of the sidebar & use 35 as a default if necessary
    ReturnComponentValue wCaption, "Width:", sValue
    If Val(sValue) < 35 Then sValue = 35
    .Dimension.X = Val(sValue)
    ' get the alignment of the sidebar image
    ReturnComponentValue wCaption, "ALIGN:", sValue
    Select Case sValue
    Case "TOP": iAlign = 1
    Case "BOT": iAlign = 2
    Case Else: iAlign = 0
    End Select
    If ImgType = 1 Then ' bitmap image
        lImgCopy = CopyImage(srcBMP, 0, 0, 0, 0)
        oldBMP(1) = SelectObject(wDC, lImgCopy)
    End If
    ' when image is too large for the sidebar panel, we need to resize it
    If imgSize.X > .Dimension.X Or imgSize.Y > .Dimension.Y Then
        imgRatioX = (.Dimension.X / imgSize.X)
        imgRatioY = (.Dimension.Y / imgSize.Y)
        If imgRatioY < imgRatioX Then imgRatioX = imgRatioY
        imgAdj.X = imgRatioX * imgSize.X
        imgAdj.Y = imgRatioX * imgSize.Y
    Else    ' will fit within sidebar boundary
        imgAdj.X = imgSize.X
        imgAdj.Y = imgSize.Y
    End If
    ' depending on alignment, calculate where in the panel image will appear
    Select Case iAlign
    Case 1: xyOffset.Y = 0                              ' top of panel
    Case 2: xyOffset.Y = .Dimension.Y - imgAdj.Y        ' bottom of panel
    Case Else: xyOffset.Y = (.Dimension.Y - imgAdj.Y) \ 2   ' centered
    End Select
    xyOffset.X = (.Dimension.X - imgAdj.X) \ 2
    ' create the final sidebar image
    newBMP = CreateCompatibleBitmap(fDC, .Dimension.X, .Dimension.Y)
    oldBMP(0) = SelectObject(tDC, newBMP)
    ' figure out which background color to use & fill the sidebar with it
    ' get the back color for the sidebar image
    ReturnComponentValue wCaption, "BColor:", sValue
    BColor = Val(sValue)
    If Val(sValue) = vbNull Then
        BColor = GetSysColor(COLOR_MENU)
    Else
        If BColor <> -1 Then BColor = ExchangeVBcolor(sValue, -1)
    End If
    ' see if a gradient will be used
    ReturnComponentValue wCaption, "GColor:", sValue
    GColor = ExchangeVBcolor(sValue, vbNull)
    If GColor = vbNull Or GColor = BColor Or BColor < 0 Then
        If BColor = -1 Then BColor = GetPixel(wDC, 0, 0)
        DrawRect tDC, 0, 0, .Dimension.X, .Dimension.Y, BColor
    Else
        hRect.Right = .Dimension.X
        hRect.Bottom = .Dimension.Y
        DrawGradient BColor, GColor, False, tDC, hRect
    End If
    ' now put the image on the sidebar
    If ImgType Then
        ReturnComponentValue wCaption, "Transparent", sValue
        If Len(sValue) = 0 Then
            StretchBlt tDC, xyOffset.X, xyOffset.Y, _
                imgAdj.X, imgAdj.Y, _
                wDC, 0, 0, imgSize.X, imgSize.Y, vbSrcCopy
            SelectObject wDC, oldBMP(1)
        Else
            Dim bmpRect As RECT
            SelectObject wDC, oldBMP(1)
            hRect.Left = xyOffset.X
            hRect.Right = imgAdj.X + hRect.Left
            hRect.Top = xyOffset.Y
            hRect.Bottom = imgAdj.Y + hRect.Top
            DrawTransparentBitmap tDC, hRect, lImgCopy, bmpRect, , imgAdj.X, imgAdj.Y
        End If
        DeleteObject lImgCopy
    Else    ' icons vs bitmaps
        DrawIconEx tDC, xyOffset.X, xyOffset.Y, srcBMP, imgAdj.X, imgAdj.Y, 0, 0, &H3
    End If
    ' clean up
    SelectObject tDC, oldBMP(0)
    If vPanels(pIndex).PanelIcon Then DeleteObject vPanels(pIndex).PanelIcon
    vPanels(pIndex).PanelIcon = newBMP  ' cache the new bitmap
    Debug.Print "created new sidebar image "; newBMP, wCaption  ' " cx="; .Dimension.X; .Dimension.Y
End With
Cleanup:
' remove temporary DCs
DeleteDC tDC
DeleteDC wDC
ReleaseDC FormHwnd, fDC
Erase oldBMP
Exit Sub

CreateGenericSidebar:
' here we create a generic sidebar if the image sidebar the user passed cannot
' be made. Reasons: generally referencing bad image handle or control name where
' control isn't on the form hosting this menu item
CreateTextSidebar mIndex, pIndex, "{Sidebar|Text:La<br0>Volpe Menus|FColor:8388608|BColor:12615680|GColor:16776960|Font:Times New Roman|FSize:14|MinFSize:9|Width:32|Align:Bot}"
End Sub

Private Sub CreateTextSidebar(mIndex As Long, pIndex As Long, wCaption As String)
' ========================================================================
' This function will create a memory bitmap of the sidebar to use
' We want memory copy to speed up repainting menus as needed
' NOTE: this copy is not destroyed until the form is closed or it is
'       replaced by an updated bitmap if menu height/width changes
' TODO: rework the logic to wordwrap/shrink text
' ========================================================================
Dim sMultiLine As String, sWords() As String, sCaptionSeg As String
Dim iFontSize As Single, MinFontSize As Single, FontPrefer As Single
Dim I As Integer, iLineHT As Integer
Dim lFontM As LOGFONT, tRect As RECT, hRect As RECT
Dim hPrevFont As Long, lFont As Long, maxHT As Long
Dim FColor As Long, BColor As Long, GColor As Long
Dim sValue As String, sideBarX As Long
Dim tDC As Long, fDC As Long, oldBMP As Long, newBMP As Long

fDC = GetDC(FormHwnd)
tDC = CreateCompatibleDC(fDC)
ReturnComponentValue wCaption, "Text:", sMultiLine
If Len(sMultiLine) = 0 Then sMultiLine = "La Volpe<br>Menus"
' make sure accidental words in text don't affect the sidebar
' Example if word "bold" in text, without removing the text/caption,
' the word "bold" could be interpreted as making the font bold
wCaption = Replace$(wCaption, sMultiLine, "")
' now we look for flags within the text/caption
' let's set the font attributes now
    ReturnComponentValue wCaption, "FONT:", sValue
    If Len(sValue) = 0 Then sValue = "Arial"
    lFontM.lfCharSet = 1
    lFontM.lfFaceName = sValue & Chr$(0)
    ReturnComponentValue wCaption, "BOLD", sValue
    If Len(sValue) Then lFontM.lfWeight = 800 Else lFontM.lfWeight = 400
    ReturnComponentValue wCaption, "ITALIC", sValue
    If Len(sValue) Then lFontM.lfItalic = 1  ' sTmp = sTmp & " Italic"
    ReturnComponentValue wCaption, "UNDERLINE", sValue
    If Len(sValue) Then lFontM.lfUnderline = 1
    ReturnComponentValue wCaption, "FSize:", sValue
    If Len(sValue) Then
        iFontSize = Val(sValue)
        If iFontSize < 7 Then iFontSize = 7
        If iFontSize > 24 Then iFontSize = 24
    Else
        iFontSize = 11
    End If
    FontPrefer = iFontSize
    ReturnComponentValue wCaption, "MINFSIZE:", sValue
    If Len(sValue) Then MinFontSize = Val(sValue)
    If MinFontSize < 5 Then MinFontSize = 5
    If MinFontSize > iFontSize Then MinFontSize = iFontSize
' get the width of the sidebar & use 35 as a default if necessary
ReturnComponentValue wCaption, "Width:", sValue
If Val(sValue) < 35 Then sValue = 35
sideBarX = Val(sValue)              ' actual value for calculations
maxHT = vItems(mIndex).Dimension.Y
' get the back/fore color for the sidebar image
ReturnComponentValue wCaption, "FColor:", sValue
FColor = ExchangeVBcolor(sValue, 0)
ReturnComponentValue wCaption, "BColor:", sValue
If Val(sValue) = vbNull Then
    BColor = GetSysColor(COLOR_MENU)
Else
    BColor = ExchangeVBcolor(sValue, GetSysColor(COLOR_MENU))
End If
' see if a gradient will be used
ReturnComponentValue wCaption, "GColor:", sValue
GColor = ExchangeVBcolor(sValue, vbNull)
' routine for line breaking and or font reduction to make caption fit
ReDim sWords(0)
sWords(0) = sMultiLine
GoSub SetTestFont
sMultiLine = LineBreaks(sWords(), True, tDC, maxHT)
Do
    DrawText tDC, sMultiLine, Len(sMultiLine), tRect, DT_CALCRECT Or DT_LEFT Or DT_HIDEPREFIX Or DT_MULTILINE Or DT_NOPREFIX
    Do While tRect.Right + 6 > maxHT Or tRect.Bottom > sideBarX
        If iFontSize > MinFontSize + 1 Then
            iFontSize = iFontSize - 1
            GoSub SetTestFont
            DrawText tDC, sMultiLine, Len(sMultiLine), tRect, DT_CALCRECT Or DT_LEFT Or DT_HIDEPREFIX Or DT_MULTILINE Or DT_NOPREFIX
        Else
            Exit Do
        End If
    Loop
    ' the width may not support the current caption (if line breaks needed)
    ' but the height supports the caption. The minimum authorized font was
    ' reached so no point in breaking up the line further. Just exit the
    ' loop and adjust the width of the sidebar to display the caption
    If tRect.Right + 6 <= maxHT Then Exit Do
    If tRect.Bottom < sideBarX Then
        ' the test font doesn't use the entire width of the sidebar and also
        ' it won't fit as is. So we have room to play & try again
        iFontSize = FontPrefer
        GoSub SetTestFont
    End If
    'if following returns a nullstring then the line cannot be broken up any
    ' further; the font can't be reduced any further and the caption still
    ' won't fit on the sidebar. Oh well, tried everything, short of completely
    ' reconfiguring the sidebar to make it work. Not my problem! If the user is
    ' providing long words on a short menu and didn't authorize a small enough
    ' font size to use, so sorry.
    sMultiLine = LineBreaks(sWords(), False, tDC, maxHT)
    If sMultiLine = "" Then Exit Do
Loop
lFontM.lfEscapement = 900
lFontM.lfOrientation = 900
GoSub SetTestFont   ' create the vertical font
' now we create a bitmap to draw the text on
If tRect.Bottom + 6 > sideBarX Then sideBarX = tRect.Bottom + 6
newBMP = CreateCompatibleBitmap(fDC, sideBarX, maxHT)
oldBMP = SelectObject(tDC, newBMP)
vItems(mIndex).Dimension.X = sideBarX
' let's set the back color of the bitmap by drawing a rectangle on it
If GColor = vbNull Or GColor = BColor Then
    DrawRect tDC, 0, 0, sideBarX, maxHT, BColor
Else    ' gradient to be used
    hRect.Right = sideBarX
    hRect.Bottom = maxHT
    DrawGradient BColor, GColor, False, tDC, hRect
End If
' here we reprocess the caption and draw it on the bitmap
SetBkMode tDC, NEWTRANSPARENT
SetTextColor tDC, FColor    ' create the forecolor
' here we get the height of one line of text & will use it to space each line
DrawText tDC, "X", 1, hRect, DT_CALCRECT Or DT_LEFT Or DT_NOCLIP Or DT_SINGLELINE
' this calc seems to always give a good result
iLineHT = hRect.Right + (hRect.Right \ 2)
' set up the starting point for the 1st line of text
tRect.Left = (sideBarX - tRect.Bottom) \ 2
For I = 0 To UBound(sWords)
    ' loop thru each line of text
    sCaptionSeg = Trim$(Replace$(sWords(I), "<br>", " "))
    sCaptionSeg = Replace$(sCaptionSeg, "<br0>", " ")
    hRect.Left = 0: hRect.Top = 0   ' reset rectangles & measure line height
    DrawText tDC, sCaptionSeg, Len(sCaptionSeg), hRect, DT_CALCRECT Or DT_LEFT Or DT_NOCLIP Or DT_SINGLELINE Or DT_HIDEPREFIX
    ' now determine where the bottom edge of the drawing rectangle should be
    ReturnComponentValue wCaption, "ALIGN:", sValue
    Select Case sValue
    Case "Bot"   ' left/bottom
        tRect.Top = maxHT - 3
    Case "Top"  ' right/top
        tRect.Top = maxHT - (maxHT - hRect.Right) + 3
    Case Else  ' center
        tRect.Top = (maxHT - hRect.Right) \ 2 + hRect.Right
    End Select
    ' draw the text & set a new starting left edge
    DrawText tDC, sCaptionSeg, Len(sCaptionSeg), tRect, DT_HIDEPREFIX Or DT_SINGLELINE Or DT_NOCLIP Or DT_NOPREFIX
    tRect.Left = tRect.Left + iLineHT
Next
' Clean up
SelectObject tDC, hPrevFont     ' delete font used
DeleteObject lFont
SelectObject tDC, oldBMP        ' delete existing panel image if any
If vPanels(pIndex).PanelIcon Then DeleteObject vPanels(pIndex).PanelIcon
vPanels(pIndex).PanelIcon = newBMP
DeleteDC tDC                    ' clean up DCs
ReleaseDC FormHwnd, fDC
Erase sWords
Exit Sub

SetTestFont:
If hPrevFont Then
    SelectObject tDC, hPrevFont
    DeleteObject lFont
End If
lFontM.lfHeight = (iFontSize * -20) / Screen.TwipsPerPixelY
lFont = CreateFontIndirect(lFontM)
hPrevFont = SelectObject(tDC, lFont)
Return
End Sub

Private Function LineBreaks(Words() As String, bInitial As Boolean, tDC As Long, MaxWidth As Long) As String
' ========================================================================
' This function is a pretty good routine to break words into multiple lines
' ========================================================================

Dim I As Integer, J As Integer, K As Integer, lb As Integer, sRtn As String, tRect As RECT
Dim sSegment As String

' The coded sidebar text may have the <br!> flag & if so, this is a forced
' line break, so we process these first & will only be found the first
' time thru the routine (bInitial parameter = True)
I = InStr(Words(0), "<br!>")
Do While I
    ' simply increment the word/line count for each forced break found
    ReDim Preserve Words(UBound(Words) + 1)
    Words(UBound(Words)) = Mid$(Words(UBound(Words) - 1), I + 5)
    Words(UBound(Words) - 1) = Left$(Words(UBound(Words) - 1), I - 1)
    I = InStr(Words(UBound(Words)), "<br!>")
Loop
If Not bInitial Then
    ' on subsequent passes, the sidebar text cannot be fit in the sidebar
    ' even though the minimum font size has been reached. This is where we
    ' break the sidebar text into lines using the current font size
    For J = 1 To 2
        For I = 0 To UBound(Words)
            ' we check each line for a preferred break first: <br>
            ' and if needed loop back thru looking for spaces
            ' but spaces is where this routine does its magic!
            lb = InStr(Words(I), Choose(J, "<br>", " "))
            If lb Then
                If J = 2 Then   ' no <br> flags, so we are on spaces as line breaks
                    ' first we draw the line using the current font being tested
                    ' in the CreateTextSidebar routine.
                    DrawText tDC, Words(I), Len(Words(I)), tRect, DT_CALCRECT Or DT_LEFT Or DT_SINGLELINE Or DT_NOCLIP Or DT_NOPREFIX
                    If tRect.Right + 6 <= MaxWidth Then
                        ' if the line will fit on the sidebar, then returning a null string as
                        ' the function value will force the sidebar to widen itself to display
                        ' the text since it will fit vertically
                        lb = 0
                    Else
                        ' otherwise, we just don't want to break a line on the first space we see
                        ' If we do that each word can be a separate line & that looks crappy!
                        ' what we do is start from the middle of the line and find the first space
                        ' closest to center of the line, looking at both forward & backward directions
                        K = InStr(Len(Words(I)) \ 2, Words(I), " ")
                        If K = 0 Then K = Len(Words(I))
                        lb = InStrRev(Words(I), " ", Len(Words(I)) \ 2)
                        If lb = 0 Then lb = Len(Words(I))
                        ' now we check to see which is closest to center & use that "space" as a line break
                        If Abs(Len(Words(I)) \ 2 - lb) > Abs(K - Len(Words(I)) \ 2) Then lb = K
                    End If
                End If
                If lb Then
                    ' line break found/created -- we need to make room for the new line in
                    ' the array. The line break generally won't be at the beginning or at
                    ' the end of the line
                    ReDim Preserve Words(0 To UBound(Words) + 1)
                    For K = UBound(Words) To I + 2 Step -1
                        Words(K) = Words(K - 1)
                    Next
                    ' made room for the new line, now we add that line to
                    ' the array and truncate the current line
                    Words(I + 1) = Mid(Words(I), lb + Choose(J, 4, 1))
                    Words(I) = Left(Words(I), lb - 1)
                    Exit For
                End If
            End If
        Next
        If lb Then Exit For
    Next
End If
If lb Or bInitial Then
    ' here we simply put together the line again but supply vbNewLine
    ' for line breaks.
    For J = 0 To UBound(Words) - 1
        sRtn = sRtn & Words(J) & vbNewLine
    Next
    sRtn = sRtn & Words(UBound(Words))
    sRtn = Replace$(sRtn, "<br>", " ")
    LineBreaks = Replace$(sRtn, "<br0>", " ")
    Exit Function
End If
End Function

Private Function ValidateImage(wCaption As String, Optional ImgType As Integer, _
            Optional ImgWidth As Long, Optional ImgHeight As Long) As Long

' ========================================================================
' This function will validate an image and return the handle, type, etc
' ========================================================================

Dim tIcon As Long, sValue As String
Dim bmpInfo As BITMAP, icoInfo As ICONINFO
Dim frmObject As Control
    
' look for the image to use for the menu item/sidebar, if any.
' A bit of a long process, but we try to guarantee that if an image
' was provided it actually exists, otherwise menus look crappy!
ReturnComponentValue wCaption, "IMG:", sValue
If Len(sValue) = 0 Then Exit Function
' images can be coded 3 different ways
' 1. As an actual handle
' 2. As an ImageList index
' 3. As a control (picbox, image1, etc)
If IsNumeric(sValue) Then
    tIcon = Val(sValue) ' image handle
Else
    If IsNumeric(Mid$(sValue, 2)) And Left$(sValue, 1) = "I" Then
        ' imagelist index
        If imgList Is Nothing Then
            tIcon = 0
        Else
            tIcon = Val(Mid$(sValue, 2))
            ' ensure the icon index passed is within the imagelist's range
            If tIcon <= imgList.ListImages.Count And tIcon > 0 Then
                tIcon = imgList.ListImages(tIcon).Picture.handle
            Else
                tIcon = 0
            End If
        End If
    Else    ' could still be a handle but in Hex form
        If IsNumeric(Mid$(sValue, 3)) And Left$(sValue, 2) = "&H" Then
            tIcon = Val(sValue)
        Else    ' last resort > control name instead of a handle or index
            ' when passed as a name vs handle, the control must exist on
            ' the form originally hosting the menu and must have a Picture property
            On Error Resume Next
            Set frmObject = SplitControlIndex(sValue)
            If frmObject Is Nothing Then
                tIcon = 0
            Else
                tIcon = frmObject.Picture.handle
                Set frmObject = Nothing
            End If
        End If
    End If
End If
If tIcon <> 0 Then
    ' here we test to see if it's an icon
    GetObject tIcon, Len(bmpInfo), bmpInfo
    If bmpInfo.bmBits Then
        ' calculate the image's height & width
        ImgWidth = bmpInfo.bmWidth
        ImgHeight = bmpInfo.bmHeight
        ValidateImage = tIcon
        ImgType = 1
    Else
        If GetIconInfo(tIcon, icoInfo) Then
            ' got an icon, let's kill the bitmaps returned by API
            If icoInfo.hbmMask <> 0 Then DeleteObject icoInfo.hbmMask
            If icoInfo.hbmColor <> 0 Then
                ' calculate the image's height & width
                GetObject icoInfo.hbmColor, Len(bmpInfo), bmpInfo
                ImgWidth = bmpInfo.bmWidth
                ImgHeight = bmpInfo.bmHeight
                DeleteObject icoInfo.hbmColor
                ValidateImage = tIcon
            End If
        End If
    End If
End If
End Function

Public Function GetShortFileName(ByVal strLongPath As String) As String
    Const cchBuffer = 300
    Dim strShortPath As String
    Dim lResult As Long, nPos As Long
'-----------------------------------------------------------
' FUNCTION GetShortPathName
' Retrieve the short pathname version of a path possibly
'   containing long subdirectory and/or file names
'-----------------------------------------------------------

' used by the FILES: flag to produce faster results when displaying files lists
' Inserted by LaVolpe
On Error GoTo Function_GetShortPathName_General_ErrTrap_by_LaVolpe
    strShortPath = String$(cchBuffer, 0)
    lResult = GetShortPathName(strLongPath, strShortPath, cchBuffer)
    If lResult = 0 Then
        'Just use the long name as this is usually good enough
        GetShortFileName = strLongPath
    Else

      nPos = InStr(strShortPath, vbNullChar)
      If nPos > 0 Then
          GetShortFileName = Left$(strShortPath, nPos - 1)
      Else
          GetShortFileName = strShortPath
      End If

    End If
Exit Function

Function_GetShortPathName_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Function GetShortPathName]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Function

'-----------------------------------------------------------
' FUNCTION: GetFileDescription
' Gets the file description information.
' IN: [strFilename] - name of file to get description of.
' Returns: Description (vbNullString if not found)
'-----------------------------------------------------------
' Retrieved function from MSDN

Public Function GetFileDescription(ByVal sFile As String) As String
    Const sEXE As String = "\FileDescription"
    Dim lVerSize As Long
    Dim lTemp As Long
    Dim lRet As Long
    Dim bInfo() As Byte
    Dim lpBuffer As Long
    Dim sDesc As String
    Dim sKEY As String
    '
    'Get the size of the file version info, allocate a buffer for it, and get the
    'version info.  Next, we query the Fixed file info portion, where the internal
    'file version used by the Windows VerInstallFile API is kept.  We then copy
    'the info into a string.
    '
    lVerSize = GetFileVersionInfoSize(sFile, lTemp)
    ReDim bInfo(lVerSize)
    If lVerSize > 0 Then
        lRet = GetFileVersionInfo(sFile, lTemp, lVerSize, VarPtr(bInfo(0)))
        If lRet <> 0 Then
            sKEY = GetNLSKey(bInfo)
            lRet = VerQueryValue(VarPtr(bInfo(0)), sKEY & sEXE, lpBuffer, lVerSize)
            If lRet <> 0 Then
                sDesc = Space$(lVerSize)
                lstrcpyn sDesc, lpBuffer, lVerSize
                GetFileDescription = sDesc
            End If
        End If
    End If
End Function

Private Function GetNLSKey(byteVerData() As Byte) As String
'-----------------------------------------------------------
' Function builds translation string, including the required
' language tag in order to extract the
' file description from the executable
' Retrieved function from MSDN
'-----------------------------------------------------------
Const strTRANSLATION$ = "\VarFileInfo\Translation"
Const strSTRINGFILEINFO$ = "\StringFileInfo\"
Const strDEFAULTNLSKEY$ = "040904E4"
Const LOCALE_IDEFAULTLANGUAGE& = &H9&
Const LOCALE_IDEFAULTCODEPAGE& = &HB&

Static strLANGCP As String

Dim lpBufPtr As Long
Dim strNLSKey As String
Dim fGotNLSKey As Integer
Dim intOffset As Integer
Dim lVerSize As Long
Dim lTmp As Long
Dim lBufLen As Long
Dim lLCID As Long
Dim strTmp As String

On Error GoTo GNLSKCleanup

If VerQueryValue(VarPtr(byteVerData(0)), strTRANSLATION, lpBufPtr, lVerSize) <> 0 Then ' (Pass byteVerData array via reference to first element)
    If Len(strLANGCP) = 0 Then
        lLCID = GetUserDefaultLCID()
        If lLCID > 0 Then
            strTmp = Space$(8)

            GetLocaleInfoA lLCID, LOCALE_IDEFAULTCODEPAGE, strTmp, 8
            strLANGCP = StringFromBuffer(strTmp)
            Do While Len(strLANGCP) < 4
                strLANGCP = "0" & strLANGCP
            Loop

            GetLocaleInfoA lLCID, LOCALE_IDEFAULTLANGUAGE, strTmp, 8
            strLANGCP = StringFromBuffer(strTmp) & strLANGCP
            Do While Len(strLANGCP) < 8
                strLANGCP = "0" & strLANGCP
            Loop
        End If
    End If

    If VerQueryValue(VarPtr(byteVerData(0)), strLANGCP, lTmp, lBufLen) <> 0 Then
        strNLSKey = strLANGCP
    Else
        For intOffset = 0 To lVerSize - 1 Step 4
            CopyMemory lTmp, ByVal lpBufPtr + intOffset, 4
            strTmp = Hex$(lTmp)
            Do While Len(strTmp) < 8
                strTmp = "0" & strTmp
            Loop

            strNLSKey = strSTRINGFILEINFO & Right$(strTmp, 4) & Left$(strTmp, 4)

            If VerQueryValue(VarPtr(byteVerData(0)), strNLSKey, lTmp, lBufLen) <> 0 Then
                fGotNLSKey = True
                Exit For
            End If
        Next

        If Not fGotNLSKey Then
            strNLSKey = strSTRINGFILEINFO & strDEFAULTNLSKEY
            If VerQueryValue(VarPtr(byteVerData(0)), strNLSKey, lTmp, lBufLen) <> 0 Then
                fGotNLSKey = True
            End If
        End If
    End If
End If
GNLSKCleanup:
    If fGotNLSKey Then
        GetNLSKey = strNLSKey
    End If
End Function

Private Function GetUNCPath(strDriveLetter As String) As String
'-----------------------------------------------------------
' Function returns the UNC drive path for mapped drives
'-----------------------------------------------------------
Dim lpszLocalName As String
Dim lpszRemoteName As String
Dim cbRemoteName As Long

lpszLocalName = strDriveLetter
If Right$(lpszLocalName, 1) = "\" Then lpszLocalName = Left$(lpszLocalName, Len(lpszLocalName) - 1)
lpszRemoteName = String$(255, Chr$(32))
cbRemoteName = Len(lpszRemoteName)
If WNetGetConnection(lpszLocalName, lpszRemoteName, cbRemoteName) Then
    GetUNCPath = "Unknown"
Else
    Dim sPath As String, sName As String
    sPath = StringFromBuffer(lpszRemoteName)
    ' here we are formatting the display similar to that on Win2K: folder on 'server'
    If Len(sPath) Then
        sName = StripFile(sPath, "N")
        sPath = StripFile(sPath, "P")
        If Right(sPath, 1) = "\" Then sPath = Left(sPath, Len(sPath) - 1)
        sPath = Replace$(sPath, "\\", "")
        GetUNCPath = " " & sName & " on '" & sPath & "'"
    End If
End If
End Function

Private Function DriveType(strDriveName As String) As String
'-----------------------------------------------------------
' Function returns the drive type
'-----------------------------------------------------------
Dim lngRet As Long
Dim strDrive As String
Const DRIVE_UNKNOWN = 0
Const DRIVE_ABSENT = 1
Const DRIVE_REMOVABLE = 2
Const DRIVE_FIXED = 3
Const DRIVE_REMOTE = 4
Const DRIVE_CDROM = 5
Const DRIVE_RAMDISK = 6

lngRet = GetDriveType(strDriveName)
Select Case lngRet
    Case DRIVE_ABSENT 'The root directory does not exist.
        strDrive = "Drive does not exist"
    Case DRIVE_REMOVABLE 'The drive can be removed from the drive.
        strDrive = "Removable Media"
    Case DRIVE_FIXED 'The disk cannot be removed from the drive.
        strDrive = "Fixed Drive"
    Case DRIVE_REMOTE  'The drive is a remote (network) drive.
        strDrive = "Network Drive"
    Case DRIVE_CDROM 'The drive is a CD-ROM drive.
        strDrive = "CD Rom"
    Case DRIVE_RAMDISK 'The drive is a RAM disk.
        strDrive = "Ram Disk"
    Case Else 'The drive type cannot be determined.
        strDrive = "Unknown Drive Type"
End Select
DriveType = strDrive
End Function

Private Sub ListAllDrives(vArray As Variant)
'-----------------------------------------------------------
' Function returns each drive letter and type
'-----------------------------------------------------------
' used for the lvDrives custom menu
Dim strAllDrives As String, strTmp As String, strRtn As String
Dim lngRet As Long
ReDim vArray(0 To 0)

' function to return all drive letters
strTmp = String$(255, 0)
lngRet = GetLogicalDriveStrings(255, strTmp)
strAllDrives = Left(strTmp, lngRet)
If strAllDrives <> "" Then
    Do
        ' process one drive letter at a time
        strTmp = StringFromBuffer(strAllDrives)
        ' truncate the string
        strAllDrives = Mid$(strAllDrives, Len(strTmp) + 2)
        ' now just add drives to array. Array will be in the format of
        ' drive letter | drive type
        strRtn = strTmp
        strTmp = DriveType(strTmp)
        If strTmp = "Network Drive" Then
            strTmp = GetUNCPath(strRtn)
            strRtn = strRtn & strTmp & "|Network/Mapped Drive"
        Else
            strRtn = strRtn & "|" & strTmp
        End If
        ReDim Preserve vArray(0 To UBound(vArray) + 1)
        vArray(UBound(vArray)) = strRtn
    Loop While strAllDrives <> ""
End If
End Sub

