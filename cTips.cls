VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' REMARKS CONSUME SOOOOO MUCH WHITE SPACE  :)
' Actual class is only 10 lines of code

' Primarily used to relay menu tips back to a form, this class also relays
' the special menus. See SendCustomSelection below for more details.

' To use the Tips, simply follow these steps.... and remember to include Tips in the menu captions
' Replace the variable myTips below with whatever you decide to use...
' 1. Declare the tips class in your form like so
'       Private WithEvents myTips As cTips
' 2. On your form's Load Event initialize the class and then pass class with the SetMenu call
'       Set myTips = New cTips
'       SetMenu Me.hWnd, [ImageList], myTips
' 3. On your form's Unload Event, remove the class
'       Set myTips = Nothing
' 4. In your form's new myTips_DisplayTip Event, simply display the incoming tip where you want...
'       Example:  StatusBar1.Panel(0).Text = TipText

' Note for MDI Child forms.  No need to include the tips class. Any MDI child will have their tips
' sent automatically to the parent MDI form. If you want the menu tips sent to the MDI child and not
' the parent, you will  need to...
'   -- include the Tips class as outlined above
'   -- call the optional function:  RerouteTips hWnd, TipsClass   (hWnd is the handle to the child form, TipsClass is the MDI child's intialized class)
'       this call would replace the SetMenu call in step 2 above and is not reversible back to the parent
'       Remember that MDI children do not call the SetMenu  function

Public Event DisplayTip(TipText As String)
Public Event CustomSelection(UserID As String, Category As String, Value As Variant)
Public Event MDIKeyDown(KeyCode As Integer, Shift As Integer)
Public Event MDIKeyUp(KeyCode As Integer, Shift As Integer)

Friend Sub SendTip(sTip As String)
' Called by the main message processing routine to send the menu tip back to the ACTIVE form
' Note that the active form may be different then the parent form hosting the menu. MDI children's
' active form is the parent.  A form right clicked on that displays the menu from another form is the
' active form, not the form where the menu came from.
' Exception could be MDI children if the notes at top of class above apply
On Error Resume Next
RaiseEvent DisplayTip(sTip)
End Sub

Friend Sub SendCustomSelection(sID As String, sCat As String, NewVal As Variant)
' 6 custom menus can be created without any references to any controls
' These menus will fire the clicked menu item back thru this routine to the
' form's cTips class reference.

' You must supply a user defined string variable for parameter: sID.
' This helps positively identify which menu item is being selected if you
' were to use more than one custom menu of the same type on the same menu panel

' sCat is the category of the special menus. See below for category names
' Value could be a long, string or date type value. See below.

' The 6 are described below as is their formatting string

' All other menu flags can be added (i.e., IMG:, Tip:, etc)

' 1. Submenu Contents: 23 colors (vbwhite, vbblack, etc) and an extra menu item
'                      to select any color. You must provide the color dialog
'    Code Format: {lvColors:[C]:ID:[id]|[other flags]}
'    -- C is the color to show as checked. If none, supply -1
'    Category: Color
'    Return value: Long value of the color selected
'       Note: return value of -1 indicates the menu item "Select another color"

' 2. Submenu Contents: The 12 months in one of 3 styles
'    Code Format: {lvMonths:[M]:Group:[G]:Year:[Y]:ID:[id]|[other flags]}
'    -- M is the numerical value of the month to show as checked.
'       -- To check system month supply 0. If none to check, supply -1
'    -- Year:[Y] is an optional flag/value. Default is current system year
'       -- if provided Y is the year for the month.
'    -- Group:[G] is an optional grouping of one of two types:
'       -- G is either CYQtr to separate months into calendar year quarters
'       -- or G is FYQtr to separate months into fiscal year quarters
'    Category: Month
'    Return value: Long value of the month (1 thru 12)

' 3. Submenu Contents: the 7 days of the week
'    Code Format: {lvDays:[D]:ID:[id]|[other flags]}
'    -- D is numerical value of weekday to show as checked.
'       -- To check system week supply 0. If none to check, supply -1
'       - the weekday is depenedent on system settings 1 could be Sunday or Monday
'    Category: WeekDay
'    Return value: long value representing weekday (1 thru 7)

' 4. Submenu Contents: the 50 United States and DC
'    Code Format: {lvStates:[S]:ID:[id]|[other flags]}
'    -- S is string 2-letter abbreviation of state to show as checked. If none, supply -1
'    Category: State
'    Return value: string value representing 2-letter state code (i.e., IL)

' 5. Submenu Contents: the dates within a specific month
'    Code Format: {lvMonth:[M]:Year:[Y]:Day:[D]:ID[id]|[other flags]}
'    -- M is the numerical value of the month to display.
'       -- Supply 0 to default to the current system month
'    -- Year:[Y] is an optional flag/value. Default is current system year
'       -- if provided Y is the year for the month
'    -- Day:[D] is an optional flag/value of the date to show as checked.
'       -- if provided D is day of month to show as checked.
'       -- To check system date, supply 0.
'       -- To check none, supply -1 or exclude flag/value from caption
'    Category: DayOfMonth
'    Return value: Date value representing a date selected

' 6. Submenu Contents: the installed fonts on the system
'    Code Format: {lvFonts:[F]:Type:[T]:Group:[G]:ID:[id]|[other flags]}
'    -- F is the font name to show as checked. If none, supply -1
'    -- Type:[T] is an optional flag/value. Default is ALL fonts
'       -- if provided T must be either System or TrueType and will
'          filter only those types of fonts
'    -- Group:[G] is an optional flag/value to filter fonts by name
'       -- if provided the format is "A-Z" where A is the first letter to
'          start the filter and Z is the last letter.
'          So "E-G" will filter fonts between the letters E and G.
'    Category: Font
'    Return value: string value of the Font Name selected

On Error Resume Next
RaiseEvent CustomSelection(sID, sCat, NewVal)
End Sub


Friend Sub SendMDIKeyPress(ByVal KeyCodex As Long, ByVal ShiftCodex As Long, ByVal KeyUp As Boolean)
If KeyUp Then
    RaiseEvent MDIKeyUp(CInt(KeyCodex), CInt(ShiftCodex))
Else
    RaiseEvent MDIKeyDown(CInt(KeyCodex), CInt(ShiftCodex))
End If
End Sub

