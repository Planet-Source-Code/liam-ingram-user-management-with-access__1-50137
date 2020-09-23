VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsXPMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'* Menu properties
Private mnuName As String

'* Menu constants
Const XBuffer As Long = 4
Const YBuffer As Long = 4

Const clr_Background As Long = &HF7F8F9
Const clr_LeftMargin As Long = &HD1D8D8
Const clr_MenuBorder As Long = &H666666
Const clr_HilightBack As Long = &HD2BDB6
Const clr_HilightBorder As Long = &H6A240A

Const dim_MarginWidth As Long = 23
Const fnt_MenuItem  As String = "Tahoma"

'* Width
Private mnuWidth As Long
Private theTextHeight As Long

Private frmMenu As New frmXPMenu
Private ActivePopup As New clsXPMenu

Private bVisible As Boolean
Private bPopupShown As Boolean
Private Yhilight As Long

'* image list
Private imageLst As ImageList

'* Menu array
Private MenuItems()     As typMenuItem
Private MenuItemCount   As Long
Private TextItemCnt     As Long
Private SepItemCnt      As Long
Private hilightedItem   As Long

'* Types
Private Type typMenuItem
    IconNum     As Long
    Text        As String
    bPopupmenu  As Boolean
    mnuSubMenu  As clsXPMenu
    bSeperator  As Boolean
End Type
    
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020
Private Type POINTAPI
        x As Long
        y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Sub AddItem(IconNum As Long, Text As String, bPopupmenu As Boolean, bSeperator As Boolean, Optional mnuSubitem As clsXPMenu = Nothing)
    
    MenuItemCount = MenuItemCount + 1
    ReDim Preserve MenuItems(1 To MenuItemCount) As typMenuItem
    
    With MenuItems(MenuItemCount)
        .IconNum = IconNum
        .Text = Text
        .bPopupmenu = bPopupmenu
        .bSeperator = bSeperator
        If (mnuSubitem Is Nothing) Then Else Set .mnuSubMenu = mnuSubitem
    End With
        
    If bSeperator Then
        SepItemCnt = SepItemCnt + 1
    Else
        TextItemCnt = TextItemCnt + 1
    End If
    
    Dim theWidth As Integer
    With frmMenu
        theWidth = .TextWidth(Text) + (XBuffer * 4) + 2 + dim_MarginWidth + 2 '2=border
        
        If bPopupmenu Then
            theWidth = theWidth + (XBuffer * 2) + frmMenu.picPopup.TextWidth("4")
        End If
        
        If theWidth > mnuWidth Then mnuWidth = theWidth
    End With
    
End Sub

Function GetHilightNum() As Integer
    GetHilightNum = hilightedItem
End Function

Public Function GetItemText(itemNum As Integer) As String
    If itemNum > MenuItemCount Then
        GetItemText = ""
        Exit Function
    End If
    
    GetItemText = MenuItems(itemNum).Text
End Function

Public Function GetMenuName()
    GetMenuName = mnuName
End Function

Public Function IsTextItem(itemNum As Integer) As Boolean
    If itemNum > MenuItemCount Then
        IsTextItem = False
        Exit Function
    End If
    
    If MenuItems(itemNum).bPopupmenu Or MenuItems(itemNum).bSeperator Then
        IsTextItem = False
    Else
        IsTextItem = True
    End If
End Function

Function IsVisible() As Boolean
    IsVisible = bVisible
End Function

Public Sub KillAllMenus()
    Dim frm As Form
    For Each frm In Forms
        If frm.Tag = "XPMenu" Then
            frm.XPMenuClass.KillPopupMenus
            frm.XPMenuClass.UnloadMenu
        End If
    Next frm
End Sub

Public Sub KillPopupMenus()
    Dim I As Long
    
    For I = 1 To MenuItemCount
        If MenuItems(I).bPopupmenu Then
            MenuItems(I).mnuSubMenu.KillPopupMenus
            MenuItems(I).mnuSubMenu.UnloadMenu
        End If
    Next I
End Sub

Public Function PopupShown() As Boolean
     PopupShown = bPopupShown
End Function

Sub ShowMenu(x As Long, y As Long)
    If bVisible = True Then Exit Sub

    frmMenu.Left = x * 15
    frmMenu.Top = y * 15
    DrawMenu
    
    bVisible = True7
    bPopupShown = False
    frmMenu.tmrActive.Enabled = True
    frmMenu.tmrHover.Enabled = True
    frmMenu.Tag = "XPMenu"
End Sub

Public Sub DrawMenu()
    
    Dim oldFont As String, oldSize As Integer
    
    With frmMenu.picMenuBuffer
        .Cls
        .BackColor = clr_Background 'background
        .Height = GetHeight()
        .Width = mnuWidth
        
        '* Border
        frmMenu.picMenuBuffer.Line (0, 0)-(.ScaleWidth - 1, .ScaleHeight - 1), clr_MenuBorder, B
        
        '* Margin
        frmMenu.picMenuBuffer.Line (1, 1)-(dim_MarginWidth + 1, .ScaleHeight - 2), clr_LeftMargin, BF
        
        '* X, Y info
        Dim Xcur As Long, Ycur As Long, index As Integer
        Ycur = 3
        
        For index = 1 To MenuItemCount
            Xcur = dim_MarginWidth + (XBuffer * 2) + 1 '* 1 for the border
        
            '* hilighted?
            If hilightedItem = index And MenuItems(index).bSeperator = False Then
                Yhilight = Ycur
                frmMenu.picMenuBuffer.Line (3, Ycur)-(.ScaleWidth - 4, Ycur + (YBuffer * 2) + theTextHeight), clr_HilightBack, BF
                frmMenu.picMenuBuffer.Line (3, Ycur)-(.ScaleWidth - 4, Ycur + (YBuffer * 2) + theTextHeight), clr_HilightBorder, B
                frmMenu.picIcon.BackColor = clr_HilightBack
            Else
                frmMenu.picIcon.BackColor = clr_LeftMargin
            End If
            
            '* bit icon
            If imageLst Is Nothing Then
            Else
                If MenuItems(index).IconNum <> 0 Then
                    frmMenu.picIcon.Picture = imageLst.ListImages.item(MenuItems(index).IconNum).Picture
                    BitBlt .hdc, (dim_MarginWidth - 16) \ 2 + 3, Ycur + (((theTextHeight + (YBuffer * 2)) - 16) \ 2), 16, 16, frmMenu.picIcon.hdc, 0, 0, SRCCOPY
                End If
            End If
                
            '* popup menu
            If MenuItems(index).bPopupmenu Then
                oldFont = frmMenu.picMenuBuffer.FontName
                oldSize = frmMenu.picMenuBuffer.FontSize
                frmMenu.picMenuBuffer.FontName = "Marlett"
                frmMenu.picMenuBuffer.FontSize = 10
                TextOut .hdc, .ScaleWidth - .TextHeight("4") - XBuffer, Ycur + (((theTextHeight + (YBuffer * 2)) - 16) \ 2) + 2, "4", 1
                frmMenu.picMenuBuffer.FontName = oldFont
                frmMenu.picMenuBuffer.FontSize = oldSize
                
                'BitBlt .hdc, .ScaleWidth - frmMenu.picPopup.ScaleWidth - XBuffer, Ycur + (((theTextHeight + (YBuffer * 2)) - 16) \ 2) + 2, frmMenu.picPopup.ScaleWidth, frmMenu.picPopup.ScaleHeight, frmMenu.picPopup.hdc, 0, 0, SRCCOPY
            End If
            
            '* draw item
            If MenuItems(index).bSeperator Then
                frmMenu.picMenuBuffer.Line (dim_MarginWidth + 1, Ycur + YBuffer)-(.ScaleWidth - 1, Ycur + YBuffer), clr_LeftMargin
                Ycur = Ycur + 1 + (XBuffer * 2)
            Else
                TextOut .hdc, Xcur, Ycur + YBuffer, MenuItems(index).Text, Len(MenuItems(index).Text)
                Ycur = Ycur + theTextHeight + (YBuffer * 2)
            End If
            
        Next index
        
    End With
    
    frmMenu.Width = frmMenu.picMenuBuffer.Width * 15
    frmMenu.Height = frmMenu.picMenuBuffer.Height * 15
    frmMenu.Picture = frmMenu.picMenuBuffer.Image
    frmMenu.Show
End Sub

Function GetHeight() As Long
    Dim lngHeight As Long
    
    With frmMenu.picMenuBuffer
        Dim Ycur As Long, index As Integer
        Ycur = 3
        
        For index = 1 To MenuItemCount
            '* draw item
            If MenuItems(index).bSeperator Then
                Ycur = Ycur + 1 + (XBuffer * 2)
            Else
                Ycur = Ycur + theTextHeight + (YBuffer * 2)
            End If
        Next index
    End With
    
    lngHeight = Ycur + 4
    GetHeight = lngHeight
End Function

Public Function GetHilightedItem(y As Single) As Integer
    On Error GoTo endd
    
    With frmMenu.picMenuBuffer
        '* X, Y info
        Dim Ycur As Long, index As Integer
        Ycur = 3
        
        For index = 1 To MenuItemCount
            If MenuItems(index).bSeperator Then
                If y >= Ycur And (y <= Ycur + (YBuffer * 2) + 1) Then
                    GetHilightedItem = index
                    Exit Function
                End If
                Ycur = Ycur + 1 + (XBuffer * 2)
            Else
                'TextOut .hdc, Xcur, Ycur + YBuffer, MenuItems(index).Text, Len(MenuItems(index).Text)
                If y >= Ycur And (y <= Ycur + theTextHeight + (YBuffer * 2)) Then
                    GetHilightedItem = index
                    Exit Function
                End If
                Ycur = Ycur + theTextHeight + (YBuffer * 2)
            End If
            
        Next index
        
    End With
    Exit Function
endd:
End Function

Sub Init(strMenuName As String, Optional imageListBind As ImageList)
    
    mnuName = strMenuName
    
    Set frmMenu.XPMenuClass = Me
    
    If imageListBind Is Nothing Then Else Set imageLst = imageListBind
    frmMenu.FontName = fnt_MenuItem
    frmMenu.picMenuBuffer.FontName = fnt_MenuItem
    
    theTextHeight = frmMenu.picMenuBuffer.TextHeight("gW")

    MenuItemCount = 0
    SepItemCnt = 0
    TextItemCnt = 0
    hilightedItem = 0
    'ReDim MenuItems(MenuItemCount) As typMenuItem
End Sub


Public Sub MoveMenu(Lft As Long, Tp As Long)
    frmMenu.Left = Lft
    frmMenu.Top = Tp
End Sub

Public Sub setHilightedItem(item As Integer)
    If item = 0 Or hilightedItem = item Then Exit Sub
    
    If item = -1 Then
        hilightedItem = -1
        DrawMenu
        Exit Sub
    End If
    
    hilightedItem = item
    KillPopupMenus
    bPopupShown = False
    
    DrawMenu
    
    If MenuItems(item).bPopupmenu Then
        bPopupShown = True
        Set ActivePopup = MenuItems(item).mnuSubMenu
        If ActivePopup.IsVisible Then Exit Sub
        
        ActivePopup.ShowMenu frmMenu.Left \ 15 + frmMenu.Width \ 15 - 5, frmMenu.Top \ 15 + Yhilight
    End If
    
End Sub


Public Sub UnloadMenu()
    Unload frmMenu
    bVisible = False
    hilightedItem = 0
    
    frmMenu.tmrActive.Enabled = False
    frmMenu.tmrHover.Enabled = False
End Sub


