Attribute VB_Name = "Module1"
Option Explicit
'
' Copyright © 1997-1999 Brad Martinez, http://www.mvps.org
'
Public Type POINTAPI   ' pt
  x As Long
  y As Long
End Type

Public Type RECT   ' rct
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As KeyCodeConstants) As Integer

Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                            (ByVal hWnd As Long, _
                            ByVal wMsg As Long, _
                            ByVal wParam As Long, _
                            lParam As Any) As Long   ' <---

Declare Function PostMessage Lib "user32" Alias "PostMessageA" _
                            (ByVal hWnd As Long, _
                            ByVal wMsg As Long, _
                            ByVal wParam As Long, _
                            lParam As Any) As Long   ' <---

' =========================================================
' user-defined listview definitions

' value returned by many listview messages indicating
' the index of no listview item
Public Const LVI_NOITEM = -1

' Listview item state image index values
Public Enum LVItemStates
  lvisNoButton = 0
  lvisCollapsed = 1
  lvisExpanded = 2
End Enum

' ListView_GetRelativeItem flags
Public Enum LVRelativeItemFlags
  lvriParent = 0
  lvriChild = 1
  lvriFirstSibling = 2
  lvriLastSibling = 3
  lvriPrevSibling = 4
  lvriNextSibling = 5
End Enum

' ListView_GetItemCountEx flags
Public Enum LVItemCountFlags
  lvicParents = 0
  lvrcChildren = 1
  lvicSiblings = 2
End Enum

' =========================================================
' listview definitions

' we're using IE3 definitions...
#Const WIN32_IE = &H300

' messages
Public Const LVM_FIRST = &H1000
Public Const LVM_SETIMAGELIST = (LVM_FIRST + 3)
Public Const LVM_GETITEMCOUNT = (LVM_FIRST + 4)
Public Const LVM_GETITEM = (LVM_FIRST + 5)
Public Const LVM_SETITEM = (LVM_FIRST + 6)
Public Const LVM_GETNEXTITEM = (LVM_FIRST + 12)
Public Const LVM_HITTEST = (LVM_FIRST + 18)
Public Const LVM_ENSUREVISIBLE = (LVM_FIRST + 19)
Public Const LVM_SETCOLUMNWIDTH = (LVM_FIRST + 30)
Public Const LVM_SETITEMSTATE = (LVM_FIRST + 43)
#If (WIN32_IE >= &H300) Then
Public Const LVM_SETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 54)
#End If

' LVM_SETIMAGELIST wParam value
Public Const LVSIL_STATE = 2

' LVM_GET/SETITEM lParam
Public Type LVITEM   ' was LV_ITEM
  mask As Long
  iItem As Long
  iSubItem As Long
  state As Long
  stateMask As Long
  pszText As Long  ' if String, must be pre-allocated before before filled
  cchTextMax As Long
  iImage As Long
  lParam As Long
#If (WIN32_IE >= &H300) Then
  iIndent As Long
#End If
End Type

' LVITEM mask
Public Const LVIF_STATE = &H8
#If (WIN32_IE >= &H300) Then
Public Const LVIF_INDENT = &H10
#End If
 
' LVITEM state, stateMask
Public Const LVIS_FOCUSED = &H1
Public Const LVIS_SELECTED = &H2
Public Const LVIS_STATEIMAGEMASK = &HF000

' LVM_GETNEXTITEM lParam
Public Const LVNI_FOCUSED = &H1
Public Const LVNI_SELECTED = &H2

' LVM_HITTEST lParam
Public Type LVHITTESTINFO   ' was LV_HITTESTINFO
  pt As POINTAPI
  flags As Long
  iItem As Long
#If (WIN32_IE >= &H300) Then
  iSubItem As Long    ' this is was NOT in win95.  valid only for LVM_SUBITEMHITTEST
#End If
End Type
 
' LVHITTESTINFO flags
Public Const LVHT_ONITEMICON = &H2
Public Const LVHT_ONITEMLABEL = &H4
Public Const LVHT_ONITEMSTATEICON = &H8

' LVM_SETEXTENDEDLISTVIEWSTYLE wParam/lParam
#If (WIN32_IE >= &H300) Then
Public Const LVS_EX_FULLROWSELECT = &H20   ' // applies to report mode only
#End If
'

' =========================================================
' user-defined listview macros

' Returns the state and indent vaues of the specified listview item

Public Function Listview_GetItemStateEx(hwndLV As Long, iItem As Long, iIndent As Long) As LVItemStates
  Dim lvi As LVITEM
  
  lvi.mask = LVIF_STATE Or LVIF_INDENT
  lvi.iItem = iItem
  lvi.stateMask = LVIS_STATEIMAGEMASK
  
  If ListView_GetItem(hwndLV, lvi) Then
    iIndent = lvi.iIndent
    Listview_GetItemStateEx = STATEIMAGEMASKTOINDEX(lvi.state And LVIS_STATEIMAGEMASK)
  End If
  
End Function

' Sets the state and indent vaues of the specified listview item

Public Function Listview_SetItemStateEx(hwndLV As Long, iItem As Long, iIndent As Long, dwState As LVItemStates) As Boolean
  Dim lvi As LVITEM
  
  lvi.mask = LVIF_STATE Or LVIF_INDENT
  lvi.iItem = iItem
  lvi.state = INDEXTOSTATEIMAGEMASK(dwState)
  lvi.stateMask = LVIS_STATEIMAGEMASK
  lvi.iIndent = iIndent
  
  Listview_SetItemStateEx = ListView_SetItem(hwndLV, lvi)
  
End Function

' Returns the zero-based position (index) of the item within the listview,
' that is selected and has the focus rectangle (as opposed to the item's
' one-based ListItems collection Index value)

Public Function ListView_GetSelectedItem(hwndLV As Long) As Long
  ListView_GetSelectedItem = ListView_GetNextItem(hwndLV, -1, LVNI_FOCUSED Or LVNI_SELECTED)
End Function
 
' Selects the specified item and gives it the focus rectangle.
' does not de-select any currently selected items (user-defined).

Public Function ListView_SetFocusedItem(hwndLV As Long, i As Long) As Boolean
  ListView_SetFocusedItem = ListView_SetItemState(hwndLV, i, LVIS_FOCUSED Or LVIS_SELECTED, _
                                                                                                    LVIS_FOCUSED Or LVIS_SELECTED)
End Function

' Returns the zero-based position (index) of the item within the listview
' that has the specified relationship to the specified item. Emmulates
' TreeView Node (and real treeview item) relational properties (most of
' the code in this proc isn't used, but is just here to show how its done).

Public Function ListView_GetRelativeItem(hwndLV As Long, iItem As Long, dwRelative As LVRelativeItemFlags) As Long
  Dim iIndentSrc As Long
  Dim iIndentTarget As Long
  Dim i As Long
  Dim nItems As Long
  Dim iSave As Long
  
  ' get the source item's indent value, exit on failure
  iIndentSrc = -1
  Call Listview_GetItemStateEx(hwndLV, iItem, iIndentSrc)
  If (iIndentSrc = -1) Then
    ListView_GetRelativeItem = LVI_NOITEM
    Exit Function
  End If
  
  i = iItem
  nItems = ListView_GetItemCount(hwndLV)
  
  Select Case dwRelative
    
    ' =====================================================
    ' Works up the tree from the source item and returns the first item whose
    ' indent is less than the source item's indent, returns LVI_NOITEM otherwise.
    
    Case lvriParent
      Do
        i = i - 1
        If (i = LVI_NOITEM) Then Exit Do
        Call Listview_GetItemStateEx(hwndLV, i, iIndentTarget)
      Loop Until (iIndentTarget < iIndentSrc)
  
    ' =====================================================
    ' If the indent of the first item below the source item is greater than the
    ' source item's indent, returns that item, returns LVI_NOITEM otheriwse.
    
    Case lvriChild
      If (i = (nItems - 1)) Then
        i = LVI_NOITEM
      Else
        i = i + 1
        Call Listview_GetItemStateEx(hwndLV, i, iIndentTarget)
        If (iIndentTarget <= iIndentSrc) Then i = LVI_NOITEM
      End If
  
    ' =====================================================
    ' Works up the tree from the source item and returns the last item whose
    ' indent matches the source item's indent. Keeps going until an indent
    ' value less than the source item's is found. Returns the source item if it's
    ' the first sibling (not implemented in the real treeview)
    
    Case lvriFirstSibling
      iSave = i
      Do
        i = i - 1
        If (i = LVI_NOITEM) Then Exit Do
        Call Listview_GetItemStateEx(hwndLV, i, iIndentTarget)
        If (iIndentTarget = iIndentSrc) Then
          iSave = i
        ElseIf (iIndentTarget < iIndentSrc) Then
          Exit Do
        End If
      Loop
      i = iSave
      
    ' =====================================================
    ' Works down the tree from the source item and returns the last item whose
    ' indent matches the source item's indent. Keeps going until an indent
    ' value less than the source item's is found. Returns the source item if it's
    ' the last sibling (not implemented in the real treeview)
    
    Case lvriLastSibling
      iSave = i
      Do
        i = i + 1
        If (i = nItems) Then Exit Do
        Call Listview_GetItemStateEx(hwndLV, i, iIndentTarget)
        If (iIndentTarget = iIndentSrc) Then
          iSave = i
        ElseIf (iIndentTarget < iIndentSrc) Then
          Exit Do
        End If
      Loop
      i = iSave
  
    ' =====================================================
    ' Works up the tree from the source item and returns the first item whose
    ' indent matches the source item's indent. Keeps going until an indent
    ' value less than the source item's is found. Returns LVI_NOITEM if
    ' there is no previous sibling.
    
    Case lvriPrevSibling
      Do
        i = i - 1
        If (i = LVI_NOITEM) Then Exit Do
        Call Listview_GetItemStateEx(hwndLV, i, iIndentTarget)
        If (iIndentTarget = iIndentSrc) Then
          Exit Do
        ElseIf (iIndentTarget < iIndentSrc) Then
          i = LVI_NOITEM
          Exit Do
        End If
      Loop
      
    ' =====================================================
    ' Works down the tree from the source item and returns the first item whose
    ' indent matches the source item's indent. Keeps going until an indent
    ' value less than the source item's is found. Returns LVI_NOITEM if
    ' there is no next sibling.
    
    Case lvriNextSibling
      Do
        i = i + 1
        If (i = nItems) Then
          i = LVI_NOITEM
          Exit Do
        End If
        Call Listview_GetItemStateEx(hwndLV, i, iIndentTarget)
        If (iIndentTarget = iIndentSrc) Then
          Exit Do
        ElseIf (iIndentTarget < iIndentSrc) Then
          i = LVI_NOITEM
          Exit Do
        End If
      Loop
    
    ' =====================================================
    Case Else
      i = LVI_NOITEM
    
  End Select
  
  ListView_GetRelativeItem = i

End Function

' Returns the count of items within the listview that have the specified
' relationship to the specified item (this proc isn't used, but is just here
' to show how its done).

Public Function ListView_GetItemCountEx(hwndLV As Long, iItem As Long, dwRelative As LVItemCountFlags) As Long
  Dim i As Long
  Dim nItems As Long
  
  Select Case dwRelative
            
    ' =====================================================
    ' Also indicates the item's effective zero-based level in the tree hierarchy
    
    Case lvicParents
      nItems = -1
      Do
        nItems = nItems + 1
        i = ListView_GetRelativeItem(hwndLV, i, lvriParent)
      Loop Until (i = LVI_NOITEM)
      
    ' =====================================================
    Case lvrcChildren
      i = ListView_GetRelativeItem(hwndLV, i, lvriChild)
      Do Until (i = LVI_NOITEM)
        nItems = nItems + 1
        i = ListView_GetRelativeItem(hwndLV, i, lvriNextSibling)
      Loop
    
    ' =====================================================
    Case lvicSiblings
      i = ListView_GetRelativeItem(hwndLV, i, lvriFirstSibling)
      Do Until (i = LVI_NOITEM)
        nItems = nItems + 1
        i = ListView_GetRelativeItem(hwndLV, i, lvriNextSibling)
      Loop
    
  End Select
  
  ListView_GetItemCountEx = nItems

End Function

' =========================================================
' listview macros

Public Function ListView_SetImageList(hWnd As Long, himl As Long, iImageList As Long) As Long
  ListView_SetImageList = SendMessage(hWnd, LVM_SETIMAGELIST, ByVal iImageList, ByVal himl)
End Function
 
Public Function ListView_GetItemCount(hWnd As Long) As Long
  ListView_GetItemCount = SendMessage(hWnd, LVM_GETITEMCOUNT, 0, 0)
End Function

Public Function ListView_GetItem(hWnd As Long, pitem As LVITEM) As Boolean
  ListView_GetItem = SendMessage(hWnd, LVM_GETITEM, 0, pitem)
End Function
 
Public Function ListView_SetItem(hWnd As Long, pitem As LVITEM) As Boolean
  ListView_SetItem = SendMessage(hWnd, LVM_SETITEM, 0, pitem)
End Function
 
Public Function ListView_GetNextItem(hWnd As Long, i As Long, flags As Long) As Long
  ListView_GetNextItem = SendMessage(hWnd, LVM_GETNEXTITEM, ByVal i, ByVal flags)   ' ByVal MAKELPARAM(flags, 0))
End Function

Public Function ListView_HitTest(hwndLV As Long, pinfo As LVHITTESTINFO) As Long
  ListView_HitTest = SendMessage(hwndLV, LVM_HITTEST, 0, pinfo)
End Function
 
Public Function ListView_EnsureVisible(hwndLV As Long, i As Long, fPartialOK As Boolean) As Boolean
  ListView_EnsureVisible = SendMessage(hwndLV, LVM_ENSUREVISIBLE, ByVal i, ByVal Abs(fPartialOK))   ' ByVal MAKELPARAM(Abs(fPartialOK), 0))
End Function

Public Function ListView_SetColumnWidth(hWnd As Long, iCol As Long, cx As Long) As Boolean
  ListView_SetColumnWidth = SendMessage(hWnd, LVM_SETCOLUMNWIDTH, ByVal iCol, ByVal cx) ' ByVal MAKELPARAM(cx, 0))
End Function

Public Function ListView_SetItemState(hwndLV As Long, i As Long, state As Long, mask As Long) As Boolean
  Dim lvi As LVITEM
  lvi.state = state
  lvi.stateMask = mask
  ListView_SetItemState = SendMessage(hwndLV, LVM_SETITEMSTATE, ByVal i, lvi)
End Function

#If (WIN32_IE >= &H300) Then

Public Function ListView_SetExtendedListViewStyleEx(hwndLV As Long, dwMask As Long, dw As Long) As Long
  ListView_SetExtendedListViewStyleEx = SendMessage(hwndLV, LVM_SETEXTENDEDLISTVIEWSTYLE, _
                                                                                          ByVal dwMask, ByVal dw)
End Function
'
#End If   ' (WIN32_IE >= &H300)
'

' =========================================================
' imagelist macros

' Returns the one-based index of the specifed state image mask, shifted
' left twelve bits. A common control utility macro.

' Prepares the index of a state image so that a tree view control or list
' view control can use the index to retrieve the state image for an item.

Public Function INDEXTOSTATEIMAGEMASK(iIndex As Long) As Long
' #define INDEXTOSTATEIMAGEMASK(i) ((i) << 12)
  INDEXTOSTATEIMAGEMASK = iIndex * (2 ^ 12)
End Function

' Returns the state image index from the one-based index state image mask.
' The inverse of INDEXTOSTATEIMAGEMASK.

' A user-defined function (not in Commctrl.h)

Public Function STATEIMAGEMASKTOINDEX(iState As Long) As Long
  STATEIMAGEMASKTOINDEX = iState / (2 ^ 12)
End Function
