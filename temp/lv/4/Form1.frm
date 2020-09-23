VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "VB ListView item tree demo"
   ClientHeight    =   5490
   ClientLeft      =   1980
   ClientTop       =   1815
   ClientWidth     =   9870
   ClipControls    =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   9870
   Begin MSComctlLib.ImageList ilStateIcons 
      Left            =   4860
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ilSmallIcons 
      Left            =   4200
      Top             =   4710
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4275
      Left            =   240
      TabIndex        =   0
      Top             =   330
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   7541
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' Copyright © 1997-1999 Brad Martinez, http://www.mvps.org
'
' Demonstrates how to indent and set state images of VB ListView
' ListItems, simulating a hierarchical item tree (TreeView)
'
Private m_hwndLV  As Long    ' ListView1.hWnd


Private Sub Form_Load()
  Dim i As Integer
  
  ' Set the Form's ScaleMode to pixels for column resizing in Form_Resize
  ScaleMode = vbPixels
  
  With ilSmallIcons
    .ImageWidth = 16
    .ImageHeight = 16
    .ListImages.Add Picture:=Icon
  End With
  
  With ilStateIcons
    .ImageWidth = 16
    .ImageHeight = 16
    .ListImages.Add Picture:=LoadPicture("Collapsed.ico")
    .ListImages.Add Picture:=LoadPicture("Expanded.ico")
  End With
  
  ' Initialize and fill up the ListView

  With ListView1
    For i = 1 To 4
      .ColumnHeaders.Add Text:="column" & i
    Next
    .LabelEdit = lvwManual
    .SmallIcons = ilSmallIcons
    .View = lvwReport
    m_hwndLV = .hWnd
  End With


  ' Assign the VB ImageList as the ListView's state imagelist.
  Call ListView_SetImageList(m_hwndLV, ilStateIcons.hImageList, LVSIL_STATE)
  
  
  ' Add 10 root item's with collapsed buttons and no indent.
  Call AddChildItems(LVI_NOITEM, -1, 10)

End Sub

Private Sub Form_Resize()
  Static rc As RECT
  
  Call GetClientRect(m_hwndLV, rc)
  Call ListView_SetColumnWidth(m_hwndLV, 0, rc.Right * 0.5)
  Call ListView_SetColumnWidth(m_hwndLV, 1, (rc.Right * 0.5) \ 3)
  Call ListView_SetColumnWidth(m_hwndLV, 2, (rc.Right * 0.5) \ 3)
  Call ListView_SetColumnWidth(m_hwndLV, 3, (rc.Right * 0.5) \ 3)

End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim lvhti As LVHITTESTINFO
  Dim dwState As LVItemStates
  Dim iIndent As Long
    
  If (Button = vbLeftButton) Then
    
    ' Get the zero-based index of the item under the cursor (if any).
    lvhti.pt.x = x / Screen.TwipsPerPixelX
    lvhti.pt.y = y / Screen.TwipsPerPixelY
    If (ListView_HitTest(m_hwndLV, lvhti) <> LVI_NOITEM) Then   ' also returns iItem
      
      ' If the item's state icon was left-clicked...
      If (lvhti.flags = LVHT_ONITEMSTATEICON) Then
        
      ' Get the item's indent and state values
      dwState = Listview_GetItemStateEx(m_hwndLV, lvhti.iItem, iIndent)
      
      ' If the item is collaped, expanded it, otherwise collapse it
      If (dwState = lvisCollapsed) Then
        Call AddChildItems(lvhti.iItem, iIndent, 10)
      Else
        Call RemoveChildItems(lvhti.iItem, iIndent)
      End If
        
      End If   ' (lvhti.flags And LVHT_ONITEMSTATEICON)
    End If   ' ListView_HitTest
  End If   ' (Button = vbLeftButton)
  
End Sub

' Toggles the expanded state of the parent item whose icon or label was
' double clicked with the left mouse button.

Private Sub ListView1_DblClick()
  Dim lvhti As LVHITTESTINFO
  Dim dwState As LVItemStates
  Dim iIndent As Long
  
  ' If a left button double-click... (change to suit)
  If (GetKeyState(vbKeyLButton) And &H8000) Then
  
    ' Get the left clicked item
    Call GetCursorPos(lvhti.pt)
    Call ScreenToClient(m_hwndLV, lvhti.pt)
    Call ListView_HitTest(m_hwndLV, lvhti)  ' also returns iItem
    
    ' If the item's icon or label is double clicked...
    If (lvhti.flags And (LVHT_ONITEMICON Or LVHT_ONITEMLABEL)) Then
    
      ' Get the left clicked item's indent and state values
      dwState = Listview_GetItemStateEx(m_hwndLV, lvhti.iItem, iIndent)
      
      ' If the item is collaped, expanded it, otherwise collapse it
      If (dwState = lvisCollapsed) Then
        Call AddChildItems(lvhti.iItem, iIndent, 10)
      Else
        Call RemoveChildItems(lvhti.iItem, iIndent)
      End If
    
    End If   ' ListView_HitTest
  End If   ' (Button = vbLeftButton)
  
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim iItem As Long
  Dim dwState As LVItemStates
  Dim iIndent As Long
  
  
  ' Get the selected item...
  iItem = ListView_GetSelectedItem(m_hwndLV)
  If (iItem <> LVI_NOITEM) Then
    
    ' Get the item's indent and state values
    dwState = Listview_GetItemStateEx(m_hwndLV, iItem, iIndent)
  
    Select Case KeyCode
      
      ' ========================================================
      ' The right arrow key expands the selected item, then selects the current
      ' item's first child
      
      Case 187, 107 'plus
         KeyCode = 0
    
        ' If the item is collaped, expanded it, otherwise select
        ' the first child of the selected item (if any)
        If (dwState = lvisCollapsed) Then
          Call AddChildItems(iItem, iIndent, 10)
        
        ElseIf (dwState = lvisExpanded) Then
          iItem = ListView_GetRelativeItem(m_hwndLV, iItem, lvriChild)
          If (iItem <> LVI_NOITEM) Then Call ListView_SetFocusedItem(m_hwndLV, iItem)
        End If
        
      ' ========================================================
      ' The left arrow key collapses the selected item, then selects the current
      ' item's parent. The backspace key only selects the current item's parent
      
      Case 109, 189 'minus
         KeyCode = 0
          Call RemoveChildItems(iItem, iIndent)
    
    End Select   ' KeyCode
  End If   ' (iItem <> LVI_NOITEM)
  
End Sub

Private Sub AddChildItems(iParentItem As Long, iParentIndent As Long, nChildren As Long)
  Dim i As Integer
  Dim liChild As ListItem
  
  Screen.MousePointer = vbHourglass
  
  ' If a parent index is specified, change its button to expanded.
  If (iParentItem <> LVI_NOITEM) Then
    Call Listview_SetItemStateEx(m_hwndLV, iParentItem, iParentIndent, lvisExpanded)
  End If
  
  For i = 1 To nChildren
    ' Add child items sequentially under the parent (VB ListItems are one-based).
    Set liChild = ListView1.ListItems.Add(iParentItem + 1 + i, , "item" & Format$(i, "00 "), , 1)
    
    ' Give the new child item the specified indent and a collapsed button.
    ' (the index of real listview items are zero-based).
    Call Listview_SetItemStateEx(m_hwndLV, liChild.Index - 1, iParentIndent + 1, lvisCollapsed)
    liChild.SubItems(1) = "subitem1"
    liChild.SubItems(2) = "subitem2"
    liChild.SubItems(3) = "subitem3"
  Next
  
  ' Resize the columns
  Call Form_Resize
  
  ' Make the last inserted subfolder visible, then the parent folder visible,
  ' per default treeview behavior. Post the messages to allow the ListView
  ' to finish pocessing any mouse events it may still be in...
  DoEvents
  Call PostMessage(m_hwndLV, LVM_ENSUREVISIBLE, iParentItem + nChildren, ByVal 0&)
  Call PostMessage(m_hwndLV, LVM_ENSUREVISIBLE, iParentItem, ByVal 0&)
  
  Screen.MousePointer = vbNormal

End Sub

' Collapses the specified parent item and removes all child items under it.

'   iParentItem     - real listview index of parent item (the zero-based position of the item within the
'                            ListView, as opposed to the item's one-based ListItems collection Index value)
'   iParentIndent  - parent item's indent value

Private Sub RemoveChildItems(iParentItem As Long, iParentIndent As Long)
  Dim nItems As Long
  Dim iChildIndent As Long
  
  Screen.MousePointer = vbHourglass
          
  ' The parent is currently expanded, collapse it, and remove all children
  ' with an indent value greater than the collapsing parent.
  Call Listview_SetItemStateEx(m_hwndLV, iParentItem, iParentIndent, lvisCollapsed)
  nItems = ListView1.ListItems.Count
  
  Do
    Call Listview_GetItemStateEx(m_hwndLV, iParentItem + 1, iChildIndent)
    If (iChildIndent > iParentIndent) Then
      
      ' Remove the item directly below the collapsing parent (VB ListItems are one-based)
      ListView1.ListItems.Remove (iParentItem + 2)
      
      ' Keep a count of ListView items so we don't try to remove more
      ' items than are in the ListView (when collapsing the last parent).
      nItems = nItems - 1
    End If
  Loop While (iChildIndent > iParentIndent) And (iParentItem + 1 < nItems)

  DoEvents
  Call Form_Resize
  Screen.MousePointer = vbNormal

End Sub

