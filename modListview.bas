Attribute VB_Name = "modListview"
Option Explicit

Private Const LVM_FIRST = &H1000
Private Const LVM_GETNEXTITEM = (LVM_FIRST + 12)
Private Const LVNI_SELECTED = &H2
Private Const LVM_GETSELECTEDCOUNT = (LVM_FIRST + 50)
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55
Private Const LVS_EX_FULLROWSELECT = &H20
Private Const LVIF_TEXT = &H1
Private Const LVM_SETITEMSTATE As Long = (LVM_FIRST + 43)
Private Const LVM_GETITEMSTATE As Long = (LVM_FIRST + 44)
Private Const LVM_GETITEMTEXT As Long = (LVM_FIRST + 45)

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Public Function GetSelectedItemsFromListview(oListview As ListView) As Collection
Dim lCurSelectedItemIndex As Long
Dim myCol As Collection
Dim myListItem As ListItem
Dim i

    'begin/start position in the listview
    lCurSelectedItemIndex = -1
    'create collection to hold selected items
    Set myCol = New Collection
    
    For i = 1 To CountSelectedItemsInListview(oListview)
        'get the itemx index from the selected (current)item
        lCurSelectedItemIndex = SendMessage(oListview.hwnd, LVM_GETNEXTITEM, lCurSelectedItemIndex, ByVal LVNI_SELECTED)
        'add the listitem to the collection
        myCol.Add oListview.ListItems.Item(lCurSelectedItemIndex + 1)
    Next i
    
    'return the collection
    Set GetSelectedItemsFromListview = myCol
    
End Function
Public Function CountSelectedItemsInListview(oListview As ListView) As Long
    
    'count the selected items
    CountSelectedItemsInListview = SendMessage(oListview.hwnd, LVM_GETSELECTEDCOUNT, 0, 0)

End Function

