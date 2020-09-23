Attribute VB_Name = "modSortListView"
'******************************************Procedures in modSortListView******************************************
'*                                                                                                                *
'*Private Sub SetIconToColumnHeader(ListViewX As ListView, ByVal ColumnHeader As MSComctlLib.ColumnHeader)        *
'*Public Sub SetListViewOrder(ListViewX As ListView, ColumnHeader As MSComctlLib.ColumnHeader)                    *
'*Private Function GetItemData(ByVal lngParam As Long, ByVal hWnd As Long, IndexColumn As Long) As Variant        *
'*Private Function CompareNumeric(ByVal lngParam1 As Long, ByVal lngParam2 As Long, ByVal hWnd As Long) As Long   *
'*Private Function CompareDates(ByVal lngParam1 As Long, ByVal lngParam2 As Long, ByVal hWnd As Long) As Long     *
'*Public Property Get SortOrder() As ListSortOrderConstants                                                       *
'*Public Property Let SortOrder(NewValue As ListSortOrderConstants)                                               *
'*Public Property Get IndexHeader() As Integer                                                                    *
'*Public Property Let IndexHeader(NewValue As Integer)                                                            *
'*****************************************************************************************************************


Option Explicit

'Structures

Private Type POINT
    x As Long
    y As Long
End Type

Private Type LV_FINDINFO
    Flags As Long
    psz As String
    lParam As Long
    pt As POINT
    vkDirection As Long
End Type

Private Type LV_ITEM
    mask As Long
    iItem As Long
    iSubItem As Long
    State As Long
    stateMask As Long
    pszText As Long
    cchTextMax As Long
    iImage As Long
    lParam As Long
    iIndent As Long
End Type

'Constants
Private Const LVFI_PARAM = 1
Private Const LVIF_TEXT = &H1

Private Const LVM_FIRST = &H1000
Private Const LVM_FINDITEM = LVM_FIRST + 13
Private Const LVM_GETITEMTEXT = LVM_FIRST + 45
Private Const LVM_SORTITEMS = LVM_FIRST + 48

'API declarations

Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" ( _
        ByVal hWnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long) As Long

Private m_SortOrder As ListSortOrderConstants

Private mintIndexHeader As Integer
Public Property Get IndexHeader() As Integer
    IndexHeader = mintIndexHeader
End Property
Public Property Let IndexHeader(NewValue As Integer)
    mintIndexHeader = NewValue
End Property
Public Property Get SortOrder() As ListSortOrderConstants
    SortOrder = m_SortOrder
End Property
Public Property Let SortOrder(NewValue As ListSortOrderConstants)
    m_SortOrder = NewValue
End Property
Private Function CompareDates(ByVal lngParam1 As Long, ByVal lngParam2 As Long, ByVal hWnd As Long) As Long
'**************************************************
'   Purpose:
'       Compare two items of listview of type dates
'   Assertments:
'       the column must be of dates !
'   Inputs :
'   ByVal lngParam1 : Long
'
'   ByVal lngParam2 : Long
'
'   ByVal hWnd : Long
'
'

'   Returns :
'   Long
'
'
'
'   Usage :
'
'
'
'

'**************************************************


    Dim strName1 As String
    Dim strName2 As String
    Dim dtFirst As Date
    Dim dtSecond As Date

    'Obtain the item names and dates corresponding to the
    'input parameters

    dtFirst = GetItemData(lngParam1, hWnd, IndexHeader)
    dtSecond = GetItemData(lngParam2, hWnd, IndexHeader)

    'Compare the dates
    'Return 0 ==> Less Than
    '       1 ==> Equal
    '       2 ==> Greater Than
    Select Case True
        Case dtFirst < dtSecond
            CompareDates = 0
        Case dtFirst = dtSecond
            CompareDates = 1
        Case dtFirst > dtSecond
            CompareDates = 2
    End Select
    If SortOrder = lvwDescending Then CompareDates = 2 - CompareDates
End Function

Private Function CompareNumeric(ByVal lngParam1 As Long, ByVal lngParam2 As Long, ByVal hWnd As Long) As Long
    
    Dim lngFirst As Long
    Dim lngSecond As Long

    'Obtain the item names and dates corresponding to the
    'input parameters
    lngFirst = GetItemData(lngParam1, hWnd, IndexHeader)
    lngSecond = GetItemData(lngParam2, hWnd, IndexHeader)

    'Compare
    'Return 0 ==> Less Than
    '       1 ==> Equal
    '       2 ==> Greater Than
    Select Case True
        Case lngFirst < lngSecond
            CompareNumeric = 0
        Case lngFirst = lngSecond
            CompareNumeric = 1
        Case lngFirst > lngSecond
            CompareNumeric = 2
    End Select
    If SortOrder = lvwDescending Then CompareNumeric = 2 - CompareNumeric
End Function


Private Function GetItemData(ByVal lngParam As Long, ByVal hWnd As Long, IndexColumn As Long) As Variant
            
    Dim objFind As LV_FINDINFO
    Dim lngIndex As Long
    Dim objItem As LV_ITEM
    Dim baBuffer(32) As Byte
    Dim lngLength As Long

    '
    ' Convert the input parameter to an index in the list view
    '
    objFind.Flags = LVFI_PARAM
    objFind.lParam = lngParam
    lngIndex = SendMessage(hWnd, LVM_FINDITEM, -1, VarPtr(objFind))

    '
    ' Obtain the value of the specified list view item
    '
    objItem.mask = LVIF_TEXT
    objItem.iSubItem = IndexColumn - 1
    objItem.pszText = VarPtr(baBuffer(0))
    objItem.cchTextMax = UBound(baBuffer)
    lngLength = SendMessage(hWnd, LVM_GETITEMTEXT, lngIndex, _
            VarPtr(objItem))
    If lngLength > 0 Then
        GetItemData = Left$(StrConv(baBuffer, vbUnicode), lngLength)
    End If

End Function


Public Sub SetListViewOrder(ListViewX As ListView, ColumnHeader As MSComctlLib.ColumnHeader)
    '**************************************************
    '   Author: Ignat Andrei
    '   Comments :  set order to listview
    '   Purpose:
    '       when clicking the columnheader on listview - call this function
    '   Assertments:
    '       Exist Associated ImageView - with min 3 images "up","down","no"
    '   Inputs :
    '   ListViewX : ListView
    '
    '   ColumnHeader : MSComctlLib.ColumnHeader
    '
    '
    '
    '
    '   Usage :
    '
    '
    '
    '   RTFM
    '**************************************************
    On Error GoTo SetListViewOrder_ErrorHandler
    
    If ListViewX.SortKey = ColumnHeader.Index - 1 Then
        ' if user clicked twice same header ...
        ' invert sort order
        Select Case ListViewX.SortOrder
            Case lvwAscending
                ListViewX.SortOrder = lvwDescending
            Case lvwDescending
                ListViewX.SortOrder = lvwAscending
        End Select
    Else ' other column - ser sort order to ascending
        ListViewX.SortOrder = lvwAscending
    End If
    ' setting icons
    SetIconToColumnHeader ListViewX, ColumnHeader
    SortOrder = ListViewX.SortOrder
    IndexHeader = ColumnHeader.Index
    ListViewX.SortKey = ColumnHeader.Index - 1
    Dim lngItem As Long, strName As String
    Select Case LCase(ColumnHeader.Tag)
        Case "numeric"
            ' the column must be sorted numeric
            ListViewX.Sorted = False
            SendMessage ListViewX.hWnd, _
                    LVM_SORTITEMS, _
                    ListViewX.hWnd, _
                    AddressOf CompareNumeric
            
            'ListViewX.Refresh
            Exit Sub
        Case "date"
            ' the column must be sorted on dates
            ListViewX.Sorted = False
            SendMessage ListViewX.hWnd, _
                    LVM_SORTITEMS, _
                    ListViewX.hWnd, _
                    AddressOf CompareDates
            
            'ListViewX.Refresh
            Exit Sub

        Case Else
            ' default sorting method of listview
            ListViewX.Sorted = True
    End Select

    Exit Sub
SetListViewOrder_ErrorHandler:

End Sub


Private Sub SetIconToColumnHeader(ListViewX As ListView, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    On Error Resume Next '
    'do not raise any error if the icons key
    ' up, down and no does not exist in associated image list
    Dim clhLoop As ColumnHeader

    For Each clhLoop In ListViewX.ColumnHeaders
        clhLoop.Icon = "no"
    Next
    ColumnHeader.Icon = IIf(ListViewX.SortOrder = lvwAscending, "up", "down")
    If Err.Number <> 0 Then Err.Clear
End Sub
