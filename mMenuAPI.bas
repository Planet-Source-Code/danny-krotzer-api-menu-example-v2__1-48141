Attribute VB_Name = "mMenuAPI"
'__________________________________________________________________
'
' Author:   Danny K (danny@xi0n.com)
' Site:     http://www.xi0n.com
' Module:   ...for APIMenu example v2
'__________________________________________________________________

Option Explicit


' API declarations
'__________________________________________________________________

Public Declare Function IsWindow Lib "user32" ( _
        ByVal hwnd As Long) As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
        ByVal hwnd As Long, _
        ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long
        
Public Declare Function EnumWindows Lib "user32" ( _
        ByVal lpEnumFunc As Long, _
        ByVal lParam As Long) As Long
        
Public Declare Function GetMenuItemCount Lib "user32" ( _
        ByVal hMenu As Long) As Long
        
Public Declare Function GetSubMenu Lib "user32" ( _
        ByVal hMenu As Long, _
        ByVal nPos As Long) As Long
        
Public Declare Function GetMenu Lib "user32" ( _
        ByVal hwnd As Long) As Long

Public Declare Function SetMenu Lib "user32" ( _
        ByVal hwnd As Long, _
        ByVal hMenu As Long) As Long

Public Declare Function DestroyMenu Lib "user32" ( _
        ByVal hMenu As Long) As Long

Public Declare Function CreateMenu Lib "user32" () As Long

Public Declare Function CreatePopupMenu Lib "user32" () As Long

Public Declare Function DrawMenuBar Lib "user32" ( _
        ByVal hwnd As Long) As Long

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" ( _
        ByVal lpPrevWndFunc As Long, _
        ByVal hwnd As Long, _
        ByVal Msg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long) As Long

Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" ( _
        ByVal hwnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long) As Long

Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" ( _
        ByVal hwnd As Long) As Long

Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" ( _
        ByVal hwnd As Long, _
        ByVal lpString As String, _
        ByVal cch As Long) As Long
        
Public Declare Function IsMenu Lib "user32" ( _
        ByVal hMenu As Long) As Long

Public Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" ( _
        ByVal hMenu As Long, _
        ByVal un As Long, _
        ByVal b As Long, _
        lpMenuItemInfo As MENUITEMINFO) As Long

Public Declare Function InsertMenuItem Lib "user32" Alias "InsertMenuItemA" ( _
        ByVal hMenu As Long, _
        ByVal un As Long, _
        ByVal bool As Boolean, _
        ByRef lpcMenuItemInfo As MENUITEMINFO) As Long


' Types
'__________________________________________________________________

Public Type MENUITEMINFO
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


' Constants
'__________________________________________________________________

Public Const WM_COMMAND = &H111
Public Const GWL_WNDPROC = (-4)

Public Const MF_BYPOSITION = &H400&
Public Const MF_BYCOMMAND = &H0&
Public Const MF_POPUP = &H10&
Public Const MF_STRING = &H0&
Public Const MF_SEPARATOR = &H800&

Public Const MFS_ENABLED = &H0

Public Const MIIM_STATE = &H1
Public Const MIIM_ID = &H2
Public Const MIIM_SUBMENU = &H4
Public Const MIIM_TYPE = &H10


' Misc Variables
'__________________________________________________________________

Public lMenu As Long        'Holds handle to our dynamic menu
Public lWindowHwnd As Long  'Targeted Window Handle
Public plOldProc As Long    'Original WinProc Address (for subclassing)

'__________________________________________________________________
'
' Subclassing - This function processes our forms messages
'__________________________________________________________________

Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Select Case uMsg
    
    Case WM_COMMAND
            
        'make sure window is open yet
        If IsWindow(lWindowHwnd) <> 0 Then
            
            'wParam holds the menu item ID
            Call PostMessage(lWindowHwnd, WM_COMMAND, wParam, 0&)
        
        Else
        
            'if the window doesnt exist anymore, we better kill the menu
            'before we get ourselves into trouble...
            
            'remove menu
            frmMain.ResetMenu
            
            'reset process list
            frmMain.ListWindows
            
        End If

    Case Else
        
        ' otherwise let the original procedure handle it.
        WindowProc = CallWindowProc(plOldProc, hwnd, uMsg, wParam, lParam)
        
    End Select
    
End Function

'__________________________________________________________________
'
' Processes all returned window handles from EnumWindows()
'__________________________________________________________________

Public Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Boolean
    
 'caption length
 Dim lLength As Long
 lLength = GetWindowTextLength(hwnd)
    
 'make sure it has a menu
 Dim lMenuHwnd As Long
 lMenuHwnd = GetMenu(hwnd)
 
 'then find menu count
 Dim lMenuCount As Long
 lMenuCount = GetMenuItemCount(lMenuHwnd)
 
 'only list windows with a valid menu
 If lLength > 0 And lMenuHwnd > 0 And lMenuCount > 0 Then
    
    'get caption
    Dim strCaption As String
    strCaption = String$(lLength, vbNullChar)
    GetWindowText hwnd, strCaption, lLength + 1
 
    'don't list our own for...don't want to clone our own clone ;P
    If strCaption <> frmMain.Caption Then
        
        'add to list
        With frmMain.lstWindows
            .AddItem strCaption
            .ItemData(.NewIndex) = hwnd
        End With
    
    End If
    
 End If
 
 'all good
 EnumWindowsProc = True
 
End Function
