VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Run Menu API"
   ClientHeight    =   3780
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   8160
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   8160
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraCaption 
      Caption         =   "Target Window Caption"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   7695
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh List"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   2520
         Width           =   3495
      End
      Begin VB.ListBox lstWindows 
         Height          =   2010
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   7215
      End
      Begin VB.CommandButton cmdGrab 
         Caption         =   "Get Menu"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   2
         Top             =   2520
         Width           =   3375
      End
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run Selected"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   6480
      Width           =   2535
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'__________________________________________________________________
'
' Name:     API Menu Example v2
'
' Author:   Danny K. (danny@xi0n.com)
'           Violent (admin@sniip3r.com)
'
' Site:     http://www.xi0n.com
'           http://www.sniip3r.com
'
' Purpose:  How to use API to access & clone another apps menu, as well as
'           subclass your form to catch messages (for custom menu items)
'
' Notes:    This is much like the first version however this time I used
'           GetMenuItemInfo() and InsertMenuItem() which preserve the states
'           of the menu item (although you can change to your liking ie. setting
'           MFS_ENABLED in the .fstate to enable a disabled item).
'
'           However this version should iterate all submenus properly no matter how
'           many levels.
'
'           Also made it so you don't need to use VB's MenuEditor to create some sort
'           of dummy menu item to keep the menu persistent. This means completely
'           dynamic menus.
'
'           And as before, sometimes you'll find some menu-items don't work, such as
'           Notepad's Cut/Copy/Paste. That's usually because the app is
'           using built in system ID's.
'__________________________________________________________________
'
     
'___________________________________________________________________________
'
' Form - Subclass Window Messages to catch Menu Clicks
'___________________________________________________________________________

Private Sub Form_Load()
 
 'populate list with all open windows
 Call ListWindows
 
 'subclass - replace current WindProc address
 plOldProc = SetWindowLong(frmMain.hwnd, GWL_WNDPROC, AddressOf WindowProc)
 
End Sub

Private Sub Form_Unload(Cancel As Integer)

 'restore previous window procedure
 Call SetWindowLong(frmMain.hwnd, GWL_WNDPROC, plOldProc)
 
 'Destroy menu (if present) and free up memory
 If IsMenu(lMenu) = 1 Then DestroyMenu (lMenu)

End Sub


'___________________________________________________________________________
'
' Form - When the form refreshes, it releases the menu, so reset it
'___________________________________________________________________________

Private Sub Form_Paint()

'if we have created a menu, reset it
If IsMenu(lMenu) = 1 Then
    SetMenu Me.hwnd, lMenu
End If

End Sub


'___________________________________________________________________________
'
' Command Button - Get Window Handle from Caption
'___________________________________________________________________________

Private Sub cmdRefresh_Click()

 'refresh windows list
 Call ListWindows

End Sub

'___________________________________________________________________________
'
' Command Button - Get Window Handle from Caption
'___________________________________________________________________________

Private Sub cmdGrab_Click()
 
 'clear any menu items we might have loaded
 Call ResetMenu
 
 'try to find the window
 lWindowHwnd = lstWindows.ItemData(lstWindows.ListIndex)

 'if window found
 If IsWindow(lWindowHwnd) <> 0 Then
    
    'try to acquire menu
    Dim lMenuHwnd As Long
    lMenuHwnd = GetMenu(lWindowHwnd)

    'no menu found
    If lMenuHwnd <= 0 Then
        MsgBox "Selected window doesnt have a valid menu.", vbInformation
        Exit Sub
    End If

    GetMenuInfo lMenuHwnd, GetMenu(Me.hwnd)
    
 Else
    MsgBox "Selected window was not found. Try refreshing the list...", vbExclamation
 End If

End Sub

'___________________________________________________________________________
'
' Sub - Grab menu items and list in a ListBox w/MenuID stored in ItemData
'___________________________________________________________________________

Private Sub GetMenuInfo(lMenuHwnd As Long, lLocalMenu As Long)

 Dim uMII As MENUITEMINFO

 'set size and flags for what data we want to retrieve
 With uMII
    .cbSize = Len(uMII)
    .fMask = MIIM_ID Or MIIM_STATE Or MIIM_TYPE Or MIIM_SUBMENU
 End With

'make sure handles belong to menus
If IsMenu(lMenuHwnd) <> 1 Or IsMenu(lLocalMenu) <> 1 Then Exit Sub

'find menuitem count
Dim lMenuCount As Long
lMenuCount = GetMenuItemCount(lMenuHwnd)

'for each menu item
Dim x As Long
For x = 0 To lMenuCount - 1

    'get size of menu item text
    uMII.cch = 0
    Call GetMenuItemInfo(lMenuHwnd, x, 1&, uMII)
    
    'now allocate string for Caption
    uMII.dwTypeData = Space$(uMII.cch)
    uMII.cch = uMII.cch + 1
    
    'retrieve menu item data to our uMenuItemInfo
    Call GetMenuItemInfo(lMenuHwnd, x, 1&, uMII)
    
    'store the submenu handle because we send this struct ByRef
    'to our AddMenuItem() (which changes the .hSubMenu to our own
    'submenu handle which we create)
    Dim hSubMenu As Long
    hSubMenu = uMII.hSubMenu
    
    'add menuitem to our local menu
    Dim lNextMenu As Long
    lNextMenu = AddMenuItem(lLocalMenu, x, uMII)
    
    'if theres a submenu...loop through it as well
    If hSubMenu <> 0 Then GetMenuInfo hSubMenu, lNextMenu
    
Next 'menu items loop


End Sub


'___________________________________________________________________________
'
' Sub - Add Item to Forms Menu
'___________________________________________________________________________

Private Function AddMenuItem(hMenu As Long, lPosition As Long, uMII As MENUITEMINFO) As Long
   
 'if its a popup
 If uMII.hSubMenu <> 0 Then
    
    'create a popup
    Dim hSubMenu As Long
    hSubMenu = CreatePopupMenu()
 
    'save new popup handle to struct
    uMII.hSubMenu = hSubMenu
    uMII.wID = hSubMenu
    
    'now add new item
    InsertMenuItem hMenu, lPosition, True, uMII
    
    'return handle to new popup
    AddMenuItem = hSubMenu
    
 Else
    
    'if you wanted all menu items to be enabled regardless
    'uMII.fState = MFS_ENABLED 'use caution in doing so though
    
    'now add new item
    InsertMenuItem hMenu, lPosition, True, uMII
    
    'return current menu handle
    AddMenuItem = hMenu
    
 End If
 
 'redraw the updated menu
 DrawMenuBar Me.hwnd

End Function


'___________________________________________________________________________
'
' Sub - Destroy Current Menu and Reload
'___________________________________________________________________________

Public Sub ResetMenu()

 'grab our menu handle
 Dim hMenu As Long
 hMenu = GetMenu(Me.hwnd)

 'if loaded, destroy and free up memory
 If hMenu <> 0 Then DestroyMenu (hMenu)

 'reinitialize a blank menu
 hMenu = CreateMenu()
 Call SetMenu(Me.hwnd, hMenu)
 
 'save new menu handle
 lMenu = hMenu
 
End Sub

'___________________________________________________________________________
'
' Sub - Load all current windows into List
'___________________________________________________________________________

Public Sub ListWindows()

 'clear list
 lstWindows.Clear

 'ask for all open window handles
 Call EnumWindows(AddressOf EnumWindowsProc, 0&)
 
End Sub





