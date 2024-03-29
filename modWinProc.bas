Attribute VB_Name = "modWinProc"
Option Explicit

'api constants
Public Const MF_BYCOMMAND = &H0&
Public Const MF_BYPOSITION = &H400&
Public Const WM_COMMAND = &H111
Public Const WM_MENUSELECT = &H11F
Public Const WM_SYSCOMMAND = &H112
Public Const GWL_WNDPROC = (-4)

'add new consts for new items
Public Const IDM_ITEM1 As Long = 0
Public Const IDM_ITEM2 As Long = 1
Public Const IDM_ABOUT As Long = 2
Public strImsg As String

Public Function WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, _
ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo handle_err
    ' ***WARNING***
    ' do not attempt to debug this procedure!!
    ' ***WARNING***
    ' My implementation of the message handling routine
    ' which determines which message was recieved and acts on the menu items
    Select Case iMsg
    Case WM_COMMAND
        If lParam = 0 Then
            frmMain.txtLowParam = CStr(LOWORD(lParam))
            frmMain.txtHighParam = CStr(wParam)
            Select Case wParam
            Case mnuFileAnmlDog
                MsgBox "Dog Days"
            Case mnuFileAnmlCat
                MsgBox "Canine Critters"
            Case mnuFileAnmlHamster
                MsgBox "Hasmters Forever"
            Case mnuFileAnmlChinchilla
                MsgBox "Chinchillas are cool!"
            Case mnuFileExit
                Unload frmMain
            Case mnuFilePrchCash
                MsgBox "Got any Green?"
            Case mnuFilePrchCredit
                MsgBox "Time for plastic surgery"
            Case mnuHelpAbout
                Load frmAbout
                frmAbout.Show
            End Select
        Else
            WindowProc = CallWindowProc(ProcOld, hwnd, iMsg, wParam, lParam)
        End If
    Case Else
        'Pass all messages on to VB and then return the value to windows
        WindowProc = CallWindowProc(ProcOld, hwnd, iMsg, wParam, lParam)
    End Select
    
handle_exit:
    Exit Function
    
handle_err:
'If you come back into debug this gives a chance to get control.
    If ProcOld <> 0 Then
        Call SetWindowLong(hwnd, GWL_WNDPROC, ProcOld)
    End If
    MsgBox Err.Description
    Resume handle_exit
    
End Function

Public Function LOWORD(ByVal dwValue As Long) As Integer
    LOWORD = dwValue Mod &H10000
End Function

Function HiWord(ByVal DWord As Long) As Integer
      HiWord = (DWord And &HFFFF0000) \ &H10000
End Function

Public Function myAddMenuItem(mnuType As Long, mnuId As Long, hMnu As Long, mnuPos As Long, mnuPosType As Long, mnuStr As String) As Long
'MenuItemInfo with SetMenuItemInfo() also gives the ability to add bitmaps to menu items and check items.

    Dim mii As MENUITEMINFO

    With mii
        ' The size of this structure.
        .cbSize = Len(mii)
        ' Which elements of the structure to use.
        .fMask = MIIM_ID Or MIIM_DATA Or MIIM_TYPE Or MIIM_SUBMENU
        ' The type of item: a string.
        .fType = mnuType
        ' This item is currently enabled and is the default item.
        .fState = MFS_ENABLED Or MFS_DEFAULT
        ' Assign this item an item identifier.
        .wID = mnuId
        ' Display the following text for the item.
        .dwTypeData = mnuStr
        .cch = Len(.dwTypeData)
    End With

    myAddMenuItem = InsertMenuItem(hMnu, mnuPos, mnuPosType, mii)
    
End Function


