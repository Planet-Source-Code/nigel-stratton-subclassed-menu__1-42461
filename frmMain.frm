VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Campaign Manager"
   ClientHeight    =   1680
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtHighParam 
      Height          =   315
      Left            =   3840
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   900
      Width           =   1275
   End
   Begin VB.TextBox txtLowParam 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   900
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "CallWindowProc Parameters"
      Height          =   255
      Left            =   540
      TabIndex        =   4
      Top             =   540
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "High Param"
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Low Param"
      Height          =   255
      Left            =   540
      TabIndex        =   2
      Top             =   960
      Width           =   795
   End
   Begin VB.Menu mnuX 
      Caption         =   "X"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hMenu As Long
Dim WithEvents myMnu As Menu

Private Sub Form_Load()
    Dim hMenuWork As Long
    Dim hMenuSub As Long
    Dim hMenuItems As Long
    Dim result As Long
    
    'I had CreateMenu but menu disappears after first item selected
    'hMenu = CreateMenu
    'This works but menu still grays while msgbox or form open
    hMenu = GetMenu(Me.hwnd)
    'Indented to give an idea of menu heirarchy
    hMenuSub = CreatePopupMenu
    result = AppendMenu(hMenu, MF_POPUP, hMenuSub, "File")
        hMenuWork = CreatePopupMenu
        result = AppendMenu(hMenuSub, MF_POPUP, hMenuWork, "Animals")
            hMenuItems = CreatePopupMenu
            result = myAddMenuItem(MF_STRING, mnuFileAnmlDog, hMenuWork, 0, 1, "Dogs")
            result = myAddMenuItem(MF_STRING, mnuFileAnmlCat, hMenuWork, 1, 1, "Cats")
            result = myAddMenuItem(MF_STRING, mnuFileAnmlHamster, hMenuWork, 2, 1, "Hamsters")
            result = myAddMenuItem(MF_STRING, mnuFileAnmlChinchilla, hMenuWork, 3, 1, "Chinchillas")
        result = myAddMenuItem(MF_STRING, mnuFileExit, hMenuSub, 2, 1, "Exit")
        hMenuSub = CreatePopupMenu
        result = AppendMenu(hMenu, MF_POPUP, hMenuSub, "Purchase")
            result = myAddMenuItem(MF_STRING, mnuFilePrchCash, hMenuSub, 1, 1, "Cash...")
            result = myAddMenuItem(MF_STRING, mnuFilePrchCredit, hMenuSub, 1, 1, "Credit Card")
        hMenuSub = CreatePopupMenu
        result = AppendMenu(hMenu, MF_POPUP Or MF_HELP, hMenuSub, "Help")
            result = myAddMenuItem(MF_STRING, mnuHelpAbout, hMenuSub, 2, 1, "About...")
    
    'Get rid of the existing menu placeholder
    Me.mnuX.Visible = False
    
    'I had this for CreateMenu but menu disappears after
    'first item selected???
    
    'Show the form, without this the createMenu will not attach
    'Me.Show
    'Add the menu bar to the form
    'result = SetMenu(Me.hwnd, hMenu)

    ProcOld = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Destroy our menu
    DestroyMenu hMenu
    If ProcOld <> 0 Then
        Call SetWindowLong(hwnd, GWL_WNDPROC, ProcOld)
    End If
End Sub


