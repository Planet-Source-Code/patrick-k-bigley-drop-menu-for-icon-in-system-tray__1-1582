<div align="center">

## Drop Menu for icon in System Tray


</div>

### Description

Easy Code here: Place an icon in the System Tray and have a drop menu appear when you click the icon with the right-mouse button. If you already have code with bitmaps in your menu, just add this code to your project! This code DOES WORK, just be careful and follow the tips. Then slap yourself for other people's long and drawn out useless code from the past.
 
### More Info
 
' Created by opus@bargainbd.com

' Original source is unknown

' Before you begin!

' Make sure your form is in view within Visual Basic,

' then press Ctrl+E to open the Menu Editor.

' Next create a Main Menu item and make it's name

' property "mnu_1", without the quotes. You can

' always change this name, but make sure that you

' change it in the Form_MouseMove too. Now create a

' few sub menus under the main menu

' and name them anything that you want,

' the code will take care of the rest.

' "TIP: Make the "mnu_1" visible property = False

' Then create a second Main menu item with sub menus

' as normal (This will appear to look as though

' it is the first menu item. The Actual First

' will be seen in the System tray when clicked with

' the right mouse button.

NO SIDE EFFECTS, but make sure your first menu item's name is consistant through out your entire project. (Just remember to refer to the proper main menu item,)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Patrick K\. Bigley](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/patrick-k-bigley.md)
**Level**          |Unknown
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/patrick-k-bigley-drop-menu-for-icon-in-system-tray__1-1582/archive/master.zip)

### API Declarations

```

'Uuser defined type required by Shell_NotifyIcon API call
Public Type NOTIFYICONDATA
  cbSize As Long
  hwnd As Long
  uId As Long
  uFlags As Long
  uCallBackMessage As Long
  hIcon As Long
  szTip As String * 64
End Type
'constants required by Shell_NotifyIcon API call:
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201   'Button down
Public Const WM_LBUTTONUP = &H202    'Button up
Public Const WM_LBUTTONDBLCLK = &H203  'Double-click
Public Const WM_RBUTTONDOWN = &H204   'Button down
Public Const WM_RBUTTONUP = &H205    'Button up
Public Const WM_RBUTTONDBLCLK = &H206  'Double-click
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public nid As NOTIFYICONDATA
```


### Source Code

```

Private Sub Form_Load()
' Project Topic:
' "Add Menu to System Tray Icon"
' For VB5.0 and better....
' Created by opus@bargainbd.com
' Original source is unknown
' Before you begin!
' Make sure your form is in view within Visual Basic,
' then press Ctrl+E to open the Menu Editor.
' Next create a Main Menu item and make it's name
' property "mnu_1", without the quotes. You can
' always change this name, but make sure that you
' change it in the Form_MouseMove too. Now create a
' few sub menus under the main menu
' and name them anything that you want,
' the code will take care of the rest.
' "TIP: Make the "mnu_1" visible property = False
' Then create a second Main menu item with sub menus
' as normal (This will appear to look as though
' it is the first menu item. The Actual First
' will be seen in the System tray when clicked with
' the right mouse button.
' *---The code begins here---*
'The form must be fully visible before calling Shell_NotifyIcon
Me.Show
Me.Refresh
With nid
    .cbSize = Len(nid)
    .hwnd = Me.hwnd
    .uId = vbNull
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .uCallBackMessage = WM_MOUSEMOVE
    .hIcon = Me.Icon
    .szTip = " Click Right Mouse Button " & vbNullChar
End With
Shell_NotifyIcon NIM_ADD, nid
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'This procedure receives the callbacks from the System Tray icon.
Dim Result As Long
Dim msg As Long
'The value of X will vary depending upon the scalemode setting
If Me.ScaleMode = vbPixels Then
 msg = X
Else
 msg = X / Screen.TwipsPerPixelX
End If
  Select Case msg
    Case WM_LBUTTONUP    '514 restore form window
     Me.WindowState = vbNormal
     Result = SetForegroundWindow(Me.hwnd)
     Me.Show
    Case WM_LBUTTONDBLCLK  '515 restore form window
     Me.WindowState = vbNormal
     Result = SetForegroundWindow(Me.hwnd)
     Me.Show
    Case WM_RBUTTONUP    '517 display popup menu
     Result = SetForegroundWindow(Me.hwnd)
'***** STOP! and make sure that your first menu item
' is named "mnu_1", otherwise you will get an erro below!!! *******
     Me.PopupMenu Me.mnu_1
  End Select
End Sub
Private Sub Form_Resize()
    'this is necessary to assure that the minimized window is hidden
    If Me.WindowState = vbMinimized Then Me.Hide
End Sub
Private Sub Form_Unload(Cancel As Integer)
    'this removes the icon from the system tray
    Shell_NotifyIcon NIM_DELETE, nid
End Sub
Private Sub mPopExit_Click()
    'called when user clicks the popup menu Exit command
    Unload Me
End Sub
Private Sub mPopRestore_Click()
    'called when the user clicks the popup menu Restore command
    Me.WindowState = vbNormal
    Result = SetForegroundWindow(Me.hwnd)
    Me.Show
End Sub
```

