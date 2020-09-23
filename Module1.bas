Attribute VB_Name = "Module1"
Option Explicit

'Declaration for the Beep function when an Url is Blocked
Declare Function Beep Lib "kernel32.dll" (ByVal dwFreq As Long, _
    ByVal dwDuration As Long) As Long

'Declaration for managing the registry keys used for
'setting when the program start when Windows start
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" _
    Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As _
    String, ByVal ulOptions As Long, ByVal samDesired As Long, _
    phkResult As Long) As Long

Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal _
    hKey As Long) As Long

Public Declare Function RegSetValueEx Lib "advapi32.dll" _
    Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName _
    As String, ByVal Reserved As Long, ByVal dwType As Long, _
    lpData As Any, ByVal cbData As Long) As Long


Public Declare Function RegDeleteValue Lib "advapi32.dll" _
    Alias "RegDeleteValueA" (ByVal hKey As Long, _
    ByVal lpValueName As String) As Long

'Declaration used for managing the SysTray Icon
Public Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, _
    ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    
Public Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc _
    As Long, ByVal hWnd As Long, ByVal Msg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
    
Public Declare Function Shell_NotifyIconA Lib "shell32.dll" _
    (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

'This structure structure stores information used to
'communicate with an icon in the system tray.
Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

'Constants for managing the SysTray
Public Const NIM_ADD = &H0      ' add an icon to the system tray
Public Const NIM_MODIFY = &H1   ' modify an icon in the system tray
Public Const NIM_DELETE = &H2   ' delete an icon in the system tray
Public Const NIF_MESSAGE = &H1  ' whether a message is sent to the window procedure for events
Public Const NIF_ICON = &H2     ' whether an icon is displayed
Public Const NIF_TIP = &H4      ' tooltip availibility


'The registry key for the current User Block
Public Const HKEY_CURRENT_USER = &H80000001
Public Const KEY_WRITE = &H20006
Public Const REG_SZ = 1


Public Const GWL_WNDPROC = -4

' pointer to the previous window function using the callback WindowsProc
Public pOldProc As Long

'Constants for parameters of messages to SysTray
Public Const WM_RBUTTONUP = &H205 'Right-click
Public Const WM_LBUTTONDBLCLK = &H203 'Left-double-click
'Default Message to TrayIcon
Public Const PK_TRAYICON = &H578

'Error opening first time the file for WhiteList
Public Const ERR_FILE_NOT_FOUND = 53

'This Sub Set or UnSet that the program start when
'User is logged
Public Sub SetStartp(blnSet As Boolean)
    Dim lngHregkey As Long
    Dim strSubkey As String
    Dim strBuffer As String
    Dim lngRetval As Long
    
    strSubkey = "Software\Microsoft\Windows\CurrentVersion\Run"
    
    'Open the Registry key
    lngRetval = RegOpenKeyEx(HKEY_CURRENT_USER, strSubkey, 0, _
        KEY_WRITE, lngHregkey)
        
    If lngRetval <> 0 Then
        Exit Sub
    End If
    
    If blnSet Then
        strBuffer = App.Path & "\" & App.EXEName & ".exe" & vbNullChar
        'Creating the subkey
        lngRetval = RegSetValueEx(lngHregkey, App.Title, 0, REG_SZ, ByVal strBuffer, Len(strBuffer))
    Else
        'Deleting the subkey
        lngRetval = RegDeleteValue(lngHregkey, App.Title)
    End If
    'Closing then key
    RegCloseKey lngHregkey
End Sub



Public Function WindowProc(ByVal hWnd As Long, ByVal Msg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
    
    'Call-Back function
    'If the message is the Tray Icon and the lParam is either the
    'Right-Click or Left-Double-Click ...
    If Msg = PK_TRAYICON And lParam = WM_RBUTTONUP Then NoPopUp.PopupMenu NoPopUp.mnuTrayIconPopup, , , , NoPopUp.mnuShow
    If Msg = PK_TRAYICON And lParam = WM_LBUTTONDBLCLK Then NoPopUp.Show
    
    
    WindowProc = CallWindowProcA(pOldProc, hWnd, Msg, wParam, lParam)
End Function

