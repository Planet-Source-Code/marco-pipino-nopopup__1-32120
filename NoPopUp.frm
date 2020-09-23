VERSION 5.00
Begin VB.Form NoPopUp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "No more Pop Up"
   ClientHeight    =   7110
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   8250
   Icon            =   "NoPopUp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_Enable 
      Caption         =   "Disable"
      Height          =   375
      Left            =   4380
      TabIndex        =   10
      Top             =   6480
      Width           =   1035
   End
   Begin VB.CommandButton cmd_Quit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   6960
      TabIndex        =   9
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmd_Hide 
      Caption         =   "Hide"
      Height          =   375
      Left            =   5640
      TabIndex        =   8
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1755
      Left            =   60
      TabIndex        =   4
      Top             =   5100
      Width           =   4035
      Begin VB.CheckBox chkHideOnStart 
         Caption         =   "Do not show  this form when program start"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1260
         Width           =   3375
      End
      Begin VB.CheckBox chkStartUp 
         Caption         =   "Run at StartUp"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   2235
      End
      Begin VB.CheckBox chkBeep 
         Caption         =   "Beep on Block"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   420
         Width           =   1515
      End
   End
   Begin VB.ListBox lstWhiteList 
      Height          =   2010
      Left            =   60
      TabIndex        =   1
      ToolTipText     =   "Double click to Remove from White List"
      Top             =   2940
      Width           =   8055
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3900
      Top             =   0
   End
   Begin VB.ListBox lstBlockedUrls 
      Height          =   2010
      Left            =   60
      TabIndex        =   0
      ToolTipText     =   "Double-click to Add in White List"
      Top             =   420
      Width           =   8055
   End
   Begin VB.Label Label2 
      Caption         =   "White List"
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   2580
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Blocked PopUp"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   120
      Width           =   1305
   End
   Begin VB.Menu mnuTrayIconPopup 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "Show Options"
      End
      Begin VB.Menu mnuEnable 
         Caption         =   "Disable"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "NoPopUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'NO POP UP is an application that remove the tedious popup when surfing on the web.
'
'It use the shydocvw.dll (Microsoft Internet Control) to enumerate and grab the Internet Explorer Window.
'Then control when then parent window is an object.
'
'Also this application is a demo of how
'- to manage the Sys tray with no dll or ocx
'- to set that un aplication start at window start-up (and then how to manage the registry)
'- to read and write on a text file

Option Explicit

'Set the reference to
' - Microsoft Internet Control (shdocvw.dll)
'       to enumerate the then Shell Widows
'       and managing the Internet Explorer windows
'
' - Microsoft HTML Object library (mshtml.dll)
'       to recognize the HTMLDocument object

Dim SWs As New SHDocVw.ShellWindows
Dim IE As SHDocVw.InternetExplorer

Dim blnEnabled As Boolean
Dim blnStarting As Boolean

Private Sub cmd_Enable_Click()
    ChangeState
End Sub

Private Sub cmd_Hide_Click()
    Me.Hide
End Sub

Private Sub cmd_Quit_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    Dim nid As NOTIFYICONDATA
    With nid
        .cbSize = Len(nid)
        .hWnd = NoPopUp.hWnd
        .uID = 0
        .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
        .uCallbackMessage = PK_TRAYICON
        .hIcon = NoPopUp.Icon
        .szTip = "No PopUp Enabled" & vbNullChar
    End With

    ' Shell_NotifyIconA ID_OF_ICON, NOTIFYICONDATA
    Shell_NotifyIconA NIM_ADD, nid
    
    ' poldproc is the address(memory location) of the original window procedure
    pOldProc = SetWindowLongA(Me.hWnd, GWL_WNDPROC, AddressOf WindowProc)

    'Loading the setting
    chkBeep.Value = CInt(GetSetting(App.Title, "Options", "Beep", "0"))
    chkStartUp.Value = CInt(GetSetting(App.Title, "Options", "StartUp", "0"))
    chkHideOnStart.Value = CInt(GetSetting(App.Title, "Options", "HideOnStart", "0"))
    'Read the WhiteList file
    ReadWhiteList
    
    blnEnabled = True
    blnStarting = True
End Sub

Private Sub Form_Paint()
    If blnStarting Then
        If CInt(GetSetting(App.Title, "Options", "HideOnStart", "0")) = "1" Then
            Me.Hide
            blnStarting = False
        End If
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Me.Hide
        Cancel = 1
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim nid As NOTIFYICONDATA
    
    SaveWhiteList
    SaveSetting App.Title, "Options", "Beep", CStr(chkBeep.Value)
    SaveSetting App.Title, "Options", "StartUp", CStr(chkStartUp.Value)
    SaveSetting App.Title, "Options", "HideOnStart", CStr(chkHideOnStart.Value)
    SetStartp chkStartUp.Value
    
    With nid
        .hWnd = Me.hWnd
        .cbSize = Len(nid)
        .uID = 0
    End With
    
    Shell_NotifyIconA NIM_DELETE, nid
    SetWindowLongA Me.hWnd, -4, pOldProc
End Sub

'Add blocked Url to White List
Private Sub lstBlockedUrls_DblClick()
    If lstBlockedUrls.ListCount > 0 Then
        If MsgBox("Add " & lstBlockedUrls.List(lstBlockedUrls.ListIndex) & " to White List?", vbYesNo) = vbYes Then
            If Not isInList(lstWhiteList, lstBlockedUrls.List(lstBlockedUrls.ListIndex)) Then
                lstWhiteList.AddItem lstBlockedUrls.List(lstBlockedUrls.ListIndex)
            End If
        End If
    End If
End Sub

Private Sub lstWhiteList_DblClick()
    If lstWhiteList.ListCount > 0 Then
        If MsgBox("Remove " & lstWhiteList.List(lstWhiteList.ListIndex) & " from White List?", vbYesNo) = vbYes Then
            lstWhiteList.RemoveItem lstWhiteList.ListIndex
        End If
    End If
End Sub

Private Sub mnuAbout_Click()
    MsgBox ("NO MORE POP UP!! " & vbCrLf & _
        "ver. 1.0 20/02/2002" & vbCrLf _
        & "by Marco Pipino (marcopipino@libero.it)")
End Sub

Private Sub mnuEnable_Click()
    ChangeState
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuShow_Click()
    Me.Show
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    Dim Doc
    If Enabled Then
    'THE CORE OF APPLICAITION
        For Each IE In SWs
        'SWs enumerate the Shell Windows
            Set Doc = IE.Document
            If TypeOf Doc Is HTMLDocument Then
                'if The type Of Doc is an HTML Document (Internet Explorer)
                'then control the opener of the parent windows
                'and the Url is not in White List then ... CLOSE!!!
                If IsObject(Doc.parentWindow.opener) And Not isInList(lstWhiteList, IE.LocationURL) Then
                    If Not isInList(lstBlockedUrls, IE.LocationURL) Then lstBlockedUrls.AddItem IE.LocationURL
                    If chkBeep.Value = 1 Then Beep 800, 200
                    IE.Quit
                End If
            End If
        Next
    End If
End Sub

'Serching into the list for the string
Private Function isInList(lstObj As ListBox, strUrl As String)
    Dim i As Integer
    isInList = False
    For i = 0 To lstObj.ListCount
        If strUrl = lstObj.List(i) Then
            isInList = True
            Exit For
        End If
    Next
End Function

'Read the WhiteList text file and insert into the list
Private Function ReadWhiteList()
    Dim intFileNumber As Integer
    Dim strTempUrl As String
    
    On Error GoTo FileError
    intFileNumber = FreeFile
    
    Open App.Path & "\" & "Whitelst.txt" For Input As #intFileNumber
    Do While Not EOF(1)
        Input #intFileNumber, strTempUrl
        If Not isInList(lstWhiteList, strTempUrl) Then
            lstWhiteList.AddItem strTempUrl
        End If
    Loop
    Close #intFileNumber
    Exit Function
    
FileError:
'At first time the file is not found, and then create it
    If Err.Number = ERR_FILE_NOT_FOUND Then
        Open App.Path & "\" & "Whitelst.txt" For Output As #intFileNumber
        Close #intFileNumber
    End If
End Function

'Save the list of WhiteList Yrl to the Whitelist text file
Private Function SaveWhiteList()
    Dim intFileNumber As Integer
    Dim strTempUrl As String
    Dim i As Integer
    
    intFileNumber = FreeFile
    
    Open App.Path & "\" & "Whitelst.txt" For Output As #intFileNumber
    For i = 0 To lstWhiteList.ListCount - 1
        Print #intFileNumber, lstWhiteList.List(i)
    Next
    Close #intFileNumber
End Function

'Change the state of application
'Changing icon, menu and button caption
Private Function ChangeState()
    If blnEnabled Then
        blnEnabled = False
        cmd_Enable.Caption = "Enable"
        mnuEnable.Caption = "Enable"
        Set Me.Icon = LoadPicture(App.Path & "\disabled.ico")
    Else
        blnEnabled = True
        cmd_Enable.Caption = "Disable"
        mnuEnable.Caption = "Disable"
        Set Me.Icon = LoadPicture(App.Path & "\enabled.ico")
    End If
    UpdateIcon
End Function

'Update the IconTray when enable or disable the application
Private Sub UpdateIcon()
    Dim nid As NOTIFYICONDATA

    With nid
        .cbSize = Len(nid)
        .hWnd = NoPopUp.hWnd
        .uID = 0
        .uFlags = NIM_DELETE Or NIM_MODIFY Or NIF_TIP
        .uCallbackMessage = PK_TRAYICON
        .hIcon = NoPopUp.Icon
        .szTip = "No PopUp " & IIf(blnEnabled, "Enabled", "Diabled") & vbNullChar
    End With
    Shell_NotifyIconA NIM_MODIFY, nid

End Sub
