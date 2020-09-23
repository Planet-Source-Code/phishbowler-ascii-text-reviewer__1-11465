Attribute VB_Name = "GeneralAPI"
'This Code Was Written By: Phishbowler
'Sept. 14, 2000
'
'For Money Making Opportunities, Visit
'Http://www.dreamstruct.com/
'
'Napster Users: Tired of Incomplete Songs?
'Get the good ol' Nap v2.0 Only available at:
'Http://come.to/NapsterResume
'
'The color form fade was
'written by the same author
'who wrote cryofade.bas
'for AIM
'
'This BAS Constructed by: Phishbowler
'

Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long)
Declare Function mciSendString Lib "MMSystem" Alias "mcisendstring" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal wReturnLength As Integer, ByVal hCallback As Integer) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal cmd As Long) As Long
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal Length As Long)

Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long


Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function iswindowvisible Lib "user32" Alias "IsWindowVisible" (ByVal hwnd As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long

Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)

Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SB_PAGEDOWN = 3
Public Const SB_LINEDOWN = 1
Public Const VK_SCROLL = &H91

Const SND_SYNC = &H0
    Public Const SND_ASYNC = &H1
    Public Const SND_NODEFAULT = &H2
    Public Const SND_MEMORY = &H4
    Public Const SND_LOOP = &H8
    Public Const SND_NOSTOP = &H10
Public Const WM_CLOSE = &H10
Public Const WM_SETTEXT = &HC
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDOWN = &H201
Public Const SW_HIDE = 0
Public Const SW_SHOW = 5
Public Const SW_RESTORE = 9
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const GW_HWNDNEXT = 2
Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1

Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1

Public Const LB_GETCOUNT = &H18B
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185



Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT




Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_MENU = &H12
Public Const VK_RETURN = &HD
Public Const VK_RIGHT = &H27
Public Const VK_SHIFT = &H10
Public Const VK_SPACE = &H20
Public Const VK_UP = &H26

Public Const WM_CHAR = &H102

Public Const WM_COMMAND = &H111
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDBLCLK = &H203


Public Const WM_MOVE = &HF012

Public Const WM_SYSCOMMAND = &H112
Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1

Public Const GW_HWNDPREV = 3
Public Const GW_MAX = 5
Public Const GW_OWNER = 4

Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000

Public Const ENTER_KEY = 13
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Type POINTAPI
        X As Long
        Y As Long
End Type
'Form back color fade codes begin here
'Works best when used in the Form_Paint() sub


Sub FormFadeBlue(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 0, 255 - intLoop), B
    Next intLoop
End Sub

Sub FormFadeGreen(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 255 - intLoop, 0), B
    Next intLoop
End Sub

Sub FormFadeGrey(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 255 - intLoop, 255 - intLoop), B
    Next intLoop
End Sub

Sub FormFadePurple(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 0, 255 - intLoop), B
    Next intLoop
End Sub

Sub FormFadeRed(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 0, 0), B
    Next intLoop
End Sub

Sub FormFadeYellow(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 255 - intLoop, 0), B
    Next intLoop
End Sub
Public Function GetCaption(WindowHandle As Long) As String
    'From Dos
    Dim Buffer As String, TextLength As Long
    TextLength& = GetWindowTextLength(WindowHandle&)
    Buffer$ = String(TextLength&, 0&)
    Call GetWindowText(WindowHandle&, Buffer$, TextLength& + 1)
    GetCaption$ = Buffer$
End Function
Sub ClickIcon(Icon)

Call SendMessage(Icon, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(Icon, WM_LBUTTONUP, 0, 0&)
End Sub

Public Sub SetText(Window As Long, Text As String)
    Call SendMessageByString(Window&, WM_SETTEXT, 0&, Text$)
End Sub
Sub ClickIcon2(TheButin As Long)
    Call PostMessage(TheButin&, WM_KEYDOWN, VK_SPACE, 0&)
    Call PostMessage(TheButin&, WM_KEYUP, VK_SPACE, 0&)
End Sub
Sub EnterKey(TheWin As Long)
    Call SendMessageByNum(TheWin&, WM_CHAR, ENTA, 0&)
End Sub
Sub EnableWin(Window&)
    Dim dis
    dis = EnableWindow(Window&, 1)
End Sub

Sub DisableWin(Window&)
    Dim dis
    dis = EnableWindow(Window&, 0)
End Sub

Function FindChildByClass(parentw, childhand)
firs% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(parentw, GW_CHILD)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone

While firs%
firss% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(firss%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(firs%, 2)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
Wend
FindChildByClass = 0

bone:
Room% = firs%
FindChildByClass = Room%

End Function


Function FindChildByTitle(parentw, childhand)

If UCase(GetCaption(GetWindow(parentw, 5))) Like UCase(childhand) Then GoTo bone
firs = GetWindow(parentw, GW_CHILD)

While firs

If UCase(GetCaption(GetWindow(parentw, 5))) Like UCase(childhand) & "*" Then GoTo bone
firs = GetWindow(GetWindow(parentw, 5), 2)
If UCase(GetCaption(GetWindow(parentw, 5))) Like UCase(childhand) & "*" Then GoTo bone
Wend
FindChildByTitle = 0

bone:
Room% = firs
FindChildByTitle = Room%
End Function
Function GetText(child)
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
getstring = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)
GetText = TrimSpace$
End Function
Function GetClass(child)
Buffer$ = String$(250, 0)
getclas% = GetClassName(child, Buffer$, 250)

GetClass = Buffer$
End Function



Sub Win_OnTop(TheFrm As Form)
    Dim SetOnTop

    SetOnTop = SetWindowPos(TheFrm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
Sub Win_Killwin(TheWind&)
    Call PostMessage(TheWind&, WM_CLOSE, 0&, 0&)
End Sub

Public Function GetWindowNext(child, CycleAmount As Integer)
For X = 1 To CycleAmount
child = GetWindow(child, GW_HWNDNEXT)
Next X
GetWindowNext = child
End Function

Public Sub WindowClassList(child, List As ListBox)
' Place a Listbox on the Form,
' This will Display Subsequent Children in List
' Note Item 1 is Item 0 on List
' Use this in conjunction with Function GetWindowNext,
' Place Item # in Cycle Amount

Do

item = item + 1
List.AddItem item & " " & GetText(child)
child = GetWindow(child, GW_HWNDNEXT)

Loop Until child = 0
End Sub
Sub WriteINI(Section As String, Key As String, KeyValue As String, Directory As String)

    Call WritePrivateProfileString(Section$, (Key$), KeyValue$, Directory$)
End Sub

Function ReadINI(Section As String, Key As String, Directory As String) As String
   Dim Buff As String

   Buff = String(750, Chr(0))
   Key$ = (Key$)
   ReadINI = Left(Buff, GetPrivateProfileString(Section$, ByVal Key$, "", Buff, Len(Buff), Directory$))
End Function
Sub ListKillDup(Lst As ListBox)
    Dim i, Duplicate
    For i = 0 To Lst.ListCount - 1
        For Duplicate = 0 To Lst.ListCount - 1
        If LCase(Lst.List(i)) Like LCase(Lst.List(Duplicate)) And i <> Duplicate Then
            Lst.RemoveItem (Duplicate)
        End If
        Next Duplicate
    Next i
End Sub

Sub ComboBoxLoad(Path As String, Combo As ComboBox)
'Call Load_ComboBox("c:\windows\desktop\combo.cmb", Combo1)

    Dim what As String
    On Error Resume Next
    Open Path$ For Input As #1
    While Not EOF(1)
        Input #1, what$
        DoEvents
        Combo.AddItem what$
    Wend
    Close #1
End Sub
Sub ListLoad(Path As String, Lst As ListBox)
'Ex: Call Load_ListBox("c:\windows\desktop\list.lst", list1)

    Dim what As String
    On Error Resume Next

    Open Path$ For Input As #1
    While Not EOF(1)
        Input #1, what$
        DoEvents
        Lst.AddItem what$
    Wend
    Close #1
End Sub
Sub TextLoad(txt As TextBox, FilePath As String)
'Ex: Call load_Text(list1,"c:\windows\desktop\text.txt")

    Dim mystr As String, FilePath2 As String, textz As String, a As String
    
    Open FilePath2$ For Input As #1
    Do While Not EOF(1)
    Line Input #1, a$
        textz$ = textz$ + a$ + Chr$(13) + Chr$(10)
        Loop
        txt = textz$
    Close #1
End Sub
Sub timeout(interval)
    Dim Current
    
    Current = Timer
    Do While Timer - Current < Val(interval)
        DoEvents
    Loop
End Sub
Function ReplaceString(MyString As String, ToFind As String, ReplaceWith As String) As String
' By dos, from dos23.bas He gets all the credit for this one
    Dim Spot As Long, NewSpot As Long, LeftString As String
    Dim RightString As String, NewString As String
    Spot& = InStr(LCase(MyString$), LCase(ToFind))
    NewSpot& = Spot&
    Do
        If NewSpot& > 0& Then
            LeftString$ = Left(MyString$, NewSpot& - 1)
            If Spot& + Len(ToFind$) <= Len(MyString$) Then
                RightString$ = Right(MyString$, Len(MyString$) - NewSpot& - Len(ToFind$) + 1)
            Else
                RightString = ""
            End If
            NewString$ = LeftString$ & ReplaceWith$ & RightString$
            MyString$ = NewString$
        Else
            NewString$ = MyString$
        End If
        Spot& = NewSpot& + Len(ReplaceWith$)
        If Spot& > 0 Then
            NewSpot& = InStr(Spot&, LCase(MyString$), LCase(ToFind$))
        End If
    Loop Until NewSpot& < 1
    ReplaceString$ = NewString$
End Function
Sub RunMenuByString(Application, StringSearch)
' From Hix he gets full credit

    Dim ToSearch As Integer, MenuCount As Integer, FindString
    Dim ToSearchSub As Integer, MenuItemCount As Integer, getstring
    Dim SubCount As Integer, MenuString As String, GetStringMenu As Integer
    Dim MenuItem As Integer, RunTheMenu As Integer
    
    ToSearch% = GetMenu(Application)
    MenuCount% = GetMenuItemCount(ToSearch%)
    
    For FindString = 0 To MenuCount% - 1
        ToSearchSub% = GetSubMenu(ToSearch%, FindString)
        MenuItemCount% = GetMenuItemCount(ToSearchSub%)
        For getstring = 0 To MenuItemCount% - 1
            SubCount% = GetMenuItemID(ToSearchSub%, getstring)
            MenuString$ = String$(100, " ")
            GetStringMenu% = GetMenuString(ToSearchSub%, SubCount%, MenuString$, 100, 1)
            If InStr(UCase(MenuString$), UCase(StringSearch)) Then
                MenuItem% = SubCount%
                GoTo MatchString
            End If
    Next getstring
    Next FindString
MatchString:
    RunTheMenu% = SendMessage(Application, WM_COMMAND, MenuItem%, 0)
End Sub
Sub ComboBoxSave(Path As String, Combo As ComboBox)
'Ex: Call Save_ComboBox("c:\windows\desktop\combo.cmb", combo1)

    Dim Savez As Long
    On Error Resume Next

    Open Path$ For Output As #1
    For Savez& = 0 To Combo.ListCount - 1
        Print #1, Combo.List(Savez&)
    Next Savez&
    Close #1
End Sub
Sub ListSave(Path As String, Lst As ListBox)
'Ex: Call Save_ListBox("c:\windows\desktop\list.lst", list1)

    Dim Listz As Long
    On Error Resume Next

    Open Path$ For Output As #1
    For Listz& = 0 To Lst.ListCount - 1
        Print #1, Lst.List(Listz&)
        Next Listz&
    Close #1
End Sub
Sub TextSave(txt As String, FilePath3 As String)
'Ex: Call Save_Text(list1,"c:\windows\desktop\text.txt")
   
    Open FilePath3$ For Output As #1
        Print #1, txt
    Close 1
End Sub
Sub Win_Center(frmz As Form)

    frmz.Top = (Screen.Height * 0.85) / 2 - frmz.Height / 2
    frmz.Left = Screen.Width / 2 - frmz.Width / 2
End Sub
Sub CNTALTDEL_Disable()
     Dim ret As Integer
     Dim pOld As Boolean

     ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
End Sub
Sub CNTALTDEL_Enable()
     Dim ret As Integer
     Dim pOld As Boolean

     ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
End Sub

Sub Win_Hide(TheWin As Long)
    Call ShowWindow(TheWin&, SW_HIDE)
End Sub

Sub Win_Maximize(THeWindow As Long)
    Dim max As Long

    max& = ShowWindow(THeWindow&, SW_MAXIMIZE)
End Sub
Sub Win_Minimize(THeWindow As Long)
    Dim Mini As Long

    Mini& = ShowWindow(THeWindow&, SW_MINIMIZE)
End Sub
Sub Win_Playwav(FilePath As String)
    Dim SoundName As String, Pla, Flagz As Integer
    
    SoundName$ = FilePath$
    Flagz% = SND_ASYNC Or SND_NODEFAULT
    Pla = sndPlaySound(SoundName$, Flagz%)
End Sub

Sub Win_Restore(THeWindow As Long)
    Dim res As Long

    res& = ShowWindow(THeWindow&, SW_RESTORE)
End Sub
Sub Win_Show(TheWin As Long)
    Call ShowWindow(TheWin&, SW_SHOW)
End Sub

Sub Win_StartButtin()
    Dim WinShell As Long, StartButtin As Long, Klick As Long

    WinShell& = FindWindow("Shell_TrayWnd", "")
    StartButtin& = FindWindowEx(WinShell&, 0, "Button", vbNullString)
    Call SendMessage(StartButtin&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(StartButtin&, WM_KEYUP, VK_SPACE, 0&)
End Sub
Sub Win_Shell(TheExe As String)
    Dim Shellz As Long, NoFreeze As Long

    Shellz& = Shell(TheExe$, 1): NoFreeze& = DoEvents()
End Sub
Sub Win_Unload(TheFrm As Form)
    Unload TheFrm
    End
    End
    Unload TheFrm
End Sub
Sub Copy(TextToCopy As String)
If TextToCopy$ = "" Then: Exit Sub
a$ = TextToCopy$
Clipboard.Clear
Clipboard.SetText (a$)

End Sub
Sub ListDeleteItem(Lst As ListBox, item$)
On Error Resume Next
Do
NoFreeze% = DoEvents()
If LCase$(Lst.List(a)) = LCase$(item$) Then Lst.RemoveItem (a)
a = 1 + a
Loop Until a >= Lst.ListCount
End Sub
Public Function GetChildCount(ByVal hwnd As Long) As Long
Dim hChild As Long

Dim i As Integer
   
If hwnd = 0 Then
GoTo Return_False
End If

hChild = GetWindow(hwnd, GW_CHILD)

While hChild
hChild = GetWindow(hChild, GW_HWNDNEXT)
i = i + 1
Wend

GetChildCount = i
   
Exit Function
Return_False:
GetChildCount = 0
Exit Function
End Function
Function TextGetLineCount(Text)

theview$ = Text


For FindChar = 1 To Len(theview$)
thechar$ = Mid(theview$, FindChar, 1)

If thechar$ = Chr(13) Then
numline = numline + 1
End If

Next FindChar

If Mid(Text, Len(Text), 1) = Chr(13) Then
GetLineCount = numline
Else
GetLineCount = numline + 1
End If
End Function
Public Function ListGetIndexText(oListBox As ListBox, sText As String) As Integer

Dim iIndex As Integer

With oListBox
 For iIndex = 0 To .ListCount - 1
   If .List(iIndex) = sText Then
    GetListIndex = iIndex
    Exit Function
   End If
 Next iIndex
End With

GetListIndex = -2   '  if Item isnt found
'( I didnt want to use -1 as it evaluates to True)

End Function

Sub ListToList(source, destination)
counts = SendMessage(source, LB_GETCOUNT, 0, 0)

For Adding = 0 To counts - 1
Buffer$ = String$(250, 0)
getstrings% = SendMessageByString(source, LB_GETTEXT, Adding, Buffer$)
addstrings% = SendMessageByString(destination, LB_ADDSTRING, 0, Buffer$)
Next Adding

End Sub
Sub FORMNotOnTop(the As Form)
'If You Dont Want Your Text On Top Of Everything
'But Shitty Code So You Gotta Make This In A EXE To See
'How It Werkx
SetWinOnTop = SetWindowPos(the.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
Sub Paste(hwnd As Integer)
a$ = Clipboard.GetText()
If a$ = "" Then: Exit Sub
X = SendMessageByString(whnd%, WM_PASTE, 0, a$)

End Sub

Function LISTSearchForSelected(Lst As ListBox)
If Lst.List(0) = "" Then
counterf = 0
GoTo last
End If
counterf = -1

Start:
counterf = counterf + 1
If Lst.ListCount = counterf + 1 Then GoTo last
If Lst.Selected(counterf) = True Then GoTo last
If couterf = Lst.ListCount Then GoTo last
GoTo Start

last:
SearchForSelected = counterf
End Function
Function LISTSearchForIndexbyText(Lst As ListBox, Text As String)
If Lst.ListIndex < 0 Then
LISTSearchForIndexbyText = -1
Exit Function
End If
X = 0
Do
If Lst.List(X) = Text Then
LISTSearchForIndexbyText = X
Exit Do
Else
X = X + 1
End If
DoEvents

Loop While Lst.ListIndex < Lst.ListCount

End Function
