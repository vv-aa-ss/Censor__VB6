Attribute VB_Name = "Function"
Option Explicit
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Const VK_CONTROL As Long = &H11
Private Const VK_W = &H57
Private Const KEYEVENTF_KEYUP = &H2

Public zapret() As String
Public whitelist() As String
Public zagolovok As String
Public szg() As String
Public szp() As String
Public iskl() As String
Public history As String

Type Slovo
n As Byte
k As Byte
max As Byte
End Type
Public sl As Slovo





Function GetStr() As String             ' «апрашивает заголовок активного окна, и возвращает его
Dim hwnd As Long, leng As Long
Dim WindowText$
Dim s As String
WindowText$ = Space(256)
hwnd = GetForegroundWindow()
    If hwnd > 0 Then
    leng = GetWindowTextLength(hwnd) + 1
    GetWindowText hwnd, WindowText$, leng
    GetStr = Mid(WindowText$, 1, Len(Trim(WindowText$)) - 1)
    End If
End Function

Function Get_White(str As String) As Boolean  ' ѕровер€ет есть ли заголовок в белом списке
Dim i As Integer
For i = 0 To UBound(whitelist)
    If InStr(1, str, whitelist(i), vbTextCompare) Then
    Get_White = True
    'Debug.Print "≈сть белый"
    i = UBound(whitelist)
    End If
Next i
End Function

Function Get_iscl(str As String) As Boolean ' I?iaa?yai anou ee enee??aiea a aaca
Dim i As Integer
    For i = 0 To UBound(iskl)
        If InStr(1, str, iskl(i), vbTextCompare) And Len(iskl(i)) = Len(str) Then
        Get_iscl = True
        i = UBound(iskl)
        End If
    Next i
endf:
End Function

Function GetProc(zag As String, zap As String) As Single
Dim kolvx As Byte
Dim byfer As String
Dim proc As Single
Dim resul As Single
If Get_Len(zag, zap) = True Then
byfer = 0
sl.n = 1
sl.k = 1
sl.max = Len(zag)
Do
    If InStr(1, zap, Mid(zag, sl.n, sl.k), vbTextCompare) Then
        byfer = sl.k
        sl.k = sl.k + 1
    Else
    sl.k = 1
            If byfer > 1 Then
            kolvx = kolvx + byfer
            sl.n = sl.n + byfer
            byfer = 0
            Else
            sl.n = sl.n + 1
            End If
    End If
Loop Until sl.n > sl.max Or sl.n + sl.k - 1 > sl.max
If byfer <> 0 And byfer <> 1 Then kolvx = kolvx + byfer
If kolvx <> 0 Then resul = kolvx / Len(zap)
    If resul > 0.7 Then
    GetProc = resul
    End If
End If
End Function


Function Get_Len(str, str1 As String) As Boolean
Dim x As Byte
    If Len(str) > Len(str1) Then
        x = Len(str) - Len(str1)
        Else
        x = Len(str1) - Len(str)
    End If
    If x > 2 Then
        Get_Len = False
    Else
        Get_Len = True
    End If
End Function

Function Get_Sovp(stroka) As Boolean
Dim i As Integer
For i = 0 To UBound(zapret)
 If InStr(1, stroka, zapret(i), vbTextCompare) Then Get_Sovp = True:  Call Addbase(zapret(i) & "    " & Date & "     " & Time): i = UBound(zapret)
Next i
End Function



Sub CloseTab()
    Main.Timer1.Interval = 0
    keybd_event VK_CONTROL, 0, 0, 0
    keybd_event VK_W, 0, 0, 0
    Sleep (100)
    keybd_event VK_W, 0, KEYEVENTF_KEYUP, 0
    keybd_event VK_CONTROL, 0, KEYEVENTF_KEYUP, 0
    Sleep (3000)
    Main.Timer1.Interval = 3000
    zagolovok = LCase(GetStr)
End Sub
