Attribute VB_Name = "BD"
Option Explicit

Sub Load_Base()
ReDim zapret(0) As String
ReDim whitelist(0) As String
ReDim iskl(0) As String
Dim kol As Integer

Close #99
    Open App.Path & "\zapr.baz" For Input As #99
    While Not EOF(99)
    Line Input #99, zapret(0)
    kol = kol + 1
    Wend
    Close #99
    ReDim zapret(kol - 1)
    kol = 0
    Open App.Path & "\zapr.baz" For Input As #99
    While Not EOF(99)
    Line Input #99, zapret(kol)
    kol = kol + 1
    Wend
Close #99
kol = 0
Close #99
    Open App.Path & "\white.baz" For Input As #99
    While Not EOF(99)
    Line Input #99, whitelist(0)
    kol = kol + 1
    Wend
    Close #99
    ReDim whitelist(kol - 1) As String
    kol = 0
    Open App.Path & "\white.baz" For Input As #99
    While Not EOF(99)
    Line Input #99, whitelist(kol)
    kol = kol + 1
    Wend
    Close #99
kol = 0
Close #99
    Open App.Path & "\iskl.baz" For Input As #99
    While Not EOF(99)
    Line Input #99, iskl(0)
    kol = kol + 1
    Wend
    Close #99
    ReDim iskl(kol - 1) As String
    kol = 0
    Open App.Path & "\iskl.baz" For Input As #99
    While Not EOF(99)
    Line Input #99, iskl(kol)
    kol = kol + 1
    Wend
Close #99
kol = 0
End Sub

Sub Addbase(str As String)
    Open App.Path & "\base.baz" For Append As #99
    Print #99, str
    Close #99
End Sub

Sub Hist(str As String)
    Open App.Path & "\history.baz" For Append As #99
    Print #99, str
    Close #99
End Sub

