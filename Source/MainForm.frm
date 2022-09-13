VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   0  'None
   ClientHeight    =   555
   ClientLeft      =   4500
   ClientTop       =   4560
   ClientWidth     =   675
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   555
   ScaleWidth      =   675
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Timer1.Interval = 2500
Call Load_Base
End Sub




Private Sub Timer1_Timer()
Dim i, k, l As Integer
Dim res, vr As Single
zagolovok = LCase(GetStr)


If zagolovok <> history Then history = zagolovok: Call Hist(history & "    " & Date & "     " & Time)

'1) Проверяет есть ли вхождение слова из белого списка в заголовке
'2) Проверяет есть ли вхождение запретного слова полностью в заголовке
'3) Если все ок, проверяем содержит ли заголовок GOOGLE, и если есть, узнаем процент вхожденя по алгоритму не очень точному))
'
If Get_White(zagolovok) = False Then
    If Get_Sovp(zagolovok) = False Then
    
        If Get_iscl(zagolovok) = False And InStr(1, zagolovok, "google", vbTextCompare) Then
        szg = Split(zagolovok, " ") ' Загружаем заголовок в массив
        For i = 0 To UBound(zapret)
            'Debug.Print ("Запрет в цикле: " + zapret(i))
            If res <> 0 Then ' Обработка результата
            res = res / (UBound(szp) + 1)
            'Debug.Print "Сработка ", zapret(i - 1), " %", res, i - 1
            Call Addbase(zapret(i) & "    " & Date & "     " & Time)
            Call CloseTab
            res = 0
            End If
            szp = Split(zapret(i), " ")
            For k = 0 To UBound(szp)
                'Debug.Print ("Это szp или запрет в цикле к: " + szp(k))
                For l = 0 To UBound(szg)
                If Len(szg(l)) > 2 Then
                    'Debug.Print ("Запрашиваем процент : " + szg(l) + " и " + szp(k))
                    vr = GetProc(szg(l), szp(k))
                    If vr <> 0 Then res = res + vr
                End If
                Next l
            Next k
        Next i
        End If
    Else:
    Call CloseTab
    End If
End If
End Sub
