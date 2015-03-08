VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  '屏幕中心
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private c As cLinkList


Private Sub Form_Load()
    Set c = New cLinkList
    
    TextSpeed CLng(10) * 10000
    
    '添加项
    c.AddItem "b"
    c.AddItem "d"
    '空项
    c.AddItem ""
    c.AddItem "f"
    DisPlay
    Debug.Print "d 的 index:"; c.Find("d")
    Debug.Print "-------------------"
    '指定index的添加
    c.AddItem "c", 2
    DisPlay
    Debug.Print "d 的 index:"; c.Find("d")
    Debug.Print "-------------------"
    
    c.AddItem "a", 1
    DisPlay
    Debug.Print "d 的 index:"; c.Find("d")
    Debug.Print "-------------------"
    '替换
    c.List(1) = "a变成a1"
    '替换空项为e
    c.List(c.Find("")) = "e"
    DisPlay
    Debug.Print "d 的 index:"; c.Find("d")
    Debug.Print "-------------------"
    
    '删除b
    c.Remove (c.Find("b"))
    DisPlay
    Debug.Print "d 的 index:"; c.Find("d")
    Debug.Print "-------------------"
    
End Sub

Private Sub DisPlay()
    Dim i As Long
    i = c.Count
    Debug.Print "共有 "; i; " 项"
    For i = 1 To i
        Debug.Print "第 "; i; " 项: "; c.List(i)
    Next
End Sub

Private Sub TextSpeed(ByVal Times As Long)
    Dim i As Long
    Dim t As Long
    t = timeGetTime
    For i = 1 To Times
        c.AddItem CStr(i)
    Next i
    
    MsgBox "添加" & Times & "项用时" & (timeGetTime - t)
    '    T = timeGetTime
    '    Debug.Print c.Find; ""
    c.Clear
End Sub
