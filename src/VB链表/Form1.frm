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
   StartUpPosition =   2  '��Ļ����
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
    
    '�����
    c.AddItem "b"
    c.AddItem "d"
    '����
    c.AddItem ""
    c.AddItem "f"
    DisPlay
    Debug.Print "d �� index:"; c.Find("d")
    Debug.Print "-------------------"
    'ָ��index�����
    c.AddItem "c", 2
    DisPlay
    Debug.Print "d �� index:"; c.Find("d")
    Debug.Print "-------------------"
    
    c.AddItem "a", 1
    DisPlay
    Debug.Print "d �� index:"; c.Find("d")
    Debug.Print "-------------------"
    '�滻
    c.List(1) = "a���a1"
    '�滻����Ϊe
    c.List(c.Find("")) = "e"
    DisPlay
    Debug.Print "d �� index:"; c.Find("d")
    Debug.Print "-------------------"
    
    'ɾ��b
    c.Remove (c.Find("b"))
    DisPlay
    Debug.Print "d �� index:"; c.Find("d")
    Debug.Print "-------------------"
    
End Sub

Private Sub DisPlay()
    Dim i As Long
    i = c.Count
    Debug.Print "���� "; i; " ��"
    For i = 1 To i
        Debug.Print "�� "; i; " ��: "; c.List(i)
    Next
End Sub

Private Sub TextSpeed(ByVal Times As Long)
    Dim i As Long
    Dim t As Long
    t = timeGetTime
    For i = 1 To Times
        c.AddItem CStr(i)
    Next i
    
    MsgBox "���" & Times & "����ʱ" & (timeGetTime - t)
    '    T = timeGetTime
    '    Debug.Print c.Find; ""
    c.Clear
End Sub
