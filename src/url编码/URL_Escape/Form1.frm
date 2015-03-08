VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "编码/解码URL资源"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7185
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   7185
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   1200
      Width           =   6135
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   720
      Width           =   6135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   240
      Width           =   6135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "解码(&D)"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "编码(&E)"
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "解码URL"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "编码URL"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "源URL"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   450
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Download by http://www.codefans.net
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

Private Sub Form_Load()
    Text1.Text = "http://www.mvps.org/vbnet code lib/net code/ip address.htm"
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
