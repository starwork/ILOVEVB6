Attribute VB_Name = "Module1"
Option Explicit

Public Const MAX_PATH                   As Long = 260
Public Const ERROR_SUCCESS              As Long = 0

'将整个URL参数作为一个URL段
Public Const URL_ESCAPE_SEGMENT_ONLY    As Long = &H2000
Public Const URL_ESCAPE_PERCENT         As Long = &H1000
Public Const URL_UNESCAPE_INPLACE       As Long = &H100000

'路径中包含#
Public Const URL_INTERNAL_PATH          As Long = &H800000
Public Const URL_DONT_ESCAPE_EXTRA_INFO As Long = &H2000000
Public Const URL_ESCAPE_SPACES_ONLY     As Long = &H4000000
Public Const URL_DONT_SIMPLIFY          As Long = &H8000000

'转换不安全字符为相应的退格序列
Public Declare Function UrlEscape Lib "shlwapi" _
   Alias "UrlEscapeA" _
  (ByVal pszURL As String, _
   ByVal pszEscaped As String, _
   pcchEscaped As Long, _
   ByVal dwFlags As Long) As Long
'Download by http://www.codefans.net
'转换退格序列为普通的字符
Public Declare Function UrlUnescape Lib "shlwapi" _
   Alias "UrlUnescapeA" _
  (ByVal pszURL As String, _
   ByVal pszUnescaped As String, _
   pcchUnescaped As Long, _
   ByVal dwFlags As Long) As Long


