VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLinkList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'http://tieba.baidu.com/f?ct=335675392&tn=baiduPostBrowser&sc=9329172937&z=857679396&pn=0&rn=30&lm=0&word=vb#9329172937
'=======================
'Name_ENG LinkList
'Name_CHS 双向链表
'Version  A,v1.0
'No.Repair  4
'Author  Xup
'WorkGroup  ImageSoft
'Date  10-8-12
'-----------------------
'感谢使用,请保留版权信息
'-----------------------
'全内存处理链表
'-----------------------
'原作  LionKing1990
'-----------------------
' 这算是完全模仿狮子王
'的那个双向链表了,功能差
'不多,添加了一些接口支持
'---弄了一些通过标识查询
'读写的接口,至此.____Xup.
'=======================
Private Declare Function GlobalAlloc Lib "Kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "Kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalReAlloc Lib "Kernel32" (ByVal hMem As Long, ByVal dwBytes As Long, ByVal wFlags As Long) As Long
Private Const GMem_Fixed = &H0  '填充
Private Const GMem_ZeroInit = &H40 '置零
Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub ZeroMemory Lib "Kernel32" Alias "RtlMoveMemory" (ByVal Dest As Any, ByVal numBytes As Long)
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Type ItemType
    LastPtr  As Long   '上项指针
    NextPtr  As Long   '下项指针
    PtrSymbol As Long   '标识指针
    LenSymbol As Long   '标识长度
    PtrData  As Long   '数据指针
    LenData  As Long   '数据长度
End Type
Private Type MainType
    Count  As Long   '总数
    StartPtr  As Long   '首项指针
    EndPtr   As Long   '尾项指针
End Type
Private Const Len1 As Long = 8
Private Const Len2 As Long = 16
Private Const Len3 As Long = 24
Dim Lt As MainType
'=========================================================================================
'内存处理
'=========================================================================================
'===========================================================
'名称提示
'-----------------------------------------------------------
'A  使用Symbol
'B  使用Index
'C  使用Ptr
'===========================================================
'===================================================
'初始化
'===================================================
Private Sub Class_Initialize()
    Lt.Count = 0
    Lt.StartPtr = 0
    Lt.EndPtr = 0
End Sub
Private Sub Class_Terminate()
    Clear
End Sub
'===================================================
'内置表
'===================================================
'------------------
'申请内存并返回指针
'------------------
Private Function MemAlloc(Length As Long) As Long
    MemAlloc = GlobalAlloc(GMem_Fixed Or GMem_ZeroInit, Length)
End Function
'----------------
'返回字节数组长度
'----------------
Private Function GetLength(Bytes() As Byte) As Long
    GetLength = UBound(Bytes) - LBound(Bytes) + 1
End Function
'--------------------
'读取内存数据到缓冲区
'--------------------
Private Sub GetData(Ptr As Long, Length As Long, Buffer() As Byte)
    If Ptr = 0 Or Length < 1 Then Exit Sub
    ReDim Buffer(1 To Length)
    CopyMemory Buffer(1), ByVal Ptr, Length
End Sub
'----------------------------
'将数据保存到内存中并返回指针
'----------------------------
Private Function SaveData(Data() As Byte, PtrData As Long, LenData As Long) As Boolean
    LenData = GetLength(Data)
    PtrData = MemAlloc(LenData)
    If PtrData = 0 Then
        SaveData = False
        Exit Function
    End If
    CopyMemory ByVal PtrData, Data(LBound(Data)), LenData
    SaveData = True
End Function
'-----------------
'根据Index返回指针
'-----------------
Private Sub Scan(Index As Long, NowPtr As Long, LastPtr As Long, NextPtr As Long)
    If Index < 1 Or Index > Lt.Count Then Exit Sub
    NowPtr = Lt.StartPtr
    Dim i As Long, tL As ItemType
    For i = 2 To Index
        GetItem1 NowPtr, tL
        NowPtr = tL.NextPtr
    Next i
    GetItem1 NowPtr, tL
    LastPtr = tL.LastPtr
    NextPtr = tL.NextPtr
End Sub
'---------------------------------------------------
'简单的取值解释
'--------------
'1  下项指针
'2  下项指针+标识信息
'3  下项指针+标识信息+数据信息
'------------------------------------------------------
Private Sub GetItem1(Ptr As Long, Buffer As ItemType)
    If Ptr = 0 Then Exit Sub
    CopyMemory Buffer, ByVal Ptr, Len1
End Sub
Private Sub SetItem1(Ptr As Long, Buffer As ItemType)
    If Ptr = 0 Then Exit Sub
    CopyMemory ByVal Ptr, Buffer, Len1
End Sub
'----------------------------------------------------------
Private Sub GetItem2(Ptr As Long, Buffer As ItemType)
    If Ptr = 0 Then Exit Sub
    CopyMemory Buffer, ByVal Ptr, Len2
End Sub
Private Sub SetItem2(Ptr As Long, Buffer As ItemType)
    If Ptr = 0 Then Exit Sub
    CopyMemory ByVal Ptr, Buffer, Len2
End Sub
'----------------------------------------------------------
Private Sub GetItem3(Ptr As Long, Buffer As ItemType)
    If Ptr = 0 Then Exit Sub
    CopyMemory Buffer, ByVal Ptr, Len3
End Sub
Private Sub SetItem3(Ptr As Long, Buffer As ItemType)
    If Ptr = 0 Then Exit Sub
    CopyMemory ByVal Ptr, Buffer, Len3
End Sub
'------------------
'内存申请失败的处理
'------------------
Private Sub MemAllocFail()
    MsgBox "内存申请失败", 16
End Sub
Private Function AddItemEx(Symbol() As Byte, Data() As Byte) As Long
    Dim tL As ItemType
    With tL
        If SaveData(Symbol, .PtrSymbol, .LenSymbol) = False Then
            MemAllocFail
            Exit Function
        End If
        If SaveData(Data, .PtrData, .LenData) = False Then
            MemAllocFail
            Exit Function
        End If
    End With
    Dim iPtr As Long
    iPtr = MemAlloc(20)
    If iPtr = 0 Then
        MemAllocFail
        Exit Function
    End If
    If Lt.Count = 0 Then
        Lt.StartPtr = iPtr
        Lt.EndPtr = iPtr
    Else
        Dim tLB As ItemType
        GetItem1 Lt.EndPtr, tLB
        tLB.NextPtr = iPtr
        SetItem1 Lt.EndPtr, tLB
        Lt.EndPtr = iPtr
    End If
    CopyMemory ByVal iPtr, tL, Len3
    Lt.Count = Lt.Count + 1
    AddItemEx = iPtr
End Function
'===================================================
'过程表
'===================================================
'------
'添加项
'------
Sub AddItem(Symbol As String, Data() As Byte, Optional Index As Long)
    Dim tB() As Byte
    tB = StrConv(Symbol, vbFromUnicode)
    Dim iPtr As Long
    iPtr = AddItemEx(tB, Data)
    If Index > 0 Then MoveItemCB iPtr, Index
End Sub
'------
'清空项
'------
Sub Clear()
    Dim i As Long, tPtr As Long, tL As ItemType
    tPtr = Lt.StartPtr
    For i = 1 To Lt.Count
        GetItem3 tPtr, tL
        GlobalFree tL.PtrData
        GlobalFree tL.PtrSymbol
        GlobalFree tPtr
        tPtr = tL.NextPtr
        DoEvents
    Next i
    Lt.Count = 0
    Lt.StartPtr = 0
    Lt.EndPtr = 0
End Sub
'------
'删除项
'------
Sub RemoveItemA(Symbol As String)
    RemoveItemC GetItemPtrA(Symbol)
End Sub
Sub RemoveItemB(Index As Long)
    RemoveItemC GetItemPtrB(Index)
End Sub
Sub RemoveItemC(Ptr As Long)
    If Ptr = 0 Then Exit Sub
    Dim tL As ItemType, tLB As ItemType, tLC As ItemType
    GetItem3 Ptr, tL
    GetItem1 tL.LastPtr, tLB
    GetItem1 tL.NextPtr, tLC
    Select Case Ptr
    Case Lt.StartPtr, Lt.EndPtr
        If Ptr = Lt.StartPtr Then
            tLC.LastPtr = 0
            SetItem1 tL.NextPtr, tLC
            Lt.StartPtr = tL.NextPtr
        End If
        If Ptr = Lt.EndPtr Then
            tLB.NextPtr = 0
            SetItem1 tL.LastPtr, tLB
            Lt.EndPtr = tL.LastPtr
        End If
    Case Else
        tLB.NextPtr = tL.NextPtr
        SetItem1 tL.LastPtr, tLB
        tLC.LastPtr = tL.LastPtr
        SetItem1 tL.NextPtr, tLC
    End Select
    GlobalFree tL.PtrSymbol
    GlobalFree tL.PtrData
    GlobalFree Ptr
End Sub
'------
'移动项
'------
Sub MoveItemAA(SocSymbol As String, DestSymbol As String)
    MoveItemCC GetItemPtrA(SocSymbol), GetItemPtrA(DestSymbol)
End Sub
Sub MoveItemAB(SocSymbol As String, DestIndex As Long)
    MoveItemCC GetItemPtrA(SocSymbol), GetItemPtrB(DestIndex)
End Sub
Sub MoveItemAC(SocSymbol As String, DestPtr As Long)
    MoveItemCC GetItemPtrA(SocSymbol), DestPtr
End Sub
Sub MoveItemBA(SocIndex As Long, DestSymbol As String)
    MoveItemCC GetItemPtrB(SocIndex), GetItemPtrA(DestSymbol)
End Sub
Sub MoveItemBB(SocIndex As Long, DestIndex As Long)
    MoveItemCC GetItemPtrB(SocIndex), GetItemPtrB(DestIndex)
End Sub
Sub MoveItemBC(SocIndex As Long, DestPtr As Long)
    MoveItemCC GetItemPtrB(SocIndex), DestPtr
End Sub
Sub MoveItemCA(SocPtr As Long, DestSymbol As String)
    MoveItemCC SocPtr, GetItemPtrA(DestSymbol)
End Sub
Sub MoveItemCB(SocPtr As Long, DestIndex As Long)
    MoveItemCC SocPtr, GetItemPtrB(DestIndex)
End Sub
Sub MoveItemCC(SocPtr As Long, DestPtr As Long)
    If SocPtr = 0 Or DestPtr = 0 Then Exit Sub
    If SocPtr = DestPtr Then Exit Sub
    If Lt.Count < 2 Then Exit Sub
    Dim tL As ItemType, tLB As ItemType, tLC As ItemType
    '获取源信息
    GetItem1 SocPtr, tL
    GetItem1 tL.LastPtr, tLB
    GetItem1 tL.NextPtr, tLC
    '上脱钩
    tLB.NextPtr = tL.NextPtr
    SetItem1 tL.LastPtr, tLB
    '改变尾指针
    If tL.NextPtr = 0 Then Lt.EndPtr = tL.LastPtr
    '下脱钩
    tLC.LastPtr = tL.LastPtr
    SetItem1 tL.NextPtr, tLC
    '改变首指针
    If tL.LastPtr = 0 Then Lt.StartPtr = tL.NextPtr
    '获取目标信息
    GetItem1 DestPtr, tLC
    GetItem1 tLC.LastPtr, tLB
    '上连接
    tLB.NextPtr = SocPtr
    SetItem1 tLC.LastPtr, tLB
    '改变首指针
    If tLC.LastPtr = 0 Then Lt.StartPtr = SocPtr
    '下连接
    tLC.LastPtr = SocPtr
    SetItem1 DestPtr, tLC
    tL.NextPtr = DestPtr
    SetItem1 SocPtr, tL
End Sub
'------------
'设置数据长度
'------------
Sub SetItemLengthA(Symbol As String, NewLength As Long)
    SetItemLengthC GetItemPtrA(Symbol), NewLength
End Sub
Sub SetItemLengthB(Index As Long, NewLength As Long)
    SetItemLengthC GetItemPtrB(Index), NewLength
End Sub
Sub SetItemLengthC(Ptr As Long, NewLength As Long)
    If Ptr = 0 Or NewLength < 0 Then Exit Sub
    Dim tL As ItemType
    GetItem3 Ptr, tL
    If NewLength = 0 Then
        GlobalFree tL.PtrData
        tL.PtrData = 0
    Else
        Dim tB() As Byte
        GetData tL.PtrData, tL.LenData, tB
        If NewLength < tL.LenData Then
            tB = MidB$(tB, 1, NewLength)
        Else
            ReDim Preserve tB(1 To NewLength)
        End If
        Dim tPtr As Long
        If SaveData(tB, tPtr, tL.LenData) = False Then
            MemAllocFail
            Exit Sub
        End If
        GlobalFree tL.PtrData
        tL.PtrData = tPtr
    End If
    SetItem3 Ptr, tL
End Sub
'------------
'设置数据标识
'------------
Sub SetItemSymbolA(OldSymbol As String, NewSymbol As String)
    SetItemSymbolC GetItemPtrA(OldSymbol), NewSymbol
End Sub
Sub SetItemSymbolB(Index As Long, NewSymbol As String)
    SetItemSymbolC GetItemPtrB(Index), NewSymbol
End Sub
Sub SetItemSymbolC(Ptr As Long, NewSymbol As String)
    If Ptr = 0 Then Exit Sub
    Dim tL As ItemType
    GetItem2 Ptr, tL
    Dim tB() As Byte
    tB = StrConv(NewSymbol, vbFromUnicode)
    Dim tPtr As Long
    If SaveData(tB, tPtr, tL.LenSymbol) = False Then
        MemAllocFail
        Exit Sub
    End If
    GlobalFree tL.PtrSymbol
    tL.PtrSymbol = tPtr
    SetItem2 Ptr, tL
End Sub
'--------
'设置数据
'--------
Sub SetItemDataA(Symbol As String, Data() As Byte)
    SetItemDataC GetItemPtrA(Symbol), Data
End Sub
Sub SetItemDataB(ByVal Index As Long, Data() As Byte)
    SetItemDataC GetItemPtrB(Index), Data
End Sub
Sub SetItemDataC(ByVal Ptr As Long, Data() As Byte)
    If Ptr = 0 Then Exit Sub
    Dim tL As ItemType
    GetItem3 Ptr, tL
    Dim tPtr As Long
    If SaveData(Data, tPtr, tL.LenData) = False Then
        MemAllocFail
        Exit Sub
    End If
    GlobalFree tL.PtrData
    tL.PtrData = tPtr
    SetItem3 Ptr, tL
End Sub
'===================================================
'函数表
'===================================================
'-----------
'获取项Index
'-----------
Function GetItemIndexA(ByVal Symbol As String) As Long
    Symbol = UCase$(Symbol)
    Dim tPtr As Long
    tPtr = Lt.StartPtr
    Dim i As Long, tL As ItemType, tB() As Byte, tS As String
    For i = 1 To Lt.Count
        GetItem2 tPtr, tL
        With tL
            GetData .PtrSymbol, .LenSymbol, tB
            tS = UCase$(StrConv(tB, vbUnicode))
            If tS = Symbol Then
                GetItemIndexA = i
                Exit For
            End If
            tPtr = .NextPtr
        End With
    Next i
End Function
Function GetItemIndexC(ByVal Ptr As Long) As Long
    If Ptr = 0 Then Exit Function
    Dim tPtr As Long
    tPtr = Lt.StartPtr
    Dim i As Long, tL As ItemType
    For i = 1 To Lt.Count
        If tPtr = Ptr Then
            GetItemIndexC = i
            Exit For
        End If
        GetItem1 tPtr, tL
        tPtr = tL.NextPtr
    Next i
End Function
'------------
'获取数据指针
'------------
Function GetItemPtrA(ByVal Symbol As String) As Long
    Symbol = UCase$(Symbol)
    Dim tPtr As Long
    tPtr = Lt.StartPtr
    Dim i As Long, tL As ItemType, tB() As Byte, tS As String
    For i = 1 To Lt.Count
        GetItem2 tPtr, tL
        With tL
            GetData .PtrSymbol, .LenSymbol, tB
            tS = UCase$(StrConv(tB, vbUnicode))
            If tS = Symbol Then
                GetItemPtrA = tPtr
                Exit For
            End If
            tPtr = .NextPtr
        End With
    Next i
End Function
Function GetItemPtrB(ByVal Index As Long) As Long
    If Index < 1 Or Index > Lt.Count Then Exit Function
    Dim tPtr As Long
    tPtr = Lt.StartPtr
    Dim i As Long, tL As ItemType
    For i = 2 To Index
        GetItem1 tPtr, tL
        tPtr = tL.NextPtr
    Next i
    GetItemPtrB = tPtr
End Function
'------------
'获取数据长度
'------------
Function GetItemLengthA(ByVal Symbol As String) As Long
    'Symbol = UCase$(Symbol)
    'Dim tPtr As Long
    'tPtr = Lt.StartPtr
    'Dim i As Long, tL As ItemType, tB() As Byte, tS As String
    'For i = 1 To Lt.Count
    '  GetItem3 tPtr, tL
    '  With tL
    '  GetData .PtrSymbol, .LenSymbol, tB
    '  tS = UCase$(StrConv(tB, vbUnicode))
    '  If tS = Symbol Then
    '  GetItemLengthA = tL.LenData
    '  Exit For
    '  End If
    '  tPtr = .NextPtr
    '  End With
    'Next i
    GetItemLengthA = GetItemLengthC(GetItemPtrA(Symbol))
End Function
Function GetItemLengthB(ByVal Index As Long) As Long
    'If Index < 1 Or Index > Lt.Count Then Exit Function
    'Dim tPtr As Long
    'tPtr = Lt.StartPtr
    'Dim i As Long, tL As ItemType
    'For i = 2 To Index
    '  GetItem1 tPtr, tL
    '  tPtr = tL.NextPtr
    'Next i
    'GetItem3 tPtr, tL
    'GetItemLengthB = tL.LenData
    GetItemLengthB = GetItemLengthC(GetItemPtrB(Index))
End Function
Function GetItemLengthC(ByVal Ptr As Long) As Long
    If Ptr = 0 Then Exit Function
    Dim tL As ItemType
    GetItem3 Ptr, tL
    GetItemLengthC = tL.LenData
End Function
'------------
'获取数据标识
'------------
Function GetItemSymbolB(ByVal Index As Long) As String
    GetItemSymbolB = GetItemSymbolC(GetItemPtrB(Index))
End Function
Function GetItemSymbolC(ByVal Ptr As Long) As String
    If Ptr = 0 Then Exit Function
    Dim tL As ItemType
    GetItem2 Ptr, tL
    With tL
        Dim tB() As Byte
        GetData .PtrSymbol, .LenSymbol, tB
        GetItemSymbolC = StrConv(tB, vbUnicode)
    End With
End Function
'--------
'获取数据
'--------
Function GetItemDataA(ByVal Symbol As String) As Byte()
    GetItemDataA = GetItemDataC(GetItemPtrA(Symbol))
End Function
Function GetItemDataB(ByVal Index As Long) As Byte()
    GetItemDataB = GetItemDataC(GetItemPtrB(Index))
End Function
Function GetItemDataC(ByVal Ptr As Long) As Byte()
    If Ptr = 0 Then Exit Function
    Dim tL As ItemType
    GetItem3 Ptr, tL
    With tL
        Dim tB() As Byte
        GetData .PtrData, .LenData, tB
        GetItemDataC = tB
    End With
End Function
'===================================================
'属性表
'===================================================
'--------
'读项总数
'--------
Property Get Count() As Long
    Count = Lt.Count
End Property
'=========================================================================================
'文件读写
'=========================================================================================
Sub LoadFile(Path As String)
    On Error GoTo ErrLF
    Dim FreeF As Integer, Pos As Long, Length As Long
    FreeF = FreeFile
    Length = FileLen(Path)
    Pos = 1
    Open Path For Binary As #FreeF
    Dim tLng As Long, tS As String, tB() As Byte
    Do Until Pos > Length
        tS = ""
        Erase tB
        Get #FreeF, , tLng
        Pos = Pos + 4 + tLng
        If tLng > 0 Then
            ReDim tB(1 To tLng)
            Get #FreeF, , tB
            tS = StrConv(tB, vbUnicode)
            Get #FreeF, , tLng
            Pos = Pos + 4 + tLng
            If tLng > 0 Then
                ReDim tB(1 To tLng)
                Get #FreeF, , tB
            End If
        End If
        AddItem tS, tB
        DoEvents
    Loop
    Close #FreeF
ErrLF:
End Sub
Sub SaveFile(Path As String)
    On Error GoTo ErrSF
    Dim FreeF As Integer
    FreeF = FreeFile
    Open Path For Output As #FreeF
    Close #FreeF
    Open Path For Binary As #FreeF
    Dim tPtr As Long
    tPtr = Lt.StartPtr
    Dim i As Long, tL As ItemType, tB() As Byte
    For i = 1 To Lt.Count
        GetItem3 tPtr, tL
        Put #FreeF, , tL.LenSymbol
        GetData tL.PtrSymbol, tL.LenSymbol, tB
        Put #FreeF, , tB
        Put #FreeF, , tL.LenData
        GetData tL.PtrData, tL.LenData, tB
        Put #FreeF, , tB
        tPtr = tL.NextPtr
        DoEvents
    Next i
    Close #FreeF
ErrSF:
End Sub
'=========================================================================================
'其他接口
'=========================================================================================

'--------
'测试添加
'--------
Function TestAddItem(Count As Long, SymbolHeader As String, Data() As Byte) As Long
    Dim i As Long
    Dim Time As Long
    Time = timeGetTime
    For i = 1 To Count
        AddItem SymbolHeader & Chr$(9) & i, Data
        DoEvents
    Next i
    TestAddItem = timeGetTime - Time
    Debug.Print "添加 " & Count & " 个，标识基长度 " & Len(SymbolHeader) & " 测试数据长度 " & GetLength(Data) & " 共耗时 " & TestAddItem
End Function
'--------
'测试清空
'--------
Function TestClear() As Long
    Dim Time As Long
    Time = timeGetTime
    Clear
    TestClear = timeGetTime - Time
    Debug.Print "清空耗时 " & TestClear
End Function
'--------
'测试获取
'--------
Function TestGet(Count As Long) As Long
    Dim Time As Long
    Time = timeGetTime
    Dim i As Long, t As Long
    VBA.Randomize
    For i = 1 To Count
        t = Int(Rnd * Lt.Count) + 1
        Debug.Print "No." & i & Chr$(9) & "Index:" & t & Chr$(9) & GetItemSymbolB(t) & Chr$(9) & StrConv(GetItemDataB(t), vbUnicode)
        DoEvents
    Next i
    TestGet = timeGetTime - Time
End Function
