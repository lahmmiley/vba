Attribute VB_Name = "TestModule"
Sub beCalledFunction(ByVal x As Integer, ByRef y As Integer)
    If y = 100 Then
        y = x + y
    Else
        y = x - y
        x = x + 100
    End If
End Sub

Sub callFunctionTest()
    Dim x1 As Integer
    Dim y1 As Integer
    x1 = 12
    y1 = 100
    Call beCalledFunction(x1, y1)
    Debug.Print x1, y1
End Sub

Sub ReadDataTest()
    Dim cellX As String
    Dim i As Integer
    For i = 1 To 1 Step 1
        'MsgBox (Application.Worksheets("¹¤×÷±í1").Range("B" & i))
    Next i
End Sub

Sub DictionaryTest()
    Dim d As New Dictionary
    d.Add "a", "wtf"
    d.Add "b", "wtf1"
    d.Add "c", "wtf2"
    d.Add "d", "wtf3"
    d.Remove "a"
    d.DebugPrint
End Sub

Sub FileTest()
    'Macintosh HD:Users:user:Documents:code:vba
    Debug.Print (CurDir())
    Open "Macintosh HD:Users:user:Documents:code:vba:Friends.txt" For Output As #1
    lname = "Smith"
    fname = "Gregory"
    birthdate = #1/2/63#
    s = 3
    Write #1, lname, fname, birthdate, "\r", "1", "\n"
    lname = "Conlin"
    fname = "Janice"
    birthdate = #5/12/48#
    s = 1
    Write #1, lname, fname, birthdate, s, "\r\n"
    Close #1
End Sub

Sub PrintTest()
    Open "Macintosh HD:Users:user:Documents:code:vba:print.txt" For Output As #1
    Dim doubleFlag As String
    Dim content As String
    content = content + "xmlversion"
    content = content + Chr(13)
    content = content + "xmlversion"
    content = content + Chr(10)
    content = content + "xmlversion"
    content = content + Chr(13) + Chr(10)
    content = content + "xmlversion"
    Print #1, content
'    doubleFlag = Chr(34)
'    Print #1, "xmlversion"; Chr(34); "10";
'    Print #1, "xmlversion"; Chr(34); "3 ";
'    Print #1, "xmlversion"; Chr(34); "3 ";
    Close #1
End Sub

'gb2312±àÂë
Public Function GBKEncode(szInput) As String
    Dim i As Long
    Dim startIndex As Long
    Dim endIndex As Long
    Dim x() As Byte
     
    x = StrConv(szInput, vbFromUnicode)
     
    startIndex = LBound(x)
    endIndex = UBound(x)
    For i = startIndex To endIndex
        GBKEncode = GBKEncode & "%" & Hex(x(i))
    Next
End Function
 
'GB2312±àÂë
Public Function GBKDecode(ByVal code As String) As String
    code = Replace(code, "%", "")
    Dim bytes(1) As Byte
    Dim index As Long
    Dim length As Long
    Dim codelen As Long
    codelen = Len(code)
    While (codelen > 3)
        For index = 1 To 2
            bytes(index - 1) = val("&H" & Mid(code, index * 2 - 1, 2))
        Next index
        GBKDecode = GBKDecode & StrConv(bytes, vbUnicode)
        code = Right(code, codelen - 4)
        codelen = Len(code)
    Wend
End Function
