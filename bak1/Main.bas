Attribute VB_Name = "Main"
Option Explicit
Public NAME_ROW_NUM As Integer
Public TYPE_ROW_NUM As Integer
Public DATA_ROW_NUM As Integer
Public DATA_START_COLUMN As Integer
Public FILE_NAME As String
Public DQM As String '双引号
Public CL_LF As String '换行符
Public S_TAB As String
Public AllExporter As New CExporter
Public ClientExporter As New CExporter
Public ServerExporter As New CExporter

Sub ConstantInitialization()
    NAME_ROW_NUM = 1
    TYPE_ROW_NUM = NAME_ROW_NUM + 1
    DATA_ROW_NUM = TYPE_ROW_NUM + 1
    DATA_START_COLUMN = 1
    FILE_NAME = ActiveSheet.name
    CL_LF = Chr(13) + Chr(10)
    DQM = Chr(34)
    S_TAB = "    "
    Call ClientExporter.Init("Macintosh HD:Users:user:Documents:code:vba:client:" & FILE_NAME & ".xml", True, False)
    Call ServerExporter.Init("Macintosh HD:Users:user:Documents:code:vba:server:" & FILE_NAME & ".xml", False, True)
    Call AllExporter.Init("Macintosh HD:Users:user:Documents:code:vba:all:" & FILE_NAME & ".xml", True, True)
End Sub

Function ReadType() As Dictionary
    Dim dataType As CDataType
    Dim dataTypeDict As New Dictionary
    Dim cellContent As String
    Dim rowNum As Integer, columnNum As Integer
    columnNum = DATA_START_COLUMN
    Do
        cellContent = ActiveSheet.Cells(TYPE_ROW_NUM, columnNum)
        If cellContent = "" Then
            Exit Do
        End If
        Set dataType = New CDataType
        Call DataTypeParser(dataType, _
            ActiveSheet.Cells(TYPE_ROW_NUM, columnNum), _
            ActiveSheet.Cells(NAME_ROW_NUM, columnNum))
        dataTypeDict.Add CStr(columnNum), dataType
        columnNum = columnNum + 1
    Loop While True
    Set ReadType = dataTypeDict
End Function

Function ReadData(columnMax As Integer) As Dictionary
    Dim rowNum As Integer, columnNum As Integer
    Dim dataDict As New Dictionary, rowDataDict As Dictionary
    Dim cellContent As String
    rowNum = DATA_ROW_NUM
    Do
        cellContent = ActiveSheet.Cells(rowNum, 1) 'Id
        If cellContent = "" Then
            Exit Do
        End If
        Set rowDataDict = New Dictionary
        For columnNum = 1 To columnMax
            rowDataDict.Add CStr(columnNum), Trim(ActiveSheet.Cells(rowNum, columnNum))
        Next columnNum
        dataDict.Add CStr(rowNum), rowDataDict
        rowNum = rowNum + 1
    Loop While True
    Set ReadData = dataDict
End Function

Function CellContentIsValid(cellContent As String, dataType As String) As Boolean
    Dim elementArray() As String
    Dim dictArray() As String
    Dim dictStr As String
    Dim i As Integer, j As Integer
    Select Case dataType
        Case "int"
            If Not IsInt(cellContent) Then
                CellContentIsValid = False
                Exit Function
            End If
        Case "array"
            elementArray = Split(cellContent, ",")
            For i = 0 To UBound(elementArray)
                If elementArray(i) = "" Then
                    CellContentIsValid = False
                    Exit Function
                End If
            Next i
        Case "dict"
            dictArray = Split(cellContent, ",")
            For i = 0 To UBound(dictArray)
                dictStr = dictArray(i)
                If dictStr = "" Then
                    CellContentIsValid = False
                    Exit Function
                ElseIf InStr(dictStr, ":") = 0 Then
                    CellContentIsValid = False
                    Exit Function
                End If
                elementArray = Split(dictStr, ":")
                If UBound(elementArray) <> 1 Then
                    CellContentIsValid = False
                    Exit Function
                End If
                For j = 0 To UBound(elementArray)
                    If (elementArray(j) = "") Then
                        CellContentIsValid = False
                        Exit Function
                    End If
                Next j
            Next i
    End Select
    CellContentIsValid = True
End Function

Function GenerateXMLContent(dataTypeDict As Dictionary, _
    dataDict As Dictionary, exporter As CExporter) As Boolean
    Dim content As String
    Dim columnDict As Dictionary
    Dim rowKVP As KeyValuePair, columnKVP As KeyValuePair
    
    content = "<?xml version=" & DQM & "1.0" & _
        DQM & " encoding=" & DQM & "gb2312" & DQM & "?>" & CL_LF
    content = content & "<list>" & CL_LF
    
    Dim dataType As CDataType
    Dim key As String, cellContent As String
    For Each rowKVP In dataDict.KeyValuePairs
        Set columnDict = rowKVP.value
        content = content & S_TAB & "<" & FILE_NAME & " id=" & DQM & columnDict.Items(CStr(0)) & DQM & ">" & CL_LF
        For Each columnKVP In columnDict.KeyValuePairs
            key = columnKVP.key
            Set dataType = dataTypeDict.Item(key)
            cellContent = ModifyCellContentFormat(columnKVP.value, dataType.GetTypeWritor)
            If CellContentIsValid(cellContent, dataType.GetTypeWritor) Then
                If (exporter.ClientExport And dataType.GetClientField) Or _
                    (exporter.ServerExport And dataType.GetServerField) Then
                    content = content & S_TAB & S_TAB & "<" & dataType.GetName & _
                        " type=" & DQM & dataType.GetTypeWritor & DQM & ">" & _
                        cellContent & "</" & dataType.GetName & ">" & CL_LF
                End If
            Else
                UserForm.ResultLabel = "数据错误 行:" & rowKVP.key & " 列:" & columnKVP.key & " 内容:" & cellContent & " 类型:" & dataType.GetTypeWritor
                GenerateXMLContent = False
                Exit Function
            End If
        Next
        content = content & S_TAB & "</" & FILE_NAME & ">" & CL_LF
    Next
    content = content & "</list>"
    exporter.fileContent = content
    GenerateXMLContent = True
End Function

Sub XMLExport(dataTypeDict As Dictionary, dataDict As Dictionary, exporterArray() As CExporter)
    Dim i As Integer
    Dim content As String
    Dim exporter As CExporter
    For i = 0 To UBound(exporterArray)
        Set exporter = exporterArray(i)
        If Not GenerateXMLContent(dataTypeDict, dataDict, exporter) Then
            GoTo Exception
        End If
    Next i
    On Error GoTo Exception
    For i = 0 To UBound(exporterArray)
        Set exporter = exporterArray(i)
        Open exporter.ExportPath For Output As #1
        Print #1, exporter.fileContent;
        Close #1
    Next i
    UserForm.ResultLabel = "生成成功"
    Exit Sub
Exception:
    Close
End Sub

Sub Main(exportClient As Boolean, exportServer As Boolean)
    Call ConstantInitialization
    Dim dataTypeDict As Dictionary
    Set dataTypeDict = ReadType()
    'Call PrintTypeDictionary(dataTypeDict)
    
    Dim dataDict As Dictionary
    Set dataDict = ReadData(dataTypeDict.Count)
    'Call PrintDataDictionary(dataDict)
    
    Dim arrayIndex As Integer
    arrayIndex = 0
    Dim exporterArray() As CExporter
    ReDim exporterArray(2)
    Set exporterArray(arrayIndex) = AllExporter
    If exportClient Then
        arrayIndex = arrayIndex + 1
        Set exporterArray(arrayIndex) = ClientExporter
    End If
    If exportServer Then
        arrayIndex = arrayIndex + 1
        Set exporterArray(arrayIndex) = ServerExporter
    End If
    ReDim Preserve exporterArray(arrayIndex)
    
    Call XMLExport(dataTypeDict, dataDict, exporterArray)
End Sub

'工具类
Function IsInt(str As String) As Boolean
    If IsNumeric(str) And InStr(str, ".") = 0 Then
        IsInt = True
        Exit Function
    End If
    IsInt = False
End Function

Function ModifyCellContentFormat(cellContent As String, dataType As String) As String
    If ((dataType = "array") Or (dataType = "dict")) And (Right(cellContent, 1) = ",") Then
        ModifyCellContentFormat = Left(cellContent, Len(cellContent) - 1)
        Exit Function
    End If
    ModifyCellContentFormat = cellContent
End Function

Sub PrintTypeDictionary(dict As Dictionary)
    Dim oKVP As KeyValuePair
    Dim dataType As CDataType
   For Each oKVP In dict.KeyValuePairs
      Set dataType = oKVP.value
      Debug.Print oKVP.key; " name:"; dataType.GetName; " typeWritor:"; dataType.GetTypeWritor;
      Debug.Print " exportClient:"; dataType.GetClientField; " exportServer:"; dataType.GetServerField
   Next
End Sub

Sub PrintDataDictionary(dict As Dictionary)
    Dim printDict As Dictionary
    Dim oKVP As KeyValuePair
    
    For Each oKVP In dict.KeyValuePairs
        Set printDict = oKVP.value
        printDict.DebugPrint
    Next
End Sub

Sub DataTypeParser(dataType As CDataType, cellContent As String, name As String)
    dataType.LetName = name
    Dim splitIndex As Integer
    splitIndex = InStr(cellContent, ".")
    If splitIndex = 0 Then
        dataType.LetExportClient = True
        dataType.LetExportServer = True
        dataType.LetTypeWritor = cellContent
    Else
        If Right(cellContent, 1) = "c" Then '客户端字段
            dataType.LetExportClient = True
        Else
            dataType.LetExportServer = True
        End If
        dataType.LetTypeWritor = Left(cellContent, Len(cellContent) - 2)
    End If
End Sub
