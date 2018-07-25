Attribute VB_Name = "Baselib"
Option Explicit
Public Function IsCheckPassed(ByVal CheckId As String, ByVal CheckSheet As Worksheet, ByRef TargetBook As Workbook, ByVal nCol As Integer)
'判断该行某字段是否满足校验要求

Dim nRow, nRowStart, nRowEnd As Integer
nRowStart = 3
nRowEnd = CheckSheet.[A65535].End(xlUp).Row

IsCheckPassed = True

If CheckId <> "" Then
    For nRow = nRowStart To nRowEnd
        If CheckSheet.Cells(nRow, 1).Value = CheckId Then
            '获取校验要求类型和参数
            Dim CheckType As Integer
            Dim CheckData As String
        
            CheckType = CheckSheet.Cells(nRow, 3).Value
            CheckData = CheckSheet.Cells(nRow, 4).Value
            
            Dim vecCheckData() As String
            Dim vecCol() As Integer
            Dim vecDataArg() As String
            Dim vecDataDep() As String
        
            Select Case CheckType
            Case 1
                vecCheckData = Split(CheckData, MARK_COMMA)
                IsCheckPassed = CheckBlankForSpecificCol(TargetBook, nCol, vecCheckData)
            Case 2
                IsCheckPassed = CheckIdFormat(TargetBook, nCol, CheckData)
            Case 3
                vecCheckData = Split(CheckData, MARK_SEMICOLON)
                vecCol = Split(vecCheckData(0), MARK_COMMA)
                vecDataArg = Split(vecCheckData(1), MARK_COMMA)
                vecDataDep = Split(vecCheckData(2), MARK_COMMA)
                IsCheckPassed = CheckDependenceRelation(TargetBook, CInt(vecCol(0)), vecDataArg, CInt(vecCol(1)), vecDataDep)
            Case 4
                vecCheckData = Split(CheckData, MARK_SEMICOLON)
                vecDataArg = Split(vecCheckData(0), MARK_COMMA)
                vecDataDep = Split(vecCheckData(1), MARK_COMMA)
                IsCheckPassed = CheckMutexRelation(TargetBook, nCol, vecDataArg, vecDataDep)
            Case 5
                vecCheckData = Split(CheckData, MARK_SEMICOLON)
                vecDataArg = Split(vecCheckData(0), MARK_COMMA)
                vecDataDep = Split(vecCheckData(1), MARK_COMMA)
                IsCheckPassed = CheckDataConsistence(TargetBook, CInt(vecDataArg(0)), vecDataArg(1), CInt(vecDataDep(0)), vecDataDep(1))
            Case 6
                vecCheckData = Split(CheckData, MARK_SEMICOLON)
                vecDataArg = Split(vecCheckData(0), MARK_COMMA)
                vecDataDep = Split(vecCheckData(1), MARK_COMMA)
                IsCheckPassed = CheckSplit(TargetBook, nCol, vecDataArg, vecDataDep)
            Case 7
                Dim ResWorkBook As Workbook
                FindWorkBook ResWorkBook, CheckData
                IsCheckPassed = CheckIdExist(TargetBook, nCol, ResWorkBook)
            Case 8
                IsCheckPassed = CheckIdUnique(TargetBook, nCol)
            End Select
        End If
    Next
End If

End Function

Public Function FormatCheck(ByRef TargetBook As Workbook) As Boolean
'格式检查，空行空列，数据类型

If CheckBlankForRowAndCol(TargetBook) = False Then
    FormatCheck = False
End If

If CheckDataType(TargetBook) = False Then
    FormatCheck = False
End If

End Function

Private Function CheckBlankForRowAndCol(ByRef TargetBook As Workbook) As Boolean
'校验是否存在空行空列

'配置表目前都只有Sheet1，其实直接用Sheet(1)也是可以的
Dim Ws1 As Worksheet
Set Ws1 = TargetBook.Sheets(1)

Dim nRow, nCol, nRowEnd, nColEnd As Integer
nRowEnd = Ws1.[A65535].End(xlUp).Row
nColEnd = Ws1.[IV1].End(xlToLeft).Column

For nRow = 1 To nRowEnd
    If IsEmpty(Ws1.Cells(nRow, 1)) Then
        CheckBlankForRowAndCol = False
    End If
Next

For nCol = 1 To nColEnd
    If IsEmpty(Ws1.Cells(1, nCol)) Then
        CheckBlankForRowAndCol = False
    End If
Next

End Function

Private Function CheckBlankForSpecificCol(ByRef TargetBook As Workbook, ByVal nCol As Integer, ByRef VecExemptWords() As String) As Boolean
'校验指定列是否存在空单元格,最后一个是豁免字符串数组，当配表Id含有此其中的字符串时，允许存在空单元格

Dim Ws1 As Worksheet
Set Ws1 = TargetBook.Sheets(1)

Dim nRow, nRowEnd As Integer
nRowEnd = Ws1.[A65535].End(xlUp).Row

CheckBlankForSpecificCol = True

For nRow = 10 To nRowEnd
    If IsNecessary(Cells(nRow, 1).Value, VecExemptWords) = True Then
        If IsEmpty(Ws1.Cells(nRow, nCol)) Then
            Ws1.Cells(nRow, nCol).Interior.ColorIndex = 3
            CheckBlankForSpecificCol = False
        End If
    End If
Next

End Function

Private Function IsNecessary(ByVal StrConfId As String, ByRef VecExemptWords() As String) As Boolean
'判断该Config在制定字段是否是必填的

IsNecessary = True

Dim i As Integer
For i = LBound(VecExemptWords) To UBound(VecExemptWords)
    If InStr(StrConfId, VecExemptWords(i)) <> 0 Then
        IsNecessary = False
        Exit For
    End If
Next

End Function

Private Function CheckDataType(ByRef TargetBook As Workbook) As Boolean
'校验表内数据类型

Dim Ws1 As Worksheet
Set Ws1 = TargetBook.Sheets(1)

CheckDataType = True

Dim nColTotal, nCol
nColTotal = Ws1.[IV1].End(xlToLeft).Column

For nCol = 2 To nColTotal
     If CheckDataTypeForSpecificCol(TargetBook, nCol) = False Then
        CheckDataType = False
    End If
Next

End Function


Private Function CheckDataTypeForSpecificCol(ByRef TargetBook As Workbook, ByVal nCol As Integer) As Boolean
'校验指定列数据类型

Dim Ws1 As Worksheet
Set Ws1 = TargetBook.Sheets(1)

CheckDataTypeForSpecificCol = True

'取该列规定的数据类型，目前只有int和string
Dim ConfigData As String
ConfigData = Ws1.Cells(2, nCol).Value

If ConfigData = "int" Then
    If CheckIntType(TargetBook, nCol) = False Then
        CheckDataTypeForSpecificCol = False
    End If
ElseIf ConfigData = "string" Then
    If CheckStringType(TargetBook, nCol) = False Then
        CheckDataTypeForSpecificCol = False
    End If
ElseIf ConfigData = "float" Then
    If CheckFloatType(TargetBook, nCol) = False Then
        CheckDataTypeForSpecificCol = False
    End If
End If

End Function

Private Function CheckIntType(ByRef TargetBook As Workbook, ByVal nCol As Integer) As Boolean
'校验int型数据

Dim Ws1 As Worksheet
Set Ws1 = TargetBook.Sheets(1)

Dim nRow, nRowEnd As Integer
nRowEnd = Ws1.[A65535].End(xlUp).Row

CheckIntType = True

For nRow = 10 To nRowEnd
    If IsEmpty(Ws1.Cells(nRow, nCol)) = False Then
        Ws1.Cells(nRow, nCol).Value = Str(Ws1.Cells(nRow, nCol).Value)
        If Ws1.Cells(nRow, nCol).Value Like "*.*" Then
            Ws1.Cells(nRow, nCol).Interior.ColorIndex = 3
            CheckIntType = False
        Else
            Ws1.Cells(nRow, nCol).Value = CInt(Ws1.Cells(nRow, nCol).Value)
        End If
    End If
Next

End Function

Private Function CheckStringType(ByRef TargetBook As Workbook, ByVal nCol As Integer) As Boolean
'校验string型数据

Dim Ws1 As Worksheet
Set Ws1 = TargetBook.Sheets(1)

Dim nRow, nRowEnd As Integer
nRowEnd = Sheets(1).[A65535].End(xlUp).Row

CheckStringType = True

For nRow = 10 To nRowEnd
    If IsEmpty(Ws1.Cells(nRow, nCol)) = False Then
        If TypeName(Ws1.Cells(nRow, nCol).Value) <> "String" Then
            Ws1.Cells(nRow, nCol).Interior.ColorIndex = 3
            CheckStringType = False
        End If
    End If
Next

End Function

Private Function CheckFloatType(ByRef TargetBook As Workbook, ByVal nCol As Integer) As Boolean
'校验float型数据

Dim Ws1 As Worksheet
Set Ws1 = TargetBook.Sheets(1)

Dim nRow, nRowEnd As Integer
nRowEnd = Sheets(1).[A65535].End(xlUp).Row

CheckFloatType = True

For nRow = 10 To nRowEnd
    If IsEmpty(Ws1.Cells(nRow, nCol)) = False Then
        If TypeName(Ws1.Cells(nRow, nCol).Value) <> "Double" Then
            Ws1.Cells(nRow, nCol).Interior.ColorIndex = 3
            CheckFloatType = False
        End If
    End If
Next

End Function

Private Function CheckIdFormat(ByRef TargetBook As Workbook, ByVal nCol As Integer, ByVal IdReq As String)
'校验Id是否是否包含指定字符串
    
Dim Ws1 As Worksheet
Set Ws1 = TargetBook.Sheets(1)
    
Dim nRow, nRowEnd As Integer
nRowEnd = Ws1.[A65535].End(xlUp).Row

CheckIdFormat = True

For nRow = 10 To nRowEnd
    If Ws1.Cells(nRow, nCol).Value Like "IdReq*" = False Then
        Ws1.Cells(nRow, nCol).Interior.ColorIndex = 3
        CheckIdFormat = False
    End If
Next

End Function

Private Function CheckDependenceRelation(ByRef TargetBook As Workbook, ByVal nColArg As Integer, ByRef vecDataArg() As String, ByVal nColDep As Integer, ByRef vecDataDep() As String) As Boolean
'校验策划表两个字段的依赖关系是否满足

Dim Ws1 As Worksheet
Set Ws1 = TargetBook.Sheets(1)

Dim nRow, nRowEnd, i As Integer
nRowEnd = Ws1.[A65535].End(xlUp).Row

CheckDependenceRelation = True

For nRow = 10 To nRowEnd
    If IsEmpty(Ws1.Cells(nRow, nColArg)) = False And IsEmpty(Ws1.Cells(nRow, nColArg)) = False Then
        For i = 0 To UBound(vecDataArg)
        If vecDataArg(i) = Ws1.Cells(nRow, nColArg) Then
            If Ws1.Cells(nRow, nColDep) <> vecDataDep(i) Then
                Ws1.Cells(nRow, nColDep).Interior.ColorIndex = 3
                CheckDependenceRelation = False
            End If
        End If
    Next
    End If
Next

End Function

Private Function CheckMutexRelation(ByRef TargetBook As Workbook, ByVal nCol As Integer, ByRef vecDataArg() As String, ByRef vecDataDep() As String) As Boolean
'校验策划表两个字段的互斥关系是否满足

Dim Ws1 As Worksheet
Set Ws1 = TargetBook.Sheets(1)

Dim nRow, nRowEnd, i As Integer
nRowEnd = Ws1.[A65535].End(xlUp).Row

CheckMutexRelation = True

For nRow = 10 To nRowEnd
    If IsEmpty(Ws1.Cells(nRow, nCol)) = False Then
        Dim vecSplitData() As String
        vecSplitData = Split(Ws1.Cells(nRow, nCol), MARK_COMMA)
        For i = 0 To UBound(vecDataArg)
            If vecSplitData(0) = vecDataArg(i) And vecSplitData(0) <> vecDataDep(i) Then
                Ws1.Cells(nRow, nCol).Interior.ColorIndex = 3
                CheckMutexRelation = False
            End If
        Next
    End If
Next
End Function

Private Function CheckDataConsistence(ByRef TargetBook As Workbook, ByVal nColArg As Integer, ByVal DelimiterArg As Integer, ByVal nColDep As Integer, ByVal DelimiterDep As Integer) As Boolean
'校验策划表两个字段的多段数据的数据一致性

Dim Ws1 As Worksheet
Set Ws1 = TargetBook.Sheets(1)

Dim nRow, nRowEnd, i As Integer
nRowEnd = Ws1.[A65535].End(xlUp).Row

CheckDataConsistence = True

For nRow = 10 To nRowEnd
    If IsEmpty(Ws1.Cells(nRow, nColArg)) = False And IsEmpty(Ws1.Cells(nRow, nColArg)) = False Then
        Dim nCountArg, nCountDep As Integer
        Dim vecDataArg, vecDataDep
        vecDataArg = Split(Ws1.Cells(nRow, nColArg), DelimiterArg)
        vecDataDep = Split(Ws1.Cells(nRow, nColDep), DelimiterDep)
        nCountArg = UBound(vecDataArg) - LBound(vecDataArg) + 1
        nCountDep = UBound(vecDataDep) - LBound(vecDataDep) + 1
        If nCountArg <> nCountDep Then
            Ws1.Cells(nRow, nColDep).Interior.ColorIndex = 3
            CheckDataConsistence = False
        End If
    End If
Next
End Function

Private Function CheckSplit(ByRef TargetBook As Workbook, ByVal nCol As Integer, ByRef vecDelimiter() As String, ByRef vecSplitCount() As String) As Boolean
'校验数据分隔格式

Dim Ws1 As Worksheet
Set Ws1 = TargetBook.Sheets(1)

Dim nRow, nRowEnd As Integer
nRowEnd = Ws1.[A65535].End(xlUp).Row

CheckSplit = True

For nRow = 10 To nRowEnd
    If IsEmpty(Ws1.Cells(nRow, nCol)) = False Then
        Dim vecSplitData
        Dim nCount As Integer
        vecSplitData = Split(Ws1.Cells(nRow, nCol), vecDelimiter(0))
        nCount = UBound(vecSplitData) - LBound(vecSplitData) + 1
    
        If nCount = CInt(vecSplitCount(0)) Then
            If UBound(vecDelimiter) = 1 Then
                Dim i As Integer
                Dim vecSplitSecond
                For i = 0 To UBound(vecSplitData)
                    vecSplitSecond = Split(vecSplitData(i), vecDelimiter(1))
                    If UBound(vecSplitSecond) - LBound(vecSplitSecond) + 1 <> CInt(vecSplitCount(1)) Then
                        Ws1.Cells(nRow, nCol).Interior.ColorIndex = 3
                        CheckSplit = False
                    End If
                Next
            End If
        Else
            Ws1.Cells(nRow, nCol).Interior.ColorIndex = 3
            CheckSplit = False
        End If
    End If
Next

End Function

Private Function CheckIdExist(ByRef TargetBook As Workbook, ByVal nCol As Integer, ByRef ResBook As Workbook)
'校验Id是否存在

Dim Ws1 As Worksheet
Set Ws1 = TargetBook.Sheets(1)

Dim Ws1InRes As Worksheet
Set Ws1InRes = ResBook.Sheets(1)

Dim nRow, nRowStart, nRowEnd As Integer
nRowEnd = Ws1.[A65535].End(xlUp).Row
nRowStart = 10

CheckIdExist = True

For nRow = nRowStart To nRowEnd
    Dim nRow1, nRowEnd1 As Integer
    nRowEnd1 = ResBook.Sheets(1).[A65535].End(xlUp).Row
    
    For nRow1 = 10 To nRowEnd1
        If Ws1InRes.Cells(nRow1, 1).Value = Ws1.Cells(nRow, nCol).Value Then GoTo ForEnd
    Next
    Ws1.Cells(nRow, nCol).Interior.ColorIndex = 3
    CheckIdExist = False
ForEnd:
Next

End Function

Private Function CheckIdUnique(ByRef TargetBook As Workbook, ByVal nCol As Integer)
'校验Id是否唯一

Dim dic
Set dic = CreateObject("Scripting.Dictionary")

Dim Ws1 As Worksheet
Set Ws1 = TargetBook.Sheets(1)

Dim nRow, nRowEnd As Integer
nRowEnd = Ws1.[A65535].End(xlUp).Row

CheckIdUnique = True

For nRow = 10 To nRowEnd
    If dic(Ws1.Cells(nRow, nCol).Value) = 1 Then
        Ws1.Cells(nRow, nCol).Interior.ColorIndex = 3
        CheckIdUnique = False
    Else
        dic(Ws1.Cells(nRow, nCol).Value) = 1
    End If
Next

End Function

Public Function FindWorkBook(ByRef TargetBook As Workbook, ByVal BookName As String)
'查找指定工作薄

Dim ContentBook As Workbook
Set ContentBook = Workbooks.Open(CONFIG_WORKBOOK)

Dim Ws2 As Worksheet
Set Ws2 = ContentBook.Sheets(2)

Dim nRow, nRowEnd As Integer
nRowEnd = Ws2.[A65535].End(xlUp).Row

Dim FilePath As String

For nRow = 4 To nRowEnd
    If Ws2.Cells(nRow, 1).Value Like "*_File_PATH" Then
        FindFilePath Ws2.Cells(nRow, 1).Value, FilePath
    End If
    
    If Ws2.Cells(nRow, 1).Value = BookName Then
        Set TargetBook = Workbooks.Open(FilePath & Ws2.Cells(nRow, 1).Value)
        Exit For
    End If
Next
End Function

Public Function FindFilePath(ByVal PathName As String, ByRef PathPrfix As String)

If PathName = "COMMON_FILE_PATH" Then
    PathPrfix = COMMON_FILE_PATH
ElseIf PathName = "CLIENT_FILE_PATH" Then
    PathPrfix = CLIENT_FILE_PATH
ElseIf PathName = "SERVER_FILE_PATH" Then
    PathPrfix = "SERVER_FILE_PATH"
End If

End Function
