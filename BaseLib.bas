Attribute VB_Name = "Baselib"
Option Explicit
Public Function IsCheckPassed(ByVal CheckId As String, ByVal CheckSheet As Worksheet, ByRef TargetBook As Workbook, ByVal nCol As Integer)
'�жϸ���ĳ�ֶ��Ƿ���סУ��Ҫ��

Dim nRowTotal, nRow, nRowStart As Integer
nRowTotal = CheckSheet.[A65535].End(xlUp).Row
nRowStart = 3

If CheckId <> "" Then
    For nRow = nRowStart To nRowTotal
        If CheckSheet.Cells(nRow, 1).Value = CheckId Then
            '��ȡУ��Ҫ�����ͺͲ���
            Dim CheckType As Integer
            Dim CheckData As String
        
            CheckType = CheckSheet.Cells(nRow, 4).Value
            CheckData = CheckSheet.Cells(nRow, 5).Value
            
            Dim vecFiledData, vecCol
            Dim vecData() As String
            Dim vecData2() As String
        
            Select Case CheckType
            Case 1
                '�и�У�����
                Dim VecExemptWords() As String
                VecExemptWords = Split(CheckData, MARK_COMMA)
                IsCheckPassed = CheckBlankForSpecificCol(TargetBook, nCol, VecExemptWords)
            Case 2
                IsCheckPassed = CheckIdFormat(TargetBook, nCol, CheckData)
            Case 3
                vecFiledData = Split(CheckData, MARK_SEMICOLON)
                vecCol = Split(vecFiledData(0), MARK_COMMA)
                vecData = Split(vecFiledData(1), MARK_COMMA)
                vecData2 = Split(vecFiledData(2), MARK_COMMA)
                IsCheckPassed = CheckDependenceRelation(TargetBook, CInt(vecCol(0)), CInt(vecCol(1)), vecData, vecData2)
            Case 4
                Dim vecMutexData() As Integer
                vecMutexData = Split(CheckData, MARK_SEMICOLON)
                IsCheckPassed = CheckMutexRelation(TargetBook, nCol, vecMutexData)
            Case 5
                vecFiledData = Split(CheckData, MARK_SEMICOLON)
                vecData = Split(vecFiledData(0), MARK_COMMA)
                vecData2 = Split(vecFiledData(1), MARK_COMMA)
                IsCheckPassed = CheckDataConsistence(TargetBook, CInt(vecData(0)), vecData(1), CInt(vecData2(0)), vecData2(1))
            Case 6
                vecFiledData = Split(CheckData, MARK_SEMICOLON)
                IsCheckPassed = True
                
                
            End Select
            Exit For
        End If
    Next
End If

End Function


Public Function FormatCheck(ByRef TargetBook As Workbook) As Boolean

'��ʽ��飬���п��У���������
If CheckBlankForRowAndCol(TargetBook) = False Then
    FormatCheck = True
End If

If CheckDataType(TargetBook) = False Then
    FormatCheck = True
End If

End Function


Private Function CheckBlankForRowAndCol(ByRef TargetBook As Workbook) As Boolean
'У���Ƿ���ڿ��п���

Dim Ws As Worksheet
'���ñ�Ŀǰ��ֻ��Sheet1����ʵֱ����Sheet(1)Ҳ�ǿ��Ե�
Set Ws = TargetBook.Sheets(1)

Dim nRow, nCol, nRowEnd, nColEnd As Integer
nRowEnd = Ws.[A65535].End(xlUp).Row
nColEnd = Ws.[IV1].End(xlToLeft).Column


For nRow = 1 To nRowEnd
    If IsEmpty(Ws.Cells(nRow, 1)) Then
        CheckBlankForRowAndCol = False
    End If
Next

For nCol = 1 To nColEnd
    If IsEmpty(Ws.Cells(1, nCol)) Then
        CheckBlankForRowAndCol = False
    End If
Next

End Function

Private Function CheckBlankForSpecificCol(ByRef TargetBook As Workbook, ByVal nCol As Integer, ByRef VecExemptWords() As String) As Boolean
'У��ָ�����Ƿ���ڿյ�Ԫ��,���һ���ǻ����ַ������飬�����Id���д����е��ַ���ʱ��������ڿյ�Ԫ��

Dim Ws As Worksheet
Set Ws = TargetBook.Sheets(1)

Dim nRow, nRowEnd As Integer
nRowEnd = Ws.[A65535].End(xlUp).Row

CheckBlankForSpecificCol = True

For nRow = 10 To nRowEnd
    If IsNecessary(Cells(nRow, 1).Value, VecExemptWords) = True Then
        If IsEmpty(Ws.Cells(nRow, nCol)) Then
            Ws.Cells(nRow, nCol).Interior.ColorIndex = 3
            CheckBlankForSpecificCol = False
        End If
    End If
Next

End Function

Private Function IsNecessary(ByVal StrConfId As String, ByRef VecExemptWords() As String) As Boolean
'�жϸ�Config���ƶ��ֶ��Ƿ��Ǳ����

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
'У�������������

Dim Ws As Worksheet
Set Ws = TargetBook.Sheets(1)

Dim nColTotal, nCol
nColTotal = Ws.[IV1].End(xlToLeft).Column

For nCol = 2 To nColTotal
     If CheckDataTypeForSpecificCol(TargetBook, nCol) = False Then
        CheckDataType = False
    End If
Next

End Function


Private Function CheckDataTypeForSpecificCol(ByRef TargetBook As Workbook, ByVal nCol As Integer) As Boolean
'У��ָ������������

Dim Ws As Worksheet
Set Ws = TargetBook.Sheets(1)

'ȡ���й涨���������ͣ�Ŀǰֻ��int��string
Dim ConfigData As String
ConfigData = Ws.Cells(2, nCol).Value

If ConfigData = "int" Then
    If CheckIntType(TargetBook, nCol) = False Then
        CheckDataTypeForSpecificCol = False
    End If
ElseIf ConfigData = "string" Then
    If CheckStringType(TargetBook, nCol) = False Then
        CheckDataTypeForSpecificCol = False
    End If
End If

End Function

Private Function CheckIntType(ByRef TargetBook As Workbook, ByVal nCol As Integer) As Boolean
'У��int������

Dim Ws As Worksheet
Set Ws = TargetBook.Sheets(1)

Dim nRow, nRowEnd As Integer
nRowEnd = Ws.[A65535].End(xlUp).Row

For nRow = 10 To nRowEnd
    If IsEmpty(Ws.Cells(nRow, nCol)) = False Then
        Ws.Cells(nRow, nCol).Value = Str(Ws.Cells(nRow, nCol).Value)
        If Ws.Cells(nRow, nCol).Value Like "*.*" Then
            Ws.Cells(nRow, nCol).Interior.ColorIndex = 3
            CheckIntType = False
        Else
            Ws.Cells(nRow, nCol).Value = CInt(Ws.Cells(nRow, nCol).Value)
        End If
    End If
Next

End Function

Private Function CheckStringType(ByRef TargetBook As Workbook, ByVal nCol As Integer) As Boolean
'У��string������

Dim Ws As Worksheet
Set Ws = TargetBook.Sheets(1)

Dim nRow, nRowEnd As Integer
nRowEnd = Sheets(1).[A65535].End(xlUp).Row

For nRow = 10 To nRowEnd
    If IsEmpty(Ws.Cells(nRow, nCol)) = False Then
        If TypeName(Ws.Cells(nRow, nCol).Value) <> "String" Then
            Ws.Cells(nRow, nCol).Interior.ColorIndex = 3
            CheckStringType = False
        End If
    End If
Next

End Function

Private Function CheckIdFormat(ByRef TargetBook As Workbook, ByVal nCol As Integer, ByVal IdReq As String)
    
Dim Ws As Worksheet
Set Ws = TargetBook.Sheets(1)
    
Dim nRow, nRowEnd As Integer
nRowEnd = Ws.[A65535].End(xlUp).Row

CheckIdFormat = True

For nRow = 10 To nRowEnd
    If Ws.Cells(nRow, nCol).Value Like "IdReq*" = False Then
        Ws.Cells(nRow, nCol).Interior.ColorIndex = 3
        CheckIdFormat = False
    End If
Next

End Function

Private Function CheckDependenceRelation(ByRef TargetBook As Workbook, ByVal FieldValue As Integer, ByVal FieldValue2 As Integer, ByRef vecData() As String, ByRef vecData2() As String) As Boolean

Dim Ws As Worksheet
Set Ws = TargetBook.Sheets(1)

Dim nRow, nRowEnd, i As Integer
nRowEnd = Ws.[A65535].End(xlUp).Row

CheckDependenceRelation = True

For nRow = 10 To nRowEnd
    For i = 0 To UBound(vecData)
        If vecData(i) = Ws.Cells(nRow, FieldValue) Then
            If Ws.Cells(nRow, FieldValue) <> vecData2(i) Then
                Ws.Cells(nRow, FieldValue2).Interior.ColorIndex = 3
                CheckDependenceRelation = False
            End If
        End If
    Next
Next

End Function

Private Function CheckMutexRelation(ByRef TargetBook As Workbook, ByVal nCol As Integer, ByRef vecMutexData) As Boolean
    
Dim Ws As Worksheet
Set Ws = TargetBook.Sheets(1)

Dim nRow, nRowEnd, i As Integer
nRowEnd = Ws.[A65535].End(xlUp).Row

CheckMutexRelation = True
For nRow = 10 To nRowEnd
    For i = 0 To UBound(vecMutexData)
        If Ws.Cells(nRow, nCol) = vecMutexData(i) Then
            Ws.Cells(nRow, nCol).Interior.ColorIndex = 3
            CheckMutexRelation = False
        End If
    Next
Next
End Function

Private Function CheckDataConsistence(ByRef TargetBook As Workbook, ByVal nCol As Integer, ByVal Delimiter As Integer, ByVal nCol2 As Integer, ByVal Delimiter2 As Integer) As Boolean

Dim Ws As Worksheet
Set Ws = TargetBook.Sheets(1)

Dim nRow, nRowEnd, i As Integer
nRowEnd = Ws.[A65535].End(xlUp).Row

CheckDataConsistence = True
For nRow = 10 To nRowEnd
    Dim nCount, nCount2 As Integer
    Dim vecData, vecData2
    vecData = Split(Ws.Cells(nRow, nCol), Delimiter)
    vecData2 = Split(Ws.Cells(nRow, nCol), Delimiter2)
    nCount = UBound(vecData) - LBound(vecData) + 1
    nCount2 = UBound(vecData2) - LBound(vecData2) + 1
    If nCount <> nCount2 Then
        Ws.Cells(nRow, nCol).Interior.ColorIndex = 3
        CheckDataConsistence = False
    End If
Next
End Function

Private Function CheckSplit(ByRef TargetBook As Workbook, ByVal nCol As Integer, ByRef vecFiledData() As String) As Boolean

Dim Ws As Worksheet
Set Ws = TargetBook.Sheets(1)

Dim nRow, nRowEnd, i As Integer
nRowEnd = Ws.[A65535].End(xlUp).Row

CheckSplit = True
For nRow = 10 To nRowEnd
    Dim vecSplitData
    Dim nCount As Integer
    vecSplitData = Split(Ws.Cells(nRow, nCol), MARK_COMMA)
    nCount = UBound(vecSplitData) - LBound(vecSplitData) + 1
    If nCount <> vecFiledData(1) Then
        Ws.Cells(nRow, nCol).Interior.ColorIndex = 3
        CheckSplit = False
    End If
    
    If CInt(vecFiledData(1)) = 2 Then
        Dim i As Integer
        For i = 0 To UBound(vecSplitData)
            Dim vecSplitData2
            vecSplitData2 = Split(vecSplitData(i))
            nCount = UBound(vecSplitData2) - LBound(vecSplitData2) + 1
            
            If nCount <> vecFiledData(1) Then
                Ws.Cells(nRow, nCol).Interior.ColorIndex = 3
                CheckSplit = False
            End If

        Next
    End If
Next

End Function
