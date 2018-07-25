Attribute VB_Name = "CommonCheck"
Option Explicit

Sub CheckAll()

'找到sheet1,用与输出校验结果
Dim Ws1 As Worksheet
Set Ws1 = Sheets(1)

'找到sheet2,用于寻找指定工作薄以及校验数据
Dim Ws2 As Worksheet
Set Ws2 = Sheets(2)

'找到sheet3,用于寻找指定的校验规则
Dim Ws3 As Worksheet
Set Ws3 = Sheets(3)

'全表校验时的校验结果输出范围12-15行，1-5列
Dim nRowForOutput, nColForOutput As Integer
nRowForOutput = 12
nColForOutput = 1

'全表校验时sheet的总行数
Dim nRowStart, nRowEnd, nRow As Integer
nRowEnd = Ws2.[A65535].End(xlUp).Row
nRowStart = 4

'校验指定表时的字段的起始列号，当前列号和总列数，
Dim nColStart, nColEnd, nCol As Integer
nColStart = 2

'判断该表是否已经有错误
Dim bIsError As Boolean
bIsError = False

'文件路径前缀
Dim FilePath As String

'工作薄变量
Dim TargetBook As Workbook

For nRow = nRowStart To nRowEnd
    
    '如果是Excel文件，加上路径前缀后打开，否则更改路径路径前缀
    If Ws2.Cells(nRow, 1).Value Like "*.xlsx" Then
        Set TargetBook = Workbooks.Open(FilePath & Ws2.Cells(nRow, 1).Value)
    Else
        FindFilePath Ws2.Cells(nRow, 1).Value, FilePath
    End If
     
    '检查空行空列以及数据类型是否符合要求
    FormatCheck TargetBook
    
    '获得校验文件的总列数
    nColEnd = Ws2.Range("IV" & nRow).End(xlToLeft).Column
    For nCol = nColStart To nColEnd
        
        '获得某个表某字段校验要求
        Dim VecCheckId
        Dim strCheckId As String
        VecCheckId = Split(Ws2.Cells(nRow, nCol), ",")
                  
        Dim i As Integer
        For i = LBound(VecCheckId) To UBound(VecCheckId)
            strCheckId = VecCheckId(i)
            If IsCheckPassed(strCheckId, Ws3, TargetBook, nCol - 1) = False And bIsError = False Then
                Ws1.Cells(nRowForOutput, nColForOutput).Interior.ColorIndex = 3
                Ws1.Cells(nRowForOutput, nColForOutput).Value = Ws2.Cells(nRow, 1).Value
                nColForOutput = nColForOutput + 1
                bIsError = True
            End If
            
            If nColForOutput > 5 Then
                nColForOutput = 1
                nRowForOutput = nRowForOutput + 1
            End If
        Next
    Next
    If nColForOutput > 15 Then
        MsgBox "输出单元格已满，请先修改校验结果"
        Exit For
    End If
    bIsError = False
    
    '保存、关闭工作薄
    TargetBook.Save
    TargetBook.Close
Next
End Sub

Sub CheckPartial()

'找到sheet1,用与输出校验结果
Dim Ws1 As Worksheet
Set Ws1 = Sheets(1)

'找到sheet2,用于寻找指定工作薄以及校验数据
Dim Ws2 As Worksheet
Set Ws2 = Sheets(2)

'找到sheet3,用于寻找指定的校验规则
Dim Ws3 As Worksheet
Set Ws3 = Sheets(3)

'Sheet2的行号和列号，用于获得所校验表的校验要求，同时nCol-1也是所校验表该字段的列号
Dim nRow, nRowStart, nRowEnd, nCol, nColStart, nColEnd As Integer

'Sheet2搜索表名从第4行开始
nRowStart = 4

'条件表从sheet2第3类开始获取所查表各字段校验要求
nColStart = 3

'打开第21行1-5列中的工作薄
Dim nRowInWs1, nColInWs1 As Integer
nRowInWs1 = 21

'所要校验的表的总数
nRowEnd = Ws2.[A65535].End(xlUp).Row

'文件路径前缀
Dim FilePath As String

'工作薄变量
Dim TargetBook As Workbook

For nColInWs1 = 1 To 5
    
    '如果21行的单元格为空，说明后面没有要校验的数据了，结束校验
    If Ws1.Cells(nRowInWs1, nColInWs1).Value = "" Then
        Exit For
    End If
    
    For nRow = nRowStart To nRowEnd
            
        If Ws2.Cells(nRow, 1).Value Like "*_FILE_PATH" Then
            FindFilePath Ws2.Cells(nRow, 1).Value, FilePath
        ElseIf Ws2.Cells(nRow, 1).Value = Ws1.Cells(nRowInWs1, nColInWs1).Value Then
              
            '找到目标工作簿
            Set TargetBook = Workbooks.Open(FilePath & Ws2.Cells(nRow, 1).Value)
                
            '检查空行空列以及数据类型是否符合要求
            FormatCheck TargetBook
            
            '获得校验文件的总列数
            nColEnd = Ws2.Range("IV" & nRow).End(xlToLeft).Column
            For nCol = nColStart To nColEnd
                
                '获得某个表某字段校验要求
                Dim VecCheckId
                Dim strCheckId As String
                VecCheckId = Split(Ws2.Cells(nRow, nCol), ",")
                    
                Dim i As Integer
                For i = LBound(VecCheckId) To UBound(VecCheckId)
                    strCheckId = VecCheckId(i)
                    If IsCheckPassed(strCheckId, Ws3, TargetBook, nCol - 1) = False Then
                        Ws1.Cells(nRowInWs1 + 1, nColInWs1).Interior.ColorIndex = 3
                        Ws1.Cells(nRowInWs1 + 1, nColInWs1).Value = "又双叒叕错了-_-!"
                    Else
                        Ws1.Cells(nRowInWs1 + 1, nColInWs1).Interior.ColorIndex = 4
                        Ws1.Cells(nRowInWs1 + 1, nColInWs1).Value = "恭喜！"
                    End If
                Next
            Next
            
            '保存、关闭工作薄
            TargetBook.Save
            TargetBook.Close
        End If
    Next
Next

End Sub

Sub ClearCheckAllResult()

'清除sheet1 12-15行，1-5列的数据
Dim nRow, nCol As Integer

For nRow = 12 To 15
    For nCol = 1 To 5
        Cells(nRow, nCol).Interior.ColorIndex = 0
        Cells(nRow, nCol).Value = ""
    Next
Next

End Sub

Sub ClearCheckPartialResult()

'清除sheet1 22行，1-5列的数据
Dim nRow, nCol As Integer
nRow = 22

For nCol = 1 To 5
    Cells(nRow, nCol).Interior.ColorIndex = 0
    Cells(nRow, nCol).Value = ""
Next

End Sub

Sub OpenWB1()

'打开21行1列中的Excel工作簿
Dim nRowInWs1, nColInWs1 As Integer
nRowInWs1 = 21
nColInWs1 = 1

'设置Sheet2变量
Dim Ws2 As Worksheet
Set Ws2 = Sheets(2)

Dim nRowEnd, nRow As Integer
nRowEnd = Ws2.[A65535].End(xlUp).Row

'文件路径前缀
Dim FilePath As String

'工作薄变量
Dim TargetBook As Workbook

If Cells(nRowInWs1, nColInWs1).Value = "" Then
    MsgBox "表1内没有配置表名称", vbOKOnly, "你很皮啊"
Else
    For nRow = 4 To nRowEnd
        If Ws2.Cells(nRow, 1).Value Like "*_FILE_PATH" Then
            FilePath = Ws2.Cells(nRow, 1).Value
        ElseIf Ws2.Cells(nRow, 1).Value = Cells(nRowInWs1, nRowInWs1).Value Then
            '找到目标工作簿
            Set TargetBook = Workbooks.Open(FilePath & Ws2.Cells(nRow, 1).Value)
        End If
    Next
End If

End Sub

Sub OpenWB2()

'打开21行2列中的Excel工作簿
Dim nRowInWs1, nColInWs1 As Integer
nRowInWs1 = 21
nColInWs1 = 2

'设置Sheet2变量
Dim Ws2 As Worksheet
Set Ws2 = Sheets(2)

Dim nRowEnd, nRow As Integer
nRowEnd = Ws2.[A65535].End(xlUp).Row

'文件路径前缀
Dim FilePath As String

'工作薄变量
Dim TargetBook As Workbook

If Cells(nRowInWs1, nColInWs1).Value = "" Then
    MsgBox "表1内没有配置表名称", vbOKOnly, "你很皮啊"
Else
    For nRow = 4 To nRowEnd
        If Ws2.Cells(nRow, 1).Value Like "*_FILE_PATH" Then
            FilePath = Ws2.Cells(nRow, 1).Value
        ElseIf Ws2.Cells(nRow, 1).Value = Cells(nRowInWs1, nRowInWs1).Value Then
            '找到目标工作簿
            Set TargetBook = Workbooks.Open(FilePath & Ws2.Cells(nRow, 1).Value)
        End If
    Next
End If

End Sub

Sub OpenWB3()

'打开21行3列中的Excel工作簿
Dim nRowInWs1, nColInWs1 As Integer
nRowInWs1 = 21
nColInWs1 = 3

'设置Sheet2变量
Dim Ws2 As Worksheet
Set Ws2 = Sheets(2)

Dim nRowEnd, nRow As Integer
nRowEnd = Ws2.[A65535].End(xlUp).Row

'文件路径前缀
Dim FilePath As String

'工作薄变量
Dim TargetBook As Workbook

If Cells(nRowInWs1, nColInWs1).Value = "" Then
    MsgBox "表1内没有配置表名称", vbOKOnly, "你很皮啊"
Else
    For nRow = 4 To nRowEnd
        If Ws2.Cells(nRow, 1).Value Like "*_FILE_PATH" Then
            FilePath = Ws2.Cells(nRow, 1).Value
        ElseIf Ws2.Cells(nRow, 1).Value = Cells(nRowInWs1, nRowInWs1).Value Then
            '找到目标工作簿
            Set TargetBook = Workbooks.Open(FilePath & Ws2.Cells(nRow, 1).Value)
        End If
    Next
End If

End Sub

Sub OpenWB4()

'打开21行4列中的Excel工作簿
Dim nRowInWs1, nColInWs1 As Integer
nRowInWs1 = 21
nColInWs1 = 4

'设置Sheet2变量
Dim Ws2 As Worksheet
Set Ws2 = Sheets(2)

Dim nRowEnd, nRow As Integer
nRowEnd = Ws2.[A65535].End(xlUp).Row

'文件路径前缀
Dim FilePath As String

'工作薄变量
Dim TargetBook As Workbook

If Cells(nRowInWs1, nColInWs1).Value = "" Then
    MsgBox "表1内没有配置表名称", vbOKOnly, "你很皮啊"
Else
    For nRow = 4 To nRowEnd
        If Ws2.Cells(nRow, 1).Value Like "*_FILE_PATH" Then
            FilePath = Ws2.Cells(nRow, 1).Value
        ElseIf Ws2.Cells(nRow, 1).Value = Cells(nRowInWs1, nRowInWs1).Value Then
            '找到目标工作簿
            Set TargetBook = Workbooks.Open(FilePath & Ws2.Cells(nRow, 1).Value)
        End If
    Next
End If

End Sub

Sub OpenWB5()

'打开21行5列中的Excel工作簿
Dim nRowInWs1, nColInWs1 As Integer
nRowInWs1 = 21
nColInWs1 = 5

'设置Sheet2变量
Dim Ws2 As Worksheet
Set Ws2 = Sheets(2)

Dim nRowEnd, nRow As Integer
nRowEnd = Ws2.[A65535].End(xlUp).Row

'文件路径前缀
Dim FilePath As String

'工作薄变量
Dim TargetBook As Workbook

If Cells(nRowInWs1, nColInWs1).Value = "" Then
    MsgBox "表1内没有配置表名称", vbOKOnly, "你很皮啊"
Else
    For nRow = 4 To nRowEnd
        If Ws2.Cells(nRow, 1).Value Like "*_FILE_PATH" Then
            FilePath = Ws2.Cells(nRow, 1).Value
        ElseIf Ws2.Cells(nRow, 1).Value = Cells(nRowInWs1, nRowInWs1).Value Then
            '找到目标工作簿
            Set TargetBook = Workbooks.Open(FilePath & Ws2.Cells(nRow, 1).Value)
        End If
    Next
End If

End Sub

Sub CheckAllResourceExist()

'设置Sheet1，Sheet4变量
Dim Ws1, Ws4 As Worksheet
Set Ws1 = Sheets(1)
Set Ws4 = Sheets(4)

Dim nRowStart, nRowEnd, nRow, nCol, nColStart, nColEnd, nRowForOutput, nColForOutput As Integer

'含有资源的策划包路径始于第3行，止于第nRowEnd行
nRowStart = 3
nRowEnd = Ws4.[A65535].End(xlUp).Row

'从所校验表的第2列开始搜索资源路径字段
nColStart = 2

'结束输出在表1的30-32行的第1-5列
nRowForOutput = 30
nColForOutput = 1

For nRow = nRowStart To nRowEnd
    
    '打开对应工作簿
    Dim TargetBook As Workbook
    Set TargetBook = Workbooks.Open(Ws4.Cells(nRow, 2).Value)
    nColEnd = TargetBook.Sheets(1).Range("IV1").End(xlToLeft).Column
    
    For nCol = nColStart To nColEnd
        If TargetBook.Sheets(1).Cells(1, nCol).Value Like "*Path" Then
            Ws1.Cells(nRowForOutput, nColForOutput) = Ws4.Cells(nRow, 1).Value
            Ws1.Cells(nRowForOutput, nColForOutput).Interior.ColorIndex = 4
            
            '工作薄共有nRowEndInRes行数据
            Dim nRowInRes, nRowEndInRes As Integer
            nRowEndInRes = TargetBook.Sheets(1).[A65535].End(xlUp).Row
            
            Dim DirName As String
            For nRowInRes = 10 To nRowEndInRes
                
                '如果路径字段不为空，说明此行数据为资源，需要查找该路径下文件是否存在
                If TargetBook.Sheets(1).Cells(nRowInRes, nCol).Value <> "" Then
                    DirName = Dir(RESOURCE_FILE_PATH & TargetBook.Sheets(1).Cells(nRowInRes, nCol).Value & ".*")
                    
                    If DirName = "" Then
                        TargetBook.Sheets(1).Cells(nRowInRes, 1).Interior.ColorIndex = 3
                        Ws1.Cells(nRowForOutput, nColForOutput).Interior.ColorIndex = 3
                    End If
                End If
            Next
            '满5列换行继续输出
            nColForOutput = nColForOutput + 1
            If nColForOutput = 6 Then
                nRowForOutput = nRowForOutput + 1
                nColForOutput = 1
            End If
            Exit For
        End If
    Next
    TargetBook.Save
    TargetBook.Close
    
    '输出只能到32行，到33行了说明需要扩展输出结果单元格
    If nRowForOutput = 33 Then
        TargetBook.Save
        TargetBook.Close
        Exit For
    End If
Next

End Sub

Sub ClearCheckAllResourceExist()

'设置Sheet1变量
Dim Ws As Worksheet
Set Ws = Sheets(1)

Dim nRow, nRowStart, nRowEnd, nCol, nColStart, nColEnd As Integer
'清空Sheet1 30-32行1-5列的内容
nRowStart = 30
nRowEnd = 32
nColStart = 1
nColEnd = 5

For nRow = nRowStart To nRowEnd
    For nCol = nColStart To nColEnd
        Ws.Cells(nRow, nCol).Interior.ColorIndex = 0
        Ws.Cells(nRow, nCol).Value = ""
    Next
Next
    
End Sub
