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

'全表校验时的校验结果输出范围
Dim nRowForOutput, nColForOutput As Integer
nRowForOutput = 12
nColForOutput = 1

'全表校验时sheet的总行数
Dim nRowStart, nRowEnd, nRow As Integer
nRowEnd = Ws2.[A65535].End(xlUp).Row
nRowStart = 4

'校验指定表时的字段的起始列号，当前列号和总列数，
Dim nColStart, nColEnd, nCol As Integer
nColStart = 3

'判断该表是否已经有错误
Dim bIsError As Boolean
bIsError = False

For nRow = nRowStart To nRowEnd
    nColEnd = Ws2.Range("IV" & nRow).End(xlToLeft).Column
    
    '找到目标工作簿,在sheet2的第nROW行第2列
    Dim TargetBook As Workbook
    Set TargetBook = Workbooks.Open(Ws2.Cells(nRow, 2).Value)
     
    '检查空行空列以及数据类型是否符合要求
    FormatCheck TargetBook
    
    For nCol = nColStart To nColEnd
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
    If nColForOutput = 16 Then
        MsgBox "输出单元格已满，请先修改校验结果"
        Exit For
    End If
    bIsError = False
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

'打开21行1-5列的工作簿
Dim nRow, nRowStart, nRowEnd, nCol, nColStart, nColEnd As Integer
nRow = 21

'Sheet2搜索表名从第4行开始
nRowStart = 4

'条件表从sheet2第3类开始获取所查表各字段校验要求
nColStart = 3

'Sheet2的行号和列号，用于获得所校验表的校验要求，同时nCol1-1也是所校验表该字段的列号
Dim nRow1, nCol1 As Integer

'所要校验的表的总数
nRowEnd = Ws2.[A65535].End(xlUp).Row

'某个表索要的字段总数为nColEnd

For nCol = 1 To 5
    If Cells(nRow, nCol).Value = "" Then
        Exit For
    Else
        For nRow1 = nRowStart To nRowEnd
            If Ws2.Cells(nRow1, 1).Value = Cells(nRow, nCol).Value Then
                
                '找到目标工作簿
                Dim TargetBook As Workbook
                Set TargetBook = Workbooks.Open(Ws2.Cells(nRow1, 2).Value)
                
                '检查空行空列以及数据类型是否符合要求
                FormatCheck TargetBook
                nColEnd = Ws2.Range("IV" & nRow1).End(xlToLeft).Column
                
                For nCol1 = nColStart To nColEnd
                    '获得某个表某字段校验要求
                    Dim VecCheckId
                    Dim strCheckId As String
                    VecCheckId = Split(Ws2.Cells(nRow1, nCol1), ",")
                    
                    Dim i As Integer
                    For i = LBound(VecCheckId) To UBound(VecCheckId)
                        strCheckId = VecCheckId(i)
                        If IsCheckPassed(strCheckId, Ws3, TargetBook, nCol1 - 1) = False Then
                            Ws1.Cells(nRow + 1, nCol).Interior.ColorIndex = 3
                            Ws1.Cells(nRow + 1, nCol).Value = "又双叒叕错了-_-!"
                        Else
                            Ws1.Cells(nRow + 1, nCol).Interior.ColorIndex = 4
                            Ws1.Cells(nRow + 1, nCol).Value = "恭喜！"
                        End If
                    Next
                Next
                
                TargetBook.Save
                TargetBook.Close
            End If
        Next
    End If
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
Dim nRow, nCol As Integer
nRow = 21
nCol = 1

'设置Sheet2变量
Dim Ws As Worksheet
Set Ws = Sheets(2)

Dim nRowEnd, nRow1 As Integer
nRowEnd = Ws.[A65535].End(xlUp).Row

If Cells(nRow, nCol).Value = "" Then
    MsgBox "表1内没有配置表名称", vbOKOnly, "你很皮啊"
Else
    Dim strWb As String
    strWb = Cells(nRow, nCol).Value
    
    For nRow1 = 3 To nRowEnd
        If Ws.Cells(nRow1, 1) = strWb Then
            Dim TargetBook As Workbook
            Set TargetBook = Workbooks.Open(Ws.Cells(nRow1, 2).Value)
        End If
    Next
End If

End Sub

Sub OpenWB2()

'打开21行2列中的Excel工作簿
Dim nRow, nCol As Integer
nRow = 21
nCol = 2

'设置Sheet2变量
Dim Ws As Worksheet
Set Ws = Sheets(2)

Dim nRowEnd, nRow1 As Integer
nRowEnd = Ws.[A65535].End(xlUp).Row

If Cells(nRow, nCol).Value = "" Then
    MsgBox "表2内没有配置表名称", vbOKOnly, "你很皮啊"
Else
    Dim strWb As String
    strWb = Cells(nRow, nCol).Value
    
     For nRow1 = 3 To nRowEnd
        If Ws.Cells(nRow1, 1) = strWb Then
            Dim TargetBook As Workbook
            Set TargetBook = Workbooks.Open(Ws.Cells(nRow1, 2).Value)
        End If
    Next
End If

End Sub

Sub OpenWB3()

'打开21行3列中的Excel工作簿
Dim nRow, nCol As Integer
nRow = 21
nCol = 3

'设置Sheet2变量
Dim Ws As Worksheet
Set Ws = Sheets(2)

Dim nRowEnd, nRow1 As Integer
nRowEnd = Ws.[A65535].End(xlUp).Row

If Cells(nRow, nCol).Value = "" Then
    MsgBox "表3内没有配置表名称", vbOKOnly, "你很皮啊"
Else
    Dim strWb As String
    strWb = Cells(nRow, nCol).Value
    
     For nRow1 = 3 To nRowEnd
        If Ws.Cells(nRow1, 1) = strWb Then
            Dim TargetBook As Workbook
            Set TargetBook = Workbooks.Open(Ws.Cells(nRow1, 2).Value)
        End If
    Next
End If

End Sub

Sub OpenWB4()

'打开21行4列中的Excel工作簿
Dim nRow, nCol As Integer
nRow = 21
nCol = 4

'设置Sheet2变量
Dim Ws As Worksheet
Set Ws = Sheets(2)

Dim nRowEnd, nRow1 As Integer
nRowEnd = Ws.[A65535].End(xlUp).Row

If Cells(nRow, nCol).Value = "" Then
    MsgBox "表4内没有配置表名称", vbOKOnly, "你很皮啊"
Else
    Dim strWb As String
    strWb = Cells(nRow, nCol).Value
    
     For nRow1 = 3 To nRowEnd
        If Ws.Cells(nRow1, 1) = strWb Then
            Dim TargetBook As Workbook
            Set TargetBook = Workbooks.Open(Ws.Cells(nRow1, 2).Value)
        End If
    Next
End If

End Sub

Sub OpenWB5()

'打开21行5列中的Excel工作簿
Dim nRow, nCol As Integer
nRow = 21
nCol = 5

'设置Sheet2变量
Dim Ws As Worksheet
Set Ws = Sheets(2)

Dim nRowEnd, nRow1 As Integer
nRowEnd = Ws.[A65535].End(xlUp).Row

If Cells(nRow, nCol).Value = "" Then
    MsgBox "表5内没有配置表名称", vbOKOnly, "你很皮啊"
Else
    Dim strWb As String
    strWb = Cells(nRow, nCol).Value
    
     For nRow1 = 3 To nRowEnd
        If Ws.Cells(nRow1, 1) = strWb Then
            Dim TargetBook As Workbook
            Set TargetBook = Workbooks.Open(Ws.Cells(nRow1, 2).Value)
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
                    DirName = Dir("E:\Game301\Game301_Client_NewDesign\Src\Game301\Assets\Resources\" & TargetBook.Sheets(1).Cells(nRowInRes, nCol).Value & ".*")
                    
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
