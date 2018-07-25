Attribute VB_Name = "CommonCheck"
Option Explicit

Sub CheckAll()

'�ҵ�sheet1,�������У����
Dim Ws1 As Worksheet
Set Ws1 = Sheets(1)

'�ҵ�sheet2,����Ѱ��ָ���������Լ�У������
Dim Ws2 As Worksheet
Set Ws2 = Sheets(2)

'�ҵ�sheet3,����Ѱ��ָ����У�����
Dim Ws3 As Worksheet
Set Ws3 = Sheets(3)

'ȫ��У��ʱ��У���������Χ12-15�У�1-5��
Dim nRowForOutput, nColForOutput As Integer
nRowForOutput = 12
nColForOutput = 1

'ȫ��У��ʱsheet��������
Dim nRowStart, nRowEnd, nRow As Integer
nRowEnd = Ws2.[A65535].End(xlUp).Row
nRowStart = 4

'У��ָ����ʱ���ֶε���ʼ�кţ���ǰ�кź���������
Dim nColStart, nColEnd, nCol As Integer
nColStart = 2

'�жϸñ��Ƿ��Ѿ��д���
Dim bIsError As Boolean
bIsError = False

'�ļ�·��ǰ׺
Dim FilePath As String

'����������
Dim TargetBook As Workbook

For nRow = nRowStart To nRowEnd
    
    '�����Excel�ļ�������·��ǰ׺��򿪣��������·��·��ǰ׺
    If Ws2.Cells(nRow, 1).Value Like "*.xlsx" Then
        Set TargetBook = Workbooks.Open(FilePath & Ws2.Cells(nRow, 1).Value)
    Else
        FindFilePath Ws2.Cells(nRow, 1).Value, FilePath
    End If
     
    '�����п����Լ����������Ƿ����Ҫ��
    FormatCheck TargetBook
    
    '���У���ļ���������
    nColEnd = Ws2.Range("IV" & nRow).End(xlToLeft).Column
    For nCol = nColStart To nColEnd
        
        '���ĳ����ĳ�ֶ�У��Ҫ��
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
        MsgBox "�����Ԫ�������������޸�У����"
        Exit For
    End If
    bIsError = False
    
    '���桢�رչ�����
    TargetBook.Save
    TargetBook.Close
Next
End Sub

Sub CheckPartial()

'�ҵ�sheet1,�������У����
Dim Ws1 As Worksheet
Set Ws1 = Sheets(1)

'�ҵ�sheet2,����Ѱ��ָ���������Լ�У������
Dim Ws2 As Worksheet
Set Ws2 = Sheets(2)

'�ҵ�sheet3,����Ѱ��ָ����У�����
Dim Ws3 As Worksheet
Set Ws3 = Sheets(3)

'Sheet2���кź��кţ����ڻ����У����У��Ҫ��ͬʱnCol-1Ҳ����У�����ֶε��к�
Dim nRow, nRowStart, nRowEnd, nCol, nColStart, nColEnd As Integer

'Sheet2���������ӵ�4�п�ʼ
nRowStart = 4

'�������sheet2��3�࿪ʼ��ȡ�������ֶ�У��Ҫ��
nColStart = 3

'�򿪵�21��1-5���еĹ�����
Dim nRowInWs1, nColInWs1 As Integer
nRowInWs1 = 21

'��ҪУ��ı������
nRowEnd = Ws2.[A65535].End(xlUp).Row

'�ļ�·��ǰ׺
Dim FilePath As String

'����������
Dim TargetBook As Workbook

For nColInWs1 = 1 To 5
    
    '���21�еĵ�Ԫ��Ϊ�գ�˵������û��ҪУ��������ˣ�����У��
    If Ws1.Cells(nRowInWs1, nColInWs1).Value = "" Then
        Exit For
    End If
    
    For nRow = nRowStart To nRowEnd
            
        If Ws2.Cells(nRow, 1).Value Like "*_FILE_PATH" Then
            FindFilePath Ws2.Cells(nRow, 1).Value, FilePath
        ElseIf Ws2.Cells(nRow, 1).Value = Ws1.Cells(nRowInWs1, nColInWs1).Value Then
              
            '�ҵ�Ŀ�깤����
            Set TargetBook = Workbooks.Open(FilePath & Ws2.Cells(nRow, 1).Value)
                
            '�����п����Լ����������Ƿ����Ҫ��
            FormatCheck TargetBook
            
            '���У���ļ���������
            nColEnd = Ws2.Range("IV" & nRow).End(xlToLeft).Column
            For nCol = nColStart To nColEnd
                
                '���ĳ����ĳ�ֶ�У��Ҫ��
                Dim VecCheckId
                Dim strCheckId As String
                VecCheckId = Split(Ws2.Cells(nRow, nCol), ",")
                    
                Dim i As Integer
                For i = LBound(VecCheckId) To UBound(VecCheckId)
                    strCheckId = VecCheckId(i)
                    If IsCheckPassed(strCheckId, Ws3, TargetBook, nCol - 1) = False Then
                        Ws1.Cells(nRowInWs1 + 1, nColInWs1).Interior.ColorIndex = 3
                        Ws1.Cells(nRowInWs1 + 1, nColInWs1).Value = "��˫��������-_-!"
                    Else
                        Ws1.Cells(nRowInWs1 + 1, nColInWs1).Interior.ColorIndex = 4
                        Ws1.Cells(nRowInWs1 + 1, nColInWs1).Value = "��ϲ��"
                    End If
                Next
            Next
            
            '���桢�رչ�����
            TargetBook.Save
            TargetBook.Close
        End If
    Next
Next

End Sub

Sub ClearCheckAllResult()

'���sheet1 12-15�У�1-5�е�����
Dim nRow, nCol As Integer

For nRow = 12 To 15
    For nCol = 1 To 5
        Cells(nRow, nCol).Interior.ColorIndex = 0
        Cells(nRow, nCol).Value = ""
    Next
Next

End Sub

Sub ClearCheckPartialResult()

'���sheet1 22�У�1-5�е�����
Dim nRow, nCol As Integer
nRow = 22

For nCol = 1 To 5
    Cells(nRow, nCol).Interior.ColorIndex = 0
    Cells(nRow, nCol).Value = ""
Next

End Sub

Sub OpenWB1()

'��21��1���е�Excel������
Dim nRowInWs1, nColInWs1 As Integer
nRowInWs1 = 21
nColInWs1 = 1

'����Sheet2����
Dim Ws2 As Worksheet
Set Ws2 = Sheets(2)

Dim nRowEnd, nRow As Integer
nRowEnd = Ws2.[A65535].End(xlUp).Row

'�ļ�·��ǰ׺
Dim FilePath As String

'����������
Dim TargetBook As Workbook

If Cells(nRowInWs1, nColInWs1).Value = "" Then
    MsgBox "��1��û�����ñ�����", vbOKOnly, "���Ƥ��"
Else
    For nRow = 4 To nRowEnd
        If Ws2.Cells(nRow, 1).Value Like "*_FILE_PATH" Then
            FilePath = Ws2.Cells(nRow, 1).Value
        ElseIf Ws2.Cells(nRow, 1).Value = Cells(nRowInWs1, nRowInWs1).Value Then
            '�ҵ�Ŀ�깤����
            Set TargetBook = Workbooks.Open(FilePath & Ws2.Cells(nRow, 1).Value)
        End If
    Next
End If

End Sub

Sub OpenWB2()

'��21��2���е�Excel������
Dim nRowInWs1, nColInWs1 As Integer
nRowInWs1 = 21
nColInWs1 = 2

'����Sheet2����
Dim Ws2 As Worksheet
Set Ws2 = Sheets(2)

Dim nRowEnd, nRow As Integer
nRowEnd = Ws2.[A65535].End(xlUp).Row

'�ļ�·��ǰ׺
Dim FilePath As String

'����������
Dim TargetBook As Workbook

If Cells(nRowInWs1, nColInWs1).Value = "" Then
    MsgBox "��1��û�����ñ�����", vbOKOnly, "���Ƥ��"
Else
    For nRow = 4 To nRowEnd
        If Ws2.Cells(nRow, 1).Value Like "*_FILE_PATH" Then
            FilePath = Ws2.Cells(nRow, 1).Value
        ElseIf Ws2.Cells(nRow, 1).Value = Cells(nRowInWs1, nRowInWs1).Value Then
            '�ҵ�Ŀ�깤����
            Set TargetBook = Workbooks.Open(FilePath & Ws2.Cells(nRow, 1).Value)
        End If
    Next
End If

End Sub

Sub OpenWB3()

'��21��3���е�Excel������
Dim nRowInWs1, nColInWs1 As Integer
nRowInWs1 = 21
nColInWs1 = 3

'����Sheet2����
Dim Ws2 As Worksheet
Set Ws2 = Sheets(2)

Dim nRowEnd, nRow As Integer
nRowEnd = Ws2.[A65535].End(xlUp).Row

'�ļ�·��ǰ׺
Dim FilePath As String

'����������
Dim TargetBook As Workbook

If Cells(nRowInWs1, nColInWs1).Value = "" Then
    MsgBox "��1��û�����ñ�����", vbOKOnly, "���Ƥ��"
Else
    For nRow = 4 To nRowEnd
        If Ws2.Cells(nRow, 1).Value Like "*_FILE_PATH" Then
            FilePath = Ws2.Cells(nRow, 1).Value
        ElseIf Ws2.Cells(nRow, 1).Value = Cells(nRowInWs1, nRowInWs1).Value Then
            '�ҵ�Ŀ�깤����
            Set TargetBook = Workbooks.Open(FilePath & Ws2.Cells(nRow, 1).Value)
        End If
    Next
End If

End Sub

Sub OpenWB4()

'��21��4���е�Excel������
Dim nRowInWs1, nColInWs1 As Integer
nRowInWs1 = 21
nColInWs1 = 4

'����Sheet2����
Dim Ws2 As Worksheet
Set Ws2 = Sheets(2)

Dim nRowEnd, nRow As Integer
nRowEnd = Ws2.[A65535].End(xlUp).Row

'�ļ�·��ǰ׺
Dim FilePath As String

'����������
Dim TargetBook As Workbook

If Cells(nRowInWs1, nColInWs1).Value = "" Then
    MsgBox "��1��û�����ñ�����", vbOKOnly, "���Ƥ��"
Else
    For nRow = 4 To nRowEnd
        If Ws2.Cells(nRow, 1).Value Like "*_FILE_PATH" Then
            FilePath = Ws2.Cells(nRow, 1).Value
        ElseIf Ws2.Cells(nRow, 1).Value = Cells(nRowInWs1, nRowInWs1).Value Then
            '�ҵ�Ŀ�깤����
            Set TargetBook = Workbooks.Open(FilePath & Ws2.Cells(nRow, 1).Value)
        End If
    Next
End If

End Sub

Sub OpenWB5()

'��21��5���е�Excel������
Dim nRowInWs1, nColInWs1 As Integer
nRowInWs1 = 21
nColInWs1 = 5

'����Sheet2����
Dim Ws2 As Worksheet
Set Ws2 = Sheets(2)

Dim nRowEnd, nRow As Integer
nRowEnd = Ws2.[A65535].End(xlUp).Row

'�ļ�·��ǰ׺
Dim FilePath As String

'����������
Dim TargetBook As Workbook

If Cells(nRowInWs1, nColInWs1).Value = "" Then
    MsgBox "��1��û�����ñ�����", vbOKOnly, "���Ƥ��"
Else
    For nRow = 4 To nRowEnd
        If Ws2.Cells(nRow, 1).Value Like "*_FILE_PATH" Then
            FilePath = Ws2.Cells(nRow, 1).Value
        ElseIf Ws2.Cells(nRow, 1).Value = Cells(nRowInWs1, nRowInWs1).Value Then
            '�ҵ�Ŀ�깤����
            Set TargetBook = Workbooks.Open(FilePath & Ws2.Cells(nRow, 1).Value)
        End If
    Next
End If

End Sub

Sub CheckAllResourceExist()

'����Sheet1��Sheet4����
Dim Ws1, Ws4 As Worksheet
Set Ws1 = Sheets(1)
Set Ws4 = Sheets(4)

Dim nRowStart, nRowEnd, nRow, nCol, nColStart, nColEnd, nRowForOutput, nColForOutput As Integer

'������Դ�Ĳ߻���·��ʼ�ڵ�3�У�ֹ�ڵ�nRowEnd��
nRowStart = 3
nRowEnd = Ws4.[A65535].End(xlUp).Row

'����У���ĵ�2�п�ʼ������Դ·���ֶ�
nColStart = 2

'��������ڱ�1��30-32�еĵ�1-5��
nRowForOutput = 30
nColForOutput = 1

For nRow = nRowStart To nRowEnd
    
    '�򿪶�Ӧ������
    Dim TargetBook As Workbook
    Set TargetBook = Workbooks.Open(Ws4.Cells(nRow, 2).Value)
    nColEnd = TargetBook.Sheets(1).Range("IV1").End(xlToLeft).Column
    
    For nCol = nColStart To nColEnd
        If TargetBook.Sheets(1).Cells(1, nCol).Value Like "*Path" Then
            Ws1.Cells(nRowForOutput, nColForOutput) = Ws4.Cells(nRow, 1).Value
            Ws1.Cells(nRowForOutput, nColForOutput).Interior.ColorIndex = 4
            
            '����������nRowEndInRes������
            Dim nRowInRes, nRowEndInRes As Integer
            nRowEndInRes = TargetBook.Sheets(1).[A65535].End(xlUp).Row
            
            Dim DirName As String
            For nRowInRes = 10 To nRowEndInRes
                
                '���·���ֶβ�Ϊ�գ�˵����������Ϊ��Դ����Ҫ���Ҹ�·�����ļ��Ƿ����
                If TargetBook.Sheets(1).Cells(nRowInRes, nCol).Value <> "" Then
                    DirName = Dir(RESOURCE_FILE_PATH & TargetBook.Sheets(1).Cells(nRowInRes, nCol).Value & ".*")
                    
                    If DirName = "" Then
                        TargetBook.Sheets(1).Cells(nRowInRes, 1).Interior.ColorIndex = 3
                        Ws1.Cells(nRowForOutput, nColForOutput).Interior.ColorIndex = 3
                    End If
                End If
            Next
            '��5�л��м������
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
    
    '���ֻ�ܵ�32�У���33����˵����Ҫ��չ��������Ԫ��
    If nRowForOutput = 33 Then
        TargetBook.Save
        TargetBook.Close
        Exit For
    End If
Next

End Sub

Sub ClearCheckAllResourceExist()

'����Sheet1����
Dim Ws As Worksheet
Set Ws = Sheets(1)

Dim nRow, nRowStart, nRowEnd, nCol, nColStart, nColEnd As Integer
'���Sheet1 30-32��1-5�е�����
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
