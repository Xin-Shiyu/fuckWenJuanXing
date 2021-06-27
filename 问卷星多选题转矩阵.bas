Attribute VB_Name = "问卷星多选题转矩阵"
Const 分隔符 = "┋"

Function 处理其他(ByVal 原文 As String) As String
    If Left(原文, 2) = "其他" Then
        If Right(原文, 1) = "〗" Then
            处理其他 = Left(原文, Len(原文) - 1)
            处理其他 = Right(处理其他, Len(处理其他) - 7)
            Exit Function
        End If
        处理其他 = "其他（未填）"
        Exit Function
    End If
    处理其他 = 原文
End Function

Function 获取所有选项(ByVal 当前行 As Long, ByVal 当前列 As Long, ByRef 问卷数 As Long) As String()
    Dim 集合 As New 不重复集合
    Do
        当前行 = 当前行 + 1
        问卷数 = 问卷数 + 1
        当前单元格 = Cells(当前行, 当前列)
        If 当前单元格 = "" Then Exit Do
        出现值们 = Split(当前单元格, 分隔符)
        For Each 出现值 In 出现值们
            集合.加入 处理其他(出现值)
        Next
    Loop
    获取所有选项 = 集合.转数组
End Function

Sub 生成多选矩阵()
    原始数据列 = Selection.Column
    问题 = Cells(1, 原始数据列)
    Dim 问卷数 As Long
    问卷数 = 0
    所有选项 = 获取所有选项(1, 原始数据列, 问卷数)
    For i = 0 To UBound(所有选项)
        Columns(原始数据列 + 1).Insert xlShiftToRight
    Next
    For i = 0 To UBound(所有选项)
        当前列 = 原始数据列 + 1 + i
        Cells(1, 当前列) = 问题 + "：" + 所有选项(i)
        For j = 1 To 问卷数
            Cells(1 + j, 当前列) = InStr(1, Cells(1 + j, 原始数据列), 所有选项(i)) <> 0
        Next
    Next
End Sub
