Attribute VB_Name = "�ʾ��Ƕ�ѡ��ת����"
Const �ָ��� = "��"

Function ��������(ByVal ԭ�� As String) As String
    If Left(ԭ��, 2) = "����" Then
        If Right(ԭ��, 1) = "��" Then
            �������� = Left(ԭ��, Len(ԭ��) - 1)
            �������� = Right(��������, Len(��������) - 7)
            Exit Function
        End If
        �������� = "������δ�"
        Exit Function
    End If
    �������� = ԭ��
End Function

Function ��ȡ����ѡ��(ByVal ��ǰ�� As Long, ByVal ��ǰ�� As Long, ByRef �ʾ��� As Long) As String()
    Dim ���� As New ���ظ�����
    Do
        ��ǰ�� = ��ǰ�� + 1
        �ʾ��� = �ʾ��� + 1
        ��ǰ��Ԫ�� = Cells(��ǰ��, ��ǰ��)
        If ��ǰ��Ԫ�� = "" Then Exit Do
        ����ֵ�� = Split(��ǰ��Ԫ��, �ָ���)
        For Each ����ֵ In ����ֵ��
            ����.���� ��������(����ֵ)
        Next
    Loop
    ��ȡ����ѡ�� = ����.ת����
End Function

Sub ���ɶ�ѡ����()
    ԭʼ������ = Selection.Column
    ���� = Cells(1, ԭʼ������)
    Dim �ʾ��� As Long
    �ʾ��� = 0
    ����ѡ�� = ��ȡ����ѡ��(1, ԭʼ������, �ʾ���)
    For i = 0 To UBound(����ѡ��)
        Columns(ԭʼ������ + 1).Insert xlShiftToRight
    Next
    For i = 0 To UBound(����ѡ��)
        ��ǰ�� = ԭʼ������ + 1 + i
        Cells(1, ��ǰ��) = ���� + "��" + ����ѡ��(i)
        For j = 1 To �ʾ���
            Cells(1 + j, ��ǰ��) = InStr(1, Cells(1 + j, ԭʼ������), ����ѡ��(i)) <> 0
        Next
    Next
End Sub
