VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "不重复集合"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private 字典 As Object

Private Sub Class_Initialize()
    Set 字典 = CreateObject("Scripting.Dictionary")
End Sub

Public Sub 加入(元素)
    If Not 字典.Exists(元素) Then
        字典.Add 元素, Nothing
    End If
End Sub

Public Function 转数组() As String()
    Dim 数组() As String
    键们 = 字典.Keys
    条目数 = 字典.Count
    ReDim 数组(条目数 - 1)
    For 某 = 0 To 条目数 - 1
        数组(某) = 键们(某)
    Next
    转数组 = 数组
End Function
