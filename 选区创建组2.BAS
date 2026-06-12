Sub 选区创建组(rng As Range)
Rem CreateRowGroup
    Dim dic As Object
    Dim startRow As Long, endRow As Long
    Dim curLevel As Long, nextLevel As Long
    Dim i As Long, key As Variant
    
    ' ========== 1. 基础合法性校验 ==========
    ' 传入区域为空直接退出
    If rng Is Nothing Then Exit Sub
    ' 调用外部函数 getnum 判断是否有效区域（保留原有业务逻辑）
    On Error Resume Next
    If getnum(rng.Address) = 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    
    ' ========== 2. 大纲样式统一设置 ==========
    With ActiveSheet.Outline
        .AutomaticStyles = False
        .summaryRow = xlAbove
        .SummaryColumn = xlRight
    End With
    
    ' 获取选区起止行
    startRow = rng.row
    endRow = rng.row + rng.Rows.count - 1
    
    ' 读取当前行级别、下一行级别
    curLevel = Rows(endRow).OutlineLevel
    nextLevel = Rows(endRow + 1).OutlineLevel
    
    ' 创建字典存储 行号:新级别
    Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = vbBinaryCompare
    
    ' ========== 3. 遍历选区行，设置新分组级别 ==========
    For i = startRow To endRow
        If Not dic.Exists(i) Then
            dic(i) = Rows(i).OutlineLevel + 1
        End If
    Next i
    
    ' 原有逻辑：向后延伸匹配同级行
    If curLevel + 1 <= nextLevel Then
        i = endRow + 1
        Do While curLevel <> Rows(i).OutlineLevel
            If Not dic.Exists(i) Then
                dic(i) = Rows(i).OutlineLevel + 1
            End If
            i = i + 1
        Loop
    End If
    
    ' ========== 4. 批量应用分组级别 ==========
    For Each key In dic.keys
        Rows(key).OutlineLevel = dic(key)
    Next key
    
    Set dic = Nothing
End Sub
