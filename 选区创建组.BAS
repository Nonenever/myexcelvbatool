Sub 选区创建组(rng)
    On Error Resume Next
    'Set rng = Selection
    If getnum(rng.Address) = 0 Then
        Exit Sub
    End If
    With ActiveSheet.Outline
        .AutomaticStyles = False
        .summaryRow = xlAbove
        .SummaryColumn = xlRight
    End With
    Dim dic
    Set dic = CreateObject("scripting.dictionary")
    '数据结束行 = Cells.Find("*", , , , , xlPrevious).Row
    开始行号 = rng.Offset(0, 0).row
    结束行号 = rng.Offset(rng.Rows.count - 1, 0).row
    Dim 组级别, 下一组级别, 下组级别
    下组级别 = Rows(结束行号 + 1).OutlineLevel
    组级别 = Rows(结束行号).OutlineLevel
    For i = 开始行号 To 结束行号
        If Not dic.Exists(i) Then dic.Add i, Rows(i).OutlineLevel + 1
    Next i
    If 组级别 + 1 <= 下组级别 Then
        i = 结束行号 + 1
        Do While 组级别 <> Rows(i).OutlineLevel
            If Not dic.Exists(i) Then dic.Add i, Rows(i).OutlineLevel + 1
            i = i + 1
        Loop
    End If
    For Each e In dic.keys
        Rows(e).OutlineLevel = dic(e)
    Next
End Sub
