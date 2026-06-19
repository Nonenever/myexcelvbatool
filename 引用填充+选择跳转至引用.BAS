Sub 公式引用填充对应颜色() '可用的 2021-09-01更新：区域填充
    On Error Resume Next
    Dim rg  As Range
    Dim c As Range
    Dim str As String
    Dim arr, i
    Dim j As Integer
    j = 2
    Set rg = Selection 'Range("BC3:BC385")
    Set rg = Intersect(rg.Parent.usedRange, rg)
    'RG.Offset(0, -2).Resize(138, 3).Interior.Pattern = xlNone
    rg.Interior.pattern = xlNone
    'RG.Interior.Pattern = xlNone
    For Each c In rg
        If c.HasFormula Then
            str = c.formula
            c.Interior.colorIndex = j
            arr = Split(myadd(str), "/")
            If UBound(arr) - LBound(arr) + 1 > 0 Then
                For i = LBound(arr) To UBound(arr)
                    If Len(arr(i)) > 0 Then
                        If InStr(arr(i), "]") = 0 Then
                            If InStr(arr(i), "!") > 0 Then
                                STR0 = Split(arr(i), "!")(0)
                                str1 = Split(arr(i), "!")(1)
                                STR0 = Replace(STR0, "'", "")
                                'str0 = Replace(str0, """'", "")
                                'str0 = Replace(str0, """'", "")
                                Sheets(STR0).Range(str1).Interior.colorIndex = j
                                'Debug.Print Sheets(STR0).Range(STR1).Address '& "_" & Sheets(STR0).Range(STR1)
                                'Debug.Print STR0 & "+++" & STR1
                            Else
                                Range(arr(i)).Interior.colorIndex = j
                            End If
                        Else
                        End If
                    End If
                Next i
            End If
        End If
        j = j + 1
        j = j Mod 56
    Next c
End Sub
Sub 选择跳转至引用()
    On Error Resume Next
    Dim r As Range, rg As Range, str As String
    Set r = ActiveCell
    str = r.formula
    str = Replace(str, "=", "")
    'Dim AA
    'AA = 字符串中的地址数据数组(str)
    If InStr(字符串中的地址数据数组(str), "/") > 0 Then
        str = Split(字符串中的地址数据数组(str), "/")(1)
    Else
        str = 字符串中的地址数据数组(str)
    End If
    If InStr(str, "!") > 0 Then
        Application.GoTo Reference:=Sheets(Replace(Split(str, "!")(0), "'", "")).Range(Split(str, "!")(1)), Scroll:=True
    Else
        Application.GoTo Reference:=Range(str), Scroll:=True
    End If
    'Cells(ActiveCell.Row, "h").Select
End Sub
Function myadd(rng)
    'https://club.excelhome.net/forum.php?mod=viewthread&tid=1597844&page=1&extra=#pid10771275
    Dim r%, i%
    Dim arr, brr
    Dim reg 'As New RegExp
    Set reg = CreateObject("VBScript.RegExp")
    With reg
        .Global = True
        '.Pattern = "(\$?[A-Z]+\$?\d+:)?\$?[A-Z]+\$?\d+"
        .pattern = "((([一-龥a-zA-Z]+[0-9]+!)|(('((?!\!).)*!)?)|([一-龥a-zA-Z][\u4E00-\u9FA5A-Za-z0-9_\.]+[^']!{1}))\$?[a-zA-Z]+\$?[0-9]+)((:((([一-龥a-zA-Z]+[0-9]+!)|(('((?!\!).)*!)?)|([一-龥a-zA-Z][\u4E00-\u9FA5A-Za-z0-9_\.]+[^']!{1}))\$?[a-zA-Z]+\$?[0-9]+))?)"
    End With
    Set mh = reg.Execute(rng)
    ss = ""
    For j = 0 To mh.count - 1
        ss = ss & "/" & mh(j)
        'ss = ss & Space(1) & mh(J)
    Next
    If ss <> "" Then
        myadd = Mid(ss, 2)
    Else
        mydadd = ""
    End If
End Function
Function 字符串中的地址数据数组(字符串)                                  '2021-08-23 数组 所有单元格地址
    Dim kt, ln, SN, k, s, data, Data1, Data2, 控制器, 小数位数
    Application.Volatile True                                  '设置自定义函数为易失性函数,无论何时重新计算工作表,函数都会重新计算
    s = 字符串
    'data = s
    data = ""
    Set reg = CreateObject("vbscript.regexp")
    'reg.Pattern = "(((('((?!\!).)*!)?)|([一-龥a-zA-Z][\u4E00-\u9FA5A-Za-z0-9_\.]+[^']!{1}))\$?[a-zA-Z]+\$?[0-9]+)"
    reg.pattern = "((([一-龥a-zA-Z]+[0-9]+!)|(('((?!\!).)*!)?)|([一-龥a-zA-Z][\u4E00-\u9FA5A-Za-z0-9_\.]+[^']!{1}))\$?[a-zA-Z]+\$?[0-9]+)"
    reg.Global = True
    Set mh = reg.Execute(s)
    If mh.count = 1 Then
        k = mh(0).Value
        data = k
        'kt = EVAL2(k)
        'data = Replace(s, k, kt)
    Else
        For i = 0 To mh.count - 1
            k = mh(mh.count - 1 - i).Value
            'SN = mh(mh.count - 1 - i).FirstIndex '表示在目标字符串中,位置3上找到该匹配iphone_5s.(位置是从0开始的)
            'ln = mh(mh.count - 1 - i).length
            data = data & "/" & k
            'data1 = Left(data, SN)
            'data2 = Right(data, Len(data) - SN - ln)
            'kt = EVAL2(k)
            'data = data1 & kt & data2
        Next i
    End If
    字符串中的地址数据数组 = data                                         'Split(data, "/")
End Function
