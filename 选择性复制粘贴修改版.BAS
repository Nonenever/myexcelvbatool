Rem 增加跨工作薄复制
Public rng As Range '源选区
Public rg As Range '目标选区
Public rng_Name, rg_Name, message, title
Public wbk_rng As Workbook, wbk_rg As Workbook
Public wk_rng As Worksheet, wk_rg As Worksheet
Public str_rng As String, str_rg As String
Public 只填目标可见单元格 As Boolean, 不转置 As Boolean, 只选源可见单元格 As Boolean, 复制单元格格式 As Boolean, 复制粘贴 As Boolean
Public 值复制 As Boolean, 公式复制 As Boolean, 复制链接 As Boolean ', 格式复制 As Boolean
Public rng_equal As String
Sub u_选择性复制粘贴修改版()
    选择性复制粘贴修改版.Show 0
End Sub
Private Sub OptionButton1_Click() '值复制
    值复制 = OptionButton1.Value
    If 值复制 Then
        '    值复制 = False
        公式复制 = False
        复制链接 = False
        复制粘贴 = False
    End If
End Sub
Private Sub OptionButton2_Click() '公式复制
    公式复制 = OptionButton2.Value
    If 公式复制 Then
        值复制 = False
        '    公式复制 = False
        复制链接 = False
        复制粘贴 = False
    End If
End Sub
Private Sub OptionButton3_Click() '复制链接
    复制链接 = OptionButton3.Value
    If 复制链接 Then
        值复制 = False
        公式复制 = False
        '    复制链接 = False
        复制粘贴 = False
    End If
End Sub
Private Sub CheckBox4_Click() '复制单元格格式
    复制单元格格式 = CheckBox4.Value
End Sub
Private Sub OptionButton4_Click() '复制粘贴
    复制粘贴 = OptionButton4.Value
    If 复制粘贴 Then
        值复制 = False
        公式复制 = False
        复制链接 = False
        '    复制粘贴 = False
    End If
End Sub
Private Sub UserForm_Initialize()
    If Selection.Cells.count > 1 Then
        Set rng = Selection
        'rng_Name = "'" & Rng.Parent.Name & "'!" '& rng.Address
        'TextBox1.Value = rng_Name & Rng.Address
        rng_Name = "'[" & ActiveWorkbook.Name & "]" & rng.Parent.Name & "'!" '& rng.Address '''2021年01月05日
        str_rng = rng_Name & rng.Address
        TextBox1.Value = str_rng ' rng_Name & rng.Address '''2021年01月05日
        Set wbk_rng = ActiveWorkbook
        Set wk_rng = ActiveSheet
        '='[VLSH - 副本.xls]5-1分包结算封面'!$R$6
    End If
    CheckBox1 = True '只选源可见单元格
    CheckBox2 = True '只填目标可见单元格
    CheckBox3 = False  '不转置
    CheckBox4 = False  '复制单元格格式
    OptionButton4 = True '复制粘贴
    '    TextBox1.MultiLine = True
    '        TextBox1.WordWrap = True ' 启用自动换行
    '         TextBox2.MultiLine = True
    '       TextBox2.WordWrap = True ' 启用自动换行
    '
End Sub
Private Sub CheckBox1_Click() '只选源可见单元格
End Sub
Private Sub CheckBox2_Click() '只填目标可见单元格
End Sub
Private Sub CheckBox3_Click() '不转置
End Sub
Private Sub TextBox1_Change() '输入源选区
End Sub
Private Sub TextBox2_Change() '目标选区
End Sub
Private Sub CommandButton1_Click() '执行
    Call Selection_选区单元格_地址
End Sub
Private Sub CommandButton2_Click() '源单元格
    message = "输入源选区"
    title = "源选区"
    Set rng = Application.InputBox(message, title, , 10000, 5000, , , 8)
    rng_Name = "'[" & ActiveWorkbook.Name & "]" & rng.Parent.Name & "'!" '& rng.Address '''2021年01月05日
    TextBox1.Value = rng_Name & rng.Address '''2021年01月05日
    Set wbk_rng = ActiveWorkbook
    Set wk_rng = ActiveSheet
    '='[VLSH - 副本.xls]5-1分包结算封面'!$R$6
End Sub
Private Sub CommandButton4_Click()
    message = "输入目的地选区"
    title = "目的地选区"
    Set rg = Application.InputBox(message, title, , 10000, 5000, , , 8)
    'rg_Name = "'" & rg.Parent.Name & "'!" '& rg.Address
    'TextBox2.Value = rg_Name & rg.Address
    '
    rg_Name = "'[" & ActiveWorkbook.Name & "]" & rg.Parent.Name & "'!" '& rng.Address '''2021年01月05日
    str_rg = rg_Name & rg.Address
    TextBox2.Value = str_rg ' rg_Name & rg.Address '''2021年01月05日
    Set wbk_rg = ActiveWorkbook
    Set wk_rg = ActiveSheet
    '='[VLSH - 副本.xls]5-1分包结算封面'!$R$6
End Sub
Private Sub Selection_选区单元格_地址()
    On Error Resume Next
    Dim ls_S_S As Range
    Dim ls_T_T As Range
    只选源可见单元格 = CheckBox1.Value
    只填目标可见单元格 = CheckBox2.Value
    转置 = CheckBox3.Value
    数值格式复制 = CheckBox4.Value
    If TextBox1.Value <> "" Then '源选区
        Rem --------------------------------------------2022-10-27----------------------------
        '        str_rng = TextBox1.Value
        '
        '        If TextBox1.Value = rng_Name & rng.Address Then
        '        Else
        '            rng_Name = Split(TextBox1.Value, "!")(0) & "!"
        '            Set rng = Range(Split(TextBox1.Value, "!")(1))
        '        End If
        '        Debug.Print str_rng
        Rem --------------------------------------------2022-10-27----------------------------
    Else
        MsgBox "源选区有问题！！！"
        Exit Sub
    End If
    If TextBox2.Value <> "" Then '目标选区
        Rem --------------------------------------------2022-10-27----------------------------
        '        str_rg = TextBox2.Value
        '        Debug.Print str_rg
        '        If TextBox2.Value = rg_Name & rg.Address Then
        '        Else
        '            rg_Name = Split(TextBox2.Value, "!")(0) & "!"
        '            Set rg = Range(Split(TextBox2.Value, "!")(1))
        '        End If
        Rem --------------------------------------------2022-10-27----------------------------
    Else
        MsgBox "目标选区有问题！！！"
        Exit Sub
    End If
    Application.ScreenUpdating = False
    If wbk_rg.Name <> wbk_rng.Name Then  '是否同一工作薄
        rng_equal = "=" & rng_Name
    Else
        If wk_rg.Name <> wk_rng.Name Then  '是否同一工作表
            rng_equal = "='" & wk_rng.Name & "'!"
        Else
            rng_equal = "="
        End If
    End If
    Dim rs, js, rng_arr_RJ()
    Dim i, j, ic, jc, rng_Visible_C_C, rng_Visible_R_C
    Dim rng_R_Start, rng_C_Start, rng_R_End, rng_C_End, rng_R_Count, rng_C_Count
    Dim rng_arr()
    LSX_TXT = Replace(StrReverse(Split(StrReverse(str_rng), ".", 2)(1)), "'[", "")
    wbk_rng.Activate
    wk_rng.Activate
    If 只选源可见单元格 = True And rng.SpecialCells(xlCellTypeVisible).Areas.count > 1 Then
        rng_Visible_R_C = Intersect(rng.SpecialCells(xlCellTypeVisible), rng.Columns(1)).Cells.count
        rng_Visible_C_C = Intersect(rng.SpecialCells(xlCellTypeVisible), rng.Rows(1)).Cells.count
        rng_R_Start = rng.row
        rng_C_Start = rng.Column
        rng_R_End = rng.Rows.count + rng.row - 1
        rng_C_End = rng.Columns.count + rng.Column - 1
        ic = 0
        jc = 0
        ReDim rng_arr(1 To rng_Visible_R_C, 1 To rng_Visible_C_C)
        For i = rng_R_Start To rng_R_End
            If Rows(i).Hidden = False Then
                ic = 1 + ic
                jc = 0
                For j = rng_C_Start To rng_C_End
                    If Columns(j).Hidden = False Then
                        jc = 1 + jc
                        rng_arr(ic, jc) = rng_equal & Cells(i, j).Address
                    End If
                Next
            End If
        Next
    Else
        rng_R_Count = rng.Rows.count
        rng_C_Count = rng.Columns.count
        ReDim rng_arr(1 To rng_R_Count, 1 To rng_C_Count)
        For i = 1 To rng_R_Count
            For j = 1 To rng_C_Count
                rng_arr(i, j) = rng_equal & rng.Cells(i, j).Address
            Next
        Next
    End If
    wbk_rg.Activate
    wk_rg.Activate
    rg_R_Start = rg.row
    rg_C_Start = rg.Column
    Dim ls_1
    If 转置 = True Then '2020年05月17日新增
        ReDim rng_arr_RJ(LBound(rng_arr, 2) To UBound(rng_arr, 2), LBound(rng_arr) To UBound(rng_arr))
        '
        For js = LBound(rng_arr) To UBound(rng_arr)
            For rs = LBound(rng_arr, 2) To UBound(rng_arr, 2)
                rng_arr_RJ(rs, js) = rng_arr(js, rs)
            Next rs
        Next js
        rng_arr = rng_arr_RJ
    End If
    Dim dic_C
    Dim ls_T
    Set dic_C = CreateObject("scripting.dictionary")
    Dim dic_R
    Set dic_R = CreateObject("scripting.dictionary")
    If 只填目标可见单元格 = True Then
        ic = 1
        jc = 0
        Do Until jc = 1 'rg_R_Start - 1 + ic <= rg_R_Start + rng_Visible_C_C - 1 行
            If Rows(rg_R_Start - 1 + ic).EntireRow.Hidden = False Then
                dic_R.Add rg_R_Start - 1 + ic, Nothing
            End If
            ic = ic + 1
            If dic_R.count = UBound(rng_arr) Then  'rng_Visible_C_C Then '''''''dic_R.Count = 1 Or''''''''''''
                jc = 1
            End If
        Loop
        jc = 1
        ic = 0
        Do Until ic = 1  'rg_C_Start - 1 + jc <= rg_C_Start + rng_Visible_R_C - 1
            If Columns(rg_C_Start - 1 + jc).Hidden = False Then
                dic_C.Add rg_C_Start - 1 + jc, Nothing
            End If
            jc = jc + 1
            If dic_C.count = UBound(rng_arr, 2) Then    'rng_Visible_R_C Then ''''' dic_C.Count = 1 Or'''''''''''''''
                ic = 1
            End If
            '
        Loop
        keys_C = dic_C.keys
        keys_R = dic_R.keys
        '        ----------------------------------------
        For i = 0 To UBound(keys_C)
            For j = 0 To UBound(keys_R)
                Set ls_S_S = Range(Right(rng_arr(j + 1, i + 1), Len(rng_arr(j + 1, i + 1)) - 1))
                Set ls_T_T = Cells(keys_R(j), keys_C(i))
                If 公式复制 Then '20240626修改
                    ls_S_S.Copy
                    ls_T_T.PasteSpecial Paste:=xlPasteFormulas
                ElseIf 复制链接 Then '20240626修改
                    ls_T_T.formula = "=" & ls_S_S.Address(External:=True)
                ElseIf 值复制 Then
                    ls_T_T = ls_S_S.Value
                Else
                    ls_S_S.Copy ls_T_T                        '                    Range(Right(rng_arr(j + 1, i + 1), Len(rng_arr(j + 1, i + 1)) - 1)).Copy Cells(keys_R(j), keys_C(i))  '20240626修改之前的
                End If '20240626修改
                If 复制单元格格式 Then '20240626修改
                    ls_S_S.Copy
                    ls_T_T.PasteSpecial Paste:=xlPasteFormats
                End If
            Next
        Next
        Rem  ---------原先的-----------------------------------------------
        '        If 值复制 = True Then
        '            For i = 0 To UBound(keys_C)
        '                For j = 0 To UBound(keys_R)
        '                    Cells(keys_R(j), keys_C(i)) = rng_arr(j + 1, i + 1)
        '                Next
        '            Next
        '        Else '2020年05月17日
        '            For i = 0 To UBound(keys_C)
        '                For j = 0 To UBound(keys_R)
        '                    Set ls_S_S = Range(Right(rng_arr(j + 1, i + 1), Len(rng_arr(j + 1, i + 1)) - 1))
        '                    Set ls_T_T = Cells(keys_R(j), keys_C(i))
        '                    If 公式复制 Then '20240626修改
        '                        ls_S_S.Copy
        '                        ls_T_T.PasteSpecial Paste:=xlPasteFormulas
        '                    ElseIf 复制链接 Then '20240626修改
        '                        ls_T_T.Formula = "=" & ls_S_S.Address(External:=True)
        '                    Else
        '                        ls_S_S.Copy ls_T_T                        '                    Range(Right(rng_arr(j + 1, i + 1), Len(rng_arr(j + 1, i + 1)) - 1)).Copy Cells(keys_R(j), keys_C(i))  '20240626修改之前的
        '                    End If '20240626修改
        '                    If 复制单元格格式 Then '20240626修改
        '                        ls_S_S.Copy
        '                        ls_T_T.PasteSpecial Paste:=xlPasteFormats
        '                    End If
        '                Next
        '            Next
        '        End If
        Rem  -----------原先的---------------------------------------------
    Else '不选择可见单元格
        If 公式复制 = True Then
            rg.Resize(UBound(rng_arr) - LBound(rng_arr) + 1, UBound(rng_arr, 2) - LBound(rng_arr, 2) + 1) = rng_arr
            ' ls_T_T.Formula = "=" & ls_S_S.Address(External:=True)
        Else
            For i = LBound(rng_arr) To UBound(rng_arr)
                For j = LBound(rng_arr, 2) To UBound(rng_arr, 2)
                    Set ls_S_S = (Right(rng_arr(i, j), Len(rng_arr(i, j)) - 1))
                    Set ls_T_T = rg.Offset(i - LBound(rng_arr), j - LBound(rng_arr, 2))
                    If 公式复制 Then '20240626修改
                        ls_S_S.Copy
                        ls_T_T.PasteSpecial Paste:=xlPasteFormulas
                    ElseIf 值复制 Then
                        ls_T_T = ls_S_S.Value
                        '                        ls_S_S.Copy
                        '                        ls_T_T.Paste Link:=True
                        '
                    Else
                        ls_S_S.Copy ls_T_T
                        '                    Range(Right(rng_arr(i, j), Len(rng_arr(i, j)) - 1)).Copy rg.Offset(i - LBound(rng_arr), j - LBound(rng_arr, 2)) '20240626修改之前的
                    End If '20240626修改
                    If 复制单元格格式 Then '20240626修改
                        ls_S_S.Copy
                        ls_T_T.PasteSpecial Paste:=xlPasteFormats
                    End If
                Next
            Next
        End If
    End If
    'End With
    Application.ScreenUpdating = True
    MsgBox "OK!"
End Sub
