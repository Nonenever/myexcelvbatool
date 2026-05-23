
---

# 四、主代码文件：合同支付申请单_脱敏统一变量版.vba
## 就是我上一版给你的**最终完整代码**，我再贴一次方便你直接保存：
```vba
Option Explicit

'==================== 统一内容变量区（仅此处修改文字、编号、金额、账户信息）====================
Dim strTitle As String          '表单标题
Dim strDocNo As String          '单据编号
Dim strProjName As String       '项目名称
Dim strContractName As String   '合同名称
Dim strUnitName As String       '承建单位
Dim dblPayMoney As Double       '支付金额
Dim strPayBasis As String       '支付依据
Dim strContractNo As String     '合同编号
Dim strPayDesc As String        '支付内容说明

Dim strBankName1 As String      '工程款开户行
Dim strAccName1 As String       '工程款账户名
Dim strAccNo1 As String         '工程款账号

Dim strBankName2 As String      '工资款开户行
Dim strAccName2 As String       '工资款账户名
Dim strAccNo2 As String         '工资款账号

Dim strAuditOpinion As String   '设计院审核意见
'=============================================================================================

'初始化脱敏默认内容
Private Sub InitContentVar()
    strTitle = "合同支付申请单"
    strDocNo = "（编号：№XXXX-XXXX-XXXX号）"
    strProjName = "XX工程项目"
    strContractName = "XX项目施工合同"
    strUnitName = "XX建设有限公司"
    dblPayMoney = 0.000000
    strPayBasis = "项目正式施工合同"
    strContractNo = "XXXX-XXXX-XXXX"
    strPayDesc = "根据合同约定及当期完成工程量，按条款核算本期应付工程款，已依规扣除相关款项，满足付款审批条件。"
    
    strAccName1 = "XX建设有限公司"
    strBankName1 = "XX银行城区支行"
    strAccNo1 = "XXXXXXXXXXXXXXXXXXXX"
    
    strAccName2 = "XX建设农民工工资专户"
    strBankName2 = "XX银行政务支行"
    strAccNo2 = "XXXXXXXXXXXXXX"
    
    strAuditOpinion = "设计院审核计量金额为        元，支付金额为                 元，详见计量支付审查表。" & Chr(10) & "计量支付工程师：       单位技术负责人："
End Sub

Sub 合同支付申请单_统一变量版()
    Call InitContentVar '载入变量内容
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    On Error GoTo 错误恢复
    
    Dim 开始地址 As String
    Dim 开始行 As Long, 开始列 As Long
    Dim rg As Range
    Dim arr As Variant, brr As Variant
    
    开始地址 = Selection.Cells(1).Address
    开始行 = Selection.Row
    开始列 = Selection.Column
    
    '行高列宽
    Rows(开始行).RowHeight = 19.5
    Rows(开始行 + 1).RowHeight = 15
    Rows(开始行 + 2 & ":" & 开始行 + 6).RowHeight = 27.75
    Rows(开始行 + 7).RowHeight = 134.25
    Rows(开始行 + 8 & ":" & 开始行 + 13).RowHeight = 23.25
    Rows(开始行 + 14).RowHeight = 46.5
    Rows(开始行 + 15).RowHeight = 30
    Rows(开始行 + 16 & ":" & 开始行 + 18).RowHeight = 32.25
    Rows(开始行 + 19 & ":" & 开始行 + 21).RowHeight = 30
    Rows(开始行 + 22).RowHeight = 14.25
    
    Range(开始地址).Offset(0, 0).EntireColumn.ColumnWidth = 6.5
    Range(开始地址).Offset(0, 1).EntireColumn.ColumnWidth = 3.25
    Range(开始地址).Offset(0, 2).EntireColumn.ColumnWidth = 11.25
    Range(开始地址).Offset(0, 3).EntireColumn.ColumnWidth = 0.92
    Range(开始地址).Offset(0, 4).EntireColumn.ColumnWidth = 24.75
    Range(开始地址).Offset(0, 5).EntireColumn.ColumnWidth = 8.75
    Range(开始地址).Offset(0, 6).EntireColumn.ColumnWidth = 24.25
    
    '标题
    Set rg = Cells(开始行, 开始列).Resize(1, 7)
    合并设置 rg, strTitle, 15, True, xlCenter
    
    Set rg = Cells(开始行 + 1, 开始列).Resize(1, 7)
    合并设置 rg, strDocNo, 11, True, xlCenter
    边框设置 rg, "//LineStyle/9/1//Weight/9/-4138"
    
    '项目信息
    Set rg = Cells(开始行 + 2, 开始列).Resize(1, 2)
    合并设置 rg, "项目名称", 12, True, xlCenter
    边框设置 rg, "//LineStyle/7/1//Weight/7/-4138//LineStyle/8/1//Weight/8/-4138//LineStyle/9/1//LineStyle/10/1//LineStyle/11/1"
    
    Set rg = Cells(开始行 + 2, 开始列 + 2).Resize(1, 5)
    合并设置 rg, strProjName, 12, True, xlLeft
    边框设置 rg, "//LineStyle/7/1//LineStyle/8/1//Weight/8/-4138//LineStyle/9/1//LineStyle/10/1//Weight/10/-4138//LineStyle/11/1"
    
    Set rg = Cells(开始行 + 3, 开始列).Resize(1, 2)
    合并设置 rg, "合同名称", 12, True, xlCenter
    边框设置 rg, "//LineStyle/7/1//Weight/7/-4138//LineStyle/8/1//LineStyle/9/1//LineStyle/10/1//LineStyle/11/1"
    
    Set rg = Cells(开始行 + 3, 开始列 + 2).Resize(1, 5)
    合并设置 rg, strContractName, 12, True, xlLeft
    边框设置 rg, "//LineStyle/7/1//LineStyle/8/1//LineStyle/9/1//LineStyle/10/1//Weight/10/-4138//LineStyle/11/1"
    
    Set rg = Cells(开始行 + 4, 开始列).Resize(1, 2)
    合并设置 rg, "项目承担单位", 12, True, xlCenter
    边框设置 rg, "//LineStyle/7/1//Weight/7/-4138//LineStyle/8/1//LineStyle/9/1//LineStyle/10/1//LineStyle/11/1"
    
    Set rg = Cells(开始行 + 4, 开始列 + 2).Resize(1, 5)
    合并设置 rg, strUnitName, 12, True, xlLeft
    边框设置 rg, "//LineStyle/7/1//LineStyle/8/1//LineStyle/9/1//LineStyle/10/1//Weight/10/-4138//LineStyle/11/1"
    
    '支付金额
    Set rg = Cells(开始行 + 5, 开始列).Resize(1, 2)
    合并设置 rg, "支付金额", 12, True, xlCenter
    边框设置 rg, "//LineStyle/7/1//Weight/7/-4138//LineStyle/8/1//LineStyle/9/1//LineStyle/10/1//LineStyle/11/1"
    
    Set rg = Cells(开始行 + 5, 开始列 + 2).Resize(1, 3)
    合并设置 rg, CStr(dblPayMoney), 12, True, xlLeft, "0.000000万元"
    边框设置 rg, "//LineStyle/7/1//LineStyle/8/1//LineStyle/9/1//LineStyle/10/1//LineStyle/11/1"
    
    Cells(开始行 + 5, 开始列 + 5).Value = "大写"
    文字格式 Cells(开始行 + 5, 开始列 + 5), 12, xlCenter
    Cells(开始行 + 5, 开始列 + 6).FormulaR1C1 = "=IF((RC[-4])*10000<0,""负"","""")&SUBSTITUTE(TEXT(TRUNC(ROUND((RC[-4])*10000,2)),""[DBNum2]"")&""元""&IF(ISNUMBER(FIND(""."",ROUND((RC[-4])*10000,2))),TEXT(RIGHT(TRUNC(ROUND((RC[-4])*10000,2)*10)),""[DBNum2]"")&IF(ISNUMBER(FIND("".0"",ROUND((RC[-4])*10000,2))),"""",""角""),"""")&IF(LEFT(RIGHT(TRUNC(ROUND((RC[-4])*10000,2),2),3),1)=""."",TEXT(RIGHT(ROUND((RC[-4])*10000,2)),""[DBNum2]"")&""分"",""整""),""-"",)"
    文字格式 Cells(开始行 + 5, 开始列 + 6), 12, xlCenter, "0.000000万元"
    
    '支付依据
    Set rg = Cells(开始行 + 6, 开始列).Resize(1, 2)
    合并设置 rg, "支付依据", 12, True, xlCenter
    边框设置 rg, "//LineStyle/7/1//Weight/7/-4138//LineStyle/8/1//LineStyle/9/1//LineStyle/10/1//LineStyle/11/1"
    
    Set rg = Cells(开始行 + 6, 开始列 + 2).Resize(1, 3)
    合并设置 rg, strPayBasis, 12, True, xlLeft
    边框设置 rg, "//LineStyle/7/1//LineStyle/8/1//LineStyle/9/1//LineStyle/10/1//LineStyle/11/1"
    
    Cells(开始行 + 6, 开始列 + 5).Value = "合同号"
    文字格式 Cells(开始行 + 6, 开始列 + 5), 12, xlCenter
    Cells(开始行 + 6, 开始列 + 6).Value = strContractNo
    文字格式 Cells(开始行 + 6, 开始列 + 6), 12, xlCenter
    
    '支付内容
    Set rg = Cells(开始行 + 7, 开始列).Resize(1, 2)
    合并设置 rg, "支付内容", 12, True, xlCenter
    边框设置 rg, "//LineStyle/7/1//Weight/7/-4138//LineStyle/8/1//LineStyle/9/1//LineStyle/10/1//LineStyle/11/1"
    
    Set rg = Cells(开始行 + 7, 开始列 + 2).Resize(1, 5)
    合并设置 rg, strPayDesc, 12, True, xlLeft
    边框设置 rg, "//LineStyle/7/1//LineStyle/8/1//LineStyle/9/1//LineStyle/10/1//Weight/10/-4138//LineStyle/11/1"
    
    '申请单位栏
    Set rg = Cells(开始行 + 8, 开始列).Resize(6, 1)
    合并设置 rg, "申请单位(盖章)", 12, True, xlCenter
    边框设置 rg, "//LineStyle/7/1//Weight/7/-4138//LineStyle/8/1//LineStyle/9/1//LineStyle/10/1//LineStyle/12/1"
    
    Set rg = Cells(开始行 + 8, 开始列 + 1).Resize(1, 3)
    合并设置 rg, "工程费开户名称", 12, True, xlLeft
    边框设置 rg, "//LineStyle/7/1//LineStyle/8/1//LineStyle/9/1//LineStyle/10/1//LineStyle/11/1"
    Set rg = Cells(开始行 + 8, 开始列 + 4).Resize(1, 3)
    合并设置 rg, strAccName1, 12, True, xlLeft
    边框设置 rg, "//LineStyle/7/1//LineStyle/8/1//LineStyle/9/1//LineStyle/10/1//Weight/10/-4138//LineStyle/11/1"
    
    Set rg = Cells(开始行 + 9, 开始列 + 1).Resize(1, 3)
    合并设置 rg, "工程费开户行", 12, True, xlLeft
    边框设置 rg, "//LineStyle/7/1//LineStyle/8/1//LineStyle/9/1//LineStyle/10/1//LineStyle/11/1"
    Cells(开始行 + 9, 开始列 + 4).Value = strBankName1
    文字格式 Cells(开始行 + 9, 开始列 + 4), 12, xlLeft
    Cells(开始行 + 9, 开始列 + 5).Value = "账号"
    文字格式 Cells(开始行 + 9, 开始列 + 5), 12, xlLeft
    Cells(开始行 + 9, 开始列 + 6).NumberFormat = "@"
    Cells(开始行 + 9, 开始列 + 6).Value = strAccNo1
    文字格式 Cells(开始行 + 9, 开始列 + 6), 12, xlLeft
    
    Set rg = Cells(开始行 + 10, 开始列 + 1).Resize(1, 3)
    合并设置 rg, "农民工工资开户名称", 12, True, xlLeft
    边框设置 rg, "//LineStyle/7/1//LineStyle/8/1//LineStyle/9/1//LineStyle/10/1//LineStyle/11/1"
    Set rg = Cells(开始行 + 10, 开始列 + 4).Resize(1, 3)
    合并设置 rg, strAccName2, 12, True, xlLeft
    边框设置 rg, "//LineStyle/7/1//LineStyle/8/1//LineStyle/9/1//LineStyle/10/1//Weight/10/-4138//LineStyle/11/1"
    
    Set rg = Cells(开始行 + 11, 开始列 + 1).Resize(1, 3)
    合并设置 rg, "农民工工资开户行", 12, True, xlLeft
    边框设置 rg, "//LineStyle/7/1//LineStyle/8/1//LineStyle/9/1//LineStyle/10/1//LineStyle/11/1"
    Cells(开始行 + 11, 开始列 + 4).Value = strBankName2
    文字格式 Cells(开始行 + 11, 开始列 + 4), 12, xlLeft
    Cells(开始行 + 11, 开始列 + 5).Value = "账号"
    文字格式 Cells(开始行 + 11, 开始列 + 5), 12, xlLeft
    Cells(开始行 + 11, 开始列 + 6).NumberFormat = "@"
    Cells(开始行 + 11, 开始列 + 6).Value = strAccNo2
    文字格式 Cells(开始行 + 11, 开始列 + 6), 12, xlLeft
    
    Set rg = Cells(开始行 + 12, 开始列 + 1).Resize(1, 3)
    合并设置 rg, "经办人", 12, True, xlLeft
    边框设置 rg, "//LineStyle/7/1//LineStyle/8/1//LineStyle/9/1//LineStyle/10/1//LineStyle/11/1"
    Cells(开始行 + 12, 开始列 + 4).Value = ""
    文字格式 Cells(开始行 + 12, 开始列 + 4), 12, xlLeft
    Cells(开始行 + 12, 开始列 + 5).Value = "联系方式"
    文字格式 Cells(开始行 + 12, 开始列 + 5), 12, xlLeft
    Cells(开始行 + 12, 开始列 + 6).Value = ""
    文字格式 Cells(开始行 + 12, 开始列 + 6), 12, xlLeft
    
    Set rg = Cells(开始行 + 13, 开始列 + 1).Resize(1, 3)
    合并设置 rg, "单位负责人", 12, True, xlLeft
    边框设置 rg, "//LineStyle/7/1//LineStyle/8/1//LineStyle/9/1//LineStyle/10/1//LineStyle/11/1"
    Cells(开始行 + 13, 开始列 + 4).Value = ""
    文字格式 Cells(开始行 + 13, 开始列 + 4), 12, xlLeft
    Cells(开始行 + 13, 开始列 + 5).Value = "日期"
    文字格式 Cells(开始行 + 13, 开始列 + 5), 12, xlLeft
    Cells(开始行 + 13, 开始列 + 6).Value = ""
    文字格式 Cells(开始行 + 13, 开始列 + 6), 12, xlLeft
    
    '设计院意见
    Set rg = Cells(开始行 + 14, 开始列).Resize(2, 7)
    合并设置 rg, strAuditOpinion, 12, True, xlLeft
    边框设置 rg, "//LineStyle/7/1//Weight/7/-4138//LineStyle/8/1//LineStyle/9/1//LineStyle/10/1//Weight/10/-4138//LineStyle/11/1//LineStyle/12/1"
    
    '下属公司意见
    Set rg = Cells(开始行 + 16, 开始列).Resize(3, 1)
    合并设置 rg, "下属公司意见", 12, True, xlCenter
    边框设置 rg, "//LineStyle/7/1//Weight/7/-4138//LineStyle/8/1//LineStyle/9/1//LineStyle/10/1//LineStyle/12/1"
    
    Set rg = Cells(开始行 + 16, 开始列 + 1).Resize(1, 2)
    合并设置 rg, "合约部门", 12, True, xlLeft
    边框设置 rg, "//LineStyle/7/1//LineStyle/8/1//LineStyle/9/1//LineStyle/10/1//LineStyle/11/1"
    Set rg = Cells(开始行 + 16, 开始列 + 3).Resize(1, 4)
    合并设置 rg, "经办人：                    分管领导：", 12, True, xlLeft
    边框设置 rg, "//LineStyle/7/1//LineStyle/8/1//LineStyle/9/1//LineStyle/10/1//Weight/10/-4138//LineStyle/11/1"
    
    Set rg = Cells(开始行 + 17, 开始列 + 1).Resize(1, 2)
    合并设置 rg, "财务部门", 12, True, xlLeft
    边框设置 rg, "//LineStyle/7/1//LineStyle/8/1//LineStyle/9/1//LineStyle/10/1//LineStyle/11/1"
    Set rg = Cells(开始行 + 17, 开始列 + 3).Resize(1, 4)
    合并设置 rg, "经办人：                    分管领导：", 12, True, xlLeft
    边框设置 rg, "//LineStyle/7/1//LineStyle/8/1//LineStyle/9/1//LineStyle/10/1//Weight/10/-4138//LineStyle/11/1"
    
    Set rg = Cells(开始行 + 18, 开始列 + 1).Resize(1, 2)
    合并设置 rg, "公司负责人", 12, True, xlLeft
    边框设置 rg, "//LineStyle/7/1//LineStyle/8/1//LineStyle/9/1//LineStyle/10/1//LineStyle/11/1"
    Set rg = Cells(开始行 + 18, 开始列 + 3).Resize(1, 4)
    合并设置 rg, "", 12, True, xlLeft
    边框设置 rg, "//LineStyle/7/1//LineStyle/8/1//LineStyle/9/1//LineStyle/10/1//Weight/10/-4138//LineStyle/11/1"
    
    '集团意见
    Set rg = Cells(开始行 + 19, 开始列).Resize(5, 1)
    合并设置 rg, "集团公司意见", 12, True, xlCenter
    边框设置 rg, "//LineStyle/7/1//Weight/7/-4138//LineStyle/8/1//LineStyle/9/1//Weight/9/-4138//LineStyle/10/1//LineStyle/12/1"
    
    Set rg = Cells(开始行 + 19, 开始列 + 1).Resize(1, 2)
    合并设置 rg, "下属公司" & Chr(10) & "分管领导", 12, True, xlCenter
    边框设置 rg, "//LineStyle/7/1//LineStyle/8/1//LineStyle/9/1//LineStyle/10/1//LineStyle/11/1"
    Set rg = Cells(开始行 + 19, 开始列 + 3).Resize(1, 4)
    合并设置 rg, "", 12, True, xlLeft
    边框设置 rg, "//LineStyle/7/1//LineStyle/8/1//LineStyle/9/1//LineStyle/10/1//Weight/10/-4138//LineStyle/11/1"
    
    Set rg = Cells(开始行 + 20, 开始列 + 1).Resize(1, 2)
    合并设置 rg, "财务部", 12, True, xlCenter
    边框设置 rg, "//LineStyle/7/1//LineStyle/8/1//LineStyle/9/1//LineStyle/10/1//LineStyle/11/1"
    Set rg = Cells(开始行 + 20, 开始列 + 3).Resize(1, 4)
    合并设置 rg, "经办人：           复核：              负责人：", 12, True, xlLeft
    边框设置 rg, "//LineStyle/7/1//LineStyle/8/1//LineStyle/9/1//LineStyle/10/1//Weight/10/-4138//LineStyle/11/1"
    
    Set rg = Cells(开始行 + 21, 开始列 + 1).Resize(1, 2)
    合并设置 rg, "财务部" & Chr(10) & "分管领导", 12, True, xlCenter
    边框设置 rg, "//LineStyle/7/1//LineStyle/8/1//LineStyle/9/1//LineStyle/10/1//LineStyle/11/1"
    Set rg = Cells(开始行 + 21, 开始列 + 3).Resize(1, 4)
    合并设置 rg, "", 12, True, xlLeft
    边框设置 rg, "//LineStyle/7/1//LineStyle/8/1//LineStyle/9/1//LineStyle/10/1//Weight/10/-4138//LineStyle/11/1"
    
    Set rg = Cells(开始行 + 22, 开始列 + 1).Resize(1, 2)
    合并设置 rg, "总经理", 12, True, xlCenter
    边框设置 rg, "//LineStyle/7/1//LineStyle/8/1//LineStyle/9/1//LineStyle/10/1//LineStyle/11/1"
    Set rg = Cells(开始行 + 22, 开始列 + 3).Resize(1, 4)
    合并设置 rg, "", 12, True, xlLeft
    边框设置 rg, "//LineStyle/7/1//LineStyle/8/1//LineStyle/9/1//LineStyle/10/1//Weight/10/-4138//LineStyle/11/1"
    
    Set rg = Cells(开始行 + 23, 开始列 + 1).Resize(1, 2)
    合并设置 rg, "董事长", 12, True, xlCenter
    边框设置 rg, "//LineStyle/7/1//LineStyle/8/1//LineStyle/9/1//Weight/9/-4138//LineStyle/10/1//LineStyle/11/1"
    Set rg = Cells(开始行 + 23, 开始列 + 3).Resize(1, 4)
    合并设置 rg, "", 12, True, xlLeft
    边框设置 rg, "//LineStyle/7/1//LineStyle/8/1//LineStyle/9/1//Weight/9/-4138//LineStyle/10/1//Weight/10/-4138//LineStyle/11/1"
    
    Set rg = Cells(开始行, 开始列).Range("$F$6,$F$7,$E$10,$F$10,$E$12,$F$12,$E$13,$F$13,$E$14,$F$14")
    边框设置 rg, "//LineStyle/7/1//LineStyle/8/1//LineStyle/9/1//LineStyle/10/1"
    Set rg = Cells(开始行, 开始列).Range("$G$6,$G$7,$G$10,$G$12,$G$13,$G$14")
    边框设置 rg, "//LineStyle/7/1//LineStyle/8/1//LineStyle/9/1//LineStyle/10/1//Weight/10/-4138"

错误恢复:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    If Err.Number <> 0 Then MsgBox "运行出错：" & Err.Description, vbExclamation
End Sub

'封装函数
Private Sub 合并设置(rg As Range, 内容 As String, 字号 As Integer, 自动换行 As Boolean, 对齐方式 As Integer, Optional 数字格式 As String = "General")
    With rg
        .MergeCells = True
        .Value = 内容
        .HorizontalAlignment = 对齐方式
        .VerticalAlignment = xlCenter
        .Font.Name = "仿宋_GB2312"
        .Font.Size = 字号
        .Font.Color = 0
        .WrapText = 自动换行
        .NumberFormat = 数字格式
    End With
End Sub

Private Sub 文字格式(rg As Range, 字号 As Integer, 对齐方式 As Integer, Optional 数字格式 As String = "General")
    With rg
        .HorizontalAlignment = 对齐方式
        .VerticalAlignment = xlCenter
        .Font.Name = "仿宋_GB2312"
        .Font.Size = 字号
        .Font.Color = 0
        .WrapText = True
        .NumberFormat = 数字格式
    End With
End Sub

Private Sub 边框设置(rg As Range, 边框字符串 As String)
    Dim arr As Variant, brr As Variant, i As Integer
    arr = Split(边框字符串, "//")
    With rg
        For i = LBound(arr) To UBound(arr)
            If Len(arr(i)) > 0 Then
                brr = Split(arr(i), "/")
                Select Case True
                    Case InStr(arr(i), "LineStyle") > 0
                        .Borders(CInt(brr(1))).LineStyle = CInt(brr(2))
                    Case InStr(arr(i), "Weight") > 0
                        .Borders(CInt(brr(1))).Weight = CInt(brr(2))
                    Case InStr(arr(i), "Color") > 0
                        .Borders(CInt(brr(1))).Color = CLng(brr(2))
                End Select
            End If
        Next i
    End With
End Sub
