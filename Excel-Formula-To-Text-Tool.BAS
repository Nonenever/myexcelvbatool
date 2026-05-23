Option Explicit
REM Excel 公式转文本说明工具
'---------------------------------------------------------------------------------------
' 主函数：组合文本结果 = 左侧可见单元格计数 + 计算表达式文本
' 参数：SourceCell  源单元格（提供公式）
'       StartCell   可选，左侧统计起始单元格（默认A列）
' 返回：拼接后的字符串
'---------------------------------------------------------------------------------------
Function CombineText(SourceCell As Range, Optional StartCell As Range) As String
    CombineText = CountLeftWidth(StartCell) & GetCalcText(SourceCell)
End Function

'---------------------------------------------------------------------------------------
' 功能：统计【起始列 ~ 当前单元格左侧】所有**列宽>0（可见）**的单元格数量 + 1
' 规则：包含当前列，同一行内统计
' 参数：StartCell 可选，指定统计起始列单元格（默认从A列开始）
' 返回：可见单元格总数（Long）
'---------------------------------------------------------------------------------------
Function CountLeftWidth(Optional StartCell As Range) As Long
    Dim currRow As Long       ' 公式所在单元格 行号
    Dim currCol As Long       ' 公式所在单元格 列号
    Dim startCol As Long      ' 统计起始列号
    Dim endCol As Long        ' 统计结束列号（当前列左侧一列）
    Dim i As Long             ' 循环变量
    Dim countNum As Long      ' 可见单元格计数器
    
    ' 获取公式所在单元格的坐标
    currRow = Application.Caller.Row
    currCol = Application.Caller.Column
    
    ' 边界处理：当前是A列，直接返回1
    If currCol = 1 Then
        CountLeftWidth = 1
        Exit Function
    End If
    
    ' 结束列 = 当前列左侧一列
    endCol = currCol - 1
    
    ' 确定统计起始列
    If Not StartCell Is Nothing Then
        startCol = StartCell.Column  ' 使用用户指定列
    Else
        startCol = 1                 ' 默认从A列开始
    End If
    
    ' 遍历统计：列宽>0 表示可见
    countNum = 1  ' 初始值=1（包含当前列）
    For i = startCol To endCol
        If Cells(currRow, i).Width > 0 Then
            countNum = countNum + 1
        End If
    Next i
    
    CountLeftWidth = countNum
End Function

'---------------------------------------------------------------------------------------
' 功能：解析源单元格公式，替换单元格地址为【当前行对应列=号前文本】
' 处理：自动移除ROUND函数、替换单元格引用、提取等号前文本
' 参数：SourceCell  提供公式的源单元格
' 返回：处理后的计算表达式文本
'---------------------------------------------------------------------------------------
Function GetCalcText(SourceCell As Range) As String
    Dim currRow As Long       ' 公式所在行号
    Dim rawFormula As String  ' 源单元格原始公式
    Dim regEx As Object       ' 正则表达式对象
    Dim matches As Object     ' 匹配结果集合
    Dim i As Long             ' 循环变量
    Dim cellAddr As String    ' 匹配到的单元格地址
    Dim colPart As String     ' 单元格地址的列部分（如 A、$C）
    Dim replaceText As String ' 要替换成的文本
    
    ' 获取当前公式所在行（动态跟随公式位置）
    currRow = Application.Caller.Row
    
    ' 判断源单元格是否包含公式
    If Not SourceCell.HasFormula Then
        GetCalcText = ""
        Exit Function
    End If
    
    ' 步骤1：移除公式中所有 ROUND 函数
    rawFormula = DelRoundFunc(SourceCell.Formula)
    
    ' 步骤2：创建正则对象，匹配单元格地址（A1、$B$2、AB123格式）
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True       ' 全局匹配
    regEx.IgnoreCase = True   ' 不区分大小写
    regEx.Pattern = "(\$?[A-Z]{1,3})(\$?\d+)"  ' 分组：列 + 行
    
    Set matches = regEx.Execute(rawFormula)
    
    ' 步骤3：倒序替换，避免地址错乱
    If matches.Count >= 1 Then
        For i = matches.Count - 1 To 0 Step -1
            cellAddr = matches(i).Value       ' 完整地址：如 $D5
            colPart = matches(i).SubMatches(0) ' 列部分：如 $D
            
            ' 获取当前行、对应列的值，并提取等号前文本
            replaceText = Range(colPart & currRow).Value
            replaceText = GetStrBeforeEqual(replaceText)
            
            ' 替换公式中的单元格地址
            rawFormula = Replace(rawFormula, cellAddr, replaceText)
        Next i
    Else
        rawFormula = ""
    End If
    
    GetCalcText = rawFormula
End Function

'---------------------------------------------------------------------------------------
' 功能：递归移除公式中所有 ROUND( 表达式, 小数位数 ) 格式
' 返回：只保留内部计算式
' 例：ROUND(A1+B1,2) → A1+B1
'---------------------------------------------------------------------------------------
Function DelRoundFunc(sFormula As String) As String
    Dim regEx As Object
    Dim oldStr As String
    Dim newStr As String
    
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.IgnoreCase = True
    
    oldStr = sFormula
    
    ' 循环替换，直到没有ROUND函数（支持嵌套）
    Do
        regEx.Pattern = "ROUND\((.+?),\s*\d+\)"
        newStr = regEx.Replace(oldStr, "$1")  ' 保留括号内第一个参数
        
        If newStr = oldStr Then Exit Do  ' 无替换则退出
        oldStr = newStr
    Loop
    
    DelRoundFunc = oldStr
End Function

'---------------------------------------------------------------------------------------
' 功能：提取字符串中【第一个等号 = 左侧的文本】
' 例："数量=10" → "数量"；"测试文本" → "测试文本"
'---------------------------------------------------------------------------------------
Function GetStrBeforeEqual(s As String) As String
    Dim trimStr As String
    Dim eqPos As Long
    
    trimStr = Trim(s)  ' 去除首尾空格
    eqPos = InStr(trimStr, "=")
    
    If eqPos > 0 Then
        GetStrBeforeEqual = Left(trimStr, eqPos - 1)
    Else
        GetStrBeforeEqual = trimStr
    End If
End Function
