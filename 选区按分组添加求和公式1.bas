Sub 选区按分组添加求和公式(sel)
    Dim i, j, arr, crr() As Range
    Stop
    Sel_R_Start = sel.row
    Sel_C_Start = sel.Column
    Sel_R_End = sel.Rows.count + sel.row - 1
    Sel_C_End = sel.Columns.count + sel.Column - 1
    ReDim arr(Sel_R_Start To Sel_R_End)
    For i = LBound(arr) To UBound(arr)
        arr(i) = Rows(i).OutlineLevel
    Next

    ReDim crr(Sel_R_Start To Sel_R_End)
    For j = LBound(crr) To UBound(crr)
        For i = j + 1 To UBound(arr)
            If arr(j) = arr(i) Then
                Exit For
            Else
                If arr(j) = arr(i) - 1 Then
                    If crr(j) Is Nothing Then
                        Set crr(j) = Cells(i, Sel_C_Start)
                    Else
                        Set crr(j) = Union(crr(j), Cells(i, Sel_C_Start))
                    End If
                End If
            End If
        Next i
    Next j
    For j = LBound(crr) To UBound(crr)
        If crr(j) Is Nothing Then
        Else
            Cells(j, Sel_C_Start).formula = "=sum(" & crr(j).Address(0, 0) & ")"
            Cells(j, Sel_C_Start).Interior.colorIndex = 8
            If Sel_C_End - Sel_C_Start > 1 Then
                Cells(j, Sel_C_Start).AutoFill Destination:=Range(Cells(j, Sel_C_Start), Cells(j, Sel_C_End))
            End If
        End If
    Next j
    'Stop
End Sub
