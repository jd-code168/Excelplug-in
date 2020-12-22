Module FindModule
    '查找字符串指定表格，返回 worksheet对象类型
    Function findWorksheet(sheetName As String) As Excel.Worksheet
        Dim ws As Excel.Worksheet
        Dim wb As Excel.Workbook = app.ActiveWorkbook

        For Each ws In wb.Worksheets 'app.ActiveWorkbook.Worksheets
            If ws.Name = sheetName Then
                Exit For
            Else
                ws = Nothing
            End If
        Next
        findWorksheet = ws
        wb = Nothing
        ws = Nothing
    End Function

    '查找指定站点序号，返回单元格索引
    Function findSite(No As Integer) As Excel.Range
        Dim ncols As Integer
        Dim isFinded As Boolean = False
        Dim i As Integer

        On Error GoTo err

        With app.ActiveWorkbook.Worksheets("表2-通信工程送审工程量明细表")
            ncols = .cells(3, .columns.Count).end(1).column

            For i = 1 To ncols

                If Val(Mid(.cells(3, i).value, 3)) = No Then
                    'MsgBox("finded：" & i)
                    isFinded = True
                    Exit For
                End If
            Next

            If isFinded Then
                findSite = .cells(3, i)
            End If

        End With

        Exit Function
err:
        findSite = Nothing

    End Function
End Module
