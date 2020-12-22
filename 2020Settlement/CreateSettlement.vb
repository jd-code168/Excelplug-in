Option Explicit On

Module CreateSettlement
    '创建汇总结算表，包括将删除《主要工程量表》、包干表隐藏模板和明细表最后一列辅助列
    Function create() As Boolean
        On Error GoTo err
        Dim ncol As Integer
        Dim fileNo As String

        With app.ActiveWorkbook

            .Worksheets("主要工程量表").delete
            .Worksheets("表1-工程结算表（单价包干）模板").delete
            ncol = .Worksheets("表2-通信工程送审工程量明细表").
                cells(3, .Worksheets("表2-通信工程送审工程量明细表").columns.count).end(1).column
            .Worksheets("表2-通信工程送审工程量明细表").columns(ncol).delete
            fileNo = Format(Date.Now, "yyyy-MM-dd hhmmss")

            .SaveAs(app.ActiveWorkbook.Path & "\2020汇总结算表" & fileNo & ".xlsx")

        End With

        create = True

        Exit Function
err:
        create = False
    End Function
End Module
