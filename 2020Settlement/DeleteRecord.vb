'Option Explicit On
Module DeleteRecord
    Sub delete(deletName As String)
        Dim wname As String = deletName
        Dim ws As Excel.Worksheet

        app.DisplayAlerts = False

        On Error GoTo err

        ws = FindModule.findWorksheet("表1-工程结算表（单价包干）" & wname)

        If isDeleted(ws, wname) Then
            MsgBox("删除站点结算成功!")
        Else
            MsgBox("删除失败！请确保记录存在或站点序号正确！")
        End If

err:
        ws = Nothing
        app.DisplayAlerts = True
    End Sub

    Function isDeleted(ByRef ws As Excel.Worksheet, deleteName As String) As Boolean '删除站点记录函数返回Bool值

        Dim deleteRng As Excel.Range = Nothing
        Dim isSuccessed As Boolean = False

        On Error GoTo err

        If ws IsNot Nothing Then
            ws.Delete()
            isSuccessed = True
        End If

        deleteRng = FindModule.findSite(deleteName)

        If deleteRng IsNot Nothing Then

            'MsgBox(deleteRng.Column)
            app.ActiveWorkbook.Worksheets("表2-通信工程送审工程量明细表") _
                .Columns(deleteRng.Column).Delete

            isSuccessed = True
        End If

        If isSuccessed Then
            isDeleted = True
        Else
            GoTo err
        End If

        Exit Function
err:
        isDeleted = False
    End Function

    Sub clearFill(rng As String)    '清空指定范围单元格
        Dim strRng As String = rng

        On Error GoTo err

        If strRng.Length < 5 Then
            GoTo err
        End If

        app.ActiveSheet.range(strRng).ClearContents

        MsgBox(strRng & "删除成功！", 0, "提示")
        Exit Sub
err:
        MsgBox(strRng & "删除失败！", 0, "错误")

    End Sub
End Module
