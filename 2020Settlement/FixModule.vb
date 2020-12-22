Module FixModule
    Function fixRecord(site As String) As Boolean
        Dim ws As Excel.Worksheet
        Dim siteNo As Integer = -1
        Dim rng As Excel.Range
        Dim ncol As Integer

        On Error GoTo err

        For Each ws In app.ActiveWorkbook.Worksheets
            If ws.Name <> "主要工程量表" And ws.Name <> "表1-工程结算表（单价包干）模板" Then
                If ws.Range("A3").Value = "工程名称：" & site Then
                    siteNo = Replace(ws.Name, "表1-工程结算表（单价包干）", "")
                    Exit For
                End If
            End If
        Next

        If siteNo = -1 Then
            fixRecord = False
            Exit Function
        End If

        rng = FindModule.findSite(siteNo)

        If rng Is Nothing Then
            MsgBox("《表2-通信工程送审工程量明细表》中查找不到相应站点信息,
                      请确认[表2]或站点信息是否存在？", 0, "错误")
        End If

        Dim arr As Array

        ncol = rng.Column
        arr = app.ActiveSheet.Range("N5:N66").value

        With app.ActiveWorkbook.Worksheets("表2-通信工程送审工程量明细表")
            .cells(5, ncol).resize(62) = arr
        End With
        fixRecord = True
err:
        ws = Nothing
        Erase arr
    End Function
End Module
