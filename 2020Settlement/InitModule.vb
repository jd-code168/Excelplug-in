Option Explicit On

Module InitModule
    Function initModule() As Boolean
        Dim ncols As Integer
        Dim i As Integer = 1
        Dim delArea As String
        Dim ws As Excel.Worksheet
        Dim wsname() As String = {"表2-通信工程送审工程量明细表", "主要工程量表", "表1-工程结算表（单价包干）模板"}

        On Error GoTo err

        With app.ActiveWorkbook
            If .Worksheets("主要工程量表").range("O1").value <> "2020驻地网模板" Then
                GoTo err
            End If

            ncols = .Worksheets(wsname(0)).cells(3, .Worksheets(wsname(0)).columns.count).end(1).column

            If ncols <= 7 Then
                GoTo err
            End If

            delArea = "G:" & Split(app.Cells(1, (ncols - 1)).address, "$")(1) 'Chr(ncols + 63)

            .Worksheets(wsname(0)).columns(delArea).Delete()
            'Do While (i <= wcount)

            '    If .Sheets(i).name <> wsname(0) And .Sheets(i).name <> wsname(1) And .Sheets(i).name <> wsname(2) Then

            '        .Sheets(i).delete

            '    End If
            '    i = i + 1
            'Loop

            For Each ws In .Worksheets
                If ws.Name <> wsname(0) And ws.Name <> wsname(1) And ws.Name <> wsname(2) Then

                    .Sheets(i).delete

                End If
            Next

        End With
        ws = Nothing
        initModule = True
        Exit Function
err:
        initModule = False
        'MsgBox("出错")
    End Function

End Module
