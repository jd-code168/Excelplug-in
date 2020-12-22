Option Explicit On

Imports dov = System.Windows.Forms

Module Run
    Sub run()
        'On Error Resume Next
        app.DisplayAlerts = False
        Dim Wname As String = "2020驻地网模板"   '模板标识
        Dim Sname As String = "表2-通信工程送审工程量明细表" '自定义工作表名称
        Dim ws As Excel.Worksheet = Nothing
        Dim ncols As Integer
        Dim addColNum As Integer

        If app.ActiveSheet.range("O1").value <> Wname Then

            MsgBox("请选择《主要工程量表》再进行添加！", 0, "错误")
            dov.Application.DoEvents()
            Exit Sub

        End If

        If app.WorksheetFunction.Sum(app.ActiveSheet.range("E4:E136").value) = 0 Then
            MsgBox("请先填写工作量！", 0, "错误")
            dov.Application.DoEvents()
            Exit Sub
        End If

        If app.ActiveSheet.range("B2").value = "" Then

            MsgBox("工程名称不能为空！请在B2填写工程名！", 0, "错误")
            dov.Application.DoEvents()
            Exit Sub

        End If

        ws = FindModule.findWorksheet(Sname)

        If ws Is Nothing Then
            MsgBox("没有找到《" & Sname & "》,请确保表格存在！", 0, "提示")
            Exit Sub
        End If

        ncols = ws.Cells(3, ws.Columns.Count).end(1).column
        addColNum = ncols '+ 1   '要添加记录的列号

        If addRecord(ws, addColNum) Then
            Call setTableStyle(ws, addColNum)
            MsgBox("添加工作量成功！", 0, "提示")
        Else
            MsgBox("添加工作量失败，请重新添加！", 0, "提示")
        End If
        ws = Nothing
        app.DisplayAlerts = True
    End Sub

    Function addRecord(ByRef ws As Excel.Worksheet, addColNum As Integer) As Boolean '添加记录
        Dim arr As Array
        Dim Nums As Array
        Dim No As Integer
        Dim addDataed As Boolean

        On Error GoTo err
        If addColNum < 7 Then
            MsgBox("《表2-通信工程送审工程量明细表》格式有误！"， 0, "提示")
            Exit Function
        End If

        arr = app.ActiveSheet.Range("N3:N66").value
        Nums = Split((ws.Cells(3, addColNum - 1).value), "站点"）

        If Nums.Length = 1 Then
            If addColNum = 7 Then
                No = 1
            Else
                MsgBox("《表2-通信工程送审工程量明细表》最后一列站点序号有误！"， 0, "提示")
                'app.StatusBar = "结算送审工程量明细表最后一列站点序号有误！"
                Exit Function
            End If

        Else
            If Nums(1) = "" Then
                MsgBox("《表2-通信工程送审工程量明细表》最后一列站点序号有误！"， 0, "提示")
                'app.StatusBar = "结算送审工程量明细表最后一列站点序号有误！"
                Exit Function
            End If
            No = Nums(1) + 1
        End If

        With ws
            ws.Cells.Columns(addColNum).Insert
            ws.Cells(3, addColNum).resize(UBound(arr), 1) = arr
            ws.Cells(3, addColNum).value = arr(1, 1) & No
            Call copySheet(No)
            addDataed = addData(No, addColNum)
            If Not addDataed Then
                GoTo err
            End If
        End With
        addRecord = True
        Exit Function
err:
        ws.Cells.Columns(addColNum).delete
        addRecord = False
    End Function

    Sub copySheet(No As Integer)
        Dim ws As Excel.Worksheet
        Dim nws As Excel.Worksheet
        Dim Sname As String = "表1-工程结算表（单价包干）模板"
        Dim addSheetName As String = "表2-通信工程送审工程量明细表"

        ws = FindModule.findWorksheet(Sname)

        If ws Is Nothing Then
            MsgBox("没有找到《" & Sname & "》,请确保表格存在！", 0, "提示")
            Exit Sub
        End If

        nws = FindModule.findWorksheet("表1-工程结算表（单价包干）" & No)

        If nws IsNot Nothing Then
            nws.Activate()
            If MsgBox("站点" & No & "结算表已存在!是否删除原有结算表？", 1, "提示") = vbOK Then
                nws.Delete()
                app.ActiveWorkbook.Sheets("主要工程量表").Activate()
            Else
                Error 1
            End If
        End If

        ws.Copy(Before:=app.ActiveWorkbook.Sheets(addSheetName))

        app.ActiveWorkbook.Sheets(app.ActiveWorkbook.Sheets(addSheetName).index - 1).name = "表1-工程结算表（单价包干）" & No
        app.ActiveWorkbook.Worksheets("表1-工程结算表（单价包干）" & No).visible = True
        app.ActiveWorkbook.Sheets("主要工程量表").Select
        'app.Worksheets
        'app.ThisWorkbook.Sheets.Add() 'app.ActiveWorkbook.Worksheets.Add(Count:=1)
        'nws = app.ActiveWorkbook.Worksheets.Add(ws)

        ws = Nothing
        nws = Nothing
    End Sub

    Function addData(No As Integer, ncol As Integer) As Boolean
        Dim ws As Excel.Worksheet
        Dim sourceWs As Excel.Worksheet
        Dim letter As String
        Dim formulaStr As String
        Dim i As Integer

        On Error GoTo err

        ws = app.ActiveWorkbook.Worksheets("表1-工程结算表（单价包干）" & No)
        sourceWs = app.ActiveWorkbook.Sheets("主要工程量表")
        letter = app.Cells(1, ncol).address 'Chr((70 + No))


        ws.Cells(3, 1).value = "工程名称：" & sourceWs.Cells(2, 2).value
        ws.Cells(3, 6).value = "施工合同编号：" & sourceWs.Cells(1, 2).value

        For i = 6 To 67
            formulaStr = "='表2-通信工程送审工程量明细表'!" & Left(letter, letter.Length - 1) & (i - 1)
            ws.Cells(i, 6).Formula = formulaStr
        Next


        addData = True
        ws = Nothing
        Exit Function
err:
        ws.Delete()
        ws = Nothing
        addData = False
    End Function

    Sub setTableStyle(ByRef ws As Excel.Worksheet, colNum As Integer)
        Dim nrows As Integer

        nrows = ws.Cells(ws.Rows.Count, colNum).end(3).row
        With ws.Cells(3, colNum).resize(nrows - 2, 1)
            .Borders.LineStyle = Excel.XlLineStyle.xlContinuous
            .Font.Name = "华文仿宋"
            .Font.Size = 12
            .ColumnWidth = 11.5
            .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            .NumberFormat = "0.000"
        End With

    End Sub
End Module
