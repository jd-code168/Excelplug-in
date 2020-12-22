Imports Microsoft.Office.Tools.Ribbon

Public Class addtools

    Private Sub addtools_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Add_Btn_Click(sender As Object, e As RibbonControlEventArgs) Handles Add_Btn.Click
        Run.run()
    End Sub

    Private Sub Delete_Btn_Click(sender As Object, e As RibbonControlEventArgs) Handles Delete_Btn.Click

        If delWin Is Nothing Then
            delWin = New DelForm
            delWin.Show()
        End If

        'delWin.Top = app.ActiveWindow.PointsToScreenPixelsX(app.Left)
        'delWin.Left = app.ActiveWindow.PointsToScreenPixelsY(app.Top)
    End Sub

    Private Sub Init_Btn_Click(sender As Object, e As RibbonControlEventArgs) Handles Init_Btn.Click
        app.DisplayAlerts = False

        If MsgBox("是否清空原有数据?", 1, "提示") <> vbOK Then
            GoTo eline
        End If

        If InitModule.initModule() Then
            MsgBox("重置成功！", 0, "提示")
        Else
            MsgBox("重置失败，或已无需重置！", 0, "提示")
        End If
eline:
        app.DisplayAlerts = True
    End Sub

    Private Sub Create_Btn_Click(sender As Object, e As RibbonControlEventArgs) Handles Create_Btn.Click
        app.DisplayAlerts = False
        If CreateSettlement.create() Then
            MsgBox("生成汇总表成功！"， 0, "提示")
        Else
            MsgBox("生成汇总表失败，请选择结算模板表生成！"， 0, "提示")
        End If
        app.DisplayAlerts = True
    End Sub

    Private Sub DelFill_Btn_Click(sender As Object, e As RibbonControlEventArgs) Handles DelFill_Btn.Click
        If app.ActiveSheet.range("O1").value <> "2020驻地网模板" Then

            MsgBox("请选择《主要工程量表》再选择清空！", 0, "错误")

            Exit Sub

        End If

        DeleteRecord.clearFill("E4:E136")
    End Sub

    Private Sub Fix_Btn_Click(sender As Object, e As RibbonControlEventArgs) Handles Fix_Btn.Click
        Dim site As String

        If app.ActiveSheet.range("O1").value <> "2020驻地网模板" Then

            MsgBox("请选择《主要工程量表》再进行修改！", 0, "错误")

            Exit Sub

        End If

        site = app.ActiveSheet.range("B2").value

        If FixModule.fixRecord(site) Then
            MsgBox("《" & site & "》工作量修改成功！", 0, "提示")
        Else
            MsgBox("《" & site & "》工作量修改失败！请确认工程名是否正确或结算记录是否存在？", 0, "提示")
        End If
    End Sub
End Class
