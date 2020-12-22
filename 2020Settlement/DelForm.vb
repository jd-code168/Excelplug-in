Imports System.Windows.Forms

Public Class DelForm
    Private Sub Ok_Btn_Click(sender As Object, e As EventArgs) Handles Ok_Btn.Click
        Dim delNo As String = 0
        delNo = No_Text.Text
        If No_Text.TextLength = 0 Or No_Text.Text = "0" Then
            MsgBox("请输入正确的站点序号", 0, "提示")
            Exit Sub
        End If
        If MsgBox("确定是否删除？", 1, "提示") = vbOK Then
            DeleteRecord.delete(delNo)
            No_Text.Clear()
        End If

        'Me.Close()
    End Sub

    Private Sub DelForm_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        delWin = Nothing
    End Sub

    'Private Sub DelForm_Load(sender As Object, e As EventArgs) Handles Me.Load
    '    Me.SetDesktopLocation(app.Left + 100, app.Top + 100)
    'End Sub

    Private Sub DelForm_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        delWin = Nothing
    End Sub
End Class