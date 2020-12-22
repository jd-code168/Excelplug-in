<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DelForm
    Inherits System.Windows.Forms.Form

    'Form 重写 Dispose，以清理组件列表。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows 窗体设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。  
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.No_Text = New System.Windows.Forms.TextBox()
        Me.Ok_Btn = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 17)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(65, 12)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "站点序号："
        '
        'No_Text
        '
        Me.No_Text.Location = New System.Drawing.Point(82, 13)
        Me.No_Text.Name = "No_Text"
        Me.No_Text.Size = New System.Drawing.Size(87, 21)
        Me.No_Text.TabIndex = 1
        '
        'Ok_Btn
        '
        Me.Ok_Btn.Location = New System.Drawing.Point(183, 12)
        Me.Ok_Btn.Name = "Ok_Btn"
        Me.Ok_Btn.Size = New System.Drawing.Size(62, 23)
        Me.Ok_Btn.TabIndex = 2
        Me.Ok_Btn.Text = "确定"
        Me.Ok_Btn.UseVisualStyleBackColor = True
        '
        'DelForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(282, 48)
        Me.Controls.Add(Me.Ok_Btn)
        Me.Controls.Add(Me.No_Text)
        Me.Controls.Add(Me.Label1)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(298, 87)
        Me.MinimumSize = New System.Drawing.Size(298, 87)
        Me.Name = "DelForm"
        Me.Text = "删除站点结算"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents No_Text As Windows.Forms.TextBox
    Friend WithEvents Ok_Btn As Windows.Forms.Button
End Class
