Partial Class addtools
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Windows.Forms 类撰写设计器支持所必需的
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        '组件设计器需要此调用。
        InitializeComponent()

    End Sub

    '组件重写释放以清理组件列表。
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    '组件设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是组件设计器所必需的
    '可使用组件设计器修改它。
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.Separator1 = Me.Factory.CreateRibbonSeparator
        Me.Separator2 = Me.Factory.CreateRibbonSeparator
        Me.Init_Btn = Me.Factory.CreateRibbonButton
        Me.Create_Btn = Me.Factory.CreateRibbonButton
        Me.Add_Btn = Me.Factory.CreateRibbonButton
        Me.Delete_Btn = Me.Factory.CreateRibbonButton
        Me.Fix_Btn = Me.Factory.CreateRibbonButton
        Me.DelFill_Btn = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Label = "2020结算插件"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.Init_Btn)
        Me.Group1.Items.Add(Me.Create_Btn)
        Me.Group1.Items.Add(Me.Separator1)
        Me.Group1.Items.Add(Me.Add_Btn)
        Me.Group1.Items.Add(Me.Fix_Btn)
        Me.Group1.Items.Add(Me.Delete_Btn)
        Me.Group1.Items.Add(Me.Separator2)
        Me.Group1.Items.Add(Me.DelFill_Btn)
        Me.Group1.Label = "结算工具"
        Me.Group1.Name = "Group1"
        '
        'Separator1
        '
        Me.Separator1.Name = "Separator1"
        '
        'Separator2
        '
        Me.Separator2.Name = "Separator2"
        '
        'Init_Btn
        '
        Me.Init_Btn.Image = Global._2020Settlement.My.Resources.Resources.应用程序管理器
        Me.Init_Btn.Label = "重置结算模板"
        Me.Init_Btn.Name = "Init_Btn"
        Me.Init_Btn.ShowImage = True
        '
        'Create_Btn
        '
        Me.Create_Btn.Image = Global._2020Settlement.My.Resources.Resources.任务
        Me.Create_Btn.Label = "生成汇总表"
        Me.Create_Btn.Name = "Create_Btn"
        Me.Create_Btn.ShowImage = True
        '
        'Add_Btn
        '
        Me.Add_Btn.Image = Global._2020Settlement.My.Resources.Resources.编辑
        Me.Add_Btn.Label = "添加站点结算"
        Me.Add_Btn.Name = "Add_Btn"
        Me.Add_Btn.ShowImage = True
        '
        'Delete_Btn
        '
        Me.Delete_Btn.Image = Global._2020Settlement.My.Resources.Resources.删除
        Me.Delete_Btn.Label = "删除站点结算"
        Me.Delete_Btn.Name = "Delete_Btn"
        Me.Delete_Btn.ShowImage = True
        '
        'Fix_Btn
        '
        Me.Fix_Btn.Image = Global._2020Settlement.My.Resources.Resources.修改
        Me.Fix_Btn.Label = "修改站点结算"
        Me.Fix_Btn.Name = "Fix_Btn"
        Me.Fix_Btn.ShowImage = True
        '
        'DelFill_Btn
        '
        Me.DelFill_Btn.Image = Global._2020Settlement.My.Resources.Resources.一键清空
        Me.DelFill_Btn.Label = "清空工作量"
        Me.DelFill_Btn.Name = "DelFill_Btn"
        Me.DelFill_Btn.ShowImage = True
        '
        'addtools
        '
        Me.Name = "addtools"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Add_Btn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Delete_Btn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Init_Btn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Create_Btn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator1 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents DelFill_Btn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Fix_Btn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator2 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property addtools() As addtools
        Get
            Return Me.GetRibbon(Of addtools)()
        End Get
    End Property
End Class
