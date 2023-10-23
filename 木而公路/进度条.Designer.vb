<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class 进度条
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
        components = New ComponentModel.Container()
        ProgressBar1 = New ProgressBar()
        Timer1 = New Timer(components)
        Button1 = New Button()
        SuspendLayout()
        ' 
        ' ProgressBar1
        ' 
        ProgressBar1.ForeColor = Color.LimeGreen
        ProgressBar1.Location = New Point(-1, 6)
        ProgressBar1.Name = "ProgressBar1"
        ProgressBar1.Size = New Size(273, 29)
        ProgressBar1.Step = 1
        ProgressBar1.Style = ProgressBarStyle.Continuous
        ProgressBar1.TabIndex = 0
        ' 
        ' Timer1
        ' 
        ' 
        ' Button1
        ' 
        Button1.Location = New Point(99, 41)
        Button1.Name = "Button1"
        Button1.Size = New Size(75, 27)
        Button1.TabIndex = 1
        Button1.Text = "中止"
        Button1.UseVisualStyleBackColor = True
        ' 
        ' 进度条
        ' 
        AutoScaleDimensions = New SizeF(7F, 17F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(272, 72)
        Controls.Add(Button1)
        Controls.Add(ProgressBar1)
        FormBorderStyle = FormBorderStyle.FixedDialog
        MaximizeBox = False
        MdiChildrenMinimizedAnchorBottom = False
        MinimizeBox = False
        Name = "进度条"
        StartPosition = FormStartPosition.CenterScreen
        Text = "正在计算并导出，请稍后......"
        ResumeLayout(False)
    End Sub

    Friend WithEvents ProgressBar1 As ProgressBar
    Friend WithEvents Timer1 As Timer
    Friend WithEvents Button1 As Button
End Class
