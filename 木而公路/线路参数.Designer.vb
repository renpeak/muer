<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class 线路参数
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
        Dim resources As ComponentModel.ComponentResourceManager = New ComponentModel.ComponentResourceManager(GetType(线路参数))
        ComboBox1 = New ComboBox()
        Button1 = New Button()
        Button2 = New Button()
        Label1 = New Label()
        Label2 = New Label()
        Label3 = New Label()
        Label4 = New Label()
        ComboBox2 = New ComboBox()
        ComboBox3 = New ComboBox()
        ComboBox4 = New ComboBox()
        ComboBox5 = New ComboBox()
        Label6 = New Label()
        Label5 = New Label()
        SuspendLayout()
        ' 
        ' ComboBox1
        ' 
        ComboBox1.FormattingEnabled = True
        ComboBox1.Location = New Point(73, 9)
        ComboBox1.Name = "ComboBox1"
        ComboBox1.Size = New Size(166, 25)
        ComboBox1.TabIndex = 0
        ' 
        ' Button1
        ' 
        Button1.Location = New Point(42, 203)
        Button1.Name = "Button1"
        Button1.Size = New Size(75, 28)
        Button1.TabIndex = 1
        Button1.Text = "确定"
        Button1.UseVisualStyleBackColor = True
        ' 
        ' Button2
        ' 
        Button2.Location = New Point(164, 203)
        Button2.Name = "Button2"
        Button2.Size = New Size(75, 28)
        Button2.TabIndex = 2
        Button2.Text = "取消"
        Button2.UseVisualStyleBackColor = True
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Location = New Point(12, 12)
        Label1.Name = "Label1"
        Label1.Size = New Size(56, 17)
        Label1.TabIndex = 3
        Label1.Text = "交点法："
        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.Location = New Point(12, 48)
        Label2.Name = "Label2"
        Label2.Size = New Size(56, 17)
        Label2.TabIndex = 4
        Label2.Text = "线元法："
        ' 
        ' Label3
        ' 
        Label3.AutoSize = True
        Label3.Location = New Point(12, 87)
        Label3.Name = "Label3"
        Label3.Size = New Size(56, 17)
        Label3.TabIndex = 5
        Label3.Text = "断   链："
        ' 
        ' Label4
        ' 
        Label4.AutoSize = True
        Label4.Location = New Point(12, 120)
        Label4.Name = "Label4"
        Label4.Size = New Size(68, 17)
        Label4.TabIndex = 6
        Label4.Text = "导线成果："
        ' 
        ' ComboBox2
        ' 
        ComboBox2.FormattingEnabled = True
        ComboBox2.Location = New Point(73, 45)
        ComboBox2.Name = "ComboBox2"
        ComboBox2.Size = New Size(166, 25)
        ComboBox2.TabIndex = 7
        ' 
        ' ComboBox3
        ' 
        ComboBox3.FormattingEnabled = True
        ComboBox3.Location = New Point(73, 81)
        ComboBox3.Name = "ComboBox3"
        ComboBox3.Size = New Size(166, 25)
        ComboBox3.TabIndex = 8
        ' 
        ' ComboBox4
        ' 
        ComboBox4.FormattingEnabled = True
        ComboBox4.Location = New Point(73, 117)
        ComboBox4.Name = "ComboBox4"
        ComboBox4.Size = New Size(166, 25)
        ComboBox4.TabIndex = 9
        ' 
        ' ComboBox5
        ' 
        ComboBox5.FormattingEnabled = True
        ComboBox5.Location = New Point(73, 152)
        ComboBox5.Name = "ComboBox5"
        ComboBox5.Size = New Size(166, 25)
        ComboBox5.TabIndex = 12
        ' 
        ' Label6
        ' 
        Label6.AutoSize = True
        Label6.Location = New Point(12, 155)
        Label6.Name = "Label6"
        Label6.Size = New Size(56, 17)
        Label6.TabIndex = 11
        Label6.Text = "竖曲线："
        ' 
        ' Label5
        ' 
        Label5.AutoSize = True
        Label5.Font = New Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point)
        Label5.ForeColor = Color.Red
        Label5.Location = New Point(85, 182)
        Label5.Name = "Label5"
        Label5.Size = New Size(122, 17)
        Label5.TabIndex = 13
        Label5.Text = "正在导入，请稍后......"
        ' 
        ' 线路参数
        ' 
        AutoScaleDimensions = New SizeF(7F, 17F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(262, 239)
        Controls.Add(Label5)
        Controls.Add(ComboBox5)
        Controls.Add(Label6)
        Controls.Add(ComboBox4)
        Controls.Add(ComboBox3)
        Controls.Add(ComboBox2)
        Controls.Add(Label4)
        Controls.Add(Label3)
        Controls.Add(Label2)
        Controls.Add(Label1)
        Controls.Add(Button2)
        Controls.Add(Button1)
        Controls.Add(ComboBox1)
        FormBorderStyle = FormBorderStyle.FixedSingle
        Icon = CType(resources.GetObject("$this.Icon"), Icon)
        MaximizeBox = False
        MdiChildrenMinimizedAnchorBottom = False
        MinimizeBox = False
        Name = "线路参数"
        StartPosition = FormStartPosition.CenterScreen
        Text = "线路参数"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents ComboBox1 As ComboBox
    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents ComboBox2 As ComboBox
    Friend WithEvents ComboBox3 As ComboBox
    Friend WithEvents ComboBox4 As ComboBox
    Friend WithEvents ComboBox5 As ComboBox
    Friend WithEvents Label6 As Label
    Friend WithEvents Label5 As Label
End Class
