<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class 登录
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
        Dim resources As ComponentModel.ComponentResourceManager = New ComponentModel.ComponentResourceManager(GetType(登录))
        Button1 = New Button()
        Label1 = New Label()
        TextBox1 = New TextBox()
        TextBox2 = New TextBox()
        Label6 = New Label()
        SqlCommand1 = New Microsoft.Data.SqlClient.SqlCommand()
        PictureBox3 = New PictureBox()
        Label2 = New Label()
        Button2 = New Button()
        PictureBox1 = New PictureBox()
        PictureBox2 = New PictureBox()
        PictureBox5 = New PictureBox()
        Panel1 = New Panel()
        PictureBox6 = New PictureBox()
        PictureBox4 = New PictureBox()
        Label3 = New Label()
        ToolTip1 = New ToolTip(components)
        CType(PictureBox3, ComponentModel.ISupportInitialize).BeginInit()
        CType(PictureBox1, ComponentModel.ISupportInitialize).BeginInit()
        CType(PictureBox2, ComponentModel.ISupportInitialize).BeginInit()
        CType(PictureBox5, ComponentModel.ISupportInitialize).BeginInit()
        Panel1.SuspendLayout()
        CType(PictureBox6, ComponentModel.ISupportInitialize).BeginInit()
        CType(PictureBox4, ComponentModel.ISupportInitialize).BeginInit()
        SuspendLayout()
        ' 
        ' Button1
        ' 
        Button1.BackColor = Color.FromArgb(CByte(45), CByte(194), CByte(155))
        Button1.BackgroundImageLayout = ImageLayout.Zoom
        Button1.FlatAppearance.BorderSize = 0
        Button1.FlatStyle = FlatStyle.Flat
        Button1.Font = New Font("微软雅黑", 12F, FontStyle.Regular, GraphicsUnit.Point)
        Button1.ForeColor = Color.Transparent
        Button1.Location = New Point(24, 143)
        Button1.Name = "Button1"
        Button1.Size = New Size(172, 35)
        Button1.TabIndex = 3
        Button1.Text = "登   录"
        Button1.UseVisualStyleBackColor = False
        ' 
        ' Label1
        ' 
        Label1.BackColor = Color.Transparent
        Label1.Font = New Font("Microsoft Sans Serif", 10.5F, FontStyle.Regular, GraphicsUnit.Point)
        Label1.ForeColor = Color.White
        Label1.Location = New Point(55, 20)
        Label1.Name = "Label1"
        Label1.Size = New Size(117, 21)
        Label1.TabIndex = 2
        Label1.Text = "木而公路软件"
        ' 
        ' TextBox1
        ' 
        TextBox1.BackColor = SystemColors.Control
        TextBox1.BorderStyle = BorderStyle.None
        TextBox1.Font = New Font("宋体", 13F, FontStyle.Regular, GraphicsUnit.Point)
        TextBox1.ForeColor = SystemColors.GrayText
        TextBox1.Location = New Point(46, 24)
        TextBox1.Multiline = True
        TextBox1.Name = "TextBox1"
        TextBox1.Size = New Size(150, 26)
        TextBox1.TabIndex = 0
        ' 
        ' TextBox2
        ' 
        TextBox2.BackColor = SystemColors.Control
        TextBox2.BorderStyle = BorderStyle.None
        TextBox2.Font = New Font("宋体", 13F, FontStyle.Regular, GraphicsUnit.Point)
        TextBox2.Location = New Point(46, 86)
        TextBox2.Multiline = True
        TextBox2.Name = "TextBox2"
        TextBox2.PasswordChar = "*"c
        TextBox2.Size = New Size(150, 26)
        TextBox2.TabIndex = 1
        ' 
        ' Label6
        ' 
        Label6.BackColor = Color.Transparent
        Label6.Font = New Font("宋体", 9F, FontStyle.Regular, GraphicsUnit.Point)
        Label6.ForeColor = Color.OrangeRed
        Label6.Location = New Point(33, 119)
        Label6.Name = "Label6"
        Label6.Size = New Size(157, 17)
        Label6.TabIndex = 11
        ' 
        ' SqlCommand1
        ' 
        SqlCommand1.CommandTimeout = 30
        SqlCommand1.EnableOptimizedParameterBinding = False
        ' 
        ' PictureBox3
        ' 
        PictureBox3.BackColor = Color.Transparent
        PictureBox3.BackgroundImage = CType(resources.GetObject("PictureBox3.BackgroundImage"), Image)
        PictureBox3.Location = New Point(202, 88)
        PictureBox3.Name = "PictureBox3"
        PictureBox3.Size = New Size(23, 23)
        PictureBox3.SizeMode = PictureBoxSizeMode.StretchImage
        PictureBox3.TabIndex = 13
        PictureBox3.TabStop = False
        PictureBox3.Visible = False
        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.BackColor = Color.Transparent
        Label2.Font = New Font("Microsoft YaHei UI", 7F, FontStyle.Regular, GraphicsUnit.Point)
        Label2.Location = New Point(56, 294)
        Label2.Name = "Label2"
        Label2.Size = New Size(321, 16)
        Label2.TabIndex = 15
        Label2.Text = "成都木而工程管理咨询有限公司@版权所有  Power By Chengdu muer"
        ' 
        ' Button2
        ' 
        Button2.AutoSize = True
        Button2.BackColor = Color.Transparent
        Button2.BackgroundImageLayout = ImageLayout.Zoom
        Button2.FlatAppearance.BorderSize = 0
        Button2.FlatStyle = FlatStyle.Flat
        Button2.Image = CType(resources.GetObject("Button2.Image"), Image)
        Button2.Location = New Point(405, -1)
        Button2.Name = "Button2"
        Button2.Size = New Size(27, 24)
        Button2.TabIndex = 4
        Button2.TextImageRelation = TextImageRelation.ImageBeforeText
        Button2.UseVisualStyleBackColor = False
        ' 
        ' PictureBox1
        ' 
        PictureBox1.BackColor = SystemColors.Control
        PictureBox1.BackgroundImage = CType(resources.GetObject("PictureBox1.BackgroundImage"), Image)
        PictureBox1.Location = New Point(23, 24)
        PictureBox1.Name = "PictureBox1"
        PictureBox1.Size = New Size(23, 26)
        PictureBox1.TabIndex = 16
        PictureBox1.TabStop = False
        ' 
        ' PictureBox2
        ' 
        PictureBox2.BackColor = SystemColors.Control
        PictureBox2.BackgroundImage = CType(resources.GetObject("PictureBox2.BackgroundImage"), Image)
        PictureBox2.Location = New Point(23, 86)
        PictureBox2.Name = "PictureBox2"
        PictureBox2.Size = New Size(23, 26)
        PictureBox2.TabIndex = 17
        PictureBox2.TabStop = False
        ' 
        ' PictureBox5
        ' 
        PictureBox5.BackColor = Color.Transparent
        PictureBox5.Location = New Point(9, 10)
        PictureBox5.Name = "PictureBox5"
        PictureBox5.Size = New Size(58, 43)
        PictureBox5.TabIndex = 19
        PictureBox5.TabStop = False
        ' 
        ' Panel1
        ' 
        Panel1.BackColor = Color.Transparent
        Panel1.Controls.Add(PictureBox6)
        Panel1.Controls.Add(PictureBox4)
        Panel1.Controls.Add(Label3)
        Panel1.Controls.Add(TextBox1)
        Panel1.Controls.Add(Label6)
        Panel1.Controls.Add(PictureBox2)
        Panel1.Controls.Add(TextBox2)
        Panel1.Controls.Add(PictureBox1)
        Panel1.Controls.Add(Button1)
        Panel1.Controls.Add(PictureBox3)
        Panel1.Location = New Point(99, 68)
        Panel1.Name = "Panel1"
        Panel1.Size = New Size(234, 195)
        Panel1.TabIndex = 20
        ' 
        ' PictureBox6
        ' 
        PictureBox6.BackColor = Color.Transparent
        PictureBox6.BackgroundImage = CType(resources.GetObject("PictureBox6.BackgroundImage"), Image)
        PictureBox6.Location = New Point(202, 88)
        PictureBox6.Name = "PictureBox6"
        PictureBox6.Size = New Size(23, 23)
        PictureBox6.SizeMode = PictureBoxSizeMode.AutoSize
        PictureBox6.TabIndex = 21
        PictureBox6.TabStop = False
        PictureBox6.Visible = False
        ' 
        ' PictureBox4
        ' 
        PictureBox4.BackColor = Color.Transparent
        PictureBox4.BackgroundImage = CType(resources.GetObject("PictureBox4.BackgroundImage"), Image)
        PictureBox4.Location = New Point(23, 56)
        PictureBox4.Name = "PictureBox4"
        PictureBox4.Size = New Size(23, 23)
        PictureBox4.TabIndex = 20
        PictureBox4.TabStop = False
        PictureBox4.Visible = False
        ' 
        ' Label3
        ' 
        Label3.Font = New Font("宋体", 10F, FontStyle.Regular, GraphicsUnit.Point)
        Label3.Location = New Point(46, 56)
        Label3.Name = "Label3"
        Label3.Size = New Size(182, 23)
        Label3.TabIndex = 19
        Label3.TextAlign = ContentAlignment.MiddleLeft
        ' 
        ' 登录
        ' 
        AutoScaleDimensions = New SizeF(7F, 17F)
        AutoScaleMode = AutoScaleMode.Font
        BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), Image)
        BackgroundImageLayout = ImageLayout.Stretch
        ClientSize = New Size(433, 315)
        Controls.Add(Label1)
        Controls.Add(PictureBox5)
        Controls.Add(Panel1)
        Controls.Add(Label2)
        Controls.Add(Button2)
        DoubleBuffered = True
        ForeColor = Color.Transparent
        FormBorderStyle = FormBorderStyle.None
        Icon = CType(resources.GetObject("$this.Icon"), Icon)
        KeyPreview = True
        MaximizeBox = False
        MinimizeBox = False
        Name = "登录"
        StartPosition = FormStartPosition.CenterScreen
        Text = "登录"
        CType(PictureBox3, ComponentModel.ISupportInitialize).EndInit()
        CType(PictureBox1, ComponentModel.ISupportInitialize).EndInit()
        CType(PictureBox2, ComponentModel.ISupportInitialize).EndInit()
        CType(PictureBox5, ComponentModel.ISupportInitialize).EndInit()
        Panel1.ResumeLayout(False)
        Panel1.PerformLayout()
        CType(PictureBox6, ComponentModel.ISupportInitialize).EndInit()
        CType(PictureBox4, ComponentModel.ISupportInitialize).EndInit()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents Button1 As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents TextBox2 As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents SqlCommand1 As Microsoft.Data.SqlClient.SqlCommand
    Friend WithEvents PictureBox3 As PictureBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Button2 As Button
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents PictureBox2 As PictureBox
    Friend WithEvents PictureBox5 As PictureBox
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Label3 As Label
    Friend WithEvents PictureBox4 As PictureBox
    Friend WithEvents PictureBox6 As PictureBox
    Friend WithEvents ToolTip1 As ToolTip
End Class
