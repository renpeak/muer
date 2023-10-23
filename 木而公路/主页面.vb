Imports System.Data.OleDb
Imports System.Security
Imports Microsoft.Data.SqlClient
Imports System.Threading


Public Class 主页面
    Private IsFormBeingDragged As Boolean = False
    Private MouseDownX As Integer
    Private MouseDownY As Integer
    Public Thd As Thread
    Private Sub 主页面_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If action(2) = "系统管理员" Then
            Dim con As New SqlConnection("Server=192.168.10.36;DataBase=master;integrated security=false;uid=sa;pwd=123;Encrypt=False;")
            'Dim con As New SqlConnection("Server=renfeng.tpddns.cn,1000;DataBase=master;integrated security=false;uid=sa;pwd=sa;Encrypt=false;Trusted_Connection=false")
            con.Open()
            Dim sqlstr = "SELECT name,dbid FROM sysdatabases WHERE charindex('项目生产组数据库',name)>0"
            Dim myCommand As New SqlCommand(sqlstr, con)
            Dim reader As SqlDataAdapter = New SqlDataAdapter(myCommand)
            Dim tempDataSet As New DataTable
            reader.Fill(tempDataSet)
            ComboBox7.DataSource = tempDataSet
            ComboBox7.ValueMember = "dbid"
            ComboBox7.DisplayMember = "name"
            con.Close()
            Button7.Visible = True
        End If
        Me.Location = New System.Drawing.Point(650, 250)

        '当前账户
        Label7.Text = action(0)
        '部门
        Label9.Text = action(2)
        'logo
        System.Environment.CurrentDirectory = My.Application.Info.DirectoryPath
        FileName2 = System.IO.Path.GetFullPath("../../../")
        SetPictureBoxImage(PictureBox1, "" + FileName2 + "img\logomain.png")
        '移动
        Me.FormBorderStyle = FormBorderStyle.None
        '提示信息
        ToolTip1.SetToolTip(Button2, "关闭")
        ToolTip1.SetToolTip(Button6, "最小化")
        ToolTip1.SetToolTip(Button10, "注销")
        Panel6.BringToFront()
        Panel3.BringToFront()
    End Sub

    'picturebox图标背景设置
    Friend Sub SetPictureBoxImage(ByVal pb As PictureBox, ByVal sFileName As String)
        '定义一个Bitmap对象作为绘制的接受对象
        Dim bmp As New Bitmap(pb.Width, pb.Height)
        Dim g As Graphics = Graphics.FromImage(bmp)

        Dim img As Image = Image.FromFile(sFileName)
        Dim rectImage As New Rectangle(0, 0, bmp.Width, bmp.Height)
        '按比例缩放
        GetScaleZoomRect(img.Width, img.Height, rectImage.Width, rectImage.Height)
        g.DrawImage(img, rectImage)
        pb.Image = bmp
    End Sub


    '设置图标尺寸
    Friend Function GetScaleZoomRect(ByVal nSrcWidth As Integer, ByVal nSrcHeight As Integer, ByRef nDstWidth As Integer, ByRef nDstHeight As Integer)
        If nSrcWidth / nSrcHeight < nDstWidth / nDstHeight Then
            nDstWidth = nDstHeight * (nSrcWidth / nSrcHeight)
        Else
            nDstHeight = nDstWidth * (nSrcHeight / nSrcWidth)
        End If
    End Function

    '关闭页面
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Application.Exit()

    End Sub

    '最小化
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    '注销
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click

        Dim YesorNo = New DialogResult()
        YesorNo = MsgBox("是否注销当前账户？", 4 + 48, "警告")
        If YesorNo = vbYes Then
            Me.Close()
            'Threading.Thread.Sleep(3000)
            登录.TextBox1.Text = "账户"
            登录.TextBox2.PasswordChar = ""
            登录.TextBox2.Text = "密码"
            登录.PictureBox4.Visible = False
            登录.Label3.Text = Nothing
            登录.Label6.Text = Nothing
            登录.Show()
        End If

    End Sub


    '鼠标拖动窗体移动
    Private Sub 主页面_MouseDown(sender As Object, e As MouseEventArgs) Handles MyBase.MouseDown

        If e.Button = MouseButtons.Left Then
            IsFormBeingDragged = True
            MouseDownX = e.X
            MouseDownY = e.Y
        End If
    End Sub
    '鼠标拖动窗体移动
    Private Sub 主页面_MouseUp(sender As Object, e As MouseEventArgs) Handles MyBase.MouseUp
        If e.Button = MouseButtons.Left Then
            IsFormBeingDragged = False
        End If
    End Sub
    '鼠标拖动窗体移动
    Private Sub 主页面_MouseMove(sender As Object, e As MouseEventArgs) Handles MyBase.MouseMove
        If IsFormBeingDragged Then
            Dim temp As Point = New Point()

            temp.X = Me.Location.X + (e.X - MouseDownX)
            temp.Y = Me.Location.Y + (e.Y - MouseDownY)
            Me.Location = temp
            temp = Nothing
        End If
    End Sub

    '系统管理员管理用户
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Panel9.BringToFront()
        Panel3.BringToFront()
        Dim con As New SqlConnection("Server=192.168.10.36;DataBase=木而公路用户数据库;integrated security=false;uid=sa;pwd=123;Encrypt=False;")
        'Dim con As New SqlConnection("Server=renfeng.tpddns.cn,1000;DataBase=木而公路用户数据库;integrated security=false;uid=sa;pwd=sa;Encrypt=false;Trusted_Connection=false")
        con.Open()
        Dim sqlstr = "select * FROM login where userID<>'administrator'"
        Dim myCommand As New SqlCommand(sqlstr, con)
        reader = New SqlDataAdapter(myCommand)
        Dim tempDataSet As New DataTable
        reader.Fill(tempDataSet)
        'tempDataSet.Rows.InsertAt(tempDataSet.NewRow, 1)
        DataGridView1.DataSource = tempDataSet           '将datagridview的数据源绑定到datatable
        con.Close()

    End Sub

    '系统管理员及用户修改密码

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Panel8.BringToFront()
        Panel3.BringToFront()
        TextBox1.Text = action(0)
        TextBox3.Text = action(2)
    End Sub


    '修改密码按钮点击事件
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        If TextBox2.Text = "" Or TextBox4.Text = "" Then
            'MsgBox("用户名或密码不能为空！！", , "提示")
            Label14.Text = "密码不能为空！！"
        ElseIf TextBox2.Text <> TextBox4.Text Then
            Label14.Text = "两次输入密码不同！！请重新输入"
        Else
            Dim con As New SqlConnection("Server=192.168.10.36;DataBase=木而公路用户数据库;integrated security=false;uid=sa;pwd=123;Encrypt=False;MultipleActiveResultSets=True;")
            'Dim con As New SqlConnection("Server=renfeng.tpddns.cn,1000;DataBase=木而公路用户数据库;integrated security=false;uid=sa;pwd=sa;Encrypt=false;Trusted_Connection=false")
            con.Open()
            Dim sqlstr = "update login set password = '" + TextBox2.Text + "' WHERE userID = '" + TextBox1.Text + "'"
            Dim myCommand As New SqlCommand(sqlstr, con)
            Dim reader As Integer
            reader = myCommand.ExecuteNonQuery
            Label14.Text = "修改成功！！"
            con.Close()

            Dim vbOKOnly = New DialogResult()
            vbOKOnly = MsgBox("修改成功,请重新登录", 0 + 48, "警告")
            If vbOKOnly = vbOK Then
                Me.Close()
                'Threading.Thread.Sleep(3000)
                登录.TextBox1.Text = "账户"
                登录.TextBox2.PasswordChar = ""
                登录.TextBox2.Text = "密码"
                登录.PictureBox4.Visible = False
                登录.Label3.Text = Nothing
                登录.Label6.Text = Nothing
                登录.Show()
            End If
        End If
    End Sub
    '禁止换行
    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If e.KeyChar = Chr(13) Then
            e.Handled = True
            Button11.PerformClick()
        End If
    End Sub
    '禁止换行
    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        If e.KeyChar = Chr(13) Then
            e.Handled = True
            Button11.PerformClick()
        End If
    End Sub


    '管理用户主页面

    Dim reader As SqlDataAdapter
    Dim tempDataSet As DataTable


    '系统管理员修改用户按钮事件
    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        Dim scb As New SqlCommandBuilder(reader) '实例化新的sql指令
        scb.GetUpdateCommand()                           '获取Update功能
        Dim data As DataTable
        data = DataGridView1.DataSource
        Dim changedata = data.GetChanges
        If changedata IsNot Nothing Then
            reader.Update(changedata)
            data.AcceptChanges()
            MsgBox("数据修改成功", 0 + 64, "提示")
        End If
    End Sub

    '系统管理员删除用户按钮事件
    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        For Each r As DataGridViewRow In DataGridView1.SelectedRows
            Dim YesorNo = New DialogResult()
            YesorNo = MsgBox("是否删除选中的数据？", 4 + 48, "警告")
            If YesorNo = vbYes Then
                If Not r.IsNewRow Then
                    DataGridView1.Rows.Remove(r)
                End If
                Dim scb As New SqlCommandBuilder(reader) '实例化新的sql指令
                scb.GetUpdateCommand()                           '获取Update功能
                Dim data As DataTable
                data = DataGridView1.DataSource
                Dim changedata = data.GetChanges
                If changedata IsNot Nothing Then
                    reader.Update(changedata)
                    data.AcceptChanges()
                    MsgBox("已删除选中的数据", 0 + 64, "提示")
                End If
            End If
        Next
    End Sub

    '页面关闭修改用户状态
    Private Sub 主页面_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        Dim con As New SqlConnection("Server=192.168.10.36;DataBase=木而公路用户数据库;integrated security=false;uid=sa;pwd=123;Encrypt=False;")
        'Dim con As New SqlConnection("Server=renfeng.tpddns.cn,1000;DataBase=木而公路用户数据库;integrated security=false;uid=sa;pwd=sa;Encrypt=false;Trusted_Connection=false;")
        con.Open()
        Dim sqlstr1 = "UPDATE login SET status = 0 WHERE userID collate Chinese_PRC_CS_AS_WS= '" + action(0) + "'"
        Dim myCommand1 As New SqlCommand(sqlstr1, con)
        myCommand1.ExecuteNonQuery()
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If TorF = False Then
            ExApp.DisplayAlerts = False
            ExApp.ScreenUpdating = True
            Exbook.Close()
            ExApp.DisplayAlerts = True
            ExApp.Quit()
            GC.Collect()
            Panel3.Enabled = True
            Timer1.Enabled = False
            ProgressBar1.Visible = False
            Label1.Visible = False
            Button8.Visible = False
        ElseIf ProgressBar1.Value < ProgressBar1.Maximum Then
            ProgressBar1.Value += 20
        Else
            ProgressBar1.Value = 0
        End If
    End Sub


    '桥梁工程
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    '桥梁工程主页面
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim storename As String
        storename = Button1.Text
        projectname = {LTrim(storename.Remove(storename.LastIndexOf("工程"))), LTrim(storename.Remove(storename.LastIndexOf("工程"))) & "-"}
        '判断是否为系统管理员
        If Label9.Text = "系统管理员" Then
            ComboBox7.Visible = True
            TabControl1.BringToFront()
            Panel3.BringToFront()
            Label17.Visible = True
        Else
            TabControl1.BringToFront()
            Panel3.BringToFront()
        End If
    End Sub

    '桥梁导入数据库
    '下部构造
    '桩基础
    Private Sub Button91_Click(sender As Object, e As EventArgs) Handles Button91.Click
        Dim tablename As String
        storebase = Nothing
        tablename = Button91.Text
        storebase = projectname(1) & tablename
        Examine(tablename)
        If filenamed = "" Or TorF = False Then
            Exit Sub
        End If
        Call 桩基资料()

    End Sub

    '承台
    Private Sub Button84_Click(sender As Object, e As EventArgs) Handles Button84.Click
        Dim tablename As String
        storebase = Nothing
        tablename = Button84.Text
        storebase = projectname(1) & tablename
        Examine(tablename)
        If filenamed = "" Or TorF = False Then
            Exit Sub
        Else
            Thd = New Thread(AddressOf 承台资料)
            Thd.Start()
        End If
    End Sub

    '墩柱
    Private Sub Button85_Click(sender As Object, e As EventArgs) Handles Button85.Click
        Dim tablename As String
        storebase = Nothing
        tablename = Button85.Text
        storebase = projectname(1) & tablename
        Examine(tablename)
        If filenamed = "" Or TorF = False Then
            Exit Sub
        End If
        Call 墩柱资料()
    End Sub

    '系梁
    Private Sub Button86_Click(sender As Object, e As EventArgs) Handles Button86.Click
        Dim tablename As String
        storebase = Nothing
        tablename = Button86.Text
        storebase = projectname(1) & tablename
        Examine(tablename)
        If filenamed = "" Or TorF = False Then
            Exit Sub
        End If
        Dim Thd As New Thread(AddressOf 系梁资料)
        Thd.Start()
    End Sub


    '台身
    Private Sub Button82_Click(sender As Object, e As EventArgs) Handles Button82.Click
        Dim tablename As String
        storebase = Nothing
        tablename = Button82.Text
        storebase = projectname(1) & tablename
        Examine(tablename)
        If filenamed = "" Or TorF = False Then
            Exit Sub
        End If
        Call 台身资料()
    End Sub


    '台帽
    Private Sub Button87_Click(sender As Object, e As EventArgs) Handles Button87.Click
        Dim tablename As String
        storebase = Nothing
        tablename = Button87.Text
        storebase = projectname(1) & tablename
        Examine(tablename)
        If filenamed = "" Or TorF = False Then
            Exit Sub
        End If
        Call 台帽资料()
    End Sub

    '耳墙
    Private Sub Button88_Click(sender As Object, e As EventArgs) Handles Button88.Click
        Dim tablename As String
        storebase = Nothing
        tablename = Button88.Text
        storebase = projectname(1) & tablename
        Examine(tablename)
        If filenamed = "" Or TorF = False Then
            Exit Sub
        End If
        Call 耳墙资料()
    End Sub

    '肋板
    Private Sub Button90_Click(sender As Object, e As EventArgs) Handles Button90.Click
        Dim tablename As String
        storebase = Nothing
        tablename = Button90.Text
        storebase = projectname(1) & tablename
        Examine(tablename)
        If filenamed = "" Or TorF = False Then
            Exit Sub
        End If
        Call 肋板资料()
    End Sub

    '支座垫石
    Private Sub Button89_Click(sender As Object, e As EventArgs) Handles Button89.Click
        Dim tablename As String
        storebase = Nothing
        tablename = Button89.Text
        storebase = projectname(1) & tablename
        Examine(tablename)
        If filenamed = "" Or TorF = False Then
            Exit Sub
        End If
        Call 支座垫石资料()
    End Sub

    '挡块
    Private Sub Button83_Click(sender As Object, e As EventArgs) Handles Button83.Click
        Dim tablename As String
        storebase = Nothing
        tablename = Button83.Text
        storebase = projectname(1) & tablename
        Examine(tablename)
        If filenamed = "" Or TorF = False Then
            Exit Sub
        End If
        Call 挡块资料()
    End Sub

    '台背回填
    Private Sub Button81_Click(sender As Object, e As EventArgs) Handles Button81.Click
        Dim tablename As String
        storebase = Nothing
        tablename = Button81.Text
        storebase = projectname(1) & tablename
        Examine(tablename)
        If filenamed = "" Or TorF = False Then
            Exit Sub
        End If

    End Sub

    '桥台扩大基础
    Private Sub Button73_Click(sender As Object, e As EventArgs) Handles Button73.Click
        Dim tablename As String
        storebase = Nothing
        tablename = Button73.Text
        storebase = projectname(1) & tablename
        Examine(tablename)
        If filenamed = "" Or TorF = False Then
            Exit Sub
        End If
        Call 桥台扩大基础资料()
    End Sub

    '背墙
    Private Sub Button107_Click(sender As Object, e As EventArgs) Handles Button107.Click
        Dim tablename As String
        storebase = Nothing
        tablename = Button107.Text
        storebase = projectname(1) & tablename
        Examine(tablename)
        If filenamed = "" Or TorF = False Then
            Exit Sub
        End If
        Call 背墙资料()
    End Sub

    '盖梁
    Private Sub Button116_Click(sender As Object, e As EventArgs) Handles Button116.Click
        Dim tablename As String
        storebase = Nothing
        tablename = Button116.Text
        storebase = projectname(1) & tablename
        Examine(tablename)
        If filenamed = "" Or TorF = False Then
            Exit Sub
        End If
        Call 盖梁资料()
    End Sub




    '上部构造
    '梁板预制
    Private Sub Button105_Click(sender As Object, e As EventArgs) Handles Button105.Click
        Dim tablename As String
        storebase = Nothing
        tablename = Button105.Text
        storebase = projectname(1) & tablename
        Examine(tablename)
        If filenamed = "" Or TorF = False Then
            Exit Sub
        End If
    End Sub

    '张拉
    Private Sub Button98_Click(sender As Object, e As EventArgs) Handles Button98.Click
        Dim tablename As String
        storebase = Nothing
        tablename = Button98.Text
        storebase = projectname(1) & tablename
        Examine(tablename)
        If filenamed = "" Or TorF = False Then
            Exit Sub
        End If
    End Sub

    '压浆
    Private Sub Button99_Click(sender As Object, e As EventArgs) Handles Button99.Click
        Dim tablename As String
        storebase = Nothing
        tablename = Button99.Text
        storebase = projectname(1) & tablename
        Examine(tablename)
        If filenamed = "" Or TorF = False Then
            Exit Sub
        End If
    End Sub

    '梁板安装
    Private Sub Button100_Click(sender As Object, e As EventArgs) Handles Button100.Click
        Dim tablename As String
        storebase = Nothing
        tablename = Button100.Text
        storebase = projectname(1) & tablename
        Examine(tablename)
        If filenamed = "" Or TorF = False Then
            Exit Sub
        End If
    End Sub

    '端横梁
    Private Sub Button96_Click(sender As Object, e As EventArgs) Handles Button96.Click
        Dim tablename As String
        storebase = Nothing
        tablename = Button96.Text
        storebase = projectname(1) & tablename
        Examine(tablename)
        If filenamed = "" Or TorF = False Then
            Exit Sub
        End If
    End Sub

    '湿接缝
    Private Sub Button101_Click(sender As Object, e As EventArgs) Handles Button101.Click
        Dim tablename As String
        storebase = Nothing
        tablename = Button101.Text
        storebase = projectname(1) & tablename
        Examine(tablename)
        If filenamed = "" Or TorF = False Then
            Exit Sub
        End If
    End Sub



    '桥面及附属
    '支座安装
    Private Sub Button115_Click(sender As Object, e As EventArgs) Handles Button115.Click
        Dim tablename As String
        storebase = Nothing
        tablename = Button115.Text
        storebase = projectname(1) & tablename
        Examine(tablename)
        If filenamed = "" Or TorF = False Then
            Exit Sub
        End If
    End Sub

    '伸缩缝
    Private Sub Button108_Click(sender As Object, e As EventArgs) Handles Button108.Click
        Dim tablename As String
        storebase = Nothing
        tablename = Button108.Text
        storebase = projectname(1) & tablename
        Examine(tablename)
        If filenamed = "" Or TorF = False Then
            Exit Sub
        End If
    End Sub

    '防水层
    Private Sub Button109_Click(sender As Object, e As EventArgs) Handles Button109.Click
        Dim tablename As String
        storebase = Nothing
        tablename = Button109.Text
        storebase = projectname(1) & tablename
        Examine(tablename)
        If filenamed = "" Or TorF = False Then
            Exit Sub
        End If
    End Sub

    '沥青铺装层
    Private Sub Button110_Click(sender As Object, e As EventArgs) Handles Button110.Click
        Dim tablename As String
        storebase = Nothing
        tablename = Button110.Text
        storebase = projectname(1) & tablename
        Examine(tablename)
        If filenamed = "" Or TorF = False Then
            Exit Sub
        End If
    End Sub

    '桥面铺装
    Private Sub Button106_Click(sender As Object, e As EventArgs) Handles Button106.Click
        Dim tablename As String
        storebase = Nothing
        tablename = Button106.Text
        storebase = projectname(1) & tablename
        Examine(tablename)
        If filenamed = "" Or TorF = False Then
            Exit Sub
        End If
    End Sub

    '桥梁总体
    Private Sub Button111_Click(sender As Object, e As EventArgs) Handles Button111.Click
        Dim tablename As String
        storebase = Nothing
        tablename = Button111.Text
        storebase = projectname(1) & tablename
        Examine(tablename)
        If filenamed = "" Or TorF = False Then
            Exit Sub
        End If
    End Sub

    '桥头搭板
    Private Sub Button112_Click(sender As Object, e As EventArgs) Handles Button112.Click
        Dim tablename As String
        storebase = Nothing
        tablename = Button112.Text
        storebase = projectname(1) & tablename
        Examine(tablename)
        If filenamed = "" Or TorF = False Then
            Exit Sub
        End If
    End Sub

    '防撞护栏
    Private Sub Button114_Click(sender As Object, e As EventArgs) Handles Button114.Click
        Dim tablename As String
        storebase = Nothing
        tablename = Button114.Text
        storebase = projectname(1) & tablename
        Examine(tablename)
        If filenamed = "" Or TorF = False Then
            Exit Sub
        End If
    End Sub

    '锥坡防护
    Private Sub Button113_Click(sender As Object, e As EventArgs) Handles Button113.Click
        Dim tablename As String
        storebase = Nothing
        tablename = Button113.Text
        storebase = projectname(1) & tablename
        Examine(tablename)
        If filenamed = "" Or TorF = False Then
            Exit Sub
        End If
    End Sub


    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    '路基工程主页面
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Dim storename As String
        storename = Button4.Text
        projectname = {LTrim(storename.Remove(storename.LastIndexOf("工程"))), LTrim(storename.Remove(storename.LastIndexOf("工程"))) & "-"}
        '判断是否为系统管理员
        If Label9.Text = "系统管理员" Then
            TabControl2.BringToFront()
            Panel3.BringToFront()
            ComboBox7.Visible = True
            Label17.Visible = True
        Else
            TabControl2.BringToFront()
            Panel3.BringToFront()
        End If
    End Sub

    '回填砂砾石
    Private Sub Button80_Click(sender As Object, e As EventArgs)

        Dim tablename As String
        store = Nothing
        tablename = Button80.Text
        Examine(tablename)
    End Sub

    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '隧道工程主页面
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        Dim storename As String
        storename = Button4.Text
        projectname = {LTrim(storename.Remove(storename.LastIndexOf("工程"))), LTrim(storename.Remove(storename.LastIndexOf("工程"))) & "-"}
        '判断是否为系统管理员
        If Label9.Text = "系统管理员" Then
            ComboBox7.Visible = True
            TabControl3.BringToFront()
            Panel3.BringToFront()
            Label17.Visible = True
        Else
            TabControl3.BringToFront()
            Panel3.BringToFront()
        End If
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Panel6.BringToFront()
        Panel3.BringToFront()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        TorF = False
    End Sub

    '路面工程主页面
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim storename As String
        storename = Button4.Text
        projectname = {LTrim(storename.Remove(storename.LastIndexOf("工程"))), LTrim(storename.Remove(storename.LastIndexOf("工程"))) & "-"}
        '判断是否为系统管理员
        If Label9.Text = "系统管理员" Then
            ComboBox7.Visible = True
            Panel2.BringToFront()
            Panel3.BringToFront()
            Label17.Visible = True
        Else
            Panel2.BringToFront()
            Panel3.BringToFront()
        End If
    End Sub

    '交安工程主页面
    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Dim storename As String
        storename = Button4.Text
        projectname = {LTrim(storename.Remove(storename.LastIndexOf("工程"))), LTrim(storename.Remove(storename.LastIndexOf("工程"))) & "-"}
        '判断是否为系统管理员
        If Label9.Text = "系统管理员" Then
            ComboBox7.Visible = True
            Panel4.BringToFront()
            Panel3.BringToFront()
            Label17.Visible = True
        Else
            Panel4.BringToFront()
            Panel3.BringToFront()
        End If
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        Dim storename As String
        storename = Button4.Text
        projectname = {LTrim(storename.Remove(storename.LastIndexOf("工程"))), LTrim(storename.Remove(storename.LastIndexOf("工程"))) & "-"}
        '判断是否为系统管理员
        If Label9.Text = "系统管理员" Then
            ComboBox7.Visible = True
            Panel5.BringToFront()
            Panel3.BringToFront()
            Label17.Visible = True
        Else
            Panel5.BringToFront()
            Panel3.BringToFront()
        End If
    End Sub



End Class