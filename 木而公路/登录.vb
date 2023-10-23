
Imports Microsoft.Data.SqlClient
'Imports Microsoft.Office.Interop.Excel
Imports System.Drawing.Image

Public Class 登录

    Dim filename2 As String
    '开始
    Private Sub 登录_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.SuspendLayout()
        SetStyle(ControlStyles.UserPaint, True)
        SetStyle(ControlStyles.AllPaintingInWmPaint, True) ' 禁止擦除背景.
        SetStyle(ControlStyles.DoubleBuffer, True) ' 双缓冲
        UpdateStyles()
        TextBox1.AcceptsReturn = False
        ToolTip1.SetToolTip(Button2, "退出")
        TextBox1.Text = "账户"
        TextBox2.PasswordChar = ""
        TextBox2.Text = "密码"
        Me.ResumeLayout()
        System.Environment.CurrentDirectory = My.Application.Info.DirectoryPath
        filename2 = System.IO.Path.GetFullPath("../../../")
        SetPictureBoxImage(PictureBox5, "" + filename2 + "img\logo.png")
    End Sub

    '登录页面移动
    Private IsFormBeingDragged As Boolean = False
    Private MouseDownX As Integer
    Private MouseDownY As Integer
    Private Sub 主页面_MouseDown(sender As Object, e As MouseEventArgs) Handles MyBase.MouseDown

        If e.Button = MouseButtons.Left Then
            IsFormBeingDragged = True
            MouseDownX = e.X
            MouseDownY = e.Y
        End If
    End Sub

    Private Sub 主页面_MouseUp(sender As Object, e As MouseEventArgs) Handles MyBase.MouseUp
        If e.Button = MouseButtons.Left Then
            IsFormBeingDragged = False
        End If
    End Sub

    Private Sub 主页面_MouseMove(sender As Object, e As MouseEventArgs) Handles MyBase.MouseMove
        If IsFormBeingDragged Then
            Dim temp As Point = New Point()

            temp.X = Me.Location.X + (e.X - MouseDownX)
            temp.Y = Me.Location.Y + (e.Y - MouseDownY)
            Me.Location = temp
            temp = Nothing
        End If
    End Sub


    '账号框的enter
    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = Chr(13) Then
            e.Handled = True
            'SendKeys.Send("{TAB}")
            Button1.PerformClick()
        End If
    End Sub


    '密码框的enter
    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If e.KeyChar = Chr(13) Then
            e.Handled = True
            Button1.PerformClick()
        End If
    End Sub


    'logo的写作
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

    'logo计算
    Friend Function GetScaleZoomRect(ByVal nSrcWidth As Integer, ByVal nSrcHeight As Integer, ByRef nDstWidth As Integer, ByRef nDstHeight As Integer)
        If nSrcWidth / nSrcHeight < nDstWidth / nDstHeight Then
            nDstWidth = nDstHeight * (nSrcWidth / nSrcHeight)
        Else
            nDstHeight = nDstWidth * (nSrcHeight / nSrcWidth)
        End If
    End Function


    '登录点击按钮
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        '判断
        If TextBox1.Text = "账户" Or TextBox2.Text = "密码" Then
            Label6.Text = "用户名或密码不能为空!"
        Else
            'Dim con As SqlConnection
            'Dim con As New SqlConnection("Server=192.168.10.36;DataBase=木而公路用户数据库;integrated security=false;uid=sa;pwd=123;Encrypt=False;")
            Dim con As New SqlConnection("Server=renfeng.tpddns.cn,1000;DataBase=木而公路用户数据库;integrated security=false;uid=sa;pwd=sa;Encrypt=false;Trusted_Connection=false;MultipleActiveResultSets=True;")
            con.Open()
            Dim sqlstr = "select * from login where userID collate Chinese_PRC_CS_AS_WS='" + TextBox1.Text + "'and password collate Chinese_PRC_CS_AS_WS='" + TextBox2.Text + "'"
            Dim myCommand As New SqlCommand(sqlstr, con)
            Dim reader As SqlDataReader
            Dim store As String
            reader = myCommand.ExecuteReader
            If reader.Read() = True Then
                If reader(3) = False Then
                    '修改状态
                    Dim sqlstr1 = "UPDATE login SET status = 1 WHERE userID collate Chinese_PRC_CS_AS_WS= '" + TextBox1.Text + "'"
                    Dim myCommand1 As New SqlCommand(sqlstr1, con)
                    myCommand1.ExecuteNonQuery()
                    store = Label3.Text & "数据库"
                    action = {TextBox1.Text, TextBox2.Text, Label3.Text, store}
                    主页面.Show()
                    Me.Hide()
                Else
                    MsgBox("该账户已登录，请将该账户注销后再登录！")
                End If
            Else
                Label6.Text = "账号或密码错误,请重新填写!"
            End If
            con.Close()
        End If

    End Sub

    '退出按钮
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        'Me.Close()
        Application.Exit()
    End Sub

    '账号框改变提示字的颜色
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Label6.Text = ""
        'TextBox2.Text = ""
        If TextBox1.Text <> "" And TextBox1.Text <> "账户" Then
            'PictureBox6.Visible = True
            TextBox1.ForeColor = Color.Black
        Else
            'PictureBox6.Visible = False
            TextBox1.ForeColor = Color.Gray
        End If

    End Sub

    '密码框改变提示字的颜色
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        Label6.Text = ""

        If TextBox2.Text <> "" And TextBox2.Text <> "密码" Then

            If PictureBox3.Visible = True Then
                TextBox2.PasswordChar = ""
            Else
                PictureBox6.Visible = True
                TextBox2.ForeColor = Color.Black
                TextBox2.PasswordChar = "*"
            End If
        Else
            PictureBox3.Visible = False
            PictureBox6.Visible = False
            TextBox2.ForeColor = Color.Gray
            TextBox2.PasswordChar = ""
        End If

    End Sub

    '账号框得到焦点
    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        If TextBox1.Text = "账户" Then
            TextBox1.Text = ""
        End If
    End Sub

    '密码框得到焦点
    Private Sub TextBox2_GotFocus(sender As Object, e As EventArgs) Handles TextBox2.GotFocus

        If TextBox2.Text = "密码" Then
            TextBox2.Text = ""
            TextBox2.PasswordChar = "*"
        End If
        If TextBox1.Text <> "账户" Then
            Try
                'Dim con As New SqlConnection("Server=192.168.10.36;DataBase=木而公路用户数据库;integrated security=false;uid=sa;pwd=123;Encrypt=False")
                Dim con As New SqlConnection("Server=renfeng.tpddns.cn,1000;DataBase=木而公路用户数据库;integrated security=false;uid=sa;pwd=sa;Encrypt=false;Trusted_Connection=false")
                con.Open()
                Dim sqlstr = "select role from login where userID collate Chinese_PRC_CS_AS_WS='" + TextBox1.Text + "'"
                Dim myCommand As New SqlCommand(sqlstr, con)
                Dim reader As SqlDataReader
                reader = myCommand.ExecuteReader
                If reader.Read() Then
                    Label3.Text = reader(0).ToString
                    PictureBox4.Visible = True
                ElseIf TextBox1.Text = "" Then
                    Label3.Text = ""
                Else
                    Label3.Text = "该账户未注册!"
                End If
                con.Close()
            Catch
                MsgBox("找不到用户，请联系系统管理员！！！")
            End Try
        End If
    End Sub

    '点击眼睛进行睁眼
    Private Sub PictureBox6_Click(sender As Object, e As EventArgs) Handles PictureBox6.Click
        TextBox2.PasswordChar = ""
        PictureBox3.Visible = True
        PictureBox6.Visible = False

    End Sub

    '点击闭眼
    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        TextBox2.PasswordChar = "*"
        PictureBox6.Visible = True
        PictureBox3.Visible = False
    End Sub

    '账户框焦点消失
    Private Sub TextBox1_Leave(sender As Object, e As EventArgs) Handles TextBox1.Leave
        If TextBox1.Text = "" Then
            TextBox1.Text = "账户"
        End If
    End Sub

    '密码框焦点消失
    Private Sub TextBox2_Leave(sender As Object, e As EventArgs) Handles TextBox2.Leave
        If TextBox2.Text = "" Then
            TextBox2.PasswordChar = ""
            TextBox2.Text = "密码"
            'TextBox2.ForeColor = Color.Gray
        End If
    End Sub


End Class