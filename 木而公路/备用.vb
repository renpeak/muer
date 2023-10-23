Imports Microsoft.Data.SqlClient
Imports System.Diagnostics.Eventing
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Module 备用


    ''管理数据库主页面
    'Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

    '    Panel4.Visible = False
    '    Panel5.Visible = False
    '    Panel11.Visible = False
    '    Panel8.Visible = False
    '    Panel7.Visible = False
    '    TabControl3.Visible = False
    '    TabControl2.Visible = False
    '    TabControl1.Visible = False
    '    Panel9.Visible = False
    '    Panel10.Visible = True
    '    Dim con As New SqlConnection("Server=192.168.10.36;DataBase=primary;integrated security=false;uid=sa;pwd=123;Encrypt=False;")
    '    con.Open()
    '    Dim sqlstr = "select * FROM primarytable"
    '    Dim myCommand As New SqlCommand(sqlstr, con)
    '    Reader = New SqlDataAdapter(myCommand)
    '    Dim tempDataSet As New DataTable
    '    Reader.Fill(tempDataSet)
    '    'tempDataSet.Rows.InsertAt(tempDataSet.NewRow, 1)
    '    DataGridView2.DataSource = tempDataSet           '将datagridview的数据源绑定到datatable
    '    con.Close()
    '    Label15.Text = ""
    '    Label16.Text = ""
    'End Sub


    '''管理数据库主页面保存添加按钮
    'Private Sub Button71_Click(sender As Object, e As EventArgs) Handles Button71.Click
    '    Dim data As DataTable
    '    data = DataGridView2.DataSource
    '    data.PrimaryKey = New DataColumn() {data.Columns("id")}
    '    Dim changedata = data.GetChanges
    '    '直接修改数据库数据0)
    '    'MsgBox(dr)

    '    If changedata IsNot Nothing Then
    '        Dim curbase As String
    '        Dim curtable As String
    '        'Try
    '        For i As Integer = 0 To changedata.Rows.Count - 1 Step i + 1
    '            'MsgBox(changedata(i)(0).ToString)
    '            Dim r As Integer = data.Rows.Count - 1
    '            Dim j As Integer
    '            For j = 0 To r
    '                'MsgBox(data.Rows(j).Item(0))

    '                If data.Rows(j).Item(0) = changedata(i)(0).ToString Then
    '                    'MsgBox("行是:" & j)
    '                    curbase = data.Rows(j)("storeid", DataRowVersion.Current).ToString()
    '                    curtable = data.Rows(j)("tableid", DataRowVersion.Current).ToString()
    '                    Dim conn As New SqlConnection("Server=192.168.10.36;DataBase=master;integrated security=false;User Id=sa;Password=123;Encrypt=False;MultipleActiveResultSets=True;")
    '                    conn.Open()
    '                    Dim sqlstr1 = "select * from sysdatabases where name ='" + curbase + "'"
    '                    Dim myCommand1 As New SqlCommand(sqlstr1, conn)
    '                    Dim reader1 As SqlDataReader
    '                    reader1 = myCommand1.ExecuteReader
    '                    '如果数据库存在
    '                    If reader1.HasRows = True Then
    '                        conn.Close()
    '                        Dim conn4 As New SqlConnection("Server=192.168.10.36;DataBase='" + curbase + "';User Id=sa;Password=123;Encrypt=False;")
    '                        conn4.Open()
    '                        'Dim sqlstr3 = "Select count(*) from sysobjects where id = object_id('" + curbase + ".Owner." + curtable + "')"
    '                        Dim sqlstr3 = "Select COUNT(*) From information_schema.TABLES Where table_catalog = '" + curbase + "' And table_name ='" + curtable + "'"
    '                        'MsgBox(sqlstr3)
    '                        Dim myCommand3 As New SqlCommand(sqlstr3, conn4)
    '                        Dim reader2 As SqlDataReader
    '                        reader2 = myCommand3.ExecuteReader
    '                        reader2.Read()
    '                        'MsgBox(reader2(0).ToString)
    '                        '查询是否有表
    '                        If reader2(0).ToString > 0 Then
    '                            conn4.Close()
    '                            MsgBox("已有该项目数据库数据表！！")
    '                        Else
    '                            'conn.Close()
    '                            'Try
    '                            Dim conn3 As New SqlConnection("Server=192.168.10.36;DataBase='" + curbase + "';User Id=sa;Password=123;Encrypt=False;")
    '                            conn3.Open()
    '                            Dim getsql = GetSqlstrs(curtable)
    '                            MsgBox(getsql)
    '                            Dim sqlstr7 = "CREATE TABLE " + curtable + " " + getsql + ""
    '                            Dim myCommand7 As New SqlCommand(sqlstr7, conn3)
    '                            myCommand7.ExecuteNonQuery()
    '                            conn3.Close()
    '                            Label16.Text = "添加成功！！"
    '                            'Catch ex As Exception
    '                            '    MsgBox("已经存在表了哟！！！")
    '                            '    Return
    '                            'End Try
    '                        End If
    '                    Else
    '                        Dim sqlstr5 = "CREATE DATABASE " + curbase + ""
    '                        Dim myCommand5 As New SqlCommand(sqlstr5, conn)
    '                        myCommand5.ExecuteNonQuery()
    '                        conn.Close()
    '                        Dim conn2 As New SqlConnection("Server=192.168.10.36;DataBase='" + curbase + "';User Id=sa;Password=123;Encrypt=False;")
    '                        conn2.Open()
    '                        Dim getsql = GetSqlstrs(curtable)
    '                        Dim sqlstr6 = "CREATE TABLE " + curtable + " " + getsql + " "
    '                        Dim myCommand6 As New SqlCommand(sqlstr6, conn2)
    '                        myCommand6.ExecuteNonQuery()
    '                        conn2.Close()

    '                        Label16.Text = "添加成功！！"
    '                    End If
    '                End If


    '                'MsgBox(3)
    '            Next

    '            'MsgBox(3)
    '        Next
    '        Dim con As New SqlConnection("Server=192.168.10.36;DataBase=primary;integrated security=false;uid=sa;pwd=123;Encrypt=False;")
    '        con.Open()
    '        Dim sqlstr = "select * FROM primarytable"

    '        Dim myCommand As New SqlCommand(sqlstr, con)
    '        Dim reader = New SqlDataAdapter(myCommand)
    '        ''实例化新的sql指令
    '        Dim scb As New SqlCommandBuilder(reader)
    '        scb.GetInsertCommand()
    '        reader.Update(changedata)
    '        data.AcceptChanges()
    '        con.Close()

    '        'Catch ex As Exception
    '        '    MsgBox("id没填或者重复了哦！！")
    '        'End Try
    '    End If



    'End Sub






    ''管理数据库主页面保存修改按钮
    'Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click


    '    Dim data As DataTable
    '    data = DataGridView2.DataSource

    '    'data.PrimaryKey = New DataColumn() {data.Columns("id")}

    '    Dim changedata = data.GetChanges
    '    'MsgBox(changedata(0)(0).ToString)
    '    If changedata IsNot Nothing Then

    '        Dim curbase As String
    '        Dim curtable As String
    '        Dim curbase1 As String
    '        Dim curtable1 As String

    '        For i As Integer = 0 To changedata.Rows.Count - 1 Step i + 1
    '            Dim r As Integer = data.Rows.Count - 1
    '            Dim j As Integer
    '            For j = 0 To r
    '                'MsgBox(data.Rows(j).Item(0))
    '                Try
    '                    If data.Rows(j).Item(0) = changedata(i)(0).ToString Then
    '                        'MsgBox("行是:" & j)
    '                        curbase = data.Rows(j)("storeid", DataRowVersion.Original).ToString()
    '                        curtable = data.Rows(j)("tableid", DataRowVersion.Original).ToString()
    '                        curbase1 = data.Rows(j)("storeid", DataRowVersion.Current).ToString()
    '                        curtable1 = data.Rows(j)("tableid", DataRowVersion.Current).ToString()
    '                        'MsgBox(curbase)
    '                        'MsgBox(curbase1)
    '                        'MsgBox(curtable)
    '                        'MsgBox(curtable1)
    '                        Dim conn As New SqlConnection("Server=192.168.10.36;DataBase=master;integrated security=false;User Id=sa;Password=123;Encrypt=False;MultipleActiveResultSets=True;")
    '                        conn.Open()
    '                        Dim sqlstr1 = "select * from sysdatabases where name ='" + curbase + "'"
    '                        Dim myCommand1 As New SqlCommand(sqlstr1, conn)
    '                        Dim reader1 As SqlDataReader

    '                        reader1 = myCommand1.ExecuteReader
    '                        'MsgBox(reader1.HasRows)
    '                        '如果数据库存在
    '                        If reader1.HasRows = True Then


    '                            If curbase <> curbase1 Then



    '                                '更改数据库名
    '                                Dim sqlstr2 = "ALTER DATABASE " + curbase + " SET SINGLE_USER WITH ROLLBACK IMMEDIATE ALTER DATABASE " + curbase + " MODIFY NAME =" + curbase1 + ""
    '                                Dim myCommand2 As New SqlCommand(sqlstr2, conn)
    '                                myCommand2.ExecuteNonQuery()

    '                                'MsgBox(curbase1)
    '                                Dim sqlstr5 = " ALTER DATABASE " + curbase1 + " SET MULTI_USER"
    '                                Dim myCommand5 As New SqlCommand(sqlstr5, conn)
    '                                myCommand5.ExecuteNonQuery()





    '                                '更改逻辑名
    '                                Dim sqlstr7 = "ALTER DATABASE " + curbase1 + " MODIFY FILE(NAME='" + curbase + "',NEWNAME='" + curbase1 + "')"
    '                                Dim myCommand7 As New SqlCommand(sqlstr7, conn)
    '                                myCommand7.ExecuteNonQuery()


    '                                Dim sqlstr6 = "ALTER DATABASE " + curbase1 + " MODIFY FILE(NAME='" + curbase + "_log',NEWNAME='" + curbase1 + "_log')"
    '                                Dim myCommand6 As New SqlCommand(sqlstr6, conn)
    '                                myCommand6.ExecuteNonQuery()




    '                                '分离数据库
    '                                Dim sqlstr13 = "use master exec sp_detach_db " + curbase1 + ""
    '                                Dim myCommand13 As New SqlCommand(sqlstr13, conn)
    '                                myCommand13.ExecuteNonQuery()




    '                                '更改物理文件名
    '                                Dim sqlstr8 = "use master exec sp_configure 'show advanced options',1 reconfigure with override exec sp_configure 'xp_cmdshell',1 reconfigure with override "
    '                                Dim myCommand8 As New SqlCommand(sqlstr8, conn)
    '                                myCommand8.ExecuteNonQuery()



    '                                Dim sqlstr9 = "exec master.dbo.xp_cmdshell 'D: & cd D:\Program Files\Microsoft SQL Server\MSSQL16.MSSQLSERVER\MSSQL\DATA & rename " + curbase + ".mdf " + curbase1 + ".mdf & rename " + curbase + "_log.ldf " + curbase1 + "_log.ldf'"
    '                                Dim myCommand9 As New SqlCommand(sqlstr9, conn)
    '                                myCommand9.ExecuteNonQuery()




    '                                Dim sqlstr11 = "use master exec sp_configure 'xp_cmdshell',0 reconfigure with override exec sp_configure 'show advanced options',0 reconfigure with override"
    '                                Dim myCommand11 As New SqlCommand(sqlstr11, conn)
    '                                myCommand11.ExecuteNonQuery()




    '                                '附和数据库
    '                                Dim sqlstr12 = "use master exec sp_attach_db " + curbase1 + ", N'D:\Program Files\Microsoft SQL Server\MSSQL16.MSSQLSERVER\MSSQL\DATA\" + curbase1 + ".mdf',N'D:\Program Files\Microsoft SQL Server\MSSQL16.MSSQLSERVER\MSSQL\DATA\" + curbase1 + "_log.ldf'"
    '                                Dim myCommand12 As New SqlCommand(sqlstr12, conn)
    '                                myCommand12.ExecuteNonQuery()




    '                                'Dim sqlstr12 = "alter database" + curbase1 + " add file (name=" + curbase1 + ",filename ='C:\Program Files\Microsoft SQL Server\MSSQL16.MSSQLSERVER\MSSQL\DATA\" + curbase1 + ".mdf',size=1GB,maxsize=10GB,filegrowth=1GB) to filegroup PRIMARY"
    '                                'Dim myCommand12 As New SqlCommand(sqlstr12, conn)
    '                                'myCommand12.ExecuteNonQuery()

    '                                'Dim sqlstr13 = "alter database" + curbase1 + " add log file (name=" + curbase1 + "_log,filename ='C:\Program Files\Microsoft SQL Server\MSSQL16.MSSQLSERVER\MSSQL\DATA\" + curbase1 + "_log.ldf',size=1GB,maxsize=10GB,filegrowth=1GB) to filegroup PRIMARY"
    '                                'Dim myCommand13 As New SqlCommand(sqlstr13, conn)
    '                                'myCommand13.ExecuteNonQuery()



    '                                conn.Close()
    '                                'Dim conn4 As New SqlConnection("Server=192.168.10.36;DataBase=primary;integrated security=false;uid=sa;pwd=123;Encrypt=False;")
    '                                'conn4.Open()
    '                                'Dim sqlstr8 = " Update primarytable Set storeid = '" + curbase1 + "'WHERE storeid = '" + curbase + "'"
    '                                'Dim myCommand8 As New SqlCommand(sqlstr8, conn4)
    '                                'myCommand8.ExecuteNonQuery()
    '                                'conn4.Close()
    '                                'conn.Open()

    '                            End If




    '                            conn.Close()
    '                            Dim conn8 As New SqlConnection("Server=192.168.10.36;DataBase='" + curbase1 + "';User Id=sa;Password=123;Encrypt=False;")
    '                            conn8.Open()
    '                            'Dim sqlstr3 = "Select count(*) from sysobjects where id = object_id('" + curbase + ".Owner." + curtable + "')"
    '                            Dim sqlstr3 = "Select COUNT(*) From information_schema.TABLES Where table_catalog = '" + curbase1 + "' And table_name ='" + curtable + "'"
    '                            'MsgBox(sqlstr3)
    '                            Dim myCommand3 As New SqlCommand(sqlstr3, conn8)
    '                            Dim reader2 As SqlDataReader
    '                            reader2 = myCommand3.ExecuteReader
    '                            reader2.Read()
    '                            MsgBox(reader2(0).ToString)
    '                            '查询是否有表
    '                            If reader2(0).ToString > 0 Then
    '                                conn8.Close()
    '                                If curtable <> curtable1 Then
    '                                    Dim conn1 As New SqlConnection("Server=192.168.10.36;DataBase='" + curbase1 + "';User Id=sa;Password=123;Encrypt=False;")
    '                                    conn1.Open()
    '                                    Dim sqlstr4 = "EXEC sp_rename '" + curtable + "', '" + curtable1 + "'"
    '                                    Dim myCommand4 As New SqlCommand(sqlstr4, conn1)
    '                                    myCommand4.ExecuteNonQuery()
    '                                    conn1.Close()
    '                                Else

    '                                End If


    '                                '直接修改数据库数据

    '                                Dim conn10 As New SqlConnection("Server=192.168.10.36;DataBase=primary;integrated security=false;uid=sa;pwd=123;Encrypt=False;")
    '                                conn10.Open()
    '                                Dim sqlstr10 = "select * FROM primarytable"

    '                                Dim myCommand10 As New SqlCommand(sqlstr10, conn10)
    '                                'Dim reader = New SqlDataAdapter(myCommand)
    '                                ''实例化新的sql指令
    '                                'Dim scb As New SqlCommandBuilder(reader)
    '                                'reader.Update(changedata)
    '                                Reader = New SqlDataAdapter(myCommand10)
    '                                Dim tempDataSet1 As New DataTable
    '                                Dim scb1 As New SqlCommandBuilder(Reader)
    '                                Reader.Update(changedata)
    '                                data.AcceptChanges()
    '                                Reader.Fill(tempDataSet1)
    '                                'tempDataSet.Rows.InsertAt(tempDataSet.NewRow, 1)
    '                                DataGridView2.DataSource = tempDataSet1           '将datagridview的数据源绑定到datatable

    '                                conn10.Close()
    '                                Label16.Text = "修改成功！！"

    '                                Return


    '                            Else
    '                                Dim conn3 As New SqlConnection("Server=192.168.10.36;DataBase='" + curbase1 + "';User Id=sa;Password=123;Encrypt=False;")
    '                                conn3.Open()
    '                                Dim sqlstr7 = "CREATE TABLE " + curtable1 + "( prod_id CHAR(10) NOT NULL,vend_id  CHAR(10) NOT NULL,prod_name  CHAR(254) NOT NULL,prod_price DECIMAL(8,2)  NOT NULL,)"
    '                                Dim myCommand7 As New SqlCommand(sqlstr7, conn3)
    '                                myCommand7.ExecuteNonQuery()
    '                                conn3.Close()
    '                            End If
    '                        Else
    '                            Dim sqlstr5 = "CREATE DATABASE " + curbase1 + ""
    '                            Dim myCommand5 As New SqlCommand(sqlstr5, conn)
    '                            myCommand5.ExecuteNonQuery()
    '                            conn.Close()
    '                            Dim conn2 As New SqlConnection("Server=192.168.10.36;DataBase='" + curbase1 + "';User Id=sa;Password=123;Encrypt=False;")
    '                            conn2.Open()
    '                            Dim sqlstr6 = "CREATE TABLE " + curtable1 + "( prod_id CHAR(10) NOT NULL,vend_id  CHAR(10) NOT NULL,prod_name  CHAR(254) NOT NULL,prod_price DECIMAL(8,2)  NOT NULL,)"
    '                            Dim myCommand6 As New SqlCommand(sqlstr6, conn2)
    '                            myCommand6.ExecuteNonQuery()
    '                            conn2.Close()

    '                        End If

    '                        '
    '                        '直接修改数据库数据

    '                        Dim con As New SqlConnection("Server=192.168.10.36;DataBase=primary;integrated security=false;uid=sa;pwd=123;Encrypt=False;")
    '                        con.Open()
    '                        Dim sqlstr = "select * FROM primarytable"

    '                        Dim myCommand As New SqlCommand(sqlstr, con)
    '                        'Dim reader = New SqlDataAdapter(myCommand)
    '                        ''实例化新的sql指令
    '                        'Dim scb As New SqlCommandBuilder(reader)
    '                        'reader.Update(changedata)
    '                        Reader = New SqlDataAdapter(myCommand)
    '                        Dim tempDataSet As New DataTable
    '                        'Dim scb As New SqlCommandBuilder(reader)
    '                        'reader.Update(changedata)
    '                        'data.AcceptChanges()
    '                        Reader.Fill(tempDataSet)
    '                        'tempDataSet.Rows.InsertAt(tempDataSet.NewRow, 1)
    '                        DataGridView2.DataSource = tempDataSet           '将datagridview的数据源绑定到datatable

    '                        con.Close()



    '                    End If
    '                Catch ex As Exception
    '                    MsgBox("如要添加数据请点击保存添加按钮哦！！")
    '                End Try
    '                'MsgBox(3)
    '            Next

    '        Next
    '        Label16.Text = "修改成功！！"
    '    Else
    '        MsgBox("您未对数据边进行修改！！")

    '    End If
    'End Sub



    ''管理数据库主页面删除按钮
    'Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click

    '    Dim curbase As String
    '    Dim response
    '    'Dim id As String
    '    Dim j As Integer
    '    j = DataGridView2.Rows.Count
    '    'Dim curtable As String
    '    response = MsgBox("是否确定删除所选行！！！", vbOKCancel, "提示！！！")
    '    If response = vbOK Then
    '        For i As Integer = 0 To j
    '            If i >= j Then Exit For
    '            If DataGridView2.Rows(i).Selected = True Then
    '                Try
    '                    curbase = DataGridView2.Rows(i).Cells(1).Value.ToString  '获取第二列

    '                    Try
    '                        Dim conn As New SqlConnection("Server=192.168.10.36;DataBase=master;integrated security=false;User Id=sa;Password=123;Encrypt=False;MultipleActiveResultSets=True;")
    '                        conn.Open()


    '                        '    Dim sqlstr12 = "USE " + curbase + " DBCC SHRINKFILE (N'" + curbase + "' , EMPTYFILE)"
    '                        '    Dim myCommand12 As New SqlCommand(sqlstr12, conn)
    '                        '    myCommand12.ExecuteNonQuery()



    '                        '    Dim sqlstr13 = "USE master ALTER DATABASE " + curbase + " REMOVE FILE " + curbase + ""
    '                        '    Dim myCommand13 As New SqlCommand(sqlstr13, conn)
    '                        '    myCommand13.ExecuteNonQuery()



    '                        '    Dim sqlstr10 = "USE " + curbase + " DBCC SHRINKFILE (N'" + curbase + "_log' , EMPTYFILE)"
    '                        '    Dim myCommand10 As New SqlCommand(sqlstr10, conn)
    '                        '    myCommand10.ExecuteNonQuery()





    '                        '    Dim sqlstr11 = "USE master ALTER DATABASE " + curbase + " REMOVE FILE " + curbase + "_log"
    '                        '    Dim myCommand11 As New SqlCommand(sqlstr11, conn)
    '                        '    myCommand11.ExecuteNonQuery()




    '                        Dim sqlstr6 = " USE master ALTER DATABASE " + curbase + " SET SINGLE_USER WITH ROLLBACK IMMEDIATE DROP DATABASE " + curbase + ""
    '                        Dim myCommand6 As New SqlCommand(sqlstr6, conn)
    '                        myCommand6.ExecuteNonQuery()


    '                        Dim sqlstr8 = " ALTER DATABASE db_database SET MULTI_USER"
    '                        Dim myCommand9 As New SqlCommand(sqlstr8, conn)
    '                        myCommand9.ExecuteNonQuery()
    '                        conn.Close()

    '                    Catch
    '                        MsgBox("数据库正在使用，请重试！！！")
    '                        Return
    '                    End Try
    '                    DataGridView2.Rows.RemoveAt(i)

    '                    '删除存储表
    '                    Dim con As New SqlConnection("Server=192.168.10.36;DataBase=primary;integrated security=false;uid=sa;pwd=123;Encrypt=False;MultipleActiveResultSets=True;")
    '                    con.Open()
    '                    Dim sqlstr1 = "select id from primarytable where storeid='" + curbase + "' "
    '                    'MsgBox(curbase)
    '                    Dim myCommand As New SqlCommand(sqlstr1, con)
    '                    Dim reader As SqlDataReader
    '                    reader = myCommand.ExecuteReader

    '                    While (reader.Read())
    '                        'MsgBox(reader(1).ToString)
    '                        Dim sqlstr = "DELETE From primarytable Where id = '" + reader(0).ToString + "'"
    '                        Dim myCommand1 As New SqlCommand(sqlstr, con)
    '                        myCommand1.ExecuteNonQuery()

    '                    End While
    '                    con.Close()
    '                Catch
    '                    MsgBox("删除失败哦！！！")
    '                End Try
    '                i -= 1
    '                j -= 1
    '            End If

    '        Next
    '    Else
    '        Return
    '    End If



    '    Dim conn9 As New SqlConnection("Server=192.168.10.36;DataBase=primary;integrated security=false;uid=sa;pwd=123;Encrypt=False;")
    '    conn9.Open()
    '    Dim sqlstr7 = "select * FROM primarytable"

    '    Dim myCommand8 As New SqlCommand(sqlstr7, conn9)
    '    'Dim reader = New SqlDataAdapter(myCommand)
    '    ''实例化新的sql指令
    '    'Dim scb As New SqlCommandBuilder(reader)
    '    'reader.Update(changedata)
    '    Reader = New SqlDataAdapter(myCommand8)
    '    Dim tempDataSet As New DataTable
    '    Dim scb As New SqlCommandBuilder(Reader)
    '    Reader.Fill(tempDataSet)
    '    'tempDataSet.Rows.InsertAt(tempDataSet.NewRow, 1)
    '    DataGridView2.DataSource = tempDataSet
    '    MsgBox("删除成功！！！")
    '    conn9.Close()



    '    'Dim scb As New SqlCommandBuilder(reader) '实例化新的sql指令
    '    'scb.GetUpdateCommand()                           '获取Update功能
    '    'Dim data As DataTable
    '    'data = DataGridView2.DataSource
    '    'Dim changedata = data.GetChanges
    '    'If changedata IsNot Nothing Then
    '    '    reader.Update(changedata)
    '    '    data.AcceptChanges()
    '    '    MsgBox("删除成功！！！")
    '    'End If

    'End Sub



    ''管理数据库主页面返回按钮
    'Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
    '    Dim data As DataTable
    '    data = DataGridView2.DataSource
    '    Dim changedata = data.GetChanges
    '    Dim response
    '    If changedata IsNot Nothing Then

    '        response = MsgBox("修改未保存，确认退出吗？", vbOKCancel, "提示！！！")
    '        If response = vbOK Then

    '            Panel4.Visible = False
    '            Panel5.Visible = False
    '            Panel11.Visible = False
    '            Panel10.Visible = False
    '            Panel9.Visible = False
    '            Panel8.Visible = False
    '            TabControl3.Visible = False
    '            TabControl2.Visible = False
    '            TabControl1.Visible = False
    '            Panel7.Visible = True
    '        Else

    '            Panel4.Visible = False
    '            Panel5.Visible = False
    '            Panel11.Visible = False
    '            Panel7.Visible = False
    '            Panel8.Visible = False
    '            TabControl3.Visible = False
    '            TabControl2.Visible = False
    '            TabControl1.Visible = False
    '            Panel9.Visible = False
    '            Panel10.Visible = True
    '        End If
    '    Else

    '        Panel4.Visible = False
    '        Panel5.Visible = False
    '        Panel11.Visible = False
    '        Panel10.Visible = False
    '        Panel9.Visible = False
    '        Panel8.Visible = False
    '        TabControl3.Visible = False
    '        TabControl2.Visible = False
    '        TabControl1.Visible = False
    '        Panel7.Visible = True
    '    End If

    'End Sub




    ''系统管理员桥梁工程选择数据库后自动选择表
    'Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
    '    'Label20.Text = ""
    '    Dim con As New SqlConnection("Server=192.168.10.36;DataBase=primary;integrated security=false;uid=sa;pwd=123;Encrypt=False;")
    '    con.Open()
    '    Dim sqlstr = "SELECT  tablename,id FROM primarytable where storename='" + ComboBox1.Text + "'"
    '    Dim myCommand As New SqlCommand(sqlstr, con)
    '    Dim reader As SqlDataAdapter = New SqlDataAdapter(myCommand)
    '    Dim tempDataSet As New DataTable
    '    reader.Fill(tempDataSet)

    '    ComboBox2.DataSource = tempDataSet
    '    ComboBox2.ValueMember = "id"
    '    ComboBox2.DisplayMember = "tablename"
    '    con.Close()

    'End Sub



    ''系统管理员桥梁工程导入数据库按钮
    'Private Sub Button41_Click(sender As Object, e As EventArgs) Handles Button41.Click


    '    Dim con As New SqlConnection("Server=192.168.10.36;DataBase=primary;integrated security=false;uid=sa;pwd=123;Encrypt=False;")
    '    'Dim con As New SqlConnection("Server=192.168.10.11;DataBase=primary;integrated security=false;uid=sa;pwd=sa;Encrypt=False")
    '    con.Open()
    '    Dim sqlstr = "SELECT  storeid,tableid FROM primarytable where storename='" + ComboBox1.Text + "'And tablename='" + ComboBox2.Text + "'"
    '    'MsgBox(sqlstr)
    '    Dim myCommand As New SqlCommand(sqlstr, con)
    '    Dim reader As SqlDataReader
    '    reader = myCommand.ExecuteReader
    '    reader.Read()
    '    Try
    '        store = {reader(0).ToString, reader(1).ToString}
    '        OpenFileDialog1.Title = "请选择需要导入的数据表"
    '        OpenFileDialog1.Filter = "(*.xls)|*.xls|(*.xlsx)|*.xlsx|All files (*.*)|*.*"
    '        OpenFileDialog1.FilterIndex = 1
    '        OpenFileDialog1.FileName = ""
    '        OpenFileDialog1.ShowDialog()
    '    Catch ex As Exception
    '        MsgBox("没有数据库或数据表，请联系技术开发部！！")
    '    End Try

    '    con.Close()
    'End Sub




    ''系统管理员路基工程选择数据库后自动选择表
    'Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
    '    Dim con As New SqlConnection("Server=192.168.10.36;DataBase=primary;integrated security=false;uid=sa;pwd=123;Encrypt=False;")
    '    con.Open()
    '    Dim sqlstr = "SELECT  tablename id FROM primarytable where storename='" + ComboBox4.Text + "'"
    '    Dim myCommand As New SqlCommand(sqlstr, con)
    '    Dim reader As SqlDataAdapter = New SqlDataAdapter(myCommand)
    '    Dim tempDataSet As New DataTable
    '    reader.Fill(tempDataSet)

    '    ComboBox3.DataSource = tempDataSet
    '    ComboBox3.ValueMember = "id"
    '    ComboBox3.DisplayMember = "tablename"
    '    con.Close()
    'End Sub

    ''系统管理员路基工程导入数据库按钮
    'Private Sub Button68_Click(sender As Object, e As EventArgs) Handles Button68.Click
    '    Dim con As New SqlConnection("Server=192.168.10.36;DataBase=primary;integrated security=false;uid=sa;pwd=123;Encrypt=False;")
    '    'Dim con As New SqlConnection("Server=192.168.10.11;DataBase=primary;integrated security=false;uid=sa;pwd=sa;Encrypt=False")
    '    con.Open()
    '    Dim sqlstr = "SELECT  storeid,tableid FROM primarytable where storename='" + ComboBox4.Text + "'And tablename='" + ComboBox3.Text + "'"
    '    'MsgBox(sqlstr)
    '    Dim myCommand As New SqlCommand(sqlstr, con)
    '    Dim reader As SqlDataReader
    '    reader = myCommand.ExecuteReader
    '    reader.Read()
    '    Try
    '        store = {reader(0).ToString, reader(1).ToString}
    '        OpenFileDialog1.Title = "请找到需要导入的数据表"
    '        OpenFileDialog1.Filter = "(*.xls)|*.xls|(*.xlsx)|*.xlsx|All files (*.*)|*.*"
    '        OpenFileDialog1.FilterIndex = 2
    '        OpenFileDialog1.ShowDialog()
    '    Catch ex As Exception
    '        MsgBox("没有数据库或数据表，请联系相关部门！！")
    '    End Try

    '    con.Close()
    'End Sub


    ''系统管理员隧道工程选择数据库后自动选择表
    'Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
    '    Dim con As New SqlConnection("Server=192.168.10.36;DataBase=primary;integrated security=false;uid=sa;pwd=123;Encrypt=False;")
    '    con.Open()
    '    Dim sqlstr = "SELECT  tablename id FROM primarytable where storename='" + ComboBox6.Text + "'"
    '    Dim myCommand As New SqlCommand(sqlstr, con)
    '    Dim reader As SqlDataAdapter = New SqlDataAdapter(myCommand)
    '    Dim tempDataSet As New DataTable
    '    reader.Fill(tempDataSet)

    '    ComboBox5.DataSource = tempDataSet
    '    ComboBox5.ValueMember = "id"
    '    ComboBox5.DisplayMember = "tablename"
    '    con.Close()
    'End Sub


    ''系统管理员隧道工程导入数据库按钮
    'Private Sub Button70_Click(sender As Object, e As EventArgs) Handles Button70.Click
    '    Dim con As New SqlConnection("Server=192.168.10.36;DataBase=primary;integrated security=false;uid=sa;pwd=123;Encrypt=False;")
    '    'Dim con As New SqlConnection("Server=192.168.10.11;DataBase=primary;integrated security=false;uid=sa;pwd=sa;Encrypt=False")
    '    con.Open()
    '    Dim sqlstr = "SELECT  storeid,tableid FROM primarytable where storename='" + ComboBox6.Text + "'And tablename='" + ComboBox5.Text + "'"
    '    'MsgBox(sqlstr)
    '    Dim myCommand As New SqlCommand(sqlstr, con)
    '    Dim reader As SqlDataReader
    '    reader = myCommand.ExecuteReader
    '    reader.Read()
    '    Try
    '        store = {reader(0).ToString, reader(1).ToString}
    '        OpenFileDialog1.Title = "请找到需要导入的数据表"
    '        OpenFileDialog1.Filter = "(*.xls)|*.xls|(*.xlsx)|*.xlsx|All files (*.*)|*.*"
    '        OpenFileDialog1.FilterIndex = 2
    '        OpenFileDialog1.ShowDialog()
    '    Catch ex As Exception
    '        MsgBox("没有数据库或数据表，请联系相关部门！！")
    '    End Try

    '    con.Close()
    'End Sub

End Module
