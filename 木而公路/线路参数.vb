Imports Microsoft.Data.SqlClient
Imports System.Security.Cryptography.X509Certificates
Imports System.Text.Json.Serialization
Imports System.Threading
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class 线路参数
    Dim storename As String
    Dim storeku1 As String

    Private Sub 线路参数_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label5.Visible = False
        TorF = False
        If action(2) = "系统管理员" Then
            storeku1 = storeku.Remove(storeku.LastIndexOf("数据库"))
            storename = storeku1 & "线路参数"
        Else
            storename = action(2) & "线路参数"
        End If
        Dim dr As DataRow
        Dim dr1 As DataRow
        Dim dr2 As DataRow
        Dim dr3 As DataRow
        ' Dim con As New SqlConnection("Server=192.168.10.36;DataBase=master;integrated security=false;uid=sa;pwd=123;Encrypt=False;MultipleActiveResultSets=True;")
        Dim con As New SqlConnection("Server=renfeng.tpddns.cn,1000;DataBase=master;integrated security=false;uid=sa;pwd=sa;Encrypt=false;Trusted_Connection=false")
        con.Open()
        '交点法
        Dim sqlstr = "SELECT name,object_id FROM " + storename + ".sys.tables WHERE charindex('交点法',name)>0 "
        Dim myCommand As New SqlCommand(sqlstr, con)
        Dim reader As SqlDataAdapter = New SqlDataAdapter(myCommand)
        Dim tempDataSet As New DataTable
        reader.Fill(tempDataSet)
        dr = tempDataSet.NewRow
        dr("object_id") = 0
        dr("name") = "--请选择--"
        tempDataSet.Rows.InsertAt(dr, 0)
        ComboBox1.DataSource = tempDataSet
        ComboBox1.ValueMember = "object_id"
        ComboBox1.DisplayMember = "name"

        '线元法
        Dim sqlstr1 = "SELECT name,object_id FROM " + storename + ".sys.tables WHERE charindex('线元法',name)>0 "
        Dim myCommand1 As New SqlCommand(sqlstr1, con)
        Dim reader1 As SqlDataAdapter = New SqlDataAdapter(myCommand1)
        Dim tempDataSet1 As New DataTable
        reader1.Fill(tempDataSet1)
        dr1 = tempDataSet1.NewRow
        dr1("object_id") = 0
        dr1("name") = "--请选择--"
        tempDataSet1.Rows.InsertAt(dr1, 0)
        ComboBox2.DataSource = tempDataSet1
        ComboBox2.ValueMember = "object_id"
        ComboBox2.DisplayMember = "name"

        '断链
        Dim sqlstr2 = "SELECT name,object_id FROM " + storename + ".sys.tables WHERE charindex('断链',name)>0 "
        Dim myCommand2 As New SqlCommand(sqlstr2, con)
        Dim reader2 As SqlDataAdapter = New SqlDataAdapter(myCommand2)
        Dim tempDataSet2 As New DataTable
        reader2.Fill(tempDataSet2)
        dr2 = tempDataSet2.NewRow
        dr2("object_id") = 0
        dr2("name") = "--请选择--"
        tempDataSet2.Rows.InsertAt(dr2, 0)
        ComboBox3.DataSource = tempDataSet2
        ComboBox3.ValueMember = "object_id"
        ComboBox3.DisplayMember = "name"

        '导线成果
        Dim sqlstr3 = "SELECT name,object_id FROM " + storename + ".sys.tables WHERE charindex('导线成果',name)>0 "
        Dim myCommand3 As New SqlCommand(sqlstr3, con)
        Dim reader3 As SqlDataAdapter = New SqlDataAdapter(myCommand3)
        Dim tempDataSet3 As New DataTable
        reader3.Fill(tempDataSet3)
        dr3 = tempDataSet3.NewRow
        dr3("object_id") = 0
        dr3("name") = "--请选择--"
        tempDataSet3.Rows.InsertAt(dr3, 0)
        ComboBox4.DataSource = tempDataSet3
        ComboBox4.ValueMember = "object_id"
        ComboBox4.DisplayMember = "name"
        con.Close()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Label5.Visible = True
        'Me.Hide()
        TorF = True
        If action(2) = "系统管理员" Then
            storeku1 = storeku.Remove(storeku.LastIndexOf("数据库"))
            storename = storeku1 & "线路参数"
        Else
            storename = action(2) & "线路参数"
        End If
        'Dim con As New SqlConnection("Server=192.168.10.36;DataBase=" + storename + ";integrated security=false;uid=sa;pwd=123;Encrypt=False;MultipleActiveResultSets=True;")
        Dim con As New SqlConnection("Server=renfeng.tpddns.cn,1000;DataBase=" + storename + ";integrated security=false;uid=sa;pwd=sa;Encrypt=false;Trusted_Connection=false")
        con.Open()

        ExApp = CreateObject("Excel.Application")
        ExApp.Visible = True
        Exbook = ExApp.Workbooks.Open(FileName1,,,, 123)


        If ComboBox1.Text <> "--请选择--" Then
            Dim sqlstr = "SELECT * FROM " + ComboBox1.Text + ""
            Dim myCommand As New SqlCommand(sqlstr, con)
            Dim reader As SqlDataAdapter = New SqlDataAdapter(myCommand)
            Dim tempDataSet As New DataTable
            reader.Fill(tempDataSet)
            Try
                sheet1 = Exbook.Worksheets("交点法")
                For i = 0 To tempDataSet.Rows.Count - 1
                    For j = 0 To tempDataSet.Columns.Count - 1
                        sheet1.Cells(i + 4, j + 1) =
                        tempDataSet.Rows(i).Item(j)
                    Next
                Next

            Catch
            End Try
        End If

        If ComboBox2.Text <> "--请选择--" Then

            Dim sqlstr = "SELECT * FROM " + ComboBox2.Text + ""
            Dim myCommand As New SqlCommand(sqlstr, con)
            Dim reader As SqlDataAdapter = New SqlDataAdapter(myCommand)
            Dim tempDataSet As New DataTable
            reader.Fill(tempDataSet)
            Try
                sheet2 = Exbook.Worksheets("线元法")
                For i = 0 To tempDataSet.Rows.Count - 1
                    For j = 0 To tempDataSet.Columns.Count - 1
                        sheet2.Cells(i + 3, j + 1) =
                            tempDataSet.Rows(i).Item(j)
                    Next
                Next
            Catch
            End Try
        End If

        If ComboBox3.Text <> "--请选择--" Then

            Dim sqlstr = "SELECT * FROM " + ComboBox3.Text + ""
            Dim myCommand As New SqlCommand(sqlstr, con)
            Dim reader As SqlDataAdapter = New SqlDataAdapter(myCommand)
            Dim tempDataSet As New DataTable
            reader.Fill(tempDataSet)
            Try
                sheet3 = Exbook.Worksheets("断链")
                For i = 0 To tempDataSet.Rows.Count - 1
                    For j = 0 To tempDataSet.Columns.Count - 1
                        sheet3.Cells(i + 3, j + 1) =
                            tempDataSet.Rows(i).Item(j)
                    Next
                Next
            Catch
            End Try
        End If

        If ComboBox4.Text <> "--请选择--" Then

            Dim sqlstr = "SELECT * FROM " + ComboBox4.Text + ""
            Dim myCommand As New SqlCommand(sqlstr, con)
            Dim reader As SqlDataAdapter = New SqlDataAdapter(myCommand)
            Dim tempDataSet As New DataTable
            reader.Fill(tempDataSet)

            sheet4 = Exbook.Worksheets("导线成果表")
            For i = 0 To tempDataSet.Rows.Count - 1
                For j = 0 To tempDataSet.Columns.Count - 1
                    sheet4.Cells(i + 3, j + 1) =
                        tempDataSet.Rows(i).Item(j)
                Next
            Next
        End If
        con.Close()
        Try
            sheet0 = Exbook.Worksheets("线元法") '数据库
            If ComboBox2.Text = "--请选择--" Then
                ComboBox1.Enabled = True
                sheet0.Range("J2").Value = ""
            Else
                ComboBox1.Enabled = False
                sheet0.Range("J2").Value = "是"
            End If
        Catch
        End Try
        Me.Close()
        交点法()
        上传数据()
        主页面.Panel6.BringToFront()
        主页面.Panel3.BringToFront()
        主页面.Panel3.Enabled = False
        主页面.Timer1.Enabled = True
        主页面.Timer1.Interval = 100
        主页面.ProgressBar1.Visible = True
        主页面.Label1.Visible = True
        主页面.Button8.Visible = True
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Text = "--请选择--" Then
            ComboBox2.Enabled = True
        Else
            ComboBox2.Enabled = False
        End If
    End Sub

End Class