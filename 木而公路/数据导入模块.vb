Imports Microsoft.Data.SqlClient
Imports System.Data.OleDb
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Module 数据导入模块

    '数据表查询后导入到模板数据库
    Sub Examine(tablename As String)
        'Try
        Dim FolderDialogObject As New FolderBrowserDialog()
        FolderDialogObject.Description = "请选择保存文件的目录"
        FolderDialogObject.UseDescriptionForTitle = True
        FolderDialogObject.ShowDialog()   '显示选择文件夹对话框 
        Filepath = FolderDialogObject.SelectedPath
        If Filepath = Nothing Then
            Exit Sub
        End If

        主页面.OpenFileDialog1.Title = "请选择需要导入的数据表"
        主页面.OpenFileDialog1.Filter = "Excel文件|*.xls;*.xlsx;*.xlt;*.xltx;*.xltm;*.xlsm"
        主页面.OpenFileDialog1.FilterIndex = 1
        主页面.OpenFileDialog1.FileName = ""

        If 主页面.OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            filenamed = 主页面.OpenFileDialog1.FileName
            storeku = 主页面.ComboBox7.Text
            Dim storeku1 As String
            'fileName1 = "" + Application.StartupPath + "木而公路模板表/" + action(2) + "\" + projectname(0) + "\" + projectname(1) + ".xlsx"
            System.Environment.CurrentDirectory = My.Application.Info.DirectoryPath
            FileName2 = System.IO.Path.GetFullPath("../../../")
            If 主页面.Label9.Text = "系统管理员" Then
                storeku1 = 主页面.ComboBox7.Text.Remove(主页面.ComboBox7.Text.LastIndexOf("数据库"))
                FileName1 = "" + FileName2 + "木而公路模板表/" + storeku1 + "/" + projectname(0) + "/" + tablename + ".xlsx"
            Else
                FileName1 = "" + FileName2 + "木而公路模板表/" + action(2) + "/" + projectname(0) + "/" + tablename + ".xlsx"
            End If

            '取出选取数据表中的数据
            线路参数.ShowDialog()

            If TorF = False Then
                Exit Sub
            End If
        Else
            Exit Sub
        End If
        'Catch
        '    MsgBox("系统错误，请联系系统管理员！！！")
        'End Try
    End Sub



    '上传数据到服务器数据库
    Sub 上传数据（）
        '建立EXCEL连接，读入数据，支持 Microsoft Excel 2010
        Dim strConn As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & filenamed & "';Extended Properties='Excel 12.0;HDR=NO;'"
        Dim cnn As OleDbConnection = New OleDbConnection(strConn)
        cnn.Open()
        Dim dt As DataTable = cnn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
        Dim workSheetName As String = dt.Rows(0)("TABLE_NAME")
        Dim oda As New OleDb.OleDbDataAdapter("Select * FROM [" + workSheetName + "A1:H5]", strConn)
        Dim ds1 As DataSet = New DataSet
        oda.Fill(ds1, storebase)
        Dim da As New OleDb.OleDbDataAdapter("Select * FROM [" + workSheetName + "A8:AZ10000]", strConn)
        Dim ds As DataSet = New DataSet
        da.Fill(ds, storebase)

        '打开模板表后上传数据
        sheet0 = Exbook.Worksheets("数据库")
        For i = 0 To ds1.Tables(0).Rows.Count - 1
            For j = 0 To ds1.Tables(0).Columns.Count - 1
                sheet0.Cells(i + 1, j + 1) =
                ds1.Tables(0).Rows(i).Item(j)
            Next
        Next
        For i = 0 To ds.Tables(0).Rows.Count - 1
            For j = 0 To ds.Tables(0).Columns.Count - 1
                sheet0.Cells(i + 8, j + 1) =
                ds.Tables(0).Rows(i).Item(j)
            Next
        Next
        Dim CnnStr As String
        If action(2) = "系统管理员" Then
            CnnStr = "Server=renfeng.tpddns.cn,1000;DataBase='" + storeku + "';integrated security=false;uid=sa;pwd=sa;Encrypt=False;Trusted_Connection=false"
            cnn.Close()
        Else
            CnnStr = "Server=renfeng.tpddns.cn,1000;DataBase='" + action(3) + "';integrated security=false;uid=sa;pwd=sa;Encrypt=False;Trusted_Connection=false"

            cnn.Close()
        End If

        Dim bcp As SqlBulkCopy
        bcp = New SqlBulkCopy(CnnStr)
        bcp.BatchSize = 100    '每次传输的行数
        bcp.DestinationTableName = storebase    '目标表
        bcp.WriteToServer(ds.Tables(0))

    End Sub


End Module
