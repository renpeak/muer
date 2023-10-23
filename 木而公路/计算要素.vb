Imports Microsoft.Office.Interop.Excel

Module 计算要素


    Sub 交点法()
        Exbook = ExApp.ActiveWorkbook
        sheet0 = Exbook.Worksheets("交点法") '数据库
        sheet1 = Exbook.Worksheets("线元法") '数据库

        Dim i As Integer, Jsfwj2, Jsfwj_d, Jsfwj_f, Jsfwj_m, Zj2, Ly2, p21, p22  'Jsfwj2 方位角,Jsfwj_d,Jsfwj_f,Jsfwj_m 方位角 度 分 秒,Zj2 转角值,Ly2圆曲线长，P2切线长，
        sheet0.Range("H5:S500").Value = Nothing
        sheet1.Range("j2").Value = "否"
        '方位角计算
        i = 5
        While sheet0.Range("b" & i).Value <> Nothing And sheet0.Range("c" & i).Value <> Nothing
            Jsfwj2 = Jsfwj_1j(Val(sheet0.Range("b" & i - 1).Value), Val(sheet0.Range("c" & i - 1).Value), Val(sheet0.Range("b" & i).Value), Val(sheet0.Range("c" & i).Value)) '方位角计算
            Jsfwj_d = Fix(Jsfwj2)
            Jsfwj_f = Fix((Jsfwj2 - Fix(Jsfwj2)) * 60)
            Jsfwj_m = Math.Round(((Jsfwj2 - Fix(Jsfwj2)) - Fix((Jsfwj2 - Fix(Jsfwj2)) * 60) / 60) * 3600, 0)
            sheet0.Range("k" & i).Value = Jsfwj_d & "°" & Math.Abs(Jsfwj_f) & "'" & Math.Abs(Jsfwj_m) & """" '方位角计算
            sheet0.Range("p" & i).Value = Jsfwj2
            i = i + 1
        End While
        sheet0.Range("l" & i - 1).Value = sheet0.Range("d" & i - 1).Value
        i = 5
        While sheet0.Range("c" & i + 1).Value <> Nothing And sheet0.Range("b" & i + 1).Value <> Nothing
            Zj2 = Val(sheet0.Range("p" & i + 1).Value) - Val(sheet0.Range("p" & i).Value) '转角值计算
            If Zj2 < -180 Then
                Zj2 = Zj2 + 360
            Else
                If Zj2 > 180 Then
                    Zj2 = Zj2 - 360
                End If
            End If
            Ly2 = Yqx_Ly(Val(sheet0.Range("E" & i).Value), Val(sheet0.Range("F" & i).Value), Val(sheet0.Range("G" & i).Value), Zj2) '圆曲线长计算
            p21 = Qx_T(Val(sheet0.Range("E" & i).Value), Val(sheet0.Range("F" & i).Value), Val(sheet0.Range("G" & i).Value), Zj2)  '切线长计算
            p22 = Qx_T2(Val(sheet0.Range("E" & i).Value), Val(sheet0.Range("F" & i).Value), Val(sheet0.Range("G" & i).Value), Zj2) '切线长计算
            Jsfwj_d = Fix(Zj2)
            Jsfwj_f = Fix((Zj2 - Fix(Zj2)) * 60)
            Jsfwj_m = Math.Round(((Zj2 - Fix(Zj2)) - Fix((Zj2 - Fix(Zj2)) * 60) / 60) * 3600, 0)
            ' sheet0.Range("d" & i) = Round( sheet0.Range("d" & i - 1) + Sqr(( sheet0.Range("b" & i - 1) -  sheet0.Range("b" & i)) ^ 2 + ( sheet0.Range("c" & i - 1) -  sheet0.Range("c" & i)) ^ 2) + ( sheet0.Range("F" & i - 1) +  sheet0.Range("G" & i - 1) +  sheet0.Range("i" & i - 1) -  sheet0.Range("r" & i - 1) -  sheet0.Range("s" & i - 1)), 3)
            sheet0.Range("H" & i).Value = Jsfwj_d & "°" & Math.Abs(Jsfwj_f) & "'" & Math.Abs(Jsfwj_m) & """"  '转角值
            sheet0.Range("q" & i).Value = Zj2 '转角值
            sheet0.Range("i" & i).Value = Math.Round(Ly2, 3) '圆曲线
            If sheet0.Range("F" & i).Value = sheet0.Range("G" & i).Value Then '切线长
                sheet0.Range("j" & i).Value = Math.Round(p21, 3)
                sheet0.Range("r" & i).Value = Math.Round(p21, 3)
                sheet0.Range("s" & i).Value = Math.Round(p21, 3)
            Else
                sheet0.Range("j" & i).Value = Math.Round(p21, 3) & "/" & Math.Round(p22, 3) '切线
                sheet0.Range("r" & i).Value = Math.Round(p21, 3)
                sheet0.Range("s" & i).Value = Math.Round(p22, 3)
            End If
            sheet0.Range("l" & i).Value = Math.Round(sheet0.Range("d" & i).Value - Math.Round(p21, 3), 3)
            sheet0.Range("m" & i).Value = Math.Round(sheet0.Range("l" & i).Value + sheet0.Range("F" & i).Value, 3)
            sheet0.Range("n" & i).Value = Math.Round(sheet0.Range("m" & i).Value + sheet0.Range("i" & i).Value, 3)
            sheet0.Range("o" & i).Value = Math.Round(sheet0.Range("n" & i).Value + sheet0.Range("G" & i).Value, 3)
            i = i + 1
        End While
        ' sheet0.Range("d" & i).value = math.Round( sheet0.Range("d" & i - 1).value + Sqr(( sheet0.Range("b" & i - 1).value -  sheet0.Range("b" & i).value) ^ 2 + ( sheet0.Range("c" & i - 1).value -  sheet0.Range("c" & i).value) ^ 2) + ( sheet0.Range("F" & i - 1).value +  sheet0.Range("G" & i - 1).value +  sheet0.Range("i" & i - 1).value -  sheet0.Range("r" & i - 1) .value-  sheet0.Range("s" & i - 1).value), 3)

    End Sub



    Sub 线元法()
        Exbook = ExApp.ActiveWorkbook
        sheet0 = Exbook.Worksheets("交点法") '数据库
        sheet1 = Exbook.Worksheets("线元法") '数据库

        sheet1.Range("B3:I32767").Value = Nothing
        Dim i, j, zx1, zx2, zy1, zy2, zxd, pdo
        pdo = 0
        i = 5
        j = 3
        While sheet0.Range("b" & i).Value <> Nothing
            '直线
            If sheet0.Range("L" & i + 1).Value - sheet0.Range("d" & i).Value <> 0 Then
                If i = 5 Then
                    sheet1.Range("b" & j).Value = sheet0.Range("D" & i - 1).Value '起点桩号
                    sheet1.Range("c" & j).Value = sheet0.Range("b" & i - 1).Value '起点坐标X
                    sheet1.Range("d" & j).Value = sheet0.Range("C" & i - 1).Value '起点坐标Y
                    sheet1.Range("e" & j).Value = Zxfwj_B0(sheet1.Range("b" & j).Value) '起点方位角
                    sheet1.Range("f" & j).Value = sheet0.Range("L" & i).Value - sheet0.Range("d" & i - 1).Value '线元长
                    sheet1.Range("g" & j).Value = 0 '开始半径
                    sheet1.Range("h" & j).Value = 0 '结束半径
                    sheet1.Range("i" & j).Value = "QD" '备注
                    j = j + 1
                Else
                    zx1 = ZSZB_X0j(sheet0.Range("l" & i).Value, 0, 90)
                    zy1 = ZSZB_Y0j(sheet0.Range("l" & i).Value, 0, 90)
                    zx2 = ZSZB_X0j("*" & sheet0.Range("o" & i - 1).Value, 0, 90)
                    zy2 = ZSZB_Y0j("*" & sheet0.Range("o" & i - 1).Value, 0, 90)
                    zxd = Math.Round(Math.Sqrt((zx1 - zx2) ^ 2 + (zy1 - zy2) ^ 2), 4)
                    If zxd > 0 Then
                        sheet1.Range("b" & j).Value = sheet0.Range("o" & i - 1).Value  '起点桩号
                        sheet1.Range("c" & j).Value = ZSZB_X0j("*" & sheet1.Range("b" & j).Value, 0, 90) '起点坐标X
                        sheet1.Range("d" & j).Value = ZSZB_Y0j("*" & sheet1.Range("b" & j).Value, 0, 90) '起点坐标Y
                        sheet1.Range("e" & j).Value = Zxfwj_B0("*" & sheet1.Range("b" & j).Value + 0.00001) '起点走向方位角
                        sheet1.Range("f" & j).Value = zxd '线元长
                        sheet1.Range("g" & j).Value = 0 ' zy2 '0 '开始半径
                        sheet1.Range("h" & j).Value = 0 '结束半径
                        sheet1.Range("i" & j).Value = "直线" '备注
                        j = j + 1
                    End If
                End If

            End If
            '缓和1
            If Val(sheet0.Range("F" & i).Value) > 0 Then
                sheet1.Range("b" & j).Value = sheet0.Range("l" & i).Value  '起点桩号
                sheet1.Range("c" & j).Value = ZSZB_X0j("*" & sheet1.Range("b" & j).Value, 0, 90) '起点坐标X
                sheet1.Range("d" & j).Value = ZSZB_Y0j("*" & sheet1.Range("b" & j).Value, 0, 90) '起点坐标Y
                sheet1.Range("e" & j).Value = Zxfwj_B0("*" & sheet1.Range("b" & j).Value + 0.00001)  '起点走向方位角
                sheet1.Range("f" & j).Value = sheet0.Range("F" & i).Value  '线元长
                sheet1.Range("g" & j).Value = 0 '开始半径
                If sheet0.Range("q" & i).Value < 0 Then
                    sheet1.Range("h" & j).Value = -sheet0.Range("E" & i).Value  '结束半径
                Else
                    sheet1.Range("h" & j).Value = sheet0.Range("E" & i).Value  '结束半径
                End If
                sheet1.Range("i" & j).Value = "缓和1" '备注
                j = j + 1
            End If
            '圆
            If Val(sheet0.Range("E" & i).Value) > 0 Then
                sheet1.Range("b" & j).Value = sheet0.Range("m" & i).Value '起点桩号
                sheet1.Range("c" & j).Value = ZSZB_X0j("*" & sheet1.Range("b" & j).Value, 0, 0) '起点坐标X
                sheet1.Range("d" & j).Value = ZSZB_Y0j("*" & sheet1.Range("b" & j).Value, 0, 0) '起点坐标Y
                sheet1.Range("e" & j).Value = Zxfwj_B0("*" & sheet1.Range("b" & j).Value + 0.00001)  '起点走向方位角
                sheet1.Range("f" & j).Value = sheet0.Range("i" & i).Value '线元长
                If sheet0.Range("q" & i).Value < 0 Then
                    sheet1.Range("g" & j).Value = -sheet0.Range("E" & i).Value '结束半径 '开始半径
                    sheet1.Range("h" & j).Value = -sheet0.Range("E" & i).Value  '结束半径 '结束半径
                Else
                    sheet1.Range("g" & j).Value = sheet0.Range("E" & i).Value '结束半径 '开始半径
                    sheet1.Range("h" & j).Value = sheet0.Range("E" & i).Value '结束半径 '结束半径
                End If
                If sheet0.Range("E" & i).Value <= 1 Then '判断半径小于0
                    pdo = 1
                End If
                sheet1.Range("i" & j).Value = "圆" '备注
                j = j + 1
            End If
            '缓和2
            If Val(sheet0.Range("G" & i).Value) > 0 Then
                sheet1.Range("b" & j).Value = sheet0.Range("m" & i).Value  '起点桩号
                sheet1.Range("c" & j).Value = ZSZB_X0j("*" & sheet1.Range("b" & j).Value, 0, 0) '起点坐标X
                sheet1.Range("d" & j).Value = ZSZB_Y0j("*" & sheet1.Range("b" & j).Value, 0, 0) '起点坐标Y
                sheet1.Range("e" & j).Value = Zxfwj_B0("*" & sheet1.Range("b" & j).Value + 0.00001)  '起点走向方位角
                sheet1.Range("f" & j).Value = sheet0.Range("G" & i).Value '线元长
                If sheet0.Range("q" & i).Value < 0 Then
                    sheet1.Range("g" & j).Value = -sheet0.Range("E" & i).Value  '结束半径 '开始半径
                Else
                    sheet1.Range("g" & j).Value = sheet0.Range("E" & i).Value  '结束半径 '开始半径
                End If
                sheet1.Range("h" & j).Value = 0 '结束半径
                sheet1.Range("i" & j).Value = "缓和2" '备注
                j = j + 1
            End If
            i = i + 1
        End While
        If pdo = 1 Then
            MsgBox("部分线元有误，请手动调整（R<1）")
        End If


    End Sub



    Sub 竖曲线()

        Exbook = ExApp.ActiveWorkbook
        sheet0 = Exbook.Worksheets("竖曲线") '数据库
        sheet1 = Exbook.Worksheets("断链") '数据库

        Dim n As Integer, i, pd, dL
        sheet0.Range("e5 : j328").Value = Nothing '清屏
        If sheet0.Range("b" & 4).Value = vbNull Or sheet0.Range("c" & 4).Value = Nothing Then
            MsgBox（"请输数据", vbOKOnly, "无数据提示"）
            Exit Sub
        Else
            If sheet0.Range("b" & 4).Value > Val(sheet0.Range("b" & 4).Value) Or sheet0.Range("c" & 4).Value > Val(sheet0.Range("c" & 4).Value) Then
                MsgBox（"请在第4行正确输入纯数字", vbOKOnly, "数据错误提示"）
                Exit Sub
            End If
        End If
        n = 5
        While sheet0.Range("b" & n).Value <> Nothing And sheet0.Range("c" & n).Value <> Nothing
            If sheet0.Range("b" & n).Value > Val(sheet0.Range("b" & n).Value) Or sheet0.Range("c" & n).Value > Val(sheet0.Range("c" & n).Value) Or sheet0.Range("d" & n).Value > Val(sheet0.Range("d" & n).Value) Then
                MsgBox（"请在第" & n & "行正确输入纯数字", vbOKOnly, "数据错误提示"）
                Exit Sub
            Else
                If sheet0.Range("c" & (n + 1)).Value = Nothing Then
                    i = 3
                    pd = 1
                    dL = 0
                    While sheet1.Range("b" & i).Value <> Nothing And pd = 1
                        If sheet1.Range("b" & i).Value >= sheet0.Range("b" & n - 1).Value And sheet1.Range("c" & i).Value < sheet0.Range("b" & n).Value Then
                            dL = sheet1.Range("c" & i).Value - sheet1.Range("b" & i).Value
                            pd = 0
                        End If
                        i = i + 1
                    End While
                    sheet0.Range("j" & n).Value = (sheet0.Range("c" & n).Value - sheet0.Range("c" & (n - 1)).Value) / (sheet0.Range("b" & n).Value - sheet0.Range("b" & (n - 1)).Value - dL) '纵坡
                Else
                    i = 3
                    pd = 1
                    dL = 0
                    While sheet1.Range("b" & i).Value <> Nothing And pd = 1
                        If sheet1.Range("b" & i).Value >= sheet0.Range("b" & n - 1).Value And sheet1.Range("c" & i).Value < sheet0.Range("b" & n).Value Then
                            dL = sheet1.Range("c" & i).Value - sheet1.Range("b" & i).Value
                            pd = 0
                        End If
                        i = i + 1
                    End While
                    sheet0.Range("j" & n).Value = (sheet0.Range("c" & n).Value - sheet0.Range("c" & (n - 1)).Value) / (sheet0.Range("b" & n).Value - sheet0.Range("b" & (n - 1)).Value - dL) '纵坡

                End If
            End If
            n = n + 1
        End While

        n = 5
        While sheet0.Range("b" & n + 1).Value <> Nothing And sheet0.Range("c" & n + 1).Value <> Nothing
            sheet0.Range("i" & n).Value = sheet0.Range("j" & n).Value - sheet0.Range("j" & n + 1).Value '转坡角
            sheet0.Range("e" & n).Value = sheet0.Range("d" & n).Value * Math.Abs(sheet0.Range("i" & n).Value) / 2 '切线长
            sheet0.Range("f" & n).Value = sheet0.Range("e" & n).Value ^ 2 / sheet0.Range("d" & n).Value / 2 '外距
            sheet0.Range("g" & n).Value = sheet0.Range("b" & n).Value - sheet0.Range("e" & n).Value '竖曲线起点
            sheet0.Range("h" & n).Value = sheet0.Range("b" & n).Value + sheet0.Range("e" & n).Value '竖曲线终点

            n = n + 1
        End While


    End Sub

End Module
