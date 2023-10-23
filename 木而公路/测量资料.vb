
Imports Microsoft.Office.Interop.Excel
Module 测量资料

    Sub 钢筋底高程（）
        Dim t, a, i1, j, i2, i3, i4 As Integer, z(238) As Object, s(238) As Object
        Dim HSL, ZSL, QSL, SXGL, SCGL, BZL, SJH, SJGL, PCL, BGH, SJWZH, sumz, sumh, ZH0, min1, min2, szdz, szdh, szdm, Ming, Maxg, TCmind, TCminx, TCmaxd, TCmaxx

        Exbook = ExApp.ActiveWorkbook
        sheet0 = Exbook.Worksheets("数据库") '数据库
        sheet1 = Exbook.Worksheets("参数表") '参数表
        sheet2 = Exbook.Worksheets("导线成果表") '导线成果表
        sheet3 = Exbook.Worksheets("钢筋检表") '钢筋检表
        sheet4 = Exbook.Worksheets("钢筋顶水准") '钢筋顶水准

        Try
            HSL = "b" '后视列
            ZSL = "c" '中视列
            QSL = "d" '前视列
            SXGL = "e" '视线高列
            SCGL = "f" '实测高列
            SJGL = "g" '设计高列
            PCL = "h" '偏差列
            BZL = "i" '备注列
            SJH = 23  '数据行数，表格计算数据的行数
            BGH = 32 '表格总行数
            SJWZH = 9 '数据位置开始行，表格数据从哪行开始
            TCmind = 1000 '塔尺小读数大值
            TCminx = 500 '塔尺小读数小值
            TCmaxd = 4600 '塔尺大读数大值
            TCmaxx = 3200 '塔尺大读数小值
            j = 1
            i1 = 3      '数据表的开始行

            j = 1
            i1 = 3      '数据表的开始行

            SJWZH = SJWZH - 2
            Ming = -10
            Maxg = 10
            i2 = 3
        min1 = 1000  '允许桩号偏差范围内的控制点
        min2 = 100   '允许高程偏差范围内的控制点
        szdz = 0
            szdh = 0
            szdm = 0
            j = j + BGH '表格行数''
            sheet4.Select()
            CType(sheet4.Rows("33:99999"), Range).Delete()  '删除表格''
            CType(sheet4.Rows("1:32"), Range).Copy()            '复制表格''
            sheet4.Range("a" & j).PasteSpecial()
            '钢筋顶水准点桩号
            sumz = Val(ExApp.WorksheetFunction.Substitute(sheet1.Range("J3").Value, "*", ""))
            sumh = Val(sheet1.Range("B32").Value)
            For i2 = 3 To sheet2.Range("B1048576").End(XlDirection.xlUp).Row + 1


                If Math.Abs(sumz - sheet2.Range("F" & i2).Value) < min1 And Math.Abs(sumh - sheet2.Range("E" & i2).Value) < min2 And sheet2.Range("E" & i2).Value <> Nothing Then
                    'min1 = Math.Abs(sumz - sheet2.Range("F" & i2).value)    '就近取点、隧道时启用
                    min2 = Math.Abs(sumh - sheet2.Range("E" & i2).Value)
                    szdz = sheet2.Range("F" & i2).Value '水准点桩号
                    szdh = sheet2.Range("E" & i2).Value '水准点高程
                    szdm = sheet2.Range("B" & i2).Value '水准点名称

                End If
            Next i2

            '开始''''''''''''''''''''''''''''
            '开始''''''''''''''''''''''''''''
            '开始''''''''''''''''''''''''''''
            i4 = 1
            '' 改，表头数据绝对行均减2
            sheet4.Range("E" & j + 3 + i4).Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & sheet1.Range("F6").Value & "钢筋" '工程部位
            sheet4.Range("I" & j + 3 + i4).Value = sheet3.Range("S7").Value '检测日期

            If szdh > sheet1.Range("B32").Value Then
                Randomize()
                sheet4.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)   '后视
            Else
                Randomize()
                sheet4.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)
            End If
            sheet4.Range(SJGL & j + SJWZH + i4).Value = szdh  '水准点高程
            sheet4.Range("A" & j + SJWZH + i4).Value = szdm  '水准点名称
            sheet4.Range(SXGL & j + SJWZH + i4).Value = sheet4.Range(SJGL & j + SJWZH + i4).Value + sheet4.Range(HSL & j + SJWZH + i4).Value      '视线高
            a = 0
            s(a) = Val(sheet4.Range(SXGL & j + SJWZH + i4).Value)
            z(a) = Val(szdz) '桩号值
            '计算
            i2 = 1

            i3 = 0
            i4 = i4 + 1
            ''''
            For t = i1 To i1 + i2 - 1

                '设转点
                ZH0 = Val(ExApp.WorksheetFunction.Substitute(sheet1.Range("J3").Value, "*", ""))

                If Val(ZH0) - z(a) > 200 Then '5
                    sheet4.Range("a" & j + SJWZH + i4).Value = "ZD" & (a + 1)   '桩号
                    If s(a) - Val(sheet1.Range("B32").Value) > 4.5 Then
                        Randomize()
                        sheet4.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)     '前视
                        Randomize()
                        sheet4.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)    '后视
                    Else
                        If s(a) - Val(sheet1.Range("B32").Value) < 0.3 Then
                            Randomize()
                            sheet4.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)   '前视
                            Randomize()
                            sheet4.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)    '后视
                        Else
                            Randomize()
                            sheet4.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)   '前视
                            Randomize()
                            sheet4.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)    '后视
                        End If
                    End If
                    sheet4.Range(SCGL & j + SJWZH + i4).Value = s(a) - Val(sheet4.Range(QSL & j + SJWZH + i4).Value)       '实测高程
                    a = a + 1 '交点数加1
                    s(a) = Val(sheet4.Range(SCGL & j + SJWZH + i4).Value) + Val(sheet4.Range(HSL & j + SJWZH + i4).Value)       '视线高
                    sheet4.Range(SXGL & j + SJWZH + i4).Value = s(a)
                    sheet4.Range(SCGL & j + SJWZH + i4).Value = ""   ''''''''实测高
                    z(a) = z(a - 1) + 200
                    i4 = i4 + 1
                    t = t - 1


                Else '距离设转点
                    If Val(ZH0) - z(a) < -200 Then '4
                        sheet4.Range("a" & j + SJWZH + i4).Value = "ZD" & (a + 1)  '桩号
                        If s(a) - Val(sheet1.Range("B32").Value) > 4.5 Then
                            Randomize()
                            sheet4.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)     '前视
                            Randomize()
                            sheet4.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)    '后视
                        Else
                            If s(a) - Val(sheet1.Range("B32").Value) < 0.3 Then
                                Randomize()
                                sheet4.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)   '前视
                                Randomize()
                                sheet4.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)    '后视
                            Else
                                Randomize()
                                sheet4.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)   '前视
                                Randomize()
                                sheet4.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)    '后视
                            End If
                        End If
                        sheet4.Range(SCGL & j + SJWZH + i4).Value = s(a) - Val(sheet4.Range(QSL & j + SJWZH + i4).Value)       '实测高程
                        a = a + 1 '交点数加1
                        s(a) = Val(sheet4.Range(SCGL & j + SJWZH + i4).Value) + Val(sheet4.Range(HSL & j + SJWZH + i4).Value)       '视线高
                        sheet4.Range(SXGL & j + SJWZH + i4).Value = s(a)
                        sheet4.Range(SCGL & j + SJWZH + i4).Value = ""    ''''''''实测高
                        z(a) = z(a - 1) - 200
                        i4 = i4 + 1
                        t = t - 1


                    Else '设转点
                        If s(a) - Val(sheet1.Range("B32").Value) > 4.5 Then '3
                            sheet4.Range("a" & j + SJWZH + i4).Value = "ZD" & (a + 1)     '桩号
                            Randomize()
                            sheet4.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)     '前视
                            sheet4.Range(SCGL & j + SJWZH + i4).Value = s(a) - Val(sheet4.Range(QSL & j + SJWZH + i4).Value)       '实测高程
                            Randomize()
                            sheet4.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)     '后视
                            a = a + 1 '交点数加1
                            s(a) = Val(sheet4.Range(SCGL & j + SJWZH + i4).Value) + Val(sheet4.Range(HSL & j + SJWZH + i4).Value)       '视线高
                            sheet4.Range(SXGL & j + SJWZH + i4).Value = s(a)
                            sheet4.Range(SCGL & j + SJWZH + i4).Value = ""    ''''''''实测高
                            z(a) = z(a - 1) '   桩号值
                            i4 = i4 + 1
                            t = t - 1

                        Else '设转点
                            If s(a) - Val(sheet1.Range("B32").Value) < 0.3 Then  '2
                                sheet4.Range("a" & j + SJWZH + i4).Value = "ZD" & (a + 1)   '桩号
                                Randomize()
                                sheet4.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)    '前视
                                sheet4.Range(SCGL & j + SJWZH + i4).Value = s(a) - sheet4.Range(QSL & j + SJWZH + i4).Value      '实测高程
                                Randomize()
                                sheet4.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)    '后视
                                a = a + 1 '交点数加1
                                s(a) = Val(sheet4.Range(SCGL & j + SJWZH + i4).Value) + Val(sheet4.Range(HSL & j + SJWZH + i4).Value)      '视线高
                                sheet4.Range(SXGL & j + SJWZH + i4).Value = s(a)
                                sheet4.Range(SCGL & j + SJWZH + i4).Value = ""    ''''''''实测高
                                z(a) = z(a - 1)  '     桩号值
                                i4 = i4 + 1
                                t = t - 1

                            Else '非转点求值
                                If t <= i1 + i2 Then '1
                                    sheet4.Range("a" & j + SJWZH + i4).Value = "骨架顶"
                                    sheet4.Range(SJGL & j + SJWZH + i4).Value = Val(sheet1.Range("B30").Value)     '设计高程
                                    'sheet4.Range(SCGL & j + SJWZH + i4).value = Math.Round(sheet4.Range(SJGL & j + SJWZH + i4).value + sheet4.Range(PCL & j + SJWZH + i4).value / 1000, 3)          '实测高程
                                    sheet4.Range(SCGL & j + SJWZH + i4).Value = sheet1.Range（"B32"）.Value   '实测高程
                                    sheet4.Range(PCL & j + SJWZH + i4).Value = (sheet4.Range(SCGL & j + SJWZH + i4).Value - sheet4.Range(SJGL & j + SJWZH + i4).Value) * 1000   '偏差
                                    sheet4.Range(ZSL & j + SJWZH + i4).Value = Math.Round(s(a) - Val(sheet4.Range(SCGL & j + SJWZH + i4).Value), 3)     '中视
                                    sheet4.Range(BZL & j + SJWZH + i4).Value = "骨架长" & Math.Round(sheet1.Range("B13").Value, 3) & "m"  '备注

                                    i4 = i4 + 1
                                End If '1
                            End If '2
                        End If '3
                    End If '4
                End If '5

                '新增表格
                If i4 > SJH Then
                    j = j + BGH  '表格行数''
                    sheet4.Range("a" & j).PasteSpecial()
                    i4 = 1
                    '''''''''''''''''''''''改改改改改改改改
                    sheet4.Range("e" & j + 3 + i4).Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & sheet1.Range("F6").Value & "钢筋" '工程部位
                    sheet4.Range("i" & j + 3 + i4).Value = sheet3.Range("S7").Value '检测日期
                    sheet4.Range("E" & j + i4 + 1).Value = "续上页" '编号
                    sheet4.Range("E" & j + i4 + 1).Interior.Color = 255
                    sheet4.Range("E" & j + i4 + 1).Font.Color = -16711681
                    ''''''''''''''''''''''''''''''''''
                End If

            Next t

            ''''闭合'''''''''''''''''''''''''''''''''''''''''''
            For t = i1 + i2 - 1 To i1 + i2
                '设转点

                If Val(szdz) - z(a) > 200 Then '设转点
                    sheet4.Range("a" & j + SJWZH + i4).Value = "ZD" & (a + 1) '桩号
                    If s(a) < Val(szdh) Then
                        Randomize()
                        sheet4.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)     '前视
                        Randomize()
                        sheet4.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)    '后视
                    Else
                        Randomize()
                        sheet4.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)    '前视
                        Randomize()
                        sheet4.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)     '后视
                    End If
                    sheet4.Range(SCGL & j + SJWZH + i4).Value = s(a) - Val(sheet4.Range(QSL & j + SJWZH + i4).Value)      '实测高程
                    a = a + 1 '交点数加1
                    s(a) = Val(sheet4.Range(SCGL & j + SJWZH + i4).Value) + Val(sheet4.Range(HSL & j + SJWZH + i4).Value)    '视线高
                    sheet4.Range(SXGL & j + SJWZH + i4).Value = s(a)
                    sheet4.Range(SCGL & j + SJWZH + i4).Value = ""    ''''''''实测高
                    z(a) = z(a - 1) + 200
                    i4 = i4 + 1
                    t = t - 1
                Else '设转点
                    If Val(szdz) - z(a) < -200 Then
                        sheet4.Range("a" & j + SJWZH + i4).Value = "ZD" & (a + 1)   '桩号
                        If s(a) < Val(szdh) Then
                            Randomize()
                            sheet4.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)     '前视
                            Randomize()
                            sheet4.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)    '后视
                        Else
                            Randomize()
                            sheet4.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)    '前视
                            Randomize()
                            sheet4.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)     '后视
                        End If
                        sheet4.Range(SCGL & j + SJWZH + i4).Value = s(a) - Val(sheet4.Range(QSL & j + SJWZH + i4).Value)       '实测高程
                        a = a + 1 '交点数加1
                        s(a) = Val(sheet4.Range(SCGL & j + SJWZH + i4).Value) + Val(sheet4.Range(HSL & j + SJWZH + i4).Value)     '视线高
                        sheet4.Range(SXGL & j + SJWZH + i4).Value = s(a)
                        sheet4.Range(SCGL & j + SJWZH + i4).Value = ""    ''''''''实测高
                        i4 = i4 + 1
                        t = t - 1
                        z(a) = z(a - 1) - 200
                    Else '设转点
                        If s(a) - Val(szdh) > 4.6 Then
                            sheet4.Range("a" & j + SJWZH + i4).Value = "ZD" & (a + 1)     '桩号
                            Randomize()
                            sheet4.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)     '前视
                            sheet4.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)    '后视
                            sheet4.Range(SCGL & j + SJWZH + i4).Value = s(a) - Val(sheet4.Range(QSL & j + SJWZH + i4).Value)       '实测高程
                            a = a + 1 '交点数加1
                            s(a) = Val(sheet4.Range(SCGL & j + SJWZH + i4).Value) + Val(sheet4.Range(HSL & j + SJWZH + i4).Value)       '视线高
                            sheet4.Range(SXGL & j + SJWZH + i4).Value = s(a)
                            sheet4.Range(SCGL & j + SJWZH + i4).Value = ""    ''''''''实测高
                            z(a) = z(a - 1)  '   桩号值
                            i4 = i4 + 1
                            t = t - 1
                        Else '设转点
                            If s(a) - Val(szdh) < 0.4 Then
                                sheet4.Range("a" & j + SJWZH + i4).Value = "ZD" & (a + 1)   '桩号
                                Randomize()
                                sheet4.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)     '前视
                                sheet4.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)   '后视
                                sheet4.Range(SCGL & j + SJWZH + i4).Value = s(a) - Val(sheet4.Range(QSL & j + SJWZH + i4).Value)       '实测高程
                                a = a + 1 '交点数加1
                                s(a) = Val(sheet4.Range(SCGL & j + SJWZH + i4).Value) + Val(sheet4.Range(HSL & j + SJWZH + i4).Value)     '视线高
                                sheet4.Range(SXGL & j + SJWZH + i4).Value = s(a)
                                sheet4.Range(SCGL & j + SJWZH + i4).Value = ""    ''''''''实测高
                                z(a) = z(a - 1)  ' 桩号值
                                i4 = i4 + 1
                                t = t - 1
                            Else
                                Randomize()
                                'sheet4.Range(PCL & j + SJWZH + i4).value = Int((p_z - p_f + 1) * Rnd() + p_f)   '偏差
                                sheet4.Range(PCL & j + SJWZH + i4).Value = 0   '偏差
                                sheet4.Range(SJGL & j + SJWZH + i4).Value = szdh    '设计高程
                                sheet4.Range(SCGL & j + SJWZH + i4).Value = Math.Round(Val(sheet4.Range(SJGL & j + SJWZH + i4).Value) + Val(sheet4.Range(PCL & j + SJWZH + i4).Value) / 1000, 3)       '实测高程
                                sheet4.Range(QSL & j + SJWZH + i4).Value = Math.Round(s(a) - Val(sheet4.Range(SCGL & j + SJWZH + i4).Value), 3)      '前视
                                sheet4.Range("a" & j + SJWZH + i4).Value = szdm

                                sheet4.Range("a" & j + SJWZH + i4 + 1).Value = "骨架底"
                                sheet4.Range(SCGL & j + SJWZH + i4 + 1).Value = sheet1.Range("B12").Value
                                sheet4.Range(SJGL & j + SJWZH + i4 + 1).Value = sheet1.Range("B31").Value
                                sheet4.Range(PCL & j + SJWZH + i4 + 1).Value = （sheet4.Range(SCGL & j + SJWZH + i4 + 1).Value - sheet4.Range(SJGL & j + SJWZH + i4 + 1).Value） * 1000

                            End If
                        End If
                    End If
                End If

                '新增表格
                If i4 > SJH Then
                    j = j + BGH '表格行数
                    sheet4.Range("a" & j).PasteSpecial()
                    i4 = 1
                    '‘’‘’‘’‘’‘改改改改改改改改
                    sheet4.Range("e" & j + 3 + i4).Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & sheet1.Range("F6").Value & "钢筋" '工程部位
                    sheet4.Range("i" & j + 3 + i4).Value = sheet3.Range("S7").Value '检测日期
                    sheet4.Range("E" & j + i4 + 1).Value = "续上页" '编号
                    sheet4.Range("E" & j + i4 + 1).Interior.Color = 255
                    sheet4.Range("E" & j + i4 + 1).Font.Color = -16711681

                End If
            Next t
            i1 = i1 + 1


            '非转点求值
            'Dim p_z As Single, p_f As Single
            ''p_z = 2 * Math.Sqrt(a + 1)
            ''p_f = -2 * Math.Sqrt(a + 1)
            ''If ExApp.Sheets("数据库").range("O2").value = "二等" Then
            ''    p_z = 0
            ''    p_f = 0
            ''Else
            ''If ExApp.Sheets("数据库").range("O2").value = "三等" Then
            'p_z = 2
            'p_f = -2
            ''Else
            ''If ExApp.Sheets("数据库").range("O2").value = "四等" Then
            ''                p_z = 3
            ''                p_f = -3
            ''            End If
            ''End If

            ''End If




            '备注列
            '        sheet4.Range("B" & j + 28).value = "闭合或附合差fh=" & sheet4.Range(PCL & j + SJWZH + i4).value
            'sheet4.Range("B" & j + 30).value = "|fh|<|f允|，符合水准测量技术规定。"
            'sheet4.Range("B" & j + 29).value = "容许误差f允=6√n=6√" & a + 1 & "=" & Math.Round(6 * Math.Sqrt(a + 1), 0)


        Catch Exclerror As Exception   '错误时弹出提示
            MsgBox(Exclerror.Message)
            TorF = False
            Exit Sub
        End Try
    End Sub

    Sub 孔底高程（）
        Dim t, a, i1, j, i2, i3, i4 As Integer, z(238) As Object, s(238) As Object
        Dim HSL, ZSL, QSL, SXGL, SCGL, BZL, SJH, SJGL, PCL, BGH, SJWZH, sumz, sumh, ZH0, min1, min2, szdz, szdh, szdm, Ming, Maxg, TCmind, TCminx, TCmaxd, TCmaxx

        Exbook = ExApp.ActiveWorkbook
        sheet0 = Exbook.Worksheets("数据库") '数据库
        sheet1 = Exbook.Worksheets("参数表") '参数表
        sheet2 = Exbook.Worksheets("导线成果表") '导线成果表
        sheet3 = Exbook.Worksheets("孔底水准") '孔底水准

        Try

            HSL = "b" '后视列
            ZSL = "c" '中视列
            QSL = "d" '前视列
            SXGL = "e" '视线高列
            SCGL = "f" '实测高列
            SJGL = "g" '设计高列
            PCL = "h" '偏差列
            BZL = "i" '备注列
            SJH = 23  '数据行数，表格计算数据的行数
            BGH = 32 '表格总行数
            SJWZH = 9 '数据位置开始行，表格数据从哪行开始
            TCmind = 1000 '塔尺小读数大值
            TCminx = 500 '塔尺小读数小值
            TCmaxd = 4600 '塔尺大读数大值
            TCmaxx = 3200 '塔尺大读数小值
            j = 1
            i1 = 3      '数据表的开始行
            j = 1
            i1 = 3      '数据表的开始行
            SJWZH = SJWZH - 2
            Ming = -10
            Maxg = 10
            i2 = 3
            min1 = 1000  '允许桩号偏差范围内的控制点
            min2 = 100   '允许高程偏差范围内的控制点
            szdz = 0
            szdh = 0
            szdm = 0
            j = j + BGH '表格行数''

            sheet3.Select()
            CType(sheet3.Rows("33:99999"), Range).Delete()  '删除表格''
            CType(sheet3.Rows("1:32"), Range).Copy()            '复制表格''
            sheet3.Range("a" & j).PasteSpecial()
            '孔底水准点桩号
            sumz = Val(ExApp.WorksheetFunction.Substitute(sheet1.Range("J3").Value, "*", ""))
            sumh = Val(sheet1.Range("B10").Value)

            For i2 = 3 To sheet2.Range("B1048576").End(XlDirection.xlUp).Row + 1

                If Math.Abs(sumz - sheet2.Range("F" & i2).Value) < min1 And Math.Abs(sumh - sheet2.Range("E" & i2).Value) < min2 And sheet2.Range("E" & i2).Value <> Nothing Then
                    'min1 = Math.Abs(sumz - sheet2.Range("F" & i2).value)   '就近取点、隧道时启用
                    min2 = Math.Abs(sumh - sheet2.Range("E" & i2).Value)
                    szdz = sheet2.Range("F" & i2).Value '水准点桩号
                    szdh = sheet2.Range("E" & i2).Value '水准点高程
                    szdm = sheet2.Range("B" & i2).Value '水准点名称

                End If
            Next i2

            '开始''''''''''''''''''''''''''''
            '开始''''''''''''''''''''''''''''
            '开始''''''''''''''''''''''''''''
            i4 = 1
            '' 改，表头数据绝对行均减2
            sheet3.Range("E" & j + 3 + i4).Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & sheet1.Range("F6").Value & "孔底" '工程部位
            sheet3.Range("I" & j + 3 + i4).Value = sheet1.Range("B23").Value '检测日期

            If szdh > sheet1.Range("B10").Value Then
                Randomize()
                sheet3.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)   '后视
            Else
                Randomize()
                sheet3.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)
            End If
            sheet3.Range(SJGL & j + SJWZH + i4).Value = szdh  '水准点高程
            sheet3.Range("A" & j + SJWZH + i4).Value = szdm  '水准点名称
            sheet3.Range(SXGL & j + SJWZH + i4).Value = sheet3.Range(SJGL & j + SJWZH + i4).Value + sheet3.Range(HSL & j + SJWZH + i4).Value      '视线高
            a = 0
            s(a) = Val(sheet3.Range(SXGL & j + SJWZH + i4).Value)
            z(a) = Val(szdz) '桩号值
            '计算
            i2 = 1

            i3 = 0
            i4 = i4 + 1
            ''''
            For t = i1 To i1 + i2 - 1

                '设转点
                ZH0 = Val(ExApp.WorksheetFunction.Substitute(sheet1.Range("J3").Value, "*", ""))

                If Val(ZH0) - z(a) > 200 Then '5
                    sheet3.Range("a" & j + SJWZH + i4).Value = "ZD" & (a + 1)   '桩号
                    If s(a) - Val(sheet1.Range("B10").Value) > 4.5 Then
                        Randomize()
                        sheet3.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)     '前视
                        Randomize()
                        sheet3.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)    '后视
                    Else
                        If s(a) - Val(sheet1.Range("B10").Value) < 0.3 Then
                            Randomize()
                            sheet3.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)   '前视
                            Randomize()
                            sheet3.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)    '后视
                        Else
                            Randomize()
                            sheet3.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)   '前视
                            Randomize()
                            sheet3.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)    '后视
                        End If
                    End If
                    sheet3.Range(SCGL & j + SJWZH + i4).Value = s(a) - Val(sheet3.Range(QSL & j + SJWZH + i4).Value)       '实测高程
                    a = a + 1 '交点数加1
                    s(a) = Val(sheet3.Range(SCGL & j + SJWZH + i4).Value) + Val(sheet3.Range(HSL & j + SJWZH + i4).Value)       '视线高
                    sheet3.Range(SXGL & j + SJWZH + i4).Value = s(a)
                    sheet3.Range(SCGL & j + SJWZH + i4).Value = ""   ''''''''实测高
                    z(a) = z(a - 1) + 200
                    i4 = i4 + 1
                    t = t - 1


                Else '距离设转点
                    If Val(ZH0) - z(a) < -200 Then '4
                        sheet3.Range("a" & j + SJWZH + i4).Value = "ZD" & (a + 1)  '桩号
                        If s(a) - Val(sheet1.Range("B10").Value) > 4.5 Then
                            Randomize()
                            sheet3.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)     '前视
                            Randomize()
                            sheet3.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)    '后视
                        Else
                            If s(a) - Val(sheet1.Range("B10").Value) < 0.3 Then
                                Randomize()
                                sheet3.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)   '前视
                                Randomize()
                                sheet3.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)    '后视
                            Else
                                Randomize()
                                sheet3.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)   '前视
                                Randomize()
                                sheet3.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)    '后视
                            End If
                        End If
                        sheet3.Range(SCGL & j + SJWZH + i4).Value = s(a) - Val(sheet3.Range(QSL & j + SJWZH + i4).Value)       '实测高程
                        a = a + 1 '交点数加1
                        s(a) = Val(sheet3.Range(SCGL & j + SJWZH + i4).Value) + Val(sheet3.Range(HSL & j + SJWZH + i4).Value)       '视线高
                        sheet3.Range(SXGL & j + SJWZH + i4).Value = s(a)
                        sheet3.Range(SCGL & j + SJWZH + i4).Value = ""    ''''''''实测高
                        z(a) = z(a - 1) - 200
                        i4 = i4 + 1
                        t = t - 1


                    Else '设转点
                        If s(a) - Val(sheet1.Range("B10").Value) > 4.5 Then '3
                            sheet3.Range("a" & j + SJWZH + i4).Value = "ZD" & (a + 1)     '桩号
                            Randomize()
                            sheet3.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)     '前视
                            sheet3.Range(SCGL & j + SJWZH + i4).Value = s(a) - Val(sheet3.Range(QSL & j + SJWZH + i4).Value)       '实测高程
                            Randomize()
                            sheet3.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)     '后视
                            a = a + 1 '交点数加1
                            s(a) = Val(sheet3.Range(SCGL & j + SJWZH + i4).Value) + Val(sheet3.Range(HSL & j + SJWZH + i4).Value)       '视线高
                            sheet3.Range(SXGL & j + SJWZH + i4).Value = s(a)
                            sheet3.Range(SCGL & j + SJWZH + i4).Value = ""    ''''''''实测高
                            z(a) = z(a - 1) '   桩号值
                            i4 = i4 + 1
                            t = t - 1

                        Else '设转点
                            If s(a) - Val(sheet1.Range("B10").Value) < 0.3 Then  '2
                                sheet3.Range("a" & j + SJWZH + i4).Value = "ZD" & (a + 1)   '桩号
                                Randomize()
                                sheet3.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)    '前视
                                sheet3.Range(SCGL & j + SJWZH + i4).Value = s(a) - sheet3.Range(QSL & j + SJWZH + i4).Value      '实测高程
                                Randomize()
                                sheet3.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)    '后视
                                a = a + 1 '交点数加1
                                s(a) = Val(sheet3.Range(SCGL & j + SJWZH + i4).Value) + Val(sheet3.Range(HSL & j + SJWZH + i4).Value)      '视线高
                                sheet3.Range(SXGL & j + SJWZH + i4).Value = s(a)
                                sheet3.Range(SCGL & j + SJWZH + i4).Value = ""    ''''''''实测高
                                z(a) = z(a - 1)  '     桩号值
                                i4 = i4 + 1
                                t = t - 1

                            Else '非转点求值
                                If t <= i1 + i2 Then '1
                                    'sheet3.Range(SJGL & j + SJWZH + i4).value = Val(sheet1.Range("B4").value)     '设计高程
                                    sheet3.Range(SCGL & j + SJWZH + i4).Value = sheet1.Range（"B10"）.Value   '实测高程
                                    'sheet3.Range(PCL & j + SJWZH + i4).value = (sheet3.Range(SCGL & j + SJWZH + i4).value - sheet3.Range(SJGL & j + SJWZH + i4).value) * 100   '偏差
                                    sheet3.Range(ZSL & j + SJWZH + i4).Value = Math.Round(s(a) - Val(sheet1.Range（"B10"）.Value), 3)     '中视
                                    sheet3.Range(BZL & j + SJWZH + i4).Value = "测绳" & Math.Round(sheet1.Range("B10").Value - sheet1.Range("B11").Value, 3) & "m"  '备注
                                    sheet3.Range("a" & j + SJWZH + i4).Value = "护筒顶"
                                    i4 = i4 + 1
                                End If '1
                            End If '2
                        End If '3
                    End If '4
                End If '5


                '新增表格
                If i4 > SJH Then
                    j = j + BGH  '表格行数''
                    sheet3.Range("a" & j).PasteSpecial()
                    i4 = 1
                    '''''''''''''''''''''''改改改改改改改改
                    sheet3.Range("e" & j + 3 + i4).Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & sheet1.Range("F6").Value  '工程部位
                    sheet3.Range("i" & j + 3 + i4).Value = sheet1.Range("B23").Value '检测日期
                    sheet3.Range("E" & j + i4 + 1).Value = "续上页" '编号
                    sheet3.Range("E" & j + i4 + 1).Interior.Color = 255
                    sheet3.Range("E" & j + i4 + 1).Font.Color = -16711681
                    ''''''''''''''''''''''''''''''''''
                End If
            Next t


            ''''闭合'''''''''''''''''''''''''''''''''''''''''''
            For t = i1 + i2 - 1 To i1 + i2
                '设转点

                If Val(szdz) - z(a) > 200 Then '设转点
                    sheet3.Range("a" & j + SJWZH + i4).Value = "ZD" & (a + 1) '桩号
                    If s(a) < Val(szdh) Then
                        Randomize()
                        sheet3.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)     '前视
                        Randomize()
                        sheet3.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)    '后视
                    Else
                        Randomize()
                        sheet3.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)    '前视
                        Randomize()
                        sheet3.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)     '后视
                    End If
                    sheet3.Range(SCGL & j + SJWZH + i4).Value = s(a) - Val(sheet3.Range(QSL & j + SJWZH + i4).Value)      '实测高程
                    a = a + 1 '交点数加1
                    s(a) = Val(sheet3.Range(SCGL & j + SJWZH + i4).Value) + Val(sheet3.Range(HSL & j + SJWZH + i4).Value)    '视线高
                    sheet3.Range(SXGL & j + SJWZH + i4).Value = s(a)
                    sheet3.Range(SCGL & j + SJWZH + i4).Value = ""    ''''''''实测高
                    z(a) = z(a - 1) + 200
                    i4 = i4 + 1
                    t = t - 1
                Else '设转点
                    If Val(szdz) - z(a) < -200 Then
                        sheet3.Range("a" & j + SJWZH + i4).Value = "ZD" & (a + 1)   '桩号
                        If s(a) < Val(szdh) Then
                            Randomize()
                            sheet3.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)     '前视
                            Randomize()
                            sheet3.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)    '后视
                        Else
                            Randomize()
                            sheet3.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)    '前视
                            Randomize()
                            sheet3.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)     '后视
                        End If
                        sheet3.Range(SCGL & j + SJWZH + i4).Value = s(a) - Val(sheet3.Range(QSL & j + SJWZH + i4).Value)       '实测高程
                        a = a + 1 '交点数加1
                        s(a) = Val(sheet3.Range(SCGL & j + SJWZH + i4).Value) + Val(sheet3.Range(HSL & j + SJWZH + i4).Value)     '视线高
                        sheet3.Range(SXGL & j + SJWZH + i4).Value = s(a)
                        sheet3.Range(SCGL & j + SJWZH + i4).Value = ""    ''''''''实测高
                        i4 = i4 + 1
                        t = t - 1
                        z(a) = z(a - 1) - 200
                    Else '设转点
                        If s(a) - Val(szdh) > 4.6 Then
                            sheet3.Range("a" & j + SJWZH + i4).Value = "ZD" & (a + 1)     '桩号
                            Randomize()
                            sheet3.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)     '前视
                            sheet3.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)    '后视
                            sheet3.Range(SCGL & j + SJWZH + i4).Value = s(a) - Val(sheet3.Range(QSL & j + SJWZH + i4).Value)       '实测高程
                            a = a + 1 '交点数加1
                            s(a) = Val(sheet3.Range(SCGL & j + SJWZH + i4).Value) + Val(sheet3.Range(HSL & j + SJWZH + i4).Value)       '视线高
                            sheet3.Range(SXGL & j + SJWZH + i4).Value = s(a)
                            sheet3.Range(SCGL & j + SJWZH + i4).Value = ""    ''''''''实测高
                            z(a) = z(a - 1)  '   桩号值
                            i4 = i4 + 1
                            t = t - 1
                        Else '设转点
                            If s(a) - Val(szdh) < 0.4 Then
                                sheet3.Range("a" & j + SJWZH + i4).Value = "ZD" & (a + 1)   '桩号
                                Randomize()
                                sheet3.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)     '前视
                                sheet3.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)   '后视
                                sheet3.Range(SCGL & j + SJWZH + i4).Value = s(a) - Val(sheet3.Range(QSL & j + SJWZH + i4).Value)       '实测高程
                                a = a + 1 '交点数加1
                                s(a) = Val(sheet3.Range(SCGL & j + SJWZH + i4).Value) + Val(sheet3.Range(HSL & j + SJWZH + i4).Value)     '视线高
                                sheet3.Range(SXGL & j + SJWZH + i4).Value = s(a)
                                sheet3.Range(SCGL & j + SJWZH + i4).Value = ""    ''''''''实测高
                                z(a) = z(a - 1)  ' 桩号值
                                i4 = i4 + 1
                                t = t - 1
                            Else
                                Randomize()
                                'sheet3.Range(PCL & j + SJWZH + i4).value = Int((p_z - p_f + 1) * Rnd() + p_f)   '偏差
                                sheet3.Range(PCL & j + SJWZH + i4).Value = 0   '偏差
                                sheet3.Range(SJGL & j + SJWZH + i4).Value = szdh    '设计高程
                                sheet3.Range(SCGL & j + SJWZH + i4).Value = Math.Round(Val(sheet3.Range(SJGL & j + SJWZH + i4).Value) + Val(sheet3.Range(PCL & j + SJWZH + i4).Value) / 1000, 3)       '实测高程
                                sheet3.Range(QSL & j + SJWZH + i4).Value = Math.Round(s(a) - Val(sheet3.Range(SCGL & j + SJWZH + i4).Value), 3)      '前视
                                sheet3.Range("a" & j + SJWZH + i4).Value = szdm
                                sheet3.Range("a" & j + SJWZH + i4 + 1).Value = "孔底"
                                sheet3.Range(SCGL & j + SJWZH + i4 + 1).Value = sheet1.Range("B11").Value
                                sheet3.Range(SJGL & j + SJWZH + i4 + 1).Value = sheet1.Range("B8").Value
                                sheet3.Range(PCL & j + SJWZH + i4 + 1).Value = （sheet3.Range(SCGL & j + SJWZH + i4 + 1).Value - sheet3.Range(SJGL & j + SJWZH + i4 + 1).Value） * 1000

                            End If
                        End If
                    End If
                End If

                '新增表格
                If i4 > SJH Then
                    j = j + BGH '表格行数
                    sheet3.Range("a" & j).PasteSpecial()
                    i4 = 1
                    '‘’‘’‘’‘’‘改改改改改改改改
                    sheet3.Range("e" & j + 3 + i4).Value = sheet1.Range("D6").Value & sheet1.Range("F6").Value '工程部位
                    sheet3.Range("i" & j + 3 + i4).Value = sheet1.Range("B23").Value '检测日期
                    sheet3.Range("E" & j + i4 + 1).Value = "续上页" '编号
                    sheet3.Range("E" & j + i4 + 1).Interior.Color = 255
                    sheet3.Range("E" & j + i4 + 1).Font.Color = -16711681
                End If
            Next t
            i1 = i1 + 1


            '非转点求值
            'Dim p_z As Single, p_f As Single
            ''p_z = 2 * Math.Sqrt(a + 1)
            ''p_f = -2 * Math.Sqrt(a + 1)
            ''If ExApp.Sheets("数据库").range("O2").value = "二等" Then
            ''    p_z = 0
            ''    p_f = 0
            ''Else
            ''If ExApp.Sheets("数据库").range("O2").value = "三等" Then
            'p_z = 2
            'p_f = -2
            ''Else
            ''If ExApp.Sheets("数据库").range("O2").value = "四等" Then
            ''                p_z = 3
            ''                p_f = -3
            ''            End If
            ''End If

            ''End If

            '备注列
            '        sheet3.Range("B" & j + 28).value = "闭合或附合差fh=" & sheet3.Range(PCL & j + SJWZH + i4).value
            'sheet3.Range("B" & j + 30).value = "|fh|<|f允|，符合水准测量技术规定。"
            'sheet3.Range("B" & j + 29).value = "容许误差f允=6√n=6√" & a + 1 & "=" & Math.Round(6 * Math.Sqrt(a + 1), 0)

        Catch Exclerror As Exception   '错误时弹出提示
            MsgBox(Exclerror.Message)
            TorF = False
            Exit Sub
        End Try
    End Sub

    Sub 桩顶高程（）
        Dim t, a, i1, j, i2, i3, i4 As Integer, z(238) As Object, s(238) As Object
        Dim HSL, ZSL, QSL, SXGL, SCGL, BZL, SJH, SJGL, PCL, BGH, SJWZH, sumz, sumh, ZH0, min1, min2, szdz, szdh, szdm, Ming, Maxg, TCmind, TCminx, TCmaxd, TCmaxx

        Exbook = ExApp.ActiveWorkbook
        sheet0 = Exbook.Worksheets("数据库") '数据库
        sheet1 = Exbook.Worksheets("参数表") '参数表
        sheet2 = Exbook.Worksheets("导线成果表") '导线成果表
        sheet3 = Exbook.Worksheets("水下灌注记录") '水下灌注记录
        sheet4 = Exbook.Worksheets("桩顶水准") '桩顶水准

        Try
            HSL = "b" '后视列
            ZSL = "c" '中视列
            QSL = "d" '前视列
            SXGL = "e" '视线高列
            SCGL = "f" '实测高列
            SJGL = "g" '设计高列
            PCL = "h" '偏差列
            BZL = "i" '备注列
            SJH = 23  '数据行数，表格计算数据的行数
            BGH = 32 '表格总行数
            SJWZH = 9 '数据位置开始行，表格数据从哪行开始
            TCmind = 1000 '塔尺小读数大值
            TCminx = 500 '塔尺小读数小值
            TCmaxd = 4600 '塔尺大读数大值
            TCmaxx = 3200 '塔尺大读数小值
            j = 1
            i1 = 3      '数据表的开始行
            j = 1
            i1 = 3      '数据表的开始行
            SJWZH = SJWZH - 2
            Ming = -10
            Maxg = 10
            i2 = 3
            min1 = 1000  '允许桩号偏差范围内的控制点
            min2 = 100   '允许高程偏差范围内的控制点
            szdz = 0
            szdh = 0
            szdm = 0
            j = j + BGH '表格行数''

            sheet4.Select()
            CType(sheet4.Rows("33:99999"), Range).Delete()  '删除表格''
            CType(sheet4.Rows("1:32"), Range).Copy()            '复制表格''
            sheet4.Range("a" & j).PasteSpecial()

            '桩顶水准点桩号
            sumz = Val(ExApp.WorksheetFunction.Substitute(sheet1.Range("J3").Value, "*", ""))
            sumh = Val(sheet1.Range("B7").Value)

            For i2 = 3 To sheet2.Range("B1048576").End(XlDirection.xlUp).Row + 1

                If Math.Abs(sumz - sheet2.Range("F" & i2).Value) < min1 And Math.Abs(sumh - sheet2.Range("E" & i2).Value) < min2 And sheet2.Range("E" & i2).Value <> Nothing Then
                    'min1 = Math.Abs(sumz - sheet2.Range("F" & i2).value)   '就近取点、隧道时启用
                    min2 = Math.Abs(sumh - sheet2.Range("E" & i2).Value)
                    szdz = sheet2.Range("F" & i2).Value '水准点桩号
                    szdh = sheet2.Range("E" & i2).Value '水准点高程
                    szdm = sheet2.Range("B" & i2).Value '水准点名称

                End If
            Next i2

            '开始''''''''''''''''''''''''''''
            '开始''''''''''''''''''''''''''''
            '开始''''''''''''''''''''''''''''
            i4 = 1
            sheet4.Range("E" & j + 3 + i4).Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & sheet1.Range("F6").Value & "浇筑后桩顶" '工程部位
            sheet4.Range("I" & j + 3 + i4).Value = sheet1.Range("B23").Value '检测日期

            If szdh > sheet1.Range("B7").Value Then
                Randomize()
                sheet4.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)   '后视
            Else
                Randomize()
                sheet4.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)
            End If
            sheet4.Range(SJGL & j + SJWZH + i4).Value = szdh  '水准点高程
            sheet4.Range("A" & j + SJWZH + i4).Value = szdm  '水准点名称
            sheet4.Range(SXGL & j + SJWZH + i4).Value = sheet4.Range(SJGL & j + SJWZH + i4).Value + sheet4.Range(HSL & j + SJWZH + i4).Value      '视线高
            a = 0
            s(a) = Val(sheet4.Range(SXGL & j + SJWZH + i4).Value)
            z(a) = Val(szdz) '桩号值
            '计算
            i2 = 1

            i3 = 0
            i4 = i4 + 1
            ''''
            For t = i1 To i1 + i2 - 1

                '设转点
                ZH0 = Val(ExApp.WorksheetFunction.Substitute(sheet1.Range("J3").Value, "*", ""))

                If Val(ZH0) - z(a) > 200 Then '5
                    sheet4.Range("a" & j + SJWZH + i4).Value = "ZD" & (a + 1)   '桩号
                    If s(a) - Val(sheet1.Range("B7").Value) > 4.5 Then
                        Randomize()
                        sheet4.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)     '前视
                        Randomize()
                        sheet4.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)    '后视
                    Else
                        If s(a) - Val(sheet1.Range("B7").Value) < 0.3 Then
                            Randomize()
                            sheet4.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)   '前视
                            Randomize()
                            sheet4.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)    '后视
                        Else
                            Randomize()
                            sheet4.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)   '前视
                            Randomize()
                            sheet4.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)    '后视
                        End If
                    End If
                    sheet4.Range(SCGL & j + SJWZH + i4).Value = s(a) - Val(sheet4.Range(QSL & j + SJWZH + i4).Value)       '实测高程
                    a = a + 1 '交点数加1
                    s(a) = Val(sheet4.Range(SCGL & j + SJWZH + i4).Value) + Val(sheet4.Range(HSL & j + SJWZH + i4).Value)       '视线高
                    sheet4.Range(SXGL & j + SJWZH + i4).Value = s(a)
                    sheet4.Range(SCGL & j + SJWZH + i4).Value = ""   ''''''''实测高
                    z(a) = z(a - 1) + 200
                    i4 = i4 + 1
                    t = t - 1


                Else '距离设转点
                    If Val(ZH0) - z(a) < -200 Then '4
                        sheet4.Range("a" & j + SJWZH + i4).Value = "ZD" & (a + 1)  '桩号
                        If s(a) - Val(sheet1.Range("B7").Value) > 4.5 Then
                            Randomize()
                            sheet4.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)     '前视
                            Randomize()
                            sheet4.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)    '后视
                        Else
                            If s(a) - Val(sheet1.Range("B7").Value) < 0.3 Then
                                Randomize()
                                sheet4.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)   '前视
                                Randomize()
                                sheet4.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)    '后视
                            Else
                                Randomize()
                                sheet4.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)   '前视
                                Randomize()
                                sheet4.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)    '后视
                            End If
                        End If
                        sheet4.Range(SCGL & j + SJWZH + i4).Value = s(a) - Val(sheet4.Range(QSL & j + SJWZH + i4).Value)       '实测高程
                        a = a + 1 '交点数加1
                        s(a) = Val(sheet4.Range(SCGL & j + SJWZH + i4).Value) + Val(sheet4.Range(HSL & j + SJWZH + i4).Value)       '视线高
                        sheet4.Range(SXGL & j + SJWZH + i4).Value = s(a)
                        sheet4.Range(SCGL & j + SJWZH + i4).Value = ""    ''''''''实测高
                        z(a) = z(a - 1) - 200
                        i4 = i4 + 1
                        t = t - 1


                    Else '设转点
                        If s(a) - Val(sheet1.Range("B7").Value) > 4.5 Then '3
                            sheet4.Range("a" & j + SJWZH + i4).Value = "ZD" & (a + 1)     '桩号
                            Randomize()
                            sheet4.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)     '前视
                            sheet4.Range(SCGL & j + SJWZH + i4).Value = s(a) - Val(sheet4.Range(QSL & j + SJWZH + i4).Value)       '实测高程
                            Randomize()
                            sheet4.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)     '后视
                            a = a + 1 '交点数加1
                            s(a) = Val(sheet4.Range(SCGL & j + SJWZH + i4).Value) + Val(sheet4.Range(HSL & j + SJWZH + i4).Value)       '视线高
                            sheet4.Range(SXGL & j + SJWZH + i4).Value = s(a)
                            sheet4.Range(SCGL & j + SJWZH + i4).Value = ""    ''''''''实测高
                            z(a) = z(a - 1) '   桩号值
                            i4 = i4 + 1
                            t = t - 1

                        Else '设转点
                            If s(a) - Val(sheet1.Range("B7").Value) < 0.3 Then  '2
                                sheet4.Range("a" & j + SJWZH + i4).Value = "ZD" & (a + 1)   '桩号
                                Randomize()
                                sheet4.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)    '前视
                                sheet4.Range(SCGL & j + SJWZH + i4).Value = s(a) - sheet4.Range(QSL & j + SJWZH + i4).Value      '实测高程
                                Randomize()
                                sheet4.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)    '后视
                                a = a + 1 '交点数加1
                                s(a) = Val(sheet4.Range(SCGL & j + SJWZH + i4).Value) + Val(sheet4.Range(HSL & j + SJWZH + i4).Value)      '视线高
                                sheet4.Range(SXGL & j + SJWZH + i4).Value = s(a)
                                sheet4.Range(SCGL & j + SJWZH + i4).Value = ""    ''''''''实测高
                                z(a) = z(a - 1)  '     桩号值
                                i4 = i4 + 1
                                t = t - 1

                            Else '非转点求值
                                If t <= i1 + i2 Then '1
                                    'sheet4.Range(SJGL & j + SJWZH + i4).value = Val(sheet1.Range("B7").value)     '设计高程
                                    sheet4.Range(SCGL & j + SJWZH + i4).Value = sheet3.Range（"P9"）.Value   '实测高程
                                    'sheet4.Range(PCL & j + SJWZH + i4).value = Math.Round((sheet4.Range(SCGL & j + SJWZH + i4).value - sheet4.Range(SJGL & j + SJWZH + i4).value) * 1000, 0)  '偏差
                                    sheet4.Range(ZSL & j + SJWZH + i4).Value = Math.Round(s(a) - Val(sheet1.Range（"B10"）.Value), 3)     '中视
                                    'sheet4.Range(BZL & j + SJWZH + i4).value = "测绳" & Math.Round(sheet1.Range("B7").value - sheet1.Range("B8").value, 3) & "米"  '备注
                                    sheet4.Range("a" & j + SJWZH + i4).Value = "桩顶"
                                    i4 = i4 + 1
                                End If '1
                            End If '2
                        End If '3
                    End If '4
                End If '5


                '新增表格
                If i4 > SJH Then
                    j = j + BGH  '表格行数''
                    sheet4.Range("a" & j).PasteSpecial()
                    i4 = 1
                    '''''''''''''''''''''''改改改改改改改改
                    sheet4.Range("e" & j + 3 + i4).Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & sheet1.Range("F6").Value & "浇筑后桩顶"  '工程部位
                    sheet4.Range("i" & j + 3 + i4).Value = sheet1.Range("B23").Value '检测日期
                    sheet4.Range("E" & j + i4 + 1).Value = "续上页" '编号
                    sheet4.Range("E" & j + i4 + 1).Interior.Color = 255
                    sheet4.Range("E" & j + i4 + 1).Font.Color = -16711681
                    ''''''''''''''''''''''''''''''''''
                End If

            Next t


            ''''闭合'''''''''''''''''''''''''''''''''''''''''''
            For t = i1 + i2 - 1 To i1 + i2
                '设转点

                If Val(szdz) - z(a) > 200 Then '设转点
                    sheet4.Range("a" & j + SJWZH + i4).Value = "ZD" & (a + 1) '桩号
                    If s(a) < Val(szdh) Then
                        Randomize()
                        sheet4.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)     '前视
                        Randomize()
                        sheet4.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)    '后视
                    Else
                        Randomize()
                        sheet4.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)    '前视
                        Randomize()
                        sheet4.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)     '后视
                    End If
                    sheet4.Range(SCGL & j + SJWZH + i4).Value = s(a) - Val(sheet4.Range(QSL & j + SJWZH + i4).Value)      '实测高程
                    a = a + 1 '交点数加1
                    s(a) = Val(sheet4.Range(SCGL & j + SJWZH + i4).Value) + Val(sheet4.Range(HSL & j + SJWZH + i4).Value)    '视线高
                    sheet4.Range(SXGL & j + SJWZH + i4).Value = s(a)
                    sheet4.Range(SCGL & j + SJWZH + i4).Value = ""    ''''''''实测高
                    z(a) = z(a - 1) + 200
                    i4 = i4 + 1
                    t = t - 1
                Else '设转点
                    If Val(szdz) - z(a) < -200 Then
                        sheet4.Range("a" & j + SJWZH + i4).Value = "ZD" & (a + 1)   '桩号
                        If s(a) < Val(szdh) Then
                            Randomize()
                            sheet4.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)     '前视
                            Randomize()
                            sheet4.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)    '后视
                        Else
                            Randomize()
                            sheet4.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)    '前视
                            Randomize()
                            sheet4.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)     '后视
                        End If
                        sheet4.Range(SCGL & j + SJWZH + i4).Value = s(a) - Val(sheet4.Range(QSL & j + SJWZH + i4).Value)       '实测高程
                        a = a + 1 '交点数加1
                        s(a) = Val(sheet4.Range(SCGL & j + SJWZH + i4).Value) + Val(sheet4.Range(HSL & j + SJWZH + i4).Value)     '视线高
                        sheet4.Range(SXGL & j + SJWZH + i4).Value = s(a)
                        sheet4.Range(SCGL & j + SJWZH + i4).Value = ""    ''''''''实测高
                        i4 = i4 + 1
                        t = t - 1
                        z(a) = z(a - 1) - 200
                    Else '设转点
                        If s(a) - Val(szdh) > 4.6 Then
                            sheet4.Range("a" & j + SJWZH + i4).Value = "ZD" & (a + 1)     '桩号
                            Randomize()
                            sheet4.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)     '前视
                            sheet4.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)    '后视
                            sheet4.Range(SCGL & j + SJWZH + i4).Value = s(a) - Val(sheet4.Range(QSL & j + SJWZH + i4).Value)       '实测高程
                            a = a + 1 '交点数加1
                            s(a) = Val(sheet4.Range(SCGL & j + SJWZH + i4).Value) + Val(sheet4.Range(HSL & j + SJWZH + i4).Value)       '视线高
                            sheet4.Range(SXGL & j + SJWZH + i4).Value = s(a)
                            sheet4.Range(SCGL & j + SJWZH + i4).Value = ""    ''''''''实测高
                            z(a) = z(a - 1)  '   桩号值
                            i4 = i4 + 1
                            t = t - 1
                        Else '设转点
                            If s(a) - Val(szdh) < 0.4 Then
                                sheet4.Range("a" & j + SJWZH + i4).Value = "ZD" & (a + 1)   '桩号
                                Randomize()
                                sheet4.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)     '前视
                                sheet4.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)   '后视
                                sheet4.Range(SCGL & j + SJWZH + i4).Value = s(a) - Val(sheet4.Range(QSL & j + SJWZH + i4).Value)       '实测高程
                                a = a + 1 '交点数加1
                                s(a) = Val(sheet4.Range(SCGL & j + SJWZH + i4).Value) + Val(sheet4.Range(HSL & j + SJWZH + i4).Value)     '视线高
                                sheet4.Range(SXGL & j + SJWZH + i4).Value = s(a)
                                sheet4.Range(SCGL & j + SJWZH + i4).Value = ""    ''''''''实测高
                                z(a) = z(a - 1)  ' 桩号值
                                i4 = i4 + 1
                                t = t - 1
                            Else

                                Randomize()
                                'sheet4.Range(PCL & j + SJWZH + i4).value = Int((p_z - p_f + 1) * Rnd() + p_f)   '偏差
                                sheet4.Range(PCL & j + SJWZH + i4).Value = 0  '偏差
                                sheet4.Range(SJGL & j + SJWZH + i4).Value = szdh    '设计高程
                                sheet4.Range(SCGL & j + SJWZH + i4).Value = Math.Round(Val(sheet4.Range(SJGL & j + SJWZH + i4).Value) + Val(sheet4.Range(PCL & j + SJWZH + i4).Value) / 1000, 3)       '实测高程
                                sheet4.Range(QSL & j + SJWZH + i4).Value = Math.Round(s(a) - Val(sheet4.Range(SCGL & j + SJWZH + i4).Value), 3)      '前视
                                sheet4.Range("a" & j + SJWZH + i4).Value = szdm
                            End If
                        End If
                    End If
                End If


                '新增表格
                If i4 > SJH Then
                    j = j + BGH '表格行数
                    sheet4.Range("a" & j).PasteSpecial()
                    i4 = 1
                    '‘’‘’‘’‘’‘改改改改改改改改
                    sheet4.Range("e" & j + 3 + i4).Value = sheet1.Range("D6").Value & sheet1.Range("F6").Value & "浇筑后桩顶"  '工程部位
                    sheet4.Range("i" & j + 3 + i4).Value = sheet1.Range("B23").Value '检测日期
                    sheet4.Range("E" & j + i4 + 1).Value = "续上页" '编号
                    sheet4.Range("E" & j + i4 + 1).Interior.Color = 255
                    sheet4.Range("E" & j + i4 + 1).Font.Color = -16711681

                End If

            Next t
            i1 = i1 + 1


            '非转点求值
            'Dim p_z As Single, p_f As Single
            ''p_z = 2 * Math.Sqrt(a + 1)
            ''p_f = -2 * Math.Sqrt(a + 1)
            ''If ExApp.Sheets("数据库").range("O2").value = "二等" Then
            ''    p_z = 0
            ''    p_f = 0
            ''Else
            ''If ExApp.Sheets("数据库").range("O2").value = "三等" Then
            'p_z = 2
            'p_f = -2
            ''Else
            ''If ExApp.Sheets("数据库").range("O2").value = "四等" Then
            ''                p_z = 3
            ''                p_f = -3
            ''            End If
            ''End If

            'End If




            '备注列
            '        sheet4.Range("B" & j + 28).value = "闭合或附合差fh=" & sheet4.Range(PCL & j + SJWZH + i4).value
            'sheet4.Range("B" & j + 30).value = "|fh|<|f允|，符合水准测量技术规定。"
            'sheet4.Range("B" & j + 29).value = "容许误差f允=6√n=6√" & a + 1 & "=" & Math.Round(6 * Math.Sqrt(a + 1), 0)

        Catch Exclerror As Exception   '错误时弹出提示
            MsgBox(Exclerror.Message)
            TorF = False
            Exit Sub
        End Try
    End Sub

    Sub 成孔平面()

        Dim pd, i1, i2, i3, j As Integer, z(238) As Object, s(238) As Object
        Dim min1, Sumz1， DXCGz, DXCGx, DXCGy, DXCGh, DXCGm, DXCGz1, DXCGx1, DXCGy1, DXCGh1, DXCGm1， JD, jdd, jdf, jdm, pdz， SJXL, SJYL, SCXL, SCYL, PCXL, PCYL, PWL, SJH, SJWZH， Minp, Maxp

        Exbook = ExApp.ActiveWorkbook
        sheet0 = Exbook.Worksheets("数据库") '数据库
        sheet1 = Exbook.Worksheets("参数表") '参数表
        sheet2 = Exbook.Worksheets("导线成果表") '导线成果表
        sheet3 = Exbook.Worksheets("钻孔桩检表") '钻孔桩检表
        sheet4 = Exbook.Worksheets("成孔平面") '成孔平面
        sheet4.Rows("12:21") = Nothing
        SJXL = "C"  '设计坐标X列
        SJYL = "E"  '设计坐标Y列
        SCXL = "G"  '实测坐标X列
        SCYL = "I"  '实测坐标Y列
        PCXL = "K"  '差　　值X列
        PCYL = "L"  '差　　值Y列
        PWL = "M"   '偏位 √(△X2+△Y2 列
        SJH = 10    '数据行数，表格计算数据的行数
        SJWZH = 12  '数据位置开始行，表格数据从哪行开始
        j = 1
        i1 = 3
        pd = 0
        SJWZH = SJWZH - 2
        Minp = 0
        Maxp = 15

        Try
            i3 = 1
            pd = pd + 1
            ''''''改，表头数据绝对行均减2
            sheet4.Range("G" & j + 3 + i3).Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & sheet1.Range("F6").Value & "成孔" '工程部位
            sheet4.Range("O" & j + 3 + i3).Value = sheet1.Range("B23").Value '时间
            sheet4.Range("O" & j + 5 + i3).Value = "P=" & ExApp.WorksheetFunction.RandBetween(715, 720) & "mmhg"   '气压
            sheet4.Range("O" & j + 6 + i3).Value = "T=" & ExApp.WorksheetFunction.RandBetween(25, 28) & "℃"   '温度
            sheet4.Range("O" & j + 7 + i3).Value = "I=" & Math.Round(1 + ExApp.WorksheetFunction.RandBetween(550, 680) * 0.001, 3) & "m"   '仪高
            ''''桩号
            Sumz1 = Val(ExApp.WorksheetFunction.Substitute(sheet1.Range("J3").Value, "*", ""))
            i2 = 3

            '站点
            min1 = 9999
            For i2 = 3 To sheet2.Range("B1048576").End(XlDirection.xlUp).Row + 1
                If Math.Abs(Sumz1 - sheet2.Range("F" & i2).Value) < min1 And sheet2.Range("C" & i2).Value <> Nothing Then
                    min1 = Math.Abs(Sumz1 - sheet2.Range("f" & i2).Value)
                    DXCGz = sheet2.Range("F" & i2).Value '导线点桩号
                    DXCGx = sheet2.Range("C" & i2).Value '导线点X
                    DXCGy = sheet2.Range("D" & i2).Value '导线点Y
                    DXCGh = sheet2.Range("E" & i2).Value '导线点h
                    DXCGm = sheet2.Range("B" & i2).Value '导线点
                End If
            Next i2
            '后视
            i2 = 3
            min1 = 9999
            For i2 = 3 To sheet2.Range("B1048576").End(XlDirection.xlUp).Row + 1
                If Math.Abs(Sumz1 - sheet2.Range("f" & i2).Value) < min1 And DXCGm <> sheet2.Range("b" & i2).Value And sheet2.Range("C" & i2).Value <> Nothing Then
                    min1 = Math.Abs(Sumz1 - sheet2.Range("f" & i2).Value)
                    DXCGz1 = sheet2.Range("f" & i2).Value '后视导线点桩号
                    DXCGx1 = sheet2.Range("C" & i2).Value '后导线点X
                    DXCGy1 = sheet2.Range("D" & i2).Value '后导线点Y
                    DXCGh1 = sheet2.Range("e" & i2).Value '后导线点h
                    DXCGm1 = sheet2.Range("B" & i2).Value '后导线点
                End If
            Next i2

            '方位角角度
            JD = ExApp.WorksheetFunction.Atan2((-DXCGx + DXCGx1), (-DXCGy + DXCGy1)) * 180 / ExApp.WorksheetFunction.Pi
            If JD < 0 Then
                JD = JD + 360
            End If
            jdd = Int(JD)
            jdf = Int((JD - Int(JD)) * 60)
            jdm = Int(((JD - Int(JD)) * 60 - Int((JD - Int(JD)) * 60)) * 60)
            '改测点后视点
            sheet4.Range("B" & j + 5 + i3).Value = DXCGm
            sheet4.Range("F" & j + 5 + i3).Value = DXCGx
            sheet4.Range("F" & j + 6 + i3).Value = DXCGy
            sheet4.Range("F" & j + 7 + i3).Value = DXCGh
            sheet4.Range("H" & j + 5 + i3).Value = DXCGm1
            sheet4.Range("K" & j + 5 + i3).Value = DXCGx1
            sheet4.Range("K" & j + 6 + i3).Value = DXCGy1
            sheet4.Range("K" & j + 7 + i3).Value = DXCGh1
            '距离
            sheet4.Range("M" & j + 5 + i3).Value = "S=" & Math.Round(Math.Sqrt((DXCGx - DXCGx1) ^ 2 + (DXCGy - DXCGy1) ^ 2), 3) & "m"
            sheet4.Range("M" & j + 6 + i3).Value = "ɑ=" & jdd & "°" & jdf & "′" & jdm & "″"


            '设计坐标
            sheet4.Range(SJXL & j + SJWZH + i3).Value = sheet1.Range("B27").Value
            sheet4.Range(SJYL & j + SJWZH + i3).Value = sheet1.Range("D27").Value
            '偏差
            If sheet1.Range("B19").Value = "排架桩" Then
                sheet4.Range(PWL & j + SJWZH + i3).Value = sheet3.Range("H14").Value
            Else
                sheet4.Range(PWL & j + SJWZH + i3).Value = sheet3.Range("H13").Value
            End If
            pdz = ExApp.WorksheetFunction.RandBetween(-1, 1)
            If pdz < 0 Then
                pdz = -1
            Else
                pdz = 1
            End If
            sheet4.Range(PCXL & j + SJWZH + i3).Value = pdz * Math.Round(Math.Sqrt(ExApp.WorksheetFunction.RandBetween(0, sheet4.Range(PWL & j + SJWZH + i3).Value * sheet4.Range(PWL & j + SJWZH + i3).Value / 2)), 0)
            pdz = ExApp.WorksheetFunction.RandBetween(-1, 1)
            If pdz < 0 Then
                pdz = -1
            Else
                pdz = 1
            End If
            sheet4.Range(PCYL & j + SJWZH + i3).Value = pdz * Math.Round(Math.Sqrt(sheet4.Range(PWL & j + SJWZH + i3).Value * sheet4.Range(PWL & j + SJWZH + i3).Value - sheet4.Range(PCXL & j + SJWZH + i3).Value * sheet4.Range(PCXL & j + SJWZH + i3).Value), 0)
            '实测
            sheet4.Range(SCXL & j + SJWZH + i3).Value = sheet4.Range(SJXL & j + SJWZH + i3).Value + sheet4.Range(PCXL & j + SJWZH + i3).Value / 1000
            sheet4.Range(SCYL & j + SJWZH + i3).Value = sheet4.Range(SJYL & j + SJWZH + i3).Value + sheet4.Range(PCYL & j + SJWZH + i3).Value / 1000
            '桩号
            sheet4.Range("a" & j + SJWZH + i3).Value = sheet1.Range("D6").Value & sheet1.Range("F6").Value & "孔位"

        Catch Exclerror As Exception   '错误时弹出提示
            MsgBox(Exclerror.Message)
            TorF = False
            Exit Sub
        End Try
    End Sub

    Sub 质检资料设计坐标()

        Dim i, n, ZH0 As Integer
        Dim sjxq, sjyq, sjxh, sjyh, sjxz, sjyz, sjxy, sjyy, sjxzz, sjyzz As Double
        i = 3
        Try
            Exbook = ExApp.ActiveWorkbook
            sheet0 = Exbook.Worksheets("数据库")
            sheet1 = Exbook.Worksheets("参数表")
            sheet2 = Exbook.Worksheets("导线成果表")
            sheet3 = Exbook.Worksheets("交点法")
            sheet4 = Exbook.Worksheets("线元法")


            ExApp.ScreenUpdating = False   '关闭屏幕刷新
            Do While sheet0.Range("C" & i).Value <> Nothing
                If TorF = False Then
                    Exit Sub
                End If

                ''''''''交点法，   线元法时n=1
                ZH0 = Val(ExApp.WorksheetFunction.Substitute(sheet0.Range("C" & i).Value, "*", ""))
                If sheet4.Range("J2").Value <> "是" Then
                    n = Pd_YSw(ZH0, sheet3.Range("O5 : O500").Value)   '对应桩号所在线元的位置
                Else
                    n = 1 ''''线元法
                End If
                If n = -1 Then
                    MsgBox("请在交点法表内输入数据")
                    TorF = False
                    Exit Sub
                Else
                    If sheet4.Range("J2").Value <> "是" Then
                        ''''''''交点法
                        sjxq = ZSZB_X0j(sheet0.Range("BA" & i).Value, sheet0.Range("BB" & i).Value, 90)  '前偏距坐标
                        sjyq = ZSZB_Y0j(sheet0.Range("BA" & i).Value, sheet0.Range("BB" & i).Value, 90)
                        sjxh = ZSZB_X0j(sheet0.Range("BE" & i).Value, sheet0.Range("BF" & i).Value, 90)  '后偏距坐标
                        sjyh = ZSZB_Y0j(sheet0.Range("BE" & i).Value, sheet0.Range("BF" & i).Value, 90)
                        sjxz = ZSZB_X0j(sheet0.Range("BI" & i).Value, sheet0.Range("BJ" & i).Value, 90)  '左偏距坐标
                        sjyz = ZSZB_Y0j(sheet0.Range("BI" & i).Value, sheet0.Range("BJ" & i).Value, 90)
                        sjxy = ZSZB_X0j(sheet0.Range("BM" & i).Value, sheet0.Range("BN" & i).Value, 90)  '右偏距坐标
                        sjyy = ZSZB_Y0j(sheet0.Range("BM" & i).Value, sheet0.Range("BN" & i).Value, 90)
                        sjxzz = ZSZB_X0j(sheet0.Range("BQ" & i).Value, sheet0.Range("BR" & i).Value, 90)  '中偏距坐标
                        sjyzz = ZSZB_Y0j(sheet0.Range("BQ" & i).Value, sheet0.Range("BR" & i).Value, 90)

                    Else
                        '''''''''''线元法
                        sjxq = XYF_X(sheet0.Range("BA" & i).Value, sheet0.Range("BB" & i).Value, 90)  '前偏距坐标
                        sjyq = XYF_Y(sheet0.Range("BA" & i).Value, sheet0.Range("BB" & i).Value, 90)
                        sjxh = XYF_X(sheet0.Range("BE" & i).Value, sheet0.Range("BF" & i).Value, 90)  '后偏距坐标
                        sjyh = XYF_Y(sheet0.Range("BE" & i).Value, sheet0.Range("BF" & i).Value, 90)
                        sjxz = XYF_X(sheet0.Range("BI" & i).Value, sheet0.Range("BJ" & i).Value, 90)  '左偏距坐标
                        sjyz = XYF_Y(sheet0.Range("BI" & i).Value, sheet0.Range("BJ" & i).Value, 90)
                        sjxy = XYF_X(sheet0.Range("BM" & i).Value, sheet0.Range("BN" & i).Value, 90)  '右偏距坐标
                        sjyy = XYF_Y(sheet0.Range("BM" & i).Value, sheet0.Range("BN" & i).Value, 90)
                        sjxzz = XYF_X(sheet0.Range("BQ" & i).Value, sheet0.Range("BR" & i).Value, 90)  '中偏距坐标
                        sjyzz = XYF_Y(sheet0.Range("BQ" & i).Value, sheet0.Range("BR" & i).Value, 90)
                    End If
                    '偏距坐标赋值
                    sheet0.Range("BC" & i).Value = Math.Round(sjxq, 3)
                    sheet0.Range("BD" & i).Value = Math.Round(sjyq, 3)
                    sheet0.Range("BG" & i).Value = Math.Round(sjxh, 3)
                    sheet0.Range("BH" & i).Value = Math.Round(sjyh, 3)
                    sheet0.Range("BK" & i).Value = Math.Round(sjxz, 3)
                    sheet0.Range("BL" & i).Value = Math.Round(sjyz, 3)
                    sheet0.Range("BO" & i).Value = Math.Round(sjxy, 3)
                    sheet0.Range("BP" & i).Value = Math.Round(sjyy, 3)
                    sheet0.Range("BS" & i).Value = Math.Round(sjxzz, 3)
                    sheet0.Range("BT" & i).Value = Math.Round(sjyzz, 3)
                End If
                i = i + 1
            Loop
            ExApp.ScreenUpdating = True   '开启屏幕刷新
        Catch Exclerror As Exception   '错误时弹出提示
            MsgBox(Exclerror.Message)
            TorF = False
            Exit Sub
        End Try
    End Sub

    Sub 水准测量记录表（）

        Dim i, a, ZS, pd, i1, j, i2, i3, i4 As Integer, z(238) As Object, s(238) As Object
        Dim mc2, HSL, ZSL, QSL, SXGL, SCGL, BZL, SJGL, PCL, SJH, BGH, SJWZH, sumz, sumh, ZH0, min1, min2, szdz, szdh, szdm, Ming, Maxg, TCmind, TCminx, TCmaxd, TCmaxx

        Try
            Exbook = ExApp.ActiveWorkbook
            sheet1 = Exbook.Worksheets("参数表") '参数表
            sheet2 = Exbook.Worksheets("导线成果表") '导线成果表
            sheet3 = Exbook.Worksheets("水准表")

            ExApp.ScreenUpdating = False '关闭代码运行屏显
            HSL = "b" '后视列
            ZSL = "c" '中视列
            QSL = "d" '前视列
            SXGL = "e" '视线高列
            SCGL = "f" '实测高列
            SJGL = "g" '设计高列
            PCL = "h" '偏差列
            BZL = "i" '备注列
            SJH = 23  '数据行数，表格计算数据的行数
            BGH = 32 '表格总行数
            SJWZH = 9 '数据位置开始行，表格数据从哪行开始
            TCmind = 1000 '塔尺小读数大值
            TCminx = 500 '塔尺小读数小值
            TCmaxd = 4600 '塔尺大读数大值
            TCmaxx = 3200 '塔尺大读数小值
            j = 1
            i1 = 3      '数据表的开始行
            pd = 0
            SJWZH = SJWZH - 2
            Ming = -10
            Maxg = 10

            sheet3.Select()
            CType(sheet3.Rows("33:99999"), Range).Delete()  '删除表格
            CType(sheet3.Rows("1:32"), Range).Copy()        '复制表格

            While sheet1.Range("K" & i1).Value <> Nothing    'while123
                If TorF = False Then
                    Exit Sub
                End If
                If sheet1.Range("L" & i1).Value <> Nothing Then
                    pd = pd + 1
                    j = j + BGH '表格行数''
                    sheet3.Range("a" & j).PasteSpecial() '粘贴表格
                    i2 = 1
                    '水准点桩号
                    sumz = Val(ExApp.WorksheetFunction.Substitute(sheet1.Range("I" & i1).Value, "*", ""))
                    sumh = Val(sheet1.Range("K" & i1).Value)
                    While sheet1.Range("K" & i1 + i2).Value <> Nothing And sheet1.Range("L" & i1 + i2).Value = Nothing
                        sumz = sumz + Val(ExApp.WorksheetFunction.Substitute(sheet1.Range("I" & i1 + i2).Value, "*", ""))
                        sumh = sumh + Val(sheet1.Range("K" & i1 + i2).Value)
                        i2 = i2 + 1
                    End While
                    sumz = sumz / i2 '桩号平均值
                    sumh = sumh / i2 '高程平均值
                    i2 = 3

                    min1 = 1000  '允许桩号偏差范围内的控制点
                    min2 = 100   '允许高程偏差范围内的控制点
                    szdz = 0
                    szdh = 0
                    szdm = 0

                    For i2 = 3 To sheet2.Range("B1048576").End(XlDirection.xlUp).Row + 1

                        If sheet1.Range("J1").Value = "进口" Then

                            If Math.Abs(sumz - sheet2.Range("F" & i2).Value) < min1 And Math.Abs(sumh - sheet2.Range("E" & i2).Value) < min2 And sumz - sheet2.Range("F" & i2).Value >= 0 And sheet2.Range("E" & i2).Value <> Nothing Then
                                'min1 = Math.Abs(sumz - sheet2.Range("F" & i2).value)  '就近取点、隧道时启用
                                min2 = Math.Abs(sumh - sheet2.Range("E" & i2).Value)
                                szdz = sheet2.Range("F" & i2).Value '水准点桩号
                                szdh = sheet2.Range("E" & i2).Value '水准点高程
                                szdm = sheet2.Range("B" & i2).Value '水准点名称
                            End If
                        Else
                            If sheet1.Range("J1").Value = "出口" Then
                                If Math.Abs(sumz - sheet2.Range("F" & i2).Value) < min1 And Math.Abs(sumh - sheet2.Range("E" & i2).Value) < min2 And sheet2.Range("F" & i2).Value - sumz >= 0 And sheet2.Range("E" & i2).Value <> Nothing Then
                                    'min1 = Math.Abs(sumz - sheet2.Range("F" & i2).value)   '就近取点、隧道时启用
                                    min2 = Math.Abs(sumh - sheet2.Range("E" & i2).Value)
                                    szdz = sheet2.Range("F" & i2).Value '水准点桩号
                                    szdh = sheet2.Range("E" & i2).Value '水准点高程
                                    szdm = sheet2.Range("B" & i2).Value '水准点名称
                                End If
                            Else
                                If Math.Abs(sumz - sheet2.Range("F" & i2).Value) < min1 And Math.Abs(sumh - sheet2.Range("E" & i2).Value) < min2 And sheet2.Range("E" & i2).Value <> Nothing Then
                                    min1 = Math.Abs(sumz - sheet2.Range("F" & i2).Value)
                                    min2 = Math.Abs(sumh - sheet2.Range("E" & i2).Value)
                                    szdz = sheet2.Range("F" & i2).Value '水准点桩号
                                    szdh = sheet2.Range("E" & i2).Value '水准点高程
                                    szdm = sheet2.Range("B" & i2).Value '水准点名称
                                End If
                            End If
                        End If
                    Next i2

                    '开始''''''''''''''''''''''''''''
                    '开始''''''''''''''''''''''''''''
                    '开始''''''''''''''''''''''''''''
                    i4 = 1
                    '' 改，表头数据绝对行均减2
                    sheet3.Range("e" & j + 3 + i4).Value = sheet1.Range("L3").Value   '工程部位
                    sheet3.Range("i" & j + 3 + i4).Value = sheet1.Range("P1").Value '检测日期
                    'Range("I" & j + 2 + i4) = pd '序号
                    ''''''''''''''''''''''''''''''''''''
                    If szdh > sheet1.Range("K" & i1).Value Then
                        Randomize()
                        sheet3.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)   '后视
                    Else
                        Randomize()
                        sheet3.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)
                    End If
                    sheet3.Range(SJGL & j + SJWZH + i4).Value = szdh  '水准测量记录表点高程
                    sheet3.Range("a" & j + SJWZH + i4).Value = szdm  '水准测量记录表点名称
                    sheet3.Range(SXGL & j + SJWZH + i4).Value = sheet3.Range(SJGL & j + SJWZH + i4).Value + sheet3.Range(HSL & j + SJWZH + i4).Value      '视线高
                    a = 0
                    s(a) = Val(sheet3.Range(SXGL & j + SJWZH + i4).Value)
                    z(a) = Val(szdz) '桩号值
                    '计算
                    i2 = 1
                    While sheet1.Range("L" & (i1 + i2)).Value = Nothing And sheet1.Range("K" & (i1 + i2)).Value <> Nothing
                        i2 = i2 + 1
                    End While
                    i3 = 0
                    i4 = i4 + 1
                    ''''
                    For i = i1 To i1 + i2 - 1

                        '设转点
                        ZH0 = Val(ExApp.WorksheetFunction.Substitute(sheet1.Range("I" & i).Value, "*", ""))

                        If Val(ZH0) - z(a) > 200 Then '5
                            sheet3.Range("a" & j + SJWZH + i4).Value = "ZD" & (a + 1)   '桩号
                            If (s(a) - Val(sheet1.Range("K" & i).Value) > -0.3 And sheet1.Range("Q" & i).Value = "是") Or (s(a) - Val(sheet1.Range("K" & i).Value) > 4.5 And sheet1.Range("Q" & i).Value <> "是") Then
                                Randomize()
                                sheet3.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)     '前视
                                Randomize()
                                sheet3.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)    '后视
                            Else
                                If (s(a) - Val(sheet1.Range("K" & i).Value) < -4.5 And sheet1.Range("Q" & i).Value = "是") Or (s(a) - Val(sheet1.Range("K" & i).Value) < 0.3 And sheet1.Range("Q" & i).Value <> "是") Then
                                    Randomize()
                                    sheet3.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)   '前视
                                    Randomize()
                                    sheet3.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)    '后视
                                Else
                                    Randomize()
                                    sheet3.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)   '前视
                                    Randomize()
                                    sheet3.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)    '后视
                                End If
                            End If
                            sheet3.Range(SCGL & j + SJWZH + i4).Value = s(a) - Val(sheet3.Range(QSL & j + SJWZH + i4).Value)       '实测高程
                            a = a + 1 '交点数加1
                            s(a) = Val(sheet3.Range(SCGL & j + SJWZH + i4).Value) + Val(sheet3.Range(HSL & j + SJWZH + i4).Value)       '视线高
                            sheet3.Range(SXGL & j + SJWZH + i4).Value = s(a)
                            sheet3.Range(SCGL & j + SJWZH + i4).Value = ""   ''''''''实测高
                            z(a) = z(a - 1) + 200
                            i4 = i4 + 1
                            i = i - 1

                        Else '距离设转点
                            If Val(ZH0) - z(a) < -200 Then '4
                                sheet3.Range("a" & j + SJWZH + i4).Value = "ZD" & (a + 1)  '桩号
                                If (s(a) - Val(sheet1.Range("K" & i).Value) > -0.3 And sheet1.Range("Q" & i).Value = "是") Or (s(a) - Val(sheet1.Range("K" & i).Value) > 4.5 And sheet1.Range("Q" & i).Value <> "是") Then
                                    Randomize()
                                    sheet3.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)     '前视
                                    Randomize()
                                    sheet3.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)    '后视
                                Else
                                    If (s(a) - Val(sheet1.Range("K" & i).Value) < -4.5 And sheet1.Range("Q" & i).Value = "是") Or (s(a) - Val(sheet1.Range("K" & i).Value) < 0.3 And sheet1.Range("Q" & i).Value <> "是") Then
                                        Randomize()
                                        sheet3.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)   '前视
                                        Randomize()
                                        sheet3.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)    '后视
                                    Else
                                        Randomize()
                                        sheet3.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)   '前视
                                        Randomize()
                                        sheet3.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)    '后视
                                    End If
                                End If
                                sheet3.Range(SCGL & j + SJWZH + i4).Value = s(a) - Val(sheet3.Range(QSL & j + SJWZH + i4).Value)       '实测高程
                                a = a + 1 '交点数加1
                                s(a) = Val(sheet3.Range(SCGL & j + SJWZH + i4).Value) + Val(sheet3.Range(HSL & j + SJWZH + i4).Value)       '视线高
                                sheet3.Range(SXGL & j + SJWZH + i4).Value = s(a)
                                sheet3.Range(SCGL & j + SJWZH + i4).Value = ""    ''''''''实测高
                                z(a) = z(a - 1) - 200
                                i4 = i4 + 1
                                i = i - 1

                            Else '设转点
                                If (s(a) - Val(sheet1.Range("K" & i).Value) > -0.3 And sheet1.Range("Q" & i).Value = "是") Or (s(a) - Val(sheet1.Range("K" & i).Value) > 4.5 And sheet1.Range("Q" & i).Value <> "是") Then '3
                                    sheet3.Range("a" & j + SJWZH + i4).Value = "ZD" & (a + 1)     '桩号
                                    Randomize()
                                    sheet3.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)     '前视
                                    sheet3.Range(SCGL & j + SJWZH + i4).Value = s(a) - Val(sheet3.Range(QSL & j + SJWZH + i4).Value)       '实测高程
                                    Randomize()
                                    sheet3.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)     '后视
                                    a = a + 1 '交点数加1
                                    s(a) = Val(sheet3.Range(SCGL & j + SJWZH + i4).Value) + Val(sheet3.Range(HSL & j + SJWZH + i4).Value)       '视线高
                                    sheet3.Range(SXGL & j + SJWZH + i4).Value = s(a)
                                    sheet3.Range(SCGL & j + SJWZH + i4).Value = ""    ''''''''实测高

                                    z(a) = z(a - 1) '   桩号值
                                    i4 = i4 + 1
                                    i = i - 1

                                Else '设转点
                                    If (s(a) - Val(sheet1.Range("K" & i).Value) < -4.5 And sheet1.Range("Q" & i).Value = "是") Or (s(a) - Val(sheet1.Range("K" & i).Value) < 0.3 And sheet1.Range("Q" & i).Value <> "是") Then  '2
                                        sheet3.Range("a" & j + SJWZH + i4).Value = "ZD" & (a + 1)   '桩号
                                        Randomize()
                                        sheet3.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)    '前视
                                        sheet3.Range(SCGL & j + SJWZH + i4).Value = s(a) - sheet3.Range(QSL & j + SJWZH + i4).Value      '实测高程
                                        Randomize()
                                        sheet3.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)    '后视
                                        a = a + 1 '交点数加1
                                        s(a) = Val(sheet3.Range(SCGL & j + SJWZH + i4).Value) + Val(sheet3.Range(HSL & j + SJWZH + i4).Value)      '视线高
                                        sheet3.Range(SXGL & j + SJWZH + i4).Value = s(a)
                                        sheet3.Range(SCGL & j + SJWZH + i4).Value = ""    ''''''''实测高

                                        z(a) = z(a - 1)  '     桩号值
                                        i4 = i4 + 1
                                        i = i - 1

                                    Else '非转点求值
                                        If i <= i1 + i2 Then '1
                                            If sheet1.Range("O" & i).Value = Nothing And sheet1.Range("O" & i).Value <> "0" Then
                                                sheet3.Range(PCL & j + SJWZH + i4).Value = Int((Maxg - Ming + 1) * Rnd() + Ming)   '偏差
                                            Else
                                                sheet3.Range(PCL & j + SJWZH + i4).Value = sheet1.Range("O" & (i)).Value
                                            End If
                                            sheet3.Range(SJGL & j + SJWZH + i4).Value = Val(sheet1.Range("K" & (i)).Value)     '设计高程
                                            sheet3.Range(SCGL & j + SJWZH + i4).Value = Math.Round(sheet3.Range(SJGL & j + SJWZH + i4).Value + sheet3.Range(PCL & j + SJWZH + i4).Value / 1000, 3)          '实测高程
                                            sheet3.Range(ZSL & j + SJWZH + i4).Value = Math.Round(s(a) - Val(sheet3.Range(SCGL & j + SJWZH + i4).Value), 3)     '中视

                                            If sheet1.Range("Q" & i).Value = "是" Then     '判断是否为倒尺
                                                sheet3.Range(BZL & j + SJWZH + i4).Value = "倒尺"
                                            Else
                                                sheet3.Range(BZL & j + SJWZH + i4).Value = ""
                                            End If

                                            If sheet1.Range("R" & (i)).Value <> Nothing Then
                                                mc2 = "（" & sheet1.Range("R" & (i)).Value & "）"
                                            Else : mc2 = ""
                                            End If
                                            If Val(sheet1.Range("J" & (i)).Value) < 0 Then
                                                sheet3.Range("a" & j + SJWZH + i4).Value = ZJZH(sheet1.Range("I" & (i)).Value) & Chr(10) & "左" & XSWS(Math.Abs(Val(sheet1.Range("J" & (i)).Value))) & "m" & mc2

                                            Else
                                                If Val(sheet1.Range("J" & (i)).Value) > 0 Then
                                                    sheet3.Range("a" & j + SJWZH + i4).Value = ZJZH(sheet1.Range("I" & (i)).Value) & Chr(10) & "右" & XSWS(Math.Abs(Val(sheet1.Range("J" & (i)).Value))) & "m" & mc2

                                                Else : sheet3.Range("a" & j + SJWZH + i4).Value = ZJZH(sheet1.Range("I" & (i)).Value) & Chr(10) & "中" & mc2

                                                End If
                                            End If
                                            i4 = i4 + 1
                                        End If '1
                                    End If '2
                                End If '3
                            End If '4
                        End If '5
                        '新增表格
                        If i4 > SJH Then
                            j = j + BGH  '表格行数''
                            sheet3.Range("a" & j).PasteSpecial()
                            i4 = 1
                            sheet3.Range("E" & j + 3 + i4).Value = sheet1.Range("L" & (i1)).Value '工程部位
                            sheet3.Range("I" & j + 3 + i4).Value = sheet1.Range("P1").Value '时间
                            sheet3.Range("E" & j + i4 + 1).Value = "续上页" '编号
                            sheet3.Range("E" & j + i4 + 1).Interior.Color = 255
                            sheet3.Range("E" & j + i4 + 1).Font.Color = -16711681
                        End If

                    Next i

                    '闭合
                    For i = i1 + i2 - 1 To i1 + i2
                        '设转点
                        'Range("J" & j + 8 + i4) = szdz - Z(A) & "," & s(A) - Val(szdh)
                        If Val(szdz) - z(a) > 200 Then '设转点
                            sheet3.Range("a" & j + SJWZH + i4).Value = "ZD" & (a + 1) '桩号
                            If s(a) < Val(szdh) Then
                                Randomize()
                                sheet3.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)     '前视
                                Randomize()
                                sheet3.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)    '后视
                            Else
                                Randomize()
                                sheet3.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)    '前视
                                Randomize()
                                sheet3.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)     '后视
                            End If
                            sheet3.Range(SCGL & j + SJWZH + i4).Value = s(a) - Val(sheet3.Range(QSL & j + SJWZH + i4).Value)      '实测高程
                            a = a + 1 '交点数加1
                            s(a) = Val(sheet3.Range(SCGL & j + SJWZH + i4).Value) + Val(sheet3.Range(HSL & j + SJWZH + i4).Value)    '视线高
                            sheet3.Range(SXGL & j + SJWZH + i4).Value = s(a)
                            sheet3.Range(SCGL & j + SJWZH + i4).Value = ""    ''''''''实测高
                            z(a) = z(a - 1) + 200
                            i4 = i4 + 1
                            i = i - 1
                        Else '设转点
                            If Val(szdz) - z(a) < -200 Then
                                sheet3.Range("a" & j + SJWZH + i4).Value = "ZD" & (a + 1)   '桩号
                                If s(a) < Val(szdh) Then
                                    Randomize()
                                    sheet3.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)     '前视
                                    Randomize()
                                    sheet3.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)    '后视
                                Else
                                    Randomize()
                                    sheet3.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)    '前视
                                    Randomize()
                                    sheet3.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)     '后视
                                End If
                                sheet3.Range(SCGL & j + SJWZH + i4).Value = s(a) - Val(sheet3.Range(QSL & j + SJWZH + i4).Value)       '实测高程
                                a = a + 1 '交点数加1
                                s(a) = Val(sheet3.Range(SCGL & j + SJWZH + i4).Value) + Val(sheet3.Range(HSL & j + SJWZH + i4).Value)     '视线高
                                sheet3.Range(SXGL & j + SJWZH + i4).Value = s(a)
                                sheet3.Range(SCGL & j + SJWZH + i4).Value = ""    ''''''''实测高
                                i4 = i4 + 1
                                i = i - 1
                                z(a) = z(a - 1) - 200
                            Else '设转点
                                If s(a) - Val(szdh) > 4.6 Then
                                    sheet3.Range("a" & j + SJWZH + i4).Value = "ZD" & (a + 1)     '桩号
                                    Randomize()
                                    sheet3.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)     '前视
                                    sheet3.Range(SCGL & j + SJWZH + i4).Value = s(a) - Val(sheet3.Range(QSL & j + SJWZH + i4).Value)       '实测高程
                                    Randomize()
                                    sheet3.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)    '后视
                                    a = a + 1 '交点数加1
                                    s(a) = Val(sheet3.Range(SCGL & j + SJWZH + i4).Value) + Val(sheet3.Range(HSL & j + SJWZH + i4).Value)       '视线高
                                    sheet3.Range(SXGL & j + SJWZH + i4).Value = s(a)
                                    sheet3.Range(SCGL & j + SJWZH + i4).Value = ""    ''''''''实测高

                                    z(a) = z(a - 1)  '   桩号值

                                    i4 = i4 + 1
                                    i = i - 1
                                Else '设转点
                                    If s(a) - Val(szdh) < 0.4 Then
                                        sheet3.Range("a" & j + SJWZH + i4).Value = "ZD" & (a + 1)   '桩号
                                        Randomize()
                                        sheet3.Range(QSL & j + SJWZH + i4).Value = Math.Round(Int((TCmind - TCminx + 1) * Rnd() + TCminx) / 1000, 3)     '前视
                                        sheet3.Range(SCGL & j + SJWZH + i4).Value = s(a) - Val(sheet3.Range(QSL & j + SJWZH + i4).Value)      '实测高程
                                        Randomize()
                                        sheet3.Range(HSL & j + SJWZH + i4).Value = Math.Round(Int((TCmaxd - TCmaxx + 1) * Rnd() + TCmaxx) / 1000, 3)   '后视
                                        a = a + 1 '交点数加1
                                        s(a) = Val(sheet3.Range(SCGL & j + SJWZH + i4).Value) + Val(sheet3.Range(HSL & j + SJWZH + i4).Value)     '视线高
                                        sheet3.Range(SXGL & j + SJWZH + i4).Value = s(a)
                                        sheet3.Range(SCGL & j + SJWZH + i4).Value = ""    ''''''''实测高

                                        z(a) = z(a - 1)  ' 桩号值

                                        i4 = i4 + 1
                                        i = i - 15
                                    Else
                                        '非转点求值闭合差
                                        'Dim p_z As Single, p_f As Single
                                        ''p_z = 2 * Math.Sqrt(a + 1)
                                        ''p_f = -2 * Math.Sqrt(a + 1)
                                        'If ExApp.Sheets("数据库").range("O2").value = "二等" Then
                                        '    p_z = 0
                                        '    p_f = 0
                                        'Else
                                        '    If ExApp.Sheets("数据库").range("O2").value = "三等" Then
                                        '        p_z = 2
                                        '        p_f = -2
                                        '    Else
                                        '        If ExApp.Sheets("数据库").range("O2").value = "四等" Then
                                        '            p_z = 3
                                        '            p_f = -3
                                        '        End If
                                        '    End If

                                        'End If

                                        Randomize()
                                        'sheet3.Range(PCL & j + SJWZH + i4).value = Int((p_z - p_f + 1) * Rnd() + p_f)   '偏差
                                        sheet3.Range(PCL & j + SJWZH + i4).Value = 0   '偏差
                                        sheet3.Range(SJGL & j + SJWZH + i4).Value = szdh    '设计高程
                                        sheet3.Range(SCGL & j + SJWZH + i4).Value = Math.Round(Val(sheet3.Range(SJGL & j + SJWZH + i4).Value) + Val(sheet3.Range(PCL & j + SJWZH + i4).Value) / 1000, 3)       '实测高程
                                        sheet3.Range(QSL & j + SJWZH + i4).Value = Math.Round(s(a) - Val(sheet3.Range(SCGL & j + SJWZH + i4).Value), 3)      '前视
                                        sheet3.Range("a" & j + SJWZH + i4).Value = szdm

                                        ''备注列
                                        'sheet3.Range("B" & j + 28).value = "闭合或附合差fh=" &sheet3.Range(PCL & j + SJWZH + i4).value
                                        'sheet3.Range("B" & j + 30).value = "|fh|<|f允|，符合水准测量记录表测量技术规定。"
                                        'If ExApp.Sheets("数据库").range("O2").value = "三等" Then
                                        '    If a = 0 Then
                                        '       sheet3.Range("B" & j + 29).value = "容许误差f允=3.5√n=3.5√" & a + 1 & "=" & 3.5
                                        '    Else
                                        '       sheet3.Range("B" & j + 29).value = "容许误差f允=3.5√n=3.5√" & a + 1 & "=" & Math.Round(3.5 * Math.Sqrt(a + 1), 0)
                                        '    End If

                                        'Else
                                        '    If ExApp.Sheets("数据库").range("O2").value = "四等" Then
                                        '       sheet3.Range("B" & j + 29).value = "容许误差f允=6√n=6√" & a + 1 & "=" & Math.Round(6 * Math.Sqrt(a + 1), 0)
                                        '    End If
                                        '    If ExApp.Sheets("数据库").range("O2").value = "二等" Then
                                        '       sheet3.Range("B" & j + 29).value = "容许误差f允=4√l=4√" & Math.Round（(Math.Abs(szdz - ZH0) * 2 + ExApp.WorksheetFunction.RANDBETWEEN(1, 10)) / 1000， 3） & "=" & Math.Round(4 * Math.Sqrt((Math.Abs(szdz - ZH0) * 2 + ExApp.WorksheetFunction.RANDBETWEEN(1, 10)) / 1000), 0)
                                        '    End If

                                        'End If



                                    End If
                                End If
                            End If
                        End If
                        '新增表格
                        If i4 > SJH Then
                            j = j + BGH '表格行数
                            sheet3.Range("a" & j).PasteSpecial()
                            i4 = 1
                            sheet3.Range("e" & j + 3 + i4).Value = sheet1.Range("L" & (i1)).Value '工程部位
                            sheet3.Range("i" & j + 3 + i4).Value = sheet1.Range("P1").Value '时间
                            sheet3.Range("E" & j + i4 + 1).Value = "续上页" '编号
                            sheet3.Range("E" & j + i4 + 1).Interior.Color = 255
                            sheet3.Range("E" & j + i4 + 1).Font.Color = -16711681
                        End If
                    Next i
                    ''''闭合
                End If 'if123
                i1 = i1 + 1
            End While 'while123
        Catch Exclerror As Exception   '错误时弹出提示
            MsgBox(Exclerror.Message)
            TorF = False
            Exit Sub
        End Try
    End Sub

    Sub 全站仪平面位置检测表（）

        Dim zs As Double
        Dim i, pd, i1, i2, i3, j As Integer, jjj As Integer, z(238) As Object, s(238) As Object
        Dim min1, Sumz1， DXCGz, DXCGx, DXCGy, DXCGh, DXCGm, DXCGz1, DXCGx1, DXCGy1, DXCGh1, DXCGm1， JD, jdd, jdf, jdm, mc2, pdz， SJXL, SJYL, SCXL, SCYL, PCXL, PCYL, PWL, SJH, BGH, SJWZH， Minp, Maxp
        SJXL = "C"  '设计坐标X列
        SJYL = "E"  '设计坐标Y列
        SCXL = "G"  '实测坐标X列
        SCYL = "I"  '实测坐标Y列
        PCXL = "K"  '差　　值X列
        PCYL = "L"  '差　　值Y列
        PWL = "M"   '偏位 √(△X2+△Y2 列
        SJH = 10    '数据行数，表格计算数据的行数
        BGH = 22    '表格总行数
        SJWZH = 12  '数据位置开始行，表格数据从哪行开始
        j = 1
        i1 = 3
        pd = 0
        SJWZH = SJWZH - 2
        Minp = 0
        Maxp = 5
        Try
            Exbook = ExApp.ActiveWorkbook
            sheet1 = Exbook.Worksheets("参数表") '参数表
            sheet2 = Exbook.Worksheets("导线成果表") '导线成果表
            sheet3 = Exbook.Worksheets("平面表") '成孔平面
            ExApp.ScreenUpdating = False '关闭代码运行屏显

            sheet3.Select()
            CType(sheet3.Rows("23:99999"), Range).Delete()  '删除表格
            CType(sheet3.Rows("1:22"), Range).Copy()        '复制表格

            While sheet1.Range("I" & i1).Value <> Nothing 'while123
                If TorF = False Then
                    Exit Sub
                End If
                i3 = 1
                If sheet1.Range("L" & i1).Value <> Nothing Then 'if123
                    j = j + BGH  '表格行数
                    sheet3.Range("a" & j).PasteSpecial() '粘贴表格
                    pd = pd + 1
                    sheet3.Range("G" & j + 3 + i3).Value = sheet1.Range("L" & (i1)).Value '工程部位
                    sheet3.Range("O" & j + 3 + i3).Value = sheet1.Range("P1").Value '时间
                    sheet3.Range("O" & j + 5 + i3).Value = "P=" & ExApp.WorksheetFunction.RandBetween(715, 720) & "mmhg"   '气压
                    sheet3.Range("O" & j + 6 + i3).Value = "T=" & ExApp.WorksheetFunction.RandBetween(24, 28) & "℃"   '温度
                    sheet3.Range("O" & j + 7 + i3).Value = "I=" & Math.Round(1 + ExApp.WorksheetFunction.RandBetween(550, 680) * 0.001, 3) & "m"   '仪高
                    i2 = 1
                    ''''桩号
                    Sumz1 = Val(ExApp.WorksheetFunction.Substitute(sheet1.Range("I" & i1).Value, "*", ""))
                    While sheet1.Range("I" & i1 + i2).Value <> Nothing And sheet1.Range("L" & i1 + i2).Value = Nothing
                        Sumz1 = Sumz1 + Val(ExApp.WorksheetFunction.Substitute(sheet1.Range("I" & i1 + i2).Value, "*", ""))
                        i2 = i2 + 1
                    End While
                    Sumz1 = Sumz1 / i2 '桩号平均值
                    i2 = 3
                    '站点
                    min1 = 9999
                    For i2 = 3 To sheet2.Range("B1048576").End(XlDirection.xlUp).Row + 1

                        If sheet1.Range("J1").Value = "进口" Then

                            If Math.Abs(Sumz1 - sheet2.Range("F" & i2).Value) < min1 And sheet2.Range("E" & i2).Value <> Nothing And Sumz1 - sheet2.Range("F" & i2).Value >= 0 Then
                                min1 = Math.Abs(Sumz1 - sheet2.Range("F" & i2).Value)
                                DXCGz = sheet2.Range("F" & i2).Value  '导线点桩号
                                DXCGx = sheet2.Range("C" & i2).Value '导线点X
                                DXCGy = sheet2.Range("D" & i2).Value  '导线点Y
                                DXCGh = sheet2.Range("E" & i2).Value  '导线点h
                                DXCGm = sheet2.Range("B" & i2).Value  '导线点

                                DXCGz1 = sheet2.Range("F" & i2 - 1).Value '后视导线点桩号
                                DXCGx1 = sheet2.Range("C" & i2 - 1).Value '后导线点X
                                DXCGy1 = sheet2.Range("D" & i2 - 1).Value '后导线点Y
                                DXCGh1 = sheet2.Range("E" & i2 - 1).Value '后导线点h
                                DXCGm1 = sheet2.Range("B" & i2 - 1).Value '后导线点
                            End If

                        Else
                            If sheet1.Range("J1").Value = "出口" Then

                                If Math.Abs(Sumz1 - sheet2.Range("f" & i2).Value) < min1 And sheet2.Range("E" & i2).Value <> Nothing And sheet2.Range("f" & i2).Value - Sumz1 >= 0 Then
                                    min1 = Math.Abs(Sumz1 - sheet2.Range("F" & i2).Value)
                                    DXCGz = sheet2.Range("F" & i2).Value '导线点桩号
                                    DXCGx = sheet2.Range("C" & i2).Value '导线点X
                                    DXCGy = sheet2.Range("D" & i2).Value '导线点Y
                                    DXCGh = sheet2.Range("E" & i2).Value '导线点h
                                    DXCGm = sheet2.Range("B" & i2).Value '导线点

                                    DXCGz1 = sheet2.Range("F" & i2 + 1).Value '后视导线点桩号
                                    DXCGx1 = sheet2.Range("C" & i2 + 1).Value '后导线点X
                                    DXCGy1 = sheet2.Range("D" & i2 + 1).Value '后导线点Y
                                    DXCGh1 = sheet2.Range("E" & i2 + 1).Value '后导线点h
                                    DXCGm1 = sheet2.Range("B" & i2 + 1).Value '后导线点
                                End If
                            Else
                                If Math.Abs(Sumz1 - sheet2.Range("F" & i2).Value) < min1 And sheet2.Range("C" & i2).Value <> Nothing Then
                                    min1 = Math.Abs(Sumz1 - sheet2.Range("f" & i2).Value)
                                    DXCGz = sheet2.Range("F" & i2).Value '导线点桩号
                                    DXCGx = sheet2.Range("C" & i2).Value '导线点X
                                    DXCGy = sheet2.Range("D" & i2).Value '导线点Y
                                    DXCGh = sheet2.Range("E" & i2).Value '导线点h
                                    DXCGm = sheet2.Range("B" & i2).Value '导线点
                                End If
                            End If
                        End If
                    Next i2

                    '后视
                    If sheet1.Range("J1").Value = "就近" Then
                        i2 = 3
                        min1 = 9999
                        For i2 = 3 To 100
                            If Math.Abs(Sumz1 - sheet2.Range("f" & i2).Value) < min1 And DXCGm <> sheet2.Range("b" & i2).Value And sheet2.Range("C" & i2).Value <> Nothing Then
                                min1 = Math.Abs(Sumz1 - sheet2.Range("f" & i2).Value)
                                DXCGz1 = sheet2.Range("f" & i2).Value '后视导线点桩号
                                DXCGx1 = sheet2.Range("C" & i2).Value '后导线点X
                                DXCGy1 = sheet2.Range("D" & i2).Value '后导线点Y
                                DXCGh1 = sheet2.Range("e" & i2).Value '后导线点h
                                DXCGm1 = sheet2.Range("B" & i2).Value '后导线点
                            End If
                        Next i2
                    End If

                    '方位角角度
                    JD = ExApp.WorksheetFunction.Atan2((-DXCGx + DXCGx1), (-DXCGy + DXCGy1)) * 180 / ExApp.WorksheetFunction.Pi
                    If JD < 0 Then
                        JD = JD + 360
                    End If
                    jdd = Int(JD)
                    jdf = Int((JD - Int(JD)) * 60)
                    jdm = Int(((JD - Int(JD)) * 60 - Int((JD - Int(JD)) * 60)) * 60)
                    '改测点后视点
                    sheet3.Range("B" & j + 5 + i3).Value = DXCGm
                    sheet3.Range("F" & j + 5 + i3).Value = DXCGx
                    sheet3.Range("F" & j + 6 + i3).Value = DXCGy
                    sheet3.Range("F" & j + 7 + i3).Value = DXCGh
                    sheet3.Range("H" & j + 5 + i3).Value = DXCGm1
                    sheet3.Range("K" & j + 5 + i3).Value = DXCGx1
                    sheet3.Range("K" & j + 6 + i3).Value = DXCGy1
                    sheet3.Range("K" & j + 7 + i3).Value = DXCGh1
                    '距离
                    sheet3.Range("M" & j + 5 + i3).Value = "S=" & Math.Round(Math.Sqrt((DXCGx - DXCGx1) ^ 2 + (DXCGy - DXCGy1) ^ 2), 3) & "m"
                    sheet3.Range("M" & j + 6 + i3).Value = "ɑ=" & jdd & "°" & jdf & "′" & jdm & "″"
                    i2 = 1
                    While sheet1.Range("L" & (i1 + i2)).Value = Nothing And sheet1.Range("I" & (i1 + i2)).Value <> Nothing
                        i2 = i2 + 1
                    End While
                    For i = i1 To i1 + i2 - 1
                        '设计坐标
                        sheet3.Range(SJXL & j + SJWZH + i3).Value = sheet1.Range("M" & i).Value
                        sheet3.Range(SJYL & j + SJWZH + i3).Value = sheet1.Range("N" & i).Value
                        '偏差
                        If sheet1.Range("P" & i).Value = Nothing And sheet1.Range("P" & i).Value <> "0" Then '自动偏差
                            sheet3.Range(PWL & j + SJWZH + i3).Value = ExApp.WorksheetFunction.RandBetween(Minp, Maxp)
                            pdz = ExApp.WorksheetFunction.RandBetween(-1, 1)
                            If pdz < 0 Then
                                pdz = -1
                            Else
                                pdz = 1
                            End If
                            sheet3.Range(PCXL & j + SJWZH + i3).Value = pdz * Math.Round(Math.Sqrt(ExApp.WorksheetFunction.RandBetween(0, sheet3.Range(PWL & j + SJWZH + i3).Value * sheet3.Range(PWL & j + SJWZH + i3).Value / 2)), 0)
                            pdz = ExApp.WorksheetFunction.RandBetween(-1, 1)
                            If pdz < 0 Then
                                pdz = -1
                            Else
                                pdz = 1
                            End If
                            sheet3.Range(PCYL & j + SJWZH + i3).Value = pdz * Math.Round(Math.Sqrt(sheet3.Range(PWL & j + SJWZH + i3).Value * sheet3.Range(PWL & j + SJWZH + i3).Value - sheet3.Range(PCXL & j + SJWZH + i3).Value * sheet3.Range(PCXL & j + SJWZH + i3).Value), 0)
                        Else
                            '手动偏差
                            sheet3.Range(PWL & j + SJWZH + i3).Value = sheet1.Range("P" & i).Value
                            pdz = ExApp.WorksheetFunction.RandBetween(-1, 1)
                            If pdz < 0 Then
                                pdz = -1
                            Else
                                pdz = 1
                            End If
                            sheet3.Range(PCXL & j + SJWZH + i3).Value = pdz * Math.Round(Math.Sqrt(ExApp.WorksheetFunction.RandBetween(0, sheet3.Range(PWL & j + SJWZH + i3).Value * sheet3.Range(PWL & j + SJWZH + i3).Value / 2)), 0)
                            pdz = ExApp.WorksheetFunction.RandBetween(-1, 1)
                            If pdz < 0 Then
                                pdz = -1
                            Else
                                pdz = 1
                            End If
                            sheet3.Range(PCYL & j + SJWZH + i3).Value = pdz * Math.Round(Math.Sqrt(sheet3.Range(PWL & j + SJWZH + i3).Value * sheet3.Range(PWL & j + SJWZH + i3).Value - sheet3.Range(PCXL & j + SJWZH + i3).Value * sheet3.Range(PCXL & j + SJWZH + i3).Value), 0)
                        End If
                        '实测
                        sheet3.Range(SCXL & j + SJWZH + i3).Value = sheet3.Range(SJXL & j + SJWZH + i3).Value + sheet3.Range(PCXL & j + SJWZH + i3).Value / 1000
                        sheet3.Range(SCYL & j + SJWZH + i3).Value = sheet3.Range(SJYL & j + SJWZH + i3).Value + sheet3.Range(PCYL & j + SJWZH + i3).Value / 1000
                        '桩号
                        If sheet1.Range("R" & (i)).Value <> Nothing Then
                            mc2 = "（" & sheet1.Range("R" & (i)).Value & "）"
                        Else : mc2 = ""
                        End If
                        If sheet1.Range("J" & (i)).Value < 0 Then
                            sheet3.Range("a" & j + SJWZH + i3).Value = ZJZH(sheet1.Range("I" & (i)).Value) & "左" & XSWS(Math.Abs(Val(sheet1.Range("J" & (i)).Value))) & "m" & mc2
                        Else
                            If sheet1.Range("J" & (i)).Value > 0 Then
                                sheet3.Range("a" & j + SJWZH + i3).Value = ZJZH(sheet1.Range("I" & (i)).Value) & "右" & XSWS(Math.Abs(Val(sheet1.Range("J" & (i)).Value))) & "m" & mc2
                            Else
                                sheet3.Range("a" & j + SJWZH + i3).Value = ZJZH(sheet1.Range("I" & (i)).Value) & "中" & mc2
                            End If
                        End If
                        i3 = i3 + 1


                        '新增表格
                        If i3 > SJH And sheet1.Range("L" & i + 1).Value = Nothing Then 'if1
                            j = j + BGH  '表格行数
                            sheet3.Range("a" & j).PasteSpecial() '粘贴表格
                            i3 = 1
                            sheet3.Range("G" & j + 3 + i3).Value = sheet1.Range("L" & (i1)).Value  '桩号
                            sheet3.Range("I" & j + i3 + 1).Value = "续上页" '序号
                            sheet3.Range("I" & j + i3 + 1).Interior.Color = 255
                            sheet3.Range("I" & j + i3 + 1).Font.Color = -16711681
                            sheet3.Range("O" & j + 3 + i3).Value = sheet1.Range("P1").Value '时间
                            sheet3.Range("O" & j + 5 + i3).Value = "P=" & ExApp.WorksheetFunction.RandBetween(715, 720) & "mmhg"   '气压
                            sheet3.Range("O" & j + 6 + i3).Value = "T=" & ExApp.WorksheetFunction.RandBetween(24, 28) & "℃"   '温度
                            sheet3.Range("O" & j + 7 + i3).Value = "I=" & 1 + ExApp.WorksheetFunction.RandBetween(550, 680) * 0.001 & "m"   '仪高
                            sheet3.Range("B" & j + 5 + i3).Value = DXCGm '控制点
                            sheet3.Range("F" & j + 5 + i3).Value = DXCGx
                            sheet3.Range("F" & j + 6 + i3).Value = DXCGy
                            sheet3.Range("F" & j + 7 + i3).Value = DXCGh
                            sheet3.Range("H" & j + 5 + i3).Value = DXCGm1
                            sheet3.Range("K" & j + 5 + i3).Value = DXCGx1
                            sheet3.Range("K" & j + 6 + i3).Value = DXCGy1
                            sheet3.Range("K" & j + 7 + i3).Value = DXCGh1
                            sheet3.Range("M" & j + 5 + i3).Value = "S=" & Math.Round(Math.Sqrt((DXCGx - DXCGx1) ^ 2 + (DXCGy - DXCGy1) ^ 2), 3) & "m"
                            sheet3.Range("M" & j + 6 + i3).Value = "ɑ=" & jdd & "°" & jdf & "′" & jdm & "″"

                        End If

                    Next i
                End If 'if123
                i1 = i1 + 1
            End While 'while123
        Catch Exclerror As Exception   '错误时弹出提示
            MsgBox(Exclerror.Message)
            TorF = False
            Exit Sub
        End Try
    End Sub
End Module
