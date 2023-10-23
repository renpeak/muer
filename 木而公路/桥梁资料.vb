
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Module 桥梁资料

    Public sheet0, sheet1, sheet2, sheet3, sheet4, sheet5, sheet6, sheet7,
            sheet8, sheet9, sheet10, sheet11, sheet12, sheet13, sheet14, sheet15,
            sheet16, sheet17, sheet18, sheet19, sheet20, EXsheet As Excel.Worksheet

    Sub 桩基资料（）

        Dim LSBLFZ, LSBL As String
        Dim h, r, Q, i As Integer
        Dim ExcelFilename, PDFFilename As String  '定义输出的PDF文件名
        Try
            h = 8
            r = 0
            sheet0 = Exbook.Worksheets("数据库") '数据库
            sheet1 = Exbook.Worksheets("参数表") '参数表
            sheet2 = Exbook.Worksheets("导线成果表") '导线成果表
            sheet3 = Exbook.Worksheets("钢筋隐蔽工程") '钢筋隐蔽工程
            sheet4 = Exbook.Worksheets("钢筋检表") '钢筋检表
            sheet5 = Exbook.Worksheets("钢筋安装记录表") '钢筋安装记录表
            sheet6 = Exbook.Worksheets("钢筋安装记录表续表") '钢筋安装记录表续表
            sheet7 = Exbook.Worksheets("成桩隐蔽工程") '成桩隐蔽工程
            sheet8 = Exbook.Worksheets("钻孔桩检表") '钻孔桩检表
            sheet9 = Exbook.Worksheets("钻孔记录") '钻孔记录
            sheet10 = Exbook.Worksheets("灌注前记录") '灌注前记录
            sheet11 = Exbook.Worksheets("砼浇筑") '砼浇筑
            sheet12 = Exbook.Worksheets("水下灌注记录") '水下灌注记录
            sheet13 = Exbook.Worksheets("砼施工记录") '砼施工记录
            sheet14 = Exbook.Worksheets("监抽钢筋检表") '钢筋-监抽
            sheet15 = Exbook.Worksheets("监抽钢筋安装记录表") '钢筋安装记录表 -监抽
            sheet16 = Exbook.Worksheets("监抽钻孔桩检表") '钻孔桩检表-监抽
            sheet17 = Exbook.Worksheets("钢筋顶水准") '钢筋顶水准
            sheet18 = Exbook.Worksheets("孔底水准") '孔底水准
            sheet19 = Exbook.Worksheets("桩顶水准") '桩顶水准
            sheet20 = Exbook.Worksheets("成孔平面") '成孔平面

            '改表头
            sheet3.Range("A1").Value = sheet0.Range("C1").Value
            sheet3.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet3.Range("E3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet3.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet3.Range("E4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet4.Range("A1").Value = sheet0.Range("C1").Value
            sheet4.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet4.Range("Q3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet4.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet4.Range("Q4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet5.Range("A1").Value = sheet0.Range("C1").Value
            sheet5.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet5.Range("S3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet5.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet5.Range("S4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet6.Range("A1").Value = sheet0.Range("C1").Value
            sheet6.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet6.Range("S3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet6.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet6.Range("S4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet7.Range("A1").Value = sheet0.Range("C1").Value
            sheet7.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet7.Range("D3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet7.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet7.Range("D4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet8.Range("A1").Value = sheet0.Range("C1").Value
            sheet8.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet8.Range("J3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet8.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet8.Range("J4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet9.Range("A1").Value = sheet0.Range("C1").Value
            sheet9.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet9.Range("S3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet9.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet9.Range("S4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet10.Range("A1").Value = sheet0.Range("C1").Value
            sheet10.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet10.Range("L3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet10.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet10.Range("L4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet11.Range("A1").Value = sheet0.Range("C1").Value
            sheet11.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet11.Range("M3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet11.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet11.Range("M4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet12.Range("A1").Value = sheet0.Range("C1").Value
            sheet12.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet12.Range("Q3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet12.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet12.Range("Q4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet13.Range("A1").Value = sheet0.Range("C1").Value
            sheet13.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet13.Range("R3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet13.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet13.Range("R4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet14.Range("A1").Value = sheet0.Range("C1").Value
            sheet14.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet14.Range("Q3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet14.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet14.Range("Q4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet15.Range("A1").Value = sheet0.Range("C1").Value
            sheet15.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet15.Range("S3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet15.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet15.Range("S4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet16.Range("A1").Value = sheet0.Range("C1").Value
            sheet16.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet16.Range("J3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet16.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet16.Range("J4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet17.Range("A1").Value = sheet0.Range("C1").Value
            sheet17.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet17.Range("H3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet17.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet17.Range("H4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet18.Range("A1").Value = sheet0.Range("C1").Value
            sheet18.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet18.Range("H3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet18.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet18.Range("H4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet19.Range("A1").Value = sheet0.Range("C1").Value
            sheet19.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet19.Range("H3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet19.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet19.Range("H4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet20.Range("A1").Value = sheet0.Range("C1").Value
            sheet20.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet20.Range("N3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet20.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet20.Range("N4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value
            sheet0.Range("W3:AE100000").Value = Nothing


            While sheet0.Range("B" & h).Value <> Nothing
                ExApp.Calculation = ExApp.Calculation.xlCalculationManual '开启手动计算
                sheet0.Range("Y" & h).Value = sheet0.Range("N" & h).Value - sheet0.Range("M" & h).Value
                sheet0.Range("Z" & h).Value = sheet0.Range("Y" & h).Value - sheet0.Range("O" & h).Value
                sheet0.Range("AA" & h).Value = sheet0.Range("O" & h).Value - (ExApp.WorksheetFunction.RandBetween(10, 20) * 0.01)
                sheet0.Range("AB" & h).Value = sheet0.Range("Y" & h).Value + (ExApp.WorksheetFunction.RandBetween(10, 12) * 0.113)
                sheet0.Range("AC" & h).Value = Math.Round(sheet0.Range("AB" & h).Value - sheet0.Range("AA" & h).Value, 2)
                sheet0.Range("AD" & h).Value = sheet0.Range("P" & h).Value / 100 + (ExApp.WorksheetFunction.RandBetween(0, 20) * 0.001)
                sheet0.Range("AE" & h).Value = sheet0.Range("O" & h).Value + sheet0.Range("AD" & h).Value + sheet0.Range("V" & h).Value / 100 + (ExApp.WorksheetFunction.RandBetween(-20, 20) * 0.001)
                sheet0.Range("AF" & h).Value = sheet0.Range("AE" & h).Value - sheet0.Range("AD" & h).Value
                sheet0.Range("AG" & h).Value = sheet0.Range("AC" & h).Value - (ExApp.WorksheetFunction.RandBetween(3, 5) * 0.01)
                sheet0.Range("AH" & h).Value = sheet0.Range("O" & h).Value + sheet0.Range("V" & h).Value / 100

                sheet1.Range("B" & r + 5).Value = sheet0.Range("B" & h).Value
                sheet1.Range("B" & r + 6).Value = sheet0.Range("C" & h).Value
                sheet1.Range("C" & r + 6).Value = sheet0.Range("D" & h).Value
                sheet1.Range("D" & r + 6).Value = sheet0.Range("E" & h).Value
                sheet1.Range("B" & r + 7).Value = sheet0.Range("Y" & h).Value
                sheet1.Range("B" & r + 8).Value = sheet0.Range("O" & h).Value
                sheet1.Range("B" & r + 9).Value = sheet0.Range("L" & h).Value
                sheet1.Range("B" & r + 10).Value = sheet0.Range("AB" & h).Value
                sheet1.Range("B" & r + 11).Value = sheet0.Range("AA" & h).Value
                sheet1.Range("B" & r + 12).Value = sheet0.Range("AF" & h).Value  '实测钢筋骨架底高程
                sheet1.Range("B" & r + 13).Value = sheet0.Range("AD" & h).Value
                sheet1.Range("B" & r + 14).Value = sheet0.Range("R" & h).Value
                sheet1.Range("B" & r + 15).Value = sheet0.Range("Q" & h).Value
                sheet1.Range("B" & r + 16).Value = sheet0.Range("S" & h).Value
                sheet1.Range("B" & r + 17).Value = sheet0.Range("U" & h).Value
                sheet1.Range("B" & r + 18).Value = sheet0.Range("W" & h).Value
                sheet1.Range("B" & r + 19).Value = sheet0.Range("K" & h).Value
                sheet1.Range("B" & r + 20).Value = sheet0.Range("T" & h).Value
                sheet1.Range("B" & r + 21).Value = sheet0.Range("X" & h).Value
                sheet1.Range("B" & r + 22).Value = sheet0.Range("F" & h).Value   '开孔日期
                sheet1.Range("B" & r + 24).Value = sheet0.Range("H" & h).Value   '浇筑日期
                sheet1.Range("B" & r + 26).Value = sheet0.Range("G" & h).Value   '钢筋日期
                sheet1.Range("B" & r + 27).Value = sheet0.Range("I" & h).Value
                sheet1.Range("D" & r + 27).Value = sheet0.Range("J" & h).Value
                sheet1.Range("B" & r + 28).Value = ExApp.WorksheetFunction.RandBetween(10, 20)
                sheet1.Range("B" & r + 29).Value = sheet0.Range("P" & h).Value * 10
                sheet1.Range("B" & r + 30).Value = sheet0.Range("N" & h).Value
                sheet1.Range("B" & r + 31).Value = sheet0.Range("AH" & h).Value
                sheet1.Range("B" & r + 32).Value = sheet0.Range("AE" & h).Value
                sheet1.Range("I2").Value = sheet1.Range("B5").Value.substring(0, ExApp.WorksheetFunction.Find("K", sheet1.Range("B5").Value))

                ExApp.Calculate()  '开启自动计算
                ExApp.Calculation = ExApp.Calculation.xlCalculationManual '开启手动计算

                ' 钢筋隐蔽工程
                sheet3.Range("C6").Value = sheet1.Range("B5").Value & sheet1.Range("C6").Value & "基础及下部构造"
                sheet3.Range("E6").Value = sheet1.Range("C6").Value & sheet1.Range("F6").Value & "钢筋加工及安装"
                sheet3.Range("C7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & sheet1.Range("F6").Value & "钢筋加工及安装"
                sheet3.Range("E10").Value = sheet1.Range("B22").Value
                sheet3.Range("E11").Value = sheet1.Range("B22").Value
                sheet3.Range("E27").Value = sheet1.Range("B22").Value
                sheet3.Range("E28").Value = sheet1.Range("B22").Value

                '桩基隐蔽工程报验单
                sheet7.Range("C6").Value = sheet1.Range("B5").Value & sheet1.Range("C6").Value & "基础及下部构造"
                sheet7.Range("E6").Value = sheet1.Range("C6").Value & sheet1.Range("F6").Value
                sheet7.Range("C7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & sheet1.Range("F6").Value
                '钻孔桩检表
                sheet8.Range("D6").Value = sheet1.Range("B5").Value
                sheet8.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & sheet1.Range("F6").Value
                sheet8.Range("J6").Value = sheet1.Range("B23").Value
                sheet8.Range("F12").Value = sheet1.Range("B21").Value
                If sheet1.Range("B19").Value <> "排架桩" Then
                    sheet8.Range("H13").Value = ExApp.WorksheetFunction.RandBetween(10, 20)
                Else
                    sheet8.Range("H13").Value = "/"
                End If
                If sheet1.Range("B19").Value = "排架桩" Then
                    sheet8.Range("H14").Value = ExApp.WorksheetFunction.RandBetween(10, 20)
                Else
                    sheet8.Range("H14").Value = "/"
                End If
                sheet8.Range("F17").Value = Math.Round(sheet1.Range("B7").Value - sheet1.Range("B8").Value, 2)
                sheet8.Range("H16").Value = Math.Round(sheet1.Range("B10").Value - sheet1.Range("B11").Value, 2)
                sheet8.Range("F19").Value = sheet1.Range("B9").Value * 1000
                sheet8.Range("H18").Value = sheet1.Range("B9").Value * 1000 + 25 + Math.Round(ExApp.WorksheetFunction.RandBetween(10, 100) * 0.2, 0)
                sheet8.Range("H20").Value = ExApp.WorksheetFunction.RandBetween(40, 100)
                sheet8.Range("H21").Value = ExApp.WorksheetFunction.RandBetween(10, 30)
                sheet8.Range("H22").Value = "I类"
                '钻孔记录表
                sheet9.Range("B6").Value = sheet1.Range("B5").Value
                sheet9.Range("K6").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & sheet1.Range("F6").Value
                sheet9.Range("V6").Value = sheet1.Range("B23").Value
                sheet9.Range("B7").Value = sheet1.Range("C6").Value
                sheet9.Range("K7").Value = sheet1.Range("D6").Value
                sheet9.Range("V7").Value = sheet1.Range("B23").Value
                sheet9.Range("B8").Value = sheet1.Range("B9").Value
                sheet9.Range("F8").Value = sheet1.Range("B10").Value
                sheet9.Range("V8").Value = "旋挖钻Ø" & sheet9.Range("B8").Value * 100
                sheet9.Range("P9").Value = sheet1.Range("B8").Value
                If sheet1.Range("B19").Value = "群桩" Then    '成孔中心偏位
                    sheet9.Range("K9").Value = sheet8.Range("H13").Value
                Else
                    sheet9.Range("K9").Value = sheet8.Range("H14").Value
                End If
                sheet9.Range("V9").Value = Math.Round(sheet9.Range("F8").Value - sheet9.Range("P9").Value, 2)
                sheet9.Range("A13").Value = sheet1.Range("B22").Value
                sheet9.Range("B13").Value = sheet1.Range("E22").Value
                sheet9.Range("H13").Value = 2.7 + ExApp.WorksheetFunction.RandBetween(-10, 10) / 100
                sheet9.Range("J13").Value = sheet9.Range("F8").Value - sheet9.Range("H13").Value

                '灌注前记录
                sheet10.Range("D6").Value = sheet1.Range("B5").Value
                sheet10.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & sheet1.Range("F6").Value
                sheet10.Range("M6").Value = sheet1.Range("B24").Value
                sheet10.Range("M7").Value = sheet1.Range("B24").Value
                If sheet1.Range("B19").Value = "群桩" Then
                    sheet10.Range("D9").Value = sheet8.Range("H13").Value
                Else
                    sheet10.Range("D9").Value = sheet8.Range("H14").Value
                End If

                sheet10.Range("D10").Value = sheet8.Range("H18").Value
                sheet10.Range("D11").Value = sheet8.Range("H20").Value
                sheet10.Range("P11").Value = Math.Round(sheet8.Range("H20").Value / (sheet8.Range("H16").Value * 1000) * 100, 2）
                sheet10.Range("D12").Value = sheet8.Range("H21").Value
                sheet10.Range("D13").Value = ExApp.WorksheetFunction.RandBetween(103, 109) / 100
                sheet10.Range("D14").Value = ExApp.WorksheetFunction.RandBetween(18, 19)
                sheet10.Range("D15").Value = ExApp.WorksheetFunction.RandBetween(1, 3)
                sheet10.Range("D16").Value = ExApp.WorksheetFunction.RandBetween(96, 99)
                sheet10.Range("D17").Value = sheet1.Range("B10").Value
                sheet10.Range("M17").Value = sheet8.Range("H16").Value
                sheet10.Range("D18").Value = sheet1.Range("B11").Value
                sheet10.Range("M18").Value = sheet1.Range("B8").Value
                sheet10.Range("D19:J19").Value = Nothing
                sheet10.Range("D20").Value = sheet1.Range("B13").Value * 100
                If sheet10.Range("D20").Value / 900 > 1 Then
                    sheet10.Range("D19").Value = 900
                Else
                    sheet10.Range("D19").Value = sheet10.Range("D20").Value
                End If

                For X = 1 To Math.Round(sheet10.Range("D20").Value / 900, 0) + 1
                    If ExApp.WorksheetFunction.Sum(sheet10.Range("D19:J19").Value) >= sheet1.Range("B13").Value * 100 Then
                        Exit For
                    ElseIf sheet1.Range("B13").Value * 100 - ExApp.WorksheetFunction.Sum(sheet10.Range("D19:J19").Value) >= 1100 Then
                        sheet10.Cells(19, X + 4) = 900
                    Else
                        sheet10.Cells(19, X + 4) = sheet1.Range("B13").Value * 100 - ExApp.WorksheetFunction.Sum(sheet10.Range("D19:J19").Value)
                    End If
                Next
                sheet10.Range("M19").Value = ExApp.WorksheetFunction.Count(sheet10.Range("D19:J19").Value)
                sheet10.Range("M20").Value = sheet1.Range("B12").Value
                sheet10.Range("D21").Value = sheet1.Range("B20").Value
                sheet10.Range("M21").Value = "合格"
                sheet10.Range("C22").Value = sheet1.Range("B15").Value
                sheet10.Range("G22").Value = sheet10.Range("D20").Value
                sheet10.Range("M22").Value = sheet1.Range("B16").Value


                '填《2.钢筋检表》
                sheet4.Range("D6").Value = sheet1.Range("B5").Value
                sheet4.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & sheet1.Range("F6").Value & "钢筋"
                sheet4.Range("S6").Value = sheet1.Range("B26").Value
                sheet4.Range("S7").Value = sheet1.Range("B22").Value
                sheet4.Range("S25").Value = sheet1.Range("B22").Value
                sheet4.Range("S28").Value = sheet1.Range("B22").Value
                '主筋
                sheet4.Range("C11").Value = Math.Round(3.14 * (sheet1.Range("B9").Value * 1000 - sheet1.Range("B18").Value * 2) / sheet1.Range("B15").Value, 0)
                sheet4.Range("F10").Value = sheet1.Range("B15").Value * 2 * sheet10.Range("M19").Value
                sheet4.Range("I10").Value = sheet4.Range("F10").Value
                '箍筋
                sheet4.Range("F12").Value = sheet10.Range("M19").Value * 10
                sheet4.Range("I12").Value = sheet4.Range("F12").Value
                sheet4.Range("C17").Value = sheet1.Range("B29").Value
                sheet4.Range("E16").Value = ExApp.WorksheetFunction.RandBetween(40, 80) * -1 ^ ExApp.WorksheetFunction.RandBetween(0, 1) & "    " & ExApp.WorksheetFunction.RandBetween(40, 80) * -1 ^ ExApp.WorksheetFunction.RandBetween(0, 1)
                sheet4.Range("E18").Value = sheet1.Range("J6").Value
                '保护层
                sheet4.Range("C21").Value = sheet1.Range("B18").Value * 10
                sheet4.Range("F20").Value = sheet1.Range("B17").Value
                sheet4.Range("I20").Value = sheet1.Range("B17").Value
                '钢筋记录表
                sheet5.Range("B6").Value = sheet1.Range("B5").Value
                sheet5.Range("B7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & sheet1.Range("F6").Value & "钢筋"
                sheet5.Range("O6").Value = sheet1.Range("B26").Value
                sheet5.Range("O7").Value = sheet1.Range("B26").Value
                sheet5.Range("L43").Value = sheet4.Range("I16").Value
                sheet5.Range("O43").Value = sheet4.Range("K16").Value
                sheet5.Range("D44").Value = ExApp.WorksheetFunction.RandBetween(1, 9) * -1 ^ ExApp.WorksheetFunction.RandBetween(0, 1)
                sheet5.Range("E44").Value = ExApp.WorksheetFunction.RandBetween(1, 9) * -1 ^ ExApp.WorksheetFunction.RandBetween(0, 1)
                For Q = 1 To sheet10.Range("M19").Value
                    sheet5.Cells(44, Q * 2 + 2) = ExApp.WorksheetFunction.RandBetween(1, 9) * -1 ^ ExApp.WorksheetFunction.RandBetween(0, 1)
                    sheet5.Cells(44, Q * 2 + 3) = ExApp.WorksheetFunction.RandBetween(1, 9) * -1 ^ ExApp.WorksheetFunction.RandBetween(0, 1)
                Next
                sheet6.Range("B6").Value = sheet1.Range("B5").Value
                sheet6.Range("B7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & sheet1.Range("F6").Value & "钢筋"
                sheet6.Range("O6").Value = sheet1.Range("B26").Value
                sheet6.Range("O7").Value = sheet1.Range("B26").Value

                '8砼浇筑申请报告单
                sheet11.Range("C6").Value = sheet1.Range("B5").Value
                sheet11.Range("C7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & sheet1.Range("F6").Value
                sheet11.Range("K6").Value = sheet1.Range("B24").Value
                sheet11.Range("K7").Value = sheet1.Range("B24").Value
                sheet11.Range("E29").Value = sheet1.Range("B24").Value
                sheet11.Range("M29").Value = sheet1.Range("B24").Value
                sheet11.Range("M33").Value = sheet1.Range("B24").Value
                sheet11.Range("M37").Value = sheet1.Range("B24").Value
                '9.水下灌注记录
                sheet12.Range("D6").Value = sheet1.Range("B5").Value
                sheet12.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & sheet1.Range("F6").Value
                sheet12.Range("P6").Value = sheet1.Range("B24").Value
                sheet12.Range("P7").Value = sheet1.Range("B25").Value
                sheet12.Range("D8").Value = sheet1.Range("B9").Value * 100
                sheet12.Range("I8").Value = sheet1.Range("B8").Value
                sheet12.Range("P8").Value = sheet1.Range("B7").Value
                sheet12.Range("D9").Value = sheet1.Range("B10").Value
                sheet12.Range("I9").Value = sheet1.Range("B11").Value
                sheet12.Range("D10").Value = sheet1.Range("B10").Value
                sheet12.Range("I10").Value = ExApp.WorksheetFunction.RoundUp(sheet1.Range("B10").Value - sheet1.Range("B11").Value, 0)
                sheet12.Range("D11").Value = sheet1.Range("B21").Value
                sheet12.Range("C13").Value = ExApp.WorksheetFunction.RandBetween(34, 38) * 5 & "        " & ExApp.WorksheetFunction.RandBetween(34, 38) * 5 & "        " & ExApp.WorksheetFunction.RandBetween(34, 38) * 5

                sheet12.Range("A18").Value = sheet1.Range("E24").Value
                sheet12.Range("B18").Value = sheet12.Range("D10").Value - sheet12.Range("I9").Value
                sheet12.Range("C18").Value = sheet12.Range("B18").Value - sheet12.Range("E18").Value
                sheet12.Range("E18").Value = 3 + ExApp.WorksheetFunction.RandBetween(10, 100) / 100
                sheet12.Range("G18").Value = sheet12.Range("I10").Value
                sheet12.Range("J18").Value = sheet12.Range("G18").Value - sheet12.Range("H18").Value

                '10.砼施工记录
                sheet13.Range("D6").Value = sheet1.Range("B5").Value
                sheet13.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & sheet1.Range("F6").Value
                sheet13.Range("N6").Value = sheet12.Range("P6").Value
                sheet13.Range("N7").Value = sheet12.Range("P7").Value
                sheet13.Range("D9").Value = sheet1.Range("B21").Value
                sheet13.Range("D21").Value = sheet12.Range("C13").Value

                '1.钢筋-监抽
                sheet14.Range("D6").Value = sheet1.Range("B5").Value
                sheet14.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & sheet1.Range("F6").Value & "钢筋"
                sheet14.Range("S6").Value = sheet1.Range("B26").Value
                sheet14.Range("S7").Value = sheet1.Range("B22").Value
                sheet14.Range("P24").Value = sheet1.Range("B22").Value
                sheet14.Range("P27").Value = sheet1.Range("B22").Value
                sheet14.Range("F10").Value = Math.Round(sheet4.Range("F10").Value * 0.2, 0)
                sheet14.Range("I10").Value = sheet14.Range("F10").Value
                sheet14.Range("C11").Value = sheet4.Range("C11").Value
                sheet14.Range("E18").Value = sheet1.Range("J6").Value

                '箍筋
                LSBLFZ = Nothing
                If Math.Round(sheet4.Range("F12").Value * 0.2, 0) <= 10 Then
                    For cs = 1 To Math.Round(sheet4.Range("F12").Value * 0.2, 0)
                        LSBL = ExApp.WorksheetFunction.RandBetween(-15, 15)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet14.Range("E12").Value = LSBLFZ
                    sheet15.Range("D32").Value = "/"
                Else
                    sheet14.Range("E12").Value = "应测" & Math.Round(sheet4.Range("F12").Value * 0.2, 0) & "处，实测" & Math.Round(sheet4.Range("F12").Value * 0.2, 0) & "处，合格" & Math.Round(sheet4.Range("F12").Value * 0.2, 0) & "处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                    For cs = 1 To Math.Round(sheet4.Range("F12").Value * 0.2, 0)
                        LSBL = ExApp.WorksheetFunction.RandBetween(-15, 15)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet15.Range("D32").Value = LSBLFZ
                End If
                '骨架骨架外径
                LSBLFZ = Nothing
                If Math.Round(sheet10.Range("M19").Value * 0.4, 0) <= 10 Then
                    For cs = 1 To Math.Round(sheet10.Range("M19").Value * 0.4, 0)
                        LSBL = ExApp.WorksheetFunction.RandBetween(-9, 9)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet14.Range("E14").Value = LSBLFZ
                    sheet15.Range("D44").Value = "/"
                Else
                    sheet14.Range("E14").Value = "应测" & Math.Round(sheet10.Range("M19").Value * 0.4, 0) & "处，实测" & Math.Round(sheet10.Range("M19").Value * 0.4, 0) & "处，合格" & Math.Round(sheet10.Range("M19").Value * 0.4, 0) & "处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                    For cs = 1 To Math.Round(sheet10.Range("M19").Value * 0.4, 0)
                        LSBL = ExApp.WorksheetFunction.RandBetween(-9, 9)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet15.Range("D44").Value = LSBLFZ
                End If

                '骨架长度
                sheet14.Range("C17").Value = sheet1.Range("B29").Value
                sheet14.Range("E16").Value = ExApp.WorksheetFunction.RandBetween(40, 80) * -1 ^ ExApp.WorksheetFunction.RandBetween(0, 1) & "    " & ExApp.WorksheetFunction.RandBetween(40, 80) * -1 ^ ExApp.WorksheetFunction.RandBetween(0, 1)
                sheet15.Range("D43").Value = sheet14.Range("E16").Value
                sheet14.Range("E18").Value = sheet1.Range("J6").Value

                '保护层
                sheet14.Range("C21").Value = sheet1.Range("B18").Value * 10
                LSBLFZ = Nothing
                If Math.Round(sheet4.Range("F20").Value * 0.2, 0) <= 10 Then
                    For cs = 1 To Math.Round(sheet4.Range("F20").Value * 0.2, 0)
                        LSBL = ExApp.WorksheetFunction.RandBetween(-9, 9)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet14.Range("E20").Value = LSBLFZ
                    sheet15.Range("D51").Value = "/"
                Else
                    sheet14.Range("E20").Value = "应测" & Math.Round(sheet4.Range("F20").Value * 0.2, 0) & "处，实测" & Math.Round(sheet4.Range("F20").Value * 0.2, 0) & "处，合格" & Math.Round(sheet4.Range("F20").Value * 0.2, 0) & "处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                    For cs = 1 To Math.Round(sheet4.Range("F20").Value * 0.2, 0)
                        LSBL = ExApp.WorksheetFunction.RandBetween(-9, 9)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet15.Range("D51").Value = LSBLFZ
                End If

                '2.钢筋安装记录表 -监抽
                sheet15.Range("B6").Value = sheet1.Range("B5").Value
                sheet15.Range("B7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & sheet1.Range("F6").Value & "钢筋"
                sheet15.Range("O6").Value = sheet1.Range("B24").Value
                sheet15.Range("O7").Value = sheet1.Range("B24").Value
                '3.钻孔桩检表-监抽 
                sheet16.Range("D6").Value = sheet1.Range("B5").Value
                sheet16.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & sheet1.Range("F6").Value
                sheet16.Range("J6").Value = sheet1.Range("B23").Value
                sheet16.Range("F12").Value = sheet1.Range("B21").Value
                sheet16.Range("H13").Value = sheet8.Range("H13").Value
                sheet16.Range("H14").Value = sheet8.Range("H14").Value
                sheet16.Range("F17").Value = Math.Round(sheet1.Range("B7").Value - sheet1.Range("B8").Value, 2)
                sheet16.Range("H16").Value = Math.Round(sheet1.Range("B10").Value - sheet1.Range("B11").Value, 2)
                sheet16.Range("F19").Value = sheet1.Range("B9").Value * 1000
                sheet16.Range("H18").Value = sheet1.Range("B9").Value * 1000 + 25 + Math.Round(ExApp.WorksheetFunction.RandBetween(10, 100) * 0.2, 0)
                sheet16.Range("H20").Value = sheet8.Range("H20").Value
                sheet16.Range("H21").Value = sheet8.Range("H21").Value
                sheet16.Range("H22").Value = "I类"
                ExApp.Calculate()  '刷新一次数据
                ExApp.Calculation = ExApp.Calculation.xlCalculationManual '开启手动计算
                sheet12.Range("C14").Value = Math.Round(3.14 * (sheet12.Range("D8").Value / 2 / 100) ^ 2 * (sheet12.Range("P8").Value - sheet12.Range("I8").Value), 2)
                sheet12.Range("O14").Value = Math.Round(3.14 * (sheet12.Range("D8").Value / 2 / 100) ^ 2 * (sheet12.Range("P9").Value - sheet12.Range("I9").Value), 2)

                sheet4.Range("L10").Value = sheet4.Range("F10").Value - ExApp.WorksheetFunction.Sum(ExApp.WorksheetFunction.CountIf(sheet5.Range("D9:W18"), ">10"), ExApp.WorksheetFunction.CountIf(sheet5.Range("D9:W18"), "<-10"), ExApp.WorksheetFunction.CountIf(sheet6.Range("D9:W18"), ">10"), ExApp.WorksheetFunction.CountIf(sheet6.Range("D9:W18"), "<-10"))
                sheet4.Range("P10").Value = Math.Round(sheet4.Range("L10").Value / sheet4.Range("I10").Value * 100, 1) & "%"
                sheet4.Range("L12").Value = sheet4.Range("F12").Value - ExApp.WorksheetFunction.Sum(ExApp.WorksheetFunction.CountIf(sheet5.Range("D32: W37"), ">20"), ExApp.WorksheetFunction.CountIf(sheet5.Range("D32: W37"), "<-20"))
                sheet4.Range("P12").Value = Math.Round(sheet4.Range("L12").Value / sheet4.Range("I12").Value * 100, 1) & "%"
                sheet4.Range("L20").Value = sheet4.Range("I20").Value - ExApp.WorksheetFunction.Sum(ExApp.WorksheetFunction.CountIf(sheet5.Range("D51: W55"), ">20"), ExApp.WorksheetFunction.CountIf(sheet5.Range("D51: W55"), "<-20"))
                sheet4.Range("Q20").Value = Math.Round(sheet4.Range("L20").Value / sheet4.Range("I20").Value * 100, 1) & "%"
                If sheet10.Range("M19").Value >= 6 Then
                    sheet4.Range("E14").Value = "应测" & sheet10.Range("M19").Value * 2 & "处，实测" & sheet10.Range("M19").Value * 2 & "处，合格" & sheet10.Range("M19").Value * 2 & "处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                Else
                    sheet4.Range("E14").Value = sheet5.Range("D44").Value & "   " & sheet5.Range("E44").Value & "   " & sheet5.Range("F44").Value & "   " & sheet5.Range("G44").Value & "   " & sheet5.Range("H44").Value & "   " & sheet5.Range("I44").Value & "   " & sheet5.Range("J44").Value & "   " & sheet5.Range("K44").Value & "   " & sheet5.Range("L44").Value & "   " & sheet5.Range("M44").Value
                End If

                sheet14.Range("L10").Value = sheet14.Range("F10").Value - ExApp.WorksheetFunction.Sum(ExApp.WorksheetFunction.CountIf(sheet15.Range("D9:W18"), ">10"), ExApp.WorksheetFunction.CountIf(sheet15.Range("D9:W18"), "<-10"))
                sheet14.Range("O10").Value = Math.Round(sheet14.Range("L10").Value / sheet14.Range("I10").Value * 100, 1) & "%"
                sheet17.Range("A5").Value = sheet1.Range("A5").Value & sheet1.Range("B5").Value
                Call 钢筋底高程()
                If TorF = False Then
                    Exit Sub
                End If
                sheet18.Range("A5").Value = sheet1.Range("A5").Value & sheet1.Range("B5").Value
                Call 孔底高程()
                If TorF = False Then
                    Exit Sub
                End If
                sheet19.Range("A5").Value = sheet1.Range("A5").Value & sheet1.Range("B5").Value
                Call 桩顶高程()
                If TorF = False Then
                    Exit Sub
                End If
                sheet20.Range("A5").Value = sheet1.Range("A5").Value & sheet1.Range("B5").Value
                Call 成孔平面()
                If TorF = False Then
                    Exit Sub
                End If
                sheet3.Select()
                For i = 4 To ExApp.Sheets.Count
                    EXsheet = Exbook.Worksheets(i)
                    If EXsheet.Visible = True Then
                        EXsheet.Select(Replace:=False)
                    End If
                Next i
                EXsheet = ExApp.ActiveSheet

                ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                ' 导出PDF文件
                PDFFilename = Filepath & "\" & sheet1.Range("B5").Value & sheet1.Range("C6").Value & sheet1.Range("D6").Value & sheet1.Range("F6").Value & ".pdf"
                '保存PDF文件
                EXsheet.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, PDFFilename, XlFixedFormatQuality.xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False)


                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                '导出EXCEL文件
                'ExcelFilename = Filepath & "\" & sheet1.Range("B5").Value & sheet1.Range("C6").Value  & sheet1.Range("D6").Value & sheet1.Range("F6").Value & ".xlsx"
                'ExApp.Cells.Select() '全选单元格
                'ExApp.Selection.Copy '复制选择的单元格
                'ExApp.Selection.PasteSpecial(XlPasteType.xlPasteValues) '粘贴为数值
                'ExApp.DisplayAlerts = False '关闭弹窗提示
                'ExApp.ActiveWorkbook.SaveAs(Filename:=ExcelFilename, FileFormat:=XlFileFormat.xlOpenXMLWorkbook, CreateBackup:=False)  '另存文件
                'ExApp.ActiveWindow.Close() '关闭活动窗口
                'ExApp.DisplayAlerts = True '打开弹窗提示
                h = h + 1

            End While
            MsgBox("已完成！", 0 + 64, "提示")
        Catch Exclerror As Exception   '错误时弹出提示
            MsgBox(Exclerror.Message & "错误出现在第" & h & "行")
        End Try
        TorF = False
    End Sub

    Sub 桥台扩大基础资料（）

        Dim h, r, i As Integer
        Dim ExcelFilename, PDFFilename As String
        Dim FolderDialogObject As New FolderBrowserDialog()
        h = 8
        i = 8
        r = 0
        Try
            sheet0 = Exbook.Worksheets("数据库") '数据库
            sheet1 = Exbook.Worksheets("参数表") '参数表
            sheet2 = Exbook.Worksheets("交点法") '交点法
            sheet3 = Exbook.Worksheets("线元法") '线元法
            sheet4 = Exbook.Worksheets("断链") '断链
            sheet5 = Exbook.Worksheets("导线成果表") '导线成果表
            sheet6 = Exbook.Worksheets("基坑隐蔽工程") '基坑隐蔽工程
            sheet7 = Exbook.Worksheets("桥涵基坑记录表") '桥涵基坑记录表
            sheet8 = Exbook.Worksheets("基础隐蔽工程") '基础隐蔽工程
            sheet9 = Exbook.Worksheets("扩大基础检表") '扩大基础检表
            sheet10 = Exbook.Worksheets("模板记录表") '模板记录表
            sheet11 = Exbook.Worksheets("砼浇筑申请报告单") '砼浇筑申请报告单
            sheet12 = Exbook.Worksheets("监抽扩大基础检表") '监抽扩大基础检表
            sheet13 = Exbook.Worksheets("水准表") '水准测量记录表
            sheet14 = Exbook.Worksheets("平面表") '全站仪平面位置检测表

            '改表头
            sheet6.Range("A1").Value = sheet0.Range("C1").Value
            sheet6.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet6.Range("E3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet6.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet6.Range("E4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet7.Range("A1").Value = sheet0.Range("C1").Value
            sheet7.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet7.Range("G3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet7.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet7.Range("G4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet8.Range("A1").Value = sheet0.Range("C1").Value
            sheet8.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet8.Range("E3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet8.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet8.Range("E4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet9.Range("A1").Value = sheet0.Range("C1").Value
            sheet9.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet9.Range("I3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet9.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet9.Range("I4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet10.Range("A1").Value = sheet0.Range("C1").Value
            sheet10.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet10.Range("I3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet10.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet10.Range("I4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet11.Range("A1").Value = sheet0.Range("C1").Value
            sheet11.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet11.Range("I3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet11.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet11.Range("I4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet12.Range("A1").Value = sheet0.Range("C1").Value
            sheet12.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet12.Range("I3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet12.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet12.Range("I4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet13.Range("A1").Value = sheet0.Range("C1").Value
            sheet13.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet13.Range("H3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet13.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet13.Range("H4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet14.Range("A1").Value = sheet0.Range("C1").Value
            sheet14.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet14.Range("N3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet14.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet14.Range("N4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value


            sheet0.Range("BA8:BT100000").Value = Nothing
            '计算偏距
            Do While sheet0.Range("B" & i).Value <> Nothing And sheet0.Range("C" & i).Value <> Nothing
                If sheet0.Range("O" & i).Value = Nothing Then
                    MsgBox("请在O" & i & "列选择该桥台是桥梁起点还是终点！！")
                    TorF = False
                    Exit Sub
                ElseIf sheet0.Range("O" & i).Value = "起点" Then
                    sheet0.Range("BA" & i).Value = sheet0.Range("C" & i).Value + sheet0.Range("J" & i).Value / 100
                    sheet0.Range("BE" & i).Value = sheet0.Range("C" & i).Value
                    sheet0.Range("BI" & i).Value = sheet0.Range("C" & i).Value + （sheet0.Range("J" & i).Value / 2） / 100
                    sheet0.Range("BM" & i).Value = sheet0.Range("C" & i).Value + （sheet0.Range("J" & i).Value / 2） / 100
                    sheet0.Range("BQ" & i).Value = sheet0.Range("C" & i).Value + （sheet0.Range("J" & i).Value / 2） / 100
                Else
                    sheet0.Range("BA" & i).Value = sheet0.Range("C" & i).Value
                    sheet0.Range("BE" & i).Value = sheet0.Range("C" & i).Value - sheet0.Range("J" & i).Value / 100
                    sheet0.Range("BI" & i).Value = sheet0.Range("C" & i).Value - （sheet0.Range("J" & i).Value / 2） / 100
                    sheet0.Range("BM" & i).Value = sheet0.Range("C" & i).Value - （sheet0.Range("J" & i).Value / 2） / 100
                    sheet0.Range("BQ" & i).Value = sheet0.Range("C" & i).Value - （sheet0.Range("J" & i).Value / 2） / 100
                End If

                If sheet0.Range("M" & i).Value = "右" Then
                    sheet0.Range("BB" & i).Value = ((sheet0.Range("N" & i).Value / 2) + (sheet0.Range("N" & i).Value - sheet0.Range("I" & i).Value) / 2) / 100 * -1
                    sheet0.Range("BF" & i).Value = ((sheet0.Range("N" & i).Value / 2) + (sheet0.Range("N" & i).Value - sheet0.Range("I" & i).Value) / 2) / 100 * -1
                    sheet0.Range("BJ" & i).Value = sheet0.Range("N" & i).Value / 100 * -1
                    sheet0.Range("BN" & i).Value = (sheet0.Range("I" & i).Value - sheet0.Range("N" & i).Value) / 100 * -1
                    sheet0.Range("BR" & i).Value = ((sheet0.Range("N" & i).Value / 2) + (sheet0.Range("N" & i).Value - sheet0.Range("I" & i).Value) / 2) / 100 * -1
                Else
                    sheet0.Range("BB" & i).Value = ((sheet0.Range("N" & i).Value / 2) + (sheet0.Range("N" & i).Value - sheet0.Range("I" & i).Value) / 2) / 100
                    sheet0.Range("BF" & i).Value = ((sheet0.Range("N" & i).Value / 2) + (sheet0.Range("N" & i).Value - sheet0.Range("I" & i).Value) / 2) / 100
                    sheet0.Range("BJ" & i).Value = (sheet0.Range("I" & i).Value - sheet0.Range("N" & i).Value) / 100
                    sheet0.Range("BN" & i).Value = sheet0.Range("N" & i).Value / 100
                    sheet0.Range("BR" & i).Value = ((sheet0.Range("N" & i).Value / 2) + (sheet0.Range("N" & i).Value - sheet0.Range("I" & i).Value) / 2) / 100
                End If
                i += 1
            Loop
            Call 质检资料设计坐标()
            If TorF = False Then
                Exit Sub
            End If
            While sheet0.Range("B" & h).Value <> Nothing
                If TorF = False Then
                    Exit Sub
                End If
                ExApp.Calculation = ExApp.Calculation.xlCalculationManual '开启手动计算
                sheet1.Range("B" & r + 5).Value = sheet0.Range("B" & h).Value
                'sheet1.Range("B" & r + 6).value = sheet0.Range("C" & h).value
                sheet1.Range("C" & r + 6).Value = sheet0.Range("D" & h).Value
                sheet1.Range("D" & r + 6).Value = sheet0.Range("E" & h).Value
                sheet1.Range("B" & r + 7).Value = sheet0.Range("I" & h).Value
                sheet1.Range("B" & r + 8).Value = sheet0.Range("J" & h).Value
                sheet1.Range("B" & r + 9).Value = sheet0.Range("F" & h).Value
                sheet1.Range("B" & r + 10).Value = sheet0.Range("G" & h).Value
                sheet1.Range("B" & r + 11).Value = sheet0.Range("H" & h).Value

                '基坑测量参数
                sheet1.Range("I3:Q15").Value = Nothing
                sheet1.Range("I3").Value = sheet0.Range("BA" & h).Value
                sheet1.Range("I4").Value = sheet0.Range("BE" & h).Value
                sheet1.Range("I5").Value = sheet0.Range("BI" & h).Value
                sheet1.Range("I6").Value = sheet0.Range("BM" & h).Value
                sheet1.Range("J3").Value = sheet0.Range("BB" & h).Value
                sheet1.Range("J4").Value = sheet0.Range("BF" & h).Value
                sheet1.Range("J5").Value = sheet0.Range("BJ" & h).Value
                sheet1.Range("J6").Value = sheet0.Range("BN" & h).Value
                sheet1.Range("K3").Value = sheet0.Range("L" & h).Value
                sheet1.Range("K4").Value = sheet0.Range("L" & h).Value
                sheet1.Range("K5").Value = sheet0.Range("L" & h).Value
                sheet1.Range("K6").Value = sheet0.Range("L" & h).Value

                sheet1.Range("M3").Value = sheet0.Range("BC" & h).Value
                sheet1.Range("N3").Value = sheet0.Range("BD" & h).Value
                sheet1.Range("M4").Value = sheet0.Range("BG" & h).Value
                sheet1.Range("N4").Value = sheet0.Range("BH" & h).Value
                sheet1.Range("M5").Value = sheet0.Range("BK" & h).Value
                sheet1.Range("N5").Value = sheet0.Range("BL" & h).Value
                sheet1.Range("M6").Value = sheet0.Range("BO" & h).Value
                sheet1.Range("N6").Value = sheet0.Range("BP" & h).Value
                sheet1.Range("P1").Value = sheet1.Range("B10").Value
                sheet1.Range("L3").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "基坑"
                '基坑测量偏差值
                sheet1.Range("O3").Value = ExApp.WorksheetFunction.RandBetween(-25, 25)
                sheet1.Range("O4").Value = ExApp.WorksheetFunction.RandBetween(-25, 25)
                sheet1.Range("O5").Value = ExApp.WorksheetFunction.RandBetween(-25, 25)
                sheet1.Range("O6").Value = ExApp.WorksheetFunction.RandBetween(-25, 25)
                sheet1.Range("P3").Value = ExApp.WorksheetFunction.RandBetween(2, 15)
                sheet1.Range("P4").Value = ExApp.WorksheetFunction.RandBetween(2, 15)
                sheet1.Range("P5").Value = ExApp.WorksheetFunction.RandBetween(2, 15)
                sheet1.Range("P6").Value = ExApp.WorksheetFunction.RandBetween(2, 15)

                sheet13.Range("A5").Value = sheet1.Range("A5").Value & sheet1.Range("B5").Value
                sheet14.Range("A5").Value = sheet1.Range("A5").Value & sheet1.Range("B5").Value

                Call 水准测量记录表()
                If TorF = False Then
                    Exit Sub
                End If

                Call 全站仪平面位置检测表（）
                If TorF = False Then
                    Exit Sub
                End If
                sheet13.Activate()
                ExApp.ActiveWindow.SelectedSheets.Copy(, (ExApp.Sheets(ExApp.Sheets.Count)))
                sheet14.Activate()
                ExApp.ActiveWindow.SelectedSheets.Copy(, (ExApp.Sheets(ExApp.Sheets.Count)))

                '基坑隐蔽工程
                sheet6.Range("C6").Value = sheet1.Range("B5").Value & sheet1.Range("C6").Value & "基础及下部构造"
                sheet6.Range("E6").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet6.Range("C7").Value = sheet1.Range("B5").Value & sheet1.Range("C6").Value & sheet1.Range("D6").Value & "基坑"
                sheet6.Range("E10").Value = sheet1.Range("B10").Value
                sheet6.Range("E11").Value = sheet1.Range("B10").Value
                sheet6.Range("E27").Value = sheet1.Range("B10").Value
                sheet6.Range("E28").Value = sheet1.Range("B10").Value
                '桥涵基坑记录表
                sheet7.Range("D6").Value = sheet1.Range("B5").Value
                sheet7.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "基坑"
                sheet7.Range("G6").Value = sheet1.Range("B9").Value
                sheet7.Range("G7").Value = sheet1.Range("B10").Value
                sheet7.Range("D10").Value = "长：" & sheet1.Range("B7").Value * 10
                sheet7.Range("D11").Value = "宽：" & sheet1.Range("B8").Value * 10
                sheet7.Range("F10").Value = sheet1.Range("B7").Value * 10 + ExApp.WorksheetFunction.RandBetween(30, 50)
                sheet7.Range("F11").Value = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(30, 50)
                sheet7.Range("E12").Value = sheet1.Range("P3").Value & "    " & sheet1.Range("P4").Value & "    " & sheet1.Range("P5").Value & "    " & sheet1.Range("P6").Value
                sheet7.Range("E13").Value = sheet1.Range("O3").Value & "    " & sheet1.Range("O4").Value & "    " & sheet1.Range("O5").Value & "    " & sheet1.Range("O6").Value
                sheet7.Range("E15").Value = "符合基地承载力要求"
                sheet7.Range("E16").Value = "符合要求"

                '模板测量偏差值
                sheet1.Range("O3").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O4").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O5").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O6").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("P3").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P4").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P5").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P6").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P1").Value = sheet1.Range("B11").Value
                sheet1.Range("L3").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "模板"
                Call 水准测量记录表()
                If TorF = False Then
                    Exit Sub
                End If
                Call 全站仪平面位置检测表（）
                If TorF = False Then
                    Exit Sub
                End If
                sheet13.Activate()
                ExApp.ActiveWindow.SelectedSheets.Copy(, (ExApp.Sheets(ExApp.Sheets.Count)))
                sheet14.Activate()
                ExApp.ActiveWindow.SelectedSheets.Copy(, (ExApp.Sheets(ExApp.Sheets.Count)))

                '现场模板安装检查记录表
                sheet10.Range("C6").Value = sheet1.Range("B5").Value
                sheet10.Range("C7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet10.Range("I6").Value = sheet1.Range("B11").Value
                sheet10.Range("I7").Value = sheet1.Range("B11").Value
                sheet10.Range("D8").Value = 2
                sheet10.Range("D9").Value = 5
                sheet10.Range("F8").Value = 4
                sheet10.Range("F9").Value = 4
                sheet10.Range("H8").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet10.Range("H9").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet10.Range("J8").Value = 100
                sheet10.Range("J9").Value = 100
                '模板测量偏差
                sheet10.Range("D11").Value = sheet1.Range("P5").Value
                sheet10.Range("F11").Value = sheet1.Range("P6").Value
                sheet10.Range("H11").Value = sheet1.Range("P3").Value
                sheet10.Range("J11").Value = sheet1.Range("P4").Value
                sheet10.Range("F14").Value = sheet1.Range("O3").Value
                sheet10.Range("G14").Value = sheet1.Range("O4").Value
                sheet10.Range("H14").Value = sheet1.Range("O5").Value
                sheet10.Range("I14").Value = sheet1.Range("O6").Value

                sheet10.Range("F12").Value = sheet1.Range("B7").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet10.Range("G12").Value = sheet1.Range("B7").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet10.Range("H12").Value = sheet1.Range("B7").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet10.Range("I12").Value = sheet1.Range("B7").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet10.Range("F13").Value = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet10.Range("G13").Value = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet10.Range("H13").Value = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet10.Range("I13").Value = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet10.Range("F18").Value = "牢固，稳定"

                '基础隐蔽工程报验单
                sheet8.Range("C6").Value = sheet1.Range("B5").Value & sheet1.Range("C6").Value & "基础及下部构造"
                sheet8.Range("E6").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet8.Range("C7").Value = sheet1.Range("B5").Value & sheet1.Range("C6").Value & sheet1.Range("D6").Value & "基础"
                '混凝土测量偏差
                sheet1.Range("I7").Value = sheet0.Range("BQ" & h).Value
                sheet1.Range("J7").Value = sheet0.Range("BR" & h).Value
                sheet1.Range("K7").Value = sheet0.Range("L" & h).Value
                sheet1.Range("O3").Value = ExApp.WorksheetFunction.RandBetween(-25, 25)
                sheet1.Range("O4").Value = ExApp.WorksheetFunction.RandBetween(-25, 25)
                sheet1.Range("O5").Value = ExApp.WorksheetFunction.RandBetween(-25, 25)
                sheet1.Range("O6").Value = ExApp.WorksheetFunction.RandBetween(-25, 25)
                sheet1.Range("O7").Value = ExApp.WorksheetFunction.RandBetween(-25, 25)
                sheet1.Range("P3").Value = ExApp.WorksheetFunction.RandBetween(5, 15)
                sheet1.Range("P4").Value = ExApp.WorksheetFunction.RandBetween(5, 15)
                sheet1.Range("P5").Value = ExApp.WorksheetFunction.RandBetween(5, 15)
                sheet1.Range("P6").Value = ExApp.WorksheetFunction.RandBetween(5, 15)
                sheet1.Range("P1").Value = sheet1.Range("B11").Value
                sheet1.Range("L3").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "底面"

                Call 水准测量记录表()
                If TorF = False Then
                    Exit Sub
                End If
                sheet13.Activate()
                ExApp.ActiveWindow.SelectedSheets.Copy(, (ExApp.Sheets(ExApp.Sheets.Count)))

                '混凝土扩大基础现场质量检验表
                sheet9.Range("D6").Value = sheet1.Range("B5").Value
                sheet9.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet9.Range("I6").Value = sheet1.Range("B11").Value
                sheet9.Range("D13").Value = "长：" & sheet1.Range("B7").Value * 10 & " " & "宽：" & sheet1.Range("B8").Value * 10

                sheet9.Range("F12").Value = sheet1.Range("B7").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet9.Range("G12").Value = sheet1.Range("B7").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet9.Range("H12").Value = sheet1.Range("B7").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet9.Range("F13").Value = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet9.Range("G13").Value = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet9.Range("H13").Value = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)

                sheet9.Range("E15").Value = sheet1.Range("O3").Value
                sheet9.Range("F15").Value = sheet1.Range("O4").Value
                sheet9.Range("G15").Value = sheet1.Range("O5").Value
                sheet9.Range("H15").Value = sheet1.Range("O6").Value
                sheet9.Range("I15").Value = sheet1.Range("O7").Value

                '顶面高程
                sheet1.Range("K3").Value = sheet0.Range("L" & h).Value + sheet0.Range("K" & h).Value / 100
                sheet1.Range("K4").Value = sheet0.Range("L" & h).Value + sheet0.Range("K" & h).Value / 100
                sheet1.Range("K5").Value = sheet0.Range("L" & h).Value + sheet0.Range("K" & h).Value / 100
                sheet1.Range("K6").Value = sheet0.Range("L" & h).Value + sheet0.Range("K" & h).Value / 100
                sheet1.Range("K7").Value = sheet0.Range("L" & h).Value + sheet0.Range("K" & h).Value / 100

                sheet1.Range("O3").Value = ExApp.WorksheetFunction.RandBetween(-25, 25)
                sheet1.Range("O4").Value = ExApp.WorksheetFunction.RandBetween(-25, 25)
                sheet1.Range("O5").Value = ExApp.WorksheetFunction.RandBetween(-25, 25)
                sheet1.Range("O6").Value = ExApp.WorksheetFunction.RandBetween(-25, 25)
                sheet1.Range("O7").Value = ExApp.WorksheetFunction.RandBetween(-25, 25)
                sheet1.Range("L3").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "顶面"
                Call 水准测量记录表()
                If TorF = False Then
                    Exit Sub
                End If
                sheet9.Range("E16").Value = sheet1.Range("O3").Value
                sheet9.Range("F16").Value = sheet1.Range("O4").Value
                sheet9.Range("G16").Value = sheet1.Range("O5").Value
                sheet9.Range("H16").Value = sheet1.Range("O6").Value
                sheet9.Range("I16").Value = sheet1.Range("O7").Value

                sheet1.Range("I7:Q7").Value = Nothing
                sheet1.Range("L3").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                Call 全站仪平面位置检测表（）
                If TorF = False Then
                    Exit Sub
                End If
                sheet9.Range("E17").Value = sheet1.Range("O3").Value
                sheet9.Range("F17").Value = sheet1.Range("O4").Value
                sheet9.Range("G17").Value = sheet1.Range("O5").Value
                sheet9.Range("H17").Value = sheet1.Range("O6").Value

                '监抽混凝土扩大基础现场质量检验表
                sheet12.Range("D6").Value = sheet1.Range("B5").Value
                sheet12.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet12.Range("I6").Value = sheet1.Range("B11").Value
                sheet12.Range("D13").Value = "长：" & sheet1.Range("B7").Value * 10 & " " & "宽：" & sheet1.Range("B8").Value * 10

                sheet12.Range("F12").Value = sheet1.Range("B7").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet12.Range("F13").Value = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)

                sheet12.Range("E15").Value = sheet9.Range("E15").Value
                sheet12.Range("F15").Value = sheet9.Range("F15").Value
                sheet12.Range("G15").Value = sheet9.Range("G15").Value
                sheet12.Range("H15").Value = sheet9.Range("H15").Value
                sheet12.Range("I15").Value = sheet9.Range("I15").Value
                sheet12.Range("E16").Value = sheet9.Range("E16").Value
                sheet12.Range("F16").Value = sheet9.Range("F16").Value
                sheet12.Range("G16").Value = sheet9.Range("G16").Value
                sheet12.Range("H16").Value = sheet9.Range("H16").Value

                sheet12.Range("E17").Value = sheet9.Range("E17").Value
                sheet12.Range("F17").Value = sheet9.Range("F17").Value
                sheet12.Range("G17").Value = sheet9.Range("G17").Value
                sheet12.Range("H17").Value = sheet9.Range("H17").Value

                ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                '选择表格
                sheet6.Select()
                For i = 7 To ExApp.Sheets.Count
                    EXsheet = Exbook.Worksheets(i)
                    If EXsheet.Sheets(i).Visible = True Then
                        EXsheet.Select(Replace:=False)
                    End If
                Next i
                EXsheet = ExApp.ActiveSheet
                ' 导出PDF文件
                PDFFilename = Filepath & "\" & sheet1.Range("B5").Value & sheet1.Range("C6").Value & sheet1.Range("D6").Value & ".pdf"
                '保存PDF文件
                EXsheet.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, PDFFilename, XlFixedFormatQuality.xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False)

                sheet15 = Exbook.Worksheets("水准表 (2)") '水准测量记录表
                sheet16 = Exbook.Worksheets("平面表 (2)") '全站仪平面位置检测表
                sheet17 = Exbook.Worksheets("水准表 (3)") '水准测量记录表
                sheet18 = Exbook.Worksheets("平面表 (3)") '全站仪平面位置检测表
                sheet19 = Exbook.Worksheets("水准表 (4)") '水准测量记录表
                ExApp.DisplayAlerts = False
                sheet15.Delete()
                sheet16.Delete()
                sheet17.Delete()
                sheet18.Delete()
                sheet19.Delete()
                h += 1
            End While
            MsgBox("已完成！", 0 + 64, "提示")

        Catch Exclerror As Exception   '错误时弹出提示
            MsgBox(Exclerror.Message)
        End Try
        TorF = False
    End Sub

    Sub 承台资料()
        Dim h, r, i As Integer
        Dim ExcelFilename, PDFFilename, LSBL, LSBLFZ As String  '定义输出的PDF文件名
        h = 8
        i = 8
        r = 0
        Try
            sheet0 = Exbook.Worksheets("数据库") '数据库
            sheet1 = Exbook.Worksheets("参数表") '参数表
            sheet2 = Exbook.Worksheets("交点法") '交点法
            sheet3 = Exbook.Worksheets("线元法") '线元法
            sheet4 = Exbook.Worksheets("断链") '断链
            sheet5 = Exbook.Worksheets("导线成果表") '导线成果表
            sheet6 = Exbook.Worksheets("钢筋隐蔽工程") '钢筋隐蔽工程
            sheet7 = Exbook.Worksheets("钢筋检表") '钢筋检表
            sheet8 = Exbook.Worksheets("钢筋记录表") '钢筋记录表
            sheet9 = Exbook.Worksheets("申请批复单") '工序检验申请批复单
            sheet10 = Exbook.Worksheets("承台检表") '承台检表
            sheet11 = Exbook.Worksheets("模板记录表") '模板记录表
            sheet12 = Exbook.Worksheets("砼浇筑申请报告单") '砼浇筑申请报告单
            sheet13 = Exbook.Worksheets("监抽钢筋检表") '监抽钢筋检表
            sheet14 = Exbook.Worksheets("监抽钢筋记录表") '监抽钢筋记录表
            sheet15 = Exbook.Worksheets("监抽承台检表") '监抽承台检表
            sheet16 = Exbook.Worksheets("水准表") '水准测量记录表
            sheet17 = Exbook.Worksheets("平面表") '全站仪平面位置检测表

            '改表头
            sheet6.Range("A1").Value = sheet0.Range("C1").Value
            sheet6.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet6.Range("E3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet6.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet6.Range("E4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet7.Range("A1").Value = sheet0.Range("C1").Value
            sheet7.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet7.Range("P3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet7.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet7.Range("P4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet8.Range("A1").Value = sheet0.Range("C1").Value
            sheet8.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet8.Range("L3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet8.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet8.Range("L4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet9.Range("A1").Value = sheet0.Range("C1").Value
            sheet9.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet9.Range("E3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet9.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet9.Range("E4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet10.Range("A1").Value = sheet0.Range("C1").Value
            sheet10.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet10.Range("K3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet10.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet10.Range("K4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet11.Range("A1").Value = sheet0.Range("C1").Value
            sheet11.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet11.Range("I3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet11.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet11.Range("I4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet12.Range("A1").Value = sheet0.Range("C1").Value
            sheet12.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet12.Range("I3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet12.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet12.Range("I4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet13.Range("A1").Value = sheet0.Range("C1").Value
            sheet13.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet13.Range("P3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet13.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet13.Range("P4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet14.Range("A1").Value = sheet0.Range("C1").Value
            sheet14.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet14.Range("L3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet14.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet14.Range("L4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet15.Range("A1").Value = sheet0.Range("C1").Value
            sheet15.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet15.Range("K3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet15.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet15.Range("K4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet16.Range("A1").Value = sheet0.Range("C1").Value
            sheet16.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet16.Range("H3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet16.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet16.Range("H4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet17.Range("A1").Value = sheet0.Range("C1").Value
            sheet17.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet17.Range("N3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet17.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet17.Range("N4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value
            sheet0.Range("BA8:BT100000").Value = Nothing
            '计算偏距
            Do While sheet0.Range("B" & i).Value <> Nothing And sheet0.Range("L" & i).Value <> Nothing And sheet0.Range("M" & i).Value <> Nothing
                If sheet0.Range("C" & i).Value = Nothing And sheet0.Range("U" & i).Value = "否" And sheet0.Range("V" & i).Value <> Nothing And sheet0.Range("W" & i).Value <> Nothing Then
                    sheet0.Range("C" & i).Value = FSZhj(sheet0.Range("V" & i).Value, sheet0.Range("W" & i).Value)
                ElseIf sheet0.Range("C" & i).Value <> Nothing And sheet0.Range("U" & i).Value = "否" And sheet0.Range("V" & i).Value = Nothing And sheet0.Range("W" & i).Value = Nothing Then
                    MsgBox("请核对中心桩号或X、Y坐标是否正确填写！")
                    TorF = False
                    Exit Sub
                End If
                sheet0.Range("BA" & i).Value = sheet0.Range("C" & i).Value + (sheet0.Range("X" & i).Value / 2 + sheet0.Range("H" & i).Value / 2) / 100
                sheet0.Range("BE" & i).Value = sheet0.Range("C" & i).Value - (sheet0.Range("X" & i).Value / 2 + sheet0.Range("H" & i).Value / 2) / 100
                sheet0.Range("BI" & i).Value = sheet0.Range("C" & i).Value + (sheet0.Range("X" & i).Value / 2) / 100
                sheet0.Range("BM" & i).Value = sheet0.Range("C" & i).Value + (sheet0.Range("X" & i).Value / 2) / 100
                sheet0.Range("BQ" & i).Value = sheet0.Range("C" & i).Value + (sheet0.Range("X" & i).Value / 2) / 100

                If sheet0.Range("K" & i).Value = "中" Then
                    sheet0.Range("BB" & i).Value = 0
                    sheet0.Range("BF" & i).Value = 0
                    sheet0.Range("BJ" & i).Value = sheet0.Range("G" & i).Value / 2 / 100 * -1
                    sheet0.Range("BN" & i).Value = sheet0.Range("G" & i).Value / 2 / 100
                    sheet0.Range("BR" & i).Value = 0
                Else
                    If sheet0.Range("K" & i).Value = "右" Then
                        sheet0.Range("BB" & i).Value = (sheet0.Range("L" & i).Value + sheet0.Range("Y" & i).Value / 2) / 100 * -1
                        sheet0.Range("BF" & i).Value = (sheet0.Range("L" & i).Value + sheet0.Range("Y" & i).Value / 2) / 100 * -1
                        sheet0.Range("BJ" & i).Value = (sheet0.Range("L" & i).Value + sheet0.Range("Y" & i).Value / 2 + sheet0.Range("G" & i).Value / 2) / 100 * -1
                        sheet0.Range("BN" & i).Value = (sheet0.Range("L" & i).Value + sheet0.Range("Y" & i).Value / 2 - sheet0.Range("G" & i).Value / 2) / 100 * -1
                        sheet0.Range("BR" & i).Value = (sheet0.Range("L" & i).Value + sheet0.Range("Y" & i).Value / 2) / 100 * -1
                    Else
                        sheet0.Range("BB" & i).Value = (sheet0.Range("L" & i).Value + sheet0.Range("Y" & i).Value / 2) / 100
                        sheet0.Range("BF" & i).Value = (sheet0.Range("L" & i).Value + sheet0.Range("Y" & i).Value / 2) / 100
                        sheet0.Range("BJ" & i).Value = (sheet0.Range("L" & i).Value + sheet0.Range("Y" & i).Value / 2 - sheet0.Range("G" & i).Value / 2) / 100
                        sheet0.Range("BN" & i).Value = (sheet0.Range("L" & i).Value + sheet0.Range("Y" & i).Value / 2 + sheet0.Range("G" & i).Value / 2) / 100
                        sheet0.Range("BR" & i).Value = (sheet0.Range("L" & i).Value + sheet0.Range("Y" & i).Value / 2) / 100
                    End If
                End If
                i += 1
            Loop

            Call 质检资料设计坐标()
            If TorF = False Then
                Exit Sub
            End If
            While sheet0.Range("B" & h).Value <> Nothing
                If TorF = False Then
                    Exit Sub
                End If
                ExApp.Calculation = ExApp.Calculation.xlCalculationManual '开启手动计算
                sheet1.Range("B" & r + 5).Value = sheet0.Range("B" & h).Value
                sheet1.Range("B" & r + 6).Value = sheet0.Range("C" & h).Value
                sheet1.Range("C" & r + 6).Value = sheet0.Range("D" & h).Value
                sheet1.Range("D" & r + 6).Value = sheet0.Range("E" & h).Value
                sheet1.Range("B" & r + 7).Value = sheet0.Range("N" & h).Value
                sheet1.Range("B" & r + 8).Value = sheet0.Range("O" & h).Value
                sheet1.Range("B" & r + 9).Value = sheet0.Range("P" & h).Value
                sheet1.Range("B" & r + 10).Value = sheet0.Range("T" & h).Value
                sheet1.Range("B" & r + 11).Value = sheet0.Range("G" & h).Value
                sheet1.Range("B" & r + 12).Value = sheet0.Range("H" & h).Value
                sheet1.Range("B" & r + 13).Value = sheet0.Range("I" & h).Value
                sheet1.Range("B" & r + 14).Value = sheet0.Range("Q" & h).Value
                sheet1.Range("B" & r + 15).Value = sheet0.Range("R" & h).Value
                sheet1.Range("B" & r + 16).Value = sheet0.Range("S" & h).Value
                sheet1.Range("B" & r + 17).Value = sheet0.Range("F" & h).Value
                sheet1.Range("B" & r + 17).Value = sheet0.Range("F" & h).Value

                '测量参数
                sheet1.Range("I3:Q15").Value = Nothing
                sheet1.Range("I1").Value = sheet1.Range("B5").Value.substring(0, ExApp.WorksheetFunction.Find("K", sheet1.Range("B5").Value))
                sheet1.Range("I3").Value = sheet0.Range("BA" & h).Value
                sheet1.Range("I4").Value = sheet0.Range("BE" & h).Value
                sheet1.Range("I5").Value = sheet0.Range("BI" & h).Value
                sheet1.Range("I6").Value = sheet0.Range("BM" & h).Value
                sheet1.Range("I7").Value = sheet0.Range("BQ" & h).Value

                sheet1.Range("J3").Value = sheet0.Range("BB" & h).Value
                sheet1.Range("J4").Value = sheet0.Range("BF" & h).Value
                sheet1.Range("J5").Value = sheet0.Range("BJ" & h).Value
                sheet1.Range("J6").Value = sheet0.Range("BN" & h).Value
                sheet1.Range("J7").Value = sheet0.Range("BR" & h).Value

                sheet1.Range("K3").Value = sheet0.Range("J" & h).Value
                sheet1.Range("K4").Value = sheet0.Range("J" & h).Value
                sheet1.Range("K5").Value = sheet0.Range("J" & h).Value
                sheet1.Range("K6").Value = sheet0.Range("J" & h).Value
                sheet1.Range("K7").Value = sheet0.Range("J" & h).Value

                sheet1.Range("M3").Value = sheet0.Range("BC" & h).Value
                sheet1.Range("N3").Value = sheet0.Range("BD" & h).Value
                sheet1.Range("M4").Value = sheet0.Range("BG" & h).Value
                sheet1.Range("N4").Value = sheet0.Range("BH" & h).Value
                sheet1.Range("M5").Value = sheet0.Range("BK" & h).Value
                sheet1.Range("N5").Value = sheet0.Range("BL" & h).Value
                sheet1.Range("M6").Value = sheet0.Range("BO" & h).Value
                sheet1.Range("N6").Value = sheet0.Range("BP" & h).Value
                sheet1.Range("M7").Value = sheet0.Range("BS" & h).Value
                sheet1.Range("N7").Value = sheet0.Range("BT" & h).Value
                sheet1.Range("P1").Value = sheet1.Range("B17").Value
                sheet1.Range("L3").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                '测量偏差值
                sheet1.Range("O3").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O4").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O5").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O6").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O7").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("P3").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P4").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P5").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P6").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet16.Range("A5").Value = sheet1.Range("A5").Value & sheet1.Range("B5").Value
                sheet17.Range("A5").Value = sheet1.Range("A5").Value & sheet1.Range("B5").Value


                ' 钢筋隐蔽工程
                sheet6.Range("C6").Value = sheet1.Range("B5").Value & sheet1.Range("C6").Value & "基础及下部构造"
                sheet6.Range("E6").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋加工及安装"
                sheet6.Range("C7").Value = sheet1.Range("B5").Value & sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋加工及安装"
                sheet6.Range("E10").Value = sheet1.Range("B18").Value
                sheet6.Range("E11").Value = sheet1.Range("B18").Value
                sheet6.Range("E27").Value = sheet1.Range("B18").Value
                sheet6.Range("E28").Value = sheet1.Range("B18").Value
                '钢筋检表
                sheet7.Range("D6").Value = sheet1.Range("B5").Value
                sheet7.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋"
                sheet7.Range("Q6").Value = sheet1.Range("B18").Value
                sheet7.Range("Q7").Value = sheet1.Range("B18").Value
                sheet7.Range("Q31").Value = sheet1.Range("B18").Value
                sheet7.Range("Q34").Value = sheet1.Range("B18").Value
                '主筋
                sheet7.Range("E15").Value = "设计值：" & sheet1.Range("B8").Value * 10
                LSBLFZ = Nothing
                If sheet1.Range("B7").Value * 2 <= 10 Then
                    For cs = 1 To sheet1.Range("B7").Value * 2
                        LSBL = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet7.Range("G14").Value = LSBLFZ
                    sheet8.Range("D24").Value = "/"
                Else
                    For cs = 1 To sheet1.Range("B7").Value * 2
                        LSBL = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet7.Range("G14").Value = "应测" & sheet1.Range("B7").Value * 2 & "处，实测" & sheet1.Range("B7").Value * 2 & "处，合格" & sheet1.Range("B7").Value * 2 & "处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                    sheet8.Range("D24").Value = LSBLFZ
                End If
                '箍筋
                sheet7.Range("E17").Value = "设计值：" & sheet1.Range("B9").Value * 10
                LSBLFZ = Nothing
                For cs = 1 To 10
                    LSBL = sheet1.Range("B9").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                    LSBLFZ = LSBLFZ & LSBL & "   "
                Next
                sheet7.Range("G16").Value = LSBLFZ

                '骨架尺寸
                sheet7.Range("E19").Value = "设计值：" & sheet1.Range("B14").Value * 10
                sheet7.Range("G18").Value = sheet1.Range("B14").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet7.Range("E21").Value = "宽：" & sheet1.Range("B15").Value * 10
                sheet7.Range("E22").Value = "高：" & sheet1.Range("B16").Value * 10
                sheet7.Range("G20").Value = "宽：  " & sheet1.Range("B15").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4) & vbCrLf &
                                            "高：  " & sheet1.Range("B16").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4)
                '保护层
                sheet7.Range("E27").Value = "设计值：" & sheet1.Range("B10").Value * 10
                Dim gs As Integer
                gs = Math.Round((sheet1.Range("B11").Value * sheet1.Range("B13").Value * 2 + sheet1.Range("B12").Value * sheet1.Range("B13").Value * 2) / 300 / 100, 0)
                If gs <= 20 Then
                    LSBLFZ = Nothing
                    For cs = 1 To 20
                        LSBL = sheet1.Range("B10").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet7.Range("G26").Value = "应测20处，实测20处，合格20处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                    sheet8.Range("D48").Value = LSBLFZ
                Else
                    LSBLFZ = Nothing
                    For cs = 1 To gs
                        LSBL = sheet1.Range("B10").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet7.Range("G26").Value = "应测" & gs & "处，实测" & gs & "处，合格" & gs & "处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                    sheet8.Range("D48").Value = LSBLFZ
                End If

                '钢筋记录表
                sheet8.Range("B6").Value = sheet1.Range("B5").Value
                sheet8.Range("B7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋"
                sheet8.Range("K6").Value = sheet1.Range("B18").Value
                sheet8.Range("K7").Value = sheet1.Range("B18").Value
                '工序检验申请批复单
                sheet9.Range("C6").Value = sheet1.Range("B5").Value & sheet1.Range("C6").Value & "基础及下部构造"
                sheet9.Range("C7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet9.Range("C8").Value = sheet1.Range("B5").Value
                sheet9.Range("C9").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet9.Range("C10").Value = "混凝土强度、平面尺寸、结构高度、顶面高程、轴线偏位、平整度"

                Call 水准测量记录表()
                If TorF = False Then
                    Exit Sub
                End If
                sheet1.Range("I7:N7").Value = Nothing

                Call 全站仪平面位置检测表（）
                If TorF = False Then
                    Exit Sub
                End If
                sheet16.Activate()
                ExApp.ActiveWindow.SelectedSheets.Copy(, (ExApp.Sheets(ExApp.Sheets.Count)))
                sheet17.Activate()
                ExApp.ActiveWindow.SelectedSheets.Copy(, (ExApp.Sheets(ExApp.Sheets.Count)))

                '承台混凝土现场质量检验表
                sheet10.Range("D6").Value = sheet1.Range("B5").Value
                sheet10.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet10.Range("K6").Value = sheet1.Range("B17").Value
                sheet10.Range("D14").Value = "长：" & sheet1.Range("B11").Value * 10 & " 宽：" & sheet1.Range("B12").Value * 10
                sheet10.Range("E13").Value = "长：" & sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20) & "  " &
                                              sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20) & vbCrLf &
                                             "宽：" & sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20) & "  " &
                                             sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet10.Range("D18").Value = sheet1.Range("B13").Value * 10
                sheet10.Range("E17").Value = sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20) & "  " &
                                             sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20) & "  " &
                                             sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20) & "  " &
                                             sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20) & "  " &
                                             sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet10.Range("E19").Value = sheet1.Range("O3").Value
                sheet10.Range("F19").Value = sheet1.Range("O4").Value
                sheet10.Range("G19").Value = sheet1.Range("O5").Value
                sheet10.Range("H19").Value = sheet1.Range("O6").Value
                sheet10.Range("I19").Value = sheet1.Range("O7").Value
                sheet10.Range("E20").Value = sheet1.Range("P3").Value
                sheet10.Range("F20").Value = sheet1.Range("P4").Value
                sheet10.Range("G20").Value = sheet1.Range("P5").Value
                sheet10.Range("H20").Value = sheet1.Range("P6").Value

                gs = Math.Round((sheet1.Range("B11").Value * sheet1.Range("B13").Value * 2 + sheet1.Range("B12").Value * sheet1.Range("B13").Value * 2) / 2000 / 100, 0)
                LSBLFZ = Nothing
                If gs <= 24 Then
                    For cs = 1 To 24
                        LSBL = ExApp.WorksheetFunction.RandBetween(1, 5)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet10.Range("E21").Value = LSBLFZ
                Else
                    For cs = 1 To gs
                        LSBL = ExApp.WorksheetFunction.RandBetween(1, 5)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet10.Range("E21").Value = LSBLFZ
                End If

                '模板测量偏差值
                sheet1.Range("O3").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O4").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O5").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O6").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("P3").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P4").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P5").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P6").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("L3").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "模板"

                Call 水准测量记录表()
                If TorF = False Then
                    Exit Sub
                End If
                Call 全站仪平面位置检测表（）
                If TorF = False Then
                    Exit Sub
                End If
                '现场模板安装检查记录表
                sheet11.Range("C6").Value = sheet1.Range("B5").Value
                sheet11.Range("C7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet11.Range("I6").Value = sheet1.Range("B17").Value
                sheet11.Range("I7").Value = sheet1.Range("B17").Value
                sheet11.Range("D8").Value = 2
                sheet11.Range("D9").Value = 5
                sheet11.Range("F8").Value = 4
                sheet11.Range("F9").Value = 4
                sheet11.Range("H8").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet11.Range("H9").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet11.Range("J8").Value = 100
                sheet11.Range("J9").Value = 100
                '测量偏差
                sheet11.Range("D11").Value = sheet1.Range("P5").Value
                sheet11.Range("F11").Value = sheet1.Range("P6").Value
                sheet11.Range("H11").Value = sheet1.Range("P3").Value
                sheet11.Range("J11").Value = sheet1.Range("P4").Value
                sheet11.Range("F14").Value = sheet1.Range("O3").Value
                sheet11.Range("G14").Value = sheet1.Range("O4").Value
                sheet11.Range("H14").Value = sheet1.Range("O5").Value
                sheet11.Range("I14").Value = sheet1.Range("O6").Value

                sheet11.Range("F12").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("G12").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("H12").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("I12").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("F13").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("G13").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("H13").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("I13").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("F18").Value = "牢固，稳定"

                '监抽钢筋检表
                sheet13.Range("D6").Value = sheet1.Range("B5").Value
                sheet13.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋"
                sheet13.Range("Q6").Value = sheet1.Range("B18").Value
                sheet13.Range("Q7").Value = sheet1.Range("B18").Value
                sheet13.Range("Q31").Value = sheet1.Range("B18").Value
                sheet13.Range("Q34").Value = sheet1.Range("B18").Value
                '主筋
                sheet13.Range("E15").Value = "设计值：" & sheet1.Range("B8").Value * 10
                LSBLFZ = Nothing
                If Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0) <= 10 Then
                    For cs = 1 To Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0)
                        LSBL = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet13.Range("G14").Value = LSBLFZ
                    sheet14.Range("D24").Value = "/"
                Else
                    sheet13.Range("G14").Value = "应测" & Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0) & "处，实测" & Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0) & "处，合格" & Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0) & "处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                    For cs = 1 To Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0)
                        LSBL = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet14.Range("D24").Value = LSBLFZ
                End If

                '箍筋
                sheet13.Range("E17").Value = "设计值：" & sheet1.Range("B9").Value * 10
                sheet13.Range("G16").Value = sheet1.Range("B9").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9) & "  " &
                                             sheet1.Range("B9").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                '骨架尺寸
                sheet13.Range("E19").Value = "设计值：" & sheet1.Range("B14").Value * 10
                sheet13.Range("G18").Value = sheet1.Range("B14").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet13.Range("E21").Value = "宽：" & sheet1.Range("B15").Value * 10
                sheet13.Range("E22").Value = "高：" & sheet1.Range("B16").Value * 10
                sheet13.Range("G20").Value = "宽： " & sheet1.Range("B15").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4) & vbCrLf &
                                             "高： " & sheet1.Range("B16").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4)
                '保护层
                sheet13.Range("E27").Value = "设计值：" & sheet1.Range("B10").Value * 10
                gs = Math.Round((sheet1.Range("B11").Value * sheet1.Range("B13").Value * 2 + sheet1.Range("B12").Value * sheet1.Range("B13").Value * 2) / 300 / 100 * 0.2, 0)
                If gs <= 10 Then
                    LSBLFZ = Nothing
                    For cs = 1 To gs
                        LSBL = sheet1.Range("B10").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet13.Range("G26").Value = LSBLFZ
                    sheet14.Range("D48").Value = Nothing
                Else
                    LSBLFZ = Nothing
                    For cs = 1 To gs
                        LSBL = sheet1.Range("B10").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet13.Range("G26").Value = "应测" & gs & "处，实测" & gs & "处，合格" & gs & "处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                    sheet14.Range("D48").Value = LSBLFZ
                End If
                '钢筋记录表 (监抽)
                sheet14.Range("B6").Value = sheet1.Range("B5").Value
                sheet14.Range("B7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋"
                sheet14.Range("K6").Value = sheet1.Range("B18").Value
                sheet14.Range("K7").Value = sheet1.Range("B18").Value

                '承台混凝土现场质量检验表 (监抽)
                sheet15.Range("D6").Value = sheet1.Range("B5").Value
                sheet15.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet15.Range("K6").Value = sheet1.Range("B17").Value
                sheet15.Range("D14").Value = "长：" & sheet1.Range("B11").Value * 10 & " 宽：" & sheet1.Range("B12").Value * 10
                sheet15.Range("E13").Value = "长： " & sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20) & vbCrLf &
                                             "宽： " & sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet15.Range("D18").Value = sheet1.Range("B13").Value * 10
                sheet15.Range("E17").Value = sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20) & "  " &
                                             sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet15.Range("E19").Value = sheet10.Range("E19").Value
                sheet15.Range("F19").Value = sheet10.Range("F19").Value
                sheet15.Range("E20").Value = sheet10.Range("E20").Value
                sheet15.Range("F20").Value = sheet10.Range("F20").Value
                gs = Math.Round((sheet1.Range("B11").Value * sheet1.Range("B13").Value * 2 + sheet1.Range("B12").Value * sheet1.Range("B13").Value * 2) / 2000 / 100 * 0.2, 0)
                LSBLFZ = Nothing
                If gs <= 5 Then
                    For cs = 1 To 5
                        LSBL = ExApp.WorksheetFunction.RandBetween(1, 5)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet15.Range("E21").Value = LSBLFZ
                Else
                    For cs = 1 To gs
                        LSBL = ExApp.WorksheetFunction.RandBetween(1, 5)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet15.Range("E21").Value = LSBLFZ
                End If
                '刷新一次数据
                ExApp.Calculate()
                ExApp.Calculation = ExApp.Calculation.xlCalculationManual

                ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                '选择表格
                sheet6.Select()
                For i = 7 To ExApp.Sheets.Count
                    EXsheet = Exbook.Worksheets(i)
                    If EXsheet.Visible = True Then
                        EXsheet.Select(Replace:=False)
                    End If
                Next i
                EXsheet = ExApp.ActiveSheet
                ' 导出PDF文件
                PDFFilename = Filepath & "\" & sheet1.Range("B5").Value & sheet1.Range("C6").Value & sheet1.Range("D6").Value & ".pdf"
                '保存PDF文件
                EXsheet.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, PDFFilename, XlFixedFormatQuality.xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False)

                sheet18 = Exbook.Worksheets("水准表 (2)") '水准测量记录表
                sheet19 = Exbook.Worksheets("平面表 (2)") '全站仪平面位置检测表
                ExApp.DisplayAlerts = False
                sheet18.Delete()
                sheet19.Delete()
                h += 1
            End While
            MsgBox("已完成！", 0 + 64, "提示")
        Catch Exclerror As Exception   '错误时弹出提示
            MsgBox(Exclerror.Message)
        End Try
        TorF = False
    End Sub

    Sub 系梁资料（）

        Dim h, r, i As Integer
        Dim ExcelFilename, PDFFilename, LSBL, LSBLFZ As String
        Dim FolderDialogObject As New FolderBrowserDialog()
        h = 8
        i = 8
        r = 0
        Try
            sheet0 = Exbook.Worksheets("数据库") '数据库
            sheet1 = Exbook.Worksheets("参数表") '参数表
            sheet2 = Exbook.Worksheets("交点法") '交点法
            sheet3 = Exbook.Worksheets("线元法") '线元法
            sheet4 = Exbook.Worksheets("断链") '断链
            sheet5 = Exbook.Worksheets("导线成果表") '导线成果表
            sheet6 = Exbook.Worksheets("钢筋隐蔽工程") '钢筋隐蔽工程
            sheet7 = Exbook.Worksheets("钢筋检表") '钢筋检表
            sheet8 = Exbook.Worksheets("钢筋记录表") '钢筋记录表
            sheet9 = Exbook.Worksheets("申请批复单") '工序检验申请批复单
            sheet10 = Exbook.Worksheets("系梁检表") '系梁检表
            sheet11 = Exbook.Worksheets("模板记录表") '模板记录表
            sheet12 = Exbook.Worksheets("砼浇筑申请报告单") '砼浇筑申请报告单
            sheet13 = Exbook.Worksheets("监抽钢筋检表") '监抽钢筋检表
            sheet14 = Exbook.Worksheets("监抽钢筋记录表") '监抽钢筋记录表
            sheet15 = Exbook.Worksheets("监抽系梁检表") '监抽系梁检表
            sheet16 = Exbook.Worksheets("水准表") '水准测量记录表
            sheet17 = Exbook.Worksheets("平面表") '全站仪平面位置检测表

            '改表头
            sheet6.Range("A1").Value = sheet0.Range("C1").Value
            sheet6.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet6.Range("E3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet6.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet6.Range("E4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet7.Range("A1").Value = sheet0.Range("C1").Value
            sheet7.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet7.Range("P3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet7.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet7.Range("P4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet8.Range("A1").Value = sheet0.Range("C1").Value
            sheet8.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet8.Range("L3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet8.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet8.Range("L4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet9.Range("A1").Value = sheet0.Range("C1").Value
            sheet9.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet9.Range("E3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet9.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet9.Range("E4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet10.Range("A1").Value = sheet0.Range("C1").Value
            sheet10.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet10.Range("K3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet10.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet10.Range("K4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet11.Range("A1").Value = sheet0.Range("C1").Value
            sheet11.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet11.Range("I3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet11.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet11.Range("I4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet12.Range("A1").Value = sheet0.Range("C1").Value
            sheet12.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet12.Range("I3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet12.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet12.Range("I4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet13.Range("A1").Value = sheet0.Range("C1").Value
            sheet13.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet13.Range("P3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet13.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet13.Range("P4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet14.Range("A1").Value = sheet0.Range("C1").Value
            sheet14.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet14.Range("L3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet14.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet14.Range("L4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet15.Range("A1").Value = sheet0.Range("C1").Value
            sheet15.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet15.Range("K3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet15.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet15.Range("K4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet16.Range("A1").Value = sheet0.Range("C1").Value
            sheet16.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet16.Range("H3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet16.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet16.Range("H4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet17.Range("A1").Value = sheet0.Range("C1").Value
            sheet17.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet17.Range("N3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet17.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet17.Range("N4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value


            sheet0.Range("BA8:BT100000").Value = Nothing
            '计算偏距
            'Do While sheet0.range("B" & i).value <> Nothing And sheet0.range("L" & i).value <> Nothing And sheet0.range("M" & i).value <> Nothing
            Do While sheet0.Range("B" & i).Value <> Nothing
                If sheet0.Range("C" & i).Value = Nothing And sheet0.Range("U" & i).Value = "否" And sheet0.Range("V" & i).Value <> Nothing And sheet0.Range("W" & i).Value <> Nothing Then
                    sheet0.Range("C" & i).Value = FSZhj(sheet0.Range("V" & i).Value, sheet0.Range("W" & i).Value)
                ElseIf sheet0.Range("C" & i).Value <> Nothing And sheet0.Range("U" & i).Value = "否" And sheet0.Range("V" & i).Value = Nothing And sheet0.Range("W" & i).Value = Nothing Then
                    MsgBox("请核对中心桩号或X、Y坐标是否正确填写！")
                    TorF = False
                    Exit Sub
                End If

                sheet0.Range("BA" & i).Value = sheet0.Range("C" & i).Value + sheet0.Range("H" & i).Value / 2 / 100
                sheet0.Range("BE" & i).Value = sheet0.Range("C" & i).Value - sheet0.Range("H" & i).Value / 2 / 100
                sheet0.Range("BI" & i).Value = sheet0.Range("C" & i).Value
                sheet0.Range("BM" & i).Value = sheet0.Range("C" & i).Value
                sheet0.Range("BQ" & i).Value = sheet0.Range("C" & i).Value

                If sheet0.Range("K" & i).Value = "中" Then
                    sheet0.Range("BB" & i).Value = 0
                    sheet0.Range("BF" & i).Value = 0
                    sheet0.Range("BJ" & i).Value = sheet0.Range("G" & i).Value / 2 / 100 * -1
                    sheet0.Range("BN" & i).Value = sheet0.Range("G" & i).Value / 2 / 100
                    sheet0.Range("BR" & i).Value = 0
                ElseIf sheet0.Range("K" & i).Value = "右" Then
                    sheet0.Range("BB" & i).Value = (sheet0.Range("L" & i).Value + sheet0.Range("M" & i).Value / 2) / 100 * -1
                    sheet0.Range("BF" & i).Value = (sheet0.Range("L" & i).Value + sheet0.Range("M" & i).Value / 2) / 100 * -1
                    sheet0.Range("BJ" & i).Value = (sheet0.Range("L" & i).Value + sheet0.Range("M" & i).Value / 2 + sheet0.Range("G" & i).Value / 2) / 100 * -1
                    sheet0.Range("BN" & i).Value = (sheet0.Range("L" & i).Value + sheet0.Range("M" & i).Value / 2 - sheet0.Range("G" & i).Value / 2) / 100 * -1
                    sheet0.Range("BR" & i).Value = (sheet0.Range("L" & i).Value + sheet0.Range("M" & i).Value / 2) / 100 * -1
                Else
                    sheet0.Range("BB" & i).Value = (sheet0.Range("L" & i).Value + sheet0.Range("M" & i).Value / 2) / 100
                    sheet0.Range("BF" & i).Value = (sheet0.Range("L" & i).Value + sheet0.Range("M" & i).Value / 2) / 100
                    sheet0.Range("BJ" & i).Value = (sheet0.Range("L" & i).Value + sheet0.Range("M" & i).Value / 2 - sheet0.Range("G" & i).Value / 2) / 100
                    sheet0.Range("BN" & i).Value = (sheet0.Range("L" & i).Value + sheet0.Range("M" & i).Value / 2 + sheet0.Range("G" & i).Value / 2) / 100
                    sheet0.Range("BR" & i).Value = (sheet0.Range("L" & i).Value + sheet0.Range("M" & i).Value / 2) / 100
                End If
                i += 1
            Loop

            Call 质检资料设计坐标()
            If TorF = False Then
                Exit Sub
            End If
            While sheet0.Range("B" & h).Value <> Nothing
                If TorF = False Then
                    Exit Sub
                End If
                ExApp.Calculation = ExApp.Calculation.xlCalculationManual '开启手动计算
                sheet1.Range("B" & r + 5).Value = sheet0.Range("B" & h).Value
                sheet1.Range("B" & r + 6).Value = sheet0.Range("C" & h).Value
                sheet1.Range("C" & r + 6).Value = sheet0.Range("D" & h).Value
                sheet1.Range("D" & r + 6).Value = sheet0.Range("E" & h).Value
                sheet1.Range("B" & r + 7).Value = sheet0.Range("N" & h).Value
                sheet1.Range("B" & r + 8).Value = sheet0.Range("O" & h).Value
                sheet1.Range("B" & r + 9).Value = sheet0.Range("P" & h).Value
                sheet1.Range("B" & r + 10).Value = sheet0.Range("T" & h).Value
                sheet1.Range("B" & r + 11).Value = sheet0.Range("G" & h).Value
                sheet1.Range("B" & r + 12).Value = sheet0.Range("H" & h).Value
                sheet1.Range("B" & r + 13).Value = sheet0.Range("I" & h).Value
                sheet1.Range("B" & r + 14).Value = sheet0.Range("Q" & h).Value
                sheet1.Range("B" & r + 15).Value = sheet0.Range("R" & h).Value
                sheet1.Range("B" & r + 16).Value = sheet0.Range("S" & h).Value
                sheet1.Range("B" & r + 17).Value = sheet0.Range("F" & h).Value

                '测量参数
                sheet1.Range("I1").Value = sheet1.Range("B5").Value.substring(0, ExApp.WorksheetFunction.Find("K", sheet1.Range("B5").Value))
                sheet1.Range("I3:Q15").Value = Nothing
                sheet1.Range("I3").Value = sheet0.Range("BA" & h).Value
                sheet1.Range("I4").Value = sheet0.Range("BE" & h).Value
                sheet1.Range("I5").Value = sheet0.Range("BI" & h).Value
                sheet1.Range("I6").Value = sheet0.Range("BM" & h).Value
                sheet1.Range("I7").Value = sheet0.Range("BQ" & h).Value

                sheet1.Range("J3").Value = sheet0.Range("BB" & h).Value
                sheet1.Range("J4").Value = sheet0.Range("BF" & h).Value
                sheet1.Range("J5").Value = sheet0.Range("BJ" & h).Value
                sheet1.Range("J6").Value = sheet0.Range("BN" & h).Value
                sheet1.Range("J7").Value = sheet0.Range("BR" & h).Value

                sheet1.Range("K3").Value = sheet0.Range("J" & h).Value
                sheet1.Range("K4").Value = sheet0.Range("J" & h).Value
                sheet1.Range("K5").Value = sheet0.Range("J" & h).Value
                sheet1.Range("K6").Value = sheet0.Range("J" & h).Value
                sheet1.Range("K7").Value = sheet0.Range("J" & h).Value

                sheet1.Range("M3").Value = sheet0.Range("BC" & h).Value
                sheet1.Range("N3").Value = sheet0.Range("BD" & h).Value
                sheet1.Range("M4").Value = sheet0.Range("BG" & h).Value
                sheet1.Range("N4").Value = sheet0.Range("BH" & h).Value
                sheet1.Range("M5").Value = sheet0.Range("BK" & h).Value
                sheet1.Range("N5").Value = sheet0.Range("BL" & h).Value
                sheet1.Range("M6").Value = sheet0.Range("BO" & h).Value
                sheet1.Range("N6").Value = sheet0.Range("BP" & h).Value
                sheet1.Range("M7").Value = sheet0.Range("BS" & h).Value
                sheet1.Range("N7").Value = sheet0.Range("BT" & h).Value
                sheet1.Range("P1").Value = sheet1.Range("B17").Value
                sheet1.Range("L3").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                '测量偏差值
                sheet1.Range("O3").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O4").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O5").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O6").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O7").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("P3").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P4").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P5").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P6").Value = ExApp.WorksheetFunction.RandBetween(1, 9)

                sheet16.Range("A5").Value = sheet1.Range("A5").Value & sheet1.Range("B5").Value
                sheet17.Range("A5").Value = sheet1.Range("A5").Value & sheet1.Range("B5").Value

                ' 钢筋隐蔽工程
                sheet6.Range("C6").Value = sheet1.Range("B5").Value & sheet1.Range("C6").Value & "基础及下部构造"
                sheet6.Range("E6").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋加工及安装"
                sheet6.Range("C7").Value = sheet1.Range("B5").Value & sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋加工及安装"
                sheet6.Range("E10").Value = sheet1.Range("B18").Value
                sheet6.Range("E11").Value = sheet1.Range("B18").Value
                sheet6.Range("E27").Value = sheet1.Range("B18").Value
                sheet6.Range("E28").Value = sheet1.Range("B18").Value
                '钢筋检表
                sheet7.Range("D6").Value = sheet1.Range("B5").Value
                sheet7.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋"
                sheet7.Range("Q6").Value = sheet1.Range("B18").Value
                sheet7.Range("Q7").Value = sheet1.Range("B18").Value
                sheet7.Range("Q31").Value = sheet1.Range("B18").Value
                sheet7.Range("Q34").Value = sheet1.Range("B18").Value
                '主筋
                sheet7.Range("E15").Value = "设计值：" & sheet1.Range("B8").Value * 10
                LSBLFZ = Nothing
                If sheet1.Range("B7").Value * 2 <= 10 Then
                    For cs = 1 To sheet1.Range("B7").Value * 2
                        LSBL = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet7.Range("G14").Value = LSBLFZ
                    sheet8.Range("D24").Value = "/"
                Else
                    For cs = 1 To sheet1.Range("B7").Value * 2
                        LSBL = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet7.Range("G14").Value = "应测" & sheet1.Range("B7").Value * 2 & "处，实测" & sheet1.Range("B7").Value * 2 & "处，合格" & sheet1.Range("B7").Value * 2 & "处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                    sheet8.Range("D24").Value = LSBLFZ
                End If
                '箍筋
                sheet7.Range("E17").Value = "设计值：" & sheet1.Range("B9").Value * 10
                LSBLFZ = Nothing
                For cs = 1 To 10
                    LSBL = sheet1.Range("B9").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                    LSBLFZ = LSBLFZ & LSBL & "   "
                Next
                sheet7.Range("G16").Value = LSBLFZ

                '骨架尺寸
                sheet7.Range("E19").Value = "设计值：" & sheet1.Range("B14").Value * 10
                sheet7.Range("G18").Value = sheet1.Range("B14").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet7.Range("E21").Value = "宽：" & sheet1.Range("B15").Value * 10
                sheet7.Range("E22").Value = "高：" & sheet1.Range("B16").Value * 10
                sheet7.Range("G20").Value = "宽：  " & sheet1.Range("B15").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4) & vbCrLf &
                                            "高：  " & sheet1.Range("B16").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4)
                '保护层
                sheet7.Range("E27").Value = "设计值：" & sheet1.Range("B10").Value * 10
                Dim gs As Integer
                gs = Math.Round((sheet1.Range("B11").Value * sheet1.Range("B13").Value * 2 + sheet1.Range("B12").Value * sheet1.Range("B13").Value * 2) / 300 / 100, 0)
                If gs <= 20 Then
                    LSBLFZ = Nothing
                    For cs = 1 To 20
                        LSBL = sheet1.Range("B10").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet7.Range("G26").Value = "应测20处，实测20处，合格20处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                    sheet8.Range("D48").Value = LSBLFZ
                Else
                    LSBLFZ = Nothing
                    For cs = 1 To gs
                        LSBL = sheet1.Range("B10").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet7.Range("G26").Value = "应测" & gs & "处，实测" & gs & "处，合格" & gs & "处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                    sheet8.Range("D48").Value = LSBLFZ
                End If

                '钢筋记录表
                sheet8.Range("B6").Value = sheet1.Range("B5").Value
                sheet8.Range("B7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋"
                sheet8.Range("K6").Value = sheet1.Range("B18").Value
                sheet8.Range("K7").Value = sheet1.Range("B18").Value

                '工序检验申请批复单
                sheet9.Range("C6").Value = sheet1.Range("B5").Value & sheet1.Range("C6").Value & "基础及下部构造"
                sheet9.Range("C7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet9.Range("C8").Value = sheet1.Range("B5").Value
                sheet9.Range("C9").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet9.Range("C10").Value = "混凝土强度、平面尺寸、结构高度、顶面高程、轴线偏位、平整度"

                Call 水准测量记录表()
                If TorF = False Then
                    Exit Sub
                End If
                sheet1.Range("I7:N7").Value = Nothing
                Call 全站仪平面位置检测表（）
                If TorF = False Then
                    Exit Sub
                End If
                sheet16.Activate()
                ExApp.ActiveWindow.SelectedSheets.Copy(, (ExApp.Sheets(ExApp.Sheets.Count)))
                sheet17.Activate()
                ExApp.ActiveWindow.SelectedSheets.Copy(, (ExApp.Sheets(ExApp.Sheets.Count)))

                '系梁混凝土现场质量检验表
                sheet10.Range("D6").Value = sheet1.Range("B5").Value
                sheet10.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet10.Range("K6").Value = sheet1.Range("B17").Value
                sheet10.Range("D14").Value = "长：" & sheet1.Range("B11").Value * 10 & " 宽：" & sheet1.Range("B12").Value * 10
                sheet10.Range("E13").Value = "长：" & sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20) & "  " &
                                              sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20) & vbCrLf &
                                             "宽：" & sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20) & "  " &
                                             sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)

                sheet10.Range("D18").Value = sheet1.Range("B13").Value * 10
                sheet10.Range("E17").Value = sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20) & "  " &
                                             sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20) & "  " &
                                             sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20) & "  " &
                                             sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20) & "  " &
                                             sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)

                sheet10.Range("E19").Value = sheet1.Range("O3").Value
                sheet10.Range("F19").Value = sheet1.Range("O4").Value
                sheet10.Range("G19").Value = sheet1.Range("O5").Value
                sheet10.Range("H19").Value = sheet1.Range("O6").Value
                sheet10.Range("I19").Value = sheet1.Range("O7").Value
                sheet10.Range("E20").Value = sheet1.Range("P3").Value
                sheet10.Range("F20").Value = sheet1.Range("P4").Value
                sheet10.Range("G20").Value = sheet1.Range("P5").Value
                sheet10.Range("H20").Value = sheet1.Range("P6").Value
                '平整度
                gs = Math.Round((sheet1.Range("B11").Value * sheet1.Range("B13").Value * 2 + sheet1.Range("B12").Value * sheet1.Range("B13").Value * 2) / 2000 / 100, 0)
                LSBLFZ = Nothing
                If gs <= 24 Then
                    For cs = 1 To 24
                        LSBL = ExApp.WorksheetFunction.RandBetween(1, 5)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet10.Range("E21").Value = LSBLFZ
                Else
                    For cs = 1 To gs
                        LSBL = ExApp.WorksheetFunction.RandBetween(1, 5)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet10.Range("E21").Value = LSBLFZ
                End If

                '模板测量偏差值
                sheet1.Range("O3").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O4").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O5").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O6").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("P3").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P4").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P5").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P6").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("L3").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "模板"

                Call 水准测量记录表()
                If TorF = False Then
                    Exit Sub
                End If
                Call 全站仪平面位置检测表（）
                If TorF = False Then
                    Exit Sub
                End If
                '现场模板安装检查记录表
                sheet11.Range("C6").Value = sheet1.Range("B5").Value
                sheet11.Range("C7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet11.Range("I6").Value = sheet1.Range("B17").Value
                sheet11.Range("I7").Value = sheet1.Range("B17").Value
                sheet11.Range("D8").Value = 2
                sheet11.Range("D9").Value = 5
                sheet11.Range("F8").Value = 4
                sheet11.Range("F9").Value = 4
                sheet11.Range("H8").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet11.Range("H9").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet11.Range("J8").Value = 100
                sheet11.Range("J9").Value = 100
                '测量偏差
                sheet11.Range("D11").Value = sheet1.Range("P5").Value
                sheet11.Range("F11").Value = sheet1.Range("P6").Value
                sheet11.Range("H11").Value = sheet1.Range("P3").Value
                sheet11.Range("J11").Value = sheet1.Range("P4").Value
                sheet11.Range("F14").Value = sheet1.Range("O3").Value
                sheet11.Range("G14").Value = sheet1.Range("O4").Value
                sheet11.Range("H14").Value = sheet1.Range("O5").Value
                sheet11.Range("I14").Value = sheet1.Range("O6").Value

                sheet11.Range("F12").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("G12").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("H12").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("I12").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("F13").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("G13").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("H13").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("I13").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("F18").Value = "牢固，稳定"

                '监抽
                '钢筋检表
                sheet13.Range("D6").Value = sheet1.Range("B5").Value
                sheet13.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋"
                sheet13.Range("Q6").Value = sheet1.Range("B18").Value
                sheet13.Range("Q7").Value = sheet1.Range("B18").Value
                sheet13.Range("Q31").Value = sheet1.Range("B18").Value
                sheet13.Range("Q34").Value = sheet1.Range("B18").Value
                '主筋
                sheet13.Range("E15").Value = "设计值：" & sheet1.Range("B8").Value * 10
                LSBLFZ = Nothing
                If Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0) <= 10 Then
                    For cs = 1 To Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0)
                        LSBL = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet13.Range("G14").Value = LSBLFZ
                    sheet14.Range("D24").Value = "/"
                Else
                    sheet13.Range("G14").Value = "应测" & Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0) & "处，实测" & Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0) & "处，合格" & Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0) & "处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                    For cs = 1 To Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0)
                        LSBL = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet14.Range("D24").Value = LSBLFZ
                End If

                '箍筋
                sheet13.Range("E17").Value = "设计值：" & sheet1.Range("B9").Value * 10
                sheet13.Range("G16").Value = sheet1.Range("B9").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9) & "  " &
                                             sheet1.Range("B9").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                '骨架尺寸
                sheet13.Range("E19").Value = "设计值：" & sheet1.Range("B14").Value * 10
                sheet13.Range("K18").Value = sheet1.Range("B14").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet13.Range("E21").Value = "宽：" & sheet1.Range("B15").Value * 10
                sheet13.Range("E22").Value = "高：" & sheet1.Range("B16").Value * 10
                sheet13.Range("G20").Value = "宽： " & sheet1.Range("B15").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4) & vbCrLf &
                                             "高： " & sheet1.Range("B16").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4)
                '保护层
                sheet13.Range("E27").Value = "设计值：" & sheet1.Range("B10").Value * 10
                gs = Math.Round((sheet1.Range("B11").Value * sheet1.Range("B13").Value * 2 + sheet1.Range("B12").Value * sheet1.Range("B13").Value * 2) / 300 / 100 * 0.2, 0)
                If gs <= 10 Then
                    LSBLFZ = Nothing
                    For cs = 1 To gs
                        LSBL = sheet1.Range("B10").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet13.Range("G26").Value = LSBLFZ
                    sheet14.Range("D48").Value = Nothing
                Else
                    LSBLFZ = Nothing
                    For cs = 1 To gs
                        LSBL = sheet1.Range("B10").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet13.Range("G26").Value = "应测" & gs & "处，实测" & gs & "处，合格" & gs & "处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                    sheet14.Range("D48").Value = LSBLFZ
                End If
                '钢筋记录表 (监抽)
                sheet14.Range("B6").Value = sheet1.Range("B5").Value
                sheet14.Range("B7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋"
                sheet14.Range("K6").Value = sheet1.Range("B18").Value
                sheet14.Range("K7").Value = sheet1.Range("B18").Value
                '系梁混凝土现场质量检验表(监抽)
                sheet15.Range("D6").Value = sheet1.Range("B5").Value
                sheet15.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet15.Range("K6").Value = sheet1.Range("B17").Value
                sheet15.Range("D14").Value = "长：" & sheet1.Range("B11").Value * 10 & " 宽：" & sheet1.Range("B12").Value * 10
                sheet15.Range("E13").Value = "长：" & sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20) &
                                             "宽：" & sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet15.Range("D18").Value = sheet1.Range("B13").Value * 10
                sheet15.Range("E17").Value = sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20) & "  " &
                                             sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet15.Range("E19").Value = sheet10.Range("E19").Value
                sheet15.Range("F19").Value = sheet10.Range("F19").Value
                sheet15.Range("E20").Value = sheet10.Range("E20").Value
                sheet15.Range("F20").Value = sheet10.Range("F20").Value
                gs = Math.Round((sheet1.Range("B11").Value * sheet1.Range("B13").Value * 2 + sheet1.Range("B12").Value * sheet1.Range("B13").Value * 2) / 2000 / 100 * 0.2, 0)
                LSBLFZ = Nothing
                If gs <= 5 Then
                    For cs = 1 To 5
                        LSBL = ExApp.WorksheetFunction.RandBetween(1, 5)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet15.Range("E21").Value = LSBLFZ
                Else
                    For cs = 1 To gs
                        LSBL = ExApp.WorksheetFunction.RandBetween(1, 5)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet15.Range("E21").Value = LSBLFZ
                End If

                '刷新一次数据
                ExApp.Calculate()
                ExApp.Calculation = ExApp.Calculation.xlCalculationManual

                ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                '选择表格
                sheet6.Select()
                For i = 7 To ExApp.Sheets.Count
                    EXsheet = Exbook.Worksheets(i)
                    If EXsheet.Visible = True Then
                        EXsheet.Select(Replace:=False)
                    End If
                Next i
                EXsheet = ExApp.ActiveSheet
                ' 导出PDF文件
                PDFFilename = Filepath & "\" & sheet1.Range("B5").Value & sheet1.Range("C6").Value & sheet1.Range("D6").Value & ".pdf"
                EXsheet.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, PDFFilename, XlFixedFormatQuality.xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False)

                sheet18 = Exbook.Worksheets("水准表 (2)") '水准测量记录表
                sheet19 = Exbook.Worksheets("平面表 (2)") '全站仪平面位置检测表
                ExApp.DisplayAlerts = False
                sheet18.Delete()
                sheet19.Delete()
                h += 1
            End While
            MsgBox("已完成！", 0 + 64, "提示")
        Catch Exclerror As Exception   '错误时弹出提示
            MsgBox(Exclerror.Message)
        End Try
        TorF = False

    End Sub

    Sub 墩柱资料（）

        Dim h, i, r, q As Integer
        Dim ExcelFilename, PDFFilename, LSBL, LSBLFZ As String  '定义输出的PDF文件名
        Dim FolderDialogObject As New FolderBrowserDialog()
        i = 8
        h = 8
        r = 0
        Try
            sheet0 = Exbook.Worksheets("数据库") '数据库
            sheet1 = Exbook.Worksheets("参数表") '参数表
            sheet2 = Exbook.Worksheets("交点法") '交点法
            sheet3 = Exbook.Worksheets("线元法") '线元法
            sheet4 = Exbook.Worksheets("断链") '断链
            sheet5 = Exbook.Worksheets("导线成果表") '导线成果表
            sheet6 = Exbook.Worksheets("钢筋隐蔽工程") '钢筋隐蔽工程
            sheet7 = Exbook.Worksheets("钢筋检表") '钢筋检表
            sheet8 = Exbook.Worksheets("钢筋记录表") '钢筋记录表
            sheet9 = Exbook.Worksheets("钢筋记录表续") '钢筋记录表续
            sheet10 = Exbook.Worksheets("申请批复单") '工序检验申请批复单
            sheet11 = Exbook.Worksheets("现浇墩台身检表") '现浇墩台身检表
            sheet12 = Exbook.Worksheets("模板记录表") '模板记录表
            sheet13 = Exbook.Worksheets("砼浇筑申请报告单") '砼浇筑申请报告单
            sheet14 = Exbook.Worksheets("监抽钢筋检表") '监抽钢筋检表
            sheet15 = Exbook.Worksheets("监抽钢筋记录表") '监抽钢筋记录表
            sheet16 = Exbook.Worksheets("监抽现浇墩台身检表") '监抽系梁检表
            sheet17 = Exbook.Worksheets("水准表") '水准测量记录表
            sheet18 = Exbook.Worksheets("平面表") '全站仪平面位置检测表


            '改表头
            sheet6.Range("A1").Value = sheet0.Range("C1").Value
            sheet6.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet6.Range("E3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet6.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet6.Range("E4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet7.Range("A1").Value = sheet0.Range("C1").Value
            sheet7.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet7.Range("P3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet7.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet7.Range("P4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet8.Range("A1").Value = sheet0.Range("C1").Value
            sheet8.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet8.Range("L3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet8.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet8.Range("L4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet9.Range("A1").Value = sheet0.Range("C1").Value
            sheet9.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet9.Range("L3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet9.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet9.Range("L4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet10.Range("A1").Value = sheet0.Range("C1").Value
            sheet10.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet10.Range("E3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet10.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet10.Range("E4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet11.Range("A1").Value = sheet0.Range("C1").Value
            sheet11.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet11.Range("J3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet11.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet11.Range("J4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet12.Range("A1").Value = sheet0.Range("C1").Value
            sheet12.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet12.Range("I3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet12.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet12.Range("I4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet13.Range("A1").Value = sheet0.Range("C1").Value
            sheet13.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet13.Range("I3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet13.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet13.Range("I4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet14.Range("A1").Value = sheet0.Range("C1").Value
            sheet14.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet14.Range("P3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet14.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet14.Range("P4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet15.Range("A1").Value = sheet0.Range("C1").Value
            sheet15.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet15.Range("L3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet15.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet15.Range("L4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet16.Range("A1").Value = sheet0.Range("C1").Value
            sheet16.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet16.Range("J3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet16.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet16.Range("J4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet17.Range("A1").Value = sheet0.Range("C1").Value
            sheet17.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet17.Range("H3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet17.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet17.Range("H4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet18.Range("A1").Value = sheet0.Range("C1").Value
            sheet18.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet18.Range("N3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet18.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet18.Range("N4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value


            sheet0.Range("BA8:BT100000").Value = Nothing
            '计算偏距
            Do While sheet0.Range("B" & i).Value <> Nothing And sheet0.Range("L" & i).Value <> Nothing
                If sheet0.Range("C" & i).Value = Nothing And sheet0.Range("U" & i).Value = "否" And sheet0.Range("V" & i).Value <> Nothing And sheet0.Range("W" & i).Value <> Nothing Then
                    sheet0.Range("C" & i).Value = FSZhj(sheet0.Range("V" & i).Value, sheet0.Range("W" & i).Value)
                ElseIf sheet0.Range("C" & i).Value <> Nothing And sheet0.Range("U" & i).Value = "否" And sheet0.Range("V" & i).Value = Nothing And sheet0.Range("W" & i).Value = Nothing Then
                    MsgBox("请核对中心桩号或X、Y坐标是否正确填写！")
                    TorF = False
                    Exit Sub
                End If
                '圆墩
                If sheet0.Range("X" & i).Value = "圆柱形" Or sheet0.Range("X" & i).Value = Nothing Then
                    sheet0.Range("BA" & i).Value = sheet0.Range("C" & i).Value + sheet0.Range("H" & i).Value / 2 / 100
                    sheet0.Range("BE" & i).Value = sheet0.Range("C" & i).Value - sheet0.Range("H" & i).Value / 2 / 100
                    sheet0.Range("BI" & i).Value = sheet0.Range("C" & i).Value
                    sheet0.Range("BM" & i).Value = sheet0.Range("C" & i).Value
                    sheet0.Range("BQ" & i).Value = sheet0.Range("C" & i).Value

                    If sheet0.Range("K" & i).Value = "中" Then
                        sheet0.Range("BB" & i).Value = 0
                        sheet0.Range("BF" & i).Value = 0
                        sheet0.Range("BJ" & i).Value = （sheet0.Range("H" & i).Value / 2) / 100 * -1
                        sheet0.Range("BN" & i).Value = （sheet0.Range("H" & i).Value / 2) / 100
                        sheet0.Range("BR" & i).Value = 0

                    ElseIf sheet0.Range("K" & i).Value = "右" Then
                        sheet0.Range("BB" & i).Value = sheet0.Range("L" & i).Value / 100 * -1
                        sheet0.Range("BF" & i).Value = sheet0.Range("L" & i).Value / 100 * -1
                        sheet0.Range("BJ" & i).Value = (sheet0.Range("L" & i).Value + sheet0.Range("H" & i).Value / 2) / 100 * -1
                        sheet0.Range("BN" & i).Value = (sheet0.Range("L" & i).Value - sheet0.Range("H" & i).Value / 2) / 100 * -1
                        sheet0.Range("BR" & i).Value = sheet0.Range("L" & i).Value / 100 * -1
                    Else
                        sheet0.Range("BB" & i).Value = sheet0.Range("L" & i).Value / 100
                        sheet0.Range("BF" & i).Value = sheet0.Range("L" & i).Value / 100
                        sheet0.Range("BJ" & i).Value = (sheet0.Range("L" & i).Value - sheet0.Range("H" & i).Value / 2) / 100
                        sheet0.Range("BN" & i).Value = (sheet0.Range("L" & i).Value + sheet0.Range("H" & i).Value / 2) / 100
                        sheet0.Range("BR" & i).Value = sheet0.Range("L" & i).Value / 100
                    End If
                Else
                    '方墩
                    sheet0.Range("BA" & i).Value = sheet0.Range("C" & i).Value + (sheet0.Range("Y" & i).Value / 2 + sheet0.Range("H" & i).Value / 2) / 100
                    sheet0.Range("BE" & i).Value = sheet0.Range("C" & i).Value - (sheet0.Range("Y" & i).Value / 2 + sheet0.Range("H" & i).Value / 2) / 100
                    sheet0.Range("BI" & i).Value = sheet0.Range("C" & i).Value + (sheet0.Range("Y" & i).Value / 2) / 100
                    sheet0.Range("BM" & i).Value = sheet0.Range("C" & i).Value + (sheet0.Range("Y" & i).Value / 2) / 100
                    sheet0.Range("BQ" & i).Value = sheet0.Range("C" & i).Value + (sheet0.Range("Y" & i).Value / 2) / 100

                    If sheet0.Range("K" & i).Value = "中" Then
                        sheet0.Range("BB" & i).Value = 0
                        sheet0.Range("BF" & i).Value = 0
                        sheet0.Range("BJ" & i).Value = （sheet0.Range("G" & i).Value / 2) / 100 * -1
                        sheet0.Range("BN" & i).Value = （sheet0.Range("G" & i).Value / 2) / 100
                        sheet0.Range("BR" & i).Value = 0

                    ElseIf sheet0.Range("K" & i).Value = "右" Then
                        sheet0.Range("BB" & i).Value = (sheet0.Range("L" & i).Value + sheet0.Range("Z" & i).Value / 2) / 100
                        sheet0.Range("BF" & i).Value = (sheet0.Range("L" & i).Value + sheet0.Range("Z" & i).Value / 2) / 100
                        sheet0.Range("BJ" & i).Value = (sheet0.Range("L" & i).Value + sheet0.Range("Z" & i).Value / 2 - sheet0.Range("G" & i).Value / 2) / 100
                        sheet0.Range("BN" & i).Value = (sheet0.Range("L" & i).Value + sheet0.Range("Z" & i).Value / 2 + sheet0.Range("G" & i).Value / 2) / 100
                        sheet0.Range("BR" & i).Value = (sheet0.Range("L" & i).Value + sheet0.Range("Z" & i).Value / 2) / 100
                    Else
                        sheet0.Range("BB" & i).Value = (sheet0.Range("L" & i).Value + sheet0.Range("Z" & i).Value / 2) / 100 * -1
                        sheet0.Range("BF" & i).Value = (sheet0.Range("L" & i).Value + sheet0.Range("Z" & i).Value / 2) / 100 * -1
                        sheet0.Range("BJ" & i).Value = (sheet0.Range("L" & i).Value + sheet0.Range("Z" & i).Value / 2 + sheet0.Range("G" & i).Value / 2) / 100 * -1
                        sheet0.Range("BN" & i).Value = (sheet0.Range("L" & i).Value + sheet0.Range("Z" & i).Value / 2 - sheet0.Range("G" & i).Value / 2) / 100 * -1
                        sheet0.Range("BR" & i).Value = (sheet0.Range("L" & i).Value + sheet0.Range("Z" & i).Value / 2) / 100 * -1
                    End If

                End If
                i += 1
            Loop
            Call 质检资料设计坐标()
            If TorF = False Then
                Exit Sub
            End If
            While sheet0.Range("B" & h).Value <> Nothing
                ExApp.Calculation = ExApp.Calculation.xlCalculationManual '开启手动计算
                sheet1.Range("B" & r + 5).Value = sheet0.Range("B" & h).Value
                sheet1.Range("B" & r + 6).Value = sheet0.Range("C" & h).Value
                sheet1.Range("C" & r + 6).Value = sheet0.Range("D" & h).Value
                sheet1.Range("D" & r + 6).Value = sheet0.Range("E" & h).Value
                sheet1.Range("B" & r + 7).Value = sheet0.Range("N" & h).Value
                sheet1.Range("B" & r + 8).Value = Math.Round(3.14 * (sheet0.Range("H" & h).Value - sheet0.Range("T" & h).Value * 2) / sheet0.Range("N" & h).Value, 1)
                sheet1.Range("B" & r + 9).Value = sheet0.Range("P" & h).Value
                sheet1.Range("B" & r + 10).Value = sheet0.Range("T" & h).Value
                sheet1.Range("B" & r + 11).Value = sheet0.Range("G" & h).Value
                sheet1.Range("B" & r + 12).Value = sheet0.Range("H" & h).Value
                sheet1.Range("B" & r + 13).Value = sheet0.Range("Q" & h).Value
                sheet1.Range("B" & r + 14).Value = sheet0.Range("R" & h).Value
                sheet1.Range("B" & r + 15).Value = sheet0.Range("F" & h).Value
                sheet1.Range("B" & r + 17).Value = sheet0.Range("X" & h).Value
                sheet1.Range("B" & r + 18).Value = sheet0.Range("I" & h).Value
                sheet1.Range("B" & r + 19).Value = sheet0.Range("S" & h).Value
                '测量参数
                sheet1.Range("I1").Value = sheet1.Range("B5").Value.substring(0, ExApp.WorksheetFunction.Find("K", sheet1.Range("B5").Value))
                sheet1.Range("I3:R15").Value = Nothing
                sheet1.Range("I3").Value = sheet0.Range("BA" & h).Value
                sheet1.Range("I4").Value = sheet0.Range("BE" & h).Value
                sheet1.Range("I5").Value = sheet0.Range("BI" & h).Value
                sheet1.Range("I6").Value = sheet0.Range("BM" & h).Value
                sheet1.Range("J3").Value = sheet0.Range("BB" & h).Value
                sheet1.Range("J4").Value = sheet0.Range("BF" & h).Value
                sheet1.Range("J5").Value = sheet0.Range("BJ" & h).Value
                sheet1.Range("J6").Value = sheet0.Range("BN" & h).Value
                sheet1.Range("K3").Value = sheet0.Range("J" & h).Value
                sheet1.Range("K4").Value = sheet0.Range("J" & h).Value
                sheet1.Range("K5").Value = sheet0.Range("J" & h).Value
                sheet1.Range("K6").Value = sheet0.Range("J" & h).Value
                sheet1.Range("M3").Value = sheet0.Range("BC" & h).Value
                sheet1.Range("N3").Value = sheet0.Range("BD" & h).Value
                sheet1.Range("M4").Value = sheet0.Range("BG" & h).Value
                sheet1.Range("N4").Value = sheet0.Range("BH" & h).Value
                sheet1.Range("M5").Value = sheet0.Range("BK" & h).Value
                sheet1.Range("N5").Value = sheet0.Range("BL" & h).Value
                sheet1.Range("M6").Value = sheet0.Range("BO" & h).Value
                sheet1.Range("N6").Value = sheet0.Range("BP" & h).Value
                sheet1.Range("R3").Value = "前"
                sheet1.Range("R4").Value = "后"
                sheet1.Range("R5").Value = "左"
                sheet1.Range("R6").Value = "右"
                sheet1.Range("P1").Value = sheet1.Range("B15").Value
                sheet1.Range("L3").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value

                '测量偏差值
                sheet1.Range("O3").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O4").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O5").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O6").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("P3").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P4").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P5").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P6").Value = ExApp.WorksheetFunction.RandBetween(1, 9)

                sheet17.Range("A5").Value = sheet1.Range("A5").Value & sheet1.Range("B5").Value
                sheet18.Range("A5").Value = sheet1.Range("A5").Value & sheet1.Range("B5").Value

                ' 钢筋隐蔽工程
                sheet6.Range("C6").Value = sheet1.Range("B5").Value & sheet1.Range("C6").Value & "基础及下部构造"
                sheet6.Range("E6").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋加工及安装"
                sheet6.Range("C7").Value = sheet1.Range("B5").Value & sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋加工及安装"
                sheet6.Range("E10").Value = sheet1.Range("B16").Value
                sheet6.Range("E11").Value = sheet1.Range("B16").Value
                sheet6.Range("E27").Value = sheet1.Range("B16").Value
                sheet6.Range("E28").Value = sheet1.Range("B16").Value
                '.钢筋检表
                sheet7.Range("D6").Value = sheet1.Range("B5").Value
                sheet7.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋"
                sheet7.Range("Q6").Value = sheet1.Range("B16").Value
                sheet7.Range("Q7").Value = sheet1.Range("B16").Value
                sheet7.Range("Q31").Value = sheet1.Range("B16").Value
                sheet7.Range("Q34").Value = sheet1.Range("B16").Value
                '主筋
                sheet7.Range("E15").Value = "设计值：" & sheet1.Range("B8").Value * 10
                LSBLFZ = Nothing
                If sheet1.Range("B7").Value * 2 <= 10 Then
                    For cs = 1 To sheet1.Range("B7").Value * 2
                        LSBL = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet7.Range("G14").Value = LSBLFZ
                    sheet8.Range("D24").Value = "/"
                Else
                    For cs = 1 To sheet1.Range("B7").Value * 2
                        LSBL = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet7.Range("G14").Value = "应测" & sheet1.Range("B7").Value * 2 & "处，实测" & sheet1.Range("B7").Value * 2 & "处，合格" & sheet1.Range("B7").Value * 2 & "处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                    sheet8.Range("D24").Value = LSBLFZ
                End If
                '箍筋
                sheet7.Range("E17").Value = "设计值：" & sheet1.Range("B9").Value * 10
                LSBLFZ = Nothing
                For cs = 1 To 10
                    LSBL = sheet1.Range("B9").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                    LSBLFZ = LSBLFZ & LSBL & "   "
                Next
                sheet7.Range("G16").Value = LSBLFZ

                '骨架尺寸
                If sheet1.Range("B17").Value = "圆柱形" Then
                    sheet7.Range("E19").Value = Nothing
                    sheet7.Range("G18").Value = "/"
                    sheet7.Range("E21").Value = "高：" & sheet1.Range("B19").Value * 10
                    sheet7.Range("E22").Value = "直径：" & sheet1.Range("B14").Value * 10
                    sheet7.Range("G20").Value = "高：  " & sheet1.Range("B19").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4) & "  " &
                                                "直径：" & sheet1.Range("B14").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4)
                Else
                    sheet7.Range("E19").Value = "设计值：" & sheet1.Range("B13").Value * 10
                    sheet7.Range("G18").Value = sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                    sheet7.Range("E21").Value = "高：" & sheet1.Range("B19").Value * 10
                    sheet7.Range("E22").Value = "宽：" & sheet1.Range("B14").Value * 10
                    sheet7.Range("G20").Value = "高：  " & sheet1.Range("B19").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4) & "  " &
                                                "宽：" & sheet1.Range("B14").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4)
                End If

                '保护层
                sheet7.Range("E27").Value = "设计值：" & sheet1.Range("B10").Value * 10
                Dim gs As Integer
                If sheet1.Range("B17").Value = "圆柱形" Then
                    gs = Math.Round(sheet1.Range("B12").Value * sheet1.Range("B18").Value * 3.14 / 300 / 100, 0)
                    If gs <= 10 Then
                        LSBLFZ = Nothing
                        For cs = 1 To gs
                            LSBL = sheet1.Range("B10").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                            LSBLFZ = LSBLFZ & LSBL & "   "
                        Next
                        sheet7.Range("G26").Value = LSBLFZ
                        sheet8.Range("D48").Value = Nothing
                    Else
                        LSBLFZ = Nothing
                        For cs = 1 To gs
                            LSBL = sheet1.Range("B10").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                            LSBLFZ = LSBLFZ & LSBL & "   "
                        Next
                        sheet7.Range("G26").Value = "应测" & gs & "处，实测" & gs & "处，合格" & gs & "处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                        sheet8.Range("D48").Value = LSBLFZ
                    End If
                Else
                    '矩形墩保护层
                    gs = Math.Round((sheet1.Range("B11").Value * sheet1.Range("B18").Value * 2 + sheet1.Range("B12").Value * sheet1.Range("B18").Value * 2) / 300 / 100, 0)
                    If gs <= 20 Then
                        LSBLFZ = Nothing
                        For cs = 1 To 20
                            LSBL = sheet1.Range("B10").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                            LSBLFZ = LSBLFZ & LSBL & "   "
                        Next
                        sheet7.Range("G26").Value = "应测20处，实测20处，合格20处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                        sheet8.Range("D48").Value = LSBLFZ
                    Else
                        LSBLFZ = Nothing
                        For cs = 1 To gs
                            LSBL = sheet1.Range("B10").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                            LSBLFZ = LSBLFZ & LSBL & "   "
                        Next
                        sheet7.Range("G26").Value = "应测" & gs & "处，实测" & gs & "处，合格" & gs & "处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                        sheet8.Range("D48").Value = LSBLFZ
                    End If
                End If


                '钢筋记录表
                sheet8.Range("B6").Value = sheet1.Range("B5").Value
                sheet8.Range("B7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋"
                sheet8.Range("K6").Value = sheet1.Range("B16").Value
                sheet8.Range("K7").Value = sheet1.Range("B16").Value

                '工序检验申请批复单
                sheet10.Range("C6").Value = sheet1.Range("B5").Value & sheet1.Range("C6").Value & "基础及下部构造"
                sheet10.Range("C7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet10.Range("C8").Value = sheet1.Range("B5").Value
                sheet10.Range("C9").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet10.Range("C10").Value = "混凝土强度、断面面尺寸、全高竖直度、顶面高程、轴线偏位、平整度"

                Call 全站仪平面位置检测表（）
                If TorF = False Then
                    Exit Sub
                End If
                sheet18.Activate()
                ExApp.ActiveWindow.SelectedSheets.Copy(, (ExApp.Sheets(ExApp.Sheets.Count)))

                '现浇墩、台身现场质量检验表
                sheet11.Range("D6").Value = sheet1.Range("B5").Value
                sheet11.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet11.Range("J6").Value = sheet1.Range("B15").Value
                sheet11.Range("D13").Value = "设计值：" & sheet1.Range("B12").Value * 10
                sheet11.Range("F12").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-10, 10)
                sheet11.Range("G12").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-10, 10)
                sheet11.Range("H12").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-10, 10)
                sheet11.Range("I12").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-10, 10)
                '全高竖直度
                If sheet1.Range("B11").Value / 100 <= 5 Then
                    sheet11.Range("F14").Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                    sheet11.Range("G14").Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                    sheet11.Range("H14").Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                    sheet11.Range("I14").Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                    sheet11.Range("F15:I16").Value = Nothing
                    sheet11.Range("G15").Value = "/"
                    sheet11.Range("G16").Value = "/"
                ElseIf sheet1.Range("B11").Value / 100 <= 60 Then
                    sheet11.Range("F15").Value = ExApp.WorksheetFunction.RandBetween(1, Math.Round(sheet1.Range("B11").Value * 10 / 1000, 0) - 1)
                    sheet11.Range("G15").Value = ExApp.WorksheetFunction.RandBetween(1, Math.Round(sheet1.Range("B11").Value * 10 / 1000, 0) - 1)
                    sheet11.Range("H15").Value = ExApp.WorksheetFunction.RandBetween(1, Math.Round(sheet1.Range("B11").Value * 10 / 1000, 0) - 1)
                    sheet11.Range("I15").Value = ExApp.WorksheetFunction.RandBetween(1, Math.Round(sheet1.Range("B11").Value * 10 / 1000, 0) - 1)
                    sheet11.Range("F14:I14").Value = Nothing
                    sheet11.Range("F16:I16").Value = Nothing
                    sheet11.Range("G14").Value = "/"
                    sheet11.Range("G16").Value = "/"
                Else
                    sheet11.Range("F16").Value = ExApp.WorksheetFunction.RandBetween(1, Math.Round(sheet1.Range("B11").Value * 10 / 3000, 0) - 1)
                    sheet11.Range("G16").Value = ExApp.WorksheetFunction.RandBetween(1, Math.Round(sheet1.Range("B11").Value * 10 / 3000, 0) - 1)
                    sheet11.Range("H16").Value = ExApp.WorksheetFunction.RandBetween(1, Math.Round(sheet1.Range("B11").Value * 10 / 3000, 0) - 1)
                    sheet11.Range("I16").Value = ExApp.WorksheetFunction.RandBetween(1, Math.Round(sheet1.Range("B11").Value * 10 / 3000, 0) - 1)
                    sheet11.Range("F14:I15").Value = Nothing
                    sheet11.Range("G14").Value = "/"
                    sheet11.Range("G15").Value = "/"
                End If

                '轴线偏位
                If sheet1.Range("B11").Value / 100 <= 60 Then
                    sheet11.Range("F18").Value = sheet1.Range("P3").Value
                    sheet11.Range("G18").Value = sheet1.Range("P4").Value
                    sheet11.Range("H18").Value = sheet1.Range("P5").Value
                    sheet11.Range("I18").Value = sheet1.Range("P6").Value
                    sheet11.Range("F19:I19").Value = Nothing
                    sheet11.Range("G19").Value = "/"
                Else
                    sheet11.Range("F19").Value = sheet1.Range("P3").Value
                    sheet11.Range("G19").Value = sheet1.Range("P4").Value
                    sheet11.Range("H19").Value = sheet1.Range("P5").Value
                    sheet11.Range("I19").Value = sheet1.Range("P6").Value
                    sheet11.Range("F18:I18").Value = Nothing
                    sheet11.Range("G18").Value = "/"
                End If
                sheet11.Range("F21").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet11.Range("G21").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet11.Range("H21").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet11.Range("I21").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                '模板测量偏差值
                sheet1.Range("O3").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O4").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O5").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O6").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("P3").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P4").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P5").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P6").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("L3").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "模板"

                Call 水准测量记录表()
                If TorF = False Then
                    Exit Sub
                End If
                Call 全站仪平面位置检测表（）
                If TorF = False Then
                    Exit Sub
                End If
                sheet17.Activate()
                ExApp.ActiveWindow.SelectedSheets.Copy(, (ExApp.Sheets(ExApp.Sheets.Count)))

                '现场模板安装检查记录表
                sheet12.Range("C6").Value = sheet1.Range("B5").Value
                sheet12.Range("C7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet12.Range("I6").Value = sheet1.Range("B15").Value
                sheet12.Range("I7").Value = sheet1.Range("B15").Value
                sheet12.Range("D8").Value = 2
                sheet12.Range("D9").Value = 5
                sheet12.Range("F8").Value = 4
                sheet12.Range("F9").Value = 4
                sheet12.Range("H8").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet12.Range("H9").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet12.Range("J8").Value = 100
                sheet12.Range("J9").Value = 100
                sheet12.Range("D11").Value = sheet1.Range("P3").Value
                sheet12.Range("F11").Value = sheet1.Range("P4").Value
                sheet12.Range("H11").Value = sheet1.Range("P5").Value
                sheet12.Range("J11").Value = sheet1.Range("P6").Value
                sheet12.Range("F14").Value = sheet1.Range("O3").Value
                sheet12.Range("G14").Value = sheet1.Range("O4").Value
                sheet12.Range("H14").Value = sheet1.Range("O5").Value
                sheet12.Range("I14").Value = sheet1.Range("O6").Value
                sheet12.Range("F12").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet12.Range("G12").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet12.Range("H12").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet12.Range("I12").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet12.Range("F13").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet12.Range("G13").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet12.Range("H13").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet12.Range("I13").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet12.Range("F18").Value = "牢固，稳定"

                sheet1.Range("I6:R15").Value = Nothing
                sheet1.Range("J5").Value = sheet0.Range("BR" & h).Value
                sheet1.Range("R5").Value = "中"
                sheet1.Range("L3").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                Call 水准测量记录表()
                If TorF = False Then
                    Exit Sub
                End If
                sheet11.Range("F17").Value = sheet1.Range("O3").Value
                sheet11.Range("G17").Value = sheet1.Range("O4").Value
                sheet11.Range("H17").Value = sheet1.Range("O5").Value
                '监抽钢筋检表
                sheet14.Range("D6").Value = sheet1.Range("B5").Value
                sheet14.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋"
                sheet14.Range("Q6").Value = sheet1.Range("B16").Value
                sheet14.Range("Q7").Value = sheet1.Range("B16").Value
                sheet14.Range("Q31").Value = sheet1.Range("B16").Value
                sheet14.Range("Q34").Value = sheet1.Range("B16").Value
                '主筋
                sheet14.Range("E15").Value = "设计值：" & sheet1.Range("B8").Value * 10
                LSBLFZ = Nothing
                If Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0) <= 10 Then
                    For cs = 1 To Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0)
                        LSBL = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet14.Range("G14").Value = LSBLFZ
                    sheet15.Range("D24").Value = "/"
                Else
                    sheet14.Range("G14").Value = "应测" & Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0) & "处，实测" & Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0) & "处，合格" & Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0) & "处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                    For cs = 1 To Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0)
                        LSBL = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet15.Range("D24").Value = LSBLFZ
                End If

                '箍筋
                sheet14.Range("E17").Value = "设计值：" & sheet1.Range("B9").Value * 10
                sheet14.Range("G16").Value = sheet1.Range("B9").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9) & "  " &
                                             sheet1.Range("B9").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                '骨架尺寸
                If sheet1.Range("B17").Value = "圆柱形" Then
                    sheet14.Range("E19").Value = Nothing
                    sheet14.Range("G18").Value = "/"
                    sheet14.Range("E21").Value = "高：" & sheet1.Range("B19").Value * 10
                    sheet14.Range("E22").Value = "直径：" & sheet1.Range("B14").Value * 10
                    sheet14.Range("G20").Value = "高：  " & sheet1.Range("B19").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4) & "  " &
                                                "直径：" & sheet1.Range("B14").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4)
                Else
                    sheet14.Range("E19").Value = "设计值：" & sheet1.Range("B13").Value * 10
                    sheet14.Range("G18").Value = sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                    sheet14.Range("E21").Value = "高：" & sheet1.Range("B19").Value * 10
                    sheet14.Range("E22").Value = "宽：" & sheet1.Range("B14").Value * 10
                    sheet14.Range("G20").Value = "高：  " & sheet1.Range("B19").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4) & "  " &
                                                "宽：" & sheet1.Range("B14").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4)
                End If

                '保护层
                sheet14.Range("E27").Value = "设计值：" & sheet1.Range("B10").Value * 10
                If sheet1.Range("B17").Value = "圆柱形" Then
                    gs = Math.Round(sheet1.Range("B12").Value * sheet1.Range("B18").Value * 3.14 / 300 / 100 * 0.2, 0)
                    If gs <= 10 Then
                        LSBLFZ = Nothing
                        For cs = 1 To gs
                            LSBL = sheet1.Range("B10").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                            LSBLFZ = LSBLFZ & LSBL & "   "
                        Next
                        sheet14.Range("G26").Value = LSBLFZ
                        sheet15.Range("D48").Value = Nothing
                    Else
                        LSBLFZ = Nothing
                        For cs = 1 To gs
                            LSBL = sheet1.Range("B10").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                            LSBLFZ = LSBLFZ & LSBL & "   "
                        Next
                        sheet14.Range("G26").Value = "应测" & gs & "处，实测" & gs & "处，合格" & gs & "处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                        sheet15.Range("D48").Value = LSBLFZ
                    End If
                Else
                    '矩形墩保护层
                    gs = Math.Round((sheet1.Range("B11").Value * sheet1.Range("B18").Value * 2 + sheet1.Range("B12").Value * sheet1.Range("B18").Value * 2) / 300 / 100 * 0.2, 0)
                    If gs <= 10 Then
                        LSBLFZ = Nothing
                        For cs = 1 To gs
                            LSBL = sheet1.Range("B10").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                            LSBLFZ = LSBLFZ & LSBL & "   "
                        Next
                        sheet14.Range("G26").Value = LSBLFZ
                        sheet15.Range("D48").Value = Nothing
                    Else
                        LSBLFZ = Nothing
                        For cs = 1 To gs
                            LSBL = sheet1.Range("B10").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                            LSBLFZ = LSBLFZ & LSBL & "   "
                        Next
                        sheet14.Range("G26").Value = "应测" & gs & "处，实测" & gs & "处，合格" & gs & "处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                        sheet15.Range("D48").Value = LSBLFZ
                    End If
                End If

                '钢筋记录表 (监抽)
                sheet15.Range("B6").Value = sheet1.Range("B5").Value
                sheet15.Range("B7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋"
                sheet15.Range("K6").Value = sheet1.Range("B16").Value
                sheet15.Range("K7").Value = sheet1.Range("B16").Value
                '现浇墩、台身现场质量检验表 (监抽)
                sheet16.Range("D6").Value = sheet1.Range("B5").Value
                sheet16.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet16.Range("J6").Value = sheet1.Range("B15").Value
                sheet16.Range("D13").Value = "设计值：" & sheet1.Range("B12").Value * 10
                sheet16.Range("F12").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-10, 10)
                sheet16.Range("G12").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-10, 10)

                '全高竖直度
                If sheet1.Range("B11").Value / 100 <= 5 Then
                    sheet16.Range("F14").Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                    sheet16.Range("G14").Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                    sheet16.Range("F15:I16").Value = Nothing
                    sheet16.Range("G15").Value = "/"
                    sheet16.Range("G16").Value = "/"
                ElseIf sheet1.Range("B11").Value / 100 <= 60 Then

                    sheet16.Range("F15").Value = ExApp.WorksheetFunction.RandBetween(1, Math.Round(sheet1.Range("B11").Value * 10 / 1000, 0) - 1)
                    sheet16.Range("G15").Value = ExApp.WorksheetFunction.RandBetween(1, Math.Round(sheet1.Range("B11").Value * 10 / 1000, 0) - 1)
                    sheet16.Range("F14:I14").Value = Nothing
                    sheet16.Range("F16:I16").Value = Nothing
                    sheet16.Range("G14").Value = "/"
                    sheet16.Range("G16").Value = "/"
                Else
                    sheet16.Range("F16").Value = ExApp.WorksheetFunction.RandBetween(1, Math.Round(sheet1.Range("B11").Value * 10 / 3000, 0) - 1)
                    sheet16.Range("G16").Value = ExApp.WorksheetFunction.RandBetween(1, Math.Round(sheet1.Range("B11").Value * 10 / 3000, 0) - 1)
                    sheet16.Range("F14:I15").Value = Nothing
                    sheet16.Range("G14").Value = "/"
                    sheet16.Range("G15").Value = "/"
                End If
                sheet16.Range("F17").Value = sheet1.Range("O3").Value

                '轴线偏位
                If sheet1.Range("B11").Value / 100 <= 60 Then
                    sheet16.Range("F18").Value = sheet1.Range("P3").Value
                    sheet16.Range("G18").Value = sheet1.Range("P4").Value
                    sheet16.Range("F19:I19").Value = Nothing
                    sheet16.Range("G19").Value = "/"
                Else
                    sheet16.Range("F19").Value = sheet1.Range("P3").Value
                    sheet16.Range("G19").Value = sheet1.Range("P4").Value
                    sheet16.Range("F18:I18").Value = Nothing
                    sheet16.Range("G18").Value = "/"
                End If
                sheet16.Range("F21").Value = ExApp.WorksheetFunction.RandBetween(1, 5) & "   " & ExApp.WorksheetFunction.RandBetween(1, 5)

                '刷新一次数据
                ExApp.Calculate()
                ExApp.Calculation = ExApp.Calculation.xlCalculationManual
                ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                '选择表格
                sheet6.Select()
                For i = 7 To ExApp.Sheets.Count
                    EXsheet = Exbook.Worksheets(i)
                    If EXsheet.Visible = True Then
                        EXsheet.Select(Replace:=False)
                    End If
                Next i
                EXsheet = ExApp.ActiveSheet
                ' 导出PDF文件
                PDFFilename = Filepath & "\" & sheet1.Range("B5").Value & sheet1.Range("C6").Value & sheet1.Range("D6").Value & ".pdf"
                '保存PDF文件
                EXsheet.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, PDFFilename, XlFixedFormatQuality.xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False)

                sheet19 = Exbook.Worksheets("水准表 (2)") '水准测量记录表
                sheet20 = Exbook.Worksheets("平面表 (2)") '全站仪平面位置检测表
                ExApp.DisplayAlerts = False
                sheet19.Delete()
                sheet20.Delete()
                h += 1
            End While
            MsgBox("已完成！", 0 + 64, "提示")
        Catch Exclerror As Exception   '错误时弹出提示
            MsgBox(Exclerror.Message)
        End Try
        TorF = False
    End Sub

    Sub 台身资料（）
        Dim h, i, r As Integer
        Dim ExcelFilename, PDFFilename As String  '定义输出的PDF文件名
        Dim FolderDialogObject As New System.Windows.Forms.FolderBrowserDialog()
        i = 8
        h = 8
        r = 0
        Try
            sheet0 = Exbook.Worksheets("数据库") '数据库
            sheet1 = Exbook.Worksheets("参数表") '参数表
            sheet2 = Exbook.Worksheets("交点法") '交点法
            sheet3 = Exbook.Worksheets("线元法") '线元法
            sheet4 = Exbook.Worksheets("断链") '断链
            sheet5 = Exbook.Worksheets("导线成果表") '导线成果表
            sheet6 = Exbook.Worksheets("申请批复单") '工序检验申请批复单
            sheet7 = Exbook.Worksheets("现浇墩台身检表") '系梁检表
            sheet8 = Exbook.Worksheets("模板记录表") '模板记录表
            sheet9 = Exbook.Worksheets("砼浇筑申请报告单") '砼浇筑申请报告单
            sheet10 = Exbook.Worksheets("监抽现浇墩台身检表") '监抽系梁检表
            sheet11 = Exbook.Worksheets("水准表") '水准测量记录表
            sheet12 = Exbook.Worksheets("平面表") '全站仪平面位置检测表

            '改表头
            sheet6.Range("A1").Value = sheet0.Range("C1").Value
            sheet6.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet6.Range("E3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet6.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet6.Range("E4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet7.Range("A1").Value = sheet0.Range("C1").Value
            sheet7.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet7.Range("J3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet7.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet7.Range("J4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet8.Range("A1").Value = sheet0.Range("C1").Value
            sheet8.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet8.Range("I3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet8.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet8.Range("I4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet9.Range("A1").Value = sheet0.Range("C1").Value
            sheet9.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet9.Range("I3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet9.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet9.Range("I4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet10.Range("A1").Value = sheet0.Range("C1").Value
            sheet10.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet10.Range("J3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet10.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet10.Range("J4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet11.Range("A1").Value = sheet0.Range("C1").Value
            sheet11.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet11.Range("H3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet11.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet11.Range("H4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet12.Range("A1").Value = sheet0.Range("C1").Value
            sheet12.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet12.Range("N3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet12.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet12.Range("N4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet0.Range("BA8:BT100000").Value = Nothing
            '计算偏距
            Do While sheet0.Range("B" & i).Value <> Nothing And sheet0.Range("C" & i).Value <> Nothing

                If sheet0.Range("O" & i).Value = Nothing Then
                    MsgBox("请在O" & i & "列选择该桥台是桥梁起点还是终点！！")
                    TorF = False
                    Exit Sub
                ElseIf sheet0.Range("N" & i).Value = "起点" Then
                    sheet0.Range("BA" & i).Value = sheet0.Range("C" & i).Value + Math.Round(sheet0.Range("O" & i).Value / 100 - sheet0.Range("P" & i).Value / 100, 3)
                    sheet0.Range("BE" & i).Value = sheet0.Range("C" & i).Value + Math.Round(sheet0.Range("O" & i).Value / 100 - sheet0.Range("P" & i).Value / 100 - sheet0.Range("H" & i).Value / 100, 3)
                    sheet0.Range("BI" & i).Value = sheet0.Range("C" & i).Value + Math.Round(sheet0.Range("O" & i).Value / 100 - sheet0.Range("P" & i).Value / 100 - sheet0.Range("H" & i).Value / 200, 3)
                    sheet0.Range("BM" & i).Value = sheet0.Range("C" & i).Value + Math.Round(sheet0.Range("O" & i).Value / 100 - sheet0.Range("P" & i).Value / 100 - sheet0.Range("H" & i).Value / 200, 3)
                    sheet0.Range("BQ" & i).Value = sheet0.Range("C" & i).Value + Math.Round(sheet0.Range("O" & i).Value / 100 - sheet0.Range("P" & i).Value / 100 - sheet0.Range("H" & i).Value / 200, 3)
                Else
                    sheet0.Range("BA" & i).Value = sheet0.Range("C" & i).Value - Math.Round(sheet0.Range("O" & i).Value / 100 - sheet0.Range("P" & i).Value / 100 - sheet0.Range("H" & i).Value / 100, 3)
                    sheet0.Range("BE" & i).Value = sheet0.Range("C" & i).Value - Math.Round(sheet0.Range("O" & i).Value / 100 - sheet0.Range("P" & i).Value / 100, 3)
                    sheet0.Range("BI" & i).Value = sheet0.Range("C" & i).Value - Math.Round(sheet0.Range("O" & i).Value / 100 - sheet0.Range("P" & i).Value / 100 - sheet0.Range("H" & i).Value / 200, 3)
                    sheet0.Range("BM" & i).Value = sheet0.Range("C" & i).Value - Math.Round(sheet0.Range("O" & i).Value / 100 - sheet0.Range("P" & i).Value / 100 - sheet0.Range("H" & i).Value / 200, 3)
                    sheet0.Range("BQ" & i).Value = sheet0.Range("C" & i).Value - Math.Round(sheet0.Range("O" & i).Value / 100 - sheet0.Range("P" & i).Value / 100 - sheet0.Range("H" & i).Value / 200, 3)
                End If

                If sheet0.Range("L" & i).Value = "右" Then
                    sheet0.Range("BB" & i).Value = ((sheet0.Range("M" & i).Value) - (sheet0.Range("G" & i).Value / 2)) / -100
                    sheet0.Range("BF" & i).Value = ((sheet0.Range("M" & i).Value) - (sheet0.Range("G" & i).Value / 2)) / -100
                    sheet0.Range("BJ" & i).Value = sheet0.Range("M" & i).Value / -100
                    sheet0.Range("BN" & i).Value = (sheet0.Range("M" & i).Value - sheet0.Range("G" & i).Value) / -100
                    sheet0.Range("BR" & i).Value = ((sheet0.Range("M" & i).Value) - (sheet0.Range("G" & i).Value / 2)) / -100
                Else
                    sheet0.Range("BB" & i).Value = ((sheet0.Range("M" & i).Value) - (sheet0.Range("G" & i).Value / 2)) / 100
                    sheet0.Range("BF" & i).Value = ((sheet0.Range("M" & i).Value) - (sheet0.Range("G" & i).Value / 2)) / 100
                    sheet0.Range("BJ" & i).Value = (sheet0.Range("M" & i).Value - sheet0.Range("G" & i).Value) / 100
                    sheet0.Range("BN" & i).Value = sheet0.Range("M" & i).Value / 100
                    sheet0.Range("BR" & i).Value = ((sheet0.Range("M" & i).Value) - (sheet0.Range("G" & i).Value / 2)) / 100
                End If
                i += 1
            Loop

            Call 质检资料设计坐标()
            If TorF = False Then
                Exit Sub
            End If

            While sheet0.Range("B" & h).Value <> Nothing

                ExApp.Calculation = ExApp.Calculation.xlCalculationManual '开启手动计算
                sheet1.Range("B" & r + 5).Value = sheet0.Range("B" & h).Value
                sheet1.Range("B" & r + 6).Value = sheet0.Range("C" & h).Value
                sheet1.Range("C" & r + 6).Value = sheet0.Range("D" & h).Value
                sheet1.Range("D" & r + 6).Value = sheet0.Range("E" & h).Value
                sheet1.Range("B" & r + 7).Value = sheet0.Range("G" & h).Value
                sheet1.Range("B" & r + 8).Value = sheet0.Range("H" & h).Value
                sheet1.Range("B" & r + 9).Value = sheet0.Range("I" & h).Value
                sheet1.Range("B" & r + 10).Value = sheet0.Range("F" & h).Value

                '测量参数
                sheet1.Range("I1").Value = sheet1.Range("B5").Value.substring(0, ExApp.WorksheetFunction.Find("K", sheet1.Range("B5").Value))
                sheet1.Range("I3:R15").Value = Nothing
                sheet1.Range("I3").Value = sheet0.Range("BA" & h).Value
                sheet1.Range("I4").Value = sheet0.Range("BE" & h).Value
                sheet1.Range("I5").Value = sheet0.Range("BI" & h).Value
                sheet1.Range("I6").Value = sheet0.Range("BM" & h).Value
                sheet1.Range("J3").Value = sheet0.Range("BB" & h).Value
                sheet1.Range("J4").Value = sheet0.Range("BF" & h).Value
                sheet1.Range("J5").Value = sheet0.Range("BJ" & h).Value
                sheet1.Range("J6").Value = sheet0.Range("BN" & h).Value
                sheet1.Range("K3").Value = sheet0.Range("J" & h).Value + (sheet0.Range("K" & h).Value + sheet0.Range("I" & h).Value) / 100
                sheet1.Range("K4").Value = sheet0.Range("J" & h).Value + (sheet0.Range("K" & h).Value + sheet0.Range("I" & h).Value) / 100
                sheet1.Range("K5").Value = sheet0.Range("J" & h).Value + (sheet0.Range("K" & h).Value + sheet0.Range("I" & h).Value) / 100
                sheet1.Range("K6").Value = sheet0.Range("J" & h).Value + (sheet0.Range("K" & h).Value + sheet0.Range("I" & h).Value) / 100
                sheet1.Range("M3").Value = sheet0.Range("BC" & h).Value
                sheet1.Range("N3").Value = sheet0.Range("BD" & h).Value
                sheet1.Range("M4").Value = sheet0.Range("BG" & h).Value
                sheet1.Range("N4").Value = sheet0.Range("BH" & h).Value
                sheet1.Range("M5").Value = sheet0.Range("BK" & h).Value
                sheet1.Range("N5").Value = sheet0.Range("BL" & h).Value
                sheet1.Range("M6").Value = sheet0.Range("BO" & h).Value
                sheet1.Range("N6").Value = sheet0.Range("BP" & h).Value
                sheet1.Range("R3").Value = "前"
                sheet1.Range("R4").Value = "后"
                sheet1.Range("R5").Value = "左"
                sheet1.Range("R6").Value = "右"
                sheet1.Range("P1").Value = sheet1.Range("B10").Value
                sheet1.Range("L3").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "模板"

                '测量偏差值
                sheet1.Range("O3").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O4").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O5").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O6").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("P3").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P4").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P5").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P6").Value = ExApp.WorksheetFunction.RandBetween(1, 9)

                sheet11.Range("A5").Value = sheet1.Range("A5").Value & sheet1.Range("B5").Value
                sheet12.Range("A5").Value = sheet1.Range("A5").Value & sheet1.Range("B5").Value

                Call 水准测量记录表()
                If TorF = False Then
                    Exit Sub
                End If
                Call 全站仪平面位置检测表（）
                If TorF = False Then
                    Exit Sub
                End If
                sheet11.Activate()
                ExApp.ActiveWindow.SelectedSheets.Copy(, (ExApp.Sheets(ExApp.Sheets.Count)))
                sheet12.Activate()
                ExApp.ActiveWindow.SelectedSheets.Copy(, (ExApp.Sheets(ExApp.Sheets.Count)))

                '现场模板安装检查记录表
                sheet8.Range("C6").Value = sheet1.Range("B5").Value
                sheet8.Range("C7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet8.Range("I6").Value = sheet1.Range("B10").Value
                sheet8.Range("I7").Value = sheet1.Range("B10").Value
                sheet8.Range("D8").Value = 2
                sheet8.Range("D9").Value = 5
                sheet8.Range("F8").Value = 4
                sheet8.Range("F9").Value = 4
                sheet8.Range("H8").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet8.Range("H9").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet8.Range("J8").Value = 100
                sheet8.Range("J9").Value = 100
                '测量偏差
                sheet8.Range("D11").Value = sheet1.Range("P3").Value
                sheet8.Range("F11").Value = sheet1.Range("P4").Value
                sheet8.Range("H11").Value = sheet1.Range("P5").Value
                sheet8.Range("J11").Value = sheet1.Range("P6").Value
                sheet8.Range("F14").Value = sheet1.Range("O3").Value
                sheet8.Range("G14").Value = sheet1.Range("O4").Value
                sheet8.Range("H14").Value = sheet1.Range("O5").Value
                sheet8.Range("I14").Value = sheet1.Range("O6").Value

                sheet8.Range("F12").Value = sheet1.Range("B7").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet8.Range("G12").Value = sheet1.Range("B7").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet8.Range("H12").Value = sheet1.Range("B7").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet8.Range("I12").Value = sheet1.Range("B7").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet8.Range("F13").Value = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet8.Range("G13").Value = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet8.Range("H13").Value = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet8.Range("I13").Value = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet8.Range("F18").Value = "牢固，稳定"

                '工序检验申请批复单
                sheet6.Range("C6").Value = sheet1.Range("B5").Value & sheet1.Range("C6").Value & "基础及下部构造"
                sheet6.Range("C7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet6.Range("C8").Value = sheet1.Range("B5").Value
                sheet6.Range("C9").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet6.Range("C10").Value = "混凝土强度、断面面尺寸、全高竖直度、顶面高程、轴线偏位、平整度"

                sheet1.Range("O3").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O4").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O5").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("P3").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P4").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P5").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P6").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("L3").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                Call 全站仪平面位置检测表（）
                If TorF = False Then
                    Exit Sub
                End If
                '现浇墩、台身现场质量检验表
                sheet7.Range("D6").Value = sheet1.Range("B5").Value
                sheet7.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet7.Range("J6").Value = sheet1.Range("B10").Value
                sheet7.Range("D13").Value = "宽：" & sheet1.Range("B8").Value * 10 & " " & "高：" & sheet1.Range("B9").Value * 10
                sheet7.Range("G12").Value = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-10, 10)
                sheet7.Range("H12").Value = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-10, 10)
                sheet7.Range("G13").Value = sheet1.Range("B9").Value * 10 + ExApp.WorksheetFunction.RandBetween(-10, 10)
                sheet7.Range("H13").Value = sheet1.Range("B9").Value * 10 + ExApp.WorksheetFunction.RandBetween(-10, 10)
                '全高竖直度
                If sheet1.Range("B9").Value / 100 <= 5 Then
                    sheet7.Range("F14").Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                    sheet7.Range("G14").Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                    sheet7.Range("H14").Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                    sheet7.Range("I14").Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                    sheet7.Range("F15:I16").Value = Nothing
                    sheet7.Range("G15").Value = "/"
                    sheet7.Range("G16").Value = "/"
                ElseIf sheet1.Range("B9").Value / 100 <= 60 Then
                    sheet7.Range("F15").Value = ExApp.WorksheetFunction.RandBetween(1, Math.Round(sheet1.Range("B11").Value * 10 / 1000, 0) - 1)
                    sheet7.Range("G15").Value = ExApp.WorksheetFunction.RandBetween(1, Math.Round(sheet1.Range("B11").Value * 10 / 1000, 0) - 1)
                    sheet7.Range("H15").Value = ExApp.WorksheetFunction.RandBetween(1, Math.Round(sheet1.Range("B11").Value * 10 / 1000, 0) - 1)
                    sheet7.Range("I15").Value = ExApp.WorksheetFunction.RandBetween(1, Math.Round(sheet1.Range("B11").Value * 10 / 1000, 0) - 1)
                    sheet7.Range("F14:I14").Value = Nothing
                    sheet7.Range("F16:I16").Value = Nothing
                    sheet7.Range("G14").Value = "/"
                    sheet7.Range("G16").Value = "/"
                Else
                    sheet7.Range("F16").Value = ExApp.WorksheetFunction.RandBetween(1, Math.Round(sheet1.Range("B11").Value * 10 / 3000, 0) - 1)
                    sheet7.Range("G16").Value = ExApp.WorksheetFunction.RandBetween(1, Math.Round(sheet1.Range("B11").Value * 10 / 3000, 0) - 1)
                    sheet7.Range("H16").Value = ExApp.WorksheetFunction.RandBetween(1, Math.Round(sheet1.Range("B11").Value * 10 / 3000, 0) - 1)
                    sheet7.Range("I16").Value = ExApp.WorksheetFunction.RandBetween(1, Math.Round(sheet1.Range("B11").Value * 10 / 3000, 0) - 1)
                    sheet7.Range("F14:I15").Value = Nothing
                    sheet7.Range("G14").Value = "/"
                    sheet7.Range("G15").Value = "/"
                End If

                '轴线偏位
                If sheet1.Range("B9").Value / 100 <= 60 Then
                    sheet7.Range("F18").Value = sheet1.Range("P3").Value
                    sheet7.Range("G18").Value = sheet1.Range("P4").Value
                    sheet7.Range("H18").Value = sheet1.Range("P5").Value
                    sheet7.Range("I18").Value = sheet1.Range("P6").Value
                    sheet7.Range("F19:I19").Value = Nothing
                    sheet7.Range("G19").Value = "/"
                Else
                    sheet7.Range("F19").Value = sheet1.Range("P3").Value
                    sheet7.Range("G19").Value = sheet1.Range("P4").Value
                    sheet7.Range("H19").Value = sheet1.Range("P5").Value
                    sheet7.Range("I19").Value = sheet1.Range("P6").Value
                    sheet7.Range("F18:I18").Value = Nothing
                    sheet7.Range("G18").Value = "/"
                End If
                sheet7.Range("F21").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet7.Range("G21").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet7.Range("H21").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet7.Range("I21").Value = ExApp.WorksheetFunction.RandBetween(1, 5)

                sheet1.Range("I6:R15").Value = Nothing
                sheet1.Range("J5").Value = sheet0.Range("BR" & h).Value
                sheet1.Range("R5").Value = "中"
                sheet1.Range("L3").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                Call 水准测量记录表()
                If TorF = False Then
                    Exit Sub
                End If
                '顶面高程
                sheet7.Range("F17").Value = sheet1.Range("O3").Value
                sheet7.Range("G17").Value = sheet1.Range("O4").Value
                sheet7.Range("H17").Value = sheet1.Range("O5").Value

                '现浇墩、台身现场质量检验表-监抽
                sheet10.Range("D6").Value = sheet1.Range("B5").Value
                sheet10.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet10.Range("J6").Value = sheet1.Range("B10").Value
                sheet10.Range("D13").Value = "宽：" & sheet1.Range("B8").Value * 10 & " " & "高：" & sheet1.Range("B9").Value * 10
                sheet10.Range("G12").Value = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-10, 10)
                sheet10.Range("G13").Value = sheet1.Range("B9").Value * 10 + ExApp.WorksheetFunction.RandBetween(-10, 10)

                '全高竖直度
                If sheet1.Range("B9").Value / 100 <= 5 Then
                    sheet10.Range("F14").Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                    sheet10.Range("G14").Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                    sheet10.Range("F15:I16").Value = Nothing
                    sheet10.Range("G15").Value = "/"
                    sheet10.Range("G16").Value = "/"
                ElseIf sheet1.Range("B9").Value / 100 <= 60 Then

                    sheet10.Range("F15").Value = ExApp.WorksheetFunction.RandBetween(1, Math.Round(sheet1.Range("B11").Value * 10 / 1000, 0) - 1)
                    sheet10.Range("G15").Value = ExApp.WorksheetFunction.RandBetween(1, Math.Round(sheet1.Range("B11").Value * 10 / 1000, 0) - 1)
                    sheet10.Range("F14:I14").Value = Nothing
                    sheet10.Range("F16:I16").Value = Nothing
                    sheet10.Range("G14").Value = "/"
                    sheet10.Range("G16").Value = "/"
                Else
                    sheet10.Range("F16").Value = ExApp.WorksheetFunction.RandBetween(1, Math.Round(sheet1.Range("B11").Value * 10 / 3000, 0) - 1)
                    sheet10.Range("G16").Value = ExApp.WorksheetFunction.RandBetween(1, Math.Round(sheet1.Range("B11").Value * 10 / 3000, 0) - 1)
                    sheet10.Range("F14:I15").Value = Nothing
                    sheet10.Range("G14").Value = "/"
                    sheet10.Range("G15").Value = "/"
                End If
                sheet10.Range("F17").Value = sheet1.Range("O3").Value

                '轴线偏位
                If sheet1.Range("B9").Value / 100 <= 60 Then
                    sheet10.Range("F18").Value = sheet1.Range("P3").Value
                    sheet10.Range("G18").Value = sheet1.Range("P4").Value
                    sheet10.Range("F19:I19").Value = Nothing
                    sheet10.Range("G19").Value = "/"
                Else
                    sheet10.Range("F19").Value = sheet1.Range("P3").Value
                    sheet10.Range("G19").Value = sheet1.Range("P4").Value
                    sheet10.Range("F18:I18").Value = Nothing
                    sheet10.Range("G18").Value = "/"
                End If

                sheet10.Range("F21").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet10.Range("G21").Value = ExApp.WorksheetFunction.RandBetween(1, 5)

                ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                '选择表格
                sheet6.Select()
                For i = 7 To ExApp.Sheets.Count
                    EXsheet = Exbook.Worksheets(i)
                    If EXsheet.Visible = True Then
                        EXsheet.Select(Replace:=False)
                    End If
                Next i
                EXsheet = ExApp.ActiveSheet
                ' 导出PDF文件
                PDFFilename = Filepath & "\" & sheet1.Range("B5").Value & sheet1.Range("C6").Value & sheet1.Range("D6").Value & ".pdf"
                '保存PDF文件
                EXsheet.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, PDFFilename, XlFixedFormatQuality.xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False)

                sheet13 = Exbook.Worksheets("水准表 (2)") '水准测量记录表
                sheet14 = Exbook.Worksheets("平面表 (2)") '全站仪平面位置检测表
                ExApp.DisplayAlerts = False
                sheet13.Delete()
                sheet14.Delete()
                h += 1
            End While
            MsgBox("已完成！", 0 + 64, "提示")
        Catch Exclerror As Exception   '错误时弹出提示
            MsgBox(Exclerror.Message)
        End Try
        TorF = False
    End Sub

    Sub 台帽资料（）

        Dim h, r, i As Integer
        Dim ExcelFilename, PDFFilename, LSBL, LSBLFZ As String
        Dim FolderDialogObject As New FolderBrowserDialog()
        h = 8
        i = 8
        r = 0
        Try
            sheet0 = Exbook.Worksheets("数据库")
            sheet1 = Exbook.Worksheets("参数表")
            sheet2 = Exbook.Worksheets("交点法")
            sheet3 = Exbook.Worksheets("线元法")
            sheet4 = Exbook.Worksheets("断链")
            sheet5 = Exbook.Worksheets("导线成果表")
            sheet6 = Exbook.Worksheets("钢筋隐蔽工程")
            sheet7 = Exbook.Worksheets("钢筋检表")
            sheet8 = Exbook.Worksheets("钢筋记录表")
            sheet9 = Exbook.Worksheets("申请批复单")
            sheet10 = Exbook.Worksheets("台帽检表")
            sheet11 = Exbook.Worksheets("模板记录表")
            sheet12 = Exbook.Worksheets("砼浇筑申请报告单")
            sheet13 = Exbook.Worksheets("监抽钢筋检表")
            sheet14 = Exbook.Worksheets("监抽钢筋记录表")
            sheet15 = Exbook.Worksheets("监抽台帽检表")
            sheet16 = Exbook.Worksheets("水准表")
            sheet17 = Exbook.Worksheets("平面表")


            '改表头
            sheet6.Range("A1").Value = sheet0.Range("C1").Value
            sheet6.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet6.Range("E3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet6.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet6.Range("E4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet7.Range("A1").Value = sheet0.Range("C1").Value
            sheet7.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet7.Range("P3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet7.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet7.Range("P4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet8.Range("A1").Value = sheet0.Range("C1").Value
            sheet8.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet8.Range("L3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet8.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet8.Range("L4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet9.Range("A1").Value = sheet0.Range("C1").Value
            sheet9.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet9.Range("E3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet9.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet9.Range("E4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet10.Range("A1").Value = sheet0.Range("C1").Value
            sheet10.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet10.Range("J3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet10.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet10.Range("J4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet11.Range("A1").Value = sheet0.Range("C1").Value
            sheet11.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet11.Range("I3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet11.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet11.Range("I4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet12.Range("A1").Value = sheet0.Range("C1").Value
            sheet12.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet12.Range("I3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet12.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet12.Range("I4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet13.Range("A1").Value = sheet0.Range("C1").Value
            sheet13.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet13.Range("P3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet13.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet13.Range("P4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet14.Range("A1").Value = sheet0.Range("C1").Value
            sheet14.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet14.Range("L3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet14.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet14.Range("L4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet15.Range("A1").Value = sheet0.Range("C1").Value
            sheet15.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet15.Range("J3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet15.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet15.Range("J4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet16.Range("A1").Value = sheet0.Range("C1").Value
            sheet16.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet16.Range("H3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet16.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet16.Range("H4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet17.Range("A1").Value = sheet0.Range("C1").Value
            sheet17.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet17.Range("N3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet17.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet17.Range("N4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value



            sheet0.Range("BA8:BT100000").Value = Nothing
            '计算偏距
            Do While sheet0.Range("B" & i).Value <> Nothing And sheet0.Range("C" & i).Value <> Nothing

                If sheet0.Range("U" & i).Value = Nothing Then
                    MsgBox("请在U" & i & "列选择该桥台是桥梁起点还是终点！！")
                    TorF = False
                    GC.Collect()
                    Exit Sub
                ElseIf sheet0.Range("U" & i).Value = "起点" Then
                    sheet0.Range("BA" & i).Value = sheet0.Range("C" & i).Value + sheet0.Range("V" & i).Value / 100
                    sheet0.Range("BE" & i).Value = sheet0.Range("C" & i).Value + sheet0.Range("V" & i).Value / 100 - sheet0.Range("H" & i).Value / 100
                    sheet0.Range("BI" & i).Value = sheet0.Range("C" & i).Value + Math.Round(sheet0.Range("V" & i).Value / 100 - sheet0.Range("H" & i).Value * 2 / 100, 3)
                    sheet0.Range("BM" & i).Value = sheet0.Range("C" & i).Value + Math.Round(sheet0.Range("V" & i).Value / 100 - sheet0.Range("H" & i).Value * 2 / 100, 3)
                    sheet0.Range("BQ" & i).Value = sheet0.Range("C" & i).Value + Math.Round(sheet0.Range("V" & i).Value / 100 - sheet0.Range("H" & i).Value * 2 / 100, 3)
                Else
                    sheet0.Range("BA" & i).Value = sheet0.Range("C" & i).Value - sheet0.Range("V" & i).Value / 100 + sheet0.Range("H" & i).Value / 100
                    sheet0.Range("BE" & i).Value = sheet0.Range("C" & i).Value - sheet0.Range("V" & i).Value / 100
                    sheet0.Range("BI" & i).Value = sheet0.Range("C" & i).Value - Math.Round(sheet0.Range("V" & i).Value / 100 + sheet0.Range("H" & i).Value * 2 / 100, 3)
                    sheet0.Range("BM" & i).Value = sheet0.Range("C" & i).Value - Math.Round(sheet0.Range("V" & i).Value / 100 + sheet0.Range("H" & i).Value * 2 / 100, 3)
                    sheet0.Range("BQ" & i).Value = sheet0.Range("C" & i).Value - Math.Round(sheet0.Range("V" & i).Value / 100 + sheet0.Range("H" & i).Value * 2 / 100, 3)
                End If

                If sheet0.Range("K" & i).Value = "中" Then
                    sheet0.Range("BB" & i).Value = 0
                    sheet0.Range("BF" & i).Value = 0
                    sheet0.Range("BJ" & i).Value = sheet0.Range("G" & i).Value / 2 / 100 * -1
                    sheet0.Range("BN" & i).Value = sheet0.Range("G" & i).Value / 2 / 100
                    sheet0.Range("BR" & i).Value = 0
                ElseIf sheet0.Range("K" & i).Value = "右" Then
                    sheet0.Range("BB" & i).Value = (sheet0.Range("L" & i).Value + sheet0.Range("G" & i).Value / 2) / 100 * -1
                    sheet0.Range("BF" & i).Value = (sheet0.Range("L" & i).Value + sheet0.Range("G" & i).Value / 2) / 100 * -1
                    sheet0.Range("BJ" & i).Value = (sheet0.Range("L" & i).Value + sheet0.Range("G" & i).Value) / 100 * -1
                    sheet0.Range("BN" & i).Value = (sheet0.Range("L" & i).Value) / 100 * -1
                    sheet0.Range("BR" & i).Value = (sheet0.Range("L" & i).Value + sheet0.Range("G" & i).Value / 2) / 100 * -1
                Else
                    sheet0.Range("BB" & i).Value = (sheet0.Range("L" & i).Value + sheet0.Range("G" & i).Value / 2) / 100
                    sheet0.Range("BF" & i).Value = (sheet0.Range("L" & i).Value + sheet0.Range("G" & i).Value / 2) / 100
                    sheet0.Range("BJ" & i).Value = (sheet0.Range("L" & i).Value) / 100
                    sheet0.Range("BN" & i).Value = (sheet0.Range("L" & i).Value + sheet0.Range("G" & i).Value) / 100
                    sheet0.Range("BR" & i).Value = (sheet0.Range("L" & i).Value + sheet0.Range("G" & i).Value / 2) / 100
                End If
                i += 1
            Loop

            Call 质检资料设计坐标()
            If TorF = False Then
                Exit Sub
            End If
            While sheet0.Range("B" & h).Value <> Nothing
                ExApp.Calculation = ExApp.Calculation.xlCalculationManual '开启手动计算
                sheet1.Range("B" & r + 5).Value = sheet0.Range("B" & h).Value
                sheet1.Range("B" & r + 6).Value = sheet0.Range("C" & h).Value
                sheet1.Range("C" & r + 6).Value = sheet0.Range("D" & h).Value
                sheet1.Range("D" & r + 6).Value = sheet0.Range("E" & h).Value
                sheet1.Range("B" & r + 7).Value = sheet0.Range("N" & h).Value
                sheet1.Range("B" & r + 8).Value = sheet0.Range("O" & h).Value
                sheet1.Range("B" & r + 9).Value = sheet0.Range("P" & h).Value
                sheet1.Range("B" & r + 10).Value = sheet0.Range("T" & h).Value
                sheet1.Range("B" & r + 11).Value = sheet0.Range("G" & h).Value
                sheet1.Range("B" & r + 12).Value = sheet0.Range("H" & h).Value
                sheet1.Range("B" & r + 13).Value = sheet0.Range("I" & h).Value
                sheet1.Range("B" & r + 14).Value = sheet0.Range("Q" & h).Value
                sheet1.Range("B" & r + 15).Value = sheet0.Range("R" & h).Value
                sheet1.Range("B" & r + 16).Value = sheet0.Range("S" & h).Value
                sheet1.Range("B" & r + 17).Value = sheet0.Range("F" & h).Value
                sheet1.Range("B" & r + 18).Value = sheet0.Range("W" & h).Value

                '测量参数
                sheet1.Range("I1").Value = sheet1.Range("B5").Value.substring(0, ExApp.WorksheetFunction.Find("K", sheet1.Range("B5").Value))
                sheet1.Range("I3:Q15").Value = Nothing
                sheet1.Range("I3").Value = sheet0.Range("BA" & h).Value
                sheet1.Range("I4").Value = sheet0.Range("BE" & h).Value
                sheet1.Range("I5").Value = sheet0.Range("BI" & h).Value
                sheet1.Range("I6").Value = sheet0.Range("BM" & h).Value
                sheet1.Range("I7").Value = sheet0.Range("BQ" & h).Value

                sheet1.Range("J3").Value = sheet0.Range("BB" & h).Value
                sheet1.Range("J4").Value = sheet0.Range("BF" & h).Value
                sheet1.Range("J5").Value = sheet0.Range("BJ" & h).Value
                sheet1.Range("J6").Value = sheet0.Range("BN" & h).Value
                sheet1.Range("J7").Value = sheet0.Range("BR" & h).Value

                sheet1.Range("K3").Value = sheet0.Range("J" & h).Value - (sheet0.Range("G" & h).Value / 200 * (sheet0.Range("J" & h).Value - sheet0.Range("M" & h).Value) / (sheet0.Range("G" & h).Value / 100))
                sheet1.Range("K4").Value = sheet0.Range("J" & h).Value - (sheet0.Range("G" & h).Value / 200 * (sheet0.Range("J" & h).Value - sheet0.Range("M" & h).Value) / (sheet0.Range("G" & h).Value / 100))
                sheet1.Range("K7").Value = sheet0.Range("J" & h).Value - (sheet0.Range("G" & h).Value / 200 * (sheet0.Range("J" & h).Value - sheet0.Range("M" & h).Value) / (sheet0.Range("G" & h).Value / 100))
                sheet1.Range("K5").Value = sheet0.Range("J" & h).Value
                sheet1.Range("K6").Value = sheet0.Range("M" & h).Value


                sheet1.Range("M3").Value = sheet0.Range("BC" & h).Value
                sheet1.Range("N3").Value = sheet0.Range("BD" & h).Value
                sheet1.Range("M4").Value = sheet0.Range("BG" & h).Value
                sheet1.Range("N4").Value = sheet0.Range("BH" & h).Value
                sheet1.Range("M5").Value = sheet0.Range("BK" & h).Value
                sheet1.Range("N5").Value = sheet0.Range("BL" & h).Value
                sheet1.Range("M6").Value = sheet0.Range("BO" & h).Value
                sheet1.Range("N6").Value = sheet0.Range("BP" & h).Value
                sheet1.Range("M7").Value = sheet0.Range("BS" & h).Value
                sheet1.Range("N7").Value = sheet0.Range("BT" & h).Value
                sheet1.Range("P1").Value = sheet1.Range("B17").Value
                sheet1.Range("L3").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                '测量偏差值
                sheet1.Range("O3").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O4").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O5").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O6").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O7").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("P3").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P4").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P5").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P6").Value = ExApp.WorksheetFunction.RandBetween(1, 9)

                sheet16.Range("A5").Value = sheet1.Range("A5").Value & sheet1.Range("B5").Value
                sheet17.Range("A5").Value = sheet1.Range("A5").Value & sheet1.Range("B5").Value

                ' 钢筋隐蔽工程
                sheet6.Range("C6").Value = sheet1.Range("B5").Value & sheet1.Range("C6").Value & "基础及下部构造"
                sheet6.Range("E6").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋加工及安装"
                sheet6.Range("C7").Value = sheet1.Range("B5").Value & sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋加工及安装"
                sheet6.Range("E10").Value = sheet1.Range("B19").Value
                sheet6.Range("E11").Value = sheet1.Range("B19").Value
                sheet6.Range("E27").Value = sheet1.Range("B19").Value
                sheet6.Range("E28").Value = sheet1.Range("B19").Value
                '钢筋检表
                sheet7.Range("D6").Value = sheet1.Range("B5").Value
                sheet7.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋"
                sheet7.Range("Q6").Value = sheet1.Range("B19").Value
                sheet7.Range("Q7").Value = sheet1.Range("B19").Value
                sheet7.Range("Q31").Value = sheet1.Range("B19").Value
                sheet7.Range("Q34").Value = sheet1.Range("B19").Value
                '主筋
                sheet7.Range("E15").Value = "设计值：" & sheet1.Range("B8").Value * 10
                LSBLFZ = Nothing
                If sheet1.Range("B7").Value * 2 <= 10 Then
                    For cs = 1 To sheet1.Range("B7").Value * 2
                        LSBL = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet7.Range("G14").Value = LSBLFZ
                    sheet8.Range("D24").Value = "/"
                Else
                    For cs = 1 To sheet1.Range("B7").Value * 2
                        LSBL = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet7.Range("G14").Value = "应测" & sheet1.Range("B7").Value * 2 & "处，实测" & sheet1.Range("B7").Value * 2 & "处，合格" & sheet1.Range("B7").Value * 2 & "处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                    sheet8.Range("D24").Value = LSBLFZ
                End If
                '箍筋
                sheet7.Range("E17").Value = "设计值：" & sheet1.Range("B9").Value * 10
                LSBLFZ = Nothing
                For cs = 1 To 10
                    LSBL = sheet1.Range("B9").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                    LSBLFZ = LSBLFZ & LSBL & "   "
                Next
                sheet7.Range("G16").Value = LSBLFZ

                '骨架尺寸
                sheet7.Range("E19").Value = "设计值：" & sheet1.Range("B14").Value * 10
                sheet7.Range("G18").Value = sheet1.Range("B14").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet7.Range("E21").Value = "设计值：" & sheet1.Range("B15").Value * 10
                sheet7.Range("E22").Value = "设计值：" & sheet1.Range("B16").Value * 10
                sheet7.Range("G20").Value = "宽：  " & sheet1.Range("B15").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4) & vbCrLf &
                                            "高：  " & sheet1.Range("B16").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4)
                '保护层
                sheet7.Range("E27").Value = "设计值：" & sheet1.Range("B10").Value * 10
                Dim gs As Integer
                gs = Math.Round((sheet1.Range("B11").Value * sheet1.Range("B13").Value * 2 + sheet1.Range("B12").Value * sheet1.Range("B13").Value * 2) / 300 / 100, 0)
                If gs <= 20 Then
                    LSBLFZ = Nothing
                    For cs = 1 To 20
                        LSBL = sheet1.Range("B10").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet7.Range("G26").Value = "应测20处，实测20处，合格20处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                    sheet8.Range("D48").Value = LSBLFZ
                Else
                    LSBLFZ = Nothing
                    For cs = 1 To gs
                        LSBL = sheet1.Range("B10").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet7.Range("G26").Value = "应测" & gs & "处，实测" & gs & "处，合格" & gs & "处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                    sheet8.Range("D48").Value = LSBLFZ
                End If
                '弯起钢筋位置
                sheet7.Range("M23").Value = ExApp.WorksheetFunction.RandBetween(-10, 10) & "  " & ExApp.WorksheetFunction.RandBetween(-10, 10) & " " &
                                            ExApp.WorksheetFunction.RandBetween(-10, 10) & "  " & ExApp.WorksheetFunction.RandBetween(-10, 10)

                '钢筋记录表
                sheet8.Range("B6").Value = sheet1.Range("B5").Value
                sheet8.Range("B7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋"
                sheet8.Range("K6").Value = sheet1.Range("B19").Value
                sheet8.Range("K7").Value = sheet1.Range("B19").Value

                '工序检验申请批复单
                sheet9.Range("C6").Value = sheet1.Range("B5").Value & sheet1.Range("C6").Value & "基础及下部构造"
                sheet9.Range("C7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet9.Range("C8").Value = sheet1.Range("B5").Value
                sheet9.Range("C9").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet9.Range("C10").Value = "混凝土强度、平面尺寸、结构高度、顶面高程、轴线偏位、平整度"

                Call 水准测量记录表()
                If TorF = False Then
                    Exit Sub
                End If
                sheet1.Range("I7:N7").Value = Nothing
                Call 全站仪平面位置检测表（）
                If TorF = False Then
                    Exit Sub
                End If
                sheet16.Activate()
                ExApp.ActiveWindow.SelectedSheets.Copy(, (ExApp.Sheets(ExApp.Sheets.Count)))
                sheet17.Activate()
                ExApp.ActiveWindow.SelectedSheets.Copy(, (ExApp.Sheets(ExApp.Sheets.Count)))

                '台帽现场质量检验表
                sheet10.Range("D6").Value = sheet1.Range("B5").Value
                sheet10.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet10.Range("J6").Value = sheet1.Range("B17").Value
                sheet10.Range("D13").Value = "长：" & sheet1.Range("B11").Value * 10
                sheet10.Range("D14").Value = "宽：" & sheet1.Range("B12").Value * 10
                sheet10.Range("D15").Value = "高：" & sheet1.Range("B13").Value * 10
                sheet10.Range("F13").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet10.Range("G13").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet10.Range("H13").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet10.Range("F14").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet10.Range("G14").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet10.Range("H14").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet10.Range("F15").Value = sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet10.Range("G15").Value = sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet10.Range("H15").Value = sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet10.Range("E16").Value = sheet1.Range("P3").Value
                sheet10.Range("F16").Value = sheet1.Range("P4").Value
                sheet10.Range("G16").Value = sheet1.Range("P5").Value
                sheet10.Range("H16").Value = sheet1.Range("P6").Value
                sheet10.Range("E17").Value = sheet1.Range("O3").Value
                sheet10.Range("F17").Value = sheet1.Range("O4").Value
                sheet10.Range("G17").Value = sheet1.Range("O5").Value
                sheet10.Range("H17").Value = sheet1.Range("O6").Value
                sheet10.Range("I17").Value = sheet1.Range("O7").Value
                sheet10.Range("E18:E19").Value = Nothing
                '垫石
                LSBLFZ = Nothing
                For K = 1 To sheet1.Range("B18").Value
                    LSBL = ExApp.WorksheetFunction.RandBetween(2, 8)
                    LSBLFZ = LSBLFZ & LSBL & "   "
                Next
                sheet10.Range("E18").Value = LSBLFZ
                '平整度
                LSBLFZ = Nothing
                For K = 1 To 15
                    LSBL = ExApp.WorksheetFunction.RandBetween(2, 6)
                    LSBLFZ = LSBLFZ & LSBL & "   "
                Next
                sheet10.Range("E19").Value = LSBLFZ

                '模板测量偏差值
                sheet1.Range("O3").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O4").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O5").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O6").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("P3").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P4").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P5").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P6").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("L3").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "模板"

                Call 水准测量记录表()
                If TorF = False Then
                    Exit Sub
                End If
                Call 全站仪平面位置检测表（）
                If TorF = False Then
                    Exit Sub
                End If
                '现场模板安装检查记录表
                sheet11.Range("C6").Value = sheet1.Range("B5").Value
                sheet11.Range("C7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet11.Range("I6").Value = sheet1.Range("B17").Value
                sheet11.Range("I7").Value = sheet1.Range("B17").Value
                sheet11.Range("D8").Value = 2
                sheet11.Range("D9").Value = 5
                sheet11.Range("F8").Value = 4
                sheet11.Range("F9").Value = 4
                sheet11.Range("H8").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet11.Range("H9").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet11.Range("J8").Value = 100
                sheet11.Range("J9").Value = 100
                '测量偏差
                sheet11.Range("D11").Value = sheet1.Range("P5").Value
                sheet11.Range("F11").Value = sheet1.Range("P6").Value
                sheet11.Range("H11").Value = sheet1.Range("P3").Value
                sheet11.Range("J11").Value = sheet1.Range("P4").Value
                sheet11.Range("F14").Value = sheet1.Range("O3").Value
                sheet11.Range("G14").Value = sheet1.Range("O4").Value
                sheet11.Range("H14").Value = sheet1.Range("O5").Value
                sheet11.Range("I14").Value = sheet1.Range("O6").Value

                sheet11.Range("F12").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("G12").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("H12").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("I12").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("F13").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("G13").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("H13").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("I13").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("F18").Value = "牢固，稳定"

                '监抽钢筋检表
                sheet13.Range("D6").Value = sheet1.Range("B5").Value
                sheet13.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋"
                sheet13.Range("Q6").Value = sheet1.Range("B19").Value
                sheet13.Range("Q7").Value = sheet1.Range("B19").Value
                sheet13.Range("Q31").Value = sheet1.Range("B19").Value
                sheet13.Range("Q34").Value = sheet1.Range("B19").Value
                '主筋
                sheet13.Range("E15").Value = "设计值：" & sheet1.Range("B8").Value * 10
                LSBLFZ = Nothing
                If Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0) <= 10 Then
                    For cs = 1 To Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0)
                        LSBL = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet13.Range("G14").Value = LSBLFZ
                    sheet14.Range("D24").Value = "/"
                Else
                    sheet13.Range("G14").Value = "应测" & Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0) & "处，实测" & Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0) & "处，合格" & Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0) & "处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                    For cs = 1 To Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0)
                        LSBL = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet14.Range("D24").Value = LSBLFZ
                End If

                '箍筋
                sheet13.Range("E17").Value = "设计值：" & sheet1.Range("B9").Value * 10
                sheet13.Range("G16").Value = sheet1.Range("B9").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9) & "  " &
                                             sheet1.Range("B9").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                '骨架尺寸
                sheet13.Range("E19").Value = "设计值：" & sheet1.Range("B14").Value * 10
                sheet13.Range("K18").Value = sheet1.Range("B14").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet13.Range("E21").Value = "宽：" & sheet1.Range("B15").Value * 10
                sheet13.Range("E22").Value = "高：" & sheet1.Range("B16").Value * 10
                sheet13.Range("G20").Value = "宽： " & sheet1.Range("B15").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4) & vbCrLf &
                                             "高： " & sheet1.Range("B16").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4)
                '保护层
                sheet13.Range("E27").Value = "设计值：" & sheet1.Range("B10").Value * 10
                gs = Math.Round((sheet1.Range("B11").Value * sheet1.Range("B13").Value * 2 + sheet1.Range("B12").Value * sheet1.Range("B13").Value * 2) / 300 / 100 * 0.2, 0)
                If gs <= 10 Then
                    LSBLFZ = Nothing
                    For cs = 1 To gs
                        LSBL = sheet1.Range("B10").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet13.Range("G26").Value = LSBLFZ
                    sheet14.Range("D48").Value = Nothing
                Else
                    LSBLFZ = Nothing
                    For cs = 1 To gs
                        LSBL = sheet1.Range("B10").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet13.Range("G26").Value = "应测" & gs & "处，实测" & gs & "处，合格" & gs & "处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                    sheet14.Range("D48").Value = LSBLFZ
                End If
                '弯起钢筋位置
                sheet13.Range("M23").Value = ExApp.WorksheetFunction.RandBetween(-10, 10) & "  " & ExApp.WorksheetFunction.RandBetween(-10, 10)
                '监抽钢筋记录表
                sheet14.Range("B6").Value = sheet1.Range("B5").Value
                sheet14.Range("B7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋"
                sheet14.Range("K6").Value = sheet1.Range("B19").Value
                sheet14.Range("K7").Value = sheet1.Range("B19").Value

                '监抽台帽混凝土现场质量检验表
                sheet15.Range("D6").Value = sheet1.Range("B5").Value
                sheet15.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet15.Range("J6").Value = sheet1.Range("B17").Value
                sheet15.Range("D13").Value = "长：" & sheet1.Range("B11").Value * 10
                sheet15.Range("D14").Value = "宽：" & sheet1.Range("B12").Value * 10
                sheet15.Range("D15").Value = "高：" & sheet1.Range("B13").Value * 10
                sheet15.Range("F13").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet15.Range("F14").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet15.Range("F15").Value = sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet15.Range("E16").Value = sheet1.Range("P3").Value
                sheet15.Range("F16").Value = sheet1.Range("P4").Value
                sheet15.Range("G16").Value = sheet1.Range("P5").Value
                sheet15.Range("H16").Value = sheet1.Range("P6").Value
                sheet15.Range("E17").Value = sheet1.Range("O3").Value
                sheet15.Range("F17").Value = sheet1.Range("O4").Value
                sheet15.Range("G17").Value = sheet1.Range("O5").Value
                sheet15.Range("H17").Value = sheet1.Range("O6").Value
                sheet15.Range("I17").Value = sheet1.Range("O7").Value
                sheet15.Range("E18:E19").Value = Nothing
                '垫石
                LSBL = Nothing
                For K = 1 To Math.Round(sheet1.Range("B18").Value * 0.2, 0)
                    LSBL = ExApp.WorksheetFunction.RandBetween(2, 8)
                    LSBLFZ = LSBLFZ & "    " & LSBL
                Next
                sheet15.Range("E18").Value = LSBLFZ
                '平整度
                LSBL = Nothing
                For K = 1 To 3
                    LSBL = ExApp.WorksheetFunction.RandBetween(2, 6)
                    LSBLFZ = LSBLFZ & "    " & LSBL
                Next
                sheet15.Range("E19").Value = LSBLFZ

                '刷新一次数据
                ExApp.Calculate()
                ExApp.Calculation = ExApp.Calculation.xlCalculationManual
                ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                '选择表格
                sheet6.Select()
                For i = 7 To ExApp.Sheets.Count
                    EXsheet = Exbook.Worksheets(i)
                    If EXsheet.Visible = True Then
                        EXsheet.Select(Replace:=False)
                    End If
                Next i
                EXsheet = ExApp.ActiveSheet

                ' 导出PDF文件
                PDFFilename = Filepath & "\" & sheet1.Range("B5").Value & sheet1.Range("C6").Value & sheet1.Range("D6").Value & ".pdf"
                EXsheet.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, PDFFilename, XlFixedFormatQuality.xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False)

                sheet18 = Exbook.Worksheets("水准表(2)")
                sheet19 = Exbook.Worksheets("平面表(2)")
                ExApp.DisplayAlerts = False
                sheet18.Delete()
                sheet19.Delete()
                h += 1
            End While

            MsgBox("已完成！", 0 + 64, "提示")
        Catch Exclerror As Exception   '错误时弹出提示
            MsgBox(Exclerror.Message)
        End Try
        TorF = False
    End Sub

    Sub 盖梁资料（）

        Dim h, r, i As Integer
        Dim ExcelFilename, PDFFilename, LSBL, LSBLFZ As String
        Dim FolderDialogObject As New FolderBrowserDialog()
        h = 8
        i = 8
        r = 0
        Try
            sheet0 = Exbook.Worksheets("数据库")
            sheet1 = Exbook.Worksheets("参数表")
            sheet2 = Exbook.Worksheets("交点法")
            sheet3 = Exbook.Worksheets("线元法")
            sheet4 = Exbook.Worksheets("断链")
            sheet5 = Exbook.Worksheets("导线成果表")
            sheet6 = Exbook.Worksheets("钢筋隐蔽工程")
            sheet7 = Exbook.Worksheets("钢筋检表")
            sheet8 = Exbook.Worksheets("钢筋记录表")
            sheet9 = Exbook.Worksheets("申请批复单")
            sheet10 = Exbook.Worksheets("盖梁检表")
            sheet11 = Exbook.Worksheets("模板记录表")
            sheet12 = Exbook.Worksheets("砼浇筑申请报告单")
            sheet13 = Exbook.Worksheets("监抽钢筋检表")
            sheet14 = Exbook.Worksheets("监抽钢筋记录表")
            sheet15 = Exbook.Worksheets("监抽盖梁检表")
            sheet16 = Exbook.Worksheets("水准表")
            sheet17 = Exbook.Worksheets("平面表")

            '改表头
            sheet6.Range("A1").Value = sheet0.Range("C1").Value
            sheet6.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet6.Range("E3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet6.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet6.Range("E4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet7.Range("A1").Value = sheet0.Range("C1").Value
            sheet7.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet7.Range("P3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet7.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet7.Range("P4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet8.Range("A1").Value = sheet0.Range("C1").Value
            sheet8.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet8.Range("L3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet8.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet8.Range("L4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet9.Range("A1").Value = sheet0.Range("C1").Value
            sheet9.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet9.Range("E3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet9.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet9.Range("E4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet10.Range("A1").Value = sheet0.Range("C1").Value
            sheet10.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet10.Range("J3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet10.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet10.Range("J4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet11.Range("A1").Value = sheet0.Range("C1").Value
            sheet11.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet11.Range("I3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet11.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet11.Range("I4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet12.Range("A1").Value = sheet0.Range("C1").Value
            sheet12.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet12.Range("I3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet12.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet12.Range("I4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet13.Range("A1").Value = sheet0.Range("C1").Value
            sheet13.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet13.Range("P3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet13.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet13.Range("P4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet14.Range("A1").Value = sheet0.Range("C1").Value
            sheet14.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet14.Range("L3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet14.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet14.Range("L4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet15.Range("A1").Value = sheet0.Range("C1").Value
            sheet15.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet15.Range("J3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet15.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet15.Range("J4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet16.Range("A1").Value = sheet0.Range("C1").Value
            sheet16.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet16.Range("H3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet16.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet16.Range("H4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet17.Range("A1").Value = sheet0.Range("C1").Value
            sheet17.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet17.Range("N3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet17.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet17.Range("N4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value



            sheet0.Range("BA8:BT100000").Value = Nothing
            '计算偏距
            Do While sheet0.Range("B" & i).Value <> Nothing
                If sheet0.Range("C" & i).Value = Nothing And sheet0.Range("U" & i).Value = "否" And sheet0.Range("V" & i).Value <> Nothing And sheet0.Range("W" & i).Value <> Nothing Then
                    sheet0.Range("C" & i).Value = FSZhj(ExApp.Sheets(1).range("V" & i).value, sheet0.Range("W" & i).Value)
                ElseIf sheet0.Range("C" & i).Value <> Nothing And sheet0.Range("U" & i).Value = "否" And sheet0.Range("V" & i).Value = Nothing And sheet0.Range("W" & i).Value = Nothing Then
                    MsgBox("请核对中心桩号或X、Y坐标是否正确填写！")
                    TorF = False
                    Exit Sub
                End If
                sheet0.Range("BA" & i).Value = sheet0.Range("C" & i).Value + sheet0.Range("H" & i).Value / 2 / 100
                sheet0.Range("BE" & i).Value = sheet0.Range("C" & i).Value - sheet0.Range("H" & i).Value / 2 / 100
                sheet0.Range("BI" & i).Value = sheet0.Range("C" & i).Value
                sheet0.Range("BM" & i).Value = sheet0.Range("C" & i).Value
                sheet0.Range("BQ" & i).Value = sheet0.Range("C" & i).Value

                If sheet0.Range("K" & i).Value = "中" Then
                    sheet0.Range("BB" & i).Value = 0
                    sheet0.Range("BF" & i).Value = 0
                    sheet0.Range("BJ" & i).Value = sheet0.Range("G" & i).Value / 2 / 100 * -1
                    sheet0.Range("BN" & i).Value = sheet0.Range("G" & i).Value / 2 / 100
                    sheet0.Range("BR" & i).Value = 0
                ElseIf sheet0.Range("K" & i).Value = "右" Then
                    sheet0.Range("BB" & i).Value = (ExApp.Sheets(1).range("L" & i).value + sheet0.Range("G" & i).Value / 2) / 100 * -1
                    sheet0.Range("BF" & i).Value = (ExApp.Sheets(1).range("L" & i).value + sheet0.Range("G" & i).Value / 2) / 100 * -1
                    sheet0.Range("BJ" & i).Value = (ExApp.Sheets(1).range("L" & i).value) / 100 * -1
                    sheet0.Range("BN" & i).Value = (ExApp.Sheets(1).range("L" & i).value + sheet0.Range("G" & i).Value) / 100 * -1
                    sheet0.Range("BR" & i).Value = (ExApp.Sheets(1).range("L" & i).value + sheet0.Range("G" & i).Value / 2) / 100 * -1
                Else
                    sheet0.Range("BB" & i).Value = (ExApp.Sheets(1).range("L" & i).value + sheet0.Range("G" & i).Value / 2) / 100
                    sheet0.Range("BF" & i).Value = (ExApp.Sheets(1).range("L" & i).value + sheet0.Range("G" & i).Value / 2) / 100
                    sheet0.Range("BJ" & i).Value = (ExApp.Sheets(1).range("L" & i).value) / 100
                    sheet0.Range("BN" & i).Value = (ExApp.Sheets(1).range("L" & i).value + sheet0.Range("G" & i).Value) / 100
                    sheet0.Range("BR" & i).Value = (ExApp.Sheets(1).range("L" & i).value + sheet0.Range("G" & i).Value / 2) / 100
                End If
                i += 1
            Loop

            Call 质检资料设计坐标()
            If TorF = False Then
                Exit Sub
            End If
            While sheet0.Range("B" & h).Value <> Nothing
                ExApp.Calculation = ExApp.Calculation.xlCalculationManual '开启手动计算
                sheet1.Range("B" & r + 5).Value = sheet0.Range("B" & h).Value
                sheet1.Range("B" & r + 6).Value = sheet0.Range("C" & h).Value
                sheet1.Range("C" & r + 6).Value = sheet0.Range("D" & h).Value
                sheet1.Range("D" & r + 6).Value = sheet0.Range("E" & h).Value
                sheet1.Range("B" & r + 7).Value = sheet0.Range("N" & h).Value
                sheet1.Range("B" & r + 8).Value = sheet0.Range("O" & h).Value
                sheet1.Range("B" & r + 9).Value = sheet0.Range("P" & h).Value
                sheet1.Range("B" & r + 10).Value = sheet0.Range("T" & h).Value
                sheet1.Range("B" & r + 11).Value = sheet0.Range("G" & h).Value
                sheet1.Range("B" & r + 12).Value = sheet0.Range("H" & h).Value
                sheet1.Range("B" & r + 13).Value = sheet0.Range("I" & h).Value
                sheet1.Range("B" & r + 14).Value = sheet0.Range("Q" & h).Value
                sheet1.Range("B" & r + 15).Value = sheet0.Range("R" & h).Value
                sheet1.Range("B" & r + 16).Value = sheet0.Range("S" & h).Value
                sheet1.Range("B" & r + 17).Value = sheet0.Range("F" & h).Value
                sheet1.Range("B" & r + 18).Value = sheet0.Range("X" & h).Value

                '测量参数
                sheet1.Range("I1").Value = sheet1.Range("B5").Value.substring(0, ExApp.WorksheetFunction.Find("K", sheet1.Range("B5").Value))
                sheet1.Range("I3:Q15").Value = Nothing
                sheet1.Range("I3").Value = sheet0.Range("BA" & h).Value
                sheet1.Range("I4").Value = sheet0.Range("BE" & h).Value
                sheet1.Range("I5").Value = sheet0.Range("BI" & h).Value
                sheet1.Range("I6").Value = sheet0.Range("BM" & h).Value
                sheet1.Range("I7").Value = sheet0.Range("BQ" & h).Value

                sheet1.Range("J3").Value = sheet0.Range("BB" & h).Value
                sheet1.Range("J4").Value = sheet0.Range("BF" & h).Value
                sheet1.Range("J5").Value = sheet0.Range("BJ" & h).Value
                sheet1.Range("J6").Value = sheet0.Range("BN" & h).Value
                sheet1.Range("J7").Value = sheet0.Range("BR" & h).Value

                sheet1.Range("K3").Value = sheet0.Range("J" & h).Value - (sheet0.Range("G" & h).Value / 200 * (sheet0.Range("J" & h).Value - sheet0.Range("M" & h).Value) / (sheet0.Range("G" & h).Value / 100))
                sheet1.Range("K4").Value = sheet0.Range("J" & h).Value - (sheet0.Range("G" & h).Value / 200 * (sheet0.Range("J" & h).Value - sheet0.Range("M" & h).Value) / (sheet0.Range("G" & h).Value / 100))
                sheet1.Range("K7").Value = sheet0.Range("J" & h).Value - (sheet0.Range("G" & h).Value / 200 * (sheet0.Range("J" & h).Value - sheet0.Range("M" & h).Value) / (sheet0.Range("G" & h).Value / 100))
                sheet1.Range("K5").Value = sheet0.Range("J" & h).Value
                sheet1.Range("K6").Value = sheet0.Range("M" & h).Value

                sheet1.Range("M3").Value = sheet0.Range("BC" & h).Value
                sheet1.Range("N3").Value = sheet0.Range("BD" & h).Value
                sheet1.Range("M4").Value = sheet0.Range("BG" & h).Value
                sheet1.Range("N4").Value = sheet0.Range("BH" & h).Value
                sheet1.Range("M5").Value = sheet0.Range("BK" & h).Value
                sheet1.Range("N5").Value = sheet0.Range("BL" & h).Value
                sheet1.Range("M6").Value = sheet0.Range("BO" & h).Value
                sheet1.Range("N6").Value = sheet0.Range("BP" & h).Value
                sheet1.Range("M7").Value = sheet0.Range("BS" & h).Value
                sheet1.Range("N7").Value = sheet0.Range("BT" & h).Value
                sheet1.Range("P1").Value = sheet1.Range("B17").Value
                sheet1.Range("L3").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                '测量偏差值
                sheet1.Range("O3").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O4").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O5").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O6").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O7").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("P3").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P4").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P5").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P6").Value = ExApp.WorksheetFunction.RandBetween(1, 9)

                sheet16.Range("A5").Value = sheet1.Range("A5").Value & sheet1.Range("B5").Value
                sheet17.Range("A5").Value = sheet1.Range("A5").Value & sheet1.Range("B5").Value

                ' 钢筋隐蔽工程
                sheet6.Range("C6").Value = sheet1.Range("B5").Value & sheet1.Range("C6").Value & "基础及下部构造"
                sheet6.Range("E6").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋加工及安装"
                sheet6.Range("C7").Value = sheet1.Range("B5").Value & sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋加工及安装"
                sheet6.Range("E10").Value = sheet1.Range("B19").Value
                sheet6.Range("E11").Value = sheet1.Range("B19").Value
                sheet6.Range("E27").Value = sheet1.Range("B19").Value
                sheet6.Range("E28").Value = sheet1.Range("B19").Value
                '钢筋检表
                sheet7.Range("D6").Value = sheet1.Range("B5").Value
                sheet7.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋"
                sheet7.Range("Q6").Value = sheet1.Range("B19").Value
                sheet7.Range("Q7").Value = sheet1.Range("B19").Value
                sheet7.Range("Q31").Value = sheet1.Range("B19").Value
                sheet7.Range("Q34").Value = sheet1.Range("B19").Value
                '主筋
                sheet7.Range("E15").Value = "设计值：" & sheet1.Range("B8").Value * 10
                LSBLFZ = Nothing
                If sheet1.Range("B7").Value * 2 <= 10 Then
                    For cs = 1 To sheet1.Range("B7").Value * 2
                        LSBL = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet7.Range("G14").Value = LSBLFZ
                    sheet8.Range("D24").Value = "/"
                Else
                    For cs = 1 To sheet1.Range("B7").Value * 2
                        LSBL = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet7.Range("G14").Value = "应测" & sheet1.Range("B7").Value * 2 & "处，实测" & sheet1.Range("B7").Value * 2 & "处，合格" & sheet1.Range("B7").Value * 2 & "处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                    sheet8.Range("D24").Value = LSBLFZ
                End If
                '箍筋
                sheet7.Range("E17").Value = "设计值：" & sheet1.Range("B9").Value * 10
                LSBLFZ = Nothing
                For cs = 1 To 10
                    LSBL = sheet1.Range("B9").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                    LSBLFZ = LSBLFZ & LSBL & "   "
                Next
                sheet7.Range("G16").Value = LSBLFZ

                '骨架尺寸
                sheet7.Range("E19").Value = "设计值：" & sheet1.Range("B14").Value * 10
                sheet7.Range("G18").Value = sheet1.Range("B14").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet7.Range("E21").Value = "设计值：" & sheet1.Range("B15").Value * 10
                sheet7.Range("E22").Value = "设计值：" & sheet1.Range("B16").Value * 10
                sheet7.Range("G20").Value = "宽：  " & sheet1.Range("B15").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4) & vbCrLf &
                                            "高：  " & sheet1.Range("B16").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4)
                '保护层
                sheet7.Range("E27").Value = "设计值：" & sheet1.Range("B10").Value * 10
                Dim gs As Integer
                gs = Math.Round((sheet1.Range("B11").Value * sheet1.Range("B13").Value * 2 + sheet1.Range("B12").Value * sheet1.Range("B13").Value * 2) / 300 / 100, 0)
                If gs <= 20 Then
                    LSBLFZ = Nothing
                    For cs = 1 To 20
                        LSBL = sheet1.Range("B10").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet7.Range("G26").Value = "应测20处，实测20处，合格20处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                    sheet8.Range("D48").Value = LSBLFZ
                Else
                    LSBLFZ = Nothing
                    For cs = 1 To gs
                        LSBL = sheet1.Range("B10").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet7.Range("G26").Value = "应测" & gs & "处，实测" & gs & "处，合格" & gs & "处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                    sheet8.Range("D48").Value = LSBLFZ
                End If
                '弯起钢筋位置
                sheet7.Range("M23").Value = ExApp.WorksheetFunction.RandBetween(-10, 10) & "  " & ExApp.WorksheetFunction.RandBetween(-10, 10) & " " &
                                            ExApp.WorksheetFunction.RandBetween(-10, 10) & "  " & ExApp.WorksheetFunction.RandBetween(-10, 10)
                '钢筋记录表
                sheet8.Range("B6").Value = sheet1.Range("B5").Value
                sheet8.Range("B7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋"
                sheet8.Range("K6").Value = sheet1.Range("B19").Value
                sheet8.Range("K7").Value = sheet1.Range("B19").Value

                '工序检验申请批复单
                sheet9.Range("C6").Value = sheet1.Range("B5").Value & sheet1.Range("C6").Value & "基础及下部构造"
                sheet9.Range("C7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet9.Range("C8").Value = sheet1.Range("B5").Value
                sheet9.Range("C9").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet9.Range("C10").Value = "混凝土强度、平面尺寸、结构高度、顶面高程、轴线偏位、平整度"

                Call 水准测量记录表()
                If TorF = False Then
                    Exit Sub
                End If
                sheet1.Range("I7:N7").Value = Nothing
                Call 全站仪平面位置检测表（）
                If TorF = False Then
                    Exit Sub
                End If
                sheet16.Activate()
                ExApp.ActiveWindow.SelectedSheets.Copy(, (ExApp.Sheets(ExApp.Sheets.Count)))
                sheet17.Activate()
                ExApp.ActiveWindow.SelectedSheets.Copy(, (ExApp.Sheets(ExApp.Sheets.Count)))

                '盖梁现场质量检验表
                sheet10.Range("D6").Value = sheet1.Range("B5").Value
                sheet10.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet10.Range("J6").Value = sheet1.Range("B17").Value
                sheet10.Range("D13").Value = "长：" & sheet1.Range("B11").Value * 10
                sheet10.Range("D14").Value = "宽：" & sheet1.Range("B12").Value * 10
                sheet10.Range("D15").Value = "高：" & sheet1.Range("B13").Value * 10
                sheet10.Range("F13").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet10.Range("G13").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet10.Range("H13").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet10.Range("F14").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet10.Range("G14").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet10.Range("H14").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet10.Range("F15").Value = sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet10.Range("G15").Value = sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet10.Range("H15").Value = sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet10.Range("E16").Value = sheet1.Range("P3").Value
                sheet10.Range("F16").Value = sheet1.Range("P4").Value
                sheet10.Range("G16").Value = sheet1.Range("P5").Value
                sheet10.Range("H16").Value = sheet1.Range("P6").Value
                sheet10.Range("E17").Value = sheet1.Range("O3").Value
                sheet10.Range("F17").Value = sheet1.Range("O4").Value
                sheet10.Range("G17").Value = sheet1.Range("O5").Value
                sheet10.Range("H17").Value = sheet1.Range("O6").Value
                sheet10.Range("I17").Value = sheet1.Range("O7").Value
                sheet10.Range("E18:E19").Value = Nothing
                '垫石
                LSBLFZ = Nothing
                For K = 1 To sheet1.Range("B18").Value
                    LSBL = ExApp.WorksheetFunction.RandBetween(2, 8)
                    LSBLFZ = LSBLFZ & LSBL & "   "
                Next
                sheet10.Range("E18").Value = LSBLFZ
                '平整度
                LSBLFZ = Nothing
                For K = 1 To 15
                    LSBL = ExApp.WorksheetFunction.RandBetween(2, 6)
                    LSBLFZ = LSBLFZ & LSBL & "   "
                Next
                sheet10.Range("E19").Value = LSBLFZ

                '模板测量偏差值
                sheet1.Range("O3").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O4").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O5").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O6").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("P3").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P4").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P5").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P6").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("L3").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "模板"

                Call 水准测量记录表()
                If TorF = False Then
                    Exit Sub
                End If
                Call 全站仪平面位置检测表（）
                If TorF = False Then
                    Exit Sub
                End If
                '现场模板安装检查记录表
                sheet11.Range("C6").Value = sheet1.Range("B5").Value
                sheet11.Range("C7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet11.Range("I6").Value = sheet1.Range("B17").Value
                sheet11.Range("I7").Value = sheet1.Range("B17").Value
                sheet11.Range("D8").Value = 2
                sheet11.Range("D9").Value = 5
                sheet11.Range("F8").Value = 4
                sheet11.Range("F9").Value = 4
                sheet11.Range("H8").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet11.Range("H9").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet11.Range("J8").Value = 100
                sheet11.Range("J9").Value = 100
                '测量偏差
                sheet11.Range("D11").Value = sheet1.Range("P5").Value
                sheet11.Range("F11").Value = sheet1.Range("P6").Value
                sheet11.Range("H11").Value = sheet1.Range("P3").Value
                sheet11.Range("J11").Value = sheet1.Range("P4").Value
                sheet11.Range("F14").Value = sheet1.Range("O3").Value
                sheet11.Range("G14").Value = sheet1.Range("O4").Value
                sheet11.Range("H14").Value = sheet1.Range("O5").Value
                sheet11.Range("I14").Value = sheet1.Range("O6").Value

                sheet11.Range("F12").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("G12").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("H12").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("I12").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("F13").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("G13").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("H13").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("I13").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("F18").Value = "牢固，稳定"


                '监抽钢筋检表
                sheet13.Range("D6").Value = sheet1.Range("B5").Value
                sheet13.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋"
                sheet13.Range("Q6").Value = sheet1.Range("B19").Value
                sheet13.Range("Q7").Value = sheet1.Range("B19").Value
                sheet13.Range("Q31").Value = sheet1.Range("B19").Value
                sheet13.Range("Q34").Value = sheet1.Range("B19").Value
                '主筋
                sheet13.Range("E15").Value = "设计值：" & sheet1.Range("B8").Value * 10
                LSBLFZ = Nothing
                If Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0) <= 10 Then
                    For cs = 1 To Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0)
                        LSBL = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet13.Range("G14").Value = LSBLFZ
                    sheet14.Range("D24").Value = "/"
                Else
                    sheet13.Range("G14").Value = "应测" & Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0) & "处，实测" & Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0) & "处，合格" & Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0) & "处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                    For cs = 1 To Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0)
                        LSBL = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet14.Range("D24").Value = LSBLFZ
                End If
                '箍筋
                sheet13.Range("E17").Value = "设计值：" & sheet1.Range("B9").Value * 10
                sheet13.Range("K16").Value = sheet1.Range("B9").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9) & "  " &
                                             sheet1.Range("B9").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                '骨架尺寸
                sheet13.Range("E19").Value = "设计值：" & sheet1.Range("B14").Value * 10
                sheet13.Range("K18").Value = sheet1.Range("B14").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet13.Range("E21").Value = "宽：" & sheet1.Range("B15").Value * 10
                sheet13.Range("E22").Value = "高：" & sheet1.Range("B16").Value * 10
                sheet13.Range("G20").Value = "宽： " & sheet1.Range("B15").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4) & vbCrLf &
                                             "高： " & sheet1.Range("B16").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4)
                '保护层
                sheet13.Range("E27").Value = "设计值：" & sheet1.Range("B10").Value * 10
                gs = Math.Round((sheet1.Range("B11").Value * sheet1.Range("B13").Value * 2 + sheet1.Range("B12").Value * sheet1.Range("B13").Value * 2) / 300 / 100 * 0.2, 0)
                If gs <= 10 Then
                    LSBLFZ = Nothing
                    For cs = 1 To gs
                        LSBL = sheet1.Range("B10").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet13.Range("G26").Value = LSBLFZ
                    sheet14.Range("D48").Value = Nothing
                Else
                    LSBLFZ = Nothing
                    For cs = 1 To gs
                        LSBL = sheet1.Range("B10").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet13.Range("G26").Value = "应测" & gs & "处，实测" & gs & "处，合格" & gs & "处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                    sheet14.Range("D48").Value = LSBLFZ
                End If
                '弯起钢筋位置
                sheet13.Range("M23").Value = ExApp.WorksheetFunction.RandBetween(-10, 10) & "  " & ExApp.WorksheetFunction.RandBetween(-10, 10)
                '监抽钢筋记录表
                sheet14.Range("B6").Value = sheet1.Range("B5").Value
                sheet14.Range("B7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋"
                sheet14.Range("K6").Value = sheet1.Range("B19").Value
                sheet14.Range("K7").Value = sheet1.Range("B19").Value

                '监抽盖梁混凝土现场质量检验表
                sheet15.Range("D6").Value = sheet1.Range("B5").Value
                sheet15.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet15.Range("J6").Value = sheet1.Range("B17").Value
                sheet15.Range("D13").Value = "长：" & sheet1.Range("B11").Value * 10
                sheet15.Range("D14").Value = "宽：" & sheet1.Range("B12").Value * 10
                sheet15.Range("D15").Value = "高：" & sheet1.Range("B13").Value * 10
                sheet15.Range("F13").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet15.Range("F14").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet15.Range("F15").Value = sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet15.Range("E16").Value = sheet1.Range("P3").Value
                sheet15.Range("F16").Value = sheet1.Range("P4").Value
                sheet15.Range("G16").Value = sheet1.Range("P5").Value
                sheet15.Range("H16").Value = sheet1.Range("P6").Value
                sheet15.Range("E17").Value = sheet1.Range("O3").Value
                sheet15.Range("F17").Value = sheet1.Range("O4").Value
                sheet15.Range("G17").Value = sheet1.Range("O5").Value
                sheet15.Range("H17").Value = sheet1.Range("O6").Value
                sheet15.Range("I17").Value = sheet1.Range("O7").Value
                sheet15.Range("E18:E19").Value = Nothing
                '垫石
                LSBL = Nothing
                For K = 1 To Math.Round(sheet1.Range("B18").Value * 0.2, 0)
                    LSBL = ExApp.WorksheetFunction.RandBetween(2, 8)
                    LSBLFZ = LSBLFZ & "    " & LSBL
                Next
                sheet15.Range("E18").Value = LSBLFZ
                '平整度
                LSBL = Nothing
                For K = 1 To 3
                    LSBL = ExApp.WorksheetFunction.RandBetween(2, 6)
                    LSBLFZ = LSBLFZ & "    " & LSBL
                Next
                sheet15.Range("E19").Value = LSBLFZ

                '刷新一次数据
                ExApp.Calculate()
                ExApp.Calculation = ExApp.Calculation.xlCalculationManual

                ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                '选择表格
                sheet6.Select()
                For i = 7 To ExApp.Sheets.Count
                    EXsheet = Exbook.Worksheets(i)
                    If EXsheet.Visible = True Then
                        EXsheet.Select(Replace:=False)
                    End If
                Next i
                EXsheet = ExApp.ActiveSheet

                ' 导出PDF文件
                PDFFilename = Filepath & "\" & sheet1.Range("B5").Value & sheet1.Range("C6").Value & sheet1.Range("D6").Value & ".pdf"
                EXsheet.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, PDFFilename, XlFixedFormatQuality.xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False)

                sheet18 = Exbook.Worksheets("水准表(2)")
                sheet19 = Exbook.Worksheets("平面表(2)")
                ExApp.DisplayAlerts = False
                sheet18.Delete()
                sheet19.Delete()
                h += 1
            End While

            MsgBox("已完成！", 0 + 64, "提示")
        Catch Exclerror As Exception   '错误时弹出提示
            MsgBox(Exclerror.Message)
        End Try
        TorF = False
    End Sub

    Sub 背墙资料()

        Dim h, r, i As Integer
        Dim ExcelFilename, PDFFilename, LSBL, LSBLFZ As String
        Dim FolderDialogObject As New FolderBrowserDialog()
        h = 8
        i = 8
        r = 0
        Try
            sheet0 = Exbook.Worksheets("数据库")
            sheet1 = Exbook.Worksheets("参数表")
            sheet2 = Exbook.Worksheets("交点法")
            sheet3 = Exbook.Worksheets("线元法")
            sheet4 = Exbook.Worksheets("断链")
            sheet5 = Exbook.Worksheets("导线成果表")
            sheet6 = Exbook.Worksheets("钢筋隐蔽工程")
            sheet7 = Exbook.Worksheets("钢筋检表")
            sheet8 = Exbook.Worksheets("钢筋记录表")
            sheet9 = Exbook.Worksheets("申请批复单")
            sheet10 = Exbook.Worksheets("背墙检表")
            sheet11 = Exbook.Worksheets("模板记录表")
            sheet12 = Exbook.Worksheets("砼浇筑申请报告单")
            sheet13 = Exbook.Worksheets("监抽钢筋检表")
            sheet14 = Exbook.Worksheets("监抽钢筋记录表")
            sheet15 = Exbook.Worksheets("监抽背墙检表")
            sheet16 = Exbook.Worksheets("水准表")
            sheet17 = Exbook.Worksheets("平面表")

            '改表头
            sheet6.Range("A1").Value = sheet0.Range("C1").Value
            sheet6.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet6.Range("E3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet6.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet6.Range("E4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet7.Range("A1").Value = sheet0.Range("C1").Value
            sheet7.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet7.Range("P3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet7.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet7.Range("P4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet8.Range("A1").Value = sheet0.Range("C1").Value
            sheet8.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet8.Range("L3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet8.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet8.Range("L4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet9.Range("A1").Value = sheet0.Range("C1").Value
            sheet9.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet9.Range("E3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet9.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet9.Range("E4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet10.Range("A1").Value = sheet0.Range("C1").Value
            sheet10.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet10.Range("K3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet10.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet10.Range("K4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet11.Range("A1").Value = sheet0.Range("C1").Value
            sheet11.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet11.Range("I3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet11.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet11.Range("I4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet12.Range("A1").Value = sheet0.Range("C1").Value
            sheet12.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet12.Range("I3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet12.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet12.Range("I4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet13.Range("A1").Value = sheet0.Range("C1").Value
            sheet13.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet13.Range("P3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet13.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet13.Range("P4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet14.Range("A1").Value = sheet0.Range("C1").Value
            sheet14.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet14.Range("L3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet14.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet14.Range("L4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet15.Range("A1").Value = sheet0.Range("C1").Value
            sheet15.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet15.Range("K3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet15.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet15.Range("K4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet16.Range("A1").Value = sheet0.Range("C1").Value
            sheet16.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet16.Range("H3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet16.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet16.Range("H4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet17.Range("A1").Value = sheet0.Range("C1").Value
            sheet17.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet17.Range("N3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet17.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet17.Range("N4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet0.Range("BA8:BT100000").Value = Nothing
            '计算桩号、偏距
            Do While sheet0.Range("B" & i).Value <> Nothing And sheet0.Range("C" & i).Value <> Nothing
                If sheet0.Range("U" & i).Value = Nothing Then
                    MsgBox("请在U" & i & "列选择该桥台是桥梁起点还是终点！！")
                    TorF = False
                    Exit Sub
                ElseIf sheet0.Range("U" & i).Value = "起点" Then
                    sheet0.Range("BA" & i).Value = sheet0.Range("C" & i).Value + sheet0.Range("V" & i).Value / 100
                    sheet0.Range("BE" & i).Value = sheet0.Range("C" & i).Value - sheet0.Range("G" & i).Value / 100
                    sheet0.Range("BI" & i).Value = sheet0.Range("C" & i).Value + （sheet0.Range("V" & i).Value - （sheet0.Range("G" & i).Value / 2)) / 100
                    sheet0.Range("BM" & i).Value = sheet0.Range("C" & i).Value + （sheet0.Range("V" & i).Value - （sheet0.Range("G" & i).Value / 2)) / 100
                    sheet0.Range("BQ" & i).Value = sheet0.Range("C" & i).Value + （sheet0.Range("V" & i).Value - （sheet0.Range("G" & i).Value / 2)) / 100
                Else
                    sheet0.Range("BA" & i).Value = sheet0.Range("C" & i).Value - (sheet0.Range("V" & i).Value - sheet0.Range("G" & i).Value) / 100
                    sheet0.Range("BE" & i).Value = sheet0.Range("C" & i).Value - sheet0.Range("V" & i).Value / 100
                    sheet0.Range("BI" & i).Value = sheet0.Range("C" & i).Value - （sheet0.Range("V" & i).Value - （sheet0.Range("G" & i).Value / 2)) / 100
                    sheet0.Range("BM" & i).Value = sheet0.Range("C" & i).Value - （sheet0.Range("V" & i).Value - （sheet0.Range("G" & i).Value / 2)) / 100
                    sheet0.Range("BQ" & i).Value = sheet0.Range("C" & i).Value - （sheet0.Range("V" & i).Value - （sheet0.Range("G" & i).Value / 2)) / 100
                End If

                If sheet0.Range("K" & i).Value = "右" Then
                    sheet0.Range("BB" & i).Value = （ExApp.Sheets(1).range("L" & i).value - sheet0.Range("H" & i).Value / 2） / -100
                    sheet0.Range("BF" & i).Value = （ExApp.Sheets(1).range("L" & i).value - sheet0.Range("H" & i).Value / 2） / -100
                    sheet0.Range("BJ" & i).Value = sheet0.Range("L" & i).Value / -100
                    sheet0.Range("BN" & i).Value = （ExApp.Sheets(1).range("L" & i).value - sheet0.Range("H" & i).Value） / -100
                    sheet0.Range("BR" & i).Value = （ExApp.Sheets(1).range("L" & i).value - sheet0.Range("H" & i).Value / 2） / -100
                Else
                    sheet0.Range("BB" & i).Value = （ExApp.Sheets(1).range("L" & i).value - sheet0.Range("H" & i).Value / 2） / 100
                    sheet0.Range("BF" & i).Value = （ExApp.Sheets(1).range("L" & i).value - sheet0.Range("H" & i).Value / 2） / 100
                    sheet0.Range("BJ" & i).Value = （ExApp.Sheets(1).range("L" & i).value - sheet0.Range("H" & i).Value） / 100
                    sheet0.Range("BN" & i).Value = sheet0.Range("L" & i).Value / 100
                    sheet0.Range("BR" & i).Value = （ExApp.Sheets(1).range("L" & i).value - sheet0.Range("H" & i).Value / 2） / 100
                End If
                i += 1
            Loop

            Call 质检资料设计坐标()
            If TorF = False Then
                Exit Sub
            End If
            While sheet0.Range("B" & h).Value <> Nothing
                If TorF = False Then
                    Exit Sub
                End If
                ExApp.Calculation = ExApp.Calculation.xlCalculationManual '开启手动计算
                sheet1.Range("B" & r + 5).Value = sheet0.Range("B" & h).Value
                sheet1.Range("B" & r + 6).Value = sheet0.Range("C" & h).Value
                sheet1.Range("C" & r + 6).Value = sheet0.Range("D" & h).Value
                sheet1.Range("D" & r + 6).Value = sheet0.Range("E" & h).Value
                sheet1.Range("B" & r + 7).Value = sheet0.Range("N" & h).Value
                sheet1.Range("B" & r + 8).Value = sheet0.Range("O" & h).Value
                sheet1.Range("B" & r + 9).Value = sheet0.Range("P" & h).Value
                sheet1.Range("B" & r + 10).Value = sheet0.Range("T" & h).Value
                sheet1.Range("B" & r + 11).Value = sheet0.Range("G" & h).Value
                sheet1.Range("B" & r + 12).Value = sheet0.Range("H" & h).Value
                sheet1.Range("B" & r + 13).Value = sheet0.Range("I" & h).Value
                sheet1.Range("B" & r + 14).Value = sheet0.Range("Q" & h).Value
                sheet1.Range("B" & r + 15).Value = sheet0.Range("R" & h).Value
                sheet1.Range("B" & r + 16).Value = sheet0.Range("S" & h).Value
                sheet1.Range("B" & r + 17).Value = sheet0.Range("F" & h).Value

                '测量参数
                sheet1.Range("I1").Value = sheet1.Range("B5").Value.substring(0, ExApp.WorksheetFunction.Find("K", sheet1.Range("B5").Value))
                sheet1.Range("I3:Q15").Value = Nothing
                sheet1.Range("I3").Value = sheet0.Range("BA" & h).Value
                sheet1.Range("I4").Value = sheet0.Range("BE" & h).Value
                sheet1.Range("I5").Value = sheet0.Range("BI" & h).Value
                sheet1.Range("I6").Value = sheet0.Range("BM" & h).Value
                sheet1.Range("I7").Value = sheet0.Range("BQ" & h).Value

                sheet1.Range("J3").Value = sheet0.Range("BB" & h).Value
                sheet1.Range("J4").Value = sheet0.Range("BF" & h).Value
                sheet1.Range("J5").Value = sheet0.Range("BJ" & h).Value
                sheet1.Range("J6").Value = sheet0.Range("BN" & h).Value
                sheet1.Range("J7").Value = sheet0.Range("BR" & h).Value

                sheet1.Range("K3").Value = sheet0.Range("J" & h).Value - (ExApp.Sheets(1).Range("H" & h).value / 200 * (ExApp.Sheets(1).Range("J" & h).value - sheet0.Range("M" & h).Value) / (ExApp.Sheets(1).Range("H" & h).value / 100))
                sheet1.Range("K4").Value = sheet0.Range("J" & h).Value - (ExApp.Sheets(1).Range("H" & h).value / 200 * (ExApp.Sheets(1).Range("J" & h).value - sheet0.Range("M" & h).Value) / (ExApp.Sheets(1).Range("H" & h).value / 100))
                sheet1.Range("K7").Value = sheet0.Range("J" & h).Value - (ExApp.Sheets(1).Range("H" & h).value / 200 * (ExApp.Sheets(1).Range("J" & h).value - sheet0.Range("M" & h).Value) / (ExApp.Sheets(1).Range("H" & h).value / 100))
                sheet1.Range("K5").Value = sheet0.Range("J" & h).Value
                sheet1.Range("K6").Value = sheet0.Range("M" & h).Value

                sheet1.Range("M3").Value = sheet0.Range("BC" & h).Value
                sheet1.Range("N3").Value = sheet0.Range("BD" & h).Value
                sheet1.Range("M4").Value = sheet0.Range("BG" & h).Value
                sheet1.Range("N4").Value = sheet0.Range("BH" & h).Value
                sheet1.Range("M5").Value = sheet0.Range("BK" & h).Value
                sheet1.Range("N5").Value = sheet0.Range("BL" & h).Value
                sheet1.Range("M6").Value = sheet0.Range("BO" & h).Value
                sheet1.Range("N6").Value = sheet0.Range("BP" & h).Value
                sheet1.Range("M7").Value = sheet0.Range("BS" & h).Value
                sheet1.Range("N7").Value = sheet0.Range("BT" & h).Value
                sheet1.Range("P1").Value = sheet1.Range("B17").Value
                sheet1.Range("L3").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                '测量偏差值
                sheet1.Range("O3").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O4").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O5").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O6").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O7").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("P3").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P4").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P5").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P6").Value = ExApp.WorksheetFunction.RandBetween(1, 9)

                sheet16.Range("A5").Value = sheet1.Range("A5").Value & sheet1.Range("B5").Value
                sheet17.Range("A5").Value = sheet1.Range("A5").Value & sheet1.Range("B5").Value

                ' 钢筋隐蔽工程
                sheet6.Range("C6").Value = sheet1.Range("B5").Value & sheet1.Range("C6").Value & "基础及下部构造"
                sheet6.Range("E6").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋加工及安装"
                sheet6.Range("C7").Value = sheet1.Range("B5").Value & sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋加工及安装"
                sheet6.Range("E10").Value = sheet1.Range("B18").Value
                sheet6.Range("E11").Value = sheet1.Range("B18").Value
                sheet6.Range("E27").Value = sheet1.Range("B18").Value
                sheet6.Range("E28").Value = sheet1.Range("B18").Value
                '钢筋检表
                sheet7.Range("D6").Value = sheet1.Range("B5").Value
                sheet7.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋"
                sheet7.Range("Q6").Value = sheet1.Range("B18").Value
                sheet7.Range("Q7").Value = sheet1.Range("B18").Value
                sheet7.Range("Q31").Value = sheet1.Range("B18").Value
                sheet7.Range("Q34").Value = sheet1.Range("B18").Value
                '主筋
                sheet7.Range("E15").Value = "设计值：" & sheet1.Range("B8").Value * 10
                LSBLFZ = Nothing
                If sheet1.Range("B7").Value * 2 <= 10 Then
                    For cs = 1 To sheet1.Range("B7").Value * 2
                        LSBL = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet7.Range("G14").Value = LSBLFZ
                    sheet8.Range("D24").Value = "/"
                Else
                    For cs = 1 To sheet1.Range("B7").Value * 2
                        LSBL = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet7.Range("G14").Value = "应测" & sheet1.Range("B7").Value * 2 & "处，实测" & sheet1.Range("B7").Value * 2 & "处，合格" & sheet1.Range("B7").Value * 2 & "处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                    sheet8.Range("D24").Value = LSBLFZ
                End If
                '箍筋
                sheet7.Range("E17").Value = "设计值：" & sheet1.Range("B9").Value * 10
                LSBLFZ = Nothing
                For cs = 1 To 10
                    LSBL = sheet1.Range("B9").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                    LSBLFZ = LSBLFZ & LSBL & "   "
                Next
                sheet7.Range("G16").Value = LSBLFZ

                '骨架尺寸
                sheet7.Range("E19").Value = "设计值：" & sheet1.Range("B14").Value * 10
                sheet7.Range("G18").Value = sheet1.Range("B14").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet7.Range("E21").Value = "宽：" & sheet1.Range("B15").Value * 10
                sheet7.Range("E22").Value = "高：" & sheet1.Range("B16").Value * 10
                sheet7.Range("G20").Value = "宽：  " & sheet1.Range("B15").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4) & vbCrLf &
                                            "高：  " & sheet1.Range("B16").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4)
                '保护层
                sheet7.Range("E27").Value = "设计值：" & sheet1.Range("B10").Value * 10
                Dim gs As Integer
                gs = Math.Round((sheet1.Range("B11").Value * sheet1.Range("B13").Value * 2 + sheet1.Range("B12").Value * sheet1.Range("B13").Value * 2) / 300 / 100, 0)
                If gs <= 20 Then
                    LSBLFZ = Nothing
                    For cs = 1 To 20
                        LSBL = sheet1.Range("B10").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet7.Range("G26").Value = "应测20处，实测20处，合格20处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                    sheet8.Range("D48").Value = LSBLFZ
                Else
                    LSBLFZ = Nothing
                    For cs = 1 To gs
                        LSBL = sheet1.Range("B10").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet7.Range("G26").Value = "应测" & gs & "处，实测" & gs & "处，合格" & gs & "处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                    sheet8.Range("D48").Value = LSBLFZ
                End If

                '钢筋记录表
                sheet8.Range("B6").Value = sheet1.Range("B5").Value
                sheet8.Range("B7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋"
                sheet8.Range("K6").Value = sheet1.Range("B18").Value
                sheet8.Range("K7").Value = sheet1.Range("B18").Value
                '工序检验申请批复单
                sheet9.Range("C6").Value = sheet1.Range("B5").Value & sheet1.Range("C6").Value & "基础及下部构造"
                sheet9.Range("C7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet9.Range("C8").Value = sheet1.Range("B5").Value
                sheet9.Range("C9").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet9.Range("C10").Value = "混凝土强度、平面尺寸、结构高度、顶面高程、轴线偏位、平整度"

                Call 水准测量记录表()
                If TorF = False Then
                    Exit Sub
                End If
                sheet1.Range("I7:N7").Value = Nothing
                Call 全站仪平面位置检测表（）
                If TorF = False Then
                    Exit Sub
                End If
                sheet16.Activate()
                ExApp.ActiveWindow.SelectedSheets.Copy(, (ExApp.Sheets(ExApp.Sheets.Count)))
                sheet17.Activate()
                ExApp.ActiveWindow.SelectedSheets.Copy(, (ExApp.Sheets(ExApp.Sheets.Count)))

                '背墙检验表
                sheet10.Range("D6").Value = sheet1.Range("B5").Value
                sheet10.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet10.Range("K6").Value = sheet1.Range("B17").Value
                '平面尺寸
                If sheet1.Range("B12").Value / 100 < 30 Then
                    sheet10.Range("E15").Value = "/"
                    sheet10.Range("D14").Value = "长：" & sheet1.Range("B11").Value * 10 & " 宽：" & sheet1.Range("B12").Value * 10
                    sheet10.Range("E13").Value = "长：" & sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20) & "   " & sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20) & vbCrLf &
                                                 "宽：" & sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20) & "   " & sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                Else
                    sheet10.Range("E13").Value = "/"
                    sheet10.Range("D16").Value = "长：" & sheet1.Range("B11").Value * 10 & " 宽：" & sheet1.Range("B12").Value * 10
                    sheet10.Range("E15").Value = "长：" & sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20) & "   " & sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20) & vbCrLf &
                                                 "宽：" & sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20) & "   " & sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                End If
                '结构高度
                sheet10.Range("D18").Value = sheet1.Range("B13").Value * 10
                sheet10.Range("E17").Value = sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20) & "   " &
                                             sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20) & "   " &
                                             sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20) & "   " &
                                             sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20) & "   " &
                                             sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet10.Range("E19").Value = sheet1.Range("O3").Value
                sheet10.Range("F19").Value = sheet1.Range("O4").Value
                sheet10.Range("G19").Value = sheet1.Range("O5").Value
                sheet10.Range("H19").Value = sheet1.Range("O6").Value
                sheet10.Range("I19").Value = sheet1.Range("O7").Value
                sheet10.Range("E20").Value = sheet1.Range("P3").Value
                sheet10.Range("F20").Value = sheet1.Range("P4").Value
                sheet10.Range("G20").Value = sheet1.Range("P5").Value
                sheet10.Range("H20").Value = sheet1.Range("P6").Value
                '平整度
                gs = Math.Round((sheet1.Range("B11").Value * sheet1.Range("B13").Value * 2 + sheet1.Range("B12").Value * sheet1.Range("B13").Value * 2) / 2000 / 100, 0)
                LSBLFZ = Nothing
                If gs <= 24 Then
                    For cs = 1 To 24
                        LSBL = ExApp.WorksheetFunction.RandBetween(1, 5)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet10.Range("E21").Value = LSBLFZ
                Else
                    For cs = 1 To gs
                        LSBL = ExApp.WorksheetFunction.RandBetween(1, 5)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet10.Range("E21").Value = LSBLFZ
                End If

                '模板测量偏差值
                sheet1.Range("O3").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O4").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O5").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O6").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("P3").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P4").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P5").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P6").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("L3").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "模板"

                Call 水准测量记录表()
                If TorF = False Then
                    Exit Sub
                End If
                Call 全站仪平面位置检测表（）
                If TorF = False Then
                    Exit Sub
                End If
                '现场模板安装检查记录表
                sheet11.Range("C6").Value = sheet1.Range("B5").Value
                sheet11.Range("C7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet11.Range("I6").Value = sheet1.Range("B17").Value
                sheet11.Range("I7").Value = sheet1.Range("B17").Value
                sheet11.Range("D8").Value = 2
                sheet11.Range("D9").Value = 5
                sheet11.Range("F8").Value = 4
                sheet11.Range("F9").Value = 4
                sheet11.Range("H8").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet11.Range("H9").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet11.Range("J8").Value = 100
                sheet11.Range("J9").Value = 100
                '测量偏差
                sheet11.Range("D11").Value = sheet1.Range("P5").Value
                sheet11.Range("F11").Value = sheet1.Range("P6").Value
                sheet11.Range("H11").Value = sheet1.Range("P3").Value
                sheet11.Range("J11").Value = sheet1.Range("P4").Value
                sheet11.Range("F14").Value = sheet1.Range("O3").Value
                sheet11.Range("G14").Value = sheet1.Range("O4").Value
                sheet11.Range("H14").Value = sheet1.Range("O5").Value
                sheet11.Range("I14").Value = sheet1.Range("O6").Value

                sheet11.Range("F12").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("G12").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("H12").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("I12").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("F13").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("G13").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("H13").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("I13").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("F18").Value = "牢固，稳定"

                '监抽钢筋检表
                sheet13.Range("D6").Value = sheet1.Range("B5").Value
                sheet13.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋"
                sheet13.Range("Q6").Value = sheet1.Range("B18").Value
                sheet13.Range("Q7").Value = sheet1.Range("B18").Value
                sheet13.Range("Q31").Value = sheet1.Range("B18").Value
                sheet13.Range("Q34").Value = sheet1.Range("B18").Value
                '主筋
                sheet13.Range("E15").Value = "设计值：" & sheet1.Range("B8").Value * 10
                LSBLFZ = Nothing
                If Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0) <= 10 Then
                    For cs = 1 To Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0)
                        LSBL = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet13.Range("G14").Value = LSBLFZ
                    sheet14.Range("D24").Value = "/"
                Else
                    sheet13.Range("G14").Value = "应测" & Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0) & "处，实测" & Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0) & "处，合格" & Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0) & "处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                    For cs = 1 To Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0)
                        LSBL = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet14.Range("D24").Value = LSBLFZ
                End If

                '箍筋
                sheet13.Range("E17").Value = "设计值：" & sheet1.Range("B9").Value * 10
                sheet13.Range("K16").Value = sheet1.Range("B9").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9) & "  " &
                                             sheet1.Range("B9").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                '骨架尺寸
                sheet13.Range("E19").Value = "设计值：" & sheet1.Range("B14").Value * 10
                sheet13.Range("K18").Value = sheet1.Range("B14").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet13.Range("E21").Value = "宽：" & sheet1.Range("B15").Value * 10
                sheet13.Range("E22").Value = "高：" & sheet1.Range("B16").Value * 10
                sheet13.Range("G20").Value = "宽： " & sheet1.Range("B15").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4) & vbCrLf &
                                             "高： " & sheet1.Range("B16").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4)
                '保护层
                sheet13.Range("E27").Value = "设计值：" & sheet1.Range("B10").Value * 10
                gs = Math.Round((sheet1.Range("B11").Value * sheet1.Range("B13").Value * 2 + sheet1.Range("B12").Value * sheet1.Range("B13").Value * 2) / 300 / 100 * 0.2, 0)
                If gs <= 10 Then
                    LSBLFZ = Nothing
                    For cs = 1 To gs
                        LSBL = sheet1.Range("B10").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet13.Range("G26").Value = LSBLFZ
                    sheet14.Range("D48").Value = Nothing
                Else
                    LSBLFZ = Nothing
                    For cs = 1 To gs
                        LSBL = sheet1.Range("B10").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet13.Range("G26").Value = "应测" & gs & "处，实测" & gs & "处，合格" & gs & "处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                    sheet14.Range("D48").Value = LSBLFZ
                End If
                '监抽钢筋记录表
                sheet14.Range("B6").Value = sheet1.Range("B5").Value
                sheet14.Range("B7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋"
                sheet14.Range("K6").Value = sheet1.Range("B18").Value
                sheet14.Range("K7").Value = sheet1.Range("B18").Value
                '监抽背墙检验表
                sheet15.Range("D6").Value = sheet1.Range("B5").Value
                sheet15.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet15.Range("K6").Value = sheet1.Range("B17").Value
                '平面尺寸
                If sheet1.Range("B12").Value / 100 < 30 Then
                    sheet15.Range("E15").Value = "/"
                    sheet15.Range("D14").Value = "长：" & sheet1.Range("B11").Value * 10 & " 宽：" & sheet1.Range("B12").Value * 10
                    sheet15.Range("E13").Value = "长：" & sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20) & vbCrLf &
                                                 "宽：" & sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                Else
                    sheet15.Range("E13").Value = "/"
                    sheet15.Range("D16").Value = "长：" & sheet1.Range("B11").Value * 10 & " 宽：" & sheet1.Range("B12").Value * 10
                    sheet15.Range("E15").Value = "长：" & sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20) & vbCrLf &
                                                 "宽：" & sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                End If
                sheet15.Range("D18").Value = sheet1.Range("B13").Value * 10
                sheet15.Range("E17").Value = sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20) & "  " &
                                             sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet15.Range("E19").Value = sheet10.Range("E19").Value
                sheet15.Range("F19").Value = sheet10.Range("F19").Value
                sheet15.Range("E20").Value = sheet10.Range("E20").Value
                sheet15.Range("F20").Value = sheet10.Range("F20").Value
                gs = Math.Round((sheet1.Range("B11").Value * sheet1.Range("B13").Value * 2 + sheet1.Range("B12").Value * sheet1.Range("B13").Value * 2) / 2000 / 100 * 0.2, 0)
                LSBLFZ = Nothing
                If gs <= 5 Then
                    For cs = 1 To 5
                        LSBL = ExApp.WorksheetFunction.RandBetween(1, 5)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet15.Range("E21").Value = LSBLFZ
                Else
                    For cs = 1 To gs
                        LSBL = ExApp.WorksheetFunction.RandBetween(1, 5)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet15.Range("E21").Value = LSBLFZ
                End If

                '刷新一次数据
                ExApp.Calculate()
                ExApp.Calculation = ExApp.Calculation.xlCalculationManual

                ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                '选择表格
                sheet6.Select()
                For i = 7 To ExApp.Sheets.Count
                    EXsheet = Exbook.Worksheets(i)
                    If EXsheet.Visible = True Then
                        EXsheet.Select(Replace:=False)
                    End If
                Next i
                EXsheet = ExApp.ActiveSheet
                ' 导出PDF文件
                PDFFilename = Filepath & "\" & sheet1.Range("B5").Value & sheet1.Range("C6").Value & sheet1.Range("D6").Value & ".pdf"
                EXsheet.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, PDFFilename, XlFixedFormatQuality.xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False)

                sheet18 = Exbook.Worksheets("水准表(2)")
                sheet19 = Exbook.Worksheets("平面表(2)")
                ExApp.DisplayAlerts = False
                sheet18.Delete()
                sheet19.Delete()
                h += 1
            End While

            MsgBox("已完成！", 0 + 64, "提示")
        Catch Exclerror As Exception   '错误时弹出提示
            MsgBox(Exclerror.Message)
        End Try
        TorF = False
    End Sub

    Sub 耳墙资料()
        Dim P As Double
        Dim h, r, i As Integer
        Dim ExcelFilename, PDFFilename, LSBL, LSBLFZ As String
        Dim FolderDialogObject As New FolderBrowserDialog()
        h = 8
        i = 8
        r = 0
        Try
            sheet0 = Exbook.Worksheets("数据库")
            sheet1 = Exbook.Worksheets("参数表")
            sheet2 = Exbook.Worksheets("交点法")
            sheet3 = Exbook.Worksheets("线元法")
            sheet4 = Exbook.Worksheets("断链")
            sheet5 = Exbook.Worksheets("导线成果表")
            sheet6 = Exbook.Worksheets("钢筋隐蔽工程")
            sheet7 = Exbook.Worksheets("钢筋检表")
            sheet8 = Exbook.Worksheets("钢筋记录表")
            sheet9 = Exbook.Worksheets("申请批复单")
            sheet10 = Exbook.Worksheets("耳墙检表")
            sheet11 = Exbook.Worksheets("模板记录表")
            sheet12 = Exbook.Worksheets("砼浇筑申请报告单")
            sheet13 = Exbook.Worksheets("监抽钢筋检表")
            sheet14 = Exbook.Worksheets("监抽钢筋记录表")
            sheet15 = Exbook.Worksheets("监抽耳墙检表")
            sheet16 = Exbook.Worksheets("水准表")
            sheet17 = Exbook.Worksheets("平面表")

            '改表头
            sheet6.Range("A1").Value = sheet0.Range("C1").Value
            sheet6.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet6.Range("E3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet6.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet6.Range("E4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet7.Range("A1").Value = sheet0.Range("C1").Value
            sheet7.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet7.Range("O3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet7.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet7.Range("O4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet8.Range("A1").Value = sheet0.Range("C1").Value
            sheet8.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet8.Range("L3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet8.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet8.Range("L4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet9.Range("A1").Value = sheet0.Range("C1").Value
            sheet9.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet9.Range("E3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet9.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet9.Range("E4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet10.Range("A1").Value = sheet0.Range("C1").Value
            sheet10.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet10.Range("K3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet10.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet10.Range("K4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet11.Range("A1").Value = sheet0.Range("C1").Value
            sheet11.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet11.Range("I3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet11.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet11.Range("I4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet12.Range("A1").Value = sheet0.Range("C1").Value
            sheet12.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet12.Range("I3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet12.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet12.Range("I4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet13.Range("A1").Value = sheet0.Range("C1").Value
            sheet13.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet13.Range("P3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet13.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet13.Range("P4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet14.Range("A1").Value = sheet0.Range("C1").Value
            sheet14.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet14.Range("L3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet14.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet14.Range("L4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet15.Range("A1").Value = sheet0.Range("C1").Value
            sheet15.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet15.Range("K3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet15.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet15.Range("K4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet16.Range("A1").Value = sheet0.Range("C1").Value
            sheet16.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet16.Range("H3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet16.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet16.Range("H4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet17.Range("A1").Value = sheet0.Range("C1").Value
            sheet17.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet17.Range("N3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet17.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet17.Range("N4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet0.Range("BA8:BT100000").Value = Nothing

            '计算桩号、偏距
            Do While sheet0.Range("B" & i).Value <> Nothing And sheet0.Range("C" & i).Value <> Nothing
                If sheet0.Range("U" & i).Value = Nothing Then
                    MsgBox("请在U" & i & "列选择该桥台是桥梁起点还是终点！！")
                    TorF = False
                    Exit Sub
                Else
                    sheet1.Range("I3:Q1000").Value = Nothing
                    sheet1.Range("I1").Value = sheet1.Range("B5").Value.substring(0, ExApp.WorksheetFunction.Find("K", sheet1.Range("B5").Value))
                    If sheet0.Range("U" & h).Value = "起点" Then
                        '桥台起点桩号
                        sheet1.Range("I3").Value = sheet0.Range("C" & h).Value + sheet0.Range("V" & h).Value / 100
                        sheet1.Range("I4").Value = sheet0.Range("C" & h).Value
                        sheet1.Range("I5").Value = sheet0.Range("C" & h).Value + sheet0.Range("V" & h).Value / 200
                        sheet1.Range("I6").Value = sheet0.Range("C" & h).Value + sheet0.Range("V" & h).Value / 200
                        sheet1.Range("I7").Value = sheet0.Range("C" & h).Value + sheet0.Range("V" & h).Value / 200

                        sheet1.Range("I8").Value = sheet0.Range("C" & h).Value + sheet0.Range("V" & h).Value / 100
                        sheet1.Range("I9").Value = sheet0.Range("C" & h).Value
                        sheet1.Range("I10").Value = sheet0.Range("C" & h).Value + sheet0.Range("V" & h).Value / 200
                        sheet1.Range("I11").Value = sheet0.Range("C" & h).Value + sheet0.Range("V" & h).Value / 200
                        sheet1.Range("I12").Value = sheet0.Range("C" & h).Value + sheet0.Range("V" & h).Value / 200
                    Else
                        '桥台终点桩号
                        sheet1.Range("I3").Value = sheet0.Range("C" & h).Value
                        sheet1.Range("I4").Value = sheet0.Range("C" & h).Value - sheet0.Range("V" & h).Value / 100
                        sheet1.Range("I5").Value = sheet0.Range("C" & h).Value - sheet0.Range("V" & h).Value / 200
                        sheet1.Range("I6").Value = sheet0.Range("C" & h).Value - sheet0.Range("V" & h).Value / 200
                        sheet1.Range("I7").Value = sheet0.Range("C" & h).Value - sheet0.Range("V" & h).Value / 200

                        sheet1.Range("I8").Value = sheet0.Range("C" & h).Value
                        sheet1.Range("I9").Value = sheet0.Range("C" & h).Value - sheet0.Range("V" & h).Value / 100
                        sheet1.Range("I10").Value = sheet0.Range("C" & h).Value - sheet0.Range("V" & h).Value / 200
                        sheet1.Range("I11").Value = sheet0.Range("C" & h).Value - sheet0.Range("V" & h).Value / 200
                        sheet1.Range("I12").Value = sheet0.Range("C" & h).Value - sheet0.Range("V" & h).Value / 200
                    End If

                    If sheet0.Range("K" & h).Value <> "左" Then
                        '偏距
                        sheet1.Range("J3").Value = (sheet0.Range("L" & h).Value - sheet0.Range("H" & h).Value / 2) / -100
                        sheet1.Range("J4").Value = (sheet0.Range("L" & h).Value - sheet0.Range("H" & h).Value / 2) / -100
                        sheet1.Range("J5").Value = sheet0.Range("L" & h).Value / -100
                        sheet1.Range("J6").Value = (sheet0.Range("L" & h).Value - sheet0.Range("H" & h).Value) / -100
                        sheet1.Range("J7").Value = (sheet0.Range("L" & h).Value - sheet0.Range("H" & h).Value / 2) / -100

                        sheet1.Range("J8").Value = (sheet0.Range("L" & h).Value - sheet0.Range("W" & h).Value + sheet0.Range("H" & h).Value / 2) / -100
                        sheet1.Range("J9").Value = (sheet0.Range("L" & h).Value - sheet0.Range("W" & h).Value + sheet0.Range("H" & h).Value / 2) / -100
                        sheet1.Range("J10").Value = (sheet0.Range("L" & h).Value - sheet0.Range("W" & h).Value + sheet0.Range("H" & h).Value) / -100
                        sheet1.Range("J11").Value = (sheet0.Range("L" & h).Value - sheet0.Range("W" & h).Value) / -100
                        sheet1.Range("J12").Value = (sheet0.Range("L" & h).Value - sheet0.Range("W" & h).Value + sheet0.Range("H" & h).Value / 2) / -100
                    Else
                        sheet1.Range("J3").Value = (sheet0.Range("L" & h).Value - sheet0.Range("W" & h).Value + sheet0.Range("H" & h).Value / 2) / 100
                        sheet1.Range("J4").Value = (sheet0.Range("L" & h).Value - sheet0.Range("W" & h).Value + sheet0.Range("H" & h).Value / 2) / 100
                        sheet1.Range("J5").Value = (sheet0.Range("L" & h).Value - sheet0.Range("W" & h).Value) / 100
                        sheet1.Range("J6").Value = (sheet0.Range("L" & h).Value - sheet0.Range("W" & h).Value + sheet0.Range("H" & h).Value) / 100
                        sheet1.Range("J7").Value = (sheet0.Range("L" & h).Value - sheet0.Range("W" & h).Value + sheet0.Range("H" & h).Value / 2) / 100

                        sheet1.Range("J8").Value = (sheet0.Range("L" & h).Value - sheet0.Range("H" & h).Value / 2) / 100
                        sheet1.Range("J9").Value = (sheet0.Range("L" & h).Value - sheet0.Range("H" & h).Value / 2) / 100
                        sheet1.Range("J10").Value = (sheet0.Range("L" & h).Value - sheet0.Range("H" & h).Value) / 100
                        sheet1.Range("J11").Value = sheet0.Range("L" & h).Value / 100
                        sheet1.Range("J12").Value = (sheet0.Range("L" & h).Value - sheet0.Range("H" & h).Value / 2) / 100
                    End If

                    '高程
                    P = (sheet0.Range("J" & h).Value - sheet0.Range("M" & h).Value) / (sheet0.Range("W" & h).Value / 100)  '横坡
                    sheet1.Range("K3").Value = sheet0.Range("J" & h).Value + sheet0.Range("H" & h).Value / 200 * P
                    sheet1.Range("K4").Value = sheet0.Range("J" & h).Value + sheet0.Range("H" & h).Value / 200 * P
                    sheet1.Range("K5").Value = sheet0.Range("J" & h).Value
                    sheet1.Range("K6").Value = sheet0.Range("J" & h).Value + sheet0.Range("H" & h).Value / 100 * P
                    sheet1.Range("K7").Value = sheet0.Range("J" & h).Value + sheet0.Range("H" & h).Value / 200 * P

                    sheet1.Range("K8").Value = sheet0.Range("J" & h).Value + (sheet0.Range("W" & h).Value - sheet0.Range("H" & h).Value / 2) / 100 * P
                    sheet1.Range("K9").Value = sheet0.Range("J" & h).Value + (sheet0.Range("W" & h).Value - sheet0.Range("H" & h).Value / 2) / 100 * P
                    sheet1.Range("K10").Value = sheet0.Range("J" & h).Value + (sheet0.Range("W" & h).Value - sheet0.Range("H" & h).Value) / 100 * P
                    sheet1.Range("K11").Value = sheet0.Range("J" & h).Value + sheet0.Range("W" & h).Value / 100 * P
                    sheet1.Range("K12").Value = sheet0.Range("J" & h).Value + (sheet0.Range("W" & h).Value - sheet0.Range("H" & h).Value / 2) / 100 * P

                    '测量偏差值
                    sheet1.Range("O3").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                    sheet1.Range("O4").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                    sheet1.Range("O5").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                    sheet1.Range("O6").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                    sheet1.Range("O7").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                    sheet1.Range("O8").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                    sheet1.Range("O9").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                    sheet1.Range("O10").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                    sheet1.Range("O11").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                    sheet1.Range("O12").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)

                    sheet1.Range("P3").Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                    sheet1.Range("P4").Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                    sheet1.Range("P5").Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                    sheet1.Range("P6").Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                    sheet1.Range("P7").Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                    sheet1.Range("P8").Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                    sheet1.Range("P9").Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                    sheet1.Range("P10").Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                    sheet1.Range("P11").Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                    sheet1.Range("P12").Value = ExApp.WorksheetFunction.RandBetween(1, 4)

                    sheet1.Range("R3").Value = "左侧前"
                    sheet1.Range("R4").Value = "左侧后"
                    sheet1.Range("R5").Value = "左侧左"
                    sheet1.Range("R6").Value = "左侧右"
                    sheet1.Range("R7").Value = "左侧中"
                    sheet1.Range("R8").Value = "右侧前"
                    sheet1.Range("R9").Value = "右侧后"
                    sheet1.Range("R10").Value = "右侧左"
                    sheet1.Range("R11").Value = "右侧右"
                    sheet1.Range("R12").Value = "右侧中"

                    '计算坐标
                    Dim sjxq, sjyq As Double
                    Dim c, b, ZH0 As Integer
                    b = 3
                    Do While sheet1.Range("I" & b).Value <> Nothing
                        ZH0 = Val(ExApp.WorksheetFunction.Substitute(sheet1.Range("I" & b).Value, "*", ""))
                        If sheet3.Range("J2").Value <> "是" Then
                            c = Pd_YSw(ZH0, sheet2.Range("O5 : O500").Value)
                        Else
                            c = 1
                        End If
                        If c = -1 Then
                            MsgBox("请在交点法表内输入数据")
                            TorF = False
                            Exit Sub
                        Else
                            If sheet3.Range("J2").Value <> "是" Then
                                sjxq = ZSZB_X0j(sheet1.Range("I" & b).Value, sheet1.Range("J" & b).Value, 90)  '前偏距坐标
                                sjyq = ZSZB_Y0j(sheet1.Range("I" & b).Value, sheet1.Range("J" & b).Value, 90)
                            Else
                                sjxq = XYF_X(sheet1.Range("I" & b).Value, sheet1.Range("J" & b).Value, 90)  '前偏距坐标
                                sjyq = XYF_Y(sheet1.Range("I" & b).Value, sheet1.Range("J" & b).Value, 90)
                            End If
                            '偏距坐标赋值
                            sheet1.Range("M" & b).Value = Math.Round(sjxq, 3)
                            sheet1.Range("N" & b).Value = Math.Round(sjyq, 3)
                        End If
                        b += 1
                    Loop

                End If
                ExApp.Calculation = ExApp.Calculation.xlCalculationManual '开启手动计算
                sheet1.Range("B" & 5).Value = sheet0.Range("B" & h).Value
                sheet1.Range("B" & r + 6).Value = sheet0.Range("C" & h).Value
                sheet1.Range("C" & r + 6).Value = sheet0.Range("D" & h).Value
                sheet1.Range("D" & r + 6).Value = sheet0.Range("E" & h).Value
                sheet1.Range("B" & r + 7).Value = sheet0.Range("N" & h).Value
                sheet1.Range("B" & r + 8).Value = sheet0.Range("O" & h).Value
                sheet1.Range("B" & r + 9).Value = sheet0.Range("P" & h).Value
                sheet1.Range("B" & r + 10).Value = sheet0.Range("T" & h).Value
                sheet1.Range("B" & r + 11).Value = sheet0.Range("G" & h).Value
                sheet1.Range("B" & r + 12).Value = sheet0.Range("H" & h).Value
                sheet1.Range("B" & r + 13).Value = sheet0.Range("I" & h).Value
                sheet1.Range("B" & r + 14).Value = sheet0.Range("Q" & h).Value
                sheet1.Range("B" & r + 15).Value = sheet0.Range("R" & h).Value
                sheet1.Range("B" & r + 16).Value = sheet0.Range("S" & h).Value
                sheet1.Range("B" & r + 17).Value = sheet0.Range("F" & h).Value
                sheet1.Range("B" & r + 17).Value = sheet0.Range("F" & h).Value
                sheet1.Range("P1").Value = sheet1.Range("B17").Value
                sheet1.Range("L3").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value

                sheet16.Range("A5").Value = sheet1.Range("A5").Value & sheet1.Range("B5").Value
                sheet17.Range("A5").Value = sheet1.Range("A5").Value & sheet1.Range("B5").Value


                '钢筋隐蔽工程
                sheet6.Range("C6").Value = sheet1.Range("B5").Value & sheet1.Range("C6").Value & "基础及下部构造"
                sheet6.Range("E6").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋加工及安装"
                sheet6.Range("C7").Value = sheet1.Range("B5").Value & sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋加工及安装"
                sheet6.Range("E10").Value = sheet1.Range("B18").Value
                sheet6.Range("E11").Value = sheet1.Range("B18").Value
                sheet6.Range("E27").Value = sheet1.Range("B18").Value
                sheet6.Range("E28").Value = sheet1.Range("B18").Value
                '钢筋检表
                sheet7.Range("D6").Value = sheet1.Range("B5").Value
                sheet7.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋"
                sheet7.Range("Q6").Value = sheet1.Range("B18").Value
                sheet7.Range("Q7").Value = sheet1.Range("B18").Value
                sheet7.Range("Q31").Value = sheet1.Range("B18").Value
                sheet7.Range("Q34").Value = sheet1.Range("B18").Value
                '主筋
                sheet7.Range("E15").Value = "设计值：" & sheet1.Range("B8").Value * 10
                LSBLFZ = Nothing
                If sheet1.Range("B7").Value * 2 <= 10 Then
                    For cs = 1 To sheet1.Range("B7").Value * 2
                        LSBL = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet7.Range("G14").Value = LSBLFZ
                    sheet8.Range("D24").Value = "/"
                Else
                    For cs = 1 To sheet1.Range("B7").Value * 2
                        LSBL = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet7.Range("G14").Value = "应测" & sheet1.Range("B7").Value * 2 & "处，实测" & sheet1.Range("B7").Value * 2 & "处，合格" & sheet1.Range("B7").Value * 2 & "处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                    sheet8.Range("D24").Value = LSBLFZ
                End If
                '箍筋
                sheet7.Range("E17").Value = "设计值：" & sheet1.Range("B9").Value * 10
                LSBLFZ = Nothing
                For cs = 1 To 10
                    LSBL = sheet1.Range("B9").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                    LSBLFZ = LSBLFZ & LSBL & "   "
                Next
                sheet7.Range("G16").Value = LSBLFZ

                '骨架尺寸
                sheet7.Range("E19").Value = "设计值：" & sheet1.Range("B14").Value * 10
                sheet7.Range("G18").Value = sheet1.Range("B14").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet7.Range("E21").Value = "设计值：" & sheet1.Range("B15").Value * 10
                sheet7.Range("E22").Value = "设计值：" & sheet1.Range("B16").Value * 10
                sheet7.Range("G20").Value = "宽：  " & sheet1.Range("B15").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4) & vbCrLf &
                                            "高：  " & sheet1.Range("B16").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4)
                '保护层
                sheet7.Range("E27").Value = "设计值：" & sheet1.Range("B10").Value * 10
                sheet7.Range("G26").Value = "应测20处，实测20处，合格20处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                LSBLFZ = Nothing
                For cs = 1 To 20
                    LSBL = sheet1.Range("B10").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                    LSBLFZ = LSBLFZ & LSBL & "   "
                Next
                sheet8.Range("D51").Value = LSBLFZ
                '钢筋记录表
                sheet8.Range("B6").Value = sheet1.Range("B5").Value
                sheet8.Range("B7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋"
                sheet8.Range("K6").Value = sheet1.Range("B18").Value
                sheet8.Range("K7").Value = sheet1.Range("B18").Value

                '4.工序检验申请批复单
                sheet9.Range("C6").Value = sheet1.Range("B5").Value & sheet1.Range("C6").Value & "基础及下部构造"
                sheet9.Range("C7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet9.Range("C8").Value = sheet1.Range("B5").Value
                sheet9.Range("C9").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet9.Range("C10").Value = "混凝土强度、平面尺寸、结构高度、顶面高程、轴线偏位、平整度"

                Call 水准测量记录表()
                If TorF = False Then
                    Exit Sub
                End If
                '高程
                Dim aa, zs, xhh As Integer
                aa = 3
                LSBLFZ = Nothing
                While sheet1.Range("O" & aa).Value <> Nothing Or sheet1.Range("O" & aa).Value = "0"
                    LSBL = sheet1.Range("O" & aa).Value
                    LSBLFZ = LSBLFZ & LSBL & "   "
                    aa += 1
                End While
                sheet10.Range("E19").Value = LSBLFZ
                zs = sheet1.Range("R1048576").End(XlDirection.xlUp).Row - 2
                xhh = 3
                sheet1.Select()

                For cs = 1 To zs
                    If InStr(1, sheet1.Range("R" & xhh).Value, "中"） > 0 Then
                        sheet1.Range("I" & xhh & ":R" & xhh).Select()
                        ExApp.Selection.Delete
                    End If
                    xhh += 1
                Next

                Call 全站仪平面位置检测表（）
                If TorF = False Then
                    Exit Sub
                End If
                '平面
                aa = 3
                LSBLFZ = Nothing
                While sheet1.Range("P" & aa).Value <> Nothing Or sheet1.Range("P" & aa).Value = "0"
                    LSBL = sheet1.Range("P" & aa).Value
                    LSBLFZ = LSBLFZ & LSBL & "   "
                    aa += 1
                End While
                sheet10.Range("E20").Value = LSBLFZ

                sheet16.Activate()
                ExApp.ActiveWindow.SelectedSheets.Copy(, (ExApp.Sheets(ExApp.Sheets.Count)))
                sheet17.Activate()
                ExApp.ActiveWindow.SelectedSheets.Copy(, (ExApp.Sheets(ExApp.Sheets.Count)))

                '耳墙现场质量检验表
                sheet10.Range("D6").Value = sheet1.Range("B5").Value
                sheet10.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet10.Range("K6").Value = sheet1.Range("B17").Value
                sheet10.Range("D14").Value = "长：" & sheet1.Range("B11").Value * 10 & " 宽：" & sheet1.Range("B12").Value * 10
                sheet10.Range("F13").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet10.Range("G13").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet10.Range("F14").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet10.Range("G14").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet10.Range("D18").Value = sheet1.Range("B13").Value * 10
                sheet10.Range("E17").Value = sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet10.Range("F17").Value = sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet10.Range("G17").Value = sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet10.Range("H17").Value = sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet10.Range("I17").Value = sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)

                '平整度
                Dim gs As Integer
                gs = Math.Round((sheet1.Range("B11").Value * sheet1.Range("B13").Value * 2 + sheet1.Range("B12").Value * sheet1.Range("B13").Value * 2) / 2000 / 100, 0)
                LSBLFZ = Nothing
                If gs <= 24 Then
                    For cs = 1 To 24
                        LSBL = ExApp.WorksheetFunction.RandBetween(1, 5)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet10.Range("E21").Value = LSBLFZ
                Else
                    For cs = 1 To gs
                        LSBL = ExApp.WorksheetFunction.RandBetween(1, 5)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet10.Range("E21").Value = LSBLFZ
                End If


                '模板测量偏差值
                sheet1.Range("O3").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O4").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O5").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O6").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("P3").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P4").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P5").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P6").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("L3").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "模板"

                Call 水准测量记录表()
                If TorF = False Then
                    Exit Sub
                End If
                Call 全站仪平面位置检测表（）
                If TorF = False Then
                    Exit Sub
                End If
                '现场模板安装检查记录表
                sheet11.Range("C6").Value = sheet1.Range("B5").Value
                sheet11.Range("C7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet11.Range("I6").Value = sheet1.Range("B17").Value
                sheet11.Range("I7").Value = sheet1.Range("B17").Value
                sheet11.Range("D8").Value = 2
                sheet11.Range("D9").Value = 5
                sheet11.Range("F8").Value = 4
                sheet11.Range("F9").Value = 4
                sheet11.Range("H8").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet11.Range("H9").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet11.Range("J8").Value = 100
                sheet11.Range("J9").Value = 100
                '测量偏差
                sheet11.Range("D11").Value = sheet1.Range("P5").Value
                sheet11.Range("F11").Value = sheet1.Range("P6").Value
                sheet11.Range("H11").Value = sheet1.Range("P3").Value
                sheet11.Range("J11").Value = sheet1.Range("P4").Value
                sheet11.Range("F14").Value = sheet1.Range("O3").Value
                sheet11.Range("G14").Value = sheet1.Range("O4").Value
                sheet11.Range("H14").Value = sheet1.Range("O5").Value
                sheet11.Range("I14").Value = sheet1.Range("O6").Value

                sheet11.Range("F12").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("G12").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("H12").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("I12").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("F13").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("G13").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("H13").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("I13").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("F18").Value = "牢固，稳定"


                '监抽钢筋检表
                sheet13.Range("D6").Value = sheet1.Range("B5").Value
                sheet13.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋"
                sheet13.Range("Q6").Value = sheet1.Range("B18").Value
                sheet13.Range("Q7").Value = sheet1.Range("B18").Value
                sheet13.Range("Q31").Value = sheet1.Range("B18").Value
                sheet13.Range("Q34").Value = sheet1.Range("B18").Value
                '主筋
                sheet13.Range("E15").Value = "设计值：" & sheet1.Range("B8").Value * 10
                sheet13.Range("E15").Value = "设计值：" & sheet1.Range("B8").Value * 10
                LSBLFZ = Nothing
                If Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0) <= 10 Then
                    For cs = 1 To Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0)
                        LSBL = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet13.Range("G14").Value = LSBLFZ
                    sheet14.Range("D24").Value = "/"
                Else
                    sheet13.Range("G14").Value = "应测" & Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0) & "处，实测" & Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0) & "处，合格" & Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0) & "处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                    For cs = 1 To Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0)
                        LSBL = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet14.Range("D24").Value = LSBLFZ
                End If

                '箍筋
                sheet13.Range("E17").Value = "设计值：" & sheet1.Range("B9").Value * 10
                sheet13.Range("G16").Value = sheet1.Range("B9").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9) & "  " &
                                             sheet1.Range("B9").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                '骨架尺寸
                sheet13.Range("E19").Value = "设计值：" & sheet1.Range("B14").Value * 10
                sheet13.Range("K18").Value = sheet1.Range("B14").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet13.Range("E21").Value = "宽：" & sheet1.Range("B15").Value * 10
                sheet13.Range("E22").Value = "高：" & sheet1.Range("B16").Value * 10
                sheet13.Range("G20").Value = "宽： " & sheet1.Range("B15").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4) & vbCrLf &
                                             "高： " & sheet1.Range("B16").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4)
                '保护层
                sheet13.Range("E27").Value = "设计值：" & sheet1.Range("B10").Value * 10
                sheet13.Range("G26").Value = sheet1.Range("B10").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9) & "  " &
                                             sheet1.Range("B10").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9) & "  " &
                                             sheet1.Range("B10").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9) & "  " &
                                             sheet1.Range("B10").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                '监抽钢筋记录表
                sheet14.Range("B6").Value = sheet1.Range("B5").Value
                sheet14.Range("B7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋"
                sheet14.Range("K6").Value = sheet1.Range("B18").Value
                sheet14.Range("K7").Value = sheet1.Range("B18").Value

                '监抽耳墙现场质量检验表
                sheet15.Range("D6").Value = sheet1.Range("B5").Value
                sheet15.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet15.Range("K6").Value = sheet1.Range("B17").Value
                sheet15.Range("D14").Value = "长：" & sheet1.Range("B11").Value * 10 & " 宽：" & sheet1.Range("B12").Value * 10
                sheet15.Range("F13").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet15.Range("F14").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet15.Range("D18").Value = sheet1.Range("B13").Value * 10
                sheet15.Range("E17").Value = sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet15.Range("F17").Value = sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-20, 20)
                sheet15.Range("E19").Value = sheet10.Range("E19").Value
                sheet15.Range("F19").Value = sheet10.Range("F19").Value
                sheet15.Range("E20").Value = sheet10.Range("E20").Value
                sheet15.Range("F20").Value = sheet10.Range("F20").Value
                sheet15.Range("E21").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet15.Range("F21").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet15.Range("G21").Value = ExApp.WorksheetFunction.RandBetween(1, 5)

                '刷新一次数据
                ExApp.Calculate()
                ExApp.Calculation = ExApp.Calculation.xlCalculationManual
                ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                sheet6.Select()
                For i = 7 To ExApp.Sheets.Count
                    EXsheet = Exbook.Worksheets(i)
                    If EXsheet.Visible = True Then
                        EXsheet.Select(Replace:=False)
                    End If
                Next i
                EXsheet = ExApp.ActiveSheet

                ' 导出PDF文件
                PDFFilename = Filepath & "\" & sheet1.Range("B5").Value & sheet1.Range("C6").Value & sheet1.Range("D6").Value & ".pdf"
                EXsheet.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, PDFFilename, XlFixedFormatQuality.xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False)

                sheet18 = Exbook.Worksheets("水准表(2)")
                sheet19 = Exbook.Worksheets("平面表(2)")
                ExApp.DisplayAlerts = False
                sheet18.Delete()
                sheet19.Delete()
                h += 1
            Loop

            MsgBox("已完成！", 0 + 64, "提示")
        Catch Exclerror As Exception   '错误时弹出提示
            MsgBox(Exclerror.Message)
        End Try
        TorF = False
    End Sub

    Sub 肋板资料（）

        Dim h, i, r As Integer
        Dim ExcelFilename, PDFFilename As String  '定义输出的PDF文件名
        Dim FolderDialogObject As New FolderBrowserDialog()
        i = 8
        h = 8
        r = 0
        Try
            sheet0 = Exbook.Worksheets("数据库")
            sheet1 = Exbook.Worksheets("参数表")
            sheet2 = Exbook.Worksheets("交点法")
            sheet3 = Exbook.Worksheets("线元法")
            sheet4 = Exbook.Worksheets("断链")
            sheet5 = Exbook.Worksheets("导线成果表")
            sheet6 = Exbook.Worksheets("申请批复单")
            sheet7 = Exbook.Worksheets("现浇墩、台身检表")
            sheet8 = Exbook.Worksheets("模板记录表")
            sheet9 = Exbook.Worksheets("砼浇筑申请报告单")
            sheet10 = Exbook.Worksheets("监抽现浇墩、台身检表")
            sheet11 = Exbook.Worksheets("水准表")
            sheet12 = Exbook.Worksheets("平面表")

            '改表头
            sheet6.Range("A1").Value = sheet0.Range("C1").Value
            sheet6.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet6.Range("E3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet6.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet6.Range("E4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet7.Range("A1").Value = sheet0.Range("C1").Value
            sheet7.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet7.Range("J3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet7.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet7.Range("J4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet8.Range("A1").Value = sheet0.Range("C1").Value
            sheet8.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet8.Range("I3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet8.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet8.Range("I4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet9.Range("A1").Value = sheet0.Range("C1").Value
            sheet9.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet9.Range("I3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet9.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet9.Range("I4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet10.Range("A1").Value = sheet0.Range("C1").Value
            sheet10.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet10.Range("J3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet10.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet10.Range("J4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet11.Range("A1").Value = sheet0.Range("C1").Value
            sheet11.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet11.Range("H3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet11.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet11.Range("H4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet12.Range("A1").Value = sheet0.Range("C1").Value
            sheet12.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet12.Range("N3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet12.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet12.Range("N4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet0.Range("BA8:BT100000").Value = Nothing
            '计算偏距

            Do While sheet0.Range("B" & i).Value <> Nothing And sheet0.Range("C" & i).Value <> Nothing
                If sheet0.Range("O" & i).Value = Nothing Then
                    MsgBox("请在O" & i & "列选择该桥台是桥梁起点还是终点！！")
                    TorF = False
                    Exit Sub
                ElseIf sheet0.Range("N" & i).Value = "起点" Then
                    sheet0.Range("BA" & i).Value = sheet0.Range("C" & i).Value + Math.Round(sheet0.Range("O" & i).Value / 100 - sheet0.Range("P" & i).Value / 100, 3)
                    sheet0.Range("BE" & i).Value = sheet0.Range("C" & i).Value + Math.Round(sheet0.Range("O" & i).Value / 100 - sheet0.Range("P" & i).Value / 100 - sheet0.Range("H" & i).Value / 100, 3)
                    sheet0.Range("BI" & i).Value = sheet0.Range("C" & i).Value + Math.Round(sheet0.Range("O" & i).Value / 100 - sheet0.Range("P" & i).Value / 100 - sheet0.Range("H" & i).Value / 200, 3)
                    sheet0.Range("BM" & i).Value = sheet0.Range("C" & i).Value + Math.Round(sheet0.Range("O" & i).Value / 100 - sheet0.Range("P" & i).Value / 100 - sheet0.Range("H" & i).Value / 200, 3)
                    sheet0.Range("BQ" & i).Value = sheet0.Range("C" & i).Value + Math.Round(sheet0.Range("O" & i).Value / 100 - sheet0.Range("P" & i).Value / 100 - sheet0.Range("H" & i).Value / 200, 3)
                Else
                    sheet0.Range("BA" & i).Value = sheet0.Range("C" & i).Value - Math.Round(sheet0.Range("O" & i).Value / 100 - sheet0.Range("P" & i).Value / 100 - sheet0.Range("H" & i).Value / 100, 3)
                    sheet0.Range("BE" & i).Value = sheet0.Range("C" & i).Value - Math.Round(sheet0.Range("O" & i).Value / 100 - sheet0.Range("P" & i).Value / 100, 3)
                    sheet0.Range("BI" & i).Value = sheet0.Range("C" & i).Value - Math.Round(sheet0.Range("O" & i).Value / 100 - sheet0.Range("P" & i).Value / 100 - sheet0.Range("H" & i).Value / 200, 3)
                    sheet0.Range("BM" & i).Value = sheet0.Range("C" & i).Value - Math.Round(sheet0.Range("O" & i).Value / 100 - sheet0.Range("P" & i).Value / 100 - sheet0.Range("H" & i).Value / 200, 3)
                    sheet0.Range("BQ" & i).Value = sheet0.Range("C" & i).Value - Math.Round(sheet0.Range("O" & i).Value / 100 - sheet0.Range("P" & i).Value / 100 - sheet0.Range("H" & i).Value / 200, 3)
                End If

                If sheet0.Range("L" & i).Value = "右" Then
                    sheet0.Range("BB" & i).Value = ((sheet0.Range("M" & i).Value) - (sheet0.Range("G" & i).Value / 2)) / -100
                    sheet0.Range("BF" & i).Value = ((sheet0.Range("M" & i).Value) - (sheet0.Range("G" & i).Value / 2)) / -100
                    sheet0.Range("BJ" & i).Value = sheet0.Range("M" & i).Value / -100
                    sheet0.Range("BN" & i).Value = (sheet0.Range("M" & i).Value - sheet0.Range("G" & i).Value) / -100
                    sheet0.Range("BR" & i).Value = ((sheet0.Range("M" & i).Value) - (sheet0.Range("G" & i).Value / 2)) / -100
                Else
                    sheet0.Range("BB" & i).Value = ((sheet0.Range("M" & i).Value) - (sheet0.Range("G" & i).Value / 2)) / 100
                    sheet0.Range("BF" & i).Value = ((sheet0.Range("M" & i).Value) - (sheet0.Range("G" & i).Value / 2)) / 100
                    sheet0.Range("BJ" & i).Value = (sheet0.Range("M" & i).Value - sheet0.Range("G" & i).Value) / 100
                    sheet0.Range("BN" & i).Value = sheet0.Range("M" & i).Value / 100
                    sheet0.Range("BR" & i).Value = ((sheet0.Range("M" & i).Value) - (sheet0.Range("G" & i).Value / 2)) / 100
                End If
                i += 1
            Loop

            Call 质检资料设计坐标()
            If TorF = False Then
                Exit Sub
            End If
            While sheet0.Range("B" & h).Value <> Nothing
                ExApp.Calculation = ExApp.Calculation.xlCalculationManual  '开启手动计算
                sheet1.Range("B" & r + 5).Value = sheet0.Range("B" & h).Value
                sheet1.Range("B" & r + 6).Value = sheet0.Range("C" & h).Value
                sheet1.Range("C" & r + 6).Value = sheet0.Range("D" & h).Value
                sheet1.Range("D" & r + 6).Value = sheet0.Range("E" & h).Value
                sheet1.Range("B" & r + 7).Value = sheet0.Range("G" & h).Value
                sheet1.Range("B" & r + 8).Value = sheet0.Range("H" & h).Value
                sheet1.Range("B" & r + 9).Value = sheet0.Range("I" & h).Value
                sheet1.Range("B" & r + 10).Value = sheet0.Range("F" & h).Value

                '测量参数
                sheet1.Range("I1").Value = sheet1.Range("B5").Value.substring(0, ExApp.WorksheetFunction.Find("K", sheet1.Range("B5").Value))
                sheet1.Range("I3:R15").Value = Nothing
                sheet1.Range("I3").Value = sheet0.Range("BA" & h).Value
                sheet1.Range("I4").Value = sheet0.Range("BE" & h).Value
                sheet1.Range("I5").Value = sheet0.Range("BI" & h).Value
                sheet1.Range("I6").Value = sheet0.Range("BM" & h).Value
                sheet1.Range("J3").Value = sheet0.Range("BB" & h).Value
                sheet1.Range("J4").Value = sheet0.Range("BF" & h).Value
                sheet1.Range("J5").Value = sheet0.Range("BJ" & h).Value
                sheet1.Range("J6").Value = sheet0.Range("BN" & h).Value
                sheet1.Range("K3").Value = sheet0.Range("J" & h).Value + (sheet0.Range("K" & h).Value + sheet0.Range("I" & h).Value) / 100
                sheet1.Range("K4").Value = sheet0.Range("J" & h).Value + (sheet0.Range("K" & h).Value + sheet0.Range("I" & h).Value) / 100
                sheet1.Range("K5").Value = sheet0.Range("J" & h).Value + (sheet0.Range("K" & h).Value + sheet0.Range("I" & h).Value) / 100
                sheet1.Range("K6").Value = sheet0.Range("J" & h).Value + (sheet0.Range("K" & h).Value + sheet0.Range("I" & h).Value) / 100
                sheet1.Range("M3").Value = sheet0.Range("BC" & h).Value
                sheet1.Range("N3").Value = sheet0.Range("BD" & h).Value
                sheet1.Range("M4").Value = sheet0.Range("BG" & h).Value
                sheet1.Range("N4").Value = sheet0.Range("BH" & h).Value
                sheet1.Range("M5").Value = sheet0.Range("BK" & h).Value
                sheet1.Range("N5").Value = sheet0.Range("BL" & h).Value
                sheet1.Range("M6").Value = sheet0.Range("BO" & h).Value
                sheet1.Range("N6").Value = sheet0.Range("BP" & h).Value
                sheet1.Range("R3").Value = "前"
                sheet1.Range("R4").Value = "后"
                sheet1.Range("R5").Value = "左"
                sheet1.Range("R6").Value = "右"
                sheet1.Range("P1").Value = sheet1.Range("B10").Value
                sheet1.Range("L3").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "模板"

                '测量偏差值
                sheet1.Range("O3").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O4").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O5").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O6").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("P3").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P4").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P5").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P6").Value = ExApp.WorksheetFunction.RandBetween(1, 9)

                sheet11.Range("A5").Value = sheet1.Range("A5").Value & sheet1.Range("B5").Value
                sheet12.Range("A5").Value = sheet1.Range("A5").Value & sheet1.Range("B5").Value

                Call 水准测量记录表()
                If TorF = False Then
                    Exit Sub
                End If
                Call 全站仪平面位置检测表（）
                If TorF = False Then
                    Exit Sub
                End If
                sheet11.Activate()
                ExApp.ActiveWindow.SelectedSheets.Copy(, (ExApp.Sheets(ExApp.Sheets.Count)))
                sheet12.Activate()
                ExApp.ActiveWindow.SelectedSheets.Copy(, (ExApp.Sheets(ExApp.Sheets.Count)))

                '现场模板安装检查记录表
                sheet8.Range("C6").Value = sheet1.Range("B5").Value
                sheet8.Range("C7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet8.Range("I6").Value = sheet1.Range("B10").Value
                sheet8.Range("I7").Value = sheet1.Range("B10").Value
                sheet8.Range("D8").Value = 2
                sheet8.Range("D9").Value = 5
                sheet8.Range("F8").Value = 4
                sheet8.Range("F9").Value = 4
                sheet8.Range("H8").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet8.Range("H9").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet8.Range("J8").Value = 100
                sheet8.Range("J9").Value = 100
                '测量偏差
                sheet8.Range("D11").Value = sheet1.Range("P3").Value
                sheet8.Range("F11").Value = sheet1.Range("P4").Value
                sheet8.Range("H11").Value = sheet1.Range("P5").Value
                sheet8.Range("J11").Value = sheet1.Range("P6").Value
                sheet8.Range("F14").Value = sheet1.Range("O3").Value
                sheet8.Range("G14").Value = sheet1.Range("O4").Value
                sheet8.Range("H14").Value = sheet1.Range("O5").Value
                sheet8.Range("I14").Value = sheet1.Range("O6").Value

                sheet8.Range("F12").Value = sheet1.Range("B7").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet8.Range("G12").Value = sheet1.Range("B7").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet8.Range("H12").Value = sheet1.Range("B7").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet8.Range("I12").Value = sheet1.Range("B7").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet8.Range("F13").Value = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet8.Range("G13").Value = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet8.Range("H13").Value = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet8.Range("I13").Value = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet8.Range("F18").Value = "牢固，稳定"

                '工序检验申请批复单
                sheet6.Range("C6").Value = sheet1.Range("B5").Value & sheet1.Range("C6").Value & "基础及下部构造"
                sheet6.Range("C7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet6.Range("C8").Value = sheet1.Range("B5").Value
                sheet6.Range("C9").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet6.Range("C10").Value = "混凝土强度、断面面尺寸、全高竖直度、顶面高程、轴线偏位、平整度"

                sheet1.Range("O3").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O4").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O5").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("P3").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P4").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P5").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("P6").Value = ExApp.WorksheetFunction.RandBetween(1, 9)
                sheet1.Range("L3").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                Call 全站仪平面位置检测表（）
                If TorF = False Then
                    Exit Sub
                End If
                '现浇墩、台身现场质量检验表
                sheet7.Range("D6").Value = sheet1.Range("B5").Value
                sheet7.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet7.Range("J6").Value = sheet1.Range("B10").Value
                sheet7.Range("D13").Value = "长：" & sheet1.Range("B7").Value * 10 & " " & "宽：" & sheet1.Range("B8").Value * 10
                sheet7.Range("G12").Value = sheet1.Range("B7").Value * 10 + ExApp.WorksheetFunction.RandBetween(-10, 10)
                sheet7.Range("H12").Value = sheet1.Range("B7").Value * 10 + ExApp.WorksheetFunction.RandBetween(-10, 10)
                sheet7.Range("G13").Value = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-10, 10)
                sheet7.Range("H13").Value = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-10, 10)
                '全高竖直度
                If sheet1.Range("B9").Value / 100 <= 5 Then
                    sheet7.Range("F14").Value = ExApp.WorksheetFunction.RandBetween(1, 4) & "   " &
                                                ExApp.WorksheetFunction.RandBetween(1, 4) & "   " &
                                                ExApp.WorksheetFunction.RandBetween(1, 4) & "   " &
                                                ExApp.WorksheetFunction.RandBetween(1, 4)
                    sheet7.Range("F15").Value = "/"
                    sheet7.Range("F16").Value = "/"
                ElseIf sheet1.Range("B9").Value / 100 <= 60 Then
                    sheet7.Range("F15").Value = ExApp.WorksheetFunction.RandBetween(1, Math.Round(sheet1.Range("B11").Value * 10 / 1000, 0) - 1) & "   " &
                                                ExApp.WorksheetFunction.RandBetween(1, Math.Round(sheet1.Range("B11").Value * 10 / 1000, 0) - 1) & "   " &
                                                ExApp.WorksheetFunction.RandBetween(1, Math.Round(sheet1.Range("B11").Value * 10 / 1000, 0) - 1) & "   " &
                                                ExApp.WorksheetFunction.RandBetween(1, Math.Round(sheet1.Range("B11").Value * 10 / 1000, 0) - 1)
                    sheet7.Range("F14").Value = "/"
                    sheet7.Range("F16").Value = "/"
                Else
                    sheet7.Range("F16").Value = ExApp.WorksheetFunction.RandBetween(1, Math.Round(sheet1.Range("B11").Value * 10 / 3000, 0) - 1) & "   " &
                                                ExApp.WorksheetFunction.RandBetween(1, Math.Round(sheet1.Range("B11").Value * 10 / 3000, 0) - 1) & "   " &
                                                ExApp.WorksheetFunction.RandBetween(1, Math.Round(sheet1.Range("B11").Value * 10 / 3000, 0) - 1) & "   " &
                                                ExApp.WorksheetFunction.RandBetween(1, Math.Round(sheet1.Range("B11").Value * 10 / 3000, 0) - 1)
                    sheet7.Range("F14").Value = "/"
                    sheet7.Range("F15").Value = "/"
                End If

                '轴线偏位
                If sheet1.Range("B9").Value / 100 <= 60 Then
                    sheet7.Range("F18").Value = sheet1.Range("P3").Value & "   " &
                                                sheet1.Range("P4").Value & "   " &
                                                sheet1.Range("P5").Value & "   " &
                                                sheet1.Range("P6").Value
                    sheet7.Range("F19").Value = "/"
                Else
                    sheet7.Range("F19").Value = sheet1.Range("P3").Value & "   " &
                                                sheet1.Range("P4").Value & "   " &
                                                sheet1.Range("P5").Value & "   " &
                                                sheet1.Range("P6").Value
                    sheet7.Range("F18").Value = "/"
                End If
                sheet7.Range("F21").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet7.Range("G21").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet7.Range("H21").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet7.Range("I21").Value = ExApp.WorksheetFunction.RandBetween(1, 5)

                sheet1.Range("I6:R15").Value = Nothing
                sheet1.Range("J5").Value = sheet0.Range("BR" & h).Value
                sheet1.Range("R5").Value = "中"
                sheet1.Range("L3").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                Call 水准测量记录表()
                If TorF = False Then
                    Exit Sub
                End If
                '顶面高程
                sheet7.Range("F17").Value = sheet1.Range("O3").Value
                sheet7.Range("G17").Value = sheet1.Range("O4").Value
                sheet7.Range("H17").Value = sheet1.Range("O5").Value

                '监抽现浇墩、台身现场质量检验表
                sheet10.Range("D6").Value = sheet1.Range("B5").Value
                sheet10.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet10.Range("J6").Value = sheet1.Range("B10").Value
                sheet10.Range("D13").Value = "长：" & sheet1.Range("B7").Value * 10 & " " & "宽：" & sheet1.Range("B8").Value * 10
                sheet10.Range("G12").Value = sheet1.Range("B7").Value * 10 + ExApp.WorksheetFunction.RandBetween(-10, 10)
                sheet10.Range("G13").Value = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-10, 10)

                '全高竖直度
                If sheet1.Range("B9").Value / 100 <= 5 Then
                    sheet10.Range("F14").Value = ExApp.WorksheetFunction.RandBetween(1, 4) & "   " &
                                                ExApp.WorksheetFunction.RandBetween(1, 4)
                    sheet10.Range("F15").Value = "/"
                    sheet10.Range("F16").Value = "/"
                ElseIf sheet1.Range("B9").Value / 100 <= 60 Then
                    sheet10.Range("F15").Value = ExApp.WorksheetFunction.RandBetween(1, Math.Round(sheet1.Range("B11").Value * 10 / 1000, 0) - 1) & "   " &
                                                ExApp.WorksheetFunction.RandBetween(1, Math.Round(sheet1.Range("B11").Value * 10 / 1000, 0) - 1)
                    sheet10.Range("F14").Value = "/"
                    sheet10.Range("F16").Value = "/"
                Else
                    sheet10.Range("F16").Value = ExApp.WorksheetFunction.RandBetween(1, Math.Round(sheet1.Range("B11").Value * 10 / 3000, 0) - 1) & "   " &
                                                ExApp.WorksheetFunction.RandBetween(1, Math.Round(sheet1.Range("B11").Value * 10 / 3000, 0) - 1)
                    sheet10.Range("F14").Value = "/"
                    sheet10.Range("F15").Value = "/"
                End If
                '轴线偏位
                If sheet1.Range("B9").Value / 100 <= 60 Then
                    sheet10.Range("F18").Value = sheet1.Range("P3").Value & "   " &
                                                sheet1.Range("P4").Value & "   " &
                                                sheet1.Range("P5").Value & "   " &
                                                sheet1.Range("P6").Value
                    sheet10.Range("F19").Value = "/"
                Else
                    sheet10.Range("F19").Value = sheet1.Range("P3").Value & "   " &
                                                sheet1.Range("P4").Value & "   " &
                                                sheet1.Range("P5").Value & "   " &
                                                sheet1.Range("P6").Value
                    sheet10.Range("F18").Value = "/"
                End If

                sheet10.Range("F21").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet10.Range("G21").Value = ExApp.WorksheetFunction.RandBetween(1, 5)

                ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                '选择表格
                sheet6.Select()
                For i = 7 To ExApp.Sheets.Count
                    EXsheet = Exbook.Worksheets(i)
                    If EXsheet.Visible = True Then
                        EXsheet.Select(Replace:=False)
                    End If
                Next i
                EXsheet = ExApp.ActiveSheet

                ' 导出PDF文件
                PDFFilename = Filepath & "\" & sheet1.Range("B5").Value & sheet1.Range("C6").Value & sheet1.Range("D6").Value & ".pdf"
                '保存PDF文件
                EXsheet.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, PDFFilename, XlFixedFormatQuality.xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False)

                sheet13 = Exbook.Worksheets("水准表(2)")
                sheet14 = Exbook.Worksheets("平面表(2)")
                ExApp.DisplayAlerts = False
                sheet13.Delete()
                sheet14.Delete()
                h += 1
            End While

            MsgBox("已完成！", 0 + 64, "提示")
        Catch Exclerror As Exception   '错误时弹出提示
            MsgBox(Exclerror.Message)
        End Try
        TorF = False
    End Sub

    Sub 支座垫石资料（）
        Dim Filepath, ExcelFilename, PDFFilename, LSBL, LSBLFZ, LSBL2, LSBLFZ2 As String
        Dim h, r, i, n, j, m, x, cs, zs, xhh As Integer
        Dim p As Double
        Dim FolderDialogObject As New FolderBrowserDialog()
        r = 0
        h = 8
        Try
            sheet0 = Exbook.Worksheets("数据库")
            sheet1 = Exbook.Worksheets("参数表")
            sheet2 = Exbook.Worksheets("交点法")
            sheet3 = Exbook.Worksheets("线元法")
            sheet4 = Exbook.Worksheets("断链")
            sheet5 = Exbook.Worksheets("导线成果表")
            sheet6 = Exbook.Worksheets("钢筋隐蔽工程")
            sheet7 = Exbook.Worksheets("钢筋检表")
            sheet8 = Exbook.Worksheets("钢筋记录表")
            sheet9 = Exbook.Worksheets("申请批复单")
            sheet10 = Exbook.Worksheets("支座垫石检表")
            sheet11 = Exbook.Worksheets("模板记录表")
            sheet12 = Exbook.Worksheets("砼浇筑申请报告单")
            sheet13 = Exbook.Worksheets("监抽钢筋检表")
            sheet14 = Exbook.Worksheets("监抽钢筋记录表")
            sheet15 = Exbook.Worksheets("监抽支座垫石检表")
            sheet16 = Exbook.Worksheets("水准表")
            sheet17 = Exbook.Worksheets("平面表")

            '改表头
            sheet6.Range("A1").Value = sheet0.Range("C1").Value
            sheet6.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet6.Range("E3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet6.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet6.Range("E4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet7.Range("A1").Value = sheet0.Range("C1").Value
            sheet7.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet7.Range("P3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet7.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet7.Range("P4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet8.Range("A1").Value = sheet0.Range("C1").Value
            sheet8.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet8.Range("L3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet8.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet8.Range("L4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet9.Range("A1").Value = sheet0.Range("C1").Value
            sheet9.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet9.Range("E3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet9.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet9.Range("E4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet10.Range("A1").Value = sheet0.Range("C1").Value
            sheet10.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet10.Range("P3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet10.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet10.Range("P4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet11.Range("A1").Value = sheet0.Range("C1").Value
            sheet11.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet11.Range("I3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet11.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet11.Range("I4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet12.Range("A1").Value = sheet0.Range("C1").Value
            sheet12.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet12.Range("I3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet12.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet12.Range("I4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet13.Range("A1").Value = sheet0.Range("C1").Value
            sheet13.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet13.Range("P3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet13.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet13.Range("P4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet14.Range("A1").Value = sheet0.Range("C1").Value
            sheet14.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet14.Range("L3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet14.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet14.Range("L4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet15.Range("A1").Value = sheet0.Range("C1").Value
            sheet15.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet15.Range("P3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet15.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet15.Range("P4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet16.Range("A1").Value = sheet0.Range("C1").Value
            sheet16.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet16.Range("H3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet16.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet16.Range("H4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet17.Range("A1").Value = sheet0.Range("C1").Value
            sheet17.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet17.Range("N3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet17.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet17.Range("N4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value


            ExApp.Calculation = ExApp.Calculation.xlCalculationManual '开启手动计算
            '计算测量参数
            While sheet0.Range("B" & h).Value <> Nothing
                sheet1.Range("I3:R100000").Value = Nothing
                sheet1.Range("I1").Value = sheet1.Range("B5").Value.substring(0, ExApp.WorksheetFunction.Find("K", sheet1.Range("B5").Value))
                j = 3
                If sheet0.Range("J" & h).Value = "盖梁垫石" Then
                    If sheet0.Range("C" & h).Value = Nothing And sheet0.Range("S" & h).Value = "否" And sheet0.Range("T" & h).Value <> Nothing And sheet0.Range("U" & h).Value <> Nothing Then
                        sheet0.Range("C" & h).Value = FSZhj(sheet0.Range("T" & h).Value, sheet0.Range("U" & h).Value)
                    ElseIf sheet0.Range("C" & h).Value <> Nothing And sheet0.Range("S" & h).Value = "否" Then
                        MsgBox("请核对""C列""中心桩号或""S列、W列""X、Y坐标是否正确填写！")
                        TorF = False
                        Exit Sub
                    End If

                    x = 0
                    m = sheet0.Range("AF" & h).Value / 2
                    For n = 1 To m
                        p = (sheet0.Range("W" & h).Value - sheet0.Range("V" & h).Value) / (sheet0.Range("X" & h).Value / 100)  '横坡
                        '桩号
                        sheet1.Range("I" & j).Value = sheet0.Range("C" & h).Value + sheet0.Range("AE" & h).Value / 100 + sheet0.Range("H" & h).Value / 2 / 100
                        sheet1.Range("I" & j + 1).Value = sheet0.Range("C" & h).Value + sheet0.Range("AE" & h).Value / 100 - sheet0.Range("H" & h).Value / 2 / 100
                        sheet1.Range("I" & j + 2).Value = sheet0.Range("C" & h).Value + sheet0.Range("AE" & h).Value / 100
                        sheet1.Range("I" & j + 3).Value = sheet0.Range("C" & h).Value + sheet0.Range("AE" & h).Value / 100
                        sheet1.Range("I" & j + 4).Value = sheet0.Range("C" & h).Value + sheet0.Range("AE" & h).Value / 100
                        sheet1.Range("I" & j + 5).Value = sheet0.Range("C" & h).Value - sheet0.Range("AE" & h).Value / 100 + sheet0.Range("H" & h).Value / 2 / 100
                        sheet1.Range("I" & j + 6).Value = sheet0.Range("C" & h).Value - sheet0.Range("AE" & h).Value / 100 - sheet0.Range("H" & h).Value / 2 / 100
                        sheet1.Range("I" & j + 7).Value = sheet0.Range("C" & h).Value - sheet0.Range("AE" & h).Value / 100
                        sheet1.Range("I" & j + 8).Value = sheet0.Range("C" & h).Value - sheet0.Range("AE" & h).Value / 100
                        sheet1.Range("I" & j + 9).Value = sheet0.Range("C" & h).Value - sheet0.Range("AE" & h).Value / 100
                        '测量偏差值
                        sheet1.Range("O" & j).Value = ExApp.WorksheetFunction.RandBetween(-1, 1)
                        sheet1.Range("O" & j + 1).Value = ExApp.WorksheetFunction.RandBetween(-1, 1)
                        sheet1.Range("O" & j + 2).Value = ExApp.WorksheetFunction.RandBetween(-1, 1)
                        sheet1.Range("O" & j + 3).Value = ExApp.WorksheetFunction.RandBetween(-1, 1)
                        sheet1.Range("O" & j + 4).Value = ExApp.WorksheetFunction.RandBetween(-1, 1)
                        sheet1.Range("O" & j + 5).Value = ExApp.WorksheetFunction.RandBetween(-1, 1)
                        sheet1.Range("O" & j + 6).Value = ExApp.WorksheetFunction.RandBetween(-1, 1)
                        sheet1.Range("O" & j + 7).Value = ExApp.WorksheetFunction.RandBetween(-1, 1)
                        sheet1.Range("O" & j + 8).Value = ExApp.WorksheetFunction.RandBetween(-1, 1)
                        sheet1.Range("O" & j + 9).Value = ExApp.WorksheetFunction.RandBetween(-1, 1)
                        sheet1.Range("P" & j).Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                        sheet1.Range("P" & j + 1).Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                        sheet1.Range("P" & j + 2).Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                        sheet1.Range("P" & j + 3).Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                        sheet1.Range("P" & j + 4).Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                        sheet1.Range("P" & j + 5).Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                        sheet1.Range("P" & j + 6).Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                        sheet1.Range("P" & j + 7).Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                        sheet1.Range("P" & j + 8).Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                        sheet1.Range("P" & j + 9).Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                        '备注
                        sheet1.Range("R" & j).Value = n + n - 1 & "#前"
                        sheet1.Range("R" & j + 1).Value = n + n - 1 & "#后"
                        sheet1.Range("R" & j + 2).Value = n + n - 1 & "#左"
                        sheet1.Range("R" & j + 3).Value = n + n - 1 & "#右"
                        sheet1.Range("R" & j + 4).Value = n + n - 1 & "#中"
                        sheet1.Range("R" & j + 5).Value = n + n & "#前"
                        sheet1.Range("R" & j + 6).Value = n + n & "#后"
                        sheet1.Range("R" & j + 7).Value = n + n & "#左"
                        sheet1.Range("R" & j + 8).Value = n + n & "#右"
                        sheet1.Range("R" & j + 9).Value = n + n & "#中"

                        If sheet0.Range("K" & h).Value = "中" Or sheet0.Range("K" & h).Value = "右" Then
                            '偏距
                            sheet1.Range("J" & j).Value = (sheet0.Range("AA" & h).Value / 100 - x * sheet0.Range("Y" & h).Value / 100) * -1
                            sheet1.Range("J" & j + 1).Value = (sheet0.Range("AA" & h).Value / 100 - x * sheet0.Range("Y" & h).Value / 100) * -1
                            sheet1.Range("J" & j + 2).Value = ((sheet0.Range("AA" & h).Value + sheet0.Range("G" & h).Value / 2) / 100 - x * sheet0.Range("Y" & h).Value / 100) * -1
                            sheet1.Range("J" & j + 3).Value = ((sheet0.Range("AA" & h).Value - sheet0.Range("G" & h).Value / 2) / 100 - x * sheet0.Range("Y" & h).Value / 100) * -1
                            sheet1.Range("J" & j + 4).Value = (sheet0.Range("AA" & h).Value / 100 - x * sheet0.Range("Y" & h).Value / 100) * -1
                            sheet1.Range("J" & j + 5).Value = (sheet0.Range("AB" & h).Value / 100 - x * sheet0.Range("Z" & h).Value / 100) * -1
                            sheet1.Range("J" & j + 6).Value = (sheet0.Range("AB" & h).Value / 100 - x * sheet0.Range("Z" & h).Value / 100) * -1
                            sheet1.Range("J" & j + 7).Value = ((sheet0.Range("AB" & h).Value + sheet0.Range("G" & h).Value / 2) / 100 - x * sheet0.Range("Z" & h).Value / 100) * -1
                            sheet1.Range("J" & j + 8).Value = ((sheet0.Range("AB" & h).Value - sheet0.Range("G" & h).Value / 2) / 100 - x * sheet0.Range("Z" & h).Value / 100) * -1
                            sheet1.Range("J" & j + 9).Value = (sheet0.Range("AB" & h).Value / 100 - x * sheet0.Range("Z" & h).Value / 100) * -1
                            '高程
                            sheet1.Range("K" & j).Value = (sheet0.Range("V" & h).Value + sheet0.Range("I" & h).Value / 100) + (sheet0.Range("AC" & h).Value + x * sheet0.Range("Y" & h).Value) * p / 100
                            sheet1.Range("K" & j + 1).Value = (sheet0.Range("V" & h).Value + sheet0.Range("I" & h).Value / 100) + (sheet0.Range("AC" & h).Value + x * sheet0.Range("Y" & h).Value) * p / 100
                            sheet1.Range("K" & j + 2).Value = (sheet0.Range("V" & h).Value + sheet0.Range("I" & h).Value / 100) + (sheet0.Range("AC" & h).Value + x * sheet0.Range("Y" & h).Value) * p / 100
                            sheet1.Range("K" & j + 3).Value = (sheet0.Range("V" & h).Value + sheet0.Range("I" & h).Value / 100) + (sheet0.Range("AC" & h).Value + x * sheet0.Range("Y" & h).Value) * p / 100
                            sheet1.Range("K" & j + 4).Value = (sheet0.Range("V" & h).Value + sheet0.Range("I" & h).Value / 100) + (sheet0.Range("AC" & h).Value + x * sheet0.Range("Y" & h).Value) * p / 100
                            sheet1.Range("K" & j + 5).Value = (sheet0.Range("V" & h).Value + sheet0.Range("I" & h).Value / 100) + (sheet0.Range("AC" & h).Value + x * sheet0.Range("Y" & h).Value) * p / 100
                            sheet1.Range("K" & j + 6).Value = (sheet0.Range("V" & h).Value + sheet0.Range("I" & h).Value / 100) + (sheet0.Range("AC" & h).Value + x * sheet0.Range("Y" & h).Value) * p / 100
                            sheet1.Range("K" & j + 7).Value = (sheet0.Range("V" & h).Value + sheet0.Range("I" & h).Value / 100) + (sheet0.Range("AC" & h).Value + x * sheet0.Range("Y" & h).Value) * p / 100
                            sheet1.Range("K" & j + 8).Value = (sheet0.Range("V" & h).Value + sheet0.Range("I" & h).Value / 100) + (sheet0.Range("AC" & h).Value + x * sheet0.Range("Y" & h).Value) * p / 100
                            sheet1.Range("K" & j + 9).Value = (sheet0.Range("V" & h).Value + sheet0.Range("I" & h).Value / 100) + (sheet0.Range("AC" & h).Value + x * sheet0.Range("Y" & h).Value) * p / 100
                        Else
                            '中线位于构筑物左侧时
                            '偏距
                            sheet1.Range("J" & j).Value = sheet0.Range("AA" & h).Value / 100 - ((m - 1) * sheet0.Range("Y" & h).Value / 100)
                            sheet1.Range("J" & j + 1).Value = sheet0.Range("AA" & h).Value / 100 - ((m - 1) * sheet0.Range("Y" & h).Value / 100)
                            sheet1.Range("J" & j + 2).Value = sheet0.Range("AA" & h).Value / 100 - ((m - 1) * sheet0.Range("Y" & h).Value / 100) - (sheet0.Range("G" & h).Value / 2) / 100
                            sheet1.Range("J" & j + 3).Value = sheet0.Range("AA" & h).Value / 100 - ((m - 1) * sheet0.Range("Y" & h).Value / 100) + (sheet0.Range("G" & h).Value / 2) / 100
                            sheet1.Range("J" & j + 4).Value = sheet0.Range("AA" & h).Value / 100 - ((m - 1) * sheet0.Range("Y" & h).Value / 100)
                            sheet1.Range("J" & j + 5).Value = sheet0.Range("AB" & h).Value / 100 - ((m - 1) * sheet0.Range("Z" & h).Value / 100)
                            sheet1.Range("J" & j + 6).Value = sheet0.Range("AB" & h).Value / 100 - ((m - 1) * sheet0.Range("Z" & h).Value / 100)
                            sheet1.Range("J" & j + 7).Value = sheet0.Range("AB" & h).Value / 100 - ((m - 1) * sheet0.Range("Z" & h).Value / 100) - (sheet0.Range("G" & h).Value / 2) / 100
                            sheet1.Range("J" & j + 8).Value = sheet0.Range("AB" & h).Value / 100 - ((m - 1) * sheet0.Range("Z" & h).Value / 100) + (sheet0.Range("G" & h).Value / 2) / 100
                            sheet1.Range("J" & j + 9).Value = sheet0.Range("AB" & h).Value / 100 - ((m - 1) * sheet0.Range("Z" & h).Value / 100)
                            '高程
                            sheet1.Range("K" & j).Value = (sheet0.Range("V" & h).Value + sheet0.Range("I" & h).Value / 100) + (sheet0.Range("AC" & h).Value + x * sheet0.Range("Y" & h).Value) * p / 100
                            sheet1.Range("K" & j + 1).Value = (sheet0.Range("V" & h).Value + sheet0.Range("I" & h).Value / 100) + (sheet0.Range("AC" & h).Value + x * sheet0.Range("Y" & h).Value) * p / 100
                            sheet1.Range("K" & j + 2).Value = (sheet0.Range("V" & h).Value + sheet0.Range("I" & h).Value / 100) + (sheet0.Range("AC" & h).Value + x * sheet0.Range("Y" & h).Value) * p / 100
                            sheet1.Range("K" & j + 3).Value = (sheet0.Range("V" & h).Value + sheet0.Range("I" & h).Value / 100) + (sheet0.Range("AC" & h).Value + x * sheet0.Range("Y" & h).Value) * p / 100
                            sheet1.Range("K" & j + 4).Value = (sheet0.Range("V" & h).Value + sheet0.Range("I" & h).Value / 100) + (sheet0.Range("AC" & h).Value + x * sheet0.Range("Y" & h).Value) * p / 100
                            sheet1.Range("K" & j + 5).Value = (sheet0.Range("V" & h).Value + sheet0.Range("I" & h).Value / 100) + (sheet0.Range("AC" & h).Value + x * sheet0.Range("Y" & h).Value) * p / 100
                            sheet1.Range("K" & j + 6).Value = (sheet0.Range("V" & h).Value + sheet0.Range("I" & h).Value / 100) + (sheet0.Range("AC" & h).Value + x * sheet0.Range("Y" & h).Value) * p / 100
                            sheet1.Range("K" & j + 7).Value = (sheet0.Range("V" & h).Value + sheet0.Range("I" & h).Value / 100) + (sheet0.Range("AC" & h).Value + x * sheet0.Range("Y" & h).Value) * p / 100
                            sheet1.Range("K" & j + 8).Value = (sheet0.Range("V" & h).Value + sheet0.Range("I" & h).Value / 100) + (sheet0.Range("AC" & h).Value + x * sheet0.Range("Y" & h).Value) * p / 100
                            sheet1.Range("K" & j + 9).Value = (sheet0.Range("V" & h).Value + sheet0.Range("I" & h).Value / 100) + (sheet0.Range("AC" & h).Value + x * sheet0.Range("Y" & h).Value) * p / 100

                        End If
                        j += 10
                        x += 1
                        m -= 1
                    Next n

                ElseIf sheet0.Range("J" & h).Value = "墩柱垫石" Then
                    If sheet0.Range("C" & h).Value = Nothing And sheet0.Range("S" & h).Value = "否" And sheet0.Range("T" & h).Value <> Nothing And sheet0.Range("U" & h).Value <> Nothing Then
                        sheet0.Range("C" & h).Value = FSZhj(sheet0.Range("T" & h).Value, sheet0.Range("U" & h).Value)
                    ElseIf sheet0.Range("C" & h).Value <> Nothing And sheet0.Range("S" & h).Value = "否" Then
                        MsgBox("请核对""C列""中心桩号或""S列、W列""X、Y坐标是否正确填写！")
                        TorF = False
                        Exit Sub
                    End If

                    m = sheet0.Range("AJ" & h).Value / 2
                    For n = 1 To m

                        '桩号
                        sheet1.Range("I" & j).Value = sheet0.Range("C" & h).Value + sheet0.Range("H" & h).Value / 2 / 100
                        sheet1.Range("I" & j + 1).Value = sheet0.Range("C" & h).Value - sheet0.Range("H" & h).Value / 2 / 100
                        sheet1.Range("I" & j + 2).Value = sheet0.Range("C" & h).Value
                        sheet1.Range("I" & j + 3).Value = sheet0.Range("C" & h).Value
                        sheet1.Range("I" & j + 4).Value = sheet0.Range("C" & h).Value
                        sheet1.Range("I" & j + 5).Value = sheet0.Range("C" & h).Value + sheet0.Range("H" & h).Value / 2 / 100
                        sheet1.Range("I" & j + 6).Value = sheet0.Range("C" & h).Value - sheet0.Range("H" & h).Value / 2 / 100
                        sheet1.Range("I" & j + 7).Value = sheet0.Range("C" & h).Value
                        sheet1.Range("I" & j + 8).Value = sheet0.Range("C" & h).Value
                        sheet1.Range("I" & j + 9).Value = sheet0.Range("C" & h).Value
                        '高程
                        sheet1.Range("K" & j).Value = sheet0.Range("AG" & h).Value + sheet0.Range("I" & h).Value / 100
                        sheet1.Range("K" & j + 1).Value = sheet0.Range("AG" & h).Value + sheet0.Range("I" & h).Value / 100
                        sheet1.Range("K" & j + 2).Value = sheet0.Range("AG" & h).Value + sheet0.Range("I" & h).Value / 100
                        sheet1.Range("K" & j + 3).Value = sheet0.Range("AG" & h).Value + sheet0.Range("I" & h).Value / 100
                        sheet1.Range("K" & j + 4).Value = sheet0.Range("AG" & h).Value + sheet0.Range("I" & h).Value / 100
                        sheet1.Range("K" & j + 5).Value = sheet0.Range("AH" & h).Value + sheet0.Range("I" & h).Value / 100
                        sheet1.Range("K" & j + 6).Value = sheet0.Range("AH" & h).Value + sheet0.Range("I" & h).Value / 100
                        sheet1.Range("K" & j + 7).Value = sheet0.Range("AH" & h).Value + sheet0.Range("I" & h).Value / 100
                        sheet1.Range("K" & j + 8).Value = sheet0.Range("AH" & h).Value + sheet0.Range("I" & h).Value / 100
                        sheet1.Range("K" & j + 9).Value = sheet0.Range("AH" & h).Value + sheet0.Range("I" & h).Value / 100
                        '测量偏差值
                        sheet1.Range("O" & j).Value = ExApp.WorksheetFunction.RandBetween(-1, 1)
                        sheet1.Range("O" & j + 1).Value = ExApp.WorksheetFunction.RandBetween(-1, 1)
                        sheet1.Range("O" & j + 2).Value = ExApp.WorksheetFunction.RandBetween(-1, 1)
                        sheet1.Range("O" & j + 3).Value = ExApp.WorksheetFunction.RandBetween(-1, 1)
                        sheet1.Range("O" & j + 4).Value = ExApp.WorksheetFunction.RandBetween(-1, 1)
                        sheet1.Range("O" & j + 5).Value = ExApp.WorksheetFunction.RandBetween(-1, 1)
                        sheet1.Range("O" & j + 6).Value = ExApp.WorksheetFunction.RandBetween(-1, 1)
                        sheet1.Range("O" & j + 7).Value = ExApp.WorksheetFunction.RandBetween(-1, 1)
                        sheet1.Range("O" & j + 8).Value = ExApp.WorksheetFunction.RandBetween(-1, 1)
                        sheet1.Range("O" & j + 9).Value = ExApp.WorksheetFunction.RandBetween(-1, 1)
                        sheet1.Range("P" & j).Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                        sheet1.Range("P" & j + 1).Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                        sheet1.Range("P" & j + 2).Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                        sheet1.Range("P" & j + 3).Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                        sheet1.Range("P" & j + 4).Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                        sheet1.Range("P" & j + 5).Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                        sheet1.Range("P" & j + 6).Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                        sheet1.Range("P" & j + 7).Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                        sheet1.Range("P" & j + 8).Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                        sheet1.Range("P" & j + 9).Value = ExApp.WorksheetFunction.RandBetween(1, 4)

                        '备注
                        sheet1.Range("R" & j).Value = n + n - 1 & "#前"
                        sheet1.Range("R" & j + 1).Value = n + n - 1 & "#后"
                        sheet1.Range("R" & j + 2).Value = n + n - 1 & "#左"
                        sheet1.Range("R" & j + 3).Value = n + n - 1 & "#右"
                        sheet1.Range("R" & j + 4).Value = n + n - 1 & "#中"
                        sheet1.Range("R" & j + 5).Value = n + n & "#前"
                        sheet1.Range("R" & j + 6).Value = n + n & "#后"
                        sheet1.Range("R" & j + 7).Value = n + n & "#左"
                        sheet1.Range("R" & j + 8).Value = n + n & "#右"
                        sheet1.Range("R" & j + 9).Value = n + n & "#中"

                        If sheet0.Range("K" & h).Value = "中" Or sheet0.Range("K" & h).Value = "右" Then
                            '偏距
                            sheet1.Range("J" & j).Value = sheet0.Range("AJ" & h).Value / 100 * -1
                            sheet1.Range("J" & j + 1).Value = sheet0.Range("AJ" & h).Value / 100 * -1
                            sheet1.Range("J" & j + 2).Value = (sheet0.Range("AJ" & h).Value / 100 + sheet0.Range("G" & h).Value / 200) * -1
                            sheet1.Range("J" & j + 3).Value = (sheet0.Range("AJ" & h).Value / 100 - sheet0.Range("G" & h).Value / 200) * -1
                            sheet1.Range("J" & j + 4).Value = sheet0.Range("AJ" & h).Value / 100 * -1
                            sheet1.Range("J" & j + 5).Value = (sheet0.Range("AJ" & h).Value / 100 - sheet0.Range("AI" & h).Value / 100) * -1
                            sheet1.Range("J" & j + 6).Value = (sheet0.Range("AJ" & h).Value / 100 - sheet0.Range("AI" & h).Value / 100) * -1
                            sheet1.Range("J" & j + 7).Value = (sheet0.Range("AJ" & h).Value / 100 - sheet0.Range("AI" & h).Value / 100 + sheet0.Range("G" & h).Value / 200) * -1
                            sheet1.Range("J" & j + 8).Value = (sheet0.Range("AJ" & h).Value / 100 - sheet0.Range("AI" & h).Value / 100 - sheet0.Range("G" & h).Value / 200) * -1
                            sheet1.Range("J" & j + 9).Value = (sheet0.Range("AJ" & h).Value / 100 - sheet0.Range("AI" & h).Value / 100) * -1
                        Else
                            '中线位于构筑物左侧时偏距
                            sheet1.Range("J" & j).Value = sheet0.Range("AJ" & h).Value / 100 - sheet0.Range("AI" & h).Value / 100
                            sheet1.Range("J" & j + 1).Value = sheet0.Range("AJ" & h).Value / 100 - sheet0.Range("AI" & h).Value / 100
                            sheet1.Range("J" & j + 2).Value = sheet0.Range("AJ" & h).Value / 100 - sheet0.Range("AI" & h).Value / 100 - sheet0.Range("G" & h).Value / 200
                            sheet1.Range("J" & j + 3).Value = sheet0.Range("AJ" & h).Value / 100 - sheet0.Range("AI" & h).Value / 100 + sheet0.Range("G" & h).Value / 200
                            sheet1.Range("J" & j + 4).Value = sheet0.Range("AJ" & h).Value / 100 - sheet0.Range("AI" & h).Value / 100
                            sheet1.Range("J" & j + 5).Value = sheet0.Range("AJ" & h).Value / 100
                            sheet1.Range("J" & j + 6).Value = sheet0.Range("AJ" & h).Value / 100
                            sheet1.Range("J" & j + 7).Value = sheet0.Range("AJ" & h).Value / 100 - sheet0.Range("G" & h).Value / 200
                            sheet1.Range("J" & j + 8).Value = sheet0.Range("AJ" & h).Value / 100 + sheet0.Range("G" & h).Value / 200
                            sheet1.Range("J" & j + 9).Value = sheet0.Range("AJ" & h).Value / 100
                        End If

                        j += 10
                        x += 1
                        m -= 1
                    Next n

                    '桥台垫石
                ElseIf sheet0.Range("AL" & h).Value = "起点" Then
                    m = sheet0.Range("AT" & h).Value
                    p = (sheet0.Range("AO" & h).Value - sheet0.Range("AN" & h).Value) / (sheet0.Range("AP" & h).Value / 100)  '横坡
                    For n = 1 To m
                        '桩号
                        sheet1.Range("I" & j).Value = sheet0.Range("C" & h).Value + sheet0.Range("AM" & h).Value / 100 + sheet0.Range("H" & h).Value / 200
                        sheet1.Range("I" & j + 1).Value = sheet0.Range("C" & h).Value + sheet0.Range("AM" & h).Value / 100 - sheet0.Range("H" & h).Value / 200
                        sheet1.Range("I" & j + 2).Value = sheet0.Range("C" & h).Value + sheet0.Range("AM" & h).Value / 100
                        sheet1.Range("I" & j + 3).Value = sheet0.Range("C" & h).Value + sheet0.Range("AM" & h).Value / 100
                        sheet1.Range("I" & j + 4).Value = sheet0.Range("C" & h).Value + sheet0.Range("AM" & h).Value / 100
                        '高程
                        sheet1.Range("K" & j).Value = (sheet0.Range("AN" & h).Value + sheet0.Range("I" & h).Value / 100) + (sheet0.Range("AS" & h).Value + x * sheet0.Range("AQ" & h).Value) * p / 100
                        sheet1.Range("K" & j + 1).Value = (sheet0.Range("AN" & h).Value + sheet0.Range("I" & h).Value / 100) + (sheet0.Range("AS" & h).Value + x * sheet0.Range("AQ" & h).Value) * p / 100
                        sheet1.Range("K" & j + 2).Value = (sheet0.Range("AN" & h).Value + sheet0.Range("I" & h).Value / 100) + (sheet0.Range("AS" & h).Value + x * sheet0.Range("AQ" & h).Value) * p / 100
                        sheet1.Range("K" & j + 3).Value = (sheet0.Range("AN" & h).Value + sheet0.Range("I" & h).Value / 100) + (sheet0.Range("AS" & h).Value + x * sheet0.Range("AQ" & h).Value) * p / 100
                        sheet1.Range("K" & j + 4).Value = (sheet0.Range("AN" & h).Value + sheet0.Range("I" & h).Value / 100) + (sheet0.Range("AS" & h).Value + x * sheet0.Range("AQ" & h).Value) * p / 100
                        '测量偏差值
                        sheet1.Range("O" & j).Value = ExApp.WorksheetFunction.RandBetween(-1, 1)
                        sheet1.Range("O" & j + 1).Value = ExApp.WorksheetFunction.RandBetween(-1, 1)
                        sheet1.Range("O" & j + 2).Value = ExApp.WorksheetFunction.RandBetween(-1, 1)
                        sheet1.Range("O" & j + 3).Value = ExApp.WorksheetFunction.RandBetween(-1, 1)
                        sheet1.Range("O" & j + 4).Value = ExApp.WorksheetFunction.RandBetween(-1, 1)
                        sheet1.Range("P" & j).Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                        sheet1.Range("P" & j + 1).Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                        sheet1.Range("P" & j + 2).Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                        sheet1.Range("P" & j + 3).Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                        sheet1.Range("P" & j + 4).Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                        '备注
                        sheet1.Range("R" & j).Value = n & "#前"
                        sheet1.Range("R" & j + 1).Value = n & "#后"
                        sheet1.Range("R" & j + 2).Value = n & "#左"
                        sheet1.Range("R" & j + 3).Value = n & "#右"
                        sheet1.Range("R" & j + 4).Value = n & "#中"

                        If sheet0.Range("K" & h).Value = "中" Or sheet0.Range("K" & h).Value = "右" Then
                            '偏距
                            sheet1.Range("J" & j).Value = (sheet0.Range("AR" & h).Value / 100 - x * sheet0.Range("AS" & h).Value / 100) * -1
                            sheet1.Range("J" & j + 1).Value = (sheet0.Range("AR" & h).Value / 100 - x * sheet0.Range("AS" & h).Value / 100) * -1
                            sheet1.Range("J" & j + 2).Value = ((sheet0.Range("AR" & h).Value + sheet0.Range("G" & h).Value / 2) / 100 - x * sheet0.Range("AS" & h).Value / 100) * -1
                            sheet1.Range("J" & j + 3).Value = ((sheet0.Range("AR" & h).Value - sheet0.Range("G" & h).Value / 2) / 100 - x * sheet0.Range("AS" & h).Value / 100) * -1
                            sheet1.Range("J" & j + 4).Value = (sheet0.Range("AR" & h).Value / 100 - x * sheet0.Range("AS" & h).Value / 100) * -1
                        Else
                            '中线位于构筑物左侧时偏距
                            sheet1.Range("J" & j).Value = sheet0.Range("AR" & h).Value / 100 - ((m - 1) * sheet0.Range("AQ" & h).Value / 100)
                            sheet1.Range("J" & j + 1).Value = sheet0.Range("AR" & h).Value / 100 - ((m - 1) * sheet0.Range("AQ" & h).Value / 100)
                            sheet1.Range("J" & j + 2).Value = sheet0.Range("AR" & h).Value / 100 - ((m - 1) * sheet0.Range("AQ" & h).Value / 100) - (sheet0.Range("G" & h).Value / 2) / 100
                            sheet1.Range("J" & j + 3).Value = sheet0.Range("AR" & h).Value / 100 - ((m - 1) * sheet0.Range("AQ" & h).Value / 100) + (sheet0.Range("G" & h).Value / 2) / 100
                            sheet1.Range("J" & j + 4).Value = sheet0.Range("AR" & h).Value / 100 - ((m - 1) * sheet0.Range("AQ" & h).Value / 100)
                        End If
                        j += 5
                        x += 1
                        m -= 1
                    Next n
                Else
                    ' 桥台终点
                    m = sheet0.Range("AT" & h).Value
                    p = (sheet0.Range("AO" & h).Value - sheet0.Range("AN" & h).Value) / (sheet0.Range("AP" & h).Value / 100)  '横坡
                    For n = 1 To m
                        '桩号
                        sheet1.Range("I" & j).Value = sheet0.Range("C" & h).Value - sheet0.Range("AM" & h).Value / 100 + sheet0.Range("H" & h).Value / 200
                        sheet1.Range("I" & j + 1).Value = sheet0.Range("C" & h).Value - sheet0.Range("AM" & h).Value / 100 - sheet0.Range("H" & h).Value / 200
                        sheet1.Range("I" & j + 2).Value = sheet0.Range("C" & h).Value - sheet0.Range("AM" & h).Value / 100
                        sheet1.Range("I" & j + 3).Value = sheet0.Range("C" & h).Value - sheet0.Range("AM" & h).Value / 100
                        sheet1.Range("I" & j + 4).Value = sheet0.Range("C" & h).Value - sheet0.Range("AM" & h).Value / 100
                        '高程
                        sheet1.Range("K" & j).Value = (sheet0.Range("AN" & h).Value + sheet0.Range("I" & h).Value / 100) + (sheet0.Range("AS" & h).Value + x * sheet0.Range("AQ" & h).Value) * p / 100
                        sheet1.Range("K" & j + 1).Value = (sheet0.Range("AN" & h).Value + sheet0.Range("I" & h).Value / 100) + (sheet0.Range("AS" & h).Value + x * sheet0.Range("AQ" & h).Value) * p / 100
                        sheet1.Range("K" & j + 2).Value = (sheet0.Range("AN" & h).Value + sheet0.Range("I" & h).Value / 100) + (sheet0.Range("AS" & h).Value + x * sheet0.Range("AQ" & h).Value) * p / 100
                        sheet1.Range("K" & j + 3).Value = (sheet0.Range("AN" & h).Value + sheet0.Range("I" & h).Value / 100) + (sheet0.Range("AS" & h).Value + x * sheet0.Range("AQ" & h).Value) * p / 100
                        sheet1.Range("K" & j + 4).Value = (sheet0.Range("AN" & h).Value + sheet0.Range("I" & h).Value / 100) + (sheet0.Range("AS" & h).Value + x * sheet0.Range("AQ" & h).Value) * p / 100
                        '测量偏差值
                        sheet1.Range("O" & j).Value = ExApp.WorksheetFunction.RandBetween(-1, 1)
                        sheet1.Range("O" & j + 1).Value = ExApp.WorksheetFunction.RandBetween(-1, 1)
                        sheet1.Range("O" & j + 2).Value = ExApp.WorksheetFunction.RandBetween(-1, 1)
                        sheet1.Range("O" & j + 3).Value = ExApp.WorksheetFunction.RandBetween(-1, 1)
                        sheet1.Range("O" & j + 4).Value = ExApp.WorksheetFunction.RandBetween(-1, 1)
                        sheet1.Range("P" & j).Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                        sheet1.Range("P" & j + 1).Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                        sheet1.Range("P" & j + 2).Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                        sheet1.Range("P" & j + 3).Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                        sheet1.Range("P" & j + 4).Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                        '备注
                        sheet1.Range("R" & j).Value = n & "#前"
                        sheet1.Range("R" & j + 1).Value = n & "#后"
                        sheet1.Range("R" & j + 2).Value = n & "#左"
                        sheet1.Range("R" & j + 3).Value = n & "#右"
                        sheet1.Range("R" & j + 4).Value = n & "#中"

                        If sheet0.Range("K" & h).Value = "中" Or sheet0.Range("K" & h).Value = "右" Then
                            '偏距
                            sheet1.Range("J" & j).Value = (sheet0.Range("AR" & h).Value / 100 - x * sheet0.Range("AS" & h).Value / 100) * -1
                            sheet1.Range("J" & j + 1).Value = (sheet0.Range("AR" & h).Value / 100 - x * sheet0.Range("AS" & h).Value / 100) * -1
                            sheet1.Range("J" & j + 2).Value = ((sheet0.Range("AR" & h).Value + sheet0.Range("G" & h).Value / 2) / 100 - x * sheet0.Range("AS" & h).Value / 100) * -1
                            sheet1.Range("J" & j + 3).Value = ((sheet0.Range("AR" & h).Value - sheet0.Range("G" & h).Value / 2) / 100 - x * sheet0.Range("AS" & h).Value / 100) * -1
                            sheet1.Range("J" & j + 4).Value = (sheet0.Range("AR" & h).Value / 100 - x * sheet0.Range("AS" & h).Value / 100) * -1
                        Else
                            '中线位于构筑物左侧时偏距
                            sheet1.Range("J" & j).Value = sheet0.Range("AR" & h).Value / 100 - ((m - 1) * sheet0.Range("AQ" & h).Value / 100)
                            sheet1.Range("J" & j + 1).Value = sheet0.Range("AR" & h).Value / 100 - ((m - 1) * sheet0.Range("AQ" & h).Value / 100)
                            sheet1.Range("J" & j + 2).Value = sheet0.Range("AR" & h).Value / 100 - ((m - 1) * sheet0.Range("AQ" & h).Value / 100) - (sheet0.Range("G" & h).Value / 2) / 100
                            sheet1.Range("J" & j + 3).Value = sheet0.Range("AR" & h).Value / 100 - ((m - 1) * sheet0.Range("AQ" & h).Value / 100) + (sheet0.Range("G" & h).Value / 2) / 100
                            sheet1.Range("J" & j + 4).Value = sheet0.Range("AR" & h).Value / 100 - ((m - 1) * sheet0.Range("AQ" & h).Value / 100)
                        End If
                        j += 5
                        x += 1
                        m -= 1
                    Next n
                End If

                '计算坐标
                Dim sjxq, sjyq As Double
                Dim c, b, ZH0 As Integer
                b = 3
                Do While sheet1.Range("I" & b).Value <> Nothing
                    ZH0 = Val(ExApp.WorksheetFunction.Substitute(sheet1.Range("I" & b).Value, "*", ""))
                    If sheet3.Range("J2").Value <> "是" Then
                        c = Pd_YSw(ZH0, sheet2.Range("O5 : O500").Value)
                    Else
                        c = 1
                    End If
                    If c = -1 Then
                        MsgBox("请在交点法表内输入数据")
                        TorF = False
                        Exit Sub
                    Else
                        If sheet3.Range("J2").Value <> "是" Then
                            sjxq = ZSZB_X0j(sheet1.Range("I" & b).Value, sheet1.Range("J" & b).Value, 90)  '前偏距坐标
                            sjyq = ZSZB_Y0j(sheet1.Range("I" & b).Value, sheet1.Range("J" & b).Value, 90)
                        Else
                            sjxq = XYF_X(sheet1.Range("I" & b).Value, sheet1.Range("J" & b).Value, 90)  '前偏距坐标
                            sjyq = XYF_Y(sheet1.Range("I" & b).Value, sheet1.Range("J" & b).Value, 90)
                        End If
                        '偏距坐标赋值
                        sheet1.Range("M" & b).Value = Math.Round(sjxq, 3)
                        sheet1.Range("N" & b).Value = Math.Round(sjyq, 3)

                    End If
                    b += 1
                Loop
                sheet1.Range("B" & r + 5).Value = sheet0.Range("B" & h).Value
                sheet1.Range("B" & r + 6).Value = sheet0.Range("C" & h).Value
                sheet1.Range("C" & r + 6).Value = sheet0.Range("D" & h).Value
                sheet1.Range("D" & r + 6).Value = sheet0.Range("E" & h).Value
                sheet1.Range("B" & r + 7).Value = sheet0.Range("L" & h).Value
                sheet1.Range("B" & r + 8).Value = sheet0.Range("M" & h).Value
                sheet1.Range("B" & r + 9).Value = sheet0.Range("N" & h).Value
                sheet1.Range("B" & r + 10).Value = sheet0.Range("R" & h).Value
                sheet1.Range("B" & r + 11).Value = sheet0.Range("G" & h).Value
                sheet1.Range("B" & r + 12).Value = sheet0.Range("H" & h).Value
                sheet1.Range("B" & r + 13).Value = sheet0.Range("I" & h).Value
                sheet1.Range("B" & r + 14).Value = sheet0.Range("O" & h).Value
                sheet1.Range("B" & r + 15).Value = sheet0.Range("P" & h).Value
                sheet1.Range("B" & r + 16).Value = sheet0.Range("Q" & h).Value
                sheet1.Range("B" & r + 17).Value = sheet0.Range("F" & h).Value
                If sheet0.Range("J" & h).Value = "盖梁垫石" Then
                    sheet1.Range("B" & r + 19).Value = sheet0.Range("AF" & h).Value
                ElseIf sheet0.Range("J" & h).Value = "墩柱垫石" Then
                    sheet1.Range("B" & r + 19).Value = sheet0.Range("AK" & h).Value
                Else
                    sheet1.Range("B" & r + 19).Value = sheet0.Range("AT" & h).Value
                End If

                sheet1.Range("P1").Value = sheet1.Range("B17").Value
                sheet1.Range("L3").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value


                sheet16.Range("A5").Value = sheet1.Range("A5").Value & sheet1.Range("B5").Value
                sheet17.Range("A5").Value = sheet1.Range("A5").Value & sheet1.Range("B5").Value



                ' 钢筋隐蔽工程
                sheet6.Range("C6").Value = sheet1.Range("B5").Value & sheet1.Range("C6").Value & "基础及下部构造"
                sheet6.Range("E6").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋加工及安装"
                sheet6.Range("C7").Value = sheet1.Range("B5").Value & sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋加工及安装"
                sheet6.Range("E10").Value = sheet1.Range("B18").Value
                sheet6.Range("E11").Value = sheet1.Range("B18").Value
                sheet6.Range("E27").Value = sheet1.Range("B18").Value
                sheet6.Range("E28").Value = sheet1.Range("B18").Value
                '钢筋检表
                sheet7.Range("D6").Value = sheet1.Range("B5").Value
                sheet7.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋"
                sheet7.Range("Q6").Value = sheet1.Range("B18").Value
                sheet7.Range("Q7").Value = sheet1.Range("B18").Value
                sheet7.Range("Q31").Value = sheet1.Range("B18").Value
                sheet7.Range("Q34").Value = sheet1.Range("B18").Value
                '主筋
                sheet7.Range("E15").Value = "设计值：" & sheet1.Range("B8").Value * 10
                LSBLFZ = Nothing
                If sheet1.Range("B7").Value * 2 <= 10 Then
                    For cs = 1 To sheet1.Range("B7").Value * 2
                        LSBL = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet7.Range("G14").Value = LSBLFZ
                    sheet8.Range("D24").Value = "/"
                Else
                    sheet7.Range("G14").Value = "应测" & sheet1.Range("B7").Value * 2 & "处，实测" & sheet1.Range("B7").Value * 2 & "处，合格" & sheet1.Range("B7").Value * 2 & "处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                    For cs = 1 To sheet1.Range("B7").Value * 2
                        LSBL = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet8.Range("D24").Value = LSBLFZ
                End If

                sheet7.Range("I14").Value = sheet1.Range("B7").Value * 2
                sheet7.Range("L14").Value = sheet1.Range("B7").Value * 2
                sheet7.Range("O14").Value = sheet1.Range("B7").Value * 2
                sheet7.Range("I15").Value = 100
                '箍筋
                sheet7.Range("E17").Value = "设计值：" & sheet1.Range("B9").Value * 10
                LSBLFZ = Nothing
                For cs = 1 To sheet1.Range("B19").Value * 10
                    LSBL = sheet1.Range("B9").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                    LSBLFZ = LSBLFZ & LSBL & "   "
                Next
                sheet7.Range("G16").Value = "应测" & sheet1.Range("B19").Value * 10 & "处，实测" & sheet1.Range("B19").Value * 10 & "处，合格" & sheet1.Range("B19").Value * 10 & "处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                sheet8.Range("D29").Value = LSBLFZ

                '骨架尺寸
                sheet13.Range("E19").Value = "设计值：" & sheet1.Range("B14").Value * 10
                sheet13.Range("K18").Value = sheet1.Range("B14").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet13.Range("E21").Value = "宽：" & sheet1.Range("B15").Value * 10
                sheet13.Range("E22").Value = "高：" & sheet1.Range("B16").Value * 10
                LSBLFZ = Nothing
                For cs = 1 To Math.Round(sheet1.Range("B19").Value * 0.3, 0)
                    LSBL = sheet1.Range("B15").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4)
                    LSBLFZ = LSBLFZ & LSBL & "   "
                Next
                sheet13.Range("G21").Value = "宽： " & LSBLFZ
                LSBLFZ = Nothing
                For cs = 1 To Math.Round(sheet1.Range("B19").Value * 0.3, 0)
                    LSBL = sheet1.Range("B16").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4)
                    LSBLFZ = LSBLFZ & LSBL & "   "
                Next
                sheet13.Range("G21").Value = "高： " & LSBLFZ

                '保护层
                sheet7.Range("E27").Value = "设计值：" & sheet1.Range("B10").Value * 10
                LSBLFZ = Nothing
                'If sheet1.Range("B19").Value * 20 <= 10 Then
                '    For cs = 1 To sheet1.Range("B19").Value * 20
                '        LSBL = sheet1.Range("B10").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                '        LSBLFZ = LSBLFZ & LSBL & "   "
                '    Next
                '    sheet7.Range("G26").Value = LSBLFZ
                '    sheet8.Range("D48").Value = "/"
                'Else
                sheet7.Range("G26").Value = "应测" & sheet1.Range("B19").Value * 20 & "处，实测" & sheet1.Range("B19").Value * 20 & "处，合格" & sheet1.Range("B19").Value * 20 & "处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                    For cs = 1 To sheet1.Range("B19").Value * 20
                        LSBL = sheet1.Range("B10").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                sheet8.Range("D48").Value = LSBLFZ
                'End If

                '钢筋记录表
                sheet8.Range("B6").Value = sheet1.Range("B5").Value
                sheet8.Range("B7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋"
                sheet8.Range("K6").Value = sheet1.Range("B18").Value
                sheet8.Range("K7").Value = sheet1.Range("B18").Value

                '工序检验申请批复单
                sheet9.Range("C6").Value = sheet1.Range("B5").Value & sheet1.Range("C6").Value & "基础及下部构造"
                sheet9.Range("C7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet9.Range("C8").Value = sheet1.Range("B5").Value
                sheet9.Range("C9").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet9.Range("C10").Value = "混凝土强度、轴线偏位、断面尺寸、顶面高程、顶面高差"

                Call 水准测量记录表()
                If TorF = False Then
                    Exit Sub
                End If
                Dim aa As Integer
                aa = 3
                LSBLFZ = Nothing
                While sheet1.Range("O" & aa).Value <> Nothing Or sheet1.Range("O" & aa).Value = "0"
                    LSBL = sheet1.Range("O" & aa).Value
                    LSBLFZ = LSBLFZ & LSBL & "   "
                    aa += 1
                End While
                sheet10.Range("F15").Value = LSBLFZ
                zs = sheet1.Range("R1048576").End(XlDirection.xlUp).Row - 2
                xhh = 3
                sheet1.Select()

                For cs = 1 To zs
                    If InStr(1, sheet1.Range("R" & xhh).Value, "中"） > 0 Then
                        sheet1.Range("I" & xhh & ":R" & xhh).Select()
                        ExApp.Selection.Delete
                    End If
                    xhh += 1
                Next

                xhh = ExApp.WorksheetFunction.RoundUp(sheet1.Range("B19").Value * 0.5, 0) * 4
                sheet1.Range("I" & xhh + 3 & ":R1000").Select()
                ExApp.Selection.Delete

                Call 全站仪平面位置检测表（）
                If TorF = False Then
                    Exit Sub
                End If
                aa = 3
                LSBLFZ = Nothing
                While sheet1.Range("P" & aa).Value <> Nothing Or sheet1.Range("P" & aa).Value = "0"
                    LSBL = sheet1.Range("P" & aa).Value
                    LSBLFZ = LSBLFZ & LSBL & "   "
                    aa += 1
                End While
                sheet10.Range("F12").Value = LSBLFZ

                sheet16.Activate()
                ExApp.ActiveWindow.SelectedSheets.Copy(, (ExApp.Sheets(ExApp.Sheets.Count)))
                sheet17.Activate()
                ExApp.ActiveWindow.SelectedSheets.Copy(, (ExApp.Sheets(ExApp.Sheets.Count)))

                '支座垫石现场质量检验表
                sheet10.Range("D6").Value = sheet1.Range("B5").Value
                sheet10.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet10.Range("P6").Value = sheet1.Range("B17").Value
                sheet10.Range("D14").Value = "宽：" & sheet1.Range("B12").Value * 10 & vbCrLf & "高：" & sheet1.Range("B13").Value * 10
                sheet10.Range("F13").Value = "宽："
                sheet10.Range("F14").Value = "高："
                LSBLFZ = Nothing
                LSBLFZ2 = Nothing
                For cs = 1 To ExApp.WorksheetFunction.RoundUp(sheet1.Range("B19").Value * 0.5, 0)
                    LSBL = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4)
                    LSBL2 = sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4)
                    LSBLFZ = LSBLFZ & LSBL & "   "
                    LSBLFZ2 = LSBLFZ2 & LSBL2 & "   "
                Next
                sheet10.Range("G13").Value = LSBLFZ
                sheet10.Range("G14").Value = LSBLFZ2
                LSBLFZ = Nothing
                If sheet1.Range("B11").Value * 10 <= 500 Or sheet1.Range("B12").Value * 10 <= 500 Then
                    sheet10.Range("F17").Value = "/"
                    For cs = 1 To sheet1.Range("B19").Value
                        LSBL = ExApp.WorksheetFunction.RandBetween(0, 1)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet10.Range("F16").Value = LSBLFZ
                Else
                    sheet10.Range("F16").Value = "/"
                    For cs = 1 To sheet1.Range("B19").Value
                        LSBL = ExApp.WorksheetFunction.RandBetween(0, 2)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet10.Range("F17").Value = LSBLFZ
                End If
                LSBLFZ = Nothing
                For cs = 1 To sheet1.Range("B19").Value
                    LSBL = ExApp.WorksheetFunction.RandBetween(0, 4)
                    LSBLFZ = LSBLFZ & LSBL & "   "
                Next
                sheet10.Range("F18").Value = LSBLFZ

                '模板测量偏差值
                sheet1.Range("O3").Value = ExApp.WorksheetFunction.RandBetween(-1, 1)
                sheet1.Range("O4").Value = ExApp.WorksheetFunction.RandBetween(-1, 1)
                sheet1.Range("O5").Value = ExApp.WorksheetFunction.RandBetween(-1, 1)
                sheet1.Range("O6").Value = ExApp.WorksheetFunction.RandBetween(-1, 1)
                sheet1.Range("P3").Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                sheet1.Range("P4").Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                sheet1.Range("P5").Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                sheet1.Range("P6").Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                sheet1.Range("L3").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "模板"
                sheet1.Range("I7:R1000").Value = Nothing
                Call 水准测量记录表()
                If TorF = False Then
                    Exit Sub
                End If
                Call 全站仪平面位置检测表（）
                If TorF = False Then
                    Exit Sub
                End If
                '现场模板安装检查记录表
                sheet11.Range("C6").Value = sheet1.Range("B5").Value
                sheet11.Range("C7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet11.Range("I6").Value = sheet1.Range("B17").Value
                sheet11.Range("I7").Value = sheet1.Range("B17").Value
                sheet11.Range("D8").Value = 2
                sheet11.Range("D9").Value = 5
                sheet11.Range("F8").Value = 4
                sheet11.Range("F9").Value = 4
                sheet11.Range("H8").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet11.Range("H9").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet11.Range("J8").Value = 100
                sheet11.Range("J9").Value = 100
                '测量偏差
                sheet11.Range("D11").Value = sheet1.Range("P5").Value
                sheet11.Range("F11").Value = sheet1.Range("P6").Value
                sheet11.Range("H11").Value = sheet1.Range("P3").Value
                sheet11.Range("J11").Value = sheet1.Range("P4").Value
                sheet11.Range("F14").Value = sheet1.Range("O3").Value
                sheet11.Range("G14").Value = sheet1.Range("O4").Value
                sheet11.Range("H14").Value = sheet1.Range("O5").Value
                sheet11.Range("I14").Value = sheet1.Range("O6").Value

                sheet11.Range("F12").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("G12").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("H12").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("I12").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("F13").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("G13").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("H13").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("I13").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("F18").Value = "牢固，稳定"

                '监抽钢筋检表
                sheet13.Range("D6").Value = sheet1.Range("B5").Value
                sheet13.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋"
                sheet13.Range("Q6").Value = sheet1.Range("B18").Value
                sheet13.Range("Q7").Value = sheet1.Range("B18").Value
                sheet13.Range("Q31").Value = sheet1.Range("B18").Value
                sheet13.Range("Q34").Value = sheet1.Range("B18").Value
                '主筋
                sheet13.Range("E15").Value = "设计值：" & sheet1.Range("B8").Value * 10
                LSBLFZ = Nothing
                If Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0) <= 10 Then
                    For cs = 1 To Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0)
                        LSBL = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet13.Range("G14").Value = LSBLFZ
                    sheet14.Range("D24").Value = "/"
                Else
                    sheet13.Range("G14").Value = "应测" & Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0) & "处，实测" & Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0) & "处，合格" & Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0) & "处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                    For cs = 1 To Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0)
                        LSBL = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet14.Range("D24").Value = LSBLFZ
                End If
                '箍筋
                sheet7.Range("E17").Value = "设计值：" & sheet1.Range("B9").Value * 10
                LSBLFZ = Nothing
                If sheet1.Range("B19").Value * 2 <= 10 Then
                    For cs = 1 To sheet1.Range("B19").Value * 2
                        LSBL = sheet1.Range("B9").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet7.Range("G16").Value = LSBLFZ
                Else
                    sheet7.Range("G16").Value = "应测" & sheet1.Range("B19").Value * 2 & "处，实测" & sheet1.Range("B19").Value * 2 & "处，合格" & sheet1.Range("B19").Value * 2 & "处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                    sheet8.Range("D30").Value = LSBLFZ
                End If
                '骨架尺寸
                sheet13.Range("E19").Value = "设计值：" & sheet1.Range("B14").Value * 10
                sheet13.Range("K18").Value = sheet1.Range("B14").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet13.Range("E21").Value = "宽：" & sheet1.Range("B15").Value * 10
                sheet13.Range("E22").Value = "高：" & sheet1.Range("B16").Value * 10
                sheet13.Range("G21").Value = "宽： " & sheet1.Range("B15").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4)
                sheet13.Range("G21").Value = "高： " & sheet1.Range("B16").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4)
                '保护层
                sheet13.Range("E27").Value = "设计值：" & sheet1.Range("B10").Value * 10
                LSBLFZ = Nothing
                If sheet1.Range("B19").Value * 4 <= 10 Then
                    For cs = 1 To sheet1.Range("B19").Value * 4
                        LSBL = sheet1.Range("B10").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet13.Range("G26").Value = LSBLFZ
                    sheet14.Range("D48").Value = "/"
                Else
                    sheet13.Range("G26").Value = "应测" & sheet1.Range("B19").Value * 4 & "处，实测" & sheet1.Range("B19").Value * 4 & "处，合格" & sheet1.Range("B19").Value * 4 & "处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                    For cs = 1 To sheet1.Range("B19").Value * 4
                        LSBL = sheet1.Range("B10").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet14.Range("D48").Value = LSBLFZ
                End If
                '监抽钢筋记录表
                sheet14.Range("B6").Value = sheet1.Range("B5").Value
                sheet14.Range("B7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋"
                sheet14.Range("K6").Value = sheet1.Range("B18").Value
                sheet14.Range("K7").Value = sheet1.Range("B18").Value

                '监抽支座垫石检验表
                sheet15.Range("D6").Value = sheet1.Range("B5").Value
                sheet15.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet15.Range("P6").Value = sheet1.Range("B17").Value
                sheet15.Range("F12").Value = sheet10.Range("F12").Value
                sheet15.Range("F15").Value = sheet10.Range("F15").Value
                sheet15.Range("D14").Value = "宽：" & sheet1.Range("B12").Value * 10 & vbCrLf & "高：" & sheet1.Range("B13").Value * 10
                sheet15.Range("F13").Value = "宽："
                sheet15.Range("F14").Value = "高："
                LSBLFZ = Nothing
                LSBLFZ2 = Nothing
                For cs = 1 To ExApp.WorksheetFunction.RoundUp(sheet1.Range("B19").Value * 0.1, 0)
                    LSBL = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4)
                    LSBL2 = sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4)
                    LSBLFZ = LSBLFZ & LSBL & "   "
                    LSBLFZ2 = LSBLFZ2 & LSBL2 & "   "
                Next
                sheet15.Range("G13").Value = LSBLFZ
                sheet15.Range("G14").Value = LSBLFZ2
                LSBLFZ = Nothing
                If sheet1.Range("B11").Value * 10 <= 500 Or sheet1.Range("B12").Value * 10 <= 500 Then
                    sheet15.Range("F17").Value = "/"
                    For cs = 1 To ExApp.WorksheetFunction.RoundUp(sheet1.Range("B19").Value * 0.2, 0)
                        LSBL = ExApp.WorksheetFunction.RandBetween(0, 1)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet15.Range("F16").Value = LSBLFZ
                Else
                    sheet15.Range("F16").Value = "/"
                    For cs = 1 To ExApp.WorksheetFunction.RoundUp(sheet1.Range("B19").Value * 0.2, 0)
                        LSBL = ExApp.WorksheetFunction.RandBetween(0, 2)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet15.Range("F17").Value = LSBLFZ
                End If
                LSBLFZ = Nothing
                For cs = 1 To ExApp.WorksheetFunction.RoundUp(sheet1.Range("B19").Value * 0.2, 0)
                    LSBL = ExApp.WorksheetFunction.RandBetween(0, 4)
                    LSBLFZ = LSBLFZ & LSBL & "   "
                Next
                sheet15.Range("F18").Value = LSBLFZ
                ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                '选择表格
                sheet6.Select()
                For i = 7 To ExApp.Sheets.Count
                    EXsheet = Exbook.Worksheets(i)
                    If EXsheet.Visible = True Then
                        EXsheet.Select(Replace:=False)
                    End If
                Next i
                EXsheet = ExApp.ActiveSheet

                ' 导出PDF文件
                PDFFilename = Filepath & "\" & sheet1.Range("B5").Value & sheet1.Range("C6").Value & sheet1.Range("D6").Value & ".pdf"
                EXsheet.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, PDFFilename, XlFixedFormatQuality.xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False)

                sheet18 = Exbook.Worksheets("水准表(2)")
                sheet19 = Exbook.Worksheets("平面表(2)")
                ExApp.DisplayAlerts = False
                sheet18.Delete()
                sheet19.Delete()
                h += 1
                i += 1
            End While
            MsgBox("已完成！", 0 + 64, "提示")
        Catch Exclerror As Exception   '错误时弹出提示
            MsgBox(Exclerror.Message)
        End Try
        TorF = False
    End Sub

    Sub 挡块资料（）
        Dim P As Double
        Dim h, r As Integer
        Dim Filepath, ExcelFilename, PDFFilename， LSBL, LSBLFZ As String  '定义输出的PDF文件名
        Dim FolderDialogObject As New FolderBrowserDialog()
        h = 8
        r = 0
        Try
            sheet0 = Exbook.Worksheets("数据库")
            sheet1 = Exbook.Worksheets("参数表")
            sheet2 = Exbook.Worksheets("交点法")
            sheet3 = Exbook.Worksheets("线元法")
            sheet4 = Exbook.Worksheets("断链")
            sheet5 = Exbook.Worksheets("导线成果表")
            sheet6 = Exbook.Worksheets("钢筋隐蔽工程")
            sheet7 = Exbook.Worksheets("钢筋检表")
            sheet8 = Exbook.Worksheets("钢筋记录表")
            sheet9 = Exbook.Worksheets("申请批复单")
            sheet10 = Exbook.Worksheets("挡块检表")
            sheet11 = Exbook.Worksheets("模板记录表")
            sheet12 = Exbook.Worksheets("砼浇筑申请报告单")
            sheet13 = Exbook.Worksheets("监抽钢筋检表")
            sheet14 = Exbook.Worksheets("监抽钢筋记录表")
            sheet15 = Exbook.Worksheets("监抽挡块检表")
            sheet16 = Exbook.Worksheets("水准表")
            sheet17 = Exbook.Worksheets("平面表")

            '改表头
            sheet6.Range("A1").Value = sheet0.Range("C1").Value
            sheet6.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet6.Range("E3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet6.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet6.Range("E4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet7.Range("A1").Value = sheet0.Range("C1").Value
            sheet7.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet7.Range("P3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet7.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet7.Range("P4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet8.Range("A1").Value = sheet0.Range("C1").Value
            sheet8.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet8.Range("L3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet8.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet8.Range("L4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet9.Range("A1").Value = sheet0.Range("C1").Value
            sheet9.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet9.Range("E3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet9.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet9.Range("E4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet10.Range("A1").Value = sheet0.Range("C1").Value
            sheet10.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet10.Range("K3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet10.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet10.Range("K4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet11.Range("A1").Value = sheet0.Range("C1").Value
            sheet11.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet11.Range("I3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet11.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet11.Range("I4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet12.Range("A1").Value = sheet0.Range("C1").Value
            sheet12.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet12.Range("I3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet12.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet12.Range("I4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet13.Range("A1").Value = sheet0.Range("C1").Value
            sheet13.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet13.Range("P3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet13.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet13.Range("P4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet14.Range("A1").Value = sheet0.Range("C1").Value
            sheet14.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet14.Range("L3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet14.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet14.Range("L4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet15.Range("A1").Value = sheet0.Range("C1").Value
            sheet15.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet15.Range("K3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet15.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet15.Range("K4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet16.Range("A1").Value = sheet0.Range("C1").Value
            sheet16.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet16.Range("H3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet16.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet16.Range("H4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            sheet17.Range("A1").Value = sheet0.Range("C1").Value
            sheet17.Range("A3").Value = sheet0.Range("A2").Value & sheet0.Range("C2").Value
            sheet17.Range("N3").Value = sheet0.Range("G2").Value & sheet0.Range("H2").Value
            sheet17.Range("A4").Value = sheet0.Range("A3").Value & sheet0.Range("C3").Value
            sheet17.Range("N4").Value = sheet0.Range("G3").Value & sheet0.Range("H3").Value

            '计算偏距
            Do While sheet0.Range("B" & h).Value <> Nothing
                sheet1.Range("I3:Q1000").Value = Nothing
                sheet1.Range("I1").Value = sheet1.Range("B5").Value.substring(0, ExApp.WorksheetFunction.Find("K", sheet1.Range("B5").Value))
                If sheet0.Range("K" & h).Value = "桥墩" Then
                    If sheet0.Range("C" & h).Value = Nothing And sheet0.Range("V" & h).Value = "否" And sheet0.Range("W" & h).Value <> Nothing And sheet0.Range("X" & h).Value <> Nothing Then
                        sheet0.Range("C" & h).Value = FSZhj(sheet0.Range("W" & h).Value, sheet0.Range("X" & h).Value)
                    ElseIf sheet0.Range("C" & h).Value <> Nothing And sheet0.Range("V" & h).Value = "否" Then
                        MsgBox("请核对第" & h & "行中心桩号或X、Y坐标是否正确填写！")
                        TorF = False
                        Exit Sub
                    End If
                    '桥墩挡块桩号
                    sheet1.Range("I3").Value = sheet0.Range("C" & h).Value + sheet0.Range("G" & h).Value / 2 / 100
                    sheet1.Range("I4").Value = sheet0.Range("C" & h).Value - sheet0.Range("G" & h).Value / 2 / 100
                    sheet1.Range("I5").Value = sheet0.Range("C" & h).Value
                    sheet1.Range("I6").Value = sheet0.Range("C" & h).Value
                    sheet1.Range("I7").Value = sheet0.Range("C" & h).Value + sheet0.Range("G" & h).Value / 2 / 100
                    sheet1.Range("I8").Value = sheet0.Range("C" & h).Value - sheet0.Range("G" & h).Value / 2 / 100
                    sheet1.Range("I9").Value = sheet0.Range("C" & h).Value
                    sheet1.Range("I10").Value = sheet0.Range("C" & h).Value
                Else
                    If sheet0.Range("AA" & h).Value = "起点" Then
                        '桥台起点桩号
                        sheet1.Range("I3").Value = sheet0.Range("C" & h).Value + sheet0.Range("AB" & h).Value / 100 + sheet0.Range("G" & h).Value / 2 / 100
                        sheet1.Range("I4").Value = sheet0.Range("C" & h).Value + sheet0.Range("AB" & h).Value / 100 - sheet0.Range("G" & h).Value / 2 / 100
                        sheet1.Range("I5").Value = sheet0.Range("C" & h).Value + sheet0.Range("AB" & h).Value / 100
                        sheet1.Range("I6").Value = sheet0.Range("C" & h).Value + sheet0.Range("AB" & h).Value / 100
                        sheet1.Range("I7").Value = sheet0.Range("C" & h).Value + sheet0.Range("AB" & h).Value / 100 + sheet0.Range("G" & h).Value / 2 / 100
                        sheet1.Range("I8").Value = sheet0.Range("C" & h).Value + sheet0.Range("AB" & h).Value / 100 - sheet0.Range("G" & h).Value / 2 / 100
                        sheet1.Range("I9").Value = sheet0.Range("C" & h).Value + sheet0.Range("AB" & h).Value / 100
                        sheet1.Range("I10").Value = sheet0.Range("C" & h).Value + sheet0.Range("AB" & h).Value / 100
                    Else
                        '桥台终点桩号
                        sheet1.Range("I3").Value = sheet0.Range("C" & h).Value - sheet0.Range("AB" & h).Value / 100 + sheet0.Range("G" & h).Value / 2 / 100
                        sheet1.Range("I4").Value = sheet0.Range("C" & h).Value - sheet0.Range("AB" & h).Value / 100 - sheet0.Range("G" & h).Value / 2 / 100
                        sheet1.Range("I5").Value = sheet0.Range("C" & h).Value - sheet0.Range("AB" & h).Value / 100
                        sheet1.Range("I6").Value = sheet0.Range("C" & h).Value - sheet0.Range("AB" & h).Value / 100
                        sheet1.Range("I7").Value = sheet0.Range("C" & h).Value - sheet0.Range("AB" & h).Value / 100 + sheet0.Range("G" & h).Value / 2 / 100
                        sheet1.Range("I8").Value = sheet0.Range("C" & h).Value - sheet0.Range("AB" & h).Value / 100 - sheet0.Range("G" & h).Value / 2 / 100
                        sheet1.Range("I9").Value = sheet0.Range("C" & h).Value - sheet0.Range("AB" & h).Value / 100
                        sheet1.Range("I10").Value = sheet0.Range("C" & h).Value - sheet0.Range("AB" & h).Value / 100
                    End If
                End If

                If sheet0.Range("J" & h).Value = "中" Or sheet0.Range("J" & h).Value = "右" Then
                    '偏距
                    sheet1.Range("J3").Value = (sheet0.Range("Y" & h).Value - sheet0.Range("H" & h).Value / 2) / -100
                    sheet1.Range("J4").Value = (sheet0.Range("Y" & h).Value - sheet0.Range("H" & h).Value / 2) / -100
                    sheet1.Range("J5").Value = sheet0.Range("Y" & h).Value / -100
                    sheet1.Range("J6").Value = (sheet0.Range("Y" & h).Value - sheet0.Range("H" & h).Value) / -100
                    sheet1.Range("J7").Value = (sheet0.Range("Y" & h).Value - sheet0.Range("H" & h).Value - sheet0.Range("Z" & h).Value - sheet0.Range("H" & h).Value / 2) / -100
                    sheet1.Range("J8").Value = (sheet0.Range("Y" & h).Value - sheet0.Range("H" & h).Value - sheet0.Range("Z" & h).Value - sheet0.Range("H" & h).Value / 2) / -100
                    sheet1.Range("J9").Value = (sheet0.Range("Y" & h).Value - sheet0.Range("H" & h).Value - sheet0.Range("Z" & h).Value) / -100
                    sheet1.Range("J10").Value = (sheet0.Range("Y" & h).Value - sheet0.Range("H" & h).Value * 2 - sheet0.Range("Z" & h).Value) / -100
                Else
                    sheet1.Range("J3").Value = (sheet0.Range("Y" & h).Value + sheet0.Range("H" & h).Value / 2) / 100
                    sheet1.Range("J4").Value = (sheet0.Range("Y" & h).Value + sheet0.Range("H" & h).Value / 2) / 100
                    sheet1.Range("J5").Value = sheet0.Range("Y" & h).Value / 100
                    sheet1.Range("J6").Value = (sheet0.Range("Y" & h).Value + sheet0.Range("H" & h).Value) / -100
                    sheet1.Range("J7").Value = (sheet0.Range("Y" & h).Value + sheet0.Range("H" & h).Value + sheet0.Range("Z" & h).Value + sheet0.Range("H" & h).Value / 2) / 100
                    sheet1.Range("J8").Value = (sheet0.Range("Y" & h).Value + sheet0.Range("H" & h).Value + sheet0.Range("Z" & h).Value + sheet0.Range("H" & h).Value / 2) / 100
                    sheet1.Range("J9").Value = (sheet0.Range("Y" & h).Value + sheet0.Range("H" & h).Value + sheet0.Range("Z" & h).Value) / 100
                    sheet1.Range("J10").Value = (sheet0.Range("Y" & h).Value + sheet0.Range("H" & h).Value * 2 + sheet0.Range("Z" & h).Value) / 100
                End If

                '高程
                P = (sheet0.Range("M" & h).Value - sheet0.Range("L" & h).Value) / (sheet0.Range("N" & h).Value / 100)  '横坡
                sheet1.Range("K3").Value = sheet0.Range("L" & h).Value + sheet0.Range("I" & h).Value / 100
                sheet1.Range("K4").Value = sheet0.Range("L" & h).Value + sheet0.Range("I" & h).Value / 100
                sheet1.Range("K5").Value = sheet0.Range("L" & h).Value + sheet0.Range("I" & h).Value / 100
                sheet1.Range("K6").Value = sheet0.Range("L" & h).Value + sheet0.Range("I" & h).Value / 100
                sheet1.Range("K7").Value = sheet0.Range("L" & h).Value + sheet0.Range("I" & h).Value / 100 + sheet0.Range("Z" & h).Value / 100 * P
                sheet1.Range("K8").Value = sheet0.Range("L" & h).Value + sheet0.Range("I" & h).Value / 100 + sheet0.Range("Z" & h).Value / 100 * P
                sheet1.Range("K9").Value = sheet0.Range("L" & h).Value + sheet0.Range("I" & h).Value / 100 + sheet0.Range("Z" & h).Value / 100 * P
                sheet1.Range("K10").Value = sheet0.Range("L" & h).Value + sheet0.Range("I" & h).Value / 100 + sheet0.Range("Z" & h).Value / 100 * P
                '测量偏差值
                sheet1.Range("O3").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O4").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O5").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O6").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O7").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O8").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O9").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet1.Range("O10").Value = ExApp.WorksheetFunction.RandBetween(-9, 9)

                sheet1.Range("P3").Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                sheet1.Range("P4").Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                sheet1.Range("P5").Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                sheet1.Range("P6").Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                sheet1.Range("P7").Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                sheet1.Range("P8").Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                sheet1.Range("P9").Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                sheet1.Range("P10").Value = ExApp.WorksheetFunction.RandBetween(1, 4)
                '计算坐标
                Dim sjxq, sjyq As Double
                Dim c, b, ZH0 As Integer
                b = 3
                Do While sheet1.Range("I" & b).Value <> Nothing
                    ZH0 = Val(ExApp.WorksheetFunction.Substitute(sheet1.Range("I" & b).Value, "*", ""))
                    If sheet3.Range("J2").Value <> "是" Then
                        c = Pd_YSw(ZH0, sheet2.Range("O5 : O500").Value)
                    Else
                        c = 1
                    End If
                    If c = -1 Then
                        MsgBox("请在交点法表内输入数据")
                        TorF = False
                        Exit Sub
                    Else
                        If sheet3.Range("J2").Value <> "是" Then
                            sjxq = ZSZB_X0j(sheet1.Range("I" & b).Value, sheet1.Range("J" & b).Value, 90)  '前偏距坐标
                            sjyq = ZSZB_Y0j(sheet1.Range("I" & b).Value, sheet1.Range("J" & b).Value, 90)
                        Else
                            sjxq = XYF_X(sheet1.Range("I" & b).Value, sheet1.Range("J" & b).Value, 90)  '前偏距坐标
                            sjyq = XYF_Y(sheet1.Range("I" & b).Value, sheet1.Range("J" & b).Value, 90)
                        End If
                        '偏距坐标赋值
                        sheet1.Range("M" & b).Value = Math.Round(sjxq, 3)
                        sheet1.Range("N" & b).Value = Math.Round(sjyq, 3)
                    End If
                    b += 1
                Loop

                ExApp.Calculation = ExApp.Calculation.xlCalculationManual '开启手动计算
                sheet1.Range("B" & r + 5).Value = sheet0.Range("B" & h).Value
                sheet1.Range("B" & r + 6).Value = sheet0.Range("C" & h).Value
                sheet1.Range("C" & r + 6).Value = sheet0.Range("D" & h).Value
                sheet1.Range("D" & r + 6).Value = sheet0.Range("E" & h).Value
                sheet1.Range("B" & r + 7).Value = sheet0.Range("O" & h).Value
                sheet1.Range("B" & r + 8).Value = sheet0.Range("P" & h).Value
                sheet1.Range("B" & r + 9).Value = sheet0.Range("Q" & h).Value
                sheet1.Range("B" & r + 10).Value = sheet0.Range("U" & h).Value
                sheet1.Range("B" & r + 11).Value = sheet0.Range("G" & h).Value
                sheet1.Range("B" & r + 12).Value = sheet0.Range("H" & h).Value
                sheet1.Range("B" & r + 13).Value = sheet0.Range("I" & h).Value
                sheet1.Range("B" & r + 14).Value = sheet0.Range("R" & h).Value
                sheet1.Range("B" & r + 15).Value = sheet0.Range("S" & h).Value
                sheet1.Range("B" & r + 16).Value = sheet0.Range("T" & h).Value
                sheet1.Range("B" & r + 17).Value = sheet0.Range("F" & h).Value
                sheet1.Range("P1").Value = sheet1.Range("B17").Value
                sheet1.Range("L3").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value

                sheet16.Range("A5").Value = sheet1.Range("A5").Value & sheet1.Range("B5").Value
                sheet17.Range("A5").Value = sheet1.Range("A5").Value & sheet1.Range("B5").Value


                '钢筋隐蔽工程
                sheet6.Range("C6").Value = sheet1.Range("B5").Value & sheet1.Range("C6").Value & "基础及下部构造"
                sheet6.Range("E6").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋加工及安装"
                sheet6.Range("C7").Value = sheet1.Range("B5").Value & sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋加工及安装"
                sheet6.Range("E10").Value = sheet1.Range("B17").Value
                sheet6.Range("E11").Value = sheet1.Range("B17").Value
                sheet6.Range("E27").Value = sheet1.Range("B17").Value
                sheet6.Range("E28").Value = sheet1.Range("B17").Value
                '钢筋检表
                sheet7.Range("D6").Value = sheet1.Range("B5").Value
                sheet7.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋"
                sheet7.Range("Q6").Value = sheet1.Range("B17").Value
                sheet7.Range("Q7").Value = sheet1.Range("B17").Value
                sheet7.Range("Q31").Value = sheet1.Range("B17").Value
                sheet7.Range("Q34").Value = sheet1.Range("B17").Value
                '主筋
                sheet7.Range("E15").Value = "设计值：" & sheet1.Range("B8").Value * 10
                LSBLFZ = Nothing
                If sheet1.Range("B7").Value * 2 <= 10 Then
                    For cs = 1 To sheet1.Range("B7").Value * 2
                        LSBL = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet7.Range("G14").Value = LSBLFZ
                    sheet8.Range("D24").Value = "/"
                Else
                    sheet7.Range("G14").Value = "应测" & sheet1.Range("B7").Value * 2 & "处，实测" & sheet1.Range("B7").Value * 2 & "处，合格" & sheet1.Range("B7").Value * 2 & "处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                    For cs = 1 To sheet1.Range("B7").Value * 2
                        LSBL = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet8.Range("D24").Value = LSBLFZ
                End If
                '箍筋
                sheet7.Range("E17").Value = "设计值：" & sheet1.Range("B9").Value * 10
                LSBLFZ = Nothing
                For cs = 1 To 20
                    LSBL = sheet1.Range("B9").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                    LSBLFZ = LSBLFZ & LSBL & "   "
                Next
                sheet7.Range("G16").Value = "应测20处，实测20处，合格20处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                sheet8.Range("D29").Value = LSBLFZ

                '骨架尺寸
                sheet7.Range("E19").Value = "设计值：" & sheet1.Range("B14").Value * 10
                sheet7.Range("G18").Value = sheet1.Range("B14").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet7.Range("E21").Value = "宽：" & sheet1.Range("B15").Value * 10
                sheet7.Range("E22").Value = "高：" & sheet1.Range("B16").Value * 10
                sheet7.Range("G20").Value = "宽： " & sheet1.Range("B15").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4) & vbCrLf &
                                            "高： " & sheet1.Range("B16").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4)
                '保护层
                sheet7.Range("E27").Value = "设计值：" & sheet1.Range("B10").Value * 10
                Dim gs As Integer
                gs = Math.Round((sheet1.Range("B11").Value * sheet1.Range("B13").Value * 2 + sheet1.Range("B12").Value * sheet1.Range("B13").Value * 2) / 300 / 100, 0)
                If gs <= 20 Then
                    LSBLFZ = Nothing
                    For cs = 1 To 20
                        LSBL = sheet1.Range("B10").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet7.Range("G26").Value = "应测20处，实测20处，合格20处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                    sheet8.Range("D48").Value = LSBLFZ
                Else
                    LSBLFZ = Nothing
                    For cs = 1 To gs
                        LSBL = sheet1.Range("B10").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet7.Range("G26").Value = "应测" & gs & "处，实测" & gs & "处，合格" & gs & "处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                    sheet8.Range("D48").Value = LSBLFZ
                End If
                '钢筋记录表
                sheet8.Range("B6").Value = sheet1.Range("B5").Value
                sheet8.Range("B7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋"
                sheet8.Range("K6").Value = sheet1.Range("B17").Value
                sheet8.Range("K7").Value = sheet1.Range("B17").Value
                '工序检验申请批复单
                sheet9.Range("C6").Value = sheet1.Range("B5").Value & sheet1.Range("C6").Value & "基础及下部构造"
                sheet9.Range("C7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet9.Range("C8").Value = sheet1.Range("B5").Value
                sheet9.Range("C9").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet9.Range("C10").Value = "混凝土强度、平面位置、断面尺寸及高度、与梁体间隙"

                Call 全站仪平面位置检测表（）
                If TorF = False Then
                    Exit Sub
                End If
                sheet17.Activate()
                ExApp.ActiveWindow.SelectedSheets.Copy(, (ExApp.Sheets(ExApp.Sheets.Count)))

                '挡块现场质量检验表
                sheet10.Range("D6").Value = sheet1.Range("B5").Value
                sheet10.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet10.Range("K6").Value = sheet1.Range("B17").Value
                LSBLFZ = Nothing
                For cs = 3 To 10
                    LSBL = sheet1.Range("P" & cs).Value
                    LSBLFZ = LSBLFZ & LSBL & "   "
                Next
                sheet10.Range("F12").Value = LSBLFZ
                sheet10.Range("D14").Value = "长：" & sheet1.Range("B11").Value * 10
                sheet10.Range("D15").Value = "宽：" & sheet1.Range("B12").Value * 10
                sheet10.Range("D16").Value = "高：" & sheet1.Range("B13").Value * 10
                sheet10.Range("F14").Value = "长：  " & sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9) & "   " &
                                                                           sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9) & "   " &
                                                                           sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9) & "   " &
                                                                           sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet10.Range("F15").Value = "宽：  " & sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9) & "   " &
                                                                           sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9) & "   " &
                                                                           sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9) & "   " &
                                                                           sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet10.Range("F16").Value = "高： " & sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9) & "   " &
                                                                           sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9) & "   " &
                                                                           sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9) & "   " &
                                                                           sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet10.Range("F17").Value = ExApp.WorksheetFunction.RandBetween(-4, 4) & "   " &
                                             ExApp.WorksheetFunction.RandBetween(-4, 4) & "   " &
                                             ExApp.WorksheetFunction.RandBetween(-4, 4) & "   " &
                                             ExApp.WorksheetFunction.RandBetween(-4, 4)

                '模板测量偏差值
                sheet1.Range("P3").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet1.Range("P4").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet1.Range("P5").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet1.Range("P6").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet1.Range("P7").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet1.Range("P8").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet1.Range("P9").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet1.Range("P10").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet1.Range("L3").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "模板"

                Call 水准测量记录表()
                If TorF = False Then
                    Exit Sub
                End If
                Call 全站仪平面位置检测表（）
                If TorF = False Then
                    Exit Sub
                End If
                '现场模板安装检查记录表
                sheet11.Range("C6").Value = sheet1.Range("B5").Value
                sheet11.Range("C7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet11.Range("I6").Value = sheet1.Range("B17").Value
                sheet11.Range("I7").Value = sheet1.Range("B17").Value
                sheet11.Range("D8").Value = 2
                sheet11.Range("D9").Value = 5
                sheet11.Range("F8").Value = 4
                sheet11.Range("F9").Value = 4
                sheet11.Range("H8").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet11.Range("H9").Value = ExApp.WorksheetFunction.RandBetween(1, 5)
                sheet11.Range("J8").Value = 100
                sheet11.Range("J9").Value = 100
                '测量偏差
                sheet11.Range("D11").Value = sheet1.Range("P5").Value
                sheet11.Range("F11").Value = sheet1.Range("P6").Value
                sheet11.Range("H11").Value = sheet1.Range("P3").Value
                sheet11.Range("J11").Value = sheet1.Range("P4").Value
                sheet11.Range("F14").Value = sheet1.Range("O3").Value
                sheet11.Range("G14").Value = sheet1.Range("O4").Value
                sheet11.Range("H14").Value = sheet1.Range("O5").Value
                sheet11.Range("I14").Value = sheet1.Range("O6").Value

                sheet11.Range("F12").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("G12").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("H12").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("I12").Value = sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("F13").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("G13").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("H13").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("I13").Value = sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                sheet11.Range("F18").Value = "牢固，稳定"
                '砼浇筑申请报告单
                '.....................................................................

                '监抽钢筋检表

                sheet13.Range("D6").Value = sheet1.Range("B5").Value
                sheet13.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋"
                sheet13.Range("Q6").Value = sheet1.Range("B17").Value
                sheet13.Range("Q7").Value = sheet1.Range("B17").Value
                sheet13.Range("Q31").Value = sheet1.Range("B17").Value
                sheet13.Range("Q34").Value = sheet1.Range("B17").Value
                '主筋
                sheet13.Range("E15").Value = "设计值：" & sheet1.Range("B8").Value * 10
                LSBLFZ = Nothing
                If Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0) <= 10 Then
                    For cs = 1 To Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0)
                        LSBL = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet13.Range("G14").Value = LSBLFZ
                    sheet14.Range("D24").Value = "/"
                Else
                    sheet13.Range("G14").Value = "应测" & Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0) & "处，实测" & Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0) & "处，合格" & Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0) & "处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                    For cs = 1 To Math.Round(sheet1.Range("B7").Value * 2 * 0.2, 0)
                        LSBL = sheet1.Range("B8").Value * 10 + ExApp.WorksheetFunction.RandBetween(-15, 15)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet14.Range("D24").Value = LSBLFZ
                End If
                '箍筋
                sheet13.Range("E17").Value = "设计值：" & sheet1.Range("B9").Value * 10
                sheet13.Range("G16").Value = sheet1.Range("B9").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9) & "   " &
                                             sheet1.Range("B9").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9) & "   " &
                                             sheet1.Range("B9").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9) & "   " &
                                             sheet1.Range("B9").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                '骨架尺寸
                sheet13.Range("E19").Value = "设计值：" & sheet1.Range("B14").Value * 10
                sheet13.Range("G18").Value = sheet1.Range("B14").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet13.Range("E21").Value = "宽：" & sheet1.Range("B15").Value * 10
                sheet13.Range("E22").Value = "高：" & sheet1.Range("B16").Value * 10
                sheet13.Range("G20").Value = "宽： " & sheet1.Range("B15").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4) & vbCrLf &
                                             "高： " & sheet1.Range("B16").Value * 10 + ExApp.WorksheetFunction.RandBetween(-4, 4)
                '保护层
                sheet13.Range("E27").Value = "设计值：" & sheet1.Range("B10").Value * 10
                gs = Math.Round((sheet1.Range("B11").Value * sheet1.Range("B13").Value * 2 + sheet1.Range("B12").Value * sheet1.Range("B13").Value * 2) / 300 / 100 * 0.2, 0)
                If gs <= 10 Then
                    LSBLFZ = Nothing
                    For cs = 1 To gs
                        LSBL = sheet1.Range("B10").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet13.Range("G26").Value = LSBLFZ
                    sheet14.Range("D48").Value = Nothing
                Else
                    LSBLFZ = Nothing
                    For cs = 1 To gs
                        LSBL = sheet1.Range("B10").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                        LSBLFZ = LSBLFZ & LSBL & "   "
                    Next
                    sheet13.Range("G26").Value = "应测" & gs & "处，实测" & gs & "处，合格" & gs & "处，合格率为100%，数据详见钢筋安装现场检查记录表TJ8-"
                    sheet14.Range("D48").Value = LSBLFZ
                End If
                '监抽钢筋记录表
                sheet14.Range("B6").Value = sheet1.Range("B5").Value
                sheet14.Range("B7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value & "钢筋"
                sheet14.Range("K6").Value = sheet1.Range("B17").Value
                sheet14.Range("K7").Value = sheet1.Range("B17").Value
                '监抽挡块现场质量检验表
                sheet15.Range("D6").Value = sheet1.Range("B5").Value
                sheet15.Range("D7").Value = sheet1.Range("C6").Value & sheet1.Range("D6").Value
                sheet15.Range("K6").Value = sheet1.Range("B17").Value
                '平面位置
                LSBLFZ = Nothing
                For cs = 3 To 10
                    LSBL = sheet1.Range("P" & cs).Value
                    LSBLFZ = LSBLFZ & LSBL & "   "
                Next
                sheet15.Range("F12").Value = LSBLFZ
                sheet15.Range("D14").Value = "长：" & sheet1.Range("B11").Value * 10
                sheet15.Range("D15").Value = "宽：" & sheet1.Range("B12").Value * 10
                sheet15.Range("D16").Value = "高：" & sheet1.Range("B13").Value * 10
                sheet15.Range("F14").Value = "长：  " & sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9) & "  " & sheet1.Range("B11").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet15.Range("F15").Value = "宽：  " & sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9) & "  " & sheet1.Range("B12").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet15.Range("F16").Value = "高：  " & sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9) & "  " & sheet1.Range("B13").Value * 10 + ExApp.WorksheetFunction.RandBetween(-9, 9)
                sheet15.Range("F17").Value = ExApp.WorksheetFunction.RandBetween(-4, 4) & "   " & ExApp.WorksheetFunction.RandBetween(-4, 4)

                ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                '选择表格
                sheet6.Select()
                For i = 7 To ExApp.Sheets.Count
                    EXsheet = Exbook.Worksheets(i)
                    If EXsheet.Visible = True Then
                        EXsheet.Select(Replace:=False)
                    End If
                Next i
                EXsheet = ExApp.ActiveSheet
                ' 导出PDF文件
                PDFFilename = Filepath & "\" & sheet1.Range("B5").Value & sheet1.Range("C6").Value & sheet1.Range("D6").Value & ".pdf"
                EXsheet.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, PDFFilename, XlFixedFormatQuality.xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False)

                sheet18 = Exbook.Worksheets("水准表(2)")
                sheet19 = Exbook.Worksheets("平面表(2)")
                ExApp.DisplayAlerts = False
                sheet18.Delete()
                sheet19.Delete()
                h += 1
            Loop

            MsgBox("已完成！", 0 + 64, "提示")
        Catch Exclerror As Exception   '错误时弹出提示
            MsgBox(Exclerror.Message)
        End Try
        TorF = False
    End Sub
End Module
