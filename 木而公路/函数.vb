

Imports Microsoft.Office.Interop.Excel

Module 函数

    Public Const PI As Double = 3.14159265358979
    Public ExApp As Application
    Public Exbook As Workbook
    Public action(20) As String '保存账户，密码，部门,部门数据库
    Public store(50) As String '保存用户数据库，数据表
    Public filenamed As String '保存文件名
    Public projectname(50) As String '保存项目名（如桥梁，隧道）,工程部位名
    Public storebase As String '工程部位名数据表
    Public storeku As String '保存系统管理员数据库名
    Public TorF As Boolean '判断是否选择线路参数
    Public FileName1, FileName2 As String '保存本地模板文件路径名
    Public Filepath As String  '保存文件路径名


    Public Function ZH(ZH1)  '改写桩号

        sheet0 = Exbook.Worksheets("数据库") '数据库
        Dim Qz, st
        Dim a1, a2, a3, a4, a5, a6, a7
        Qz = sheet0.Range("F2").Value '前缀
        st = Mid(ZH1, 1, 1)
        ZH1 = Val(ExApp.WorksheetFunction.Substitute(ZH1, "*", ""))
        If Qz = "" Or Qz = "K" Or Qz = "k" Then
            ZH = Format(ZH1, "K0+000.000") '"!AK0+000.000"'''"!BK0+000.000"
        Else
            If ZH1 < 0 Then
                Qz = "-" & Qz
            End If
            ZH1 = Math.Abs(Math.Round(ZH1, 3))
            a1 = Int(ZH1 / 1000) '千
            a2 = Int((ZH1 - a1 * 1000) / 100) '百
            a3 = Int((ZH1 - Int(ZH1 / 100) * 100) / 10) '十
            a4 = Int(ZH1 - Int(ZH1 / 10) * 10) '个
            a5 = Int(Math.Round(ZH1 - Int(ZH1), 3) * 10)   '小数1
            a6 = Int(Math.Round(ZH1 - Int(ZH1), 3) * 100) - a5 * 10 '小数2
            a7 = Int(Math.Round(ZH1 - Int(ZH1), 3) * 1000) - a5 * 100 - a6 * 10   '小数3
            ZH = Qz & a1 & "+" & a2 & a3 & a4 & "." & a5 & a6 & a7
        End If
        If st = "*" Then
            ZH = st & ZH
        End If
        'ZH = ZH1
    End Function

    Public Function ZJZH(ZH1)  '改写质检测量桩号

        sheet1 = Exbook.Worksheets("参数表") '参数表
        Dim Qz, st
        Dim a1, a2, a3, a4, a5, a6, a7
        Qz = sheet1.Range("I1").Value '前缀
        st = Mid(ZH1, 1, 1)
        ZH1 = Val(ExApp.WorksheetFunction.Substitute(ZH1, "*", ""))
        If Qz = "" Or Qz = "K" Or Qz = "k" Then
            ZJZH = Format(ZH1, "K0+000.000")
        Else
            If ZH1 < 0 Then
                Qz = "-" & Qz
            End If
            ZH1 = Math.Abs(Math.Round(ZH1, 3))
            a1 = Int(ZH1 / 1000) '千
            a2 = Int((ZH1 - a1 * 1000) / 100) '百
            a3 = Int((ZH1 - Int(ZH1 / 100) * 100) / 10) '十
            a4 = Int(ZH1 - Int(ZH1 / 10) * 10) '个
            a5 = Int(Math.Round(ZH1 - Int(ZH1), 3) * 10)   '小数1
            a6 = Int(Math.Round(ZH1 - Int(ZH1), 3) * 100) - a5 * 10 '小数2
            a7 = Int(Math.Round(ZH1 - Int(ZH1), 3) * 1000) - a5 * 100 - a6 * 10   '小数3
            ZJZH = Qz & a1 & "+" & a2 & a3 & a4 & "." & a5 & a6 & a7
        End If
        If st = "*" Then
            ZJZH = st & ZJZH
        End If

    End Function


    Public Function XSWS(sj) '保留小数位数
        XSWS = Format(sj, "0.00")
        'XSWS = sj
    End Function
    Public Function BP_B(Z, GC, b) '边坡宽度计算,GC为点到路面的高差，+填-挖，b左右幅判断
        Dim i, c_s, k_s, bp, pt
        c_s = 0
        k_s = 0
        pt = 0
        i = 4
        If GC >= 0 Then
            '填方
            While Math.Abs(GC) >= c_s And ExApp.Sheets("边坡参数").Range("c" & i).value <> Nothing
                c_s = c_s + ExApp.Sheets("边坡参数").Range("c" & i).value '高度
                k_s = k_s + ExApp.Sheets("边坡参数").Range("c" & i).value * ExApp.Sheets("边坡参数").Range("b" & i).value '宽度
                pt = pt + ExApp.Sheets("边坡参数").Range("d" & i).value '平台
                i = i + 1
            End While
            bp = k_s + pt + (Math.Abs(GC) - c_s) * ExApp.Sheets("边坡参数").Range("b" & i - 1).value
        Else
            '挖方
            While Math.Abs(GC) >= c_s And ExApp.Sheets("边坡参数").Range("h" & i).value <> Nothing
                c_s = c_s + ExApp.Sheets("边坡参数").Range("h" & i).value '高度
                k_s = k_s + ExApp.Sheets("边坡参数").Range("h" & i).value * ExApp.Sheets("边坡参数").Range("g" & i).value '宽度
                pt = pt + ExApp.Sheets("边坡参数").Range("i" & i).value '平台
                i = i + 1
            End While
            bp = k_s + pt + (Math.Abs(GC) - c_s) * ExApp.Sheets("边坡参数").Range("g" & i - 1).value
        End If
        If b < 0 Then
            BP_B = -bp + JK_ZY(Z, b)
        Else
            BP_B = bp + JK_ZY(Z, b)
        End If
    End Function
    Public Function BP_i(GC)   '填方边坡坡度计算,GC为点到路面的高差，+填-挖
        Dim i, c_s, pd
        i = 4
        c_s = 0
        pd = 0
        While ExApp.Sheets("边坡参数").Range("c" & i).value <> Nothing And pd = 0
            c_s = c_s + ExApp.Sheets("边坡参数").Range("c" & i).value
            If GC <= c_s Then
                BP_i = ExApp.Sheets("边坡参数").Range("b" & i).value
                pd = 1
            End If
            i = i + 1
        End While
        If ExApp.Sheets("边坡参数").Range("c" & i).value = Nothing And i > 4 Then
            BP_i = ExApp.Sheets("边坡参数").Range("b" & i - 1).value
        End If
    End Function
    Public Function FCpj_z(Z, k, H) '左侧筑面偏距,z桩号，k原地面点，h路面到填筑面高差
        Dim i, GC(0 To 328), Pj(0 To 328), m, pd, jdB(0 To 328), max0, min0
        Dim ming1, minp1, ming2, minp2, fpb
        Dim pdjg, mo, TH
        mo = 0
        FCpj_z = 32767
        pdjg = PJgc_z(Z, k) '坡脚高差
        i = 1
        m = 1
        max0 = -32767
        min0 = 32767
        While k(i) <> Nothing
            Pj(m) = k(i) '偏距
            GC(m) = k(i + 1) '高差，+填，-挖
            If GC(m) > max0 Then
                max0 = GC(m) '最大高高差
            End If
            If GC(m) < min0 Then
                min0 = GC(m) '最小高差
            End If
            i = i + 2
            m = m + 1
        End While
        If H > max0 Then
            H = max0 '填方厚度不能大于原地面
        End If
        '开始
        'if2222
        If m = 2 Or H <= pdjg Then
            FCpj_z = BP_B(Z, H, -1) '原地面为直线时直接放坡或者填土高度小于坡脚时
        Else
            '非边坡点
            i = 2
            pd = 0
            While Pj(i) <> Nothing And pd <= 2 '坡脚点
                i = 2
                pd = 0
                jdB(1) = 0
                While Pj(i) <> Nothing And pd <= 2
                    If (GC(i) - GC(i - 1)) <> 0 Then
                        TH = Pj(i - 1) + (Pj(i) - Pj(i - 1)) / (GC(i) - GC(i - 1)) * (H - GC(i - 1))
                    Else
                        TH = Pj(i - 1)
                    End If
                    If ((H <= GC(i - 1) And H >= GC(i)) Or (H >= GC(i - 1) And H <= GC(i))) And BP_B(Z, pdjg, -1) <= TH Then ' Then '
                        pd = pd + 1
                        mo = mo + 1
                        If (GC(i) - GC(i - 1)) <> 0 Then
                            jdB(pd) = Pj(i - 1) + (Pj(i) - Pj(i - 1)) / (GC(i) - GC(i - 1)) * (H - GC(i - 1))
                        Else
                            jdB(pd) = Pj(i - 1)
                        End If
                    End If
                    i = i + 1
                End While
            End While
            If mo <> 0 Then
                FCpj_z = jdB(1) '(BP_B(z, pdjg, -1) >= Pj(1) + (Pj(2) - Pj(1)) / (GC(2) - GC(1)) * (h - GC(1))) 'jdB(1)
            End If
            If FCpj_z < BP_B(Z, pdjg, -1) Then
                FCpj_z = BP_B(Z, pdjg, -1)
            Else
                If FCpj_z > BP_B(Z, PJgc_y(Z, k), 1) Then
                    FCpj_z = BP_B(Z, PJgc_y(Z, k), 1)
                End If
            End If
            'if2222
        End If
    End Function
    Public Function FCpj_y(Z, k, H) '右侧填筑面偏距,z桩号，k原地面点，h路面到填筑面高差
        Dim i, GC(0 To 328), Pj(0 To 328), m, pd, jdB(0 To 328), max0, min0
        Dim ming1, minp1, ming2, minp2, fpb
        Dim pdjg, mo, TH
        mo = 0
        FCpj_y = 32767
        pdjg = PJgc_y(Z, k) '坡脚高差
        i = 1
        m = 1
        max0 = -32767
        min0 = 32767
        While k(i) <> Nothing
            Pj(m) = k(i) '偏距
            GC(m) = k(i + 1) '高差，+填，-挖
            If GC(m) > max0 Then
                max0 = GC(m) '最大高高差
            End If
            If GC(m) < min0 Then
                min0 = GC(m) '最小高差
            End If
            i = i + 2
            m = m + 1
        End While
        If H > max0 Then
            H = max0 '填方厚度不能大于原地面
        End If
        '开始
        'if2222
        If m = 2 Or H <= pdjg Then
            FCpj_y = BP_B(Z, H, 1) '原地面为直线时直接放坡或者填土高度小于坡脚时
        Else
            '非边坡点
            i = 2
            pd = 0
            While Pj(i) <> Nothing  '坡脚点
                i = 2
                pd = 0
                jdB(1) = 0
                While Pj(i) <> Nothing
                    If (GC(i) - GC(i - 1)) <> 0 Then
                        TH = Pj(i - 1) + (Pj(i) - Pj(i - 1)) / (GC(i) - GC(i - 1)) * (H - GC(i - 1))
                    Else
                        TH = Pj(i - 1)
                    End If

                    If ((H <= GC(i - 1) And H >= GC(i)) Or (H >= GC(i - 1) And H <= GC(i))) And BP_B(Z, pdjg, 1) >= TH Then '
                        pd = pd + 1
                        If (GC(i) - GC(i - 1)) <> 0 Then
                            jdB(pd) = Pj(i - 1) + (Pj(i) - Pj(i - 1)) / (GC(i) - GC(i - 1)) * (H - GC(i - 1))
                        Else
                            jdB(pd) = Pj(i - 1)
                        End If
                        mo = pd
                    End If
                    i = i + 1
                End While
            End While
            If mo <> 0 Then
                FCpj_y = jdB(mo)
            End If
            'if2222
        End If
        If FCpj_y > BP_B(Z, pdjg, 1) Then
            FCpj_y = BP_B(Z, pdjg, 1)
        Else
            If FCpj_y < BP_B(Z, PJgc_z(Z, k), -1) Then
                FCpj_y = BP_B(Z, PJgc_z(Z, k), -1)
            End If
        End If
    End Function
    Public Function PJgc_z(Z, k)  '左侧坡脚点高差
        Dim i, GC(0 To 328), Pj(0 To 328), m, pd, jdB(0 To 328)
        Dim ming1, minp1, ming2, minp2, fpb, pjdp, pjdg, pdg, pjdg0
        Dim pd0, pd1
        PJgc_z = -1
        pd0 = 1
        i = 1
        m = 1
        While k(i) <> Nothing
            Pj(m) = k(i) '偏距
            GC(m) = k(i + 1) '高差，+填，-挖
            i = i + 2
            m = m + 1
        End While
        i = 2
        pd = 1
        pd0 = 0
        While Pj(i) <> Nothing And pd0 < 327677 And pd1 = 0 '坡脚点
            If GC(i) > GC(i - 1) Then
                ming1 = GC(i - 1) '小高差
                minp1 = Pj(i - 1) '偏距
                ming2 = GC(i) '大高差
                minp2 = Pj(i) '偏距
            Else
                ming1 = GC(i) '小高差
                minp1 = Pj(i) '偏距
                ming2 = GC(i - 1) '大高差
                minp2 = Pj(i - 1) '偏距
            End If
            pjdg0 = ming1  '边坡点高差
            fpb = BP_B(Z, pjdg0, -1) '坡脚偏距
            pjdp = minp1 - fpb '坡脚点平距差
            If (minp2 - minp1) <> 0 Then
                pjdg = (ming2 - ming1) / (minp2 - minp1) * pjdp  '边坡点差高差
            Else
                pjdg = (ming2 - ming1)
            End If
            pdg = pjdg - (ming1 - pjdg0) '计算高差与坡脚差
            If Math.Abs(pdg) <= 0.1 And pjdg0 <= ming2 And pd0 < 327677 And (minp1 < fpb And fpb < minp2 Or minp2 < fpb And fpb < minp1) Then
                PJgc_z = ming1 + pjdg
                pd1 = 1
            Else
                While Math.Abs(pdg) > 0.1 And pjdg0 <= ming2 And pd0 < 32767 And (minp1 < fpb And fpb < minp2 Or minp2 < fpb And fpb < minp1)
                    pjdg0 = pjdg0 + 0.05 '边坡点高差
                    fpb = BP_B(Z, pjdg0, -1)
                    pjdp = fpb - minp1 '坡脚点平距差
                    If (minp2 - minp1) <> 0 Then
                        pjdg = (ming2 - ming1) / (minp2 - minp1) * pjdp  '边坡点差高差
                    Else
                        pjdg = (ming2 - ming1)
                    End If
                    pdg = pjdg - (pjdg0 - ming1) '计算高差与坡脚差
                    pd0 = pd0 + 1 '中止死循环
                    If Math.Abs(pdg) <= 0.1 Then
                        PJgc_z = ming1 + pjdg
                        pd1 = 1
                    End If
                End While
            End If '
            pd0 = pd0 + 1 '中止死循环
            i = i + 1
        End While
        If PJgc_z = -1 Then '当地面线不够长时
            PJgc_z = GC(1)
        End If
        'PJgc_z = fpb
    End Function
    Public Function PJgc_y(Z, k)  '左侧坡脚点高差
        Dim i, GC(0 To 328), Pj(0 To 328), m, pd, jdB(0 To 328)
        Dim ming1, minp1, ming2, minp2, fpb, pjdp, pjdg, pdg, pjdg0
        Dim pd0, pd1
        PJgc_y = -1
        pd0 = 1
        i = 1
        m = 1
        While k(i) <> Nothing
            Pj(m) = k(i) '偏距
            GC(m) = k(i + 1) '高差，+填，-挖
            i = i + 2
            m = m + 1
        End While
        i = 2
        pd = 1
        pd1 = 0
        While Pj(i) <> Nothing And pd0 < 327677 And pd1 = 0 '坡脚点
            If GC(i) > GC(i - 1) Then
                ming1 = GC(i - 1) '小高差
                minp1 = Pj(i - 1) '偏距
                ming2 = GC(i) '大高差
                minp2 = Pj(i) '偏距
            Else
                ming1 = GC(i) '小高差
                minp1 = Pj(i) '偏距
                ming2 = GC(i - 1) '大高差
                minp2 = Pj(i - 1) '偏距
            End If
            pjdg0 = ming1  '边坡点高差
            fpb = BP_B(Z, pjdg0, 1) '坡脚偏距
            pjdp = minp1 - fpb '坡脚点平距差
            If (minp2 - minp1) <> 0 Then
                pjdg = (ming2 - ming1) / (minp2 - minp1) * pjdp  '边坡点差高差
            Else
                pjdg = (ming2 - ming1)
            End If
            pdg = pjdg - (ming1 - pjdg0) '计算高差与坡脚差
            If Math.Abs(pdg) <= 0.1 And pjdg0 <= ming2 And pd0 < 327677 And (minp1 < fpb And fpb < minp2 Or minp2 < fpb And fpb < minp1) Then
                PJgc_y = ming1 + pjdg
                pd1 = 1
            Else
                While Math.Abs(pdg) > 0.1 And pjdg0 <= ming2 And pd0 < 32767 And (minp1 < fpb And fpb < minp2 Or minp2 < fpb And fpb < minp1)
                    pjdg0 = pjdg0 + 0.05 '边坡点高差
                    fpb = BP_B(Z, pjdg0, 1)
                    pjdp = fpb - minp1 '坡脚点平距差
                    If (minp2 - minp1) <> 0 Then
                        pjdg = (ming2 - ming1) / (minp2 - minp1) * pjdp  '边坡点差高差
                    Else
                        pjdg = (ming2 - ming1)
                    End If
                    pdg = pjdg - (pjdg0 - ming1) '计算高差与坡脚差
                    pd0 = pd0 + 1 '中止死循环
                    If Math.Abs(pdg) <= 0.1 Then
                        PJgc_y = ming1 + pjdg
                        pd1 = 1
                    End If
                End While
            End If
            pd0 = pd0 + 1 '中止死循环
            i = i + 1
        End While
        If PJgc_y = -1 Then '当地面线不够长时
            PJgc_y = GC(m - 1)
        End If
        'PJgc_y = fpb
    End Function

    Public Function tf_h(Z, k) '填方高度，k原地面数据
        Dim i, m, zg, yg, zp, yp, max0
        zg = PJgc_z(Z, k) '左坡脚高
        yg = PJgc_y(Z, k) '右坡脚gap
        zp = BP_B(Z, zg, -1)
        yp = BP_B(Z, yg, 1)
        If zg > yg Then
            max0 = zg
        Else
            max0 = yg
        End If
        i = 1
        While k(i) <> Nothing
            If k(i) > zp And k(i) < yp Then
                If k(i + 1) > max0 Then
                    max0 = k(i + 1)
                End If
            End If
            i = i + 2
        End While
        tf_h = Int(max0 * 10) / 10
    End Function
    Public Function FCS0_t(H0) '土方分层总数,H0填土的厚度,扣结构层
        Dim i, slc, xlc, sld, pd, lch, sldh, xldh, lcs, slds
        H0 = Val(H0)
        xlc = 0.8 '路床
        sld = 0.7 '上路堤
        lch = ExApp.Sheets("路基填筑").Range("K2").value '路床，0.2厚
        sldh = ExApp.Sheets("路基填筑").Range("K3").value '上路堤，0.233厚
        xldh = ExApp.Sheets("路基填筑").Range("N2").value  '下路堤，0.25厚
        lcs = Int(xlc / lch) '路床层数，0.2厚
        slds = Int(sld / sldh) '上路堤层数，0.233厚
        'If H0 <= slc Then '上路床1层，0.3厚
        'FCS0_s = 1
        'Else
        If H0 <= xlc Then
            pd = Math.Round((H0) / lch, 0) '路床，0.2厚
            FCS0_t = pd
        Else
            If H0 <= sld Then
                pd = Math.Round((H0 - xlc) / sldh, 0) '上路堤，0.233厚
                FCS0_t = lcs + pd
            Else
                pd = Math.Round((H0 - xlc - sld) / xldh, 0) '下路堤，0.25厚
                FCS0_t = (lcs + slds) + pd
            End If
        End If
        'End If
    End Function
    Public Function FCS0_s(H0) '石方分层总数,H0填土的厚度,扣结构层
        Dim i, slc, xlc, sld, pd, lch, sldh, xldh, lcs, slds
        H0 = Val(H0)
        'slc = 0.3
        xlc = 0.8
        sld = 0.7
        lch = ExApp.Sheets("路基填筑").Range("q2").value '路床，0.2厚
        sldh = ExApp.Sheets("路基填筑").Range("q3").value  '上路堤，0.233厚
        xldh = ExApp.Sheets("路基填筑").Range("t2").value  '下路堤，0.25厚
        lcs = Int(xlc / lch) '路床层数，0.2厚
        slds = Int(sld / sldh) '上路堤层数，0.233厚
        'If H0 <= slc Then '上路床1层，0.3厚
        'FCS0_s = 1
        'Else
        If H0 <= xlc Then
            pd = Math.Round((H0) / lch, 0) '路床，0.2厚
            FCS0_s = pd
        Else
            If H0 <= sld Then
                pd = Math.Round((H0 - xlc) / sldh, 0)  '上路堤，0.35厚
                FCS0_s = lcs + pd
            Else
                pd = Math.Round((H0 - sld - xlc) / xldh, 0) '下路堤，0.5厚
                FCS0_s = lcs + slds + pd
            End If
        End If
        'End If
    End Function
    Public Function FCzh_q0(k, H)  '起点桩号，k原地面点，h路面到填筑面高差
        Dim i, ZH(0 To 328), GC(0 To 328), m, pd, jdz(0 To 328), max0
        i = 1
        m = 1
        max0 = 0
        While k(i) <> Nothing
            ZH(m) = k(i) '桩号
            GC(m) = k(i + 1) '高差
            If GC(m) > max0 Then
                max0 = GC(m) '最大高差
            End If
            i = i + 2
            m = m + 1
        End While
        If H > max0 Then
            H = max0
        End If
        If m = 2 Then
            FCzh_q0 = ZH(1)
        Else
            i = 2
            pd = 0
            jdz(1) = 0
            While ZH(i) <> Nothing And pd <= 2
                If (H <= GC(i - 1) And H >= GC(i)) Or (H >= GC(i - 1) And H <= GC(i)) Then '
                    pd = pd + 1
                    If (GC(i) - GC(i - 1)) <> 0 Then
                        jdz(pd) = ZH(i - 1) + (ZH(i) - ZH(i - 1)) / (GC(i) - GC(i - 1)) * (H - GC(i - 1))
                    Else
                        jdz(pd) = ZH(i - 1)
                    End If
                End If
                i = i + 1
            End While
            If H <= GC(1) Then
                FCzh_q0 = ZH(1)
            Else
                FCzh_q0 = jdz(1)
            End If
        End If
    End Function
    Public Function FCzh_z0(k, H)  '终点桩号，k原地面点，h路面到填筑面高差
        Dim i, ZH(0 To 328), GC(0 To 328), m, pd, jdz(0 To 328), max0
        Dim mo
        mo = 0
        i = 1
        m = 1
        max0 = 0
        While k(i) <> Nothing
            ZH(m) = k(i) '桩号
            GC(m) = k(i + 1) '高差
            If GC(m) > max0 Then
                max0 = GC(m) '最大高差
            End If
            i = i + 2
            m = m + 1
        End While
        If H > max0 Then
            H = max0
        End If
        If m = 2 Then
            FCzh_z0 = ZH(1)
        Else
            i = 2
            pd = 0
            jdz(1) = 0
            While ZH(i) <> Nothing
                If (H <= GC(i - 1) And H >= GC(i)) Or (H >= GC(i - 1) And H <= GC(i)) Then '
                    pd = pd + 1
                    mo = mo + 1
                    If (GC(i) - GC(i - 1)) <> 0 Then
                        jdz(pd) = ZH(i - 1) + (ZH(i) - ZH(i - 1)) / (GC(i) - GC(i - 1)) * (H - GC(i - 1))
                    Else
                        jdz(pd) = ZH(i)
                    End If
                End If
                i = i + 1
            End While
            If H <= GC(m - 1) Then
                FCzh_z0 = ZH(m - 1)
            Else
                FCzh_z0 = jdz(mo)
            End If
        End If
    End Function

    Public Function sjs(min0, max) '生成随机数
        sjs = ExApp.WorksheetFunction.RandBetween(min0 * 1000, max * 1000) / 1000
    End Function

    Public Function JD_X(X1, Y1, X2, Y2, M1, N1, M2, N2) '计算2条直线的交点
        Dim k1, b1, k2, b2
        If X1 = X2 Then
            k1 = 32767
            b1 = X1
        Else
            k1 = (Y2 - Y1) / (X2 - X1)
            b1 = Y1 - k1 * X1
        End If
        If M1 = M2 Then
            k2 = 32767
            b2 = 32767
        Else
            k2 = (N2 - N1) / (M2 - M1)
            b2 = N1 - k2 * M1
        End If

        If k1 = 32767 And k2 <> 32767 Then
            JD_X = X1
        Else
            If k2 = k1 Then
                JD_X = -32767
            Else
                JD_X = (b2 - b1) / (k1 - k2)
            End If
        End If

    End Function
    Public Function JD_Y(X1, Y1, X2, Y2, M1, N1, M2, N2) '计算2条直线的交点
        Dim k1, b1, k2, b2
        If X1 = X2 Then
            k1 = 32767
            b1 = X1
        Else
            k1 = (Y2 - Y1) / (X2 - X1)
            b1 = Y1 - k1 * X1
        End If
        If M1 = M2 Then
            k2 = 32767
            b2 = 32767
        Else
            k2 = (N2 - N1) / (M2 - M1)
            b2 = N1 - k2 * M1
        End If

        If k1 = 32767 And k2 <> 32767 Then
            JD_Y = k2 * X1 + b2
        Else
            If k2 = k1 Then
                JD_Y = -32767
            Else
                If k1 = 0 Then
                    JD_Y = b1
                Else
                    JD_Y = (b2 - k2 * b1 / k1) / (1 - k2 / k1)
                End If
            End If
        End If

    End Function





    Public Function wkj_JK(Z, b, r, w, Lk, Lk2, HY, YH) As Double  'z桩号，B路基未加宽的宽度，w全加宽值，lk加宽曲线长。计算加宽后的宽度
        If r < 0 Then '左转加宽
            If Z <= (HY - Lk) Then
                wkj_JK = b
            Else
                If Z <= HY Then
                    If b < 0 Then
                        wkj_JK = b - (Z - (HY - Lk)) / Lk * w
                    Else
                        wkj_JK = b
                    End If
                Else
                    If Z <= YH Then
                        If b < 0 Then
                            wkj_JK = b - w
                        Else
                            wkj_JK = b
                        End If
                    Else
                        If Z <= (YH + Lk2) Then
                            If b < 0 Then
                                wkj_JK = b - ((YH + Lk2) - Z) / Lk2 * w
                            Else
                                wkj_JK = b
                            End If
                        Else
                            wkj_JK = b
                        End If
                    End If
                End If
            End If
        Else '右转加宽
            If Z <= (HY - Lk) Then
                wkj_JK = b
            Else
                If Z <= HY Then
                    If b <= 0 Then
                        wkj_JK = b
                    Else
                        wkj_JK = b + (Z - (HY - Lk)) / Lk * w
                    End If
                Else
                    If Z <= YH Then
                        If b <= 0 Then
                            wkj_JK = b
                        Else
                            wkj_JK = b + w
                        End If
                    Else
                        If Z <= (YH + Lk2) Then
                            If b <= 0 Then
                                wkj_JK = b
                            Else
                                wkj_JK = b + ((YH + Lk2) - Z) / Lk2 * w
                            End If
                        Else
                            wkj_JK = b
                        End If
                    End If
                End If
            End If
        End If
    End Function


    Public Function Pd_YS(Z, L, k) As Integer '进行要素判断，最后得到一要素值的位置。z为桩号，l为超高或者加宽缓和曲线长，k为第HY点桩号
        Dim n As Integer, no As Integer
        no = 1
        n = 1
        While k(n, 1) <> vbNull
            n = n + 1
        End While
        If n = 1 Then
            Pd_YS = -1
        Else
            If Z > (k(n - 1, 1) + L(n - 1, 1)) Then
                Pd_YS = n + 1
            Else
                n = 1
                While k(n, 1) <> vbNull And no = 1
                    If Z <= (k(n, 1) + L(n, 1)) Then
                        Pd_YS = n + 2
                        no = 0
                    Else : n = n + 1
                    End If
                End While
            End If
        End If
    End Function

    Public Function Pd_YSd(Z, L, k) As Integer '进行要素判断，最后得到一要素值的位置。z为桩号，l为超高或者加宽缓和曲线长，k为第HY点桩号
        Dim n As Integer, no As Integer
        no = 1
        n = 1
        While k(n, 1) <> Nothing
            n = n + 1
        End While
        If n = 1 Then
            Pd_YSd = -1
        Else
            If Z > (k(n - 1, 1) + L(n - 1, 1)) Then
                Pd_YSd = n + 1
            Else
                n = 1
                While k(n, 1) <> Nothing And no <= 2
                    If n = 1 Then
                        If Z <= (k(n, 1) + L(n, 1)) Then
                            Pd_YSd = n + 2
                            no = no + 1
                        End If
                    Else
                        If Z <= (k(n, 1) + L(n, 1)) And Z > (k(n - 1, 1) + L(n - 1, 1)) Then
                            Pd_YSd = n + 2
                            no = no + 1
                        End If
                    End If
                    n = n + 1
                End While
            End If
        End If
    End Function
    Public Function Zhz_1(k As Object) As Double '找出最后值。k数组
        Dim i As Integer
        i = 1
        While k(i, 1) <> vbNull
            i = i + 1
        End While
        If i = 1 Then
            Zhz_1 = -1
        Else : Zhz_1 = Val(k(i - 1, 1))
        End If
    End Function
    Public Function Zhz_2(k As Object) As Double '找出最后值的位置。k数组
        Dim i As Integer
        i = 1
        While k(i, 1) <> vbNull
            i = i + 1
        End While
        If i = 1 Then
            Zhz_2 = -1
        Else : Zhz_2 = i + 3
        End If
    End Function

    Public Function JK_ZY(Z, b) '加宽后宽度,Z桩号,b左右判断
        Dim i As Integer, n As Integer, m As Integer, N1 As Integer, j As Integer, kd, st
        st = Mid(Z, 1, 1)
        Z = Val(ExApp.WorksheetFunction.Substitute(Z, "*", ""))

        '加宽计算
        'n = Pd_YS(Z, ExApp.sheets("加宽参数").Range("d3 : d101 "), ExApp.sheets("加宽参数").Range("f3 : f101 "))
        '''''''''''''''''''
        If st <> "*" Then
            n = Pd_YSd(Z, ExApp.Sheets("加宽参数").Range("d3 : d101 ").value, ExApp.Sheets("加宽参数").Range("f3 : f101 ").value)
        Else
            n = Pd_YS(Z, ExApp.Sheets("加宽参数").Range("d3 : d101 ").value, ExApp.Sheets("加宽参数").Range("f3 : f101 ").value)
        End If
        '‘’‘’‘’‘’‘’‘
        '路面参数
        N1 = Pd_YS1(Z, ExApp.Sheets("路面参数").Range("r4 : r328 ").value) - 1
        If N1 = 4 Then
            N1 = 5
        End If
        kd = 0
        If Z <= ExApp.Sheets("路面参数").Cells(N1, 18) Then
            If b < 0 Then
                '左幅
                i = 17
                While ExApp.Sheets("路面参数").Cells(N1, i) <> Nothing And i > 1

                    kd = kd - (ExApp.Sheets("路面参数").Cells(N1 - 1, i) + (Z - ExApp.Sheets("路面参数").Cells(N1 - 1, 18)) * (ExApp.Sheets("路面参数").Cells(N1, i).value - ExApp.Sheets("路面参数").Cells(N1 - 1, i).value) / (ExApp.Sheets("路面参数").Cells(N1, 18).value - ExApp.Sheets("路面参数").Cells(N1 - 1, 18).value))
                    i = i - 2
                End While
            Else
                '右幅
                i = 19
                While ExApp.Sheets("路面参数").Cells(N1, i).value <> Nothing
                    kd = kd + (ExApp.Sheets("路面参数").Cells(N1 - 1, i).value + (Z - ExApp.Sheets("路面参数").Cells(N1 - 1, 18).value) * (ExApp.Sheets("路面参数").Cells(N1, i).value - ExApp.Sheets("路面参数").Cells(N1 - 1, i).value) / (ExApp.Sheets("路面参数").Cells(N1, 18).value - ExApp.Sheets("路面参数").Cells(N1 - 1, 18).value))
                    i = i + 2
                End While
            End If
        Else
            If b < 0 Then
                '左幅
                i = 17
                While ExApp.Sheets("路面参数").Cells(N1, i).value <> Nothing And i > 1

                    kd = kd - ExApp.Sheets("路面参数").Cells(N1, i).value
                    i = i - 2
                End While
            Else
                '右幅
                i = 19
                While ExApp.Sheets("路面参数").Cells(N1, i).value <> Nothing
                    kd = kd + ExApp.Sheets("路面参数").Cells(N1, i).value
                    i = i + 2
                End While
            End If

        End If
        ''''
        If n <> 3 Then '有弯道加宽
            JK_ZY = Math.Round(wkj_JK(Z, kd, ExApp.Sheets("加宽参数").Range("b" & n).value, ExApp.Sheets("加宽参数").Range("c" & n).value, ExApp.Sheets("加宽参数").Range("d" & n).value, ExApp.Sheets("加宽参数").Range("e" & n).value, ExApp.Sheets("加宽参数").Range("f" & n).value, ExApp.Sheets("加宽参数").Range("g" & n).value), 2)  '加宽
        Else
            JK_ZY = kd
        End If
    End Function

    Public Function JK(Z, b) '仅加宽值,Z桩号,b左右判断
        Dim i As Integer, n As Integer, m As Integer, N1 As Integer, j As Integer, kd
        Dim st
        st = Mid(Z, 1, 1)
        Z = Val(ExApp.WorksheetFunction.Substitute(Z, "*", ""))

        '加宽计算
        'n = Pd_YS(Z, ExApp.sheets("加宽参数").Range("d3 : d101 "), ExApp.sheets("加宽参数").Range("f3 : f101 "))
        ''''''''''''''
        If st <> "*" Then
            n = Pd_YSd(Z, ExApp.Sheets("加宽参数").Range("d3 : d101 ").value, ExApp.Sheets("加宽参数").Range("f3 : f101 ").value)
        Else
            n = Pd_YS(Z, ExApp.Sheets("加宽参数").Range("d3 : d101 ").value, ExApp.Sheets("加宽参数").Range("f3 : f101 ").value)
        End If
        '''''''''''''''''
        If n <> 3 Then '有弯道加宽
            JK = Math.Round(wkj_JK(Z, b, ExApp.Sheets("加宽参数").Range("b" & n).value, ExApp.Sheets("加宽参数").Range("c" & n).value, ExApp.Sheets("加宽参数").Range("d" & n).value, ExApp.Sheets("加宽参数").Range("e" & n).value, ExApp.Sheets("加宽参数").Range("f" & n).value, ExApp.Sheets("加宽参数").Range("g" & n).value), 2) - b '加宽
        Else
            JK = b - b
        End If
    End Function







    Public Function BG_s(Z, Z_B, H, r, T, w, i) As Double '竖曲线高程
        If Z <= (Z_B - T) Then
            BG_s = H - (Z_B - Z) * i
        Else
            If Z <= Z_B Then
                If w > 0 Then
                    BG_s = H - (Z_B - Z) * i - (Z - (Z_B - T)) ^ 2 / 2 / r
                Else : BG_s = H - (Z_B - Z) * i + (Z - (Z_B - T)) ^ 2 / 2 / r
                End If
            Else
                If Z <= (Z_B + T) Then
                    If w > 0 Then
                        BG_s = H + (Z - Z_B) * (i - w) - ((Z_B + T) - Z) ^ 2 / 2 / r
                    Else : BG_s = H + (Z - Z_B) * (i - w) + ((Z_B + T) - Z) ^ 2 / 2 / r
                    End If
                Else : BG_s = H + (Z - Z_B) * (i - w)
                End If
            End If
        End If


    End Function
    Public Function Pd_YS1(Z As Object, k As Object) As Integer '进行要素判断，最后得到一要素值的位置。z为桩号，k为竖曲线终点
        Dim n As Integer, no As Integer
        no = 1
        n = 1
        While k(n, 1) <> Nothing Or k(n, 1) = "0"
            n = n + 1
        End While
        If n = 1 Then
            Pd_YS1 = -1
        Else
            If Z > k(n - 1, 1) Then
                Pd_YS1 = n + 3
            Else
                n = 1
                While k(n, 1) <> Nothing And no = 1
                    If Z <= k(n, 1) Then
                        Pd_YS1 = n + 4
                        no = 0
                    Else : n = n + 1
                    End If
                End While
            End If
        End If
    End Function
    Public Function Pd_YS1D(Z As Object, k As Object) As Integer '进行要素判断，最后得到一要素值的位置。z为桩号，k为竖曲线终点
        Dim n As Integer, no As Integer
        no = 1
        n = 1
        While k(n, 1) <> Nothing
            n = n + 1
        End While
        If n = 1 Then
            Pd_YS1D = -1
        Else
            If Z > k(n - 1, 1) Then
                Pd_YS1D = n + 3
            Else
                n = 1
                While k(n, 1) <> Nothing And no <= 2
                    If n = 1 Then
                        If Z <= k(n, 1) Then
                            Pd_YS1D = n + 4
                            no = no + 1
                        End If
                    Else
                        If Z <= k(n, 1) And Z > k(n - 1, 1) Then
                            Pd_YS1D = n + 4
                            no = no + 1
                        End If
                    End If
                    n = n + 1
                End While
            End If
        End If
    End Function
    Public Function Zhz_11(k As Object) As Double '找出最后值。k数组
        Dim i As Integer
        i = 1
        While k(i) <> Nothing
            i = i + 1
        End While
        If i = 1 Then
            Zhz_11 = -1
        Else : Zhz_11 = Val(k(i - 1))
        End If
    End Function

    Public Function BG_s_L(Z, Z_B1, Z_B2, h1, h2) As Double '直线高程
        BG_s_L = h1 + (Z - Z_B1) * (h2 - h1) / (Z_B2 - Z_B1)
    End Function







    '绕中桩超高横坡计算
    Public Function wkj_HP(Z, b, Z_Q, Z_Z, iz1, iz2, iy1, iy2)
        Dim i, fs, HP, e, Lc
        Lc = Z_Z - Z_Q
        fs = ExApp.Sheets("横坡参数").Range("F3").value
        If fs = "x" Or fs = "X" Then
            '线型超高
            If Z <= Z_Q Then
                If b < 0 Then '左幅
                    HP = -iz1
                Else
                    HP = iy1
                End If
            Else
                If Z <= Z_Z Then
                    If b < 0 Then '左幅
                        HP = -((Z - Z_Q) / Lc * (iz2 - iz1) + iz1)
                    Else
                        HP = ((Z - Z_Q) / Lc * (iy2 - iy1) + iy1)
                    End If
                Else
                    If b < 0 Then '左幅
                        HP = -iz2
                    Else
                        HP = iy2
                    End If
                End If
            End If
        Else '三次超高
            If Z <= Z_Q Then
                If b < 0 Then '左幅
                    HP = -iz1
                Else
                    HP = iy1
                End If
            Else
                If Z <= Z_Z Then
                    If b < 0 Then '左幅i1 +(i2-i1)×e^2/Lc^2×(3-2×e/Lc),e为桩号到起点的长度
                        HP = -(iz1 + (iz2 - iz1) * (Z - Z_Q) ^ 2 / Lc ^ 2 * (3 - 2 * (Z - Z_Q) / Lc))
                    Else
                        HP = (iy1 + (iy2 - iy1) * (Z - Z_Q) ^ 2 / Lc ^ 2 * (3 - 2 * (Z - Z_Q) / Lc))
                    End If
                Else
                    If b < 0 Then '左幅
                        HP = -iz2
                    Else
                        HP = iy2
                    End If
                End If
            End If
        End If
        wkj_HP = HP
    End Function

    '''超高横坡达到单向路拱横坡时的长度
    Public Function lX(Lc, ib, i) 'Lc为超高过渡段长度，ib为超高横坡，i为路拱横坡
        i = Math.Abs(i)
        ib = Math.Abs(ib)
        lX = Lc * 2 * i / (ib + i)
    End Function






    Public Function FSZhj(X, y) '反算坐标，计算桩号

        sheet2 = Exbook.Worksheets("交点法") '交点法
        Dim i As Integer, Qz, Fx, Fy As Double, Qz1, Jl, Jl1, Jl2, Qx, Qy, n, Zx, jz
        i = 4 '坐标循环
        Fx = X '计算x坐标
        Fy = y '计算y坐标
        Qz = sheet2.Range("D4").Value  '起点桩号
        Jl1 = 1
        While Math.Abs(Jl1) > 0.02
            n = Pd_YSw(Qz, sheet2.Range("O5 : O328").Value)   '对应桩号所在线元的位置
            If n = -1 Then
                MsgBox("请在交点法表内输入数据")
                Exit Function
            Else
                '计算起点桩号坐标
                Qx = Zzzb_X(Qz, sheet2.Range("E" & n).Value, sheet2.Range("F" & n).Value, sheet2.Range("G" & n).Value, sheet2.Range("Q" & n).Value, sheet2.Range("D" & n).Value, sheet2.Range("B" & n).Value, sheet2.Range("P" & n).Value)
                Qy = Zzzb_Y(Qz, sheet2.Range("E" & n).Value, sheet2.Range("F" & n).Value, sheet2.Range("G" & n).Value, sheet2.Range("Q" & n).Value, sheet2.Range("D" & n).Value, sheet2.Range("C" & n).Value, sheet2.Range("P" & n).Value)
            End If
            jz = Jsfwj_1j(Qx, Qy, Fx, Fy) '两点的夹角
            Zx = Zxfwj_B(Qz, sheet2.Range("E" & n).Value, sheet2.Range("F" & n).Value, sheet2.Range("G" & n).Value, sheet2.Range("Q" & n).Value, sheet2.Range("P" & n).Value, sheet2.Range("D" & n).Value)  '计算中桩走向方位角，计算后的单位为度。z为计算桩号。r为圆曲线半径。Lh为缓和曲线总长。Zj为交点转角，其单位为度。Jsfwj为直缓点方位角，其单位为度。z为计算桩号。ZH为直缓点桩号。HY为缓圆点桩号。YH为圆缓点桩号。HZ为缓直点桩号
            Jl = Math.Sqrt((Qx - Fx) ^ 2 + (Qy - Fy) ^ 2)
            Jl1 = Jl * Math.Cos((jz - Zx) / 180 * 3.1415926)
            If Jl1 < 0 Then
                FSZhj = Math.Round(Qz, 3)
                Exit Function
            End If
            Qz = Qz + Jl1
            If Math.Abs(Jl1) <= 0.02 Then
                FSZhj = Math.Round(Qz, 3)
            End If
        End While
    End Function
    Public Function FSPJj(X, y) '反算坐标，计算偏距
        sheet2 = Exbook.Worksheets("交点法") '交点法
        Dim Fx, Fy, i As Integer, Qz As Double, Qz1, Jl, Jl1, Jl2, Qx, Qy, n, Zx, jz
        i = 4 '坐标循环
        Fx = X '计算x坐标
        Fy = y '计算y坐标
        Qz = sheet2.Range("d4").Value  '起点桩号
        Jl1 = 1
        While Math.Abs(Jl1) > 0.02
            n = Pd_YSw(Qz, sheet2.Range("o5 : o328").Value)   '对应桩号所在线元的位置
            If n = -1 Then
                MsgBox("请在交点法表内输入数据")
                Exit Function
            Else
                '计算起点桩号坐标
                Qx = Zzzb_X(Qz, sheet2.Range("f" & n).Value, sheet2.Range("g" & n).Value, sheet2.Range("h" & n).Value, sheet2.Range("q" & n).Value, sheet2.Range("d" & n).Value, sheet2.Range("b" & n).Value, sheet2.Range("p" & n).Value)
                Qy = Zzzb_Y(Qz, sheet2.Range("f" & n).Value, sheet2.Range("g" & n).Value, sheet2.Range("h" & n).Value, sheet2.Range("q" & n).Value, sheet2.Range("d" & n).Value, sheet2.Range("c" & n).Value, sheet2.Range("p" & n).Value)
            End If
            jz = Jsfwj_1j(Qx, Qy, Fx, Fy) '两点的夹角
            Zx = Zxfwj_B(Qz, sheet2.Range("f" & n).Value, sheet2.Range("g" & n).Value, sheet2.Range("h" & n).Value, sheet2.Range("q" & n).Value, sheet2.Range("p" & n).Value, sheet2.Range("d" & n).Value)  '计算中桩走向方位角，计算后的单位为度。z为计算桩号。r为圆曲线半径。Lh为缓和曲线总长。Zj为交点转角，其单位为度。Jsfwj为直缓点方位角，其单位为度。z为计算桩号。ZH为直缓点桩号。HY为缓圆点桩号。YH为圆缓点桩号。HZ为缓直点桩号
            Jl = Math.Sqrt((Qx - Fx) ^ 2 + (Qy - Fy) ^ 2)
            Jl1 = Jl * Math.Cos((jz - Zx) / 180 * 3.1415926)
            Qz = Qz + Jl1
            If Math.Abs(Jl1) <= 0.02 Then
                'FSZh = math.Round(Qz, 3)
                FSPJj = Math.Round(Jl * Math.Sin((jz - Zx) / 180 * 3.1415926), 3)
            End If
        End While
    End Function

    Public Function FSZHx(X, y) '反算坐标，计算桩号
        sheet3 = Exbook.Worksheets("线元法") '线元法
        Dim Fx, Fy, Qz As Double, Qz1, Jl, Jl1, Jl2, Qx, Qy, n, Zx, jz
        Fx = X '计算x坐标
        Fy = y '计算y坐标
        Qz = sheet3.Range("b3").Value '起点桩号
        Jl1 = 1
        While Math.Abs(Jl1) > 0.02
            '计算起点桩号坐标
            Qx = XYF_X(Qz, 0, 90)
            Qy = XYF_Y(Qz, 0, 90)
            jz = Jsfwj_1x(Qx, Qy, Fx, Fy) '两点的夹角
            Zx = XYF_a(Qz) * 180 / PI
            Jl = Math.Sqrt((Qx - Fx) ^ 2 + (Qy - Fy) ^ 2)
            Jl1 = Jl * Math.Cos((jz - Zx) / 180 * PI)
            Qz = Qz + Jl1
            If Math.Abs(Jl1) <= 0.02 Then
                FSZHx = Math.Round(Qz, 3)
                'FSPJ = math.Round(Jl * math.Sin((jz - Zx) / 180 * 3.1415926), 3)
            End If
        End While
    End Function
    Public Function FSPJx(X, y) '反算坐标，计算偏距
        sheet3 = Exbook.Worksheets("线元法") '线元法
        Dim Fx, Fy, Qz As Double, Qz1, Jl, Jl1, Jl2, Qx, Qy, n, Zx, jz
        Fx = X '计算x坐标
        Fy = y '计算y坐标
        Qz = sheet3.Range("b3").Value '起点桩号
        Jl1 = 1
        While Math.Abs(Jl1) > 0.02
            '计算起点桩号坐标
            Qx = XYF_X(Qz, 0, 90)
            Qy = XYF_Y(Qz, 0, 90)
            jz = Jsfwj_1x(Qx, Qy, Fx, Fy) '两点的夹角
            Zx = XYF_a(Qz) * 180 / PI
            Jl = Math.Sqrt((Qx - Fx) ^ 2 + (Qy - Fy) ^ 2)
            Jl1 = Jl * Math.Cos((jz - Zx) / 180 * PI)
            Qz = Qz + Jl1
            If Math.Abs(Jl1) <= 0.02 Then
                'FSZH = math.Round(QZ, 3)
                FSPJx = Math.Round(Jl * Math.Sin((jz - Zx) / 180 * 3.1415926), 3)
            End If
        End While
    End Function
    Public Function FSZH(X, y) '反算桩号
        If sheet3.Range("j2").Value = "是" Then
            FSZH = FSZHx(X, y)
        Else
            FSZH = FSZhj(X, y)
        End If
    End Function
    Public Function FSPJ(X, y) '反算偏距
        If sheet3.Range("j2").Value = "是" Then
            FSPJ = FSPJx(X, y)
        Else
            FSPJ = FSPJj(X, y)
        End If
    End Function











    Public Function LMGC(Z, BJ) '路面高程
        sheet4 = Exbook.Worksheets("断链") '断链
        Dim n As Integer, m As Integer, N1, Gc_lmd, M1, st
        Dim i, GC_S, KD_S, GC, pd, JQkd, JQhp, DLi, DLpd, DLc, Z1
        Dim i12, no
        st = Mid(Z, 1, 1)
        Z = Val(ExApp.WorksheetFunction.Substitute(Z, "*", ""))
        Z1 = Z
        If st <> "*" Then
            n = Pd_YS1D(Z, ExApp.Sheets("竖曲线").Range("h5 : h328 ").value)   '对应桩号所在线元的位置
            i12 = 3
            no = 1
            While sheet4.Range("B" & i12).Value <> Nothing And no = 1
                If sheet4.Range("c" & i12).Value <= Z And sheet4.Range("B" & i12).Value > Z Then
                    n = Pd_YS1(sheet4.Range("B" & i12).Value, ExApp.Sheets("竖曲线").Range("h5 : h328 ").value)   '对应桩号所在线元的位置
                    no = 0
                End If
                i12 = i12 + 1
            End While
        Else
            n = Pd_YS1(Z, ExApp.Sheets("竖曲线").Range("h5 : h328 ").value)  '对应桩号所在线元的位置
        End If
        i12 = 3
        no = 1
        DLc = 0
        While sheet4.Range("B" & i12).Value <> Nothing And no = 1
            If sheet4.Range("B" & i12).Value < ExApp.Sheets("竖曲线").Range("b" & n).value And sheet4.Range("B" & i12).Value > ExApp.Sheets("竖曲线").Range("b" & n - 1).value Then
                DLc = (sheet4.Range("C" & i12).Value - sheet4.Range("B" & i12).Value)
                If Z > sheet4.Range("c" & i12).Value Then
                    Z1 = Z - DLc
                End If
                no = 0
            End If
            i12 = i12 + 1
        End While
        '标高计算
        If n >= 5 Then
            Gc_lmd = Math.Round(BG_s(Z1, ExApp.Sheets("竖曲线").Range("b" & n).value - DLc, ExApp.Sheets("竖曲线").Range("c" & n).value, ExApp.Sheets("竖曲线").Range("d" & n).value, ExApp.Sheets("竖曲线").Range("e" & n).value, ExApp.Sheets("竖曲线").Range("i" & n).value, ExApp.Sheets("竖曲线").Range("j" & n).value), 3)
        Else
            Gc_lmd = BG_s_L(Z1, ExApp.Sheets("竖曲线").Range("b" & 4).value, ExApp.Sheets("竖曲线").Range("b" & 5).value, ExApp.Sheets("竖曲线").Range("c" & 4).value, ExApp.Sheets("竖曲线").Range("c" & 5).value)
        End If



        ''超高计算
        'If st <> "*" Then
        '    m = Pd_YS1D(Z, ExApp.Sheets("横坡参数").Range("b4 : b328 ").value) - 1
        'Else
        '    m = Pd_YS1(Z, ExApp.Sheets("横坡参数").Range("b4 : b328 ").value) - 1
        'End If

        'If m = 4 Then
        '    M1 = 4
        '    m = 5
        'Else
        '    M1 = m - 1
        'End If
        ''路面参数
        'N1 = Pd_YS1(Z, ExApp.Sheets("路面参数").Range("r4 : r328 ").value) - 1
        'If N1 = 4 Then
        '    N1 = 5
        'End If
        ''加宽
        'If st <> "*" Then
        '    n = Pd_YSd(Z, ExApp.Sheets("加宽参数").Range("d3 : d101 ").value, ExApp.Sheets("加宽参数").Range("f3 : f101 ").value)
        'Else
        '    n = Pd_YS(Z, ExApp.Sheets("加宽参数").Range("d3 : d101 ").value, ExApp.Sheets("加宽参数").Range("f3 : f101 ").value)
        'End If
        ''wkj_HP(Z, b, Z_Q, Z_Z, iz1, iz2, iy1, iy2)
        'GC_S = 0 ' 高程合
        'KD_S = 0 '宽度合
        'GC = 0
        'If BJ <= 0 Then
        '    ''''左幅
        '    i = 17
        '    pd = 1
        '    While Math.Abs(BJ) > KD_S And ExApp.Sheets("路面参数").Cells(N1, i).value <> Nothing
        '        JQkd = Math.Abs((ExApp.Sheets("路面参数").Cells(N1 - 1, i).value + (Z - ExApp.Sheets("路面参数").Cells(N1 - 1, 18).value) * (ExApp.Sheets("路面参数").Cells(N1, i).value - ExApp.Sheets("路面参数").Cells(N1 - 1, i).value) / (ExApp.Sheets("路面参数").Cells(N1, 18).value - ExApp.Sheets("路面参数").Cells(N1 - 1, 18).value)))
        '        JQhp = (ExApp.Sheets("路面参数").Cells(N1 - 1, i - 1).value + (Z - ExApp.Sheets("路面参数").Cells(N1 - 1, 18).value) * (ExApp.Sheets("路面参数").Cells(N1, i - 1).value - ExApp.Sheets("路面参数").Cells(N1 - 1, i - 1).value) / (ExApp.Sheets("路面参数").Cells(N1, 18).value - ExApp.Sheets("路面参数").Cells(N1 - 1, 18).value))
        '        If i = 15 Then
        '            KD_S = KD_S + Math.Round(Math.Abs(wkj_JK(Z, -JQkd, ExApp.Sheets("加宽参数").Range("b" & n).value, ExApp.Sheets("加宽参数").Range("c" & n).value, ExApp.Sheets("加宽参数").Range("d" & n).value, ExApp.Sheets("加宽参数").Range("e" & n).value, ExApp.Sheets("加宽参数").Range("f" & n).value, ExApp.Sheets("加宽参数").Range("g" & n).value)), 3)
        '            GC_S = GC_S + Math.Round(Math.Abs(wkj_JK(Z, -JQkd, ExApp.Sheets("加宽参数").Range("b" & n).value, ExApp.Sheets("加宽参数").Range("c" & n).value, ExApp.Sheets("加宽参数").Range("d" & n).value, ExApp.Sheets("加宽参数").Range("e" & n).value, ExApp.Sheets("加宽参数").Range("f" & n).value, ExApp.Sheets("加宽参数").Range("g" & n).value)), 3) * wkj_HP(Z, BJ, ExApp.Sheets("横坡参数").Range("b" & M1).value, ExApp.Sheets("横坡参数").Range("b" & m).value, ExApp.Sheets("横坡参数").Range("c" & M1).value, ExApp.Sheets("横坡参数").Range("c" & m).value, ExApp.Sheets("横坡参数").Range("d" & M1).value, ExApp.Sheets("横坡参数").Range("d" & m).value) / 100
        '        Else
        '            KD_S = KD_S + JQkd
        '            GC_S = GC_S - JQkd * JQhp / 100
        '        End If
        '        i = i - 2
        '    End While
        '    If (i + 2) = 15 Then
        '        GC = GC_S + (Math.Abs(BJ) - KD_S) * wkj_HP(Z, BJ, ExApp.Sheets("横坡参数").Range("b" & M1).value, ExApp.Sheets("横坡参数").Range("b" & m).value, ExApp.Sheets("横坡参数").Range("c" & M1).value, ExApp.Sheets("横坡参数").Range("c" & m).value, ExApp.Sheets("横坡参数").Range("d" & M1).value, ExApp.Sheets("横坡参数").Range("d" & m).value) / 100
        '    Else
        '        GC = GC_S - (Math.Abs(BJ) - KD_S) * JQhp / 100
        '    End If
        '    If (i + 2) <= 9 Then
        '        GC = GC + (ExApp.Sheets("路面参数").Cells(N1 - 1, 1).value + (Z - ExApp.Sheets("路面参数").Cells(N1 - 1, 18).value) * (ExApp.Sheets("路面参数").Cells(N1, 1).value - ExApp.Sheets("路面参数").Cells(N1 - 1, 1).value) / (ExApp.Sheets("路面参数").Cells(N1, 18).value - ExApp.Sheets("路面参数").Cells(N1 - 1, 18).value))
        '    End If
        'Else
        '    ''''右幅
        '    i = 19
        '    pd = 1
        '    While Math.Abs(BJ) > KD_S And ExApp.Sheets("路面参数").Cells(N1, i).value <> Nothing
        '        JQkd = (ExApp.Sheets("路面参数").Cells(N1 - 1, i).value + (Z - ExApp.Sheets("路面参数").Cells(N1 - 1, 18).value) * (ExApp.Sheets("路面参数").Cells(N1, i).value - ExApp.Sheets("路面参数").Cells(N1 - 1, i).value) / (ExApp.Sheets("路面参数").Cells(N1, 18).value - ExApp.Sheets("路面参数").Cells(N1 - 1, 18).value))
        '        JQhp = (ExApp.Sheets("路面参数").Cells(N1 - 1, i + 1).value + (Z - ExApp.Sheets("路面参数").Cells(N1 - 1, 18).value) * (ExApp.Sheets("路面参数").Cells(N1, i + 1).value - ExApp.Sheets("路面参数").Cells(N1 - 1, i + 1).value) / (ExApp.Sheets("路面参数").Cells(N1, 18).value - ExApp.Sheets("路面参数").Cells(N1 - 1, 18).value))
        '        If i = 21 Then
        '            KD_S = KD_S + Math.Round(wkj_JK(Z, JQkd, ExApp.Sheets("加宽参数").Range("b" & n).value, ExApp.Sheets("加宽参数").Range("c" & n).value, ExApp.Sheets("加宽参数").Range("d" & n).value, ExApp.Sheets("加宽参数").Range("e" & n).value, ExApp.Sheets("加宽参数").Range("f" & n).value, ExApp.Sheets("加宽参数").Range("g" & n).value), 2)
        '            GC_S = GC_S + Math.Round(wkj_JK(Z, JQkd, ExApp.Sheets("加宽参数").Range("b" & n).value, ExApp.Sheets("加宽参数").Range("c" & n).value, ExApp.Sheets("加宽参数").Range("d" & n).value, ExApp.Sheets("加宽参数").Range("e" & n).value, ExApp.Sheets("加宽参数").Range("f" & n).value, ExApp.Sheets("加宽参数").Range("g" & n).value), 2) * wkj_HP(Z, BJ, ExApp.Sheets("横坡参数").Range("b" & M1).value, ExApp.Sheets("横坡参数").Range("b" & m).value, ExApp.Sheets("横坡参数").Range("c" & M1).value, ExApp.Sheets("横坡参数").Range("c" & m).value, ExApp.Sheets("横坡参数").Range("d" & M1).value, ExApp.Sheets("横坡参数").Range("d" & m).value) / 100
        '        Else
        '            KD_S = KD_S + JQkd
        '            GC_S = GC_S + JQkd * JQhp / 100
        '        End If
        '        i = i + 2
        '    End While
        '    If (i - 2) = 21 Then
        '        GC = (GC_S) + (Math.Abs(BJ) - KD_S) * wkj_HP(Z, BJ, ExApp.Sheets("横坡参数").Range("b" & M1).value, ExApp.Sheets("横坡参数").Range("b" & m).value, ExApp.Sheets("横坡参数").Range("c" & M1).value, ExApp.Sheets("横坡参数").Range("c" & m).value, ExApp.Sheets("横坡参数").Range("d" & M1).value, ExApp.Sheets("横坡参数").Range("d" & m).value) / 100
        '    Else
        '        GC = GC_S + (Math.Abs(BJ) - KD_S) * JQhp / 100
        '    End If
        '    ''''人行道道牙
        '    If (i - 2) >= 27 Then
        '        GC = GC + (ExApp.Sheets("路面参数").Cells(N1 - 1, 1).value + (Z - ExApp.Sheets("路面参数").Cells(N1 - 1, 18).value) * (ExApp.Sheets("路面参数").Cells(N1, 1).value - ExApp.Sheets("路面参数").Cells(N1 - 1, 1).value) / (ExApp.Sheets("路面参数").Cells(N1, 18).value - ExApp.Sheets("路面参数").Cells(N1 - 1, 18).value))
        '    End If
        'End If
        'LMGC = Gc_lmd + GC  '设计高程+横坡+超高
        LMGC = Gc_lmd   '设计高程
    End Function

    Public Function HP(Z, BJ) '路面横坡
        Dim n As Integer, m As Integer, N1, Gc_lmd, M1
        Dim i, KD_S, GC, pd, JQkd, JQhp, st
        '超高计算
        If st <> "*" Then
            m = Pd_YS1D(Z, ExApp.Sheets("横坡参数").Range("b4 : b328 ").value) - 1
        Else
            m = Pd_YS1(Z, ExApp.Sheets("横坡参数").Range("b4 : b328 ").value) - 1
        End If
        If m = 4 Then
            M1 = 4
            m = 5
        Else
            M1 = m - 1
        End If
        '加宽
        If st <> "*" Then
            n = Pd_YSd(Z, ExApp.Sheets("加宽参数").Range("d3 : d101 ").value, ExApp.Sheets("加宽参数").Range("f3 : f101 ").value)
        Else
            n = Pd_YS(Z, ExApp.Sheets("加宽参数").Range("d3 : d101 ").value, ExApp.Sheets("加宽参数").Range("f3 : f101 ").value)
        End If
        '路面参数
        N1 = Pd_YS1(Z, ExApp.Sheets("路面参数").Range("r4 : r328 ").value) - 1
        If N1 = 4 Then
            N1 = 5
        End If
        'wkj_HP(Z, b, Z_Q, Z_Z, iz1, iz2, iy1, iy2)
        KD_S = 0 '宽度合
        If BJ <= 0 Then
            ''''左幅
            i = 17
            pd = 1
            JQhp = (ExApp.Sheets("路面参数").Cells(N1 - 1, i - 1).value + (Z - ExApp.Sheets("路面参数").Cells(N1 - 1, 18).value) * (ExApp.Sheets("路面参数").Cells(N1, i - 1).value - ExApp.Sheets("路面参数").Cells(N1 - 1, i - 1).value) / (ExApp.Sheets("路面参数").Cells(N1, 18).value - ExApp.Sheets("路面参数").Cells(N1 - 1, 18).value))
            While Math.Abs(BJ) > KD_S And ExApp.Sheets("路面参数").Cells(N1, i).value <> Nothing
                JQkd = Math.Abs((ExApp.Sheets("路面参数").Cells(N1 - 1, i).value + (Z - ExApp.Sheets("路面参数").Cells(N1 - 1, 18).value) * (ExApp.Sheets("路面参数").Cells(N1, i).value - ExApp.Sheets("路面参数").Cells(N1 - 1, i).value) / (ExApp.Sheets("路面参数").Cells(N1, 18).value - ExApp.Sheets("路面参数").Cells(N1 - 1, 18).value)))
                JQhp = (ExApp.Sheets("路面参数").Cells(N1 - 1, i - 1).value + (Z - ExApp.Sheets("路面参数").Cells(N1 - 1, 18).value) * (ExApp.Sheets("路面参数").Cells(N1, i - 1).value - ExApp.Sheets("路面参数").Cells(N1 - 1, i - 1).value) / (ExApp.Sheets("路面参数").Cells(N1, 18).value - ExApp.Sheets("路面参数").Cells(N1 - 1, 18).value))
                If i = 15 Then
                    KD_S = KD_S + Math.Round(Math.Abs(wkj_JK(Z, -JQkd, ExApp.Sheets("加宽参数").Range("b" & n).value, ExApp.Sheets("加宽参数").Range("c" & n).value, ExApp.Sheets("加宽参数").Range("d" & n).value, ExApp.Sheets("加宽参数").Range("e" & n).value, ExApp.Sheets("加宽参数").Range("f" & n).value, ExApp.Sheets("加宽参数").Range("g" & n).value)), 3)
                    GC = -wkj_HP(Z, BJ, ExApp.Sheets("横坡参数").Range("b" & M1).value, ExApp.Sheets("横坡参数").Range("b" & m).value, ExApp.Sheets("横坡参数").Range("c" & M1).value, ExApp.Sheets("横坡参数").Range("c" & m).value, ExApp.Sheets("横坡参数").Range("d" & M1).value, ExApp.Sheets("横坡参数").Range("d" & m).value)
                Else
                    KD_S = KD_S + JQkd
                    GC = JQhp
                End If
                i = i - 2
            End While

        Else
            ''''右幅
            i = 19
            pd = 1
            While Math.Abs(BJ) > KD_S And ExApp.Sheets("路面参数").Cells(N1, i).value <> Nothing
                JQkd = (ExApp.Sheets("路面参数").Cells(N1 - 1, i).value + (Z - ExApp.Sheets("路面参数").Cells(N1 - 1, 18).value) * (ExApp.Sheets("路面参数").Cells(N1, i).value - ExApp.Sheets("路面参数").Cells(N1 - 1, i).value) / (ExApp.Sheets("路面参数").Cells(N1, 18).value - ExApp.Sheets("路面参数").Cells(N1 - 1, 18).value))
                JQhp = (ExApp.Sheets("路面参数").Cells(N1 - 1, i + 1).value + (Z - ExApp.Sheets("路面参数").Cells(N1 - 1, 18).value) * (ExApp.Sheets("路面参数").Cells(N1, i + 1).value - ExApp.Sheets("路面参数").Cells(N1 - 1, i + 1).value) / (ExApp.Sheets("路面参数").Cells(N1, 18).value - ExApp.Sheets("路面参数").Cells(N1 - 1, 18).value))
                If i = 21 Then
                    KD_S = KD_S + Math.Round(wkj_JK(Z, JQkd, ExApp.Sheets("加宽参数").Range("b" & n).value, ExApp.Sheets("加宽参数").Range("c" & n).value, ExApp.Sheets("加宽参数").Range("d" & n).value, ExApp.Sheets("加宽参数").Range("e" & n).value, ExApp.Sheets("加宽参数").Range("f" & n).value, ExApp.Sheets("加宽参数").Range("g" & n).value), 3)
                    GC = wkj_HP(Z, BJ, ExApp.Sheets("横坡参数").Range("b" & M1).value, ExApp.Sheets("横坡参数").Range("b" & m).value, ExApp.Sheets("横坡参数").Range("c" & M1).value, ExApp.Sheets("横坡参数").Range("c" & m).value, ExApp.Sheets("横坡参数").Range("d" & M1).value, ExApp.Sheets("横坡参数").Range("d" & m).value)
                Else
                    KD_S = KD_S + JQkd
                    GC = JQhp
                End If
                i = i + 2
            End While
        End If
        HP = GC
    End Function





    Public Function HYH_Z(Z, JJ) '寻找小于间距的HY点桩号
        sheet3 = Exbook.Worksheets("线元法") '线元法
        Dim j, XYc, pd
        Z = Math.Round(Z, 5)
        j = 3
        pd = 1
        HYH_Z = -32767
        XYc = sheet3.Range("b" & j).Value
        j = 3
        While sheet3.Range("f" & j).Value <> Nothing And pd = 1
            If sheet3.Range("b" & j).Value = Nothing Then
                XYc = Math.Round(XYc + sheet3.Range("f" & j - 1).Value, 5)
            Else
                XYc = Math.Round(sheet3.Range("b" & j).Value, 5)
            End If
            If (Z + JJ) > XYc And Z < XYc And sheet3.Range("g" & j).Value = sheet3.Range("h" & j).Value And sheet3.Range("g" & j).Value <> 0 Then
                HYH_Z = XYc
                pd = 0
            End If
            If (Z + JJ) > (XYc + sheet3.Range("f" & j).Value) And Z < Math.Round((XYc + sheet3.Range("f" & j).Value), 5) And sheet3.Range("g" & j).Value = sheet3.Range("h" & j).Value And sheet3.Range("g" & j).Value <> 0 Then
                HYH_Z = XYc + sheet3.Range("f" & j).Value
                pd = 0
            End If
            j = j + 1
        End While
        'HYH_Z = j
    End Function

    Public Function DLz(Z) '线元法断链前桩号对应线元的终桩+断链长度
        sheet3 = Exbook.Worksheets("线元法") '线元法
        sheet4 = Exbook.Worksheets("断链") '断链
        Dim i, m, j, XYc, Xyz, gs, st, pd, pd1, xyzi
        gs = Len(Z)
        st = Mid(Z, 1, 1)
        Z = Val(ExApp.WorksheetFunction.Substitute(Z, "*", ""))
        i = 3
        m = Z
        pd = 1
        xyzi = 0
        '寻找z在最小断链开始桩号
        While sheet4.Range("b" & i).Value <> Nothing
            If (sheet4.Range("c" & i).Value <= Z) Then   ' And st <> "*" Or (sheet4.Range("c" & i) >= Z And sheet4.Range("b" & i) <= Z)
                m = sheet4.Range("b" & i).Value '断链
                xyzi = i
            End If
            i = i + 1
        End While
        '寻找最小断链位于交点法的终点桩号
        pd = 1
        pd1 = 0
        XYc = sheet3.Range("b" & 3).Value
        Xyz = 0 ' sheet3.Range("b" & 3) + sheet3.Range("f" & 3)
        j = 4
        While sheet3.Range("f" & j).Value <> Nothing And pd = 1
            If sheet3.Range("b" & j).Value = Nothing Then
                XYc = XYc + sheet3.Range("f" & j - 1).Value
            Else
                XYc = sheet3.Range("b" & j).Value
            End If
            If m >= XYc Or sheet3.Range("b" & j).Value = Nothing Then
                Xyz = XYc + sheet3.Range("f" & j).Value
                'Xyz = j
            End If
            If m < XYc + sheet3.Range("f" & j).Value And sheet3.Range("b" & j + 1).Value <> Nothing Then  'And st = "*" Then
                pd = 0
            End If
            j = j + 1
        End While
        If xyzi <> 0 And st <> "*" Then
            DLz = Xyz + (sheet4.Range("c" & xyzi).Value - sheet4.Range("b" & xyzi).Value)
        Else
            DLz = 0
        End If
    End Function
    Public Function XYF_X(Z, BJ, JD) '线元法断链
        sheet3 = Exbook.Worksheets("线元法") '线元法
        sheet4 = Exbook.Worksheets("断链") '断链
        ''''''交点法
        Dim i12, no, n, x0, y0, st, gs, zi, z0, Z1
        gs = Len(Z)
        st = Mid(Z, 1, 1)
        Z = Val(ExApp.WorksheetFunction.Substitute(Z, "*", ""))
        z0 = DLz(Z) '断链点对应线元+断链长度
        Z1 = z0
        '短断链
        i12 = 3
        no = 1
        zi = 0
        While sheet4.Range("B" & i12).Value <> Nothing And no = 1
            If Z >= sheet4.Range("c" & i12).Value And Z <= z0 And st <> "*" Then
                zi = i12
            End If
            i12 = i12 + 1
        End While
        If zi <> 0 Then
            Z = Z - (sheet4.Range("c" & zi).Value - sheet4.Range("b" & zi).Value)
            Z1 = z0 - (sheet4.Range("c" & zi).Value - sheet4.Range("b" & zi).Value)
        End If
        ''''''''''''''''
        ''''''''交点法
        '开始计算坐标
        Dim R1, R2, R3, R4, R5, V1, V2, V3, V4, V5, a, AA, b, Ls, L, KA, KB, KAB, XA, YA
        Dim Qz, ZZ, R_Q, R_Z, i, pd 'ZZ位线元终点桩号,QZ线元起点桩号,R_Q为起点半径, R_Z终点半径
        Dim b1, b2, b3, b4, b5, Xyi, XYc, THi, pd1
        R1 = 0.1184634425
        R5 = R1
        R2 = 0.2393143352
        R4 = R2
        R3 = 0.2844444444
        V1 = 0.04691007
        V5 = 1 - V1
        V2 = 0.2307653449
        V4 = 1 - V2
        V3 = 0.5
        ''寻找
        Xyi = 3
        AA = sheet3.Range("e" & Xyi).Value * PI / 180
        XA = sheet3.Range("c" & Xyi).Value
        YA = sheet3.Range("d" & Xyi).Value
        Qz = sheet3.Range("b" & Xyi).Value
        ZZ = sheet3.Range("f" & Xyi).Value + Qz
        THi = Xyi
        Xyi = 4
        XYc = sheet3.Range("b" & Xyi - 1).Value
        pd = 1
        pd1 = 0
        While sheet3.Range("f" & Xyi).Value <> Nothing And pd = 1 And pd1 <= 3 ' And z < XYc + sheet3.Range("f" & Xyi - 1)
            If sheet3.Range("b" & Xyi).Value = Nothing Then
                XYc = XYc + sheet3.Range("f" & Xyi - 1).Value
            Else
                XYc = sheet3.Range("b" & Xyi).Value
            End If
            If Z >= sheet3.Range("b" & Xyi).Value And sheet3.Range("b" & Xyi).Value <> Nothing Then
                AA = sheet3.Range("e" & Xyi).Value * PI / 180
                XA = sheet3.Range("c" & Xyi).Value
                YA = sheet3.Range("d" & Xyi).Value
                Qz = sheet3.Range("b" & Xyi).Value
                ZZ = sheet3.Range("f" & Xyi).Value + Qz
                THi = Xyi
            End If
            If Z < XYc + sheet3.Range("f" & Xyi).Value And st = "*" Then '
                pd = 0
            End If
            If Z < XYc + sheet3.Range("f" & Xyi).Value And st <> "*" And Z <= Z1 Then
                pd = 0
            End If
            If Z < XYc + sheet3.Range("f" & Xyi).Value And st <> "*" And Z > z0 Then
                pd1 = pd1 + 1
            End If
            Xyi = Xyi + 1
        End While
        i = THi
        pd = 1
        While sheet3.Range("f" & i).Value <> Nothing And pd = 1 'While 1
            ''''''''''
            If sheet3.Range("g" & i).Value = Nothing Or sheet3.Range("g" & i).Value = 0 Then
                KA = 0
            Else
                KA = Math.Abs(1 / sheet3.Range("g" & i).Value)
            End If

            If sheet3.Range("h" & i).Value = Nothing Or sheet3.Range("h" & i).Value = 0 Then
                KB = 0
            Else
                KB = Math.Abs(1 / sheet3.Range("h" & i).Value)
            End If

            KAB = KB - KA
            Ls = sheet3.Range("f" & i).Value
            If Val(sheet3.Range("g" & i).Value) = 0 And Val(sheet3.Range("h" & i).Value) = 0 Then
                b = 0
            Else
                If sheet3.Range("g" & i).Value < 0 Or sheet3.Range("h" & i).Value < 0 Then
                    b = -1
                Else
                    b = 1
                End If
            End If

            ''''''''
            If Z >= Qz And Z <= ZZ Or sheet3.Range("f" & i + 1).Value = Nothing Then
                pd = 0
            Else
                L = ZZ - Qz
                b1 = b * (KA * V1 * L ^ 1 + KAB * V1 ^ 2 * L ^ 2 / (2 * Ls))
                b2 = b * (KA * V2 * L ^ 1 + KAB * V2 ^ 2 * L ^ 2 / (2 * Ls))
                b3 = b * (KA * V3 * L ^ 1 + KAB * V3 ^ 2 * L ^ 2 / (2 * Ls))
                b4 = b * (KA * V4 * L ^ 1 + KAB * V4 ^ 2 * L ^ 2 / (2 * Ls))
                b5 = b * (KA * V5 * L ^ 1 + KAB * V5 ^ 2 * L ^ 2 / (2 * Ls))
                XA = XA + L * (R1 * Math.Cos(AA + b1) + R2 * Math.Cos(AA + b2) + R3 * Math.Cos(AA + b3) + R4 * Math.Cos(AA + b4) + R5 * Math.Cos(AA + b5))
                a = AA + b * (KA * L + KAB * L ^ 2 / (2 * Ls))
                AA = a
                Qz = ZZ
                i = i + 1
                ZZ = Qz + sheet3.Range("f" & i).Value
            End If
        End While 'While 1
        '''''''''''''''''
        L = Z - Qz
        a = AA + b * (KA * L + KAB * L ^ 2 / (2 * Ls)) '任意方位角
        b1 = b * (KA * V1 * L ^ 1 + KAB * V1 ^ 2 * L ^ 2 / (2 * Ls))
        b2 = b * (KA * V2 * L ^ 1 + KAB * V2 ^ 2 * L ^ 2 / (2 * Ls))
        b3 = b * (KA * V3 * L ^ 1 + KAB * V3 ^ 2 * L ^ 2 / (2 * Ls))
        b4 = b * (KA * V4 * L ^ 1 + KAB * V4 ^ 2 * L ^ 2 / (2 * Ls))
        b5 = b * (KA * V5 * L ^ 1 + KAB * V5 ^ 2 * L ^ 2 / (2 * Ls))
        XYF_X = XA + L * (R1 * Math.Cos(AA + b1) + R2 * Math.Cos(AA + b2) + R3 * Math.Cos(AA + b3) + R4 * Math.Cos(AA + b4) + R5 * Math.Cos(AA + b5))
        JD = JD * PI / 180
        If BJ < 0 Then
            XYF_X = XYF_X + Math.Abs(BJ) * Math.Cos(a - Math.Abs(JD))
        Else
            XYF_X = XYF_X + BJ * Math.Cos(a + PI - Math.Abs(JD))
        End If
        'XYF_X = Z 'THi ' L * (R1 *math. Cos(aA + b1) + R2 *math. Cos(aA + b2) + R3 *math. Cos(aA + b3) + R4 *math. Cos(aA + b4) + R5 *math. Cos(aA + b5))  ' THi
    End Function
    Public Function XYF_Y(Z, BJ, JD) '线元法断链
        sheet3 = Exbook.Worksheets("线元法") '线元法
        sheet4 = Exbook.Worksheets("断链") '断链
        ''''''交点法
        Dim i12, no, n, x0, y0, st, gs, zi, z0, Z1
        gs = Len(Z)
        st = Mid(Z, 1, 1)
        Z = Val(ExApp.WorksheetFunction.Substitute(Z, "*", ""))
        z0 = DLz(Z) '
        Z1 = z0
        '短断链
        i12 = 3
        no = 1
        zi = 0
        While sheet4.Range("B" & i12).Value <> Nothing And no = 1
            If Z >= sheet4.Range("c" & i12).Value And Z <= z0 And st <> "*" Then
                zi = i12
            End If
            i12 = i12 + 1
        End While
        If zi <> 0 Then
            Z = Z - (sheet4.Range("c" & zi).Value - sheet4.Range("b" & zi).Value)
            Z1 = z0 - (sheet4.Range("c" & zi).Value - sheet4.Range("b" & zi).Value)
        End If
        ''''''''''''''''
        ''''''''交点法
        Dim R1, R2, R3, R4, R5, V1, V2, V3, V4, V5, a, AA, b, Ls, L, KA, KB, KAB, XA, YA
        Dim Qz, ZZ, R_Q, R_Z, i, pd 'ZZ位线元终点桩号,QZ线元起点桩号,R_Q为起点半径, R_Z终点半径
        Dim b1, b2, b3, b4, b5, Xyi, XYc, THi, pd1
        R1 = 0.1184634425
        R5 = R1
        R2 = 0.2393143352
        R4 = R2
        R3 = 0.2844444444
        V1 = 0.04691007
        V5 = 1 - V1
        V2 = 0.2307653449
        V4 = 1 - V2
        V3 = 0.5
        ''寻找
        Xyi = 3
        AA = sheet3.Range("e" & Xyi).Value * PI / 180
        XA = sheet3.Range("c" & Xyi).Value
        YA = sheet3.Range("d" & Xyi).Value
        Qz = sheet3.Range("b" & Xyi).Value
        ZZ = sheet3.Range("f" & Xyi).Value + Qz
        THi = Xyi
        Xyi = 4
        XYc = sheet3.Range("b" & Xyi - 1).Value
        pd = 1
        pd1 = 0
        While sheet3.Range("f" & Xyi).Value <> Nothing And pd = 1 And pd1 <= 3 ' And z < XYc + sheet3.Range("f" & Xyi - 1)
            If sheet3.Range("b" & Xyi).Value = Nothing Then
                XYc = XYc + sheet3.Range("f" & Xyi - 1).Value
            Else
                XYc = sheet3.Range("b" & Xyi).Value
            End If
            If Z >= sheet3.Range("b" & Xyi).Value And sheet3.Range("b" & Xyi).Value <> Nothing Then
                AA = sheet3.Range("e" & Xyi).Value * PI / 180
                XA = sheet3.Range("c" & Xyi).Value
                YA = sheet3.Range("d" & Xyi).Value
                Qz = sheet3.Range("b" & Xyi).Value
                ZZ = sheet3.Range("f" & Xyi).Value + Qz
                THi = Xyi
            End If
            If Z < XYc + sheet3.Range("f" & Xyi).Value And st = "*" Then '
                pd = 0
            End If
            If Z < XYc + sheet3.Range("f" & Xyi).Value And st <> "*" And Z <= Z1 Then
                pd = 0
            End If
            If Z < XYc + sheet3.Range("f" & Xyi).Value And st <> "*" And Z > z0 Then
                pd1 = pd1 + 1
            End If
            Xyi = Xyi + 1
        End While
        i = THi
        pd = 1
        While sheet3.Range("f" & i).Value <> Nothing And pd = 1 'While 1
            ''''''''''
            If sheet3.Range("g" & i).Value = Nothing Or sheet3.Range("g" & i).Value = 0 Then
                KA = 0
            Else
                KA = Math.Abs(1 / sheet3.Range("g" & i).Value)
            End If
            If sheet3.Range("h" & i).Value = Nothing Or sheet3.Range("h" & i).Value = 0 Then
                KB = 0
            Else
                KB = Math.Abs(1 / sheet3.Range("h" & i).Value)
            End If
            KAB = KB - KA
            Ls = sheet3.Range("f" & i).Value
            If Val(sheet3.Range("g" & i).Value) = 0 And Val(sheet3.Range("h" & i).Value) = 0 Then
                b = 0
            Else
                If sheet3.Range("g" & i).Value < 0 Or sheet3.Range("h" & i).Value < 0 Then
                    b = -1
                Else
                    b = 1
                End If
            End If

            ''''''''
            If Z >= Qz And Z <= ZZ Or sheet3.Range("f" & i + 1).Value = Nothing Then
                pd = 0
            Else
                L = ZZ - Qz
                b1 = b * (KA * V1 * L ^ 1 + KAB * V1 ^ 2 * L ^ 2 / (2 * Ls))
                b2 = b * (KA * V2 * L ^ 1 + KAB * V2 ^ 2 * L ^ 2 / (2 * Ls))
                b3 = b * (KA * V3 * L ^ 1 + KAB * V3 ^ 2 * L ^ 2 / (2 * Ls))
                b4 = b * (KA * V4 * L ^ 1 + KAB * V4 ^ 2 * L ^ 2 / (2 * Ls))
                b5 = b * (KA * V5 * L ^ 1 + KAB * V5 ^ 2 * L ^ 2 / (2 * Ls))
                YA = YA + L * (R1 * Math.Sin(AA + b1) + R2 * Math.Sin(AA + b2) + R3 * Math.Sin(AA + b3) + R4 * Math.Sin(AA + b4) + R5 * Math.Sin(AA + b5))
                a = AA + b * (KA * L + KAB * L ^ 2 / (2 * Ls))
                AA = a
                Qz = ZZ
                i = i + 1
                ZZ = Qz + sheet3.Range("f" & i).Value
            End If
        End While 'While 1
        '''''''''''''''''
        L = Z - Qz
        a = AA + b * (KA * L + KAB * L ^ 2 / (2 * Ls)) '任意方位角
        b1 = b * (KA * V1 * L ^ 1 + KAB * V1 ^ 2 * L ^ 2 / (2 * Ls))
        b2 = b * (KA * V2 * L ^ 1 + KAB * V2 ^ 2 * L ^ 2 / (2 * Ls))
        b3 = b * (KA * V3 * L ^ 1 + KAB * V3 ^ 2 * L ^ 2 / (2 * Ls))
        b4 = b * (KA * V4 * L ^ 1 + KAB * V4 ^ 2 * L ^ 2 / (2 * Ls))
        b5 = b * (KA * V5 * L ^ 1 + KAB * V5 ^ 2 * L ^ 2 / (2 * Ls))
        XYF_Y = YA + L * (R1 * Math.Sin(AA + b1) + R2 * Math.Sin(AA + b2) + R3 * Math.Sin(AA + b3) + R4 * Math.Sin(AA + b4) + R5 * Math.Sin(AA + b5))
        JD = JD * PI / 180
        If BJ < 0 Then
            XYF_Y = XYF_Y + Math.Abs(BJ) * Math.Sin(a - Math.Abs(JD))
        Else
            XYF_Y = XYF_Y + BJ * Math.Sin(a + PI - Math.Abs(JD))
        End If
        'XYF_Y = YA
    End Function

    Public Function XYF_a(Z) '线元法，角度,幅度
        sheet3 = Exbook.Worksheets("线元法") '线元法
        sheet4 = Exbook.Worksheets("断链") '断链
        ''''''交点法
        Dim i12, no, n, x0, y0, st, gs, zi, z0, Z1
        gs = Len(Z)
        st = Mid(Z, 1, 1)
        Z = Val(ExApp.WorksheetFunction.Substitute(Z, "*", ""))
        z0 = DLz(Z) '
        Z1 = z0
        '短断链
        i12 = 3
        no = 1
        zi = 0
        While sheet4.Range("B" & i12).Value <> Nothing And no = 1
            If Z >= sheet4.Range("c" & i12).Value And Z <= z0 And st <> "*" Then
                zi = i12
            End If
            i12 = i12 + 1
        End While
        If zi <> 0 Then
            Z = Z - (sheet4.Range("c" & zi).Value - sheet4.Range("b" & zi).Value)
            Z1 = z0 - (sheet4.Range("c" & zi).Value - sheet4.Range("b" & zi).Value)
        End If
        ''''''''''''''''
        ''''''''交点法
        Dim R1, R2, R3, R4, R5, V1, V2, V3, V4, V5, a, AA, b, Ls, L, KA, KB, KAB, XA, YA
        Dim Qz, ZZ, R_Q, R_Z, i, pd 'ZZ位线元终点桩号,QZ线元起点桩号,R_Q为起点半径, R_Z终点半径
        Dim b1, b2, b3, b4, b5, Xyi, XYc, THi, pd1
        R1 = 0.1184634425
        R5 = R1
        R2 = 0.2393143352
        R4 = R2
        R3 = 0.2844444444
        V1 = 0.04691007
        V5 = 1 - V1
        V2 = 0.2307653449
        V4 = 1 - V2
        V3 = 0.5
        ''寻找
        Xyi = 3
        AA = sheet3.Range("e" & Xyi).Value * PI / 180
        XA = sheet3.Range("c" & Xyi).Value
        YA = sheet3.Range("d" & Xyi).Value
        Qz = sheet3.Range("b" & Xyi).Value
        ZZ = sheet3.Range("f" & Xyi).Value + Qz
        THi = Xyi
        Xyi = 4
        XYc = sheet3.Range("b" & Xyi - 1).Value
        pd = 1
        pd1 = 0
        While sheet3.Range("f" & Xyi).Value <> Nothing And pd = 1 And pd1 <= 3 ' And z < XYc + sheet3.Range("f" & Xyi - 1)
            If sheet3.Range("b" & Xyi).Value = Nothing Then
                XYc = XYc + sheet3.Range("f" & Xyi - 1).Value
            Else
                XYc = sheet3.Range("b" & Xyi).Value
            End If
            If Z >= sheet3.Range("b" & Xyi).Value And sheet3.Range("b" & Xyi).Value <> Nothing Then
                AA = sheet3.Range("e" & Xyi).Value * PI / 180
                XA = sheet3.Range("c" & Xyi).Value
                YA = sheet3.Range("d" & Xyi).Value
                Qz = sheet3.Range("b" & Xyi).Value
                ZZ = sheet3.Range("f" & Xyi).Value + Qz
                THi = Xyi
            End If
            If Z < XYc + sheet3.Range("f" & Xyi).Value And st = "*" Then '
                pd = 0
            End If
            If Z < XYc + sheet3.Range("f" & Xyi).Value And st <> "*" And Z <= Z1 Then
                pd = 0
            End If
            If Z < XYc + sheet3.Range("f" & Xyi).Value And st <> "*" And Z > z0 Then
                pd1 = pd1 + 1
            End If
            Xyi = Xyi + 1
        End While
        i = THi
        pd = 1
        While sheet3.Range("f" & i).Value <> Nothing And pd = 1 'While 1
            ''''''''''
            If sheet3.Range("g" & i).Value = Nothing Or sheet3.Range("g" & i).Value = 0 Then
                KA = 0
            Else
                KA = Math.Abs(1 / sheet3.Range("g" & i).Value)
            End If
            If sheet3.Range("h" & i).Value = Nothing Or sheet3.Range("h" & i).Value = 0 Then
                KB = 0
            Else
                KB = Math.Abs(1 / sheet3.Range("h" & i).Value)
            End If
            KAB = KB - KA
            Ls = sheet3.Range("f" & i).Value
            If Val(sheet3.Range("g" & i).Value) = 0 And Val(sheet3.Range("h" & i).Value) = 0 Then
                b = 0
            Else
                If sheet3.Range("g" & i).Value < 0 Or sheet3.Range("h" & i).Value < 0 Then
                    b = -1
                Else
                    b = 1
                End If
            End If

            ''''''''
            If Z >= Qz And Z <= ZZ Or sheet3.Range("f" & i + 1).Value = Nothing Then
                pd = 0
            Else
                L = ZZ - Qz
                b1 = b * (KA * V1 * L ^ 1 + KAB * V1 ^ 2 * L ^ 2 / (2 * Ls))
                b2 = b * (KA * V2 * L ^ 1 + KAB * V2 ^ 2 * L ^ 2 / (2 * Ls))
                b3 = b * (KA * V3 * L ^ 1 + KAB * V3 ^ 2 * L ^ 2 / (2 * Ls))
                b4 = b * (KA * V4 * L ^ 1 + KAB * V4 ^ 2 * L ^ 2 / (2 * Ls))
                b5 = b * (KA * V5 * L ^ 1 + KAB * V5 ^ 2 * L ^ 2 / (2 * Ls))
                YA = YA + L * (R1 * Math.Sin(AA + b1) + R2 * Math.Sin(AA + b2) + R3 * Math.Sin(AA + b3) + R4 * Math.Sin(AA + b4) + R5 * Math.Sin(AA + b5))
                a = AA + b * (KA * L + KAB * L ^ 2 / (2 * Ls))
                AA = a
                Qz = ZZ
                i = i + 1
                ZZ = Qz + sheet3.Range("f" & i).Value
            End If
        End While 'While 1
        '''''''''''''''''
        L = Z - Qz
        a = AA + b * (KA * L + KAB * L ^ 2 / (2 * Ls)) '任意方位角
        b1 = b * (KA * V1 * L ^ 1 + KAB * V1 ^ 2 * L ^ 2 / (2 * Ls))
        b2 = b * (KA * V2 * L ^ 1 + KAB * V2 ^ 2 * L ^ 2 / (2 * Ls))
        b3 = b * (KA * V3 * L ^ 1 + KAB * V3 ^ 2 * L ^ 2 / (2 * Ls))
        b4 = b * (KA * V4 * L ^ 1 + KAB * V4 ^ 2 * L ^ 2 / (2 * Ls))
        b5 = b * (KA * V5 * L ^ 1 + KAB * V5 ^ 2 * L ^ 2 / (2 * Ls))
        'XYF_Y = YA + L * (R1 * math.Sin(aA + b1) + R2 * math.Sin(aA + b2) + R3 * math.Sin(aA + b3) + R4 * math.Sin(aA + b4) + R5 * math.Sin(aA + b5))
        'jd = jd * Pi / 180
        'If BJ < 0 Then
        'XYF_Y = XYF_Y + math.Abs(BJ) * math.Sin(A - math.Abs(jd))
        'Else
        'XYF_Y = XYF_Y + BJ * math.Sin(A + Pi - math.Abs(jd))
        'End If

        XYF_a = a

    End Function
    Public Function Jsfwj_1x(X1, Y1, X2, Y2) As Object  '计算直线方位角。x1、y1为起点坐标。x2、y2为终点方位角
        Dim c
        If X1 = 0 Or Y1 = 0 Or X2 = 0 Or Y2 = 0 Then
            Jsfwj_1x = Nothing
        Else
            If X2 - X1 = 0 Or Y2 - Y1 = 0 Then
                c = 0
            Else
                c = ExApp.WorksheetFunction.Degrees(ExApp.WorksheetFunction.Amath.Tan2(X2 - X1, Y2 - Y1))
            End If
            If c < 0 Then
                Jsfwj_1x = c + 360
            Else : Jsfwj_1x = c
            End If
        End If
    End Function
    Public Function Pd_YSzx(Z, i, k) As Object '进行要素判断，最后得到一要素值。z为桩号，I为第二缓和曲线终点。K为要素值
        Dim n, m, no As Integer
        no = 1
        n = 1
        While i(n) <> Nothing
            n = n + 1
        End While
        If n = 1 Then
            Pd_YSzx = Nothing
        Else
            If Z > i(n - 1) Then
                Pd_YSzx = k(n - 1)
            Else
                n = 1
                While i(m) <> Nothing And no = 1
                    If Z <= i(n) Then
                        Pd_YSzx = k(n)
                        no = 0
                    Else : n = n + 1
                    End If
                End While
            End If
        End If
    End Function
    Public Function Pd_YSwx(Z, i) As Object '进行要素判断，最后得到一要素值的位置。z为桩号，I为第二缓和曲线终点。K为要素值
        Dim n, m, no As Integer
        no = 1
        n = 1
        While i(n) <> Nothing
            n = n + 1
        End While
        If n = 1 Then
            Pd_YSwx = -1
        Else
            If Z > i(n - 1) Then
                Pd_YSwx = n + 3
            Else
                n = 1
                While i(n) <> Nothing And no = 1
                    If Z <= i(n) Then
                        Pd_YSwx = n + 4
                        no = 0
                    Else : n = n + 1
                    End If
                End While
            End If
        End If
    End Function
    Public Function Zhz_z1x(k) As Object '找出最后值。k数组
        Dim i
        i = 2
        While k(i) <> Nothing
            i = i + 1
        End While
        Zhz_z1x = k(i - 1)
    End Function
    Public Function ZSZB_X(Z, BJ, zj)
        sheet3 = Exbook.Worksheets("线元法") '线元法
        If sheet3.Range("j2").Value = "是" Then
            ZSZB_X = XYF_X(Z, BJ, zj) ''''线元法
        Else
            ZSZB_X = ZSZB_X0j(Z, BJ, zj)
            'ZSZB_Y = ZSZB_Y0j(Z, BJ, zj)
            ''''''''交点法
        End If
    End Function
    Public Function ZSZB_Y(Z, BJ, zj)
        sheet3 = Exbook.Worksheets("线元法") '线元法
        If sheet3.Range("j2").Value = "是" Then
            ZSZB_Y = XYF_Y(Z, BJ, zj) ''''线元法
        Else
            'ZSZB_X = ZSZB_X0j(Z, BJ, zj)
            ZSZB_Y = ZSZB_Y0j(Z, BJ, zj)
            ''''''''交点法
        End If
    End Function
    Public Function ZSZB_a(Z)
        sheet3 = Exbook.Worksheets("线元法") '线元法
        If sheet3.Range("j2").Value = "是" Then
            ZSZB_a = XYF_a(Z) ''''线元法
        Else
            ZSZB_a = Zxfwj_B0(Z)
            ''''''''交点法
        End If
    End Function








    Public Function Qx_T(r, Lh1, Lh2, zj) As Object '计算切线长T1。R为半径。Lh1为缓和曲线1长,Lh2为缓和曲线2长,。Zj为交点转角，其单位为度,Wz为点在缓和曲线的位置
        If Val(r) = 0 Then 'Or Val(zj) = 0 Or Val(Lh1) = 0
            Qx_T = 0
        Else
            Dim a, p1, q1, p2, q2, wz
            wz = 1
            p1 = Lh1 ^ 2 / 24 / r - Lh1 ^ 4 / 2688 / r ^ 3 + Lh1 ^ 6 / 506880 / r ^ 5 - Lh1 ^ 8 / 154828800 / r ^ 7
            q1 = Lh1 / 2 - Lh1 ^ 3 / 240 / r ^ 2 + Lh1 ^ 5 / 34560 / r ^ 4 - Lh1 ^ 7 / 8386560 / r ^ 6 + Lh1 ^ 9 / 3158507520.0# / r ^ 8
            a = Math.Abs(ExApp.WorksheetFunction.Radians(zj)) '将度转化成弧度
            p2 = Lh2 ^ 2 / 24 / r - Lh2 ^ 4 / 2688 / r ^ 3 + Lh2 ^ 6 / 506880 / r ^ 5 - Lh2 ^ 8 / 154828800 / r ^ 7
            q2 = Lh2 / 2 - Lh2 ^ 3 / 240 / r ^ 2 + Lh2 ^ 5 / 34560 / r ^ 4 - Lh2 ^ 7 / 8386560 / r ^ 6 + Lh2 ^ 9 / 3158507520.0# / r ^ 8
            If wz = 1 Then
                Qx_T = (r + p2) / Math.Sin(a) - (r + p1) / Math.Tan(a) + q1
            Else
                Qx_T = (r + p1) / Math.Sin(a) - (r + p2) / Math.Tan(a) + q2
            End If
        End If
    End Function
    Public Function Qx_T2(r, Lh1, Lh2, zj) As Object '计算切线长T2。R为半径。Lh1为缓和曲线1长,Lh2为缓和曲线2长,。Zj为交点转角，其单位为度,Wz为点在缓和曲线的位置
        If Val(r) = 0 Then 'Or Val(zj) = 0 Or Val(Lh1) = 0
            Qx_T2 = 0
        Else
            Dim a, p1, q1, p2, q2, wz
            wz = 2
            p1 = Lh1 ^ 2 / 24 / r - Lh1 ^ 4 / 2688 / r ^ 3 + Lh1 ^ 6 / 506880 / r ^ 5 - Lh1 ^ 8 / 154828800 / r ^ 7
            q1 = Lh1 / 2 - Lh1 ^ 3 / 240 / r ^ 2 + Lh1 ^ 5 / 34560 / r ^ 4 - Lh1 ^ 7 / 8386560 / r ^ 6 + Lh1 ^ 9 / 3158507520.0# / r ^ 8
            a = Math.Abs(ExApp.WorksheetFunction.Radians(zj)) '将度转化成弧度
            p2 = Lh2 ^ 2 / 24 / r - Lh2 ^ 4 / 2688 / r ^ 3 + Lh2 ^ 6 / 506880 / r ^ 5 - Lh2 ^ 8 / 154828800 / r ^ 7
            q2 = Lh2 / 2 - Lh2 ^ 3 / 240 / r ^ 2 + Lh2 ^ 5 / 34560 / r ^ 4 - Lh2 ^ 7 / 8386560 / r ^ 6 + Lh2 ^ 9 / 3158507520.0# / r ^ 8
            If wz = 1 Then
                Qx_T2 = (r + p2) / Math.Sin(a) - (r + p1) / Math.Tan(a) + q1
            Else
                Qx_T2 = (r + p1) / Math.Sin(a) - (r + p2) / Math.Tan(a) + q2
            End If
        End If
    End Function
    Public Function Yqx_Ly(r, Lh1, Lh2, zj) As Object '计算圆曲线长。r为半径。Lh为缓和曲线总长。Zj为交点转角，其单位为度
        Dim a
        If Val(r) = 0 Then 'Or Val(zj) = 0 Or Val(Lh1) = 0
            Yqx_Ly = 0
        Else
            a = Math.Abs(ExApp.WorksheetFunction.Radians(zj)) '将度转化成弧度
            Yqx_Ly = r * Math.Abs(a - (Lh1 + Lh2) / 2 / r)
        End If
    End Function

    Public Function Hhquzjfuzb_I(Z, r, Lh1, Lh2, zj, Jd_Z) As Object '计算缓和曲线支距复数坐标。z为计算桩号。r为圆曲线半径。Jd_z为交点桩号。

        Dim L, X, y, ZH, HY, YH, HZ
        ZH = Jd_Z - Qx_T(r, Lh1, Lh2, zj) 'Lh为缓和曲线总长。ZH为直缓点桩号。HY为缓圆点桩号。YH为圆缓点桩号。HZ为缓直点桩号
        HY = ZH + Lh1
        YH = HY + Yqx_Ly(r, Lh1, Lh2, zj)
        HZ = YH + Lh2
        If Z >= ZH And Z <= HY Then
            L = Z - ZH
            X = L - L ^ 5 / 40 / r ^ 2 / Lh1 ^ 2 + L ^ 9 / 3456 / r ^ 4 / Lh1 ^ 4 - L ^ 13 / 599040 / r ^ 6 / Lh1 ^ 6 + L ^ 17 / 175472640 / r ^ 8 / Lh1 ^ 8
            y = L ^ 3 / 6 / r / Lh1 - L ^ 7 / 336 / r ^ 3 / Lh1 ^ 3 + L ^ 11 / 42240 / r ^ 5 / Lh1 ^ 5 - L ^ 15 / 9676800 / r ^ 7 / Lh1 ^ 7 + L ^ 19 / 3530096640.0# / r ^ 18 / Lh1 ^ 18
            Hhquzjfuzb_I = ExApp.WorksheetFunction.Complex(X, y)
        Else
            If Z >= YH And Z <= HZ Then
                L = HZ - Z
                X = L - L ^ 5 / 40 / r ^ 2 / Lh2 ^ 2 + L ^ 9 / 3456 / r ^ 4 / Lh2 ^ 4 - L ^ 13 / 599040 / r ^ 6 / Lh2 ^ 6 + L ^ 17 / 175472640 / r ^ 8 / Lh2 ^ 8
                y = L ^ 3 / 6 / r / Lh2 - L ^ 7 / 336 / r ^ 3 / Lh2 ^ 3 + L ^ 11 / 42240 / r ^ 5 / Lh2 ^ 5 - L ^ 15 / 9676800 / r ^ 7 / Lh2 ^ 7 + L ^ 19 / 3530096640.0# / r ^ 18 / Lh2 ^ 18
                Hhquzjfuzb_I = ExApp.WorksheetFunction.Complex(X, y)
            Else : Hhquzjfuzb_I = Nothing
            End If
        End If
    End Function
    Public Function Xcfwj_A(Z, r, Lh1, Lh2, zj, Jsfwj, Jd_Z) As Object '计算中桩玄长方位角，计算后的单位为弧度。z为计算桩号。r为圆曲线半径。Lh为缓和曲线总长。Jd_z为交点桩号。Zj为交点转角，其单位为度。Jsfwj为直缓点方位角，其单位为度。
        Dim s, q, m, ZH, HY, YH, HZ

        ZH = Jd_Z - Qx_T(r, Lh1, Lh2, zj) 'Lh为缓和曲线总长。ZH为直缓点桩号。HY为缓圆点桩号。YH为圆缓点桩号。HZ为缓直点桩号
        HY = ZH + Lh1
        YH = HY + Yqx_Ly(r, Lh1, Lh2, zj)
        HZ = YH + Lh2
        s = ExApp.WorksheetFunction.Radians(zj) '将度转化成弧度
        q = ExApp.WorksheetFunction.Radians(Jsfwj) '将度转化成弧度
        If zj < 0 Then
            m = -1
        Else : m = 1
        End If
        If Z <= ZH Then
            Xcfwj_A = q
        Else
            If Z < HY Then
                Xcfwj_A = q + m * ExApp.WorksheetFunction.ImArgument(Hhquzjfuzb_I(Z, r, Lh1, Lh2, zj, Jd_Z))
            Else
                If Z <= YH Then
                    Xcfwj_A = q + m * Lh1 / 2 / r + m * (Z - HY) / 2 / r
                Else
                    If Z < HZ Then
                        Xcfwj_A = q + s - m * ExApp.WorksheetFunction.ImArgument(Hhquzjfuzb_I(Z, r, Lh1, Lh2, zj, Jd_Z))
                    Else : Xcfwj_A = q + s
                    End If
                End If
            End If
        End If
    End Function
    Public Function Xc_C(Z, r, Lh1, Lh2, zj, Jd_Z) As Object '计算中桩玄长。z为计算桩号。R为圆曲线半径。Lh为缓和曲线总长。ZH为直缓点桩号。HY为缓圆点桩号。YH为圆缓点桩号。HZ为缓直点桩号
        Dim ZH, HY, YH, HZ
        ZH = Jd_Z - Qx_T(r, Lh1, Lh2, zj) 'Lh为缓和曲线总长。ZH为直缓点桩号。HY为缓圆点桩号。YH为圆缓点桩号。HZ为缓直点桩号
        HY = ZH + Lh1
        YH = HY + Yqx_Ly(r, Lh1, Lh2, zj)
        HZ = YH + Lh2
        If Z <= ZH Then
            Xc_C = ZH - Z '以Zh为起点
        Else
            If Z < HY Then
                Xc_C = ExApp.WorksheetFunction.ImAbs(Hhquzjfuzb_I(Z, r, Lh1, Lh2, zj, Jd_Z))
            Else
                If Z <= YH Then
                    Xc_C = 2 * r * Math.Sin((Z - HY) / 2 / r)
                Else
                    If Z < HZ Then
                        Xc_C = ExApp.WorksheetFunction.ImAbs(Hhquzjfuzb_I(Z, r, Lh1, Lh2, zj, Jd_Z))
                    Else : Xc_C = Z - HZ '以hz为起点
                    End If
                End If
            End If
        End If
    End Function

    Public Function Hyzb_X(r, Lh1, Lh2, zj, Jsfwj, JD_X) As Object '计算缓圆点坐标X，R 为圆曲线半径。Lh为缓和曲线总长。Zj为交点转角，其单位为度。Jsfwj为直缓点方位角，JD_X为交点坐标X
        If r = vbNull Or zj = vbNull Or Jsfwj = vbNull Or JD_X = vbNull Or Lh1 = vbNull Then
            Hyzb_X = Nothing
        Else
            Dim q
            q = ExApp.WorksheetFunction.Radians(Jsfwj) '将度转化成弧度
            If Lh1 = 0 Then
                Hyzb_X = JD_X - Qx_T(r, Lh1, Lh2, zj) * Math.Cos(q)
            Else
                Dim X, y, w, m, s, xc
                X = Lh1 - Lh1 ^ 5 / 40 / r ^ 2 / Lh1 ^ 2 + Lh1 ^ 9 / 3456 / r ^ 4 / Lh1 ^ 4 - Lh1 ^ 13 / 599040 / r ^ 6 / Lh1 ^ 6 + Lh1 ^ 17 / 175472640 / r ^ 8 / Lh1 ^ 8 '缓圆点的支距X
                y = Lh1 ^ 3 / 6 / r / Lh1 - Lh1 ^ 7 / 336 / r ^ 3 / Lh1 ^ 3 + Lh1 ^ 11 / 42240 / r ^ 5 / Lh1 ^ 5 - Lh1 ^ 15 / 9676800 / r ^ 7 / Lh1 ^ 7 + Lh1 ^ 19 / 3530096640.0# / r ^ 18 / Lh1 ^ 18 '缓圆点的支距Y
                w = ExApp.WorksheetFunction.Complex(X, y)
                If zj < 0 Then
                    m = -1
                Else : m = 1
                End If
                s = q + m * ExApp.WorksheetFunction.ImArgument(w)
                xc = ExApp.WorksheetFunction.ImAbs(w)
                Hyzb_X = JD_X - Qx_T(r, Lh1, Lh2, zj) * Math.Cos(q) + xc * Math.Cos(s)
            End If
        End If
    End Function

    Public Function Hyzb_Y(r, Lh1, Lh2, zj, Jsfwj, JD_Y) As Object '计算缓圆点坐标Y，R 为圆曲线半径。Lh为缓和曲线总长。Zj为交点转角，其单位为度。Jsfwj为直缓点方位角，其单位为度。JD_Y为交点坐标Y
        If r = vbNull Or zj = vbNull Or Jsfwj = vbNull Or JD_Y = vbNull Or Lh1 = vbNull Then
            Hyzb_Y = Nothing
        Else
            Dim q
            q = ExApp.WorksheetFunction.Radians(Jsfwj) '将度转化成弧度
            If Lh1 = 0 Then
                Hyzb_Y = JD_Y - Qx_T(r, Lh1, Lh2, zj) * Math.Sin(q)
            Else
                Dim X, y, w, m, s, xc
                X = Lh1 - Lh1 ^ 5 / 40 / r ^ 2 / Lh1 ^ 2 + Lh1 ^ 9 / 3456 / r ^ 4 / Lh1 ^ 4 - Lh1 ^ 13 / 599040 / r ^ 6 / Lh1 ^ 6 + Lh1 ^ 17 / 175472640 / r ^ 8 / Lh1 ^ 8 '缓圆点的支距X
                y = Lh1 ^ 3 / 6 / r / Lh1 - Lh1 ^ 7 / 336 / r ^ 3 / Lh1 ^ 3 + Lh1 ^ 11 / 42240 / r ^ 5 / Lh1 ^ 5 - Lh1 ^ 15 / 9676800 / r ^ 7 / Lh1 ^ 7 + Lh1 ^ 19 / 3530096640.0# / r ^ 18 / Lh1 ^ 18 '缓圆点的支距Y
                w = ExApp.WorksheetFunction.Complex(X, y)
                If zj < 0 Then
                    m = -1
                Else : m = 1
                End If
                s = q + m * ExApp.WorksheetFunction.ImArgument(w)
                xc = ExApp.WorksheetFunction.ImAbs(w)
                Hyzb_Y = JD_Y - Qx_T(r, Lh1, Lh2, zj) * Math.Sin(q) + xc * Math.Sin(s)
            End If
        End If
    End Function

    Public Function Zzzb_X(Z, r, Lh1, Lh2, zj, Jd_Z, JD_X, Jsfwj) As Object '计算中桩X坐标。z为计算桩号。R 为圆曲线半径。Lh为缓和曲线总长。Zj为交点转角，其单位为度。Jsfwj为直缓点方位角，其单位为度。z为计算桩号。ZH为直缓点桩号。HY为缓圆点桩号。YH为圆缓点桩号。HZ为缓直点桩号。jd_X为交点X坐标
        If r = vbNull Or zj = vbNull Or Jsfwj = vbNull Or Lh1 = vbNull Or Jd_Z = vbNull Or JD_X = vbNull Then
            Zzzb_X = Nothing
        Else
            Dim s, q, ZH, HY, YH, HZ, wz
            s = ExApp.WorksheetFunction.Radians(zj) '将度转化成弧度
            q = ExApp.WorksheetFunction.Radians(Jsfwj) '将度转化成弧度
            wz = 1
            ZH = Jd_Z - Qx_T(r, Lh1, Lh2, zj) 'Lh为缓和曲线总长。ZH为直缓点桩号。HY为缓圆点桩号。YH为圆缓点桩号。HZ为缓直点桩号
            HY = ZH + Lh1
            YH = HY + Yqx_Ly(r, Lh1, Lh2, zj)
            HZ = YH + Lh2
            If Z <= ZH Then
                Zzzb_X = JD_X - (Qx_T(r, Lh1, Lh2, zj) + ZH - Z) * Math.Cos(q)
            Else
                If Z < HY Then
                    Zzzb_X = JD_X - Qx_T(r, Lh1, Lh2, zj) * Math.Cos(q) + Xc_C(Z, r, Lh1, Lh2, zj, Jd_Z) * Math.Cos(Xcfwj_A(Z, r, Lh1, Lh2, zj, Jsfwj, Jd_Z))
                Else
                    If Z <= YH Then
                        Zzzb_X = Hyzb_X(r, Lh1, Lh2, zj, Jsfwj, JD_X) + Xc_C(Z, r, Lh1, Lh2, zj, Jd_Z) * Math.Cos(Xcfwj_A(Z, r, Lh1, Lh2, zj, Jsfwj, Jd_Z))
                    Else
                        If Z < HZ Then
                            Zzzb_X = JD_X + Qx_T2(r, Lh1, Lh2, zj) * Math.Cos(q + s) - Xc_C(Z, r, Lh1, Lh2, zj, Jd_Z) * Math.Cos(Xcfwj_A(Z, r, Lh1, Lh2, zj, Jsfwj, Jd_Z))
                        Else
                            Zzzb_X = JD_X + (Qx_T2(r, Lh1, Lh2, zj) + (Z - HZ)) * Math.Cos(q + s)
                        End If
                    End If
                End If
            End If
        End If
    End Function


    Public Function Zzzb_Y(Z, r, Lh1, Lh2, zj, Jd_Z, JD_Y, Jsfwj) As Object  '计算中桩Y坐标。z为计算桩号。R 为圆曲线半径。Lh为缓和曲线总长。Zj为交点转角，其单位为度。Jsfwj为直缓点方位角，其单位为度。z为计算桩号。ZH为直缓点桩号。HY为缓圆点桩号。YH为圆缓点桩号。HZ为缓直点桩号。jd_Y为交点X坐标
        If Z = vbNull Or r = vbNull Or zj = vbNull Or Jsfwj = vbNull Or Lh1 = vbNull Or Jd_Z = vbNull Or JD_Y = vbNull Then
            Zzzb_Y = Nothing
        Else
            Dim s, q, ZH, HY, YH, HZ
            s = ExApp.WorksheetFunction.Radians(zj) '将度转化成弧度
            q = ExApp.WorksheetFunction.Radians(Jsfwj) '将度转化成弧
            ZH = Jd_Z - Qx_T(r, Lh1, Lh2, zj) 'Lh为缓和曲线总长。ZH为直缓点桩号。HY为缓圆点桩号。YH为圆缓点桩号。HZ为缓直点桩号
            HY = ZH + Lh1
            YH = HY + Yqx_Ly(r, Lh1, Lh2, zj)
            HZ = YH + Lh2
            If Z <= ZH Then
                Zzzb_Y = JD_Y - (Qx_T(r, Lh1, Lh2, zj) + ZH - Z) * Math.Sin(q)
            Else
                If Z < HY Then
                    Zzzb_Y = JD_Y - Qx_T(r, Lh1, Lh2, zj) * Math.Sin(q) + Xc_C(Z, r, Lh1, Lh2, zj, Jd_Z) * Math.Sin(Xcfwj_A(Z, r, Lh1, Lh2, zj, Jsfwj, Jd_Z))
                Else
                    If Z <= YH Then
                        Zzzb_Y = Hyzb_Y(r, Lh1, Lh2, zj, Jsfwj, JD_Y) + Xc_C(Z, r, Lh1, Lh2, zj, Jd_Z) * Math.Sin(Xcfwj_A(Z, r, Lh1, Lh2, zj, Jsfwj, Jd_Z))
                    Else
                        If Z < HZ Then
                            Zzzb_Y = JD_Y + Qx_T2(r, Lh1, Lh2, zj) * Math.Sin(q + s) - Xc_C(Z, r, Lh1, Lh2, zj, Jd_Z) * Math.Sin(Xcfwj_A(Z, r, Lh1, Lh2, zj, Jsfwj, Jd_Z))
                        Else
                            Zzzb_Y = JD_Y + (Qx_T2(r, Lh1, Lh2, zj) + (Z - HZ)) * Math.Sin(q + s)
                        End If
                    End If
                End If
            End If
        End If
    End Function
    Public Function Zxfwj_B(Z, r, Lh1, Lh2, zj, Jsfwj, Jd_Z) As Object '计算中桩走向方位角，计算后的单位为度。z为计算桩号。r为圆曲线半径。Lh为缓和曲线总长。Zj为交点转角，其单位为度。Jsfwj为直缓点方位角，其单位为度。z为计算桩号。ZH为直缓点桩号。HY为缓圆点桩号。YH为圆缓点桩号。HZ为缓直点桩号


        If r = vbNull Or zj = vbNull Or Jsfwj = vbNull Or Lh1 = vbNull Or Jd_Z = vbNull Then
            Zxfwj_B = Nothing
        Else
            Dim s, q, ZH, HY, YH, HZ, m
            s = ExApp.WorksheetFunction.Radians(zj) '将度转化成弧度
            q = ExApp.WorksheetFunction.Radians(Jsfwj) '将度转化成弧度
            ZH = Jd_Z - Qx_T(r, Lh1, Lh2, zj) 'Lh为缓和曲线总长。ZH为直缓点桩号。HY为缓圆点桩号。YH为圆缓点桩号。HZ为缓直点桩号
            HY = ZH + Lh1
            YH = HY + Yqx_Ly(r, Lh1, Lh2, zj)
            HZ = YH + Lh2
            If zj < 0 Then
                m = -1
            Else : m = 1
            End If
            If Z <= ZH Then
                Zxfwj_B = Jsfwj
            Else
                If Z < HY Then
                    Zxfwj_B = ExApp.WorksheetFunction.Degrees(q + m * (Z - ZH) ^ 2 / 2 / r / Lh1)
                Else
                    If Z <= YH Then
                        Zxfwj_B = ExApp.WorksheetFunction.Degrees(q + m * Lh1 / 2 / r + m * (Z - HY) / r)
                    Else
                        If Z < HZ Then
                            Zxfwj_B = ExApp.WorksheetFunction.Degrees(q + s - m * (HZ - Z) ^ 2 / 2 / r / Lh2)
                        Else : Zxfwj_B = ExApp.WorksheetFunction.Degrees(q + s)
                        End If
                    End If
                End If
            End If
        End If
    End Function

    Public Function Bzzb_X(Z, r, Lh1, Lh2, zj, Jd_Z, JD_X, Jsfwj, BJ, Zpj, X) As Object '计算边桩坐标x。bj为边桩离中桩的距离。zpj为左边桩与中桩走向方向的夹角。x为中桩坐标X.z为计算桩号。r为圆曲线半径。Lh为缓和曲线总长。Zj为交点转角，其单位为度。Jsfwj为直缓点方位角，其单位为度。z为计算桩号。ZH为直缓点桩号。HY为缓圆点桩号。YH为圆缓点桩号。HZ为缓直点桩号
        Dim w As Single
        If BJ = vbNull Or Zpj = vbNull Or X = vbNull Then
            Bzzb_X = Nothing
        Else
            If BJ < 0 Then

                w = ExApp.WorksheetFunction.Radians(Zxfwj_B(Z, r, Lh1, Lh2, zj, Jsfwj, Jd_Z) - Math.Abs(Zpj))
                Bzzb_X = X + Math.Abs(BJ) * Math.Cos(w)
            Else
                w = ExApp.WorksheetFunction.Radians(Zxfwj_B(Z, r, Lh1, Lh2, zj, Jsfwj, Jd_Z) + 180 - Math.Abs(Zpj))
                Bzzb_X = X + BJ * Math.Cos(w)
            End If
        End If
    End Function
    Public Function Bzzb_Y(Z, r, Lh1, Lh2, zj, Jd_Z, JD_Y, Jsfwj, BJ, Zpj, y) As Object '计算边桩坐标y。bj为边桩离中桩的距离。zpj为左边桩与中桩走向方向的夹角。y为中桩坐标Y.z为计算桩号。r为圆曲线半径。Lh为缓和曲线总长。Zj为交点转角，其单位为度。Jsfwj为直缓点方位角，其单位为度。z为计算桩号。ZH为直缓点桩号。HY为缓圆点桩号。YH为圆缓点桩号。HZ为缓直点桩号
        Dim w
        If BJ = vbNull Or Zpj = vbNull Or y = vbNull Then
            Bzzb_Y = Nothing
        Else
            If BJ < 0 Then

                w = ExApp.WorksheetFunction.Radians(Zxfwj_B(Z, r, Lh1, Lh2, zj, Jsfwj, Jd_Z) - Math.Abs(Zpj))
                Bzzb_Y = y + Math.Abs(BJ) * Math.Sin(w)
            Else
                w = ExApp.WorksheetFunction.Radians(Zxfwj_B(Z, r, Lh1, Lh2, zj, Jsfwj, Jd_Z) + 180 - Math.Abs(Zpj))
                Bzzb_Y = y + BJ * Math.Sin(w)
            End If
        End If
    End Function


    Public Function Jsfwj_1j(X1, Y1, X2, Y2) As Object  '计算直线方位角。x1、y1为起点坐标。x2、y2为终点方位角
        Dim c
        If X1 = 0 Or Y1 = 0 Or X2 = 0 Or Y2 = 0 Then
            Jsfwj_1j = Nothing
        Else
            c = ExApp.WorksheetFunction.Degrees(ExApp.WorksheetFunction.Atan2(X2 - X1, Y2 - Y1))
            If c < 0 Then
                Jsfwj_1j = c + 360
            Else : Jsfwj_1j = c
            End If
        End If
    End Function
    Public Function Pd_YSz(Z, i, k) As Object '进行要素判断，最后得到一要素值。z为桩号，I为第二缓和曲线终点。K为要素值
        Dim n, m, no As Integer
        no = 1
        n = 1
        While i(n, 1) <> Nothing
            n = n + 1
        End While
        If n = 1 Then
            Pd_YSz = Nothing
        Else
            If Z > i(n - 1, 1) Then
                Pd_YSz = k(n - 1)
            Else
                n = 1
                While i(m) <> Nothing And no = 1
                    If Z <= i(n, 1) Then
                        Pd_YSz = k(n)
                        no = 0
                    Else : n = n + 1
                    End If
                End While
            End If
        End If
    End Function
    Public Function Pd_YSw(Z, i) As Object '进行要素判断，最后得到一要素值的位置。z为桩号，I为第二缓和曲线终点。K为要素值
        Dim n, m, no As Integer
        no = 1
        n = 1
        While i(n, 1) <> Nothing
            n = n + 1
        End While
        If n = 1 Then
            Pd_YSw = -1
        Else
            If Z > i(n - 1, 1) Then
                Pd_YSw = n + 3
            Else
                n = 1
                While i(n, 1) <> Nothing And no = 1
                    If Z <= i(n, 1) Then
                        Pd_YSw = n + 4
                        no = 0
                    Else : n = n + 1
                    End If
                End While
            End If
        End If
    End Function
    Public Function Pd_YSwD(Z, i) As Object '进行要素判断，最后得到一要素值的位置。z为桩号，I为第二缓和曲线终点。K为要素值
        Dim n, m, no As Integer
        no = 1
        n = 1
        While i(n, 1) <> Nothing
            n = n + 1
        End While
        If n = 1 Then
            Pd_YSwD = -1
        Else
            If Z > i(n - 1, 1) Then
                Pd_YSwD = n + 3
            Else
                n = 1
                While i(n, 1) <> Nothing And no <= 2
                    If n = 1 Then
                        If Z <= i(n, 1) Then
                            Pd_YSwD = n + 4
                            no = no + 1
                        End If
                    Else
                        If Z <= i(n, 1) And Z > i(n - 1, 1) Then
                            Pd_YSwD = n + 4
                            no = no + 1
                        End If
                    End If
                    n = n + 1
                End While
            End If
        End If
    End Function
    Public Function Zhz_z1(k) As Object '找出最后值。k数组
        Dim i
        i = 2
        While k(i) <> Nothing
            i = i + 1
        End While
        Zhz_z1 = k(i - 1)
    End Function
    Public Function Zhz_z2(k2) As Object '找出最后第二值。k2数组
        Dim i
        i = 2
        While k2(i) <> Nothing
            i = i + 1
        End While
        Zhz_z2 = k2(i - 1)
    End Function

    '交点法在计算坐标时，在长链前桩号前加"*"
    Public Function ZSZB_X0j(Z, BJ, zj) '计算坐标X,Z桩号,BJ偏距
        sheet2 = Exbook.Worksheets("交点法") '交点法
        sheet4 = Exbook.Worksheets("断链") '断链
        ''''''交点法
        Dim i12, no, n, x0, y0, st, gs, DLc
        gs = Len(Z)
        st = Mid(Z, 1, 1)
        Z = Val(ExApp.WorksheetFunction.Substitute(Z, "*", ""))
        If st <> "*" Then
            n = Pd_YSwD(Z, sheet2.Range("o5 : o500").Value)   '对应桩号所在线元的位置
            i12 = 3
            no = 1
            While sheet4.Range("B" & i12).Value <> Nothing And no = 1
                If sheet4.Range("c" & i12).Value <= Z And sheet4.Range("B" & i12).Value > Z Then
                    n = Pd_YSw(sheet4.Range("B" & i12).Value, sheet2.Range("o5 : o500").Value)   '对应桩号所在线元的位置
                    no = 0
                End If
                i12 = i12 + 1
            End While
        Else
            n = Pd_YSw(Z, sheet2.Range("o5 : o500").Value)  '对应桩号所在线元的位置
        End If
        i12 = 3
        no = 1
        DLc = 0
        While sheet4.Range("B" & i12).Value <> Nothing And no = 1
            If sheet4.Range("B" & i12).Value < sheet2.Range("o" & n).Value And sheet4.Range("B" & i12).Value > sheet2.Range("o" & n - 1).Value Then
                DLc = (sheet4.Range("C" & i12).Value - sheet4.Range("B" & i12).Value)
                If Z > sheet4.Range("c" & i12).Value Then
                    Z = Z - DLc
                End If
                no = 0
            End If
            i12 = i12 + 1
        End While
        ''''''''交点法
        x0 = Zzzb_X(Z, Val(sheet2.Range("E" & n).Value), Val(sheet2.Range("F" & n).Value), Val(sheet2.Range("G" & n).Value), Val(sheet2.Range("q" & n).Value), Val(sheet2.Range("d" & n).Value) - DLc, Val(sheet2.Range("b" & n).Value), Val(sheet2.Range("p" & n).Value))
        y0 = Zzzb_Y(Z, Val(sheet2.Range("E" & n).Value), Val(sheet2.Range("F" & n).Value), Val(sheet2.Range("G" & n).Value), Val(sheet2.Range("q" & n).Value), Val(sheet2.Range("d" & n).Value) - DLc, Val(sheet2.Range("c" & n).Value), Val(sheet2.Range("p" & n).Value))
        ZSZB_X0j = Bzzb_X(Z, Val(sheet2.Range("E" & n).Value), Val(sheet2.Range("F" & n).Value), Val(sheet2.Range("G" & n).Value), Val(sheet2.Range("q" & n).Value), Val(sheet2.Range("d" & n).Value) - DLc, Val(sheet2.Range("b" & n).Value), Val(sheet2.Range("p" & n).Value), BJ, zj, x0)
        'ZSZB_Y0j = Bzzb_Y(Z, Val(sheet2.Range("E" & n)), Val(sheet2.Range("F" & n)), Val(sheet2.Range("G" & n)), Val(sheet2.Range("q" & n)), Val(sheet2.Range("d" & n)) - DLc, Val(sheet2.Range("c" & n)), Val(sheet2.Range("p" & n)), BJ, zj, y0)
    End Function
    '交点法在计算坐标时，在长链前桩号前加"*"
    Public Function ZSZB_Y0j(Z, BJ, zj) '计算坐标X,Z桩号,BJ偏距
        sheet2 = Exbook.Worksheets("交点法") '交点法
        sheet4 = Exbook.Worksheets("断链") '断链
        ''''''交点法
        Dim i12, no, n, x0, y0, st, gs, DLc
        gs = Len(Z)
        st = Mid(Z, 1, 1)
        Z = Val(ExApp.WorksheetFunction.Substitute(Z, "*", ""))
        If st <> "*" Then
            n = Pd_YSwD(Z, sheet2.Range("o5 : o328").Value)   '对应桩号所在线元的位置
            i12 = 3
            no = 1
            While sheet4.Range("B" & i12).Value <> Nothing And no = 1
                If sheet4.Range("c" & i12).Value <= Z And sheet4.Range("B" & i12).Value > Z Then
                    n = Pd_YSw(sheet4.Range("B" & i12).Value, sheet2.Range("o5 : o328").Value)   '对应桩号所在线元的位置
                    no = 0
                End If
                i12 = i12 + 1
            End While
        Else
            n = Pd_YSw(Z, sheet2.Range("o5 : o328").Value)  '对应桩号所在线元的位置
        End If
        i12 = 3
        no = 1
        DLc = 0
        While sheet4.Range("B" & i12).Value <> Nothing And no = 1
            If sheet4.Range("B" & i12).Value < sheet2.Range("o" & n).Value And sheet4.Range("B" & i12).Value > sheet2.Range("o" & n - 1).Value Then
                DLc = (sheet4.Range("C" & i12).Value - sheet4.Range("B" & i12).Value)
                If Z > sheet4.Range("C" & i12).Value Then
                    Z = Z - DLc
                End If
                no = 0
            End If
            i12 = i12 + 1
        End While
        ''''''''交点法
        x0 = Zzzb_X(Z, Val(sheet2.Range("E" & n).Value), Val(sheet2.Range("F" & n).Value), Val(sheet2.Range("G" & n).Value), Val(sheet2.Range("q" & n).Value), Val(sheet2.Range("d" & n).Value) - DLc, Val(sheet2.Range("b" & n).Value), Val(sheet2.Range("p" & n).Value))
        y0 = Zzzb_Y(Z, Val(sheet2.Range("E" & n).Value), Val(sheet2.Range("F" & n).Value), Val(sheet2.Range("G" & n).Value), Val(sheet2.Range("q" & n).Value), Val(sheet2.Range("d" & n).Value) - DLc, Val(sheet2.Range("c" & n).Value), Val(sheet2.Range("p" & n).Value))
        'ZSZB_X0j = Bzzb_X(Z, Val(sheet2.Range("E" & n)), Val(sheet2.Range("F" & n)), Val(sheet2.Range("G" & n)), Val(sheet2.Range("q" & n)), Val(sheet2.Range("d" & n))-dlc, Val(sheet2.Range("b" & n)),Val( sheet2.Range("p" & n)), BJ, zj, x0)
        ZSZB_Y0j = Bzzb_Y(Z, Val(sheet2.Range("E" & n).Value), Val(sheet2.Range("F" & n).Value), Val(sheet2.Range("G" & n).Value), Val(sheet2.Range("q" & n).Value), Val(sheet2.Range("d" & n).Value) - DLc, Val(sheet2.Range("c" & n).Value), Val(sheet2.Range("p" & n).Value), BJ, zj, y0)

    End Function

    Public Function Zxfwj_B0(Z) '计算中桩走向方位角
        sheet2 = Exbook.Worksheets("交点法") '交点法
        sheet4 = Exbook.Worksheets("断链") '断链
        ''''''交点法
        Dim i12, no, n, x0, y0, st, gs, DLc
        gs = Len(Z)
        st = Mid(Z, 1, 1)
        Z = Val(ExApp.WorksheetFunction.Substitute(Z, "*", ""))
        If st <> "*" Then
            n = Pd_YSwD(Z, sheet2.Range("o5 : o500").Value)   '对应桩号所在线元的位置
            i12 = 3
            no = 1
            While sheet4.Range("B" & i12).Value <> Nothing And no = 1
                If sheet4.Range("c" & i12).Value <= Z And sheet4.Range("B" & i12).Value > Z Then
                    n = Pd_YSw(sheet4.Range("B" & i12).Value, sheet2.Range("o5 : o500").Value)   '对应桩号所在线元的位置
                    no = 0
                End If
                i12 = i12 + 1
            End While
        Else
            n = Pd_YSw(Z, sheet2.Range("o5 : o500").Value)  '对应桩号所在线元的位置
        End If
        i12 = 3
        no = 1
        DLc = 0
        While sheet4.Range("B" & i12).Value <> Nothing And no = 1
            If sheet4.Range("B" & i12).Value < sheet2.Range("o" & n).Value And sheet4.Range("B" & i12).Value > sheet2.Range("o" & n - 1).Value Then
                DLc = (sheet4.Range("C" & i12).Value - sheet4.Range("B" & i12).Value)
                If Z > sheet4.Range("c" & i12).Value Then
                    Z = Z - DLc
                End If
                no = 0
            End If
            i12 = i12 + 1
        End While
        ''''''''交点法
        'x0 = Zzzb_X(Z, Val(sheet2.Range("E" & n)), Val(sheet2.Range("F" & n)), Val(sheet2.Range("G" & n)), Val(sheet2.Range("q" & n)), Val(sheet2.Range("d" & n)) - DLc, Val(sheet2.Range("b" & n)), Val(sheet2.Range("p" & n)))
        'y0 = Zzzb_Y(Z, Val(sheet2.Range("E" & n)), Val(sheet2.Range("F" & n)), Val(sheet2.Range("G" & n)), Val(sheet2.Range("q" & n)), Val(sheet2.Range("d" & n)) - DLc, Val(sheet2.Range("c" & n)), Val(sheet2.Range("p" & n)))
        'ZSZB_X0j = Bzzb_X(Z, Val(sheet2.Range("E" & n)), Val(sheet2.Range("F" & n)), Val(sheet2.Range("G" & n)), Val(sheet2.Range("q" & n)), Val(sheet2.Range("d" & n))-dlc, Val(sheet2.Range("b" & n)),Val( sheet2.Range("p" & n)), BJ, zj, x0)
        'ZSZB_Y0j = Bzzb_Y(Z, Val(sheet2.Range("E" & n)), Val(sheet2.Range("F" & n)), Val(sheet2.Range("G" & n)), Val(sheet2.Range("q" & n)), Val(sheet2.Range("d" & n)) - DLc, Val(sheet2.Range("c" & n)), Val(sheet2.Range("p" & n)), BJ, zj, y0)
        Zxfwj_B0 = Zxfwj_B(Z, Val(sheet2.Range("E" & n).Value), Val(sheet2.Range("F" & n).Value), Val(sheet2.Range("G" & n).Value), Val(sheet2.Range("q" & n).Value), Val(sheet2.Range("p" & n).Value), Val(sheet2.Range("d" & n).Value))
    End Function



End Module
