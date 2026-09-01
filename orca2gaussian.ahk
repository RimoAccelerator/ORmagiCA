#Requires AutoHotkey v2.0
#SingleInstance Force
#ErrorStdOut

; =====================================================================
; orca2gaussian.ahk
; 将 ORCA 输出文件(.out)转换为 GaussView 可完美读取的伪 Gaussian 输出
; (_fake.log)。模仿 OfakeG 的输出骨架，并补齐其不足：
;   - 扫描：可选保留每步最后一帧，或所有帧(opt 轨迹)，并标注扫描坐标值
;   - 约束(%geom Constraints)：以 ModRedundant 风格回显到日志头部
;   - TD-DFT：每个激发态写成 Gaussian 的 "Excited State" 行，
;     GaussView 可直接生成 UV-Vis 光谱；TD 优化则每步都带光谱
;   - 构型优化：每一帧几何/能量/收敛判据 + 末次振动分析 + 热化学
; =====================================================================

global g := {}

OnError(HandleError, 1)

HandleError(Exception, Mode) {
    FileAppend("RUNTIME_ERROR: " . Exception.Message . " @line " . Exception.Line . "`n", A_Temp . "\orca2g_dbg.txt")
}

; 读取 ORmagiCA_settings.ini 的 [General] Language 设置 (zh/en)，默认 zh
GetUiLang() {
    try {
        iniPath := A_ScriptDir "\ORmagiCA_settings.ini"
        if FileExist(iniPath) {
            v := IniRead(iniPath, "General", "Language", "zh")
            return (v = "en") ? "en" : "zh"
        }
    }
    return "zh"
}

Main()

Main() {
    global g
    if (A_Args.Length >= 1) {
        ; 命令行模式: orca2gaussian.ahk <file.out> [all|last] [silent]
        f := A_Args[1]
        outFile := RegExReplace(f, "\.out$") . "_fake.log"
        if (A_Args.Length >= 2 && RegExMatch(f, "i)scan")) {
            m := StrLower(A_Args[2])
            if (m = "all")
                outFile := RegExReplace(f, "\.out$") . "_scanall_fake.log"
            else if (m = "last")
                outFile := RegExReplace(f, "\.out$") . "_scanlast_fake.log"
        }
        Dbg("[1] parsing " . f)
        ParseORCA(f)
        Dbg("[2] calc=" . g.calc . " frames=" . g.frames.Length . " natoms=" . g.natoms . " nbasis=" . g.nbasis . " q=" . g.charge . " m=" . g.mult)
        if (A_Args.Length >= 2 && g.calc = "SCAN")
            g.scanMode := StrLower(A_Args[2])
        try
            AskAndWrite(outFile, f, true)
        catch Error as err
            Dbg("ERR: " . err.Message . " @line " . err.Line)
        ExitApp
    }
    lang := GetUiLang()
    f := FileSelect(3, , (lang = "en") ? "Select ORCA output file" : "选择 ORCA 输出文件", (lang = "en") ? "ORCA output (*.out; *.log)" : "ORCA 输出 (*.out; *.log)")
    if !f
        ExitApp
    outFile := RegExReplace(f, "\.out$") . "_fake.log"
    ParseORCA(f)
    AskAndWrite(outFile, f, false)
}

; ---------------------------------------------------------------
; 解析
; ---------------------------------------------------------------
ParseORCA(path) {
    global g
    g := { natoms:0, syms:[], nums:[], charge:0, mult:1
         , calc:"SP", isTS:false, isTD:false, hasFreq:false
         , nbasis:0, frames:[], curFrame:0
         , constraints:[], scanDef:"", scanVal:""
         , freqs:[], irInt:[], redMass:[], raman:[], modes:[]
         , thermo:{}, tdRoot:1
         , kwLine:"", functional:"", basis:"", solvent:"", nroots:0, normalEnd:false
         , mulC:[], mulS:[], hasMulSpin:false
         , lowC:[], lowS:[], hasLowSpin:false
         , nprocs:0, maxcore:0, kwVerbatim:[] }

    lines := StrSplit(FileRead(path), "`n", "`r")
    g.lines := lines
    n := lines.Length

    ; ---------- 逐行扫描 ----------
    i := 0
    inGeomBlock := false
    inConstraints := false
    inAbsorb := false
    absorbIdx := 0

    While (i < n) {
        i++
        L := lines[i]

        ; ---- 输入回显区 (仅当行以 | 开头才匹配) ----
        if (SubStr(L, 1, 1) = "|" && RegExMatch(L, "^\|\s*\d+>\s*(.*)$", &m)) {
            t := m[1]
            ; TD-DFT/CIS 关键字(%tddft 块)
            if RegExMatch(t, "i)(%tddft|\btddft\b)")
                g.isTD := true
            if RegExMatch(t, "i)nroots\s+(\d+)", &nr)
                g.nroots := Integer(nr[1])
            if RegExMatch(t, "^\s*!\s*(.+)$", &kwraw) {
                g.kwLine := Trim(kwraw[1])
                toks := StrSplit(g.kwLine, " ")
                For tk in toks {
                    tlk := StrLower(tk)
                    if (tlk = "" || tlk = "!")
                        continue
                    ; 含 / 或 ( 的 token (辅基/溶剂/带括号) 一律不当作 functional/basis
                    if (InStr(tk, "/") || InStr(tk, "(")) {
                        if (tlk = "scan")
                            g.functional := tk
                        else if (tlk != StrLower(g.functional) && tlk != StrLower(g.basis))
                            g.kwVerbatim.Push(tk)
                        continue
                    }
                    if (g.functional = "" && (RegExMatch(tlk, "^(b3lyp|pbe0|pbe|tpss|m06|wb97|revpbe|otpss|bp86|blyp|olyp|pw91|b97|hf|mp2|ccsd|dlpno|sos-|soss-|r2scan)") || tlk = "scan"))
                        g.functional := tk
                    else if (g.basis = "" && RegExMatch(tlk, "^(def2|cc-p|cc-pv|aug-|sto-|6-3|lanl|sarc|dz|tz|qz|svp|tzv|qzv|pc-|ma)") && !InStr(tlk, "/"))
                        g.basis := tk
                    else if RegExMatch(tlk, "^(smd|cpcm)\((.+)\)", &sv) {
                        g.solvent := sv[2]
                        ; solvent 也要逐字保留
                        if (tlk != StrLower(g.functional) && tlk != StrLower(g.basis))
                            g.kwVerbatim.Push(tk)
                    }
                    else {
                        ; 逐字保留非 functional/basis 的关键词 (optts/opt/freq/scants/TD...)
                        if (tlk != StrLower(g.functional) && tlk != StrLower(g.basis))
                            g.kwVerbatim.Push(tk)
                    }
                }
            }
            if RegExMatch(t, "^\!") {
                kw := StrLower(t)
                if RegExMatch(kw, "\boptts\b")
                    g.isTS := true, g.calc := "OPT"
                else if RegExMatch(kw, "(^|\s)(scants|scan|scan_ts)(\s|$)")
                    g.calc := "SCAN"
                else if RegExMatch(kw, "(^|\s)opt(\s|$)|\bcopt\b|\bopt\b")
                    g.calc := "OPT"
                if RegExMatch(kw, "(^|\s)(freq|numfreq)(\s|$)")
                    g.hasFreq := true
                if RegExMatch(kw, "(^|\s)(td|cis|steom|rocis)(\s|$)")
                    g.isTD := true
            }
            ; %geom 块内的约束与扫描定义
            if RegExMatch(t, "i)^\s*Constraints\s*$")
                inConstraints := true
            else if (inConstraints && RegExMatch(t, "i)^\s*end\s*$"))
                inConstraints := false
            else if inConstraints && RegExMatch(t, "i)^\s*\{\s*([BADCB])\s+(.*?)\s*\}\s*$", &c) {
                typ := StrUpper(c[1])
                rest := Trim(c[2])
                idxs := []
                val := ""
                wildcard := false
                For tok in StrSplit(rest, " ") {
                    tk := Trim(tok)
                    if (tk = "")
                        continue
                    if RegExMatch(tk, "^[\d\.]+$") {
                        if (val = "" && RegExMatch(tk, "^\d+$") && idxs.Length < 4 && !InStr(tk, "."))
                            idxs.Push(Integer(tk))
                        else
                            val := tk
                    } else if (tk ~= "^[A-Z]$") {
                        ; 结尾的 C 标志忽略（约束即冻结）
                    } else {
                        wildcard := true
                    }
                }
                g.constraints.Push({typ:typ, idx:idxs, val:val, wild:wildcard})
            }
            ; 扫描定义: Scan B 6 19 = 2.5, 1.8, 8
            if RegExMatch(t, "i)^\s*Scan\s+([BAD])\s+([\d\s]+)=\s*(-?[\d\.]+)\s*,\s*(-?[\d\.]+)\s*,\s*(\d+)", &s)
                g.scanDef := {typ:StrUpper(s[1]), atoms:Trim(s[2]), start:s[3], stop:s[4], steps:s[5]}
        }

        ; ---- 电荷/多重度 ----
        if (InStr(L, "xyz") && RegExMatch(L, "i)\*\s*xyz\s+(-?\d+)\s+(\d+)", &q)) {
            g.charge := Integer(q[1])
            g.mult := Integer(q[2])
        }
        if !g.nbasis && RegExMatch(L, "Basis Dimension\s+Dim\s+\.*\s+(\d+)", &b)
            g.nbasis := Integer(b[1])

        ; ---- nprocs / memory (%pal nprocs, %maxcore) ----
        if (InStr(L, "%pal") && RegExMatch(L, "i)%pal\s+nprocs\s+(\d+)", &np_))
            g.nprocs := Integer(np_[1])
        if (InStr(L, "%maxcore") && RegExMatch(L, "i)%maxcore\s+(\d+)", &mc_))
            g.maxcore := Integer(mc_[1])

        ; ---- 原子列表(第一次笛卡尔坐标) ----
        if (!g.natoms && InStr(L, "CARTESIAN COORDINATES (ANGSTROEM)")) {
            j := i
            ; 跳过标题下方的分隔线/空行
            While (++j <= n) {
                tl := Trim(lines[j])
                if (tl != "" && SubStr(tl, 1, 2) != "--")
                    break
            }
            While (j <= n) {
                tl := Trim(lines[j])
                if (tl = "" || SubStr(tl, 1, 2) = "--")
                    break
                if RegExMatch(tl, "^([A-Za-z]{1,3})\s+(-?[\d\.]+)\s+(-?[\d\.]+)\s+(-?[\d\.]+)$", &m) {
                    g.natoms++
                    g.syms.Push(m[1])
                    g.nums.Push(AtomNum(m[1]))
                }
                j++
            }
        }

        ; ---- 新优化帧 ----
        isNewFrame := false
        if (InStr(L, "GEOMETRY OPTIMIZATION CYCLE") && RegExMatch(L, "GEOMETRY OPTIMIZATION CYCLE\s+(\d+)", &cy)) {
            isNewFrame := true
        } else if (InStr(L, "RELAXED SURFACE SCAN STEP") && RegExMatch(L, "RELAXED SURFACE SCAN STEP\s+(\d+)", &sc)) {
            isNewFrame := true
            g.curScanStep := Integer(sc[1])
            ; 向后找扫描坐标值行
            k := i
            While (k < i + 10 && k < n) {
                k++
                if RegExMatch(lines[k], "(Bond|Angle|Dihedral)\s*\(([^)]*)\)\s*:\s*(-?[\d\.]+)", &sv) {
                    g.scanVal := sv[1] . "(" . sv[2] . ")=" . sv[3]
                    break
                }
            }
        }
        if isNewFrame {
            g.frames.Push({geom:[], energy:"", etot:"", conv:{eChg:"", rmsG:"", maxG:"", rmsS:"", maxS:""}
                         , yesE:0, yesRG:0, yesMG:0, yesRS:0, yesMS:0
                         , converged:false, td:[], scanStep:(g.HasOwnProp("curScanStep") ? g.curScanStep : 0)
                         , scanVal:g.scanVal})
            g.curFrame := g.frames.Length
            absorbIdx := 0
            continue
        }
        if (g.curFrame = 0) {
            ; SP/TD 单点没有帧头，也允许收集
            g.frames.Push({geom:[], energy:"", etot:"", conv:{eChg:"", rmsG:"", maxG:"", rmsS:"", maxS:""}
                         , yesE:0, yesRG:0, yesMG:0, yesRS:0, yesMS:0
                         , converged:true, td:[], scanStep:0})
            g.curFrame := 1
        }

        fr := (g.curFrame > 0) ? g.frames[g.curFrame] : 0

        ; ---- 当前帧几何 ----
        if (fr && InStr(L, "CARTESIAN COORDINATES (ANGSTROEM)") && g.natoms) {
            fr.geom := []
            j := i
            cnt := 0
            While (cnt < g.natoms && ++j <= n) {
                tl := Trim(lines[j])
                if (tl = "" || SubStr(tl, 1, 2) = "--")
                    continue
                if RegExMatch(tl, "^([A-Za-z]{1,3})\s+(-?[\d\.]+)\s+(-?[\d\.]+)\s+(-?[\d\.]+)$", &m2)
                    fr.geom.Push([m2[1], m2[2]+0, m2[3]+0, m2[4]+0]), cnt++
                else if (cnt > 0)
                    break
            }
        }

        ; ---- 能量 ----
        if (fr) {
            if (InStr(L, "FINAL SINGLE POINT ENERGY") && RegExMatch(L, "FINAL SINGLE POINT ENERGY\s+(-?\d+\.?\d*)", &e))
                fr.energy := e[1]+0
            if (InStr(L, "E(tot)") && RegExMatch(L, "E\(tot\)\s+=\s+(-?\d+\.\d+)", &et))
                fr.etot := et[1]+0
            if (g.isTD && InStr(L, "DE(CIS)") && RegExMatch(L, "DE\(CIS\)\s+=\s+-?\d+\.\d+\s+Eh\s+\(Root\s+(\d+)\)", &rt))
                g.tdRoot := Integer(rt[1])
        }

        ; ---- 收敛判据 ----
        if (fr && (InStr(L, "gradient") || InStr(L, "Energy change") || InStr(L, "step")) && RegExMatch(L, "i)\s(RMS gradient|MAX gradient|RMS step|MAX step|Energy change)\s+(-?[\d\.]+)\s+([\d\.]+)\s+(YES|NO)", &cv)) {
            key := cv[1]
            v := cv[2]+0
            yn := (cv[4]="YES") ? 1 : 0
            switch key {
                case "RMS gradient": fr.conv.rmsG := v, fr.yesRG := yn
                case "MAX gradient": fr.conv.maxG := v, fr.yesMG := yn
                case "RMS step":     fr.conv.rmsS := v, fr.yesRS := yn
                case "MAX step":     fr.conv.maxS := v, fr.yesMS := yn
                case "Energy change": fr.conv.eChg := v, fr.yesE := yn
            }
        }

        ; ---- 收敛标志 ----
        if InStr(L, "THE OPTIMIZATION HAS CONVERGED") && (fr)
            fr.converged := true

        ; ---- TD 激发态 ----
        if (fr && InStr(L, "STATE") && InStr(L, "au") && RegExMatch(L, "STATE\s+(\d+):\s+E=\s+(-?\d+\.?\d*)\s+au\s+(-?\d+\.?\d*)\s*eV", &st)) {
            mult2 := ""
            s2v := ""
            if RegExMatch(L, "<S\*\*2>\s+=\s+([\d\.]+)", &ss) {
                s2v := ss[1]+0
                mult2 := Round(s2v + 0.5)
            }
            stIdx := Integer(st[1])
            ; 去重：ORCA TD-opt 同一结构会输出两遍激发态(一遍带 f, 一遍 f=0)，
            ; 若同 idx 已存在则跳过第二个，保留第一个(带 f 的)。
            dup := false
            For rec in fr.td
                if (rec.idx = stIdx) {
                    dup := true
                    break
                }
            if (!dup)
                fr.td.Push({idx:stIdx, ev:st[3]+0, au:st[2]+0, fosc:0
                           , s2:s2v, multW:mult2, trans:[]})
        }
        if (fr && fr.td.Length && InStr(L, "->") && RegExMatch(L, "^\s*(\d+[ab]?)\s*->\s*(\d+[ab]?)\s*:\s*([\d\.]+)(?:\s*\(c=\s*(-?[\d\.]+)\))?", &tr))
            fr.td[fr.td.Length].trans.Push([tr[1], tr[2], tr[3]+0, (tr[4] != "") ? tr[4]+0 : tr[3]+0])

        ; ---- 吸收谱表(取 fosc) ----
        if InStr(L, "ABSORPTION SPECTRUM VIA TRANSITION ELECTRIC DIPOLE MOMENTS") {
            inAbsorb := true
            absorbIdx := 0
            continue
        }
        if inAbsorb {
            if RegExMatch(L, "^\s*-+\s*$") || InStr(L, "CD SPECTRUM") || InStr(L, "VELOCITY DIPOLE")
                (InStr(L, "VELOCITY DIPOLE") || InStr(L, "CD SPECTRUM")) ? (inAbsorb := false) : 0
            else if RegExMatch(L, "^\s*(\d+)\s*-\s*\S+\s*->\s*(\d+)\s*-\s*\S+\s+(-?\d+\.?\d*)\s+(-?\d+\.?\d*)\s+(\d+\.?\d*)\s+([\d\.]+)", &ar) {
                if (fr && fr.td.Length) {
                    absorbIdx++
                    if (absorbIdx <= fr.td.Length) {
                        fr.td[absorbIdx].ev := ar[3]+0
                        fr.td[absorbIdx].nm := ar[5]+0
                        fr.td[absorbIdx].fosc := ar[6]+0
                    }
                }
            }
        }

; ---- 频率 ----
        if InStr(L, "VIBRATIONAL FREQUENCIES") {
            g.freqs := []
            g.irInt := []
            g.redMass := []
            g.raman := []
            g.modes := []
            j := i
            While (++j <= n) {
                if RegExMatch(lines[j], "^\s*(\d+):\s*(-?[\d\.]+)\s*cm", &fm) {
                    g.freqs.Push(fm[2]+0)
                    g.irInt.Push(0)
                    g.redMass.Push(0)
                    g.raman.Push("")
                    g.modes.Push([])
                }
                if InStr(lines[j], "NORMAL MODES")
                    break
            }
        }
        if (g.freqs.Length && InStr(L, "IR SPECTRUM")) {
            j := i
            While (++j <= n) {
                if RegExMatch(lines[j], "^\s*(\d+):\s*(-?[\d\.]+)\s+([\d\.]+)\s+([\d\.]+)\s+", &im) {
                    mi := Integer(im[1]) + 1
                    if (mi <= g.freqs.Length)
                        g.irInt[mi] := im[4]+0
                }
                if InStr(lines[j], "NORMAL MODES") || InStr(lines[j], "RAMAN SPECTRUM")
                    break
            }
        }
        if (g.freqs.Length && InStr(L, "RAMAN SPECTRUM")) {
            j := i
            While (++j <= n) {
                if RegExMatch(lines[j], "^\s*(\d+):\s*(-?[\d\.]+)\s+([\d\.]+)\s+([\d\.]+)\s+", &rm) {
                    mi := Integer(rm[1]) + 1
                    if (mi <= g.freqs.Length)
                        g.raman[mi] := rm[4]+0
                }
                if InStr(lines[j], "NORMAL MODES") || InStr(lines[j], "THERMOCHEMISTRY")
                    break
            }
        }
        
        ; ---- 热化学 ----
        if InStr(L, "Eh") {
            if RegExMatch(L, "Electronic energy\s+\.*\s+(-?\d+\.\d+)\s+Eh", &ee)
                g.thermo.Eel := ee[1]+0
            if RegExMatch(L, "Zero point energy\s+\.*\s+(-?\d+\.\d+)\s+Eh", &zp)
                g.thermo.ZPE := zp[1]+0
            if RegExMatch(L, "Total thermal correction\s+(-?\d+\.\d+)\s+Eh", &tt)
                g.thermo.TE := tt[1]+0
            if RegExMatch(L, "Total thermal energy\s+\.*\s+(-?\d+\.\d+)\s+Eh", &tu)
                g.thermo.U := tu[1]+0
            if RegExMatch(L, "Total Enthalpy\s+\.*\s+(-?\d+\.\d+)\s+Eh", &th)
                g.thermo.H := th[1]+0
            if RegExMatch(L, "Final Gibbs free energy\s+\.*\s+(-?\d+\.\d+)\s+Eh", &fg)
                g.thermo.G := fg[1]+0
        }

        ; ---- 正常结束标记 ----
        if InStr(L, "ORCA TERMINATED NORMALLY")
            g.normalEnd := true

        ; ---- 原子电荷/自旋布居 (写入当前帧, 实现 pop=always 效果) ----
        if InStr(L, "MULLIKEN ATOMIC CHARGES") {
            g.mulC := [], g.mulS := [], g.hasMulSpin := InStr(L, "SPIN") ? true : false
            j := i
            While (++j <= n) {
                tl := Trim(lines[j])
                if (tl = "" || SubStr(tl, 1, 2) = "--")
                    continue
                if InStr(tl, "Sum of atomic") || InStr(tl, "REDUCED")
                    break
                if RegExMatch(tl, "^(\d+)\s+([A-Za-z]{1,3})\s*:\s*(-?\d+\.\d+)(?:\s+(-?\d+\.\d+))?", &mc) {
                    g.mulC.Push(mc[3]+0)
                    g.mulS.Push((mc[4] != "") ? mc[4]+0 : 0.0)
                }
            }
            ; 同步到当前帧, 供每帧写出电荷
            if (fr) {
                fr.mulC := g.mulC
                fr.mulS := g.mulS
                fr.hasMulSpin := g.hasMulSpin
            }
        }
        if InStr(L, "LOEWDIN ATOMIC CHARGES") {
            g.lowC := [], g.lowS := [], g.hasLowSpin := InStr(L, "SPIN") ? true : false
            j := i
            While (++j <= n) {
                tl := Trim(lines[j])
                if (tl = "" || SubStr(tl, 1, 2) = "--")
                    continue
                if InStr(tl, "Sum of atomic") || InStr(tl, "REDUCED")
                    break
                if RegExMatch(tl, "^(\d+)\s+([A-Za-z]{1,3})\s*:\s*(-?\d+\.\d+)(?:\s+(-?\d+\.\d+))?", &lc) {
                    g.lowC.Push(lc[3]+0)
                    g.lowS.Push((lc[4] != "") ? lc[4]+0 : 0.0)
                }
            }
        }
    }

    ; ---- 解析最后一个 NORMAL MODES 段的位移矩阵 ----
    ParseNormalModes(lines, n, g)
}

ParseNormalModes(lines, n, g) {
    ; 找到最后一个 "These modes are the Cartesian" 段的起点
    start := -1
    Loop n {
        if InStr(lines[A_Index], "These modes are the Cartesian")
            start := A_Index
    }
    if (start = -1) {
        g.modes := []
        return
    }
    ; 找到该段结束位置 (IR SPECTRUM / RAMAN SPECTRUM / THERMOCHEMISTRY / ORCA GEOMETRY RELAXATION / VIBRATIONAL FREQUENCIES)
    end := n + 1
    j := start
    While (++j <= n) {
        if InStr(lines[j], "IR SPECTRUM") || InStr(lines[j], "RAMAN SPECTRUM") || InStr(lines[j], "THERMOCHEMISTRY") || InStr(lines[j], "ORCA GEOMETRY RELAXATION") || InStr(lines[j], "VIBRATIONAL FREQUENCIES") {
            end := j
            break
        }
    }
    g.modes := []
    ; 页眉行 = 全部是整数(模式索引0-based)；数据行 = 前导整数 + 若干带小数的浮点数
    j := start
    While (++j < end) {
        L := lines[j]
        if !RegExMatch(L, "^\s+\d+(\s+\d+)*\s*$")
            continue
        colIdx := []
        For num in StrSplit(Trim(L), " ", "`t")
            if (num != "")
                colIdx.Push(Integer(num))
        ; 连续收集该页的数据行
        k := j
        While (k + 1 < end) {
            k++
            dl := lines[k]
            ; 非数据行(下一条页眉或段结束)则停止
            if !RegExMatch(dl, "^\s+(\d+)\s+-?[\d\.]+", &dm0)
                break
            comp := Integer(dm0[1])
            ; 去掉前导分量号后拆分余下的位移值
            rest := RegExReplace(dl, "^\s+\d+\s*", "")
            vals := []
            For tp in StrSplit(rest, " ", "`t")
                if (tp != "" && IsNum(tp))
                    vals.Push(tp + 0)
            if (vals.Length < colIdx.Length)
                break
            aIdx := comp // 3 + 1
            xyzI := Mod(comp, 3) + 1
            For ci, mv in vals {
                gi := colIdx[ci] + 1
                while (g.modes.Length < gi)
                    g.modes.Push([])
                arr := g.modes[gi]
                while (arr.Length < aIdx)
                    arr.Push([0, 0, 0])
                arr[aIdx][xyzI] := mv
            }
        }
        j := k - 1
    }
}

IsNum(s) {
    return RegExMatch(Trim(s), "^-?\d+\.?\d*$") ? true : false
}

AtomNum(sym) {
    static tbl := Map(
        "H",1,"HE",2,"LI",3,"BE",4,"B",5,"C",6,"N",7,"O",8,"F",9,"NE",10,
        "NA",11,"MG",12,"AL",13,"SI",14,"P",15,"S",16,"CL",17,"AR",18,
        "K",19,"CA",20,"SC",21,"TI",22,"V",23,"CR",24,"MN",25,"FE",26,
        "CO",27,"NI",28,"CU",29,"ZN",30,"GA",31,"GE",32,"AS",33,"SE",34,
        "BR",35,"KR",36,"RB",37,"SR",38,"Y",39,"ZR",40,"NB",41,"MO",42,
        "TC",43,"RU",44,"RH",45,"PD",46,"AG",47,"CD",48,"IN",49,"SN",50,
        "SB",51,"TE",52,"I",53,"XE",54,"CS",55,"BA",56,"LA",57,"CE",58,
        "PR",59,"ND",60,"PM",61,"SM",62,"EU",63,"GD",64,"TB",65,"DY",66,
        "HO",67,"ER",68,"TM",69,"YB",70,"LU",71,"HF",72,"TA",73,"W",74,
        "RE",75,"OS",76,"IR",77,"PT",78,"AU",79,"HG",80,"TL",81,"PB",82,
        "BI",83,"PO",84,"AT",85,"RN",86,"FR",87,"RA",88,"AC",89,"TH",90,
        "PA",91,"U",92,"NP",93,"PU",94)
    return tbl.Has(StrUpper(Trim(sym))) ? tbl[StrUpper(Trim(sym))] : 0
}

; ---------------------------------------------------------------
; 写出调度
; ---------------------------------------------------------------
AskAndWrite(outFile, srcFile, cliMode := false) {
    global g
    out := []

    ; 扫描模式询问
    scanKeepAll := false
    if (g.calc = "SCAN") {
        explicit := (g.HasOwnProp("scanMode") && (g.scanMode = "all" || g.scanMode = "last"))
        if (!explicit) {
            lang := GetUiLang()
            if (lang = "en") {
                r := MsgBox("Relaxed surface scan detected.`n`nChoose output mode:`n`nYes (Y) — keep only the last frame of each scan step (scanlast)`nNo (N)  — keep all frames (scanall)`nCancel (C) — abort conversion", "Scan conversion mode (scanlast / scanall)", "YesNoCancel Icon?")
            } else {
                r := MsgBox("检测到 RELAXED SURFACE SCAN。`n`n请选择输出方式:`n`n是(Y) —— 仅保留每个扫描步的最后一帧 (scanlast)`n否(N) —— 保留所有帧 (scanall)`n取消(C) —— 中止转换", "扫描转换模式 (scanlast / scanall)", "YesNoCancel Icon?")
            }
            if (r = "Cancel")
                ExitApp
            scanKeepAll := (r = "No")
            g.scanMode := (r = "No") ? "all" : "last"
        } else {
            scanKeepAll := (g.scanMode = "all")
        }
    }

    ; 选择要输出的帧序列
    frames := g.frames
    if (g.calc = "SCAN") {
        ; 丢弃扫描开始前的回显几何帧
        clean := []
        For f in g.frames
            if (f.scanStep > 0)
                clean.Push(f)
        if (!clean.Length)
            clean := g.frames
        if scanKeepAll {
            frames := clean               ; 所有帧(opt 轨迹)
        } else {
            ; 仅每步最后一帧: 优先取该步已收敛且有能量的帧
            frames := []
            lastOf := Map(), lastConv := Map(), order := []
            For f in clean {
                st := f.scanStep
                hasEn := (f.energy != "" || f.etot != "")
                if !lastOf.Has(st) {
                    order.Push(st)
                    lastOf[st] := ""
                    lastConv[st] := ""
                }
                if hasEn
                    lastOf[st] := f
                if (f.converged && hasEn)
                    lastConv[st] := f
            }
            For st in order
                frames.Push(lastConv[st] != "" ? lastConv[st] : lastOf[st])
        }
    }

    WriteLog(outFile, srcFile, frames)
}

Dbg(s) {
    FileAppend(s . "`n", A_Temp . "\orca2g_dbg.txt")
}

Fmt(v, w, d) {
    if (v = "" || !IsNum(v))
        v := 0
    v := v + 0
    s := Format("{:." . d . "f}", v)
    while (StrLen(s) < w)
        s := " " . s
    return s
}

WriteLog(outFile, srcFile, frames) {
    global g
    o := []

    nelec := 0
    For z in g.nums
        nelec += z
    nelec -= g.charge
    unpaired := g.mult - 1
    na := (nelec + unpaired) // 2
    nb := (nelec - unpaired) // 2

    SplitPath(srcFile, &srcName)

    ; ---------- 头部 (GaussView 识别并行/内存设置所需) ----------
    pid := DllCall("GetCurrentProcessId")
    o.Push(" Entering Gaussian System, Link 0=g16")
    o.Push(" Initial command:")
    o.Push(" /opt/app/gaussian/16-C01-avx2/g16/l1.exe `"/tmp/Gau-" . Mod(pid, 999999) . ".inp`" -scrdir=`"/tmp/`"")
    o.Push(" Entering Link 1 = /opt/app/gaussian/16-C01-avx2/g16/l1.exe PID=     " . pid . ".")
    o.Push("")

    o.Push(" This file was generated by orca2gaussian.ahk (AutoHotkey v2)")
    o.Push(" Source ORCA output: " . srcName)

    ; 构造 G16 风格 route (逐字保留 ORCA 关键字, 不做转化)
    route := "# "
    route .= (g.functional != "" ? g.functional : "HF") . "/" . (g.basis != "" ? g.basis : "Gen")
    For wv in g.kwVerbatim
        route .= " " . wv

    o.Push(" ******************************************")
    o.Push(" Gaussian 16:  ES64L-G16RevC.01 24-Jan-2017")
    o.Push("                 1-Jan-2026 ")
    o.Push(" ******************************************")
    if (g.nprocs) {
        o.Push("%nprocshared=" . g.nprocs)
        o.Push(" Will use up to   " . g.nprocs . " processors via shared memory.")
    }
    if (g.maxcore) {
        ; ORCA %maxcore 是每核内存；Gaussian %mem 是总内存 => 相乘再转 GB
        np := (g.nprocs > 0) ? g.nprocs : 1
        o.Push("%mem=" . MaxcoreGB(g.maxcore * np) . "GB")
    }
    o.Push(" ----------------------------------------------------------------------")
    o.Push(" " . route)
    o.Push(" ----------------------------------------------------------------------")
    o.Push(" 1/38=1,158=3,167=1,172=1/1;")
    o.Push(" 2/12=2,17=6,18=5,40=1/2;")
    o.Push(" 3/5=7,11=2,16=1,17=8,25=1,27=10,30=1,70=32201,72=2,74=-58,75=-4,116=1,158=3/1,2,3;")
    o.Push(" 4/60=-1/1;")
    o.Push(" 5/5=2,7=64,8=3,13=1,38=5,53=2,87=10/2,8;")
    o.Push(" 6/7=2,8=2,9=2,10=2,28=1,87=10/1;")
    o.Push(" 99/5=1,9=1/99;")
    o.Push("  using 2006 physical constants.")
    o.Push(" ----------")
    o.Push(" Title Card")
    o.Push(" ----------")
    o.Push(" Symbolic Z-matrix:")
    o.Push(" Charge = " . PadL(g.charge, 2) . " Multiplicity = " . PadL(g.mult, 2))
    ; Z-matrix 坐标回显 (符号 + xyz), 取自第一个含几何的帧
    zmGeo := []
    For frz in frames
        if (frz.HasOwnProp("geom") && frz.geom.Length) {
            zmGeo := frz.geom
            break
        }
    if (zmGeo.Length) {
        For a in zmGeo {
            o.Push(PadL(a[1], 2) . " " . PadL(Format("{:.5f}", a[2]), 12)
                 . " " . PadL(Format("{:.5f}", a[3]), 12) . "  " . PadL(Format("{:.5f}", a[4]), 12))
        }
    }
    o.Push("")

    ; 约束/扫描回显 (ModRedundant 风格, 原子号转 1-based)
    consTxt := []
    For c in g.constraints {
        if c.wild {
            consTxt.Push("! " . c.typ . " (含通配符的约束, 见 ORCA 输入)")
            continue
        }
        ids := ""
        For ix in c.idx
            ids .= " " . (ix + 1)
        valPart := (c.val != "") ? " " . c.val : ""
        consTxt.Push(c.typ . ids . valPart . " F")
    }
    if (g.scanDef != "") {
        sd := g.scanDef
        inc := 0
        try inc := (sd.stop - sd.start) / Max(sd.steps - 1, 1)
        consTxt.Push("! Relaxed scan: " . sd.typ . " " . sd.atoms . " = " . Format("{:.6f}", sd.start)
                   . " -> " . Format("{:.6f}", sd.stop) . " (" . sd.steps . " points, step " . Format("{:.6f}", inc) . ")")
        consTxt.Push(sd.typ . " " . sd.atoms . " S " . sd.steps . " " . Format("{:.6f}", sd.start) . " " . Format("{:.6f}", inc))
    }
    if (consTxt.Length) {
        o.Push("")
        For ct in consTxt
            o.Push(" ! ModRedundant/scan (from ORCA, atoms 1-based): " . ct)
    }

    ; ---------- 帧 ----------
    ; 过滤无能量帧(GaussView 需要每帧都有能量锚点)
    good := []
    For f in frames
        if (f.energy != "" || f.etot != "")
            good.Push(f)
    if (good.Length)
        frames := good

    nf := frames.Length
    if (nf = 0) {
        Dbg("ERROR: no usable geometry frames")
        ExitApp
    }

    anyConv := false
    For f in frames
        if f.converged
            anyConv := true

    ; OfakeG 顺序: [orient1][scf1] [Step1块] [orient2][scf2] [Step2块] ...
    fi := 0
    For f in frames {
        fi++
        WriteOrientation(o, f.geom, true)
        WriteOrientation(o, f.geom)
        ; 能量: TD 优化用激发态总能量 —— 必须用 G16 规范格式供 GaussView 解析
        en := (g.isTD && f.etot != "") ? f.etot : f.energy
        meth := (g.mult > 1) ? "UKS" : "RKS"
        o.Push(" SCF Done:  E(" . meth . ") =  " . Format("{:.10f}", en)
             . "     A.U. after    1 cycles")

        ; 每帧写出电荷/自旋(G16 pop=always 效果)
        if (f.HasOwnProp("mulC") && f.mulC.Length)
            WriteCharges(o, f.mulC, f.mulS, f.hasMulSpin)

        ; TD 光谱
        if (g.isTD && f.td.Length)
            WriteExcited(o, f)
        o.Push("")

        ; Step 块(优化/扫描), conv 数据属于当前帧
        if (g.calc = "OPT" || g.calc = "SCAN") {
            o.Push(GRADBAR())
            o.Push(" Berny optimization.")
            o.Push(" Internal  Forces:  Max " . Format("{:.6f}", IsNum(f.conv.maxG)?f.conv.maxG+0:0)
                 . " RMS " . Format("{:.6f}", IsNum(f.conv.rmsG)?f.conv.rmsG+0:0))
            o.Push(" Search for a local " . (g.isTS ? "maximum." : "minimum."))
            o.Push(" Step number   " . fi . " out of a maximum of  " . (nf + 10))
            o.Push(" All quantities printed in internal units (Hartrees-Bohrs-Radians)")
            WriteConvItems(o, f)
            o.Push(GRADBAR())
            ; 扫描注释放在下一几何之前
            if (g.calc = "SCAN" && f.scanVal != "" && fi < nf)
                o.Push(" ! Next scan point: " . frames[fi + 1].scanVal)
            o.Push("")
        }
    }

    ; 收敛声明
    if (anyConv && (g.calc = "OPT" || g.calc = "SCAN")) {
        o.Push(" Optimization completed.")
        o.Push(" -- Stationary point found.")
        o.Push("")
    }

    ; ---------- 频率 ----------
    if (g.hasFreq && g.freqs.Length)
        WriteFreqs(o)

    ; ---------- 热化学 ----------
    WriteThermo(o)

    ; ---------- 结束标记: ORCA 非正常结束则 log 也非正常结束 ----------
    if (g.normalEnd) {
        o.Push(" Normal termination of Gaussian 16 (fake)")
    } else {
        o.Push(" Error termination request processed by link 9999.")
        o.Push(" Error termination of Gaussian 16 (fake).")
        o.Push(" (Source ORCA output did NOT terminate normally)")
    }

    txt := ""
    For ln in o
        txt .= ln . "`n"
    if FileExist(outFile)
        FileDelete(outFile)
    FileAppend(txt, outFile)
    if (A_Args.Length >= 1) {
        Dbg("OK frames=" . frames.Length . " freqs=" . g.freqs.Length . " type=" . g.calc)
        return
    }
}

GRADBAR() {
    return "GradGradGradGradGradGradGradGradGradGradGradGradGradGradGradGradGradGrad"
}

PadL(s, w) {
    while (StrLen(s) < w)
        s := " " . s
    return s
}

WriteConvItems(o, f) {
    o.Push("         Item               Value     Threshold  Converged?")
    lab1 := " Maximum Force            "
    lab2 := " RMS     Force            "
    lab3 := " Maximum Displacement     "
    lab4 := " RMS     Displacement     "
    lab0 := " Energy Change            "
    thrF := "     0.000300     "
    thrR := "     0.000100     "
    thrD := "     0.004000     "
    thrS := "     0.002000     "
    thrE := "     0.000005     "
    ynE := f.yesE ? "YES" : "NO "
    ynM := f.yesMG ? "YES" : "NO "
    ynR := f.yesRG ? "YES" : "NO "
    ynDS := f.yesMS ? "YES" : "NO "
    ynRS := f.yesRS ? "YES" : "NO "
    vE := IsNum(f.conv.eChg) ? f.conv.eChg + 0 : 0
    vM := IsNum(f.conv.maxG) ? f.conv.maxG + 0 : 0
    vR := IsNum(f.conv.rmsG) ? f.conv.rmsG + 0 : 0
    vD := IsNum(f.conv.maxS) ? f.conv.maxS + 0 : 0
    vS := IsNum(f.conv.rmsS) ? f.conv.rmsS + 0 : 0
    hasAny := (f.conv.eChg != "" || f.conv.maxG != "")
    if (!hasAny)
        return
    if (f.conv.eChg != "")
        o.Push(lab0 . PadL(Format("{:.6f}", vE), 9) . thrE . ynE)
    o.Push(lab1 . PadL(Format("{:.6f}", vM), 9) . thrF . ynM)
    o.Push(lab2 . PadL(Format("{:.6f}", vR), 9) . thrR . ynR)
    o.Push(lab3 . PadL(Format("{:.6f}", vD), 9) . thrD . ynDS)
    o.Push(lab4 . PadL(Format("{:.6f}", vS), 9) . thrS . ynRS)
}

WriteOrientation(o, geom, input := false) {
    global g
    if (input)
        o.Push("                          Input orientation:                          ")
    else
        o.Push("                         Standard orientation:                         ")
    o.Push(" ---------------------------------------------------------------------")
    o.Push(" Center     Atomic      Atomic             Coordinates (Angstroms)")
    o.Push(" Number     Number       Type             X           Y           Z")
    o.Push(" ---------------------------------------------------------------------")
    ai := 0
    For a in geom {
        ai++
        z := (ai <= g.nums.Length) ? g.nums[ai] : AtomNum(a[1])
        o.Push(" " . PadL(ai, 6) . PadL(z, 11) . PadL(0, 12)
             . PadL(Format("{:.6f}", a[2]), 16) . PadL(Format("{:.6f}", a[3]), 12)
             . PadL(Format("{:.6f}", a[4]), 12))
    }
    o.Push(" ---------------------------------------------------------------------")
}

WriteExcited(o, f) {
    static words := Map(1, "Singlet", 2, "Doublet", 3, "Triplet", 4, "Quartet", 5, "Quintet", 6, "Sextet", 7, "Septet", 8, "Octet")
    o.Push(" Excitation energies and oscillator strengths:")
    o.Push("")
    For st in f.td {
        mw := (st.multW != "" && words.Has(st.multW)) ? words[st.multW] : "Singlet"
        sym := (st.s2 != "" && st.s2 + 0 > 0.3) ? "'" : ""
        o.Push(" Excited State " . Fmt(st.idx, 4, 0) . ":  " . mw . "-A" . sym
             . Fmt(st.ev, 11, 4) . " eV" . Fmt(1239.841984 / Abs(st.ev > 0 ? st.ev : 1), 9, 2) . " nm  f="
             . Format("{:.4f}", st.fosc)
             . ((st.s2 != "") ? "      <S**2>=" . Format("{:.3f}", st.s2 + 0) : ""))
        ti := 0
        For tr in st.trans {
            ti++
            if (ti > 24) {
                o.Push("        ... (more contributions omitted)")
                break
            }
            o.Push("      " . StrUpper(tr[1]) . " -> " . StrUpper(tr[2]) . "        " . Format("{:.5f}", Abs(tr[4])))
        }
    }
    o.Push("")
    global g
    if (g.calc = "OPT" || g.calc = "SCAN") {
        o.Push(" This state for optimization and/or subsequent second-order properties.")
        o.Push(" The state that is being optimized has been set to state number   " . g.tdRoot)
    }
    o.Push("")
}

WriteFreqs(o) {
    global g
    local freqStartIdx
    ; ---- G16/OfakeG 频率块: 无 Low frequencies 前缀, 无 Red. masses/Frc consts 行 ----
    o.Push(" Harmonic frequencies (cm**-1), IR intensities (KM/Mole), Raman scattering")
    o.Push(" activities (A**4/AMU), depolarization ratios for plane and unpolarized")
    o.Push(" incident light, reduced masses (AMU), force constants (mDyne/A),")
    o.Push(" and normal coordinates:")
    
n := g.freqs.Length
    ; 跳过前 6 个平动/转动模式 (0.0 cm-1), 显示编号从 1 开始
    freqStartIdx := 7
    if (freqStartIdx > n) {
        freqStartIdx := 1
    }
    mi := freqStartIdx
    While (mi <= n) {
        hi := Min(mi + 2, n)
        ; 模式号行与对称行: PadL 23
        l1 := ""
        Loop (hi - mi + 1) {
            k := A_Index + mi - 1
            l1 .= PadL(k - 6, 23)
        }
        l2 := ""
        Loop (hi - mi + 1) {
            l2 .= PadL("A", 23)
        }
        lf := " Frequencies --", li := " IR Inten    --"
        Loop (hi - mi + 1) {
            k := A_Index + mi - 1
            w1 := (A_Index = 1 ? 12 : 23)
            lf .= PadL(Format("{:.4f}", g.freqs[k]), w1)
            li .= PadL(Format("{:.4f}", g.irInt[k]), w1)
        }
        o.Push("")
        o.Push(l1)
        o.Push(l2)
        o.Push(lf)
        o.Push(li)
        o.Push("  Atom  AN      X      Y      Z        X      Y      Z        X      Y      Z")
        Loop g.natoms {
            at := A_Index
            row := PadL(at, 6) . PadL(g.nums[at], 4)
            Loop (hi - mi + 1) {
                k := A_Index + mi - 1
                disp := [0,0,0]
                if (k <= g.modes.Length && at <= g.modes[k].Length)
                    disp := g.modes[k][at]
                row .= PadL(Format("{:.2f}", disp[1]), 9)
                     . PadL(Format("{:.2f}", disp[2]), 7)
                     . PadL(Format("{:.2f}", disp[3]), 7)
            }
            o.Push(row)
        }
        mi := hi + 1
    }
    o.Push("")
}

MassOf(sym) {
    static tbl := Map(
        "H",1.008,"HE",4.003,"LI",6.94,"BE",9.012,"B",10.81,"C",12.011,"N",14.007,
        "O",15.999,"F",18.998,"NE",20.180,"NA",22.990,"MG",24.305,"AL",26.982,
        "SI",28.086,"P",30.974,"S",32.06,"CL",35.45,"AR",39.948,"K",39.098,
        "CA",40.078,"TI",47.867,"V",50.942,"CR",51.996,"MN",54.938,"FE",55.845,
        "CO",58.933,"NI",58.693,"CU",63.546,"ZN",65.38,"BR",79.904,"I",126.904,
        "AG",107.868,"AU",196.967,"PT",195.084,"PD",106.42,"SN",118.710)
    s := StrUpper(Trim(sym))
    return tbl.Has(s) ? tbl[s] : 12.011
}

MaxcoreGB(mb) {
    ; MB -> GB (向上取整, 最小 1)
    if !IsNum(mb)
        return 1
    mbi := Integer(mb + 0)
    gb := Ceil(mbi / 1024)
    return (gb >= 1) ? gb : 1
}

WriteCharges(o, mulC := 0, mulS := 0, hasMulSpin := 0) {
    global g
    if (mulC = 0) {
        mulC := g.mulC
        mulS := g.mulS
        hasMulSpin := g.hasMulSpin
    }
    if (mulC.Length = 0)
        return
    na_ := mulC.Length
    if (hasMulSpin) {
        ; G16 开壳层格式: 表头只有列号 1(charge) 2(spin)
        o.Push(" Mulliken charges and spin densities:")
        o.Push(PadL(1, 16) . PadL(2, 10))
        Loop na_ {
            sym := (A_Index <= g.syms.Length) ? g.syms[A_Index] : "X"
            o.Push(PadL(A_Index, 6) . "  " . sym . "   " . Format("{:.6f}", mulC[A_Index]) . "   " . Format("{:.6f}", mulS[A_Index]))
        }
    } else {
        ; G16 闭壳层格式: " Mulliken charges:" 表头只有列号 1
        o.Push(" Mulliken charges:")
        o.Push(PadL(1, 16))
        Loop na_ {
            sym := (A_Index <= g.syms.Length) ? g.syms[A_Index] : "X"
            o.Push(PadL(A_Index, 6) . "  " . sym . "   " . Format("{:.6f}", mulC[A_Index]))
        }
    }
    o.Push(" Sum of Mulliken charges =  " . Format("{:.5f}", SumArr(mulC)))
    o.Push("")
}

SumArr(arr) {
    s := 0.0
    For v in arr
        s += v
    return s
}

WriteThermo(o) {
    global g
    t := g.thermo
    if !t.HasOwnProp("ZPE")
        return
    eel := t.HasOwnProp("Eel") ? t.Eel : 0
    te := t.HasOwnProp("TE") ? t.TE : 0
    tcE := t.HasOwnProp("U") ? (t.U - eel) : (te + t.ZPE)
    tcH := t.HasOwnProp("H") ? (t.H - eel) : (tcE + 0.000944)
    tcG := t.HasOwnProp("G") ? (t.G - eel) : (tcH + 0)
    o.Push(" Temperature   298.150 Kelvin.  Pressure   1.00000 Atm.")
    o.Push(" Zero-point correction=                           " . Fmt(t.ZPE, 12, 6) . " Hartree")
    o.Push(" Thermal correction to Energy=                    " . Fmt(tcE, 12, 6))
    o.Push(" Thermal correction to Enthalpy=                  " . Fmt(tcH, 12, 6))
    o.Push(" Thermal correction to Gibbs Free Energy=         " . Fmt(tcG, 12, 6))
    o.Push(" Sum of electronic and zero-point Energies=       " . Fmt(eel + t.ZPE, 16, 6))
    o.Push(" Sum of electronic and thermal Energies=          " . Fmt(eel + tcE, 16, 6))
    o.Push(" Sum of electronic and thermal Enthalpies=        " . Fmt(eel + tcH, 16, 6))
    o.Push(" Sum of electronic and thermal Free Energies=     " . Fmt(eel + tcG, 16, 6))
    o.Push("")
}
