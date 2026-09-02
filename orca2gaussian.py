#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
orca2gaussian.py  --  ORCA output (.out) -> Gaussian16-fake log readable by GaussView

Faithful port of orca2gaussian.ahk (AutoHotkey v2). Produces a "_fake.log" that
mimics a real Gaussian 16 log so that GaussView displays real data in the
Results -> Vibrations / UV-Vis / Atomic Charges menus.

CLI:  python orca2gaussian.py <file.out> [all|last]
        - scan tasks: optional [all]=keep every scan frame,
          [last]=keep only last frame of each scan step.
          If omitted, an interactive prompt is shown (unless --yes is given).
"""

import argparse
import os
import re
import sys
import tempfile
from pathlib import Path

# ---------------------------------------------------------------------------
# 名称 / 调试
# ---------------------------------------------------------------------------
METHOD_NAME = "orca2gaussian.py"


def dbg(msg: str) -> None:
    """Mirror of AHK Dbg(): append a line to a temp debug file."""
    tmp = tempfile.gettempdir()
    dbf = os.path.join(tmp, "orca2g_dbg.txt")
    try:
        with open(dbf, "a", encoding="utf-8") as fh:
            fh.write(str(msg) + "\n")
    except OSError:
        pass


# ---------------------------------------------------------------------------
# UI 语言 (读 ORmagiCA_settings.ini 的 [General] Language: zh/en)
# ---------------------------------------------------------------------------
def settings_path() -> str:
    return os.path.join(os.path.dirname(os.path.abspath(__file__)),
                        "ORmagiCA_settings.ini")


def get_ui_lang() -> str:
    try:
        import configparser
        cp = configparser.ConfigParser()
        cp.read(settings_path(), encoding="utf-8")
        v = cp.get("General", "Language", fallback="zh")
        return "en" if v.strip().lower() == "en" else "zh"
    except Exception:
        return "zh"


def L(zh: str, en: str) -> str:
    return en if get_ui_lang() == "en" else zh


# ---------------------------------------------------------------------------
# 元素表
# ---------------------------------------------------------------------------
ATOM_NUM = {
    "H": 1, "HE": 2, "LI": 3, "BE": 4, "B": 5, "C": 6, "N": 7, "O": 8,
    "F": 9, "NE": 10, "NA": 11, "MG": 12, "AL": 13, "SI": 14, "P": 15,
    "S": 16, "CL": 17, "AR": 18, "K": 19, "CA": 20, "SC": 21, "TI": 22,
    "V": 23, "CR": 24, "MN": 25, "FE": 26, "CO": 27, "NI": 28, "CU": 29,
    "ZN": 30, "GA": 31, "GE": 32, "AS": 33, "SE": 34, "BR": 35, "KR": 36,
    "RB": 37, "SR": 38, "Y": 39, "ZR": 40, "NB": 41, "MO": 42, "TC": 43,
    "RU": 44, "RH": 45, "PD": 46, "AG": 47, "CD": 48, "IN": 49, "SN": 50,
    "SB": 51, "TE": 52, "I": 53, "XE": 54, "CS": 55, "BA": 56, "LA": 57,
    "CE": 58, "PR": 59, "ND": 60, "PM": 61, "SM": 62, "EU": 63, "GD": 64,
    "TB": 65, "DY": 66, "HO": 67, "ER": 68, "TM": 69, "YB": 70, "LU": 71,
    "HF": 72, "TA": 73, "W": 74, "RE": 75, "OS": 76, "IR": 77, "PT": 78,
    "AU": 79, "HG": 80, "TL": 81, "PB": 82, "BI": 83, "PO": 84, "AT": 85,
    "RN": 86, "FR": 87, "RA": 88, "AC": 89, "TH": 90, "PA": 91, "U": 92,
    "NP": 93, "PU": 94,
}

ATOM_MASS = {
    "H": 1.008, "HE": 4.003, "LI": 6.94, "BE": 9.012, "B": 10.81,
    "C": 12.011, "N": 14.007, "O": 15.999, "F": 18.998, "NE": 20.180,
    "NA": 22.990, "MG": 24.305, "AL": 26.982, "SI": 28.086, "P": 30.974,
    "S": 32.06, "CL": 35.45, "AR": 39.948, "K": 39.098, "CA": 40.078,
    "TI": 47.867, "V": 50.942, "CR": 51.996, "MN": 54.938, "FE": 55.845,
    "CO": 58.933, "NI": 58.693, "CU": 63.546, "ZN": 65.38, "BR": 79.904,
    "I": 126.904, "AG": 107.868, "AU": 196.967, "PT": 195.084, "PD": 106.42,
    "SN": 118.710,
}


def atom_num(sym: str) -> int:
    s = sym.strip().upper()
    return ATOM_NUM.get(s, 0)


def mass_of(sym: str) -> float:
    s = sym.strip().upper()
    return ATOM_MASS.get(s, 12.011)


def isnum(s) -> bool:
    if s is None:
        return False
    return re.match(r"^-?\d+\.?\d*$", str(s).strip()) is not None


# ---------------------------------------------------------------------------
# 通用格式化辅助 (镜像 AHK PadL / Fmt)
# ---------------------------------------------------------------------------
def padl(s, w: int) -> str:
    """Left pad string with spaces to width w (AHK PadL)."""
    s = str(s)
    if len(s) >= w:
        return s[:w] if False else s
    return " " * (w - len(s)) + s


def fmt(v, w: int, d: int) -> str:
    """Right-justify numeric v to width w with d decimals (AHK Fmt)."""
    try:
        v = float(v)
    except (ValueError, TypeError):
        v = 0.0
    s = f"{v:.{d}f}"
    return s.rjust(w)


def gradbar() -> str:
    return "Grad" * 18


def maxcore_gb(mb) -> int:
    """Convert MB to GB (round up, minimum 1). 2000 MB -> 2 GB."""
    try:
        mb = int(float(mb))
    except (ValueError, TypeError):
        return 1
    gb = -(-mb // 1024)  # ceiling division
    return gb if gb >= 1 else 1


# ---------------------------------------------------------------------------
# 数据结构
# ---------------------------------------------------------------------------
class Conv:
    def __init__(self):
        self.eChg = ""
        self.rmsG = ""
        self.maxG = ""
        self.rmsS = ""
        self.maxS = ""


class TDState:
    def __init__(self, idx, ev, au, fosc=0.0, s2="", multW="", trans=None):
        self.idx = idx
        self.ev = ev
        self.au = au
        self.nm = 0.0
        self.fosc = fosc
        self.s2 = s2
        self.multW = multW
        self.trans = trans if trans is not None else []


class Frame:
    def __init__(self, scan_step=0, scan_val=""):
        self.geom = []          # list of [sym, x, y, z]
        self.energy = ""        # FINAL SINGLE POINT ENERGY
        self.etot = ""          # E(tot)
        self.conv = Conv()
        self.yesE = 0
        self.yesMG = 0
        self.yesRG = 0
        self.yesMS = 0
        self.yesRS = 0
        self.converged = False
        self.td = []            # list of TDState
        self.scan_step = scan_step
        self.scan_val = scan_val
        # per-frame charges (pop=always), only set when ORCA provided them
        self.mulC = None
        self.mulS = None
        self.hasMulSpin = False


class Frames(list):
    """A 1-based-ish list wrapper matching AHK frame indexing via .frames[idx]."""


class Run:
    def __init__(self):
        self.natoms = 0
        self.syms = []
        self.nums = []
        self.charge = 0
        self.mult = 1
        self.calc = "SP"
        self.isTS = False
        self.isTD = False
        self.hasFreq = False
        self.nbasis = 0
        self.frames = []
        self.cur_frame = 0
        self.constraints = []
        self.scan_def = ""
        self.scan_val = ""
        self.freqs = []
        self.ir_int = []
        self.red_mass = []
        self.raman = []
        self.modes = []
        self.thermo = {}
        self.td_root = 1
        self.kw_line = ""
        self.functional = ""
        self.basis = ""
        self.solvent = ""
        self.nroots = 0
        self.normal_end = False
        self.mulC = []
        self.mulS = []
        self.hasMulSpin = False
        self.lowC = []
        self.lowS = []
        self.hasLowSpin = False
        self.cur_scan_step = 0
        self.scan_mode = ""
        self.nprocs = 0      # %pal nprocs N end
        self.maxcore = 0     # %maxcore N (MB)
        self.kw_verbatim = []  # ORCA 关键词逐字 (不含 functional/basis), 按原样保留


# ---------------------------------------------------------------------------
# 解析 (ParseORCA)
# ---------------------------------------------------------------------------
FUNC_RE = re.compile(
    r"^(b3lyp|pbe0|pbe|tpss|m06|wb97|revpbe|otpss|bp86|blyp|olyp|pw91|b97|hf|mp2|"
    r"ccsd|dlpno|sos-|soss-|r2scan)", re.I)
BASIS_RE = re.compile(
    r"^(def2|cc-p|cc-pv|aug-|sto-|6-3|lanl|sarc|dz|tz|qz|svp|tzv|qzv|pc-|ma)", re.I)


def parse_orca(path: str) -> Run:
    g = Run()
    with open(path, "r", encoding="utf-8", errors="replace") as fh:
        raw = fh.read()
    lines = raw.split("\n")
    n = len(lines)

    in_constraints = False
    in_absorb = False
    absorb_idx = 0

    i = -1
    while i + 1 < n:
        i += 1
        L = lines[i]

        # ---- 输入回显区 (仅当行以 | 开头才匹配) ----
        m = re.match(r"^\|\s*\d+>\s*(.*)$", L) if L.startswith("|") else None
        if m:
            t = m.group(1)
            if re.search(r"(%tddft|\btddft\b)", t, re.I):
                g.isTD = True
            nr = re.search(r"nroots\s+(\d+)", t, re.I)
            if nr:
                g.nroots = int(nr.group(1))
            kwraw = re.match(r"^\s*!\s*(.+)$", t)
            if kwraw:
                g.kw_line = kwraw.group(1).strip()
                for tk in g.kw_line.split(" "):
                    tlk = tk.lower()
                    if tlk == "" or tlk == "!":
                        continue
                    # 含 / 或 ( 的 token (辅基/溶剂/带括号) 一律不当作 functional/basis
                    if "/" in tk or "(" in tk:
                        if tlk == "scan":
                            g.functional = tk
                        elif tlk not in (g.functional.lower() if g.functional else "",
                                         g.basis.lower() if g.basis else ""):
                            g.kw_verbatim.append(tk)
                        continue
                    if (g.functional == "" and (FUNC_RE.match(tlk)
                                                or tlk == "scan")):
                        g.functional = tk
                    elif g.basis == "" and BASIS_RE.match(tlk) and "/" not in tlk:
                        g.basis = tk
                    else:
                        sv = re.match(r"^(smd|cpcm)\((.+)\)", tlk)
                        if sv:
                            g.solvent = sv.group(2)
                        # 逐字保留非 functional/basis 的关键词
                        if tlk not in (g.functional.lower() if g.functional else "",
                                       g.basis.lower() if g.basis else ""):
                            g.kw_verbatim.append(tk)
            if re.match(r"^!", t):
                kw = t.lower()
                if re.search(r"\boptts\b", kw):
                    g.isTS = True
                    g.calc = "OPT"
                elif re.search(r"(^|\s)(scants|scan|scan_ts)(\s|$)", kw):
                    g.calc = "SCAN"
                elif re.search(r"(^|\s)opt(\s|$)|\bcopt\b|\bopt\b", kw):
                    g.calc = "OPT"
                if re.search(r"(^|\s)(freq|numfreq)(\s|$)", kw):
                    g.hasFreq = True
                if re.search(r"(^|\s)(td|cis|steom|rocis)(\s|$)", kw):
                    g.isTD = True
            # %geom 块内的约束
            if re.match(r"^\s*Constraints\s*$", t, re.I):
                in_constraints = True
            elif in_constraints and re.match(r"^\s*end\s*$", t, re.I):
                in_constraints = False
            elif in_constraints and re.match(r"^\s*\{\s*([BADCB])\s+(.*?)\s*\}\s*$", t, re.I):
                c = re.match(r"^\s*\{\s*([BADCB])\s+(.*?)\s*\}\s*$", t, re.I)
                typ = c.group(1).upper()
                rest = c.group(2).strip()
                idxs = []
                val = ""
                wildcard = False
                for tok in rest.split(" "):
                    tk = tok.strip()
                    if tk == "":
                        continue
                    if re.match(r"^[\d\.]+$", tk):
                        if (val == "" and re.match(r"^\d+$", tk)
                                and len(idxs) < 4 and "." not in tk):
                            idxs.append(int(tk))
                        else:
                            val = tk
                    elif re.match(r"^[A-Z]$", tk):
                        pass  # 结尾的 C 标志忽略（约束即冻结）
                    else:
                        wildcard = True
                g.constraints.append({
                    "typ": typ, "idx": idxs, "val": val, "wild": wildcard})
            # 扫描定义: Scan B 6 19 = 2.5, 1.8, 8
            s = re.match(
                r"^\s*Scan\s+([BAD])\s+([\d\s]+)=\s*(-?[\d\.]+)\s*,\s*(-?[\d\.]+)"
                r"\s*,\s*(\d+)", t, re.I)
            if s:
                g.scan_def = {
                    "typ": s.group(1).upper(),
                    "atoms": s.group(2).strip(),
                    "start": s.group(3),
                    "stop": s.group(4),
                    "steps": s.group(5),
                }

        # ---- 电荷/多重度 ----
        if "xyz" in L and re.search(r"\*\s*xyz\s+(-?\d+)\s+(\d+)", L, re.I):
            q = re.search(r"\*\s*xyz\s+(-?\d+)\s+(\d+)", L, re.I)
            g.charge = int(q.group(1))
            g.mult = int(q.group(2))
        if g.nbasis == 0:
            b = re.search(r"Basis Dimension\s+Dim\s+\.*\s+(\d+)", L)
            if b:
                g.nbasis = int(b.group(1))

        # ---- nprocs / memory (%pal nprocs, %maxcore) ----
        if "%pal" in L:
            np_ = re.search(r"%pal\s+nprocs\s+(\d+)", L, re.I)
            if np_:
                g.nprocs = int(np_.group(1))
        if "%maxcore" in L:
            mc_ = re.search(r"%maxcore\s+(\d+)", L, re.I)
            if mc_:
                g.maxcore = int(mc_.group(1))

        # ---- 原子列表(第一次笛卡尔坐标) ----
        if g.natoms == 0 and "CARTESIAN COORDINATES (ANGSTROEM)" in L:
            j = i
            while j + 1 < n:
                j += 1
                tl = lines[j].strip()
                if tl != "" and not tl.startswith("--"):
                    break
            while j < n:
                tl = lines[j].strip()
                if tl == "" or tl.startswith("--"):
                    break
                mm = re.match(
                    r"^([A-Za-z]{1,3})\s+(-?[\d\.]+)\s+(-?[\d\.]+)\s+(-?[\d\.]+)$", tl)
                if mm:
                    g.natoms += 1
                    g.syms.append(mm.group(1))
                    g.nums.append(atom_num(mm.group(1)))
                j += 1

        # ---- 新优化帧 ----
        is_new_frame = False
        if "GEOMETRY OPTIMIZATION CYCLE" in L:
            cy = re.search(r"GEOMETRY OPTIMIZATION CYCLE\s+(\d+)", L)
            if cy:
                is_new_frame = True
        elif "RELAXED SURFACE SCAN STEP" in L:
            sc = re.search(r"RELAXED SURFACE SCAN STEP\s+(\d+)", L)
            if sc:
                is_new_frame = True
                g.calc = "SCAN"      # 关键：输出出现扫描步即按扫描处理(即使 ! 行只写了 opt)
                g.cur_scan_step = int(sc.group(1))
                k = i
                while k < i + 10 and k + 1 < n:
                    k += 1
                    sv = re.search(
                        r"(Bond|Angle|Dihedral)\s*\(([^)]*)\)\s*:\s*(-?[\d\.]+)",
                        lines[k])
                    if sv:
                        g.scan_val = sv.group(1) + "(" + sv.group(2) + ")=" + sv.group(3)
                        break
        if is_new_frame:
            g.frames.append(Frame(
                scan_step=g.cur_scan_step, scan_val=g.scan_val if g.scan_val else ""))
            g.cur_frame = len(g.frames)
            absorb_idx = 0
            continue
        if g.cur_frame == 0:
            # SP/TD 单点没有帧头，也允许收集
            g.frames.append(Frame())
            g.cur_frame = 1

        fr = g.frames[g.cur_frame - 1] if g.cur_frame > 0 else None

        # ---- 当前帧几何 ----
        if fr and "CARTESIAN COORDINATES (ANGSTROEM)" in L and g.natoms:
            fr.geom = []
            j = i
            cnt = 0
            while cnt < g.natoms and j + 1 < n:
                j += 1
                tl = lines[j].strip()
                if tl == "" or tl.startswith("--"):
                    continue
                m2 = re.match(
                    r"^([A-Za-z]{1,3})\s+(-?[\d\.]+)\s+(-?[\d\.]+)\s+(-?[\d\.]+)$", tl)
                if m2:
                    fr.geom.append([m2.group(1),
                                    float(m2.group(2)),
                                    float(m2.group(3)),
                                    float(m2.group(4))])
                    cnt += 1
                elif cnt > 0:
                    break

        # ---- 能量 ----
        if fr:
            if "FINAL SINGLE POINT ENERGY" in L:
                e = re.search(r"FINAL SINGLE POINT ENERGY\s+(-?\d+\.?\d*)", L)
                if e:
                    fr.energy = e.group(1)
            if "E(tot)" in L:
                et = re.search(r"E\(tot\)\s+=\s+(-?\d+\.\d+)", L)
                if et:
                    fr.etot = et.group(1)
            if g.isTD and "DE(CIS)" in L:
                rt = re.search(r"DE\(CIS\)\s+=\s+-?\d+\.\d+\s+Eh\s+\(Root\s+(\d+)\)", L)
                if rt:
                    g.td_root = int(rt.group(1))

        # ---- 收敛判据 ----
        if fr and ("gradient" in L or "Energy change" in L or "step" in L):
            cv = re.search(
                r"\s(RMS gradient|MAX gradient|RMS step|MAX step|Energy change)"
                r"\s+(-?[\d\.]+)\s+([\d\.]+)\s+(YES|NO)", L, re.I)
            if cv:
                key = cv.group(1)
                v = cv.group(2)
                yn = 1 if cv.group(4) == "YES" else 0
                if key == "RMS gradient":
                    fr.conv.rmsG = v; fr.yesRG = yn
                elif key == "MAX gradient":
                    fr.conv.maxG = v; fr.yesMG = yn
                elif key == "RMS step":
                    fr.conv.rmsS = v; fr.yesRS = yn
                elif key == "MAX step":
                    fr.conv.maxS = v; fr.yesMS = yn
                elif key == "Energy change":
                    fr.conv.eChg = v; fr.yesE = yn

        # ---- 收敛标志 ----
        if "THE OPTIMIZATION HAS CONVERGED" in L and fr:
            fr.converged = True

        # ---- TD 激发态 ----
        if fr and "STATE" in L and "au" in L:
            st = re.search(
                r"STATE\s+(\d+):\s+E=\s+(-?\d+\.?\d*)\s+au\s+(-?\d+\.?\d*)\s*eV", L)
            if st:
                mult2 = ""
                s2v = ""
                ss = re.search(r"<S\*\*2>\s+=\s+([\d\.]+)", L)
                if ss:
                    s2v = ss.group(1)
                    mult2 = round(float(s2v) + 0.5)
                st_idx = int(st.group(1))
                dup = any(rec.idx == st_idx for rec in fr.td)
                if not dup:
                    fr.td.append(TDState(
                        idx=st_idx, ev=st.group(3), au=st.group(2),
                        s2=s2v, multW=mult2))
        if fr and fr.td and "->" in L:
            tr = re.match(
                r"^\s*(\d+[ab]?)\s*->\s*(\d+[ab]?)\s*:\s*([\d\.]+)"
                r"(?:\s*\(c=\s*(-?[\d\.]+)\))?", L)
            if tr:
                c3 = (tr.group(4) if tr.group(4) not in (None, "") else tr.group(3))
                fr.td[-1].trans.append(
                    [tr.group(1), tr.group(2), tr.group(3), c3])

        # ---- 吸收谱表(取 fosc) ----
        if "ABSORPTION SPECTRUM VIA TRANSITION ELECTRIC DIPOLE MOMENTS" in L:
            in_absorb = True
            absorb_idx = 0
            continue
        if in_absorb:
            if re.match(r"^\s*-+\s*$", L) or "CD SPECTRUM" in L or "VELOCITY DIPOLE" in L:
                if "VELOCITY DIPOLE" in L or "CD SPECTRUM" in L:
                    in_absorb = False
            else:
                ar = re.match(
                    r"^\s*(\d+)\s*-\s*\S+\s*->\s*(\d+)\s*-\s*\S+\s+(-?\d+\.?\d*)"
                    r"\s+(-?\d+\.?\d*)\s+(\d+\.?\d*)\s+([\d\.]+)", L)
                if ar:
                    if fr and fr.td:
                        absorb_idx += 1
                        if absorb_idx <= len(fr.td):
                            fr.td[absorb_idx - 1].ev = ar.group(3)
                            fr.td[absorb_idx - 1].nm = ar.group(5)
                            fr.td[absorb_idx - 1].fosc = ar.group(6)

        # ---- 频率 ----
        if "VIBRATIONAL FREQUENCIES" in L:
            g.freqs = []
            g.ir_int = []
            g.red_mass = []
            g.raman = []
            g.modes = []
            j = i
            while j + 1 < n:
                j += 1
                fm = re.match(r"^\s*(\d+):\s*(-?[\d\.]+)\s*cm", lines[j])
                if fm:
                    g.freqs.append(fm.group(2))
                    g.ir_int.append(0)
                    g.red_mass.append(0)
                    g.raman.append("")
                    g.modes.append([])
                if "NORMAL MODES" in lines[j]:
                    break
        if g.freqs and "IR SPECTRUM" in L:
            j = i
            while j + 1 < n:
                j += 1
                im = re.match(
                    r"^\s*(\d+):\s*(-?[\d\.]+)\s+([\d\.]+)\s+([\d\.]+)\s+", lines[j])
                if im:
                    mi = int(im.group(1)) + 1
                    if mi <= len(g.freqs):
                        g.ir_int[mi - 1] = im.group(4)
                if "NORMAL MODES" in lines[j] or "RAMAN SPECTRUM" in lines[j]:
                    break
        if g.freqs and "RAMAN SPECTRUM" in L:
            j = i
            while j + 1 < n:
                j += 1
                rm = re.match(
                    r"^\s*(\d+):\s*(-?[\d\.]+)\s+([\d\.]+)\s+([\d\.]+)\s+", lines[j])
                if rm:
                    mi = int(rm.group(1)) + 1
                    if mi <= len(g.freqs):
                        g.raman[mi - 1] = rm.group(4)
                if "NORMAL MODES" in lines[j] or "THERMOCHEMISTRY" in lines[j]:
                    break

        # ---- 热化学 ----
        if "Eh" in L:
            ee = re.search(r"Electronic energy\s+\.*\s+(-?\d+\.\d+)\s+Eh", L)
            if ee:
                g.thermo["Eel"] = ee.group(1)
            zp = re.search(r"Zero point energy\s+\.*\s+(-?\d+\.\d+)\s+Eh", L)
            if zp:
                g.thermo["ZPE"] = zp.group(1)
            tt = re.search(r"Total thermal correction\s+(-?\d+\.\d+)\s+Eh", L)
            if tt:
                g.thermo["TE"] = tt.group(1)
            tu = re.search(r"Total thermal energy\s+\.*\s+(-?\d+\.\d+)\s+Eh", L)
            if tu:
                g.thermo["U"] = tu.group(1)
            th = re.search(r"Total Enthalpy\s+\.*\s+(-?\d+\.\d+)\s+Eh", L)
            if th:
                g.thermo["H"] = th.group(1)
            fg = re.search(r"Final Gibbs free energy\s+\.*\s+(-?\d+\.\d+)\s+Eh", L)
            if fg:
                g.thermo["G"] = fg.group(1)

        # ---- 正常结束标记 ----
        if "ORCA TERMINATED NORMALLY" in L:
            g.normal_end = True

        # ---- 原子电荷/自旋布居 (写入当前帧, pop=always) ----
        if "MULLIKEN ATOMIC CHARGES" in L:
            g.mulC = []
            g.mulS = []
            g.hasMulSpin = "SPIN" in L or "SPIN POP" in L
            j = i
            while j + 1 < n:
                j += 1
                tl = lines[j].strip()
                if tl == "" or tl.startswith("--"):
                    continue
                if "Sum of atomic" in tl or "REDUCED" in tl:
                    break
                mc = re.match(
                    r"^(\d+)\s+([A-Za-z]{1,3})\s*:\s*(-?\d+\.\d+)"
                    r"(?:\s+(-?\d+\.\d+))?", tl)
                if mc:
                    g.mulC.append(mc.group(3))
                    g.mulS.append(mc.group(4) if mc.group(4) not in (None, "") else 0.0)
            if fr is not None:
                fr.mulC = list(g.mulC)
                fr.mulS = list(g.mulS)
                fr.hasMulSpin = g.hasMulSpin
        if "LOEWDIN ATOMIC CHARGES" in L:
            g.lowC = []
            g.lowS = []
            g.hasLowSpin = "SPIN" in L or "SPIN POP" in L
            j = i
            while j + 1 < n:
                j += 1
                tl = lines[j].strip()
                if tl == "" or tl.startswith("--"):
                    continue
                if "Sum of atomic" in tl or "REDUCED" in tl:
                    break
                lc = re.match(
                    r"^(\d+)\s+([A-Za-z]{1,3})\s*:\s*(-?\d+\.\d+)"
                    r"(?:\s+(-?\d+\.\d+))?", tl)
                if lc:
                    g.lowC.append(lc.group(3))
                    g.lowS.append(lc.group(4) if lc.group(4) not in (None, "") else 0.0)

    # ---- 解析最后一个 NORMAL MODES 段的位移矩阵 ----
    parse_normal_modes(lines, n, g)
    return g


# ---------------------------------------------------------------------------
# 正则模式 (NORMAL MODES)
# ---------------------------------------------------------------------------
def parse_normal_modes(lines, n, g):
    start = -1
    for idx in range(n):
        if "These modes are the Cartesian" in lines[idx]:
            start = idx
    if start == -1:
        g.modes = []
        return
    end = n
    for j in range(start + 1, n):
        if any(k in lines[j] for k in (
                "IR SPECTRUM", "RAMAN SPECTRUM", "THERMOCHEMISTRY",
                "ORCA GEOMETRY RELAXATION", "VIBRATIONAL FREQUENCIES")):
            end = j
            break
    g.modes = []
    j = start
    while j + 1 < end:
        j += 1
        L = lines[j]
        if not re.match(r"^\s+\d+(\s+\d+)*\s*$", L):
            continue
        col_idx = [int(x) for x in L.strip().split() if x != ""]
        k = j
        while k + 1 < end:
            k += 1
            dl = lines[k]
            dm0 = re.match(r"^\s+(\d+)\s+-?[\d\.]+", dl)
            if not dm0:
                break
            comp = int(dm0.group(1))
            rest = re.sub(r"^\s+\d+\s*", "", dl)
            vals = [float(tp) for tp in rest.split() if tp != "" and isnum(tp)]
            if len(vals) < len(col_idx):
                break
            a_idx = comp // 3 + 1
            xyz_i = comp % 3 + 1
            for ci, mv in enumerate(vals):
                gi = col_idx[ci] + 1
                while len(g.modes) < gi:
                    g.modes.append([])
                arr = g.modes[gi - 1]
                while len(arr) < a_idx:
                    arr.append([0.0, 0.0, 0.0])
                arr[a_idx - 1][xyz_i - 1] = mv
        j = k - 1


# ---------------------------------------------------------------------------
# 写出调度
# ---------------------------------------------------------------------------
def ask_and_write(out_file, src_file, cli_mode=False, scan_choice=None):
    g = _GLOBAL_G
    scan_keep_all = False
    if g.calc == "SCAN":
        explicit = (g.scan_mode in ("all", "last"))
        if not explicit:
            if scan_choice:
                val = str(scan_choice).lower()
                if val == "cancel":
                    sys.exit(0)
                scan_keep_all = (val == "all")
                g.scan_mode = "all" if val == "all" else "last"
            else:
                r = input(L(
                    "检测到 RELAXED SURFACE SCAN。\n"
                    "请选择输出方式:\n"
                    "  1) all  —— 保留所有扫描帧 (生成 opt 轨迹 log)\n"
                    "  2) last —— 仅保留每个扫描步的最后一帧\n"
                    "  3) cancel —— 中止转换\n"
                    "选择 [1/2/3]: ",
                    "Relaxed surface scan detected.\n"
                    "Choose output mode:\n"
                    "  1) all  — keep all frames (opt trajectory log)\n"
                    "  2) last — keep only the last frame of each step\n"
                    "  3) cancel — abort conversion\n"
                    "Choice [1/2/3]: ")).strip().lower()
                if r in ("cancel", "c", "3"):
                    sys.exit(0)
                scan_keep_all = (r in ("all", "a", "1"))
                g.scan_mode = "all" if scan_keep_all else "last"
        else:
            scan_keep_all = (g.scan_mode == "all")

    frames = g.frames
    if g.calc == "SCAN":
        clean = [f for f in g.frames if f.scan_step > 0]
        if not clean:
            clean = g.frames
        if scan_keep_all:
            frames = clean
        else:
            last_of = {}
            last_conv = {}
            order = []
            for f in clean:
                st = f.scan_step
                has_en = (f.energy != "" or f.etot != "")
                if st not in last_of:
                    order.append(st)
                    last_of[st] = None
                    last_conv[st] = None
                if has_en:
                    last_of[st] = f
                if f.converged and has_en:
                    last_conv[st] = f
            frames = [last_conv[st] if last_conv[st] is not None else last_of[st]
                      for st in order]

    write_log(out_file, src_file, frames)


# ---------------------------------------------------------------------------
# 写日志 (WriteLog)
# ---------------------------------------------------------------------------
def write_log(out_file, src_file, frames):
    g = _GLOBAL_G
    o = []

    nelec = sum(int(z) for z in g.nums) - g.charge
    unpaired = g.mult - 1
    na = (nelec + unpaired) // 2
    nb = (nelec - unpaired) // 2

    src_name = Path(src_file).name

    # ---- GaussView 识别并行/内存设置所需的头部 (对应 AcO_SP.log) ----
    o.append(" Entering Gaussian System, Link 0=g16")
    o.append(" Initial command:")
    o.append(" /opt/app/gaussian/16-C01-avx2/g16/l1.exe \"/tmp/Gau-%d.inp\" -scrdir=\"/tmp/\"" % (os.getpid() % 999999))
    o.append(" Entering Link 1 = /opt/app/gaussian/16-C01-avx2/g16/l1.exe PID=     %d." % os.getpid())
    o.append("")

    o.append(" This file was generated by orca2gaussian.py (Python 3)")
    o.append(" Source ORCA output: " + src_name)

    # ---- route (逐字保留 ORCA 关键字, 不做转化) ----
    route = "# "
    route += (g.functional if g.functional != "" else "HF") + "/" + \
             (g.basis if g.basis != "" else "Gen")
    # 逐字追加其余 ORCA 关键字 (optts/scants/freq/smd(...) 等原样保留)
    for wv in g.kw_verbatim:
        route += " " + wv

    o.append(" ******************************************")
    o.append(" Gaussian 16:  ES64L-G16RevC.01 24-Jan-2017")
    o.append("                 1-Jan-2026 ")
    o.append(" ******************************************")
    if g.nprocs:
        o.append(f"%nprocshared={g.nprocs}")
        o.append(" Will use up to   " + str(g.nprocs) + " processors via shared memory.")
    if g.maxcore:
        # ORCA %maxcore 是每核内存；Gaussian %mem 是总内存 => 相乘再转 GB
        total_mb = int(g.maxcore) * (g.nprocs if g.nprocs else 1)
        o.append(f"%mem={maxcore_gb(total_mb)}GB")
    o.append(" ----------------------------------------------------------------------")
    o.append(" " + route)
    o.append(" ----------------------------------------------------------------------")
    o.append(" 1/38=1,158=3,167=1,172=1/1;")
    o.append(" 2/12=2,17=6,18=5,40=1/2;")
    o.append(" 3/5=7,11=2,16=1,17=8,25=1,27=10,30=1,70=32201,72=2,74=-58,75=-4,116=1,158=3/1,2,3;")
    o.append(" 4/60=-1/1;")
    o.append(" 5/5=2,7=64,8=3,13=1,38=5,53=2,87=10/2,8;")
    o.append(" 6/7=2,8=2,9=2,10=2,28=1,87=10/1;")
    o.append(" 99/5=1,9=1/99;")
    o.append("  using 2006 physical constants.")
    o.append(" ----------")
    o.append(" Title Card")
    o.append(" ----------")
    o.append(" Symbolic Z-matrix:")
    o.append(" Charge = " + padl(g.charge, 2) + " Multiplicity = " + padl(g.mult, 2))

    # Z-matrix 坐标回显 (符号 + xyz), 取自第一个含几何的帧
    zm_geo = []
    for frz in frames:
        if getattr(frz, "geom", None) and frz.geom:
            zm_geo = frz.geom
            break
    if zm_geo:
        for a in zm_geo:
            o.append(padl(a[0], 2) + " " + padl(f"{float(a[1]):.5f}", 12)
                     + " " + padl(f"{float(a[2]):.5f}", 12)
                     + "  " + padl(f"{float(a[3]):.5f}", 12))
    o.append("")

    # 约束/扫描回显
    cons_txt = []
    for c in g.constraints:
        if c["wild"]:
            cons_txt.append("! " + c["typ"] + " (含通配符的约束, 见 ORCA 输入)")
            continue
        ids = ""
        for ix in c["idx"]:
            ids += " " + str(ix + 1)
        val_part = (" " + c["val"]) if c["val"] != "" else ""
        cons_txt.append(c["typ"] + ids + val_part + " F")
    if g.scan_def != "":
        sd = g.scan_def
        start = float(sd["start"])
        stop = float(sd["stop"])
        steps = int(sd["steps"])
        inc = (stop - start) / max(steps - 1, 1)
        cons_txt.append("! Relaxed scan: " + sd["typ"] + " " + sd["atoms"]
                        + " = " + f"{start:.6f}" + " -> " + f"{stop:.6f}"
                        + " (" + str(steps) + " points, step " + f"{inc:.6f}" + ")")
        cons_txt.append(sd["typ"] + " " + sd["atoms"] + " S " + str(steps)
                        + " " + f"{start:.6f}" + " " + f"{inc:.6f}")
    if cons_txt:
        o.append("")
        for ct in cons_txt:
            o.append(" ! ModRedundant/scan (from ORCA, atoms 1-based): " + ct)

    # ---- 帧 ----
    good = [f for f in frames if f.energy != "" or f.etot != ""]
    if good:
        frames = good
    nf = len(frames)
    if nf == 0:
        dbg("ERROR: no usable geometry frames")
        sys.exit(1)

    any_conv = any(f.converged for f in frames)

    fi = 0
    for f in frames:
        fi += 1
        write_orientation(o, f.geom, True)
        write_orientation(o, f.geom)
        en = f.etot if (g.isTD and f.etot != "") else f.energy
        meth = "UKS" if g.mult > 1 else "RKS"
        o.append(" SCF Done:  E(" + meth + ") =  " + f"{float(en):.10f}"
                 + "     A.U. after    1 cycles")

        if getattr(f, "mulC", None) and f.mulC:
            write_charges(o, f.mulC, f.mulS, f.hasMulSpin)

        if g.isTD and f.td:
            write_excited(o, f)
        o.append("")

        if g.calc in ("OPT", "SCAN"):
            o.append(gradbar())
            o.append(" Berny optimization.")
            o.append(" Internal  Forces:  Max " + f"{val0(f.conv.maxG):.6f}"
                     + " RMS " + f"{val0(f.conv.rmsG):.6f}")
            o.append(" Search for a local " + ("maximum." if g.isTS else "minimum."))
            o.append(" Step number   " + str(fi) + " out of a maximum of  " + str(nf + 10))
            o.append(" All quantities printed in internal units (Hartrees-Bohrs-Radians)")
            write_conv_items(o, f)
            o.append(gradbar())
            if g.calc == "SCAN" and f.scan_val != "" and fi < nf:
                o.append(" ! Next scan point: " + frames[fi].scan_val)
            o.append("")

    if any_conv and g.calc in ("OPT", "SCAN"):
        o.append(" Optimization completed.")
        o.append(" -- Stationary point found.")
        o.append("")

    if g.hasFreq and g.freqs:
        write_freqs(o)

    write_thermo(o)

    if g.normal_end:
        o.append(" Normal termination of Gaussian 16 (fake)")
    else:
        o.append(" Error termination request processed by link 9999.")
        o.append(" Error termination of Gaussian 16 (fake).")
        o.append(" (Source ORCA output did NOT terminate normally)")

    txt = "\n".join(o) + "\n"
    if os.path.exists(out_file):
        os.remove(out_file)
    with open(out_file, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(txt)

    dbg("OK frames=" + str(len(frames)) + " freqs=" + str(len(g.freqs)) + " type=" + g.calc)


def val0(v):
    try:
        return float(v)
    except (ValueError, TypeError):
        return 0.0


# ---------------------------------------------------------------------------
# 收敛判据块
# ---------------------------------------------------------------------------
def write_conv_items(o, f):
    lab1 = " Maximum Force            "
    lab2 = " RMS     Force            "
    lab3 = " Maximum Displacement     "
    lab4 = " RMS     Displacement     "
    lab0 = " Energy Change            "
    thrF = "     0.000300     "
    thrR = "     0.000100     "
    thrD = "     0.004000     "
    thrS = "     0.002000     "
    thrE = "     0.000005     "
    ynE = "YES" if f.yesE else "NO "
    ynM = "YES" if f.yesMG else "NO "
    ynR = "YES" if f.yesRG else "NO "
    ynDS = "YES" if f.yesMS else "NO "
    ynRS = "YES" if f.yesRS else "NO "
    vE = val0(f.conv.eChg)
    vM = val0(f.conv.maxG)
    vR = val0(f.conv.rmsG)
    vD = val0(f.conv.maxS)
    vS = val0(f.conv.rmsS)
    if f.conv.eChg == "" and f.conv.maxG == "":
        return
    o.append("         Item               Value     Threshold  Converged?")
    if f.conv.eChg != "":
        o.append(lab0 + padl(f"{vE:.6f}", 9) + thrE + ynE)
    o.append(lab1 + padl(f"{vM:.6f}", 9) + thrF + ynM)
    o.append(lab2 + padl(f"{vR:.6f}", 9) + thrR + ynR)
    o.append(lab3 + padl(f"{vD:.6f}", 9) + thrD + ynDS)
    o.append(lab4 + padl(f"{vS:.6f}", 9) + thrS + ynRS)


# ---------------------------------------------------------------------------
# orientation 块
# ---------------------------------------------------------------------------
def write_orientation(o, geom, input_=False):
    g = _GLOBAL_G
    if input_:
        o.append("                          Input orientation:                          ")
    else:
        o.append("                         Standard orientation:                         ")
    o.append(" ---------------------------------------------------------------------")
    o.append(" Center     Atomic      Atomic             Coordinates (Angstroms)")
    o.append(" Number     Number       Type             X           Y           Z")
    o.append(" ---------------------------------------------------------------------")
    ai = 0
    for a in geom:
        ai += 1
        z = g.nums[ai - 1] if ai <= len(g.nums) else atom_num(a[0])
        o.append(" " + padl(ai, 6) + padl(z, 11) + padl(0, 12)
                 + padl(f"{float(a[1]):.6f}", 16) + padl(f"{float(a[2]):.6f}", 12)
                 + padl(f"{float(a[3]):.6f}", 12))
    o.append(" ---------------------------------------------------------------------")


# ---------------------------------------------------------------------------
# 激发态块
# ---------------------------------------------------------------------------
MULT_WORDS = {1: "Singlet", 2: "Doublet", 3: "Triplet", 4: "Quartet",
              5: "Quintet", 6: "Sextet", 7: "Septet", 8: "Octet"}


def write_excited(o, f):
    g = _GLOBAL_G
    o.append(" Excitation energies and oscillator strengths:")
    o.append("")
    for st in f.td:
        mw = MULT_WORDS.get(st.multW, "") if st.multW != "" else "Singlet"
        if mw == "":
            mw = "Singlet"
        sym = "'" if (st.s2 != "" and float(st.s2) + 0 > 0.3) else ""
        ev = float(st.ev) if st.ev != "" else 0.0
        lam = 1239.841984 / abs(ev) if ev > 0 else 1
        s2part = ("      <S**2>=" + f"{float(st.s2):.3f}") if st.s2 != "" else ""
        o.append(" Excited State " + fmt(st.idx, 4, 0) + ":  " + mw + "-A" + sym
                 + fmt(ev, 11, 4) + " eV" + fmt(lam, 9, 2) + " nm  f="
                 + f"{float(st.fosc):.4f}" + s2part)
        ti = 0
        for tr in st.trans:
            ti += 1
            if ti > 24:
                o.append("        ... (more contributions omitted)")
                break
            c = abs(val0(tr[3]))
            o.append("      " + tr[0].upper() + " -> " + tr[1].upper()
                     + "        " + f"{c:.5f}")
    o.append("")
    if g.calc in ("OPT", "SCAN"):
        o.append(" This state for optimization and/or subsequent second-order properties.")
        o.append(" The state that is being optimized has been set to state number   "
                 + str(g.td_root))
    o.append("")


# ---------------------------------------------------------------------------
# 频率块
# ---------------------------------------------------------------------------
def write_freqs(o):
    g = _GLOBAL_G
    o.append(" Harmonic frequencies (cm**-1), IR intensities (KM/Mole), Raman scattering")
    o.append(" activities (A**4/AMU), depolarization ratios for plane and unpolarized")
    o.append(" incident light, reduced masses (AMU), force constants (mDyne/A),")
    o.append(" and normal coordinates:")

    n = len(g.freqs)
    start_i = 6  # skip 6 trans/rot
    if start_i >= n:
        start_i = 0
    mi = start_i
    while mi < n:
        hi = min(mi + 2, n - 1)
        l1 = ""
        for i in range(mi, hi + 1):
            l1 += padl(i - 5, 23)
        l2 = ""
        for _ in range(mi, hi + 1):
            l2 += padl("A", 23)
        lf = " Frequencies --"
        li = " IR Inten    --"
        for i in range(mi, hi + 1):
            w1 = 12 if i == mi else 23
            lf += padl(f"{float(g.freqs[i]):.4f}", w1)
            li += padl(f"{float(g.ir_int[i]):.4f}", w1)
        o.append("")
        o.append(l1)
        o.append(l2)
        o.append(lf)
        o.append(li)
        o.append("  Atom  AN      X      Y      Z        X      Y      Z        X      Y      Z")
        for at in range(g.natoms):
            row = padl(at + 1, 6) + padl(g.nums[at], 4 if g.nums else 4)
            for i in range(mi, hi + 1):
                disp = [0.0, 0.0, 0.0]
                if i + 1 <= len(g.modes) and at + 1 <= len(g.modes[i]):
                    disp = g.modes[i][at]
                row += padl(f"{float(disp[0]):.2f}", 9) \
                     + padl(f"{float(disp[1]):.2f}", 7) \
                     + padl(f"{float(disp[2]):.2f}", 7)
            o.append(row)
        mi = hi + 1
    o.append("")


# ---------------------------------------------------------------------------
# 电荷块
# ---------------------------------------------------------------------------
def write_charges(o, mul_c=None, mul_s=None, has_mul_spin=None):
    g = _GLOBAL_G
    if mul_c is None:
        mul_c = g.mulC
        mul_s = g.mulS
        has_mul_spin = g.hasMulSpin
    if not mul_c or len(mul_c) == 0:
        return
    na_ = len(mul_c)
    if has_mul_spin:
        o.append(" Mulliken charges and spin densities:")
        o.append(padl(1, 16) + padl(2, 10))
        for idx in range(1, na_ + 1):
            sym = g.syms[idx - 1] if idx <= len(g.syms) else "X"
            o.append(padl(idx, 6) + "  " + sym + "   " + f"{float(mul_c[idx-1]):.6f}"
                     + "   " + f"{float(mul_s[idx-1]):.6f}")
    else:
        o.append(" Mulliken charges:")
        o.append(padl(1, 16))
        for idx in range(1, na_ + 1):
            sym = g.syms[idx - 1] if idx <= len(g.syms) else "X"
            o.append(padl(idx, 6) + "  " + sym + "   " + f"{float(mul_c[idx-1]):.6f}")
    o.append(" Sum of Mulliken charges =  " + f"{sum(float(x) for x in mul_c):.5f}")
    o.append("")


# ---------------------------------------------------------------------------
# 热化学块
# ---------------------------------------------------------------------------
def write_thermo(o):
    g = _GLOBAL_G
    t = g.thermo
    if "ZPE" not in t:
        return
    eel = float(t.get("Eel", 0))
    te = float(t.get("TE", 0))
    if "U" in t:
        tc_e = float(t["U"]) - eel
    else:
        tc_e = te + float(t["ZPE"])
    if "H" in t:
        tc_h = float(t["H"]) - eel
    else:
        tc_h = tc_e + 0.000944
    if "G" in t:
        tc_g = float(t["G"]) - eel
    else:
        tc_g = tc_h + 0
    o.append(" Temperature   298.150 Kelvin.  Pressure   1.00000 Atm.")
    o.append(" Zero-point correction=                           "
             + fmt(float(t["ZPE"]), 12, 6) + " Hartree")
    o.append(" Thermal correction to Energy=                    " + fmt(tc_e, 12, 6))
    o.append(" Thermal correction to Enthalpy=                  " + fmt(tc_h, 12, 6))
    o.append(" Thermal correction to Gibbs Free Energy=         " + fmt(tc_g, 12, 6))
    o.append(" Sum of electronic and zero-point Energies=       "
             + fmt(eel + float(t["ZPE"]), 16, 6))
    o.append(" Sum of electronic and thermal Energies=          " + fmt(eel + tc_e, 16, 6))
    o.append(" Sum of electronic and thermal Enthalpies=        " + fmt(eel + tc_h, 16, 6))
    o.append(" Sum of electronic and thermal Free Energies=     " + fmt(eel + tc_g, 16, 6))
    o.append("")


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------
_GLOBAL_G = Run()


def main(argv=None):
    if argv is None:
        argv = sys.argv[1:]

    parser = argparse.ArgumentParser(
        prog="orca2gaussian.py",
        description="Convert ORCA .out to a Gaussian16-fake log readable by GaussView.",
    )
    parser.add_argument("input", help="ORCA output .out file")
    parser.add_argument("mode", nargs="?", choices=["all", "last"],
                        help="scan mode: all (keep all frames) or last (last of each step)")
    parser.add_argument("--yes", action="store_true",
                        help="non-interactive: if scan and no mode given, default to 'last'")
    args = parser.parse_args(argv)

    f = args.input
    base = re.sub(r"\.out$", "", f)
    out_file = base + "_fake.log"
    if args.mode and re.search(r"scan", f, re.I):
        m = args.mode.lower()
        if m == "all":
            out_file = base + "_scanall_fake.log"
        elif m == "last":
            out_file = base + "_scanlast_fake.log"

    dbg("[1] parsing " + f)
    g = parse_orca(f)
    _GLOBAL_G.__dict__.update(g.__dict__)

    dbg("[2] calc=" + g.calc + " frames=" + str(len(g.frames))
        + " natoms=" + str(g.natoms) + " nbasis=" + str(g.nbasis)
        + " q=" + str(g.charge) + " m=" + str(g.mult))

    # scan 模式须写到 _GLOBAL_G (ask_and_write 读取它)
    if args.mode and _GLOBAL_G.calc == "SCAN":
        _GLOBAL_G.scan_mode = args.mode.lower()

    scan_choice = None
    if (_GLOBAL_G.calc == "SCAN" and _GLOBAL_G.scan_mode not in ("all", "last")
            and args.yes):
        scan_choice = "last"

    try:
        ask_and_write(out_file, f, cli_mode=True, scan_choice=scan_choice)
    except SystemExit:
        raise
    except Exception as err:
        dbg("ERR: " + str(err))
        print("ERR: " + str(err), file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
