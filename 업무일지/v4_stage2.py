# -*- coding: utf-8 -*-
"""Stage 2 — 사업별 현황 (예산 집행 · 단계별 진척 · 추진 타임라인)"""
import sys; sys.path.insert(0, "/home/user/live-trade/업무일지")
from v4_common import *

wb = openpyxl.load_workbook(OUT)
NT = 60                                    # 타임라인 표시 건수
PLB = R("사업추진일지", "B", PL1, PL2)
PLC = R("사업추진일지", "C", PL1, PL2)
PLE = R("사업추진일지", "E", PL1, PL2)
PLI = R("사업추진일지", "I", PL1, PL2)
PM_A = R("사업마스터", "A", PJ1, PJ2)

ps = fresh_sheet(wb, "사업별 현황")
ps.sheet_view.showGridLines = False
title(ps, "사업별 현황  ·  추진 경과 추적",
      "※ 노란 칸에서 사업번호를 고르면 그 사업이 언제 무슨 일을 했는지 전부 정리됩니다.  "
      "※ 입력할 것은 사업번호 하나뿐이며 나머지는 자동입니다.  "
      "※ 예산이 없는 사업은 집행률·잔액이 빈칸으로 남습니다.")
widths(ps, {"A":14,"B":13,"C":14,"D":34,"E":30,"F":12,"G":15,"H":18,"I":11,
            "J":3,"K":3,"L":3})
put(ps, "A4", "사업번호 선택 →", LBL_F, fill=SUB_FILL, align="left")
put(ps, "B4", None, IN_F, fill=KEY_FILL)
dv_name(ps, "B4", "목록_사업번호")

MT = f'MATCH($B$4,{PM_A},0)'
def _pm(col, miss='"※ 미등록 사업번호"'):
    return f'=IF($B$4="","",IFERROR(INDEX({R("사업마스터", col, PJ1, PJ2)},{MT}),{miss}))'

hdr(ps, 6, ["사업명","연도","사업 구분","주관팀","총괄 담당자","진행상태","현재 단계","진척률"])
put(ps, "A7", _pm("B"), LNK_F, align="left")
put(ps, "B7", _pm("C", '""'), LNK_F, "0")
put(ps, "C7", _pm("D", '""'), LNK_F)
put(ps, "D7", _pm("E", '""'), LNK_F)
put(ps, "E7", _pm("F", '""'), LNK_F)
put(ps, "F7", _pm("P", '""'), LNK_F)
put(ps, "G7", _pm("M", '""'), LNK_F)
put(ps, "H7", _pm("N", '0'), LNK_F, PCT0)

# ── 정렬 키 (숨김)
for i in range(1, (PL2 - PL1 + 1) + 1):
    s = PL1 + i - 1
    ps[f"BA{i}"] = (f'=IF(AND(사업추진일지!$C{s}=$B$4,사업추진일지!$B{s}<>""),'
                    f'사업추진일지!$B{s}*10000+{i},"")')
NP = PL2 - PL1 + 1

# ── 1. 예산 집행
section(ps, 9, "1.  예산 집행", span=8)
hdr(ps, 10, ["예산액(원)","집행액(원)","잔액(원)","집행률","추진 건수",
             "최초 기록일","최종 기록일","경과일"])
ps["BD11"] = f'=IF($B$4="",0,SUMPRODUCT(MAX(({PLC}=$B$4)*({PLB}>0)*(1000000-{PLB}))))'
ps["BD12"] = f'=IF($B$4="",0,SUMPRODUCT(MAX(({PLC}=$B$4)*{PLB})))'
put(ps, "A11", _pm("G", '""'), LNK_F, MONEY)
put(ps, "B11", _pm("H", '""'), LNK_F, MONEY)
put(ps, "C11", _pm("J", '""'), LNK_F, MONEY)
put(ps, "D11", _pm("I", '""'), LNK_F, PCT)
put(ps, "E11", _pm("O", '""'), LNK_F, NUM)
put(ps, "F11", '=IF($BD$11=0,"",1000000-$BD$11)', FX_F, DATE_)
put(ps, "G11", '=IF($BD$12=0,"",$BD$12)', FX_F, DATE_)
put(ps, "H11", '=IF(OR($F11="",$G11=""),"",$G11-$F11+1)', FX_F, NUM)
ps["A12"] = (f'=IF($B$4="","",IF(IFERROR(INDEX({R("사업마스터","G",PJ1,PJ2)},{MT}),"")="",'
             f'"※ 예산이 배정되지 않은 사업입니다 — 집행률·잔액은 표시되지 않습니다.",""))')
ps["A12"].font = Font(name=F, size=9, bold=True, color="9C5700")

# ── 2. 단계별 추진 현황
section(ps, 14, "2.  단계별 추진 현황", span=6)
hdr(ps, 15, ["추진 단계","기록 건수","집행액(원)","최초 일자","최종 일자","도달"])
NS = PS2 - PS1 + 1
for i in range(NS):
    r, cr = 16 + i, PS1 + i
    stg = f'$A{r}'
    put(ps, f"A{r}", f'=IF(코드설정!$S{cr}="","",코드설정!$S{cr})', FX_F, fill=SUB_FILL)
    g = f'IF(OR($B$4="",{stg}=""),"",'
    put(ps, f"B{r}", f'={g}COUNTIFS({PLC},$B$4,{PLE},{stg}))', LNK_F, NUM)
    put(ps, f"C{r}", f'={g}SUMIFS({PLI},{PLC},$B$4,{PLE},{stg}))', LNK_F, MONEY)
    ps[f"BD{r}"] = (f'=IF(OR($B$4="",{stg}=""),0,'
                    f'SUMPRODUCT(MAX(({PLC}=$B$4)*({PLE}={stg})*({PLB}>0)*(1000000-{PLB}))))')
    ps[f"BE{r}"] = f'=IF(OR($B$4="",{stg}=""),0,SUMPRODUCT(MAX(({PLC}=$B$4)*({PLE}={stg})*{PLB})))'
    put(ps, f"D{r}", f'=IF($BD{r}=0,"",1000000-$BD{r})', FX_F, DATE_)
    put(ps, f"E{r}", f'=IF($BE{r}=0,"",$BE{r})', FX_F, DATE_)
    put(ps, f"F{r}", f'={g}IF($B{r}>0,"○","-"))', FX_F)
R_STOT = 16 + NS
put(ps, f"A{R_STOT}", "합계", LBL_F, fill=TOT_FILL)
put(ps, f"B{R_STOT}", f'=SUM(B16:B{R_STOT-1})', FX_F, NUM, TOT_FILL)
put(ps, f"C{R_STOT}", f'=SUM(C16:C{R_STOT-1})', FX_F, MONEY, TOT_FILL)
for L in "DEF": put(ps, f"{L}{R_STOT}", "", FX_F, fill=TOT_FILL)
ps[f"F{R_STOT}"] = f'=IF($B$4="","",COUNTIF(F16:F{R_STOT-1},"○")&" / "&{NS})'
ps[f"F{R_STOT}"].font = FX_F
ps[f"A{R_STOT+1}"] = ("※ 단계를 적지 않은 기록은 단계별 표에 잡히지 않습니다. 아래 타임라인에는 모두 나옵니다 "
                      "— 그래서 단계별 합계가 전체 추진 건수보다 적을 수 있습니다.")
ps[f"A{R_STOT+1}"].font = NOTE_F
ps["F15"].comment = Comment(
    "그 단계에 기록이 하나라도 있으면 ○ 입니다. 합계 칸은 도달한 단계 수 / 전체 단계 수입니다.", "업무일지")

# ── 3. 추진 타임라인
R_TL = R_STOT + 3
section(ps, R_TL, f"3.  추진 타임라인 (날짜순 · 상위 {NT}건)", span=9)
hdr(ps, R_TL + 1, ["순번","일자","추진 단계","추진 내용","조치 · 결과",
                   "담당자","집행액(원)","증빙 · 문서번호","진행상태"])
def _pi(col, ref): return f'INDEX({R("사업추진일지", col, PL1, PL2)},${ref})'
for k in range(1, NT + 1):
    r = R_TL + 1 + k
    ps[f"BB{r}"] = f'=IFERROR(MATCH(SMALL($BA$1:$BA${NP},{k}),$BA$1:$BA${NP},0),"")'
    g = f'IF($BB{r}="","",'
    put(ps, f"A{r}", f'={g}{k})', FX_F, NUM)
    put(ps, f"B{r}", f'={g}{_pi("B", f"BB{r}")})', LNK_F, DATE_)
    put(ps, f"C{r}", f'={g}{_pi("E", f"BB{r}")})', LNK_F)
    put(ps, f"D{r}", f'={g}{_pi("F", f"BB{r}")})', LNK_F, align="left")
    put(ps, f"E{r}", f'={g}{_pi("G", f"BB{r}")})', LNK_F, align="left")
    put(ps, f"F{r}", f'={g}{_pi("H", f"BB{r}")})', LNK_F)
    put(ps, f"G{r}", f'={g}{_pi("I", f"BB{r}")})', LNK_F, MONEY)
    put(ps, f"H{r}", f'={g}{_pi("J", f"BB{r}")})', LNK_F)
    put(ps, f"I{r}", f'={g}{_pi("K", f"BB{r}")})', LNK_F)
_T1, _T2 = R_TL + 2, R_TL + 1 + NT
ps.conditional_formatting.add(f"A{_T1}:I{_T2}", FormulaRule(
    formula=[f'$I{_T1}="완료"'], fill=PatternFill("solid", fgColor="E2EFDA"), stopIfTrue=False))
ps.conditional_formatting.add(f"A{_T1}:I{_T2}", FormulaRule(
    formula=[f'$I{_T1}="보류"'], fill=PatternFill("solid", fgColor="FCE4D6"), stopIfTrue=False))
ps[f"A{_T2+2}"] = f"※ 타임라인은 상위 {NT}건까지 표시합니다. 그보다 많으면 [사업추진일지]에서 사업번호로 필터를 거세요."
ps[f"A{_T2+2}"].font = NOTE_F
for j in range(10, 60):
    ps.column_dimensions[CL(j)].hidden = True

wb.save(OUT)
print(f"Stage 2 완료 — 사업별 현황 (단계 {NS} / 타임라인 {NT}행)")
