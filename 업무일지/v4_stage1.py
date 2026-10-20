# -*- coding: utf-8 -*-
"""Stage 1 — 코드설정 확장 + 사업마스터 + 사업추진일지"""
import sys; sys.path.insert(0, "/home/user/live-trade/업무일지")
from v4_common import *

wb = openpyxl.load_workbook(SRC)
cs = wb["코드설정"]

# ══════════════════════ 코드설정: 사업 구분 · 추진 단계 추가
NEW_LISTS = [
    ("Q", "사업 구분", ["시스템 구축", "장비 도입", "교육·훈련 사업", "연구·용역",
                       "시설 개선", "정보화 사업", "국제 협력", "기타 사업"], 18, SG2),
    ("S", "추진 단계", ["계획", "예산확보", "계약·발주", "이행", "검사·납품", "정산·종료"], 14, PS2),
]
for col, name, vals, w, lim in NEW_LISTS:
    c = cs[f"{col}4"]
    c.value, c.font, c.fill, c.border = name, HDR_F, HDR_FILL, BOX
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r in range(5, lim + 1):
        cell = cs[f"{col}{r}"]
        if r - 5 < len(vals): cell.value = vals[r - 5]
        cell.font, cell.border = IN_F, BOX
        cell.alignment = Alignment(horizontal="center", vertical="center")
    cs.column_dimensions[col].width = w
    cs.column_dimensions[CL(openpyxl.utils.column_index_from_string(col) + 1)].width = 3
# T열: 단계별 진척률 (S와 같은 행끼리 짝)
c = cs["T4"]
c.value, c.font, c.fill, c.border = "단계 진척률", HDR_F, HDR_FILL, BOX
c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
for i, v in enumerate([0.10, 0.25, 0.45, 0.70, 0.90, 1.00]):
    cell = cs[f"T{5+i}"]
    cell.value, cell.font, cell.border = v, IN_F, BOX
    cell.number_format = PCT0
    cell.alignment = Alignment(horizontal="center", vertical="center")
for r in range(5 + 6, PS2 + 1):
    cell = cs[f"T{r}"]; cell.font, cell.border = IN_F, BOX
    cell.number_format = PCT0
cs.column_dimensions["T"].width = 12
cs["T4"].comment = Comment(
    "추진 단계가 그 단계에 도달했을 때의 진척률입니다. 부서 기준에 맞게 바꿔도 됩니다. "
    "[사업마스터]의 진척률이 이 값을 그대로 가져옵니다.", "업무일지")
cs[f"A{max(SG2, PS2, 36) + 2}"] = ("※ 추진 단계(S)와 단계 진척률(T)은 같은 행끼리 짝입니다.  "
                                   "※ 사업 구분(Q)·추진 단계(S)를 늘리면 [사업마스터]·[사업추진일지] 드롭다운에 바로 반영됩니다.")
cs[f"A{max(SG2, PS2, 36) + 2}"].font = NOTE_F

# 이름 정의 추가
for nm, ref in [("목록_사업구분", f"코드설정!$Q$5:$Q${SG2}"),
                ("목록_추진단계", f"코드설정!$S$5:$S${PS2}"),
                ("목록_사업번호", f"사업마스터!$A${PJ1}:$A${PJ2}")]:
    if nm in wb.defined_names: del wb.defined_names[nm]
    wb.defined_names.add(DefinedName(nm, attr_text=ref))

# ══════════════════════ 사업마스터
pm = fresh_sheet(wb, "사업마스터")
title(pm, "사업 마스터  ·  사업 등록 및 예산 관리",
      "※ 사업을 여기 한 번만 등록하면 [사업추진일지]에서 사업번호만 골라도 사업명이 자동으로 채워집니다.  "
      "※ 예산이 없는 사업은 [예산액]을 비워 두세요 — 집행률·잔액은 빈칸으로 남습니다.  "
      "※ 현재 단계·진척률·집행액은 [사업추진일지]에서 자동으로 올라옵니다.")
PM_H = ["사업번호","사업명","연도","사업 구분","주관팀","총괄 담당자",
        "예산액(원)","집행액(원)","집행률","잔액(원)",
        "착수예정일","완료예정일","현재 단계","진척률","추진 건수","진행상태","비고"]
hdr(pm, 4, PM_H)
widths(pm, {"A":16,"B":30,"C":8,"D":16,"E":13,"F":12,"G":15,"H":15,"I":10,"J":15,
            "K":13,"L":13,"M":13,"N":9,"O":10,"P":11,"Q":22,"R":3,"S":10})
PLC = R("사업추진일지", "C", PL1, PL2)      # 사업번호
PLI = R("사업추진일지", "I", PL1, PL2)      # 집행액
PLN = R("사업추진일지", "N", PL1, PL2)      # 단계순번(숨김)
FXC = {"H","I","J","M","N","O"}
for r in range(PJ1, PJ2 + 1):
    for j in range(1, 18):
        col = CL(j); c = pm.cell(row=r, column=j); c.border = BOX
        c.font = FX_F if col in FXC else IN_F
        c.alignment = Alignment(horizontal="left" if col in ("B","Q") else "center",
                                vertical="center", wrap_text=(col in ("B","Q")))
    g = f'IF($A{r}="","",'
    pm[f"C{r}"].number_format = "0"
    for L in ("G","J"): pm[f"{L}{r}"].number_format = MONEY
    for L in ("K","L"): pm[f"{L}{r}"].number_format = DATE_
    put(pm, f"H{r}", f'={g}SUMIFS({PLI},{PLC},$A{r}))', LNK_F, MONEY)
    put(pm, f"I{r}", f'=IF(OR($A{r}="",$G{r}="",$G{r}=0),"",$H{r}/$G{r})', FX_F, PCT)
    put(pm, f"J{r}", f'=IF(OR($A{r}="",$G{r}=""),"",$G{r}-$H{r})', FX_F, MONEY)
    pm[f"S{r}"] = f'=IF($A{r}="",0,SUMPRODUCT(MAX(({PLC}=$A{r})*{PLN})))'
    put(pm, f"M{r}", f'={g}IF($S{r}=0,"미착수",INDEX(코드설정!$S${PS1}:$S${PS2},$S{r})))', FX_F)
    put(pm, f"N{r}", f'={g}IF($S{r}=0,0,INDEX(코드설정!$T${PS1}:$T${PS2},$S{r})))', FX_F, PCT0)
    put(pm, f"O{r}", f'={g}COUNTIF({PLC},$A{r}))', LNK_F, NUM)
pm.freeze_panes = "C5"
pm.auto_filter.ref = f"A4:Q{PJ2}"
for c in ("R", "S"): pm.column_dimensions[c].hidden = True
pm["R4"] = "▼ 최고 추진단계 순번 (지우지 마세요)"; pm["R4"].font = NOTE_F
dv_name(pm, f"D{PJ1}:D{PJ2}", "목록_사업구분")
dv_name(pm, f"E{PJ1}:E{PJ2}", "목록_팀")
dv_name(pm, f"F{PJ1}:F{PJ2}", "목록_담당자")
dv_name(pm, f"P{PJ1}:P{PJ2}", "목록_진행상태")
pm["G4"].comment = Comment(
    "예산이 배정되지 않은 사업은 비워 두세요. 집행률과 잔액이 0%가 아니라 빈칸으로 남고, "
    "부서 집행률 집계에서도 제외됩니다.", "업무일지")
pm["M4"].comment = Comment(
    "[사업추진일지]에 기록된 추진 단계 중 가장 진행된 단계입니다. 단계를 적지 않은 사업은 '미착수'로 표시됩니다.", "업무일지")
pm["N4"].comment = Comment(
    "현재 단계에 해당하는 진척률입니다. 단계별 진척률은 [코드설정] T열에서 바꿀 수 있습니다.", "업무일지")
pm.conditional_formatting.add(f"A{PJ1}:Q{PJ2}", FormulaRule(
    formula=[f'$P{PJ1}="완료"'], fill=PatternFill("solid", fgColor="E2EFDA"), stopIfTrue=False))
pm.conditional_formatting.add(f"A{PJ1}:Q{PJ2}", FormulaRule(
    formula=[f'$P{PJ1}="보류"'], fill=PatternFill("solid", fgColor="FCE4D6"), stopIfTrue=False))
pm.conditional_formatting.add(f"A{PJ1}:Q{PJ2}", FormulaRule(
    formula=[f'AND($A{PJ1}<>"",COUNTIF($A${PJ1}:$A${PJ2},$A{PJ1})>1)'],
    fill=PatternFill("solid", fgColor="FFC7CE"), stopIfTrue=False))
pm.conditional_formatting.add(f"L{PJ1}:L{PJ2}", FormulaRule(
    formula=[f'AND($L{PJ1}<>"",$L{PJ1}<TODAY(),$P{PJ1}<>"완료")'],
    fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(color="9C0006", bold=True), stopIfTrue=False))

# ══════════════════════ 사업추진일지
pl = fresh_sheet(wb, "사업추진일지")
title(pl, "사업 추진일지  ·  사업별 추진 경과 기록",
      "※ 사업이 한 일을 날짜별로 한 줄씩 씁니다. 사업번호를 고르면 사업명이 자동으로 채워집니다.  "
      "※ 추진 단계는 선택 입력입니다 — 적으면 [사업마스터]의 현재 단계·진척률이 올라가고, 비워 두면 시간순 기록으로만 남습니다.  "
      "※ 집행액은 그 건에서 쓴 금액입니다. 예산 없는 사업이면 비워 두세요.")
PL_H = ["연번","일자","사업번호","사업명(자동)","추진 단계","추진 내용","조치 · 결과",
        "담당자","집행액(원)","증빙 · 문서번호","진행상태","비고"]
hdr(pl, 4, PL_H)
widths(pl, {"A":6,"B":12,"C":16,"D":26,"E":13,"F":40,"G":30,"H":11,"I":14,"J":18,"K":11,"L":22,
            "M":3,"N":8})
PLEFT = {"D","F","G","L"}
for r in range(PL1, PL2 + 1):
    for j in range(1, 13):
        col = CL(j); c = pl.cell(row=r, column=j); c.border = BOX
        c.font = FX_F if col in ("A","D") else IN_F
        c.alignment = Alignment(horizontal="left" if col in PLEFT else "center",
                                vertical="center", wrap_text=(col in PLEFT))
    pl[f"A{r}"] = f'=IF($B{r}="","",COUNT($B${PL1}:$B{r}))'
    pl[f"A{r}"].number_format = NUM
    pl[f"B{r}"].number_format = DATE_
    pl[f"D{r}"] = (f'=IF($C{r}="","",IFERROR(INDEX(사업마스터!$B${PJ1}:$B${PJ2},'
                   f'MATCH($C{r},사업마스터!$A${PJ1}:$A${PJ2},0)),"※ 미등록 사업번호"))')
    pl[f"I{r}"].number_format = MONEY
    pl[f"N{r}"] = f'=IFERROR(MATCH($E{r},코드설정!$S${PS1}:$S${PS2},0),0)'
pl.freeze_panes = "C5"
pl.auto_filter.ref = f"A4:L{PL2}"
for c in ("M", "N"): pl.column_dimensions[c].hidden = True
pl["M4"] = "▼ 추진단계 순번 (지우지 마세요)"; pl["M4"].font = NOTE_F
dv_name(pl, f"C{PL1}:C{PL2}", "목록_사업번호")
dv_name(pl, f"E{PL1}:E{PL2}", "목록_추진단계")
dv_name(pl, f"H{PL1}:H{PL2}", "목록_담당자")
dv_name(pl, f"K{PL1}:K{PL2}", "목록_진행상태")
pl["E4"].comment = Comment(
    "선택 입력입니다. 단계를 적으면 [사업마스터]의 현재 단계가 그 중 가장 진행된 단계로 올라갑니다. "
    "단계와 무관한 중간 기록은 비워 두세요.", "업무일지")
pl["I4"].comment = Comment(
    "그 건에서 집행한 금액입니다. 합계가 [사업마스터]의 집행액으로 올라갑니다. 예산 없는 사업은 비워 두세요.", "업무일지")
_pr = f"A{PL1}:L{PL2}"
pl.conditional_formatting.add(_pr, FormulaRule(
    formula=[f'AND($B{PL1}<>"",$K{PL1}="완료")'], fill=PatternFill("solid", fgColor="E2EFDA"), stopIfTrue=False))
pl.conditional_formatting.add(_pr, FormulaRule(
    formula=[f'AND($B{PL1}<>"",$K{PL1}="보류")'], fill=PatternFill("solid", fgColor="FCE4D6"), stopIfTrue=False))

wb.save(OUT)
print("Stage 1 완료 — 사업마스터 · 사업추진일지 추가")
