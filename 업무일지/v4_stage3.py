# -*- coding: utf-8 -*-
"""Stage 3 — 교육 이수 현황"""
import sys; sys.path.insert(0, "/home/user/live-trade/업무일지")
from v4_common import *

wb = openpyxl.load_workbook(OUT)
ed = fresh_sheet(wb, "교육 이수 현황")
ed.sheet_view.showGridLines = False
title(ed, "교육 이수 현황  ·  담당자별 연간 교육",
      "※ 입력할 것이 없습니다. [업무일지]에서 직무구분이 '교육'인 기록과 [일일실시사항]의 '교육·훈련'을 모아 셉니다.  "
      "※ 노란 칸(연간 이수 기준)만 부서 방침에 맞게 바꾸면 이수 판정이 다시 계산됩니다.")
widths(ed, {"A":13,"B":13,"C":14,"D":13,"E":13,"F":13,"G":10,"H":11,"I":11,"J":13,"K":11,"L":3,
            "M":3,"N":3})
put(ed, "A3", "기준연도", LBL_F, fill=SUB_FILL)
put(ed, "B3", "='부서 전체현황'!$B$3", LNK_F, "0", KEY_FILL)
put(ed, "C3", "연간 이수 기준(회)", LBL_F, fill=SUB_FILL)
put(ed, "D3", 2, IN_F, "0", KEY_FILL)
ed["E3"] = "← 이 횟수 이상이면 '이수'로 판정합니다"; ed["E3"].font = NOTE_F
ed["D3"].comment = Comment(
    "담당자가 한 해에 이 횟수 이상 교육에 참여하면 '이수'로 봅니다. 부서 방침에 맞게 바꾸세요. "
    "[업무일지]의 교육 회차와 [일일실시사항]의 교육·훈련 건수를 합쳐서 셉니다.", "업무일지")

YR  = f'{WB_},">="&DATE($B$3,1,1),{WB_},"<"&DATE($B$3+1,1,1)'
DYR = f'{DB_},">="&DATE($B$3,1,1),{DB_},"<"&DATE($B$3+1,1,1)'
EDU = [("외부 교육·세미나", "C"), ("내부 직무교육", "D"), ("자격·인증 과정", "E"),
       ("워크숍·학회", "F"), ("기타 교육", "G")]

# ── 1. 담당자별 교육 이수
section(ed, 5, "1.  담당자별 교육 이수 (연간 누계)", span=11)
hdr(ed, 6, ["담당자","소속팀","외부 교육·세미나","내부 직무교육","자격·인증 과정",
            "워크숍·학회","기타 교육","업무일지 계","일일 교육·훈련","총 교육 횟수","이수 판정"])
NPP = P2 - P1 + 1
for i in range(NPP):
    r, cr = 7 + i, P1 + i
    nm = f'$A{r}'
    g = f'IF({nm}="","",'
    put(ed, f"A{r}", f'=IF(코드설정!$C{cr}="","",코드설정!$C{cr})', FX_F, fill=SUB_FILL)
    put(ed, f"B{r}", f'={g}코드설정!$D{cr})', LNK_F)
    for lab, col in EDU:
        put(ed, f"{col}{r}",
            f'={g}COUNTIFS({YR},{WAA},"*"&{nm}&"*",{WD_},"교육",{WE_},"{lab}"))', LNK_F, NUM)
    put(ed, f"H{r}", f'={g}COUNTIFS({YR},{WAA},"*"&{nm}&"*",{WD_},"교육"))', LNK_F, NUM)
    put(ed, f"I{r}", f'={g}COUNTIFS({DYR},{DD_},{nm},{DF_},"교육·훈련"))', LNK_F, NUM)
    put(ed, f"J{r}", f'={g}H{r}+I{r})', FX_F, NUM)
    put(ed, f"K{r}", f'={g}IF($J{r}>=$D$3,"이수","미이수 ("&MAX(0,$D$3-$J{r})&"회 부족)"))', FX_F)
R_ETOT = 7 + NPP
put(ed, f"A{R_ETOT}", "합계", LBL_F, fill=TOT_FILL); put(ed, f"B{R_ETOT}", "", LBL_F, fill=TOT_FILL)
for L in "CDEFGHIJ":
    put(ed, f"{L}{R_ETOT}", f'=SUM({L}7:{L}{R_ETOT-1})', FX_F, NUM, TOT_FILL)
put(ed, f"K{R_ETOT}", f'=COUNTIF(K7:K{R_ETOT-1},"이수")&" / "&(({NPP})-COUNTBLANK($A$7:$A${R_ETOT-1}))&" 명 이수"',
    FX_F, fill=TOT_FILL)
ed.conditional_formatting.add(f"A7:K{R_ETOT-1}", FormulaRule(
    formula=['AND($A7<>"",LEFT($K7,3)="미이수")'],
    fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(color="9C0006", bold=True), stopIfTrue=True))
ed.conditional_formatting.add(f"A7:K{R_ETOT-1}", FormulaRule(
    formula=['$K7="이수"'], fill=PatternFill("solid", fgColor="E2EFDA"), stopIfTrue=True))
ed["H6"].comment = Comment(
    "[업무일지]에서 직무구분이 '교육'이고 참여자 명단에 이름이 있는 회차를 셉니다. 주담당·참여를 구분하지 않습니다.", "업무일지")
ed["I6"].comment = Comment(
    "[일일실시사항]에서 업무구분이 '교육·훈련'이고 담당자가 본인인 건을 셉니다.", "업무일지")

# ── 2. 교육 세부유형별
R_TY = R_ETOT + 2
section(ed, R_TY, "2.  교육 세부유형별 (연간 누계)", span=5)
hdr(ed, R_TY + 1, ["세부유형","회차","참여 연인원(명)","비중","소요일수"])
R_TYT = R_TY + 2 + len(EDU)
for i, (lab, _) in enumerate(EDU):
    r = R_TY + 2 + i
    put(ed, f"A{r}", lab, FX_F, fill=SUB_FILL, align="left")
    put(ed, f"B{r}", f'=COUNTIFS({YR},{WD_},"교육",{WE_},$A{r})', LNK_F, NUM)
    put(ed, f"C{r}", f'=SUMIFS({WZ_},{YR},{WD_},"교육",{WE_},$A{r})', LNK_F, NUM)
    put(ed, f"D{r}", f'=IFERROR(B{r}/$B${R_TYT},0)', FX_F, PCT)
    put(ed, f"E{r}", f'=SUMIFS({WAF},{YR},{WD_},"교육",{WE_},$A{r})', LNK_F, NUM)
put(ed, f"A{R_TYT}", "합계", LBL_F, fill=TOT_FILL)
for L in "BCE":
    put(ed, f"{L}{R_TYT}", f'=SUM({L}{R_TY+2}:{L}{R_TYT-1})', FX_F, NUM, TOT_FILL)
put(ed, f"D{R_TYT}", f'=IFERROR(B{R_TYT}/$B${R_TYT},0)', FX_F, PCT, TOT_FILL)

# ── 3. 팀별 교육
R_TM = R_TYT + 2
section(ed, R_TM, "3.  팀별 교육 (연간 누계)", span=6)
hdr(ed, R_TM + 1, ["소속팀","인원","교육 회차","참여 연인원(명)","이수 인원","미이수 인원"])
NTEAM = TM2 - TM1 + 1
R_TMT = R_TM + 2 + NTEAM
for i in range(NTEAM):
    r, cr = R_TM + 2 + i, TM1 + i
    tm = f'$A{r}'
    g = f'IF({tm}="","",'
    put(ed, f"A{r}", f'=IF(코드설정!$I{cr}="","",코드설정!$I{cr})', FX_F, fill=SUB_FILL)
    put(ed, f"B{r}", f'={g}COUNTIF(코드설정!$D${P1}:$D${P2},{tm}))', LNK_F, NUM)
    put(ed, f"C{r}", f'={g}SUMIFS($J$7:$J${R_ETOT-1},$B$7:$B${R_ETOT-1},{tm}))', FX_F, NUM)
    put(ed, f"D{r}", f'={g}SUMIFS({WZ_},{YR},{WD_},"교육",{WAB},{tm}))', LNK_F, NUM)
    put(ed, f"E{r}", f'={g}COUNTIFS($B$7:$B${R_ETOT-1},{tm},$K$7:$K${R_ETOT-1},"이수"))', FX_F, NUM)
    put(ed, f"F{r}", f'={g}COUNTIFS($B$7:$B${R_ETOT-1},{tm},$K$7:$K${R_ETOT-1},"미이수*"))', FX_F, NUM)
put(ed, f"A{R_TMT}", "합계", LBL_F, fill=TOT_FILL)
for L in "BCDEF":
    put(ed, f"{L}{R_TMT}", f'=SUM({L}{R_TM+2}:{L}{R_TMT-1})', FX_F, NUM, TOT_FILL)
ed[f"A{R_TMT+2}"] = "※ 교육 회차는 참여자 명단에 이름이 들어간 모든 회차를 셉니다(주담당·참여 구분 없음)."
ed[f"A{R_TMT+2}"].font = NOTE_F
ed[f"A{R_TMT+3}"] = "※ 참여 연인원은 각 교육 건의 투입인원을 합산한 값이라, 같은 사람이 여러 번 참여하면 중복 계상됩니다."
ed[f"A{R_TMT+3}"].font = NOTE_F
ed[f"A{R_TMT+4}"] = "※ 팀별 교육 회차는 소속 담당자들의 개인 회차를 더한 값이고, 참여 연인원은 주담당자의 소속팀 기준입니다 — 두 값은 다를 수 있습니다."
ed[f"A{R_TMT+4}"].font = NOTE_F
ed.freeze_panes = "A7"

wb.save(OUT)
print(f"Stage 3 완료 — 교육 이수 현황 (담당자 {NPP} / 세부유형 {len(EDU)} / 팀 {NTEAM})")
