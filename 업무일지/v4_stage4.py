# -*- coding: utf-8 -*-
"""Stage 4 — 팀별 사건수 (출장/분석/전체) + 개인별 참여 사건수"""
import sys; sys.path.insert(0, "/home/user/live-trade/업무일지")
from v4_common import *

wb = openpyxl.load_workbook(OUT)
ov, pv = wb["부서 전체현황"], wb["개인별 현황"]
KM_A = R("사건마스터", "A", K1, K2)
NCASE  = K2 - K1 + 1          # 60
NTEAM  = TM2 - TM1 + 1        # 10
NPP    = P2 - P1 + 1          # 20

# ══════════════════════ 부서 전체현황: 팀별 사건수
R_TEAM_H, R_TEAM_1 = 43, 44
R_TEAM_T = R_TEAM_1 + NTEAM                       # 54 합계
HB, HE = 85, 85 + NCASE - 1                       # 기존 사건 블록 행
CASE_COL = "O"                                    # 기존 사건번호 열
T_TRIP, T_ANAL, T_ALL = 56, 67, 78                # BD / BO / BZ
YRO = f'{WB_},">="&DATE($B$3,1,1),{WB_},"<"&DATE($B$3+1,1,1)'

for j, h in enumerate(["출장 사건", "분석 사건", "전체 사건"], start=10):
    c = ov.cell(row=R_TEAM_H, column=j, value=h)
    c.font, c.fill, c.border = HDR_F, HDR_FILL, BOX
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
section(ov, 42, "3.  팀별 현황 (연간 누계)   ※ 주담당자의 소속팀 기준 / 회차 = 줄 수, 사건 = 고유 사건 수", span=12)

for t in range(NTEAM):
    trow = R_TEAM_1 + t
    tref = f'$A${trow}'
    for c_ in range(NCASE):
        hr = HB + c_
        case = f'${CASE_COL}{hr}'
        for base, job in ((T_TRIP, "출장"), (T_ANAL, "분석")):
            ov[f"{CL(base + t)}{hr}"] = (
                f'=IF({case}="",0,IF(COUNTIFS({WF_},{case},{YRO},{WAB},{tref},{WD_},"{job}")>0,1,0))')
        ov[f"{CL(T_ALL + t)}{hr}"] = (
            f'=IF({case}="",0,IF(COUNTIFS({WF_},{case},{YRO},{WAB},{tref})>0,1,0))')
    g = f'IF({tref}="","",'
    for col, base in (("J", T_TRIP), ("K", T_ANAL), ("L", T_ALL)):
        cc = CL(base + t)
        put(ov, f"{col}{trow}", f'={g}SUM({cc}${HB}:{cc}${HE}))', FX_F, NUM)
put(ov, f"J{R_TEAM_T}", '=$C$19', LNK_F, NUM, TOT_FILL)
put(ov, f"K{R_TEAM_T}", '=$E$19', LNK_F, NUM, TOT_FILL)
put(ov, f"L{R_TEAM_T}", '=$H$19', LNK_F, NUM, TOT_FILL)
ov[f"J{R_TEAM_H}"].comment = Comment(
    "그 팀이 주담당으로 맡은 사건 중 올해 출장 작업이 있었던 사건의 개수입니다. 회차가 몇 번이든 1로 셉니다.", "업무일지")
ov[f"L{R_TEAM_H}"].comment = Comment(
    "합계 칸은 팀별 값의 단순 합이 아니라 부서 전체 고유 사건 수입니다. 한 사건에 두 팀이 걸쳐 있으면 "
    "팀별 값을 더한 수가 부서 합계보다 커지기 때문입니다.", "업무일지")
widths(ov, {"J": 11, "K": 11, "L": 11})
ov[f"{CL(T_TRIP)}84"] = "▼ 팀별 고유 사건 수 계산용 (지우지 마세요) — 출장 / 분석 / 전체"
ov[f"{CL(T_TRIP)}84"].font = NOTE_F
for j in range(55, T_ALL + NTEAM + 1):
    ov.column_dimensions[CL(j)].hidden = True

# ══════════════════════ 개인별 현황: 참여 사건수
P_H, P_1 = 6, 7
P_T = P_1 + NPP                                   # 27 합계
PB, PE = 60, 60 + NCASE - 1                       # 숨은 블록 행
PCASE_COL, P_ALL = "P", 17                        # P=사건번호, Q~AJ=담당자별
YRP = f'{WB_},">="&DATE($B$3,1,1),{WB_},"<"&DATE($B$3+1,1,1)'

c = pv.cell(row=P_H, column=12, value="참여 사건수")
c.font, c.fill, c.border = HDR_F, HDR_FILL, BOX
c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
section(pv, 5, "1.  담당자별 종합 (연간 누계)   ※ 회차 = 줄 수, 사건 = 고유 사건 수", span=12)
for c_ in range(NCASE):
    pr = PB + c_
    pv[f"{PCASE_COL}{pr}"] = f'=IF(사건마스터!$A{K1 + c_}="","",사건마스터!$A{K1 + c_})'
for p in range(NPP):
    prow = P_1 + p
    nref = f'$A${prow}'
    cc = CL(P_ALL + p)
    for c_ in range(NCASE):
        pr = PB + c_
        case = f'${PCASE_COL}{pr}'
        pv[f"{cc}{pr}"] = (
            f'=IF({case}="",0,IF(COUNTIFS({WF_},{case},{YRP},{WAA},"*"&{nref}&"*")>0,1,0))')
    put(pv, f"L{prow}", f'=IF({nref}="","",SUM({cc}${PB}:{cc}${PE}))', FX_F, NUM)
put(pv, f"L{P_T}", "='부서 전체현황'!$H$19", LNK_F, NUM, TOT_FILL)
pv.cell(row=P_H, column=12).comment = Comment(
    "본인이 참여한 사건의 개수입니다. 같은 사건을 여러 회차로 나눠 했어도 1로 셉니다. "
    "합계 칸은 개인별 값의 단순 합이 아니라 부서 전체 고유 사건 수입니다 — 한 사건에 여러 명이 "
    "참여하므로 개인별 값을 더하면 중복됩니다.", "업무일지")
widths(pv, {"L": 12})
pv[f"{PCASE_COL}59"] = "▼ 개인별 고유 사건 수 계산용 (지우지 마세요)"
pv[f"{PCASE_COL}59"].font = NOTE_F
for j in range(15, P_ALL + NPP + 1):
    pv.column_dimensions[CL(j)].hidden = True

wb.save(OUT)
print(f"Stage 4 완료 — 팀별 사건수 3종({NTEAM}팀) · 개인별 참여 사건수({NPP}명)")
print(f"  신규 헬퍼 셀: 팀 {NTEAM*NCASE*3:,} + 개인 {NPP*NCASE:,} = {NTEAM*NCASE*3 + NPP*NCASE:,}")
