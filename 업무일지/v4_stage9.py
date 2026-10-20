# -*- coding: utf-8 -*-
"""Stage 9 — 개인 이력 조회 (담당자 한 명의 모든 업무)"""
import sys; sys.path.insert(0, "/home/user/live-trade/업무일지")
from v4_common import *

wb = openpyxl.load_workbook(OUT)
NW, ND, NP = 80, 30, 20              # 표시 건수: 업무 / 일일 / 사업
NWR = W2 - W1 + 1                    # 1000
NDR = D2 - D1 + 1                    # 1000
NPR = PL2 - PL1 + 1                  # 400

ph = fresh_sheet(wb, "개인 이력 조회")
ph.sheet_view.showGridLines = False
title(ph, "개인 이력 조회  ·  담당자별 전체 업무",
      "※ 노란 칸에서 담당자와 연도를 고르면 그 사람이 한 해 동안 한 모든 일이 나옵니다.  "
      "※ [업무일지]의 출장·분석·교육, [일일실시사항], [사업추진일지] 세 곳을 모아 날짜순으로 보여줍니다.  "
      "※ 입력할 것은 담당자와 연도뿐이며 나머지는 자동입니다.")
widths(ph, {"A":7,"B":12,"C":12,"D":7,"E":10,"F":18,"G":16,"H":6,"I":30,
            "J":9,"K":9,"L":18,"M":11,"N":11,"O":3})
put(ph, "A4", "담당자", LBL_F, fill=SUB_FILL)
put(ph, "B4", None, IN_F, fill=KEY_FILL)
dv_name(ph, "B4", "목록_담당자")
put(ph, "C4", "기준연도", LBL_F, fill=SUB_FILL)
put(ph, "D4", "='부서 전체현황'!$B$3", LNK_F, "0", KEY_FILL)
put(ph, "E4", "소속팀", LBL_F, fill=SUB_FILL)
put(ph, "F4", f'=IF($B$4="","",IFERROR(INDEX(코드설정!$D${P1}:$D${P2},'
              f'MATCH($B$4,코드설정!$C${P1}:$C${P2},0)),""))', LNK_F)
ph["G4"] = "← 담당자를 고르면 아래 전부 다시 계산됩니다"; ph["G4"].font = NOTE_F

NM   = '$B$4'
Y1   = 'DATE($D$4,1,1)'
Y2   = 'DATE($D$4+1,1,1)'
WYR  = f'{WB_},">="&{Y1},{WB_},"<"&{Y2}'
DYR  = f'{DB_},">="&{Y1},{DB_},"<"&{Y2}'
PLB  = R("사업추진일지", "B", PL1, PL2)
PLC  = R("사업추진일지", "C", PL1, PL2)
PLH  = R("사업추진일지", "H", PL1, PL2)
PLI  = R("사업추진일지", "I", PL1, PL2)
PYR  = f'{PLB},">="&{Y1},{PLB},"<"&{Y2}'
LIKE = f'"*"&{NM}&"*"'

# ── 숨은 정렬 키
for i in range(1, NWR + 1):
    s = W1 + i - 1
    ph[f"AP{i}"] = (f'=IF({NM}="","",IF(AND(COUNTIF(업무일지!${WC["명단"]}{s},{LIKE})>0,'
                    f'업무일지!${WC["시작일"]}{s}>={Y1},업무일지!${WC["시작일"]}{s}<{Y2}),'
                    f'업무일지!${WC["시작일"]}{s}*10000+{i},""))')
for i in range(1, NDR + 1):
    s = D1 + i - 1
    ph[f"AS{i}"] = (f'=IF({NM}="","",IF(AND(일일실시사항!${DC["담당자"]}{s}={NM},'
                    f'일일실시사항!${DC["일자"]}{s}>={Y1},일일실시사항!${DC["일자"]}{s}<{Y2}),'
                    f'일일실시사항!${DC["일자"]}{s}*10000+{i},""))')
for i in range(1, NPR + 1):
    s = PL1 + i - 1
    ph[f"AV{i}"] = (f'=IF({NM}="","",IF(AND(사업추진일지!$H{s}={NM},'
                    f'사업추진일지!$B{s}>={Y1},사업추진일지!$B{s}<{Y2}),'
                    f'사업추진일지!$B{s}*10000+{i},""))')
# 참여 사건 플래그
for i in range(K2 - K1 + 1):
    ph[f"AY{i+1}"] = (f'=IF(OR({NM}="",사건마스터!$A{K1+i}=""),0,'
                      f'IF(COUNTIFS({WF_},사건마스터!$A{K1+i},{WYR},{WAA},{LIKE})>0,1,0))')
# 최근 활동일
ph["BA2"] = f'=IF({NM}="",0,SUMPRODUCT(MAX(({DD_}={NM})*{DB_})))'
ph["BA3"] = f'=IF({NM}="",0,SUMPRODUCT(MAX(({PLH}={NM})*{PLB})))'
ph["BA4"] = (f'=IFERROR(IF(MAX($AP$1:$AP${NWR})=0,0,'
             f'INT(MAX($AP$1:$AP${NWR})/10000)),0)')   # 키 = 날짜*10000+순번 → 날짜만 추출

# ══════════ 1. 연간 요약
section(ph, 6, "1.  연간 요약", span=14)
hdr(ph, 7, ["출장 회차","분석 회차","교육 회차","업무 계","주담당","참여",
            "참여 사건수","출장 소요일수","대상매체","일일실시","사업 추진","집행액(원)",
            "진행중","최근 활동일"])
g = f'IF({NM}="","",'
put(ph, "A8", f'={g}COUNTIFS({WYR},{WAA},{LIKE},{WD_},"출장"))', LNK_F, NUM)
put(ph, "B8", f'={g}COUNTIFS({WYR},{WAA},{LIKE},{WD_},"분석"))', LNK_F, NUM)
put(ph, "C8", f'={g}COUNTIFS({WYR},{WAA},{LIKE},{WD_},"교육"))', LNK_F, NUM)
put(ph, "D8", f'={g}A8+B8+C8)', FX_F, NUM)
put(ph, "E8", f'={g}COUNTIFS({WYR},{WJ_},{NM}))', LNK_F, NUM)
put(ph, "F8", f'={g}D8-E8)', FX_F, NUM)
put(ph, "G8", f'={g}SUM($AY$1:$AY${K2-K1+1}))', FX_F, NUM)
put(ph, "H8", f'={g}SUMIFS({WAF},{WYR},{WAA},{LIKE},{WD_},"출장"))', LNK_F, NUM)
put(ph, "I8", f'={g}SUMIFS({WAD},{WYR},{WAA},{LIKE}))', LNK_F, NUM)
put(ph, "J8", f'={g}COUNTIFS({DYR},{DD_},{NM}))', LNK_F, NUM)
put(ph, "K8", f'={g}COUNTIFS({PYR},{PLH},{NM}))', LNK_F, NUM)
put(ph, "L8", f'={g}SUMIFS({PLI},{PYR},{PLH},{NM}))', LNK_F, MONEY)
put(ph, "M8", f'={g}COUNTIFS({WYR},{WAA},{LIKE},{WAG},"진행중"))', LNK_F, NUM)
put(ph, "N8", f'={g}IF(MAX($BA$2:$BA$4)=0,"",MAX($BA$2:$BA$4)))', FX_F, DATE_)
ph["E7"].comment = Comment("본인이 주담당자로 지정된 회차입니다. [참여]는 전체 회차에서 주담당을 뺀 값입니다.", "업무일지")
ph["G7"].comment = Comment("본인이 참여한 사건의 개수입니다. 같은 사건을 여러 회차로 나눠 했어도 1로 셉니다.", "업무일지")
ph["N7"].comment = Comment("업무일지·일일실시사항·사업추진일지 가운데 가장 늦은 기록 날짜입니다.", "업무일지")

# ══════════ 2. 업무 이력
R_W = 10
section(ph, R_W, f"2.  업무 이력 — 출장 · 분석 · 교육 (날짜순 · 상위 {NW}건)", span=14)
hdr(ph, R_W + 1, ["순번","시작일","종료일","일수","직무구분","세부유형","사건번호","회차",
                  "사건명 · 업무명","역할","투입인원","장소 / 분석실","대상매체","진행상태"])
def _wi(col, ref): return f'INDEX({R("업무일지", col, W1, W2)},${ref})'
for k in range(1, NW + 1):
    r = R_W + 1 + k
    ph[f"AQ{r}"] = f'=IFERROR(MATCH(SMALL($AP$1:$AP${NWR},{k}),$AP$1:$AP${NWR},0),"")'
    gg = f'IF($AQ{r}="","",'
    put(ph, f"A{r}", f'={gg}{k})', FX_F, NUM)
    put(ph, f"B{r}", f'={gg}{_wi(WC["시작일"], f"AQ{r}")})', LNK_F, DATE_)
    put(ph, f"C{r}", f'={gg}{_wi(WC["종료일"], f"AQ{r}")})', LNK_F, DATE_)
    put(ph, f"D{r}", f'={gg}{_wi(WC["소요일수"], f"AQ{r}")})', LNK_F, NUM)
    put(ph, f"E{r}", f'={gg}{_wi(WC["직무구분"], f"AQ{r}")})', LNK_F)
    put(ph, f"F{r}", f'={gg}{_wi(WC["세부유형"], f"AQ{r}")})', LNK_F, align="left")
    put(ph, f"G{r}", f'={gg}{_wi(WC["사건번호"], f"AQ{r}")})', LNK_F)
    put(ph, f"H{r}", f'={gg}{_wi(WC["회차"], f"AQ{r}")})', LNK_F, NUM)
    put(ph, f"I{r}", f'={gg}{_wi(WC["사건명"], f"AQ{r}")})', LNK_F, align="left")
    put(ph, f"J{r}", f'={gg}IF({_wi(WC["주담당"], f"AQ{r}")}={NM},"주담당","참여"))', FX_F)
    put(ph, f"K{r}", f'={gg}{_wi(WC["투입인원"], f"AQ{r}")})', LNK_F, NUM)
    put(ph, f"L{r}", f'={gg}{_wi(WC["장소"], f"AQ{r}")})', LNK_F, align="left")
    put(ph, f"M{r}", f'={gg}{_wi(WC["매체"], f"AQ{r}")})', LNK_F, NUM)
    put(ph, f"N{r}", f'={gg}{_wi(WC["진행상태"], f"AQ{r}")})', LNK_F)
_W1r, _W2r = R_W + 2, R_W + 1 + NW
for cond, color in (('"완료"', "E2EFDA"), ('"보류"', "FCE4D6")):
    ph.conditional_formatting.add(f"A{_W1r}:N{_W2r}", FormulaRule(
        formula=[f'$N{_W1r}={cond}'], fill=PatternFill("solid", fgColor=color), stopIfTrue=False))
ph.conditional_formatting.add(f"J{_W1r}:J{_W2r}", FormulaRule(
    formula=[f'$J{_W1r}="주담당"'], font=Font(bold=True, color="1F3864"), stopIfTrue=False))

# ══════════ 3. 일일 실시사항
R_D = _W2r + 2
section(ph, R_D, f"3.  일일 실시사항 (날짜순 · 상위 {ND}건)", span=8)
hdr(ph, R_D + 1, ["순번","일자","요일","업무구분","관련 사건번호","실시 내용","조치 · 결과","진행상태"])
def _di(col, ref): return f'INDEX({R("일일실시사항", col, D1, D2)},${ref})'
for k in range(1, ND + 1):
    r = R_D + 1 + k
    ph[f"AT{r}"] = f'=IFERROR(MATCH(SMALL($AS$1:$AS${NDR},{k}),$AS$1:$AS${NDR},0),"")'
    gg = f'IF($AT{r}="","",'
    put(ph, f"A{r}", f'={gg}{k})', FX_F, NUM)
    put(ph, f"B{r}", f'={gg}{_di(DC["일자"], f"AT{r}")})', LNK_F, DATE_)
    put(ph, f"C{r}", f'={gg}{_di(DC["요일"], f"AT{r}")})', LNK_F)
    put(ph, f"D{r}", f'={gg}{_di(DC["업무구분"], f"AT{r}")})', LNK_F, align="left")
    put(ph, f"E{r}", f'={gg}{_di(DC["사건번호"], f"AT{r}")})', LNK_F)
    put(ph, f"F{r}", f'={gg}{_di(DC["내용"], f"AT{r}")})', LNK_F, align="left")
    ph.merge_cells(f"F{r}:I{r}")
    put(ph, f"J{r}", f'={gg}{_di(DC["결과"], f"AT{r}")})', LNK_F, align="left")
    ph.merge_cells(f"J{r}:M{r}")
    put(ph, f"N{r}", f'={gg}{_di(DC["진행상태"], f"AT{r}")})', LNK_F)
for _c in ("G", "H"):                      # 일일 구역도 동일
    ph[f"{_c}{R_D+1}"].value = None
for j, h in zip((6, 10), ("실시 내용", "조치 · 결과")):
    c = ph.cell(row=R_D + 1, column=j, value=h)
    c.font, c.fill, c.border = HDR_F, HDR_FILL, BOX
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
ph.cell(row=R_D + 1, column=14, value="진행상태").font = HDR_F
ph.cell(row=R_D + 1, column=14).fill = HDR_FILL
ph.cell(row=R_D + 1, column=14).border = BOX
ph.cell(row=R_D + 1, column=14).alignment = Alignment(horizontal="center", vertical="center")
ph.merge_cells(f"F{R_D+1}:I{R_D+1}"); ph.merge_cells(f"J{R_D+1}:M{R_D+1}")

# ══════════ 4. 사업 추진 참여
R_P = R_D + 1 + ND + 2
section(ph, R_P, f"4.  사업 추진 참여 (날짜순 · 상위 {NP}건)", span=8)
hdr(ph, R_P + 1, ["순번","일자","사업번호","사업명","추진 단계","추진 내용","집행액(원)","진행상태"])
def _pi(col, ref): return f'INDEX({R("사업추진일지", col, PL1, PL2)},${ref})'
for k in range(1, NP + 1):
    r = R_P + 1 + k
    ph[f"AW{r}"] = f'=IFERROR(MATCH(SMALL($AV$1:$AV${NPR},{k}),$AV$1:$AV${NPR},0),"")'
    gg = f'IF($AW{r}="","",'
    put(ph, f"A{r}", f'={gg}{k})', FX_F, NUM)
    put(ph, f"B{r}", f'={gg}{_pi("B", f"AW{r}")})', LNK_F, DATE_)
    put(ph, f"C{r}", f'={gg}{_pi("C", f"AW{r}")})', LNK_F)
    put(ph, f"D{r}", f'={gg}{_pi("D", f"AW{r}")})', LNK_F, align="left")
    ph.merge_cells(f"D{r}:E{r}")
    put(ph, f"F{r}", f'={gg}{_pi("E", f"AW{r}")})', LNK_F)
    put(ph, f"G{r}", f'={gg}{_pi("F", f"AW{r}")})', LNK_F, align="left")
    ph.merge_cells(f"G{r}:K{r}")
    put(ph, f"L{r}", f'={gg}{_pi("I", f"AW{r}")})', LNK_F, MONEY)
    put(ph, f"M{r}", f'={gg}{_pi("K", f"AW{r}")})', LNK_F)
    put(ph, f"N{r}", "", LNK_F)
for _c in ("E", "H"):                      # hdr()가 써 둔 텍스트가 병합에 가려지므로 제거
    ph[f"{_c}{R_P+1}"].value = None
for col, lab, span in (("D","사업명","E"),("F","추진 단계",None),("G","추진 내용","K"),
                       ("L","집행액(원)",None),("M","진행상태",None)):
    c = ph[f"{col}{R_P+1}"]; c.value = lab
    c.font, c.fill, c.border = HDR_F, HDR_FILL, BOX
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
ph.merge_cells(f"D{R_P+1}:E{R_P+1}"); ph.merge_cells(f"G{R_P+1}:K{R_P+1}")

# ══════════ 5. 월별 활동
R_M = R_P + 1 + NP + 2
section(ph, R_M, "5.  월별 활동", span=8)
hdr(ph, R_M + 1, ["월","출장","분석","교육","업무 계","일일실시","사업 추진","합계"])
def MOW(m): return f'{WB_},">="&DATE($D$4,{m},1),{WB_},"<"&DATE($D$4,{m}+1,1)'
def MOD_(m): return f'{DB_},">="&DATE($D$4,{m},1),{DB_},"<"&DATE($D$4,{m}+1,1)'
def MOP(m): return f'{PLB},">="&DATE($D$4,{m},1),{PLB},"<"&DATE($D$4,{m}+1,1)'
for m in range(1, 13):
    r = R_M + 1 + m
    put(ph, f"A{r}", f"{m}월", LBL_F, fill=SUB_FILL)
    put(ph, f"B{r}", f'={g}COUNTIFS({MOW(m)},{WAA},{LIKE},{WD_},"출장"))', LNK_F, NUM)
    put(ph, f"C{r}", f'={g}COUNTIFS({MOW(m)},{WAA},{LIKE},{WD_},"분석"))', LNK_F, NUM)
    put(ph, f"D{r}", f'={g}COUNTIFS({MOW(m)},{WAA},{LIKE},{WD_},"교육"))', LNK_F, NUM)
    put(ph, f"E{r}", f'={g}B{r}+C{r}+D{r})', FX_F, NUM)
    put(ph, f"F{r}", f'={g}COUNTIFS({MOD_(m)},{DD_},{NM}))', LNK_F, NUM)
    put(ph, f"G{r}", f'={g}COUNTIFS({MOP(m)},{PLH},{NM}))', LNK_F, NUM)
    put(ph, f"H{r}", f'={g}E{r}+F{r}+G{r})', FX_F, NUM)
R_MT = R_M + 14
put(ph, f"A{R_MT}", "합계", LBL_F, fill=TOT_FILL)
for L in "BCDEFGH":
    put(ph, f"{L}{R_MT}", f'=SUM({L}{R_M+2}:{L}{R_M+13})', FX_F, NUM, TOT_FILL)
ph[f"A{R_MT+2}"] = (f"※ 업무 이력은 상위 {NW}건, 일일 실시사항 {ND}건, 사업 추진 {NP}건까지 표시합니다. "
                    f"그보다 많으면 각 원본 시트에서 담당자로 필터를 거세요.")
ph[f"A{R_MT+2}"].font = NOTE_F
ph[f"A{R_MT+3}"] = "※ 업무 이력은 참여자 명단에 이름이 들어간 모든 회차를 모읍니다(주담당·참여 구분 없이). 역할 칸에서 구분을 확인하세요."
ph[f"A{R_MT+3}"].font = NOTE_F
ph[f"A{R_MT+4}"] = "※ 동명이인은 이름 부분일치로 섞일 수 있습니다. [코드설정]에서 '김디지(1팀)'처럼 구분해 등록하세요."
ph[f"A{R_MT+4}"].font = NOTE_F
for j in range(15, 56):
    ph.column_dimensions[CL(j)].hidden = True

ph.page_setup.orientation = "landscape"
ph.page_setup.paperSize = ph.PAPERSIZE_A4
ph.page_setup.fitToWidth = 1; ph.page_setup.fitToHeight = 0
ph.sheet_properties.pageSetUpPr.fitToPage = True
ph.page_margins.left = ph.page_margins.right = 0.35
ph.print_options.horizontalCentered = True
ph.print_area = f"A1:N{R_MT+4}"
ph.oddHeader.left.text = '&"Arial"&9개인 이력 조회'
ph.oddHeader.right.text = '&"Arial"&9인쇄 &D'
ph.oddFooter.center.text = '&"Arial"&9&P / &N'
ph.sheet_properties.tabColor = "BF8F00"

# 개인별 현황에 안내
pv = wb["개인별 현황"]
pv["D3"] = "← 담당자 한 명의 전체 이력은 [개인 이력 조회] 시트에서"
pv["D3"].font = Font(name=F, size=9, bold=True, color="BF8F00")

ORDER = ["대시보드","사용안내","월력표","업무일지","일일실시사항","사건마스터","지원 대장",
         "사업마스터","사업추진일지","장비·매체 관리","진행중 보드",
         "사건 이력 조회","사업별 현황","부서 전체현황","사건별 현황","개인별 현황",
         "개인 이력 조회","교육 이수 현황","입력검증","코드설정"]
assert sorted(ORDER) == sorted(wb.sheetnames), set(ORDER) ^ set(wb.sheetnames)
wb._sheets = [wb[n] for n in ORDER]
wb.active = 0
wb.save(OUT)
print(f"Stage 9 완료 — 개인 이력 조회 (업무 {NW} / 일일 {ND} / 사업 {NP} / 월별 12)")
print(f"  섹션: 요약 6 · 업무 {R_W} · 일일 {R_D} · 사업 {R_P} · 월별 {R_M} · 합계 {R_MT}")
