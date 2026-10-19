# -*- coding: utf-8 -*-
"""Stage 5 — 대시보드 (맨 앞)"""
import sys; sys.path.insert(0, "/home/user/live-trade/업무일지")
from v4_common import *

wb = openpyxl.load_workbook(OUT)
db = wb.create_sheet("대시보드")
db.sheet_view.showGridLines = False
title(db, "대시보드  ·  부서 업무 현황",
      "※ 노란 칸(연도·월)만 바꾸면 아래 전부 다시 계산됩니다.  ※ 파일을 열 때마다 오늘 날짜 기준으로 갱신됩니다.  "
      "※ 회차 = 업무일지 줄 수 / 사건 = 고유 사건 수.")
widths(db, {"A":15,"B":11,"C":11,"D":11,"E":11,"F":13,"G":15,"H":26,"I":11,"J":11,"K":13,"L":13,
            "M":13,"N":3})

put(db, "A4", "조회 연도", LBL_F, fill=SUB_FILL)
put(db, "B4", 2026, IN_F, "0", KEY_FILL)
put(db, "C4", "조회 월", LBL_F, fill=SUB_FILL)
put(db, "D4", 3, IN_F, "0", KEY_FILL)
put(db, "E4", '=TEXT(DATE($B$4,$D$4,1),"yyyy-mm-dd")&" ~ "&TEXT(DATE($B$4,$D$4+1,1)-1,"yyyy-mm-dd")',
    FX_F, align="left"); db.merge_cells("E4:G4")
put(db, "H4", "오늘", LBL_F, fill=SUB_FILL)
put(db, "I4", "=TODAY()", FX_F, DATE_)
from openpyxl.worksheet.datavalidation import DataValidation as DV
_m = DV(type="whole", operator="between", formula1=1, formula2=12,
        showErrorMessage=True, errorTitle="월 입력", error="1~12 사이의 숫자를 입력하세요.")
db.add_data_validation(_m); _m.add("D4")

ST_, EN_ = 'DATE($B$4,$D$4,1)', 'DATE($B$4,$D$4+1,1)'
PST, PEN = 'DATE($B$4,$D$4-1,1)', 'DATE($B$4,$D$4,1)'
def MM(a, b):  return f'{WB_},">="&{a},{WB_},"<"&{b}'
def DM(a, b):  return f'{DB_},">="&{a},{DB_},"<"&{b}'
YRD = f'{WB_},">="&DATE($B$4,1,1),{WB_},"<"&DATE($B$4+1,1,1)'
OVM = "'부서 전체현황'"
PLB = R("사업추진일지", "B", PL1, PL2)
PLI = R("사업추진일지", "I", PL1, PL2)

# ══════════ 1. 이번 달 한눈에
section(db, 6, "1.  이번 달 한눈에", span=13)
TILES = [("출장 회차", f'=COUNTIFS({MM(ST_,EN_)},{WD_},"출장")', f'=COUNTIFS({MM(PST,PEN)},{WD_},"출장")'),
         ("분석 회차", f'=COUNTIFS({MM(ST_,EN_)},{WD_},"분석")', f'=COUNTIFS({MM(PST,PEN)},{WD_},"분석")'),
         ("교육 회차", f'=COUNTIFS({MM(ST_,EN_)},{WD_},"교육")', f'=COUNTIFS({MM(PST,PEN)},{WD_},"교육")'),
         ("전체 회차", '=B8+C8+D8', '=B10+C10+D10'),
         ("진행중", f'=COUNTIF({WAG},"진행중")', ''),
         ("일일실시", f'=COUNTIFS({DM(ST_,EN_)})', f'=COUNTIFS({DM(PST,PEN)})'),
         ("대상매체(건)", f'=SUMIFS({WAD},{MM(ST_,EN_)})', f'=SUMIFS({WAD},{MM(PST,PEN)})')]
for i, (lab, cur, prv) in enumerate(TILES):
    col = CL(2 + i)
    c = db[f"{col}7"]; c.value = lab
    c.font = Font(name=F, size=9, color="4A5468", bold=True); c.fill = TINT
    c.border = Border(left=THIN, right=THIN, top=MED)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    v = db[f"{col}8"]; v.value = cur
    v.font = Font(name=F, size=18, bold=True, color=NAVY); v.fill = TINT
    v.number_format = NUM
    v.border = Border(left=THIN, right=THIN)
    v.alignment = Alignment(horizontal="center", vertical="center")
    d = db[f"{col}9"]
    d.value = ('' if not prv else f'=IF({col}10=0,"",IF({col}8>={col}10,"▲ ","▼ ")&TEXT(ABS({col}8-{col}10),"#,##0")&" (전월 "&{col}10&")")')
    d.font = Font(name=F, size=9, bold=True, color="186B39"); d.fill = TINT
    d.border = Border(left=THIN, right=THIN, bottom=THIN)
    d.alignment = Alignment(horizontal="center", vertical="center")
    if prv:
        db[f"{col}10"] = prv
        db[f"{col}10"].number_format = NUM
db.row_dimensions[7].height = 17
db.row_dimensions[8].height = 28
db.row_dimensions[9].height = 15
db.row_dimensions[10].hidden = True
put(db, "A7", "구분", LBL_F, fill=SUB_FILL)
put(db, "A8", "이번 달", LBL_F, fill=SUB_FILL)
put(db, "A9", "전월 대비", LBL_F, fill=SUB_FILL)

# ══════════ 2. 확인이 필요한 것   /   3. 사건 기준
section(db, 12, "2.  확인이 필요한 것", span=6)
for j, h in enumerate(["구분", "이번 달", "전체 사건", "신규 사건"], start=8):
    c = db.cell(row=12, column=j, value="")
    c.fill, c.border = SEC_FILL, BOX
db["H12"] = "3.  사건 기준 집계"; db["H12"].font = SEC_F
hdr(db, 13, ["항목","건수","상태","확인할 곳"]); db.merge_cells("D13:F13")
for j, h in enumerate(["구분","회차","사건 수","비고"], start=8):
    c = db.cell(row=13, column=j, value=h)
    c.font, c.fill, c.border = HDR_F, HDR_FILL, BOX
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALERTS = [("입력검증 오류", "='입력검증'!$C$9", "[입력검증]"),
          ("장비 반납 연체", f'=COUNTIF({R("장비·매체 관리","J",LN_F,LN_L)},"연체")', "[장비·매체 관리]"),
          ("30일 초과 진행중", "='진행중 보드'!$C$10", "[진행중 보드]"),
          ("사업 완료예정일 경과", f'=SUMPRODUCT(({R("사업마스터","L",PJ1,PJ2)}>0)*'
                                   f'({R("사업마스터","L",PJ1,PJ2)}<TODAY())*'
                                   f'({R("사업마스터","P",PJ1,PJ2)}<>"완료"))', "[사업마스터]")]
for i, (lab, f_, where) in enumerate(ALERTS):
    r = 14 + i
    put(db, f"A{r}", lab, TXT_F, align="left")
    put(db, f"B{r}", f_, LNK_F, NUM)
    put(db, f"C{r}", f'=IF($B{r}=0,"이상 없음","확인 필요")', FX_F)
    put(db, f"D{r}", where, LNK_F, align="left"); db.merge_cells(f"D{r}:F{r}")
    db.conditional_formatting.add(f"A{r}:F{r}", FormulaRule(
        formula=[f'$B{r}>0'], fill=PatternFill("solid", fgColor="FFC7CE"),
        font=Font(color="9C0006", bold=True), stopIfTrue=True))
    db.conditional_formatting.add(f"A{r}:F{r}", FormulaRule(
        formula=[f'$B{r}=0'], fill=PatternFill("solid", fgColor="E2EFDA"), stopIfTrue=True))
CASE_ROWS = [("출장", "B", "C"), ("분석", "D", "E"), ("교육", "F", "")]
for i, (lab, rc, cc) in enumerate(CASE_ROWS):
    r = 14 + i
    put(db, f"H{r}", lab, TXT_F, align="left")
    put(db, f"I{r}", f'=INDEX({OVM}!${rc}$7:${rc}$18,$D$4)', LNK_F, NUM)
    put(db, f"J{r}", (f'=INDEX({OVM}!${cc}$7:${cc}$18,$D$4)' if cc else '"-"'), LNK_F, NUM)
    put(db, f"K{r}", ("" if cc else "사건과 묶이지 않음"), NOTE_F, align="left")
    db.merge_cells(f"K{r}:M{r}")
put(db, "H17", "전체", LBL_F, fill=TOT_FILL, align="left")
put(db, "I17", f'=INDEX({OVM}!$G$7:$G$18,$D$4)', LNK_F, NUM, TOT_FILL)
put(db, "J17", f'=INDEX({OVM}!$H$7:$H$18,$D$4)', LNK_F, NUM, TOT_FILL)
put(db, "K17", f'=INDEX({OVM}!$I$7:$I$18,$D$4)&" 건 신규 착수"', FX_F, fill=TOT_FILL, align="left")
db.merge_cells("K17:M17")
db["J13"].comment = Comment(
    "고유 사건 수입니다. 같은 사건을 여러 회차로 나눠 했어도 1로 셉니다. 전체 사건은 출장 사건 + 분석 사건이 "
    "아닙니다 — 한 사건에서 둘 다 했으면 전체에는 1로만 잡힙니다.", "업무일지")

# ══════════ 4. 진행중 오래된 순
section(db, 19, "4.  진행중 · 오래된 순 TOP 5   ※ 전체는 [진행중 보드]", span=13)
hdr(db, 20, ["경과일","사건번호","회차","사건명","직무","세부유형","주담당",
             "인원","시작일","장소","진행상태","대상매체","데이터량"])
BDM = "'진행중 보드'"
for k in range(5):
    r, src = 21 + k, 13 + k
    g = f"IF({BDM}!$A{src}=\"\",\"\","
    put(db, f"A{r}", f'={g}{BDM}!$B{src})', LNK_F, NUM)
    put(db, f"B{r}", f'={g}{BDM}!$G{src})', LNK_F)
    put(db, f"C{r}", f'={g}"-")', LNK_F)
    put(db, f"D{r}", f'={g}{BDM}!$H{src})', LNK_F, align="left")
    put(db, f"E{r}", f'={g}{BDM}!$E{src})', LNK_F)
    put(db, f"F{r}", f'={g}{BDM}!$F{src})', LNK_F, align="left")
    put(db, f"G{r}", f'={g}{BDM}!$I{src})', LNK_F)
    put(db, f"H{r}", f'={g}{BDM}!$J{src})', LNK_F, NUM)
    put(db, f"I{r}", f'={g}{BDM}!$C{src})', LNK_F, DATE_)
    put(db, f"J{r}", f'={g}{BDM}!$K{src})', LNK_F, align="left")
    put(db, f"K{r}", f'={g}"진행중")', LNK_F)
    put(db, f"L{r}", f'={g}"-")', LNK_F)
    put(db, f"M{r}", f'={g}"-")', LNK_F)
    db.conditional_formatting.add(f"A{r}:M{r}", FormulaRule(
        formula=[f'AND($A{r}<>"",$A{r}>60)'], fill=PatternFill("solid", fgColor="FFC7CE"),
        font=Font(color="9C0006", bold=True), stopIfTrue=True))
    db.conditional_formatting.add(f"A{r}:M{r}", FormulaRule(
        formula=[f'AND($A{r}<>"",$A{r}>30)'], fill=PatternFill("solid", fgColor="FCE4D6"), stopIfTrue=True))

# ══════════ 5. 팀별 (이번 달 아님 · 연간)
section(db, 27, "5.  팀별 실적 (연간 누계)", span=13)
hdr(db, 28, ["소속팀","출장 회차","분석 회차","교육 회차","전체 회차","구성비",
             "출장 사건","분석 사건","전체 사건","소요일수","연인원","대상매체","진행중"])
for i in range(3):
    r, src = 29 + i, 44 + i
    g = f'IF({OVM}!$A{src}="","",'
    put(db, f"A{r}", f'={g}{OVM}!$A{src})', LNK_F, align="left")
    for col, sc in (("B","B"),("C","C"),("D","D"),("E","E")):
        put(db, f"{col}{r}", f'={g}{OVM}!${sc}{src})', LNK_F, NUM)
    put(db, f"F{r}", f'=IF(OR($E{r}="",{OVM}!$E$54=0),"",REPT("|",ROUND($E{r}/{OVM}!$E$54*20,0))&" "&TEXT($E{r}/{OVM}!$E$54,"0%"))',
        FX_F, align="left")
    for col, sc in (("G","J"),("H","K"),("I","L"),("J","F"),("K","G"),("L","H"),("M","I")):
        put(db, f"{col}{r}", f'={g}{OVM}!${sc}{src})', LNK_F, NUM)
put(db, "A32", "합계", LBL_F, fill=TOT_FILL, align="left")
for col, sc in (("B","B"),("C","C"),("D","D"),("E","E"),("G","J"),("H","K"),("I","L"),
                ("J","F"),("K","G"),("L","H"),("M","I")):
    put(db, f"{col}32", f'={OVM}!${sc}$54', LNK_F, NUM, TOT_FILL)
put(db, "F32", "100%", FX_F, fill=TOT_FILL, align="left")

# ══════════ 6. 사업 현황
section(db, 34, "6.  사업 현황", span=13)
hdr(db, 35, ["사업번호","사업명","구분","주관팀","현재 단계","진척률",
             "예산액(원)","집행액(원)","집행률","추진 건수","완료예정일","진행상태","비고"])
PMM = "사업마스터"
for i in range(8):
    r, src = 36 + i, PJ1 + i
    g = f'IF({PMM}!$A{src}="","",'
    for col, sc, fmt in (("A","A",None),("B","B",None),("C","D",None),("D","E",None),
                         ("E","M",None),("F","N",PCT0),("G","G",MONEY),("H","H",MONEY),
                         ("I","I",PCT),("J","O",NUM),("K","L",DATE_),("L","P",None),("M","Q",None)):
        put(db, f"{col}{r}", f'={g}{PMM}!${sc}{src})', LNK_F, fmt,
            align="left" if col in ("B","M") else None)
    db.conditional_formatting.add(f"A{r}:M{r}", FormulaRule(
        formula=[f'$L{r}="완료"'], fill=PatternFill("solid", fgColor="E2EFDA"), stopIfTrue=True))
    db.conditional_formatting.add(f"A{r}:M{r}", FormulaRule(
        formula=[f'AND($K{r}<>"",$K{r}<TODAY(),$L{r}<>"완료")'],
        fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(color="9C0006", bold=True), stopIfTrue=True))
put(db, "A44", "합계", LBL_F, fill=TOT_FILL, align="left")
put(db, "B44", f'=IF(COUNTBLANK({PMM}!$A${PJ1}:$A${PJ2})={PJ2-PJ1+1},"등록된 사업 없음",'
               f'({PJ2-PJ1+1})-COUNTBLANK({PMM}!$A${PJ1}:$A${PJ2})&" 개 사업")', FX_F, fill=TOT_FILL, align="left")
put(db, "G44", f'=SUM({R(PMM,"G",PJ1,PJ2)})', LNK_F, MONEY, TOT_FILL)
put(db, "H44", f'=SUM({R(PMM,"H",PJ1,PJ2)})', LNK_F, MONEY, TOT_FILL)
put(db, "I44", '=IF($G$44=0,"",$H$44/$G$44)', FX_F, PCT, TOT_FILL)
put(db, "J44", f'=SUM({R(PMM,"O",PJ1,PJ2)})', LNK_F, NUM, TOT_FILL)
for col in ("C","D","E","F","K","L","M"): put(db, f"{col}44", "", FX_F, fill=TOT_FILL)
db["A46"] = ("※ 상위 8개 사업만 표시합니다. 전체는 [사업마스터], 사업별 경과는 [사업별 현황]을 보세요.  "
             "※ 집행률 합계는 예산이 배정된 사업만 기준으로 계산됩니다.")
db["A46"].font = NOTE_F
db["A47"] = "※ 팀별 실적은 연간 누계입니다(이번 달이 아님). 월별 추이는 [부서 전체현황]을 보세요."
db["A47"].font = NOTE_F
db["I35"].comment = Comment(
    "예산이 배정되지 않은 사업은 집행률이 빈칸으로 남습니다. 합계 집행률도 예산 있는 사업만 기준입니다.", "업무일지")

# ── 시트 순서: 대시보드 맨 앞
ORDER = ["대시보드","사용안내","월력표","업무일지","일일실시사항","사건마스터",
         "사업마스터","사업추진일지","장비·매체 관리","진행중 보드",
         "사건 이력 조회","사업별 현황","부서 전체현황","사건별 현황","개인별 현황",
         "교육 이수 현황","입력검증","코드설정"]
assert sorted(ORDER) == sorted(wb.sheetnames), set(ORDER) ^ set(wb.sheetnames)
wb._sheets = [wb[n] for n in ORDER]
TAB = {"대시보드":"1F3864","사업마스터":"7030A0","사업추진일지":"7030A0","사업별 현황":"7030A0",
       "교육 이수 현황":"BF8F00"}
for n, col in TAB.items(): wb[n].sheet_properties.tabColor = col
wb.active = 0
wb.save(OUT)
print("Stage 5 완료 — 대시보드 추가 + 시트 18개 재정렬")
