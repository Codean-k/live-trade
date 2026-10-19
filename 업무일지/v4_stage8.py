# -*- coding: utf-8 -*-
"""Stage 8 — 사건별 현황에 지원 건수 연결 + 대시보드 증거물 블록 + 입력검증 확장"""
import sys; sys.path.insert(0, "/home/user/live-trade/업무일지")
from v4_common import *

wb = openpyxl.load_workbook(OUT)
SA = R("지원 대장", "A", SP1, SP2); SB = R("지원 대장", "B", SP1, SP2)
SF = R("지원 대장", "F", SP1, SP2); SN = R("지원 대장", "N", SP1, SP2)
SP_ = R("지원 대장", "P", SP1, SP2); SQ = R("지원 대장", "Q", SP1, SP2)
SI = R("지원 대장", "I", SP1, SP2); SH = R("지원 대장", "H", SP1, SP2)
SJ = R("지원 대장", "J", SP1, SP2); SM = R("지원 대장", "M", SP1, SP2)
SE = R("지원 대장", "E", SP1, SP2)
KM_A = R("사건마스터", "A", K1, K2)

# ══════════════════════ 1. 사건별 현황: 지원 건수 (T열 — 숨은 헬퍼 뒤)
cv = wb["사건별 현황"]
for j, h in enumerate(["지원 점수", "보관중"], start=20):     # T, U
    c = cv.cell(row=5, column=j, value=h)
    c.font, c.fill, c.border = HDR_F, HDR_FILL, BOX
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
for i in range(K2 - K1 + 1):
    r = 6 + i
    g = f'IF($A{r}="","",'
    put(cv, f"T{r}", f'={g}COUNTIF({SB},$A{r}))', LNK_F, NUM)
    put(cv, f"U{r}", f'={g}SUMPRODUCT(({SB}=$A{r})*({SQ}="")*({SA}<>"")))', FX_F, NUM)
put(cv, "T66", f'=SUM(T6:T65)', FX_F, NUM, TOT_FILL)
put(cv, "U66", f'=SUM(U6:U65)', FX_F, NUM, TOT_FILL)
widths(cv, {"T": 11, "U": 10})
cv["T5"].comment = Comment(
    "[지원 대장]에서 이 사건번호로 접수된 증거물 점수입니다. 한 사건에 지원번호가 여러 개 붙을 수 있습니다.", "업무일지")
cv["U5"].comment = Comment("아직 반환하지 않은(반환일이 빈) 증거물 점수입니다.", "업무일지")
# T·U 는 숨은 헬퍼(Q~S) 뒤라 보이도록 명시
for c_ in ("T", "U"): cv.column_dimensions[c_].hidden = False
cv.print_area = "A1:U92"

# ══════════════════════ 2. 대시보드: 증거물 지원 현황 블록
db = wb["대시보드"]
R_EV = 49
section(db, R_EV, "7.  증거물 지원 현황", span=13)
hdr(db, R_EV + 1, ["구분","접수 점수","보관중","분석중","분석완료","반환완료",
                   "수량 합","용량 합(GB)","미반환 30일↑","봉인 누락","중복 번호",
                   "사건 연계","사건 없음"])
r = R_EV + 2
put(db, f"A{r}", "전체", LBL_F, fill=SUB_FILL)
put(db, f"B{r}", f'={SP2-SP1+1}-COUNTBLANK({SA})', LNK_F, NUM)
put(db, f"C{r}", f'=SUMPRODUCT(({SA}<>"")*({SQ}=""))', LNK_F, NUM)
put(db, f"D{r}", f'=COUNTIF({SN},"분석중")', LNK_F, NUM)
put(db, f"E{r}", f'=COUNTIF({SN},"분석완료")', LNK_F, NUM)
put(db, f"F{r}", f'=COUNTIF({SN},"반환완료")', LNK_F, NUM)
put(db, f"G{r}", f'=SUM({SI})', LNK_F, NUM)
put(db, f"H{r}", f'=SUM({SH})', LNK_F, GB)
put(db, f"I{r}", f'=SUMPRODUCT(({SP_}>0)*({SQ}="")*(TODAY()-{SP_}>30))', LNK_F, NUM)
put(db, f"J{r}", f'=SUMPRODUCT(({SA}<>"")*({SJ}=""))', LNK_F, NUM)
put(db, f"K{r}", f'=SUMPRODUCT(({SA}<>"")*(COUNTIF({SA},{SA}&"")>1))', LNK_F, NUM)
put(db, f"L{r}", f'=SUMPRODUCT(({SA}<>"")*({SB}<>""))', LNK_F, NUM)
put(db, f"M{r}", f'=SUMPRODUCT(({SA}<>"")*({SB}=""))', LNK_F, NUM)
r2 = r + 1
put(db, f"A{r2}", "이번 달 접수", LBL_F, fill=SUB_FILL)
put(db, f"B{r2}", f'=COUNTIFS({SE},">="&DATE($B$4,$D$4,1),{SE},"<"&DATE($B$4,$D$4+1,1))', LNK_F, NUM)
for L in "CDEFGHIJKLM": put(db, f"{L}{r2}", "", FX_F)
for col in ("I", "J", "K"):
    db.conditional_formatting.add(f"{col}{r}", FormulaRule(
        formula=[f'${col}${r}>0'], fill=PatternFill("solid", fgColor="FFC7CE"),
        font=Font(color="9C0006", bold=True), stopIfTrue=True))
db[f"A{r2+2}"] = ("※ 보관중 = 반환일이 비어 있는 증거물입니다.  ※ 사건 없음 = 사건번호를 적지 않은 지원건(타 기관 지원·자체 인지 등)입니다.  "
                  "※ 상세는 [지원 대장]을 보세요.")
db[f"A{r2+2}"].font = NOTE_F
# 대시보드 경고 블록에 증거물 항목 추가 (18행)
r3 = 18
put(db, f"A{r3}", "증거물 미반환 30일↑", TXT_F, align="left")
put(db, f"B{r3}", f'=SUMPRODUCT(({SP_}>0)*({SQ}="")*(TODAY()-{SP_}>30))', LNK_F, NUM)
put(db, f"C{r3}", f'=IF($B{r3}=0,"이상 없음","확인 필요")', FX_F)
put(db, f"D{r3}", "[지원 대장]", LNK_F, align="left"); db.merge_cells(f"D{r3}:F{r3}")
db.conditional_formatting.add(f"A{r3}:F{r3}", FormulaRule(
    formula=[f'$B{r3}>0'], fill=PatternFill("solid", fgColor="FFC7CE"),
    font=Font(color="9C0006", bold=True), stopIfTrue=True))
db.conditional_formatting.add(f"A{r3}:F{r3}", FormulaRule(
    formula=[f'$B{r3}=0'], fill=PatternFill("solid", fgColor="E2EFDA"), stopIfTrue=True))
db.print_area = f"A1:M{r2+2}"

# ══════════════════════ 3. 입력검증: 지원 대장 블록 추가
vf = wb["입력검증"]
V_S1 = 2097
V_S2 = V_S1 + (SP2 - SP1)
put(vf, "A9", "지원 대장", LBL_F, fill=SUB_FILL, align="left")
put(vf, "B9", f'={SP2-SP1+1}-COUNTBLANK({SA})', LNK_F, NUM)
put(vf, "C9", f'={V_S2-V_S1+1}-COUNTBLANK($D${V_S1}:$D${V_S2})', FX_F, NUM)
put(vf, "D9", '=IF($C9=0,"이상 없음","확인 필요 — "&$C9&"개 행")', FX_F, align="left")
for r in (10,):
    put(vf, f"A{r}", "합계", LBL_F, fill=TOT_FILL, align="left")
    put(vf, f"B{r}", '=SUM(B6:B9)', FX_F, NUM, TOT_FILL)
    put(vf, f"C{r}", '=SUM(C6:C9)', FX_F, NUM, TOT_FILL)
    put(vf, f"D{r}", '=IF($C$10=0,"전체 이상 없음","총 "&$C$10&"개 행을 확인하세요")', FX_F, fill=TOT_FILL, align="left")
for rr in (9, 10):
    vf.conditional_formatting.add(f"A{rr}:D{rr}", FormulaRule(
        formula=[f'$C{rr}>0'], fill=PatternFill("solid", fgColor="FFC7CE"),
        font=Font(color="9C0006", bold=True), stopIfTrue=True))
    vf.conditional_formatting.add(f"A{rr}:D{rr}", FormulaRule(
        formula=[f'$C{rr}=0'], fill=PatternFill("solid", fgColor="E2EFDA"), stopIfTrue=True))
s = lambda c, row: f"'지원 대장'!${c}{row}"
for i in range(SP2 - SP1 + 1):
    src, r = SP1 + i, V_S1 + i
    act = f'COUNTBLANK(\'지원 대장\'!$A{src}:$T{src})<20'
    put(vf, f"A{r}", "지원 대장", TXT_F)
    put(vf, f"B{r}", src, TXT_F, NUM)
    put(vf, f"C{r}", f'=IF($D{r}="","",{s("A",src)}&"  "&{s("F",src)}&"  "&{s("G",src)})', FX_F, align="left")
    put(vf, f"D{r}", '=IF(NOT(' + act + '),"",TRIM('
        f'IF({s("A",src)}="","지원번호 누락; ","")'
        f'&IF(COUNTIF({SA},{s("A",src)}&"")>1,"지원번호 중복; ","")'
        f'&IF({s("E",src)}="","접수일 누락; ","")'
        f'&IF(AND({s("B",src)}<>"",COUNTIF({KM_A},{s("B",src)}&"")=0),"미등록 사건번호; ","")'
        f'&IF(COUNTIF(코드설정!$V${MD1}:$V${MD2},{s("F",src)})=0,"매체 종류 누락; ","")'
        f'&IF({s("J",src)}="","봉인번호 누락; ","")'
        f'&IF(AND({s("M",src)}="",{s("Q",src)}=""),"보관 위치 누락; ","")'
        f'&IF(COUNTIF(코드설정!$X${AS1}:$X${AS2},{s("N",src)})=0,"분석 상태 누락; ","")'
        f'&IF(AND({s("O",src)}<>"",{s("P",src)}<>"",{s("P",src)}<{s("O",src)}),"완료일이 착수일보다 빠름; ","")'
        f'&IF(AND({s("Q",src)}<>"",{s("E",src)}<>"",{s("Q",src)}<{s("E",src)}),"반환일이 접수일보다 빠름; ","")'
        f'&IF(AND({s("P",src)}<>"",{s("Q",src)}="",TODAY()-{s("P",src)}>30),"분석완료 후 30일 초과 미반환; ","")'
        '))', FX_F, align="left")
vf.auto_filter.ref = f"A36:D{V_S2}"
vf.conditional_formatting.add(f"A{V_S1}:D{V_S2}", FormulaRule(
    formula=[f'$D{V_S1}<>""'], fill=PatternFill("solid", fgColor="FFF2CC"), stopIfTrue=False))
# ── 지원 대장 오류 유형은 기존 목록(13~33행) 오른쪽에 별도 블록으로
#    (35행 섹션 · 36행 헤더 · 37행부터 상세 데이터이므로 아래로 늘릴 수 없음)
NEWT = ["지원번호 누락","지원번호 중복","접수일 누락","매체 종류 누락","봉인번호 누락",
        "보관 위치 누락","분석 상태 누락","완료일이 착수일보다 빠름","반환일이 접수일보다 빠름",
        "분석완료 후 30일 초과 미반환"]
widths(vf, {"E": 3, "F": 34, "G": 3, "H": 3, "I": 10})
for j, h in enumerate(["지원 대장 오류 유형", "", "", "건수"], start=6):
    c = vf.cell(row=12, column=j, value=h)
    c.font, c.fill, c.border = HDR_F, HDR_FILL, BOX
    c.alignment = Alignment(horizontal="center", vertical="center")
vf.merge_cells("F12:H12")
for i, t in enumerate(NEWT):
    r = 13 + i
    put(vf, f"F{r}", t, TXT_F, fill=SUB_FILL, align="left")
    for j in (7, 8):
        c = vf.cell(row=r, column=j); c.border = BOX; c.fill = SUB_FILL
    vf.merge_cells(f"F{r}:H{r}")
    put(vf, f"I{r}", f'=COUNTIF($D${V_S1}:$D${V_S2},"*{t}*")', LNK_F, NUM)
    vf.conditional_formatting.add(f"F{r}:I{r}", FormulaRule(
        formula=[f'$I{r}>0'], fill=PatternFill("solid", fgColor="FCE4D6"), stopIfTrue=False))

# 기존 유형별 건수의 범위도 지원 대장까지 확장
for r in range(13, 34):
    v = vf[f"D{r}"].value
    if isinstance(v, str) and "COUNTIF($D$37:$D$2096" in v:
        vf[f"D{r}"] = v.replace("$D$2096", f"$D${V_S2}")

db["B14"] = "='입력검증'!$C$10"      # 합계 행이 9 → 10 으로 내려감
wb.save(OUT)
print(f"Stage 8 완료 — 사건별 지원 점수 / 대시보드 증거물 블록 / 입력검증 +{SP2-SP1+1}행")
print(f"  입력검증 지원 대장 구간: {V_S1}~{V_S2} · 신규 오류 유형 {len(NEWT)}종")
