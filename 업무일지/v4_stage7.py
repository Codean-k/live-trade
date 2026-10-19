# -*- coding: utf-8 -*-
"""Stage 7 — 코드설정(매체 종류·분석 상태) + 지원 대장(증거물 관리)"""
import sys; sys.path.insert(0, "/home/user/live-trade/업무일지")
from v4_common import *

wb = openpyxl.load_workbook(OUT)
cs = wb["코드설정"]

# ══════════════════════ 코드설정: 매체 종류 · 분석 상태
for col, name, vals, w, lim in [
    ("V", "매체 종류", ["하드디스크", "SSD", "휴대폰", "USB 메모리", "메모리카드", "노트북",
                       "데스크톱", "서버 · NAS", "CCTV · DVR", "태블릿", "광디스크",
                       "클라우드 · 계정", "문서 · 서면", "기타"], 16, MD2),
    ("X", "분석 상태", ["접수", "분석대기", "분석중", "분석완료", "반환완료", "폐기", "보류"], 13, AS2),
]:
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

for nm, ref in [("목록_매체종류", f"코드설정!$V$5:$V${MD2}"),
                ("목록_분석상태", f"코드설정!$X$5:$X${AS2}"),
                ("목록_지원번호", f"'지원 대장'!$A${SP1}:$A${SP2}")]:
    if nm in wb.defined_names: del wb.defined_names[nm]
    wb.defined_names.add(DefinedName(nm, attr_text=ref))

# ══════════════════════ 지원 대장
sp = wb.create_sheet("지원 대장")
title(sp, "지원 대장  ·  증거물 접수 · 분석 · 반환 관리",
      "※ 증거물 1점 = 1행입니다. 지원번호를 직접 부여해 입력하세요.  "
      "※ 사건번호는 선택 입력입니다 — 한 사건에 여러 지원번호를 붙여도 되고, 사건번호 없는 지원(타 기관 지원·자체 인지)도 그냥 비워 두면 됩니다.  "
      "※ 반환일을 비워 두면 '보관중'으로 봅니다. 분석완료 후 30일이 지나도록 반환되지 않으면 빨갛게 표시됩니다.")
SP_H = ["지원번호","사건번호","사건명(자동)","요청·의뢰처","접수일","매체 종류",
        "제조사 · 모델","용량(GB)","수량","봉인번호","인계자(제출자)","인수자",
        "보관 위치","분석 상태","분석 착수일","분석 완료일","반환일",
        "보관 일수","처리 일수","비고"]
hdr(sp, 5, SP_H)
section(sp, 4, "1.  지원 · 증거물 대장", span=20)
widths(sp, {"A":15,"B":16,"C":24,"D":17,"E":12,"F":14,"G":22,"H":11,"I":7,"J":14,
            "K":13,"L":11,"M":16,"N":11,"O":12,"P":12,"Q":12,"R":10,"S":10,"T":22})
KM_A = R("사건마스터", "A", K1, K2)
KM_B = R("사건마스터", "B", K1, K2)
KM_C = R("사건마스터", "C", K1, K2)
FXC  = {"C","R","S"}
LEFT = {"C","G","M","T"}
for r in range(SP1, SP2 + 1):
    for j in range(1, 21):
        col = CL(j); c = sp.cell(row=r, column=j); c.border = BOX
        c.font = FX_F if col in FXC else (LNK_F if col == "D" else IN_F)
        c.alignment = Alignment(horizontal="left" if col in LEFT else "center",
                                vertical="center", wrap_text=(col in LEFT))
    MT = f'MATCH($B{r},{KM_A},0)'
    sp[f"C{r}"] = f'=IF($B{r}="","",IFERROR(INDEX({KM_B},{MT}),"※ 미등록 사건번호"))'
    sp[f"D{r}"] = f'=IF($B{r}="","",IFERROR(INDEX({KM_C},{MT}),""))'
    for L in ("E","O","P","Q"): sp[f"{L}{r}"].number_format = DATE_
    sp[f"H{r}"].number_format = GB
    sp[f"I{r}"].number_format = NUM
    put(sp, f"R{r}", f'=IF($E{r}="","",IF($Q{r}="",TODAY()-$E{r},$Q{r}-$E{r}))', FX_F, NUM)
    put(sp, f"S{r}", f'=IF(OR($O{r}="",$P{r}=""),"",$P{r}-$O{r}+1)', FX_F, NUM)
sp.freeze_panes = "C6"
sp.auto_filter.ref = f"A5:T{SP2}"
dv_name(sp, f"B{SP1}:B{SP2}", "목록_사건번호")
dv_name(sp, f"F{SP1}:F{SP2}", "목록_매체종류")
dv_name(sp, f"L{SP1}:L{SP2}", "목록_담당자")
dv_name(sp, f"N{SP1}:N{SP2}", "목록_분석상태")
sp["B5"].comment = Comment(
    "선택 입력입니다. 한 사건에 지원번호를 여러 개 붙여도 되고, 사건과 무관한 지원이면 비워 두세요. "
    "비워 두면 사건명·요청처도 빈칸으로 남습니다.", "업무일지")
sp["D5"].comment = Comment(
    "사건번호를 고르면 [사건마스터]의 요청처가 자동으로 들어옵니다. 사건번호가 없는 지원은 이 칸에 직접 적으세요.", "업무일지")
sp["R5"].comment = Comment(
    "접수일부터 반환일까지의 일수입니다. 아직 반환하지 않았으면 접수일부터 오늘까지를 셉니다.", "업무일지")
sp["S5"].comment = Comment("분석 착수일부터 완료일까지의 일수입니다(양쪽 날짜를 모두 포함).", "업무일지")
_sr = f"A{SP1}:T{SP2}"
sp.conditional_formatting.add(_sr, FormulaRule(
    formula=[f'AND($A{SP1}<>"",$P{SP1}<>"",$Q{SP1}="",TODAY()-$P{SP1}>30)'],
    fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(color="9C0006", bold=True), stopIfTrue=True))
sp.conditional_formatting.add(_sr, FormulaRule(
    formula=[f'AND($A{SP1}<>"",COUNTIF($A${SP1}:$A${SP2},$A{SP1})>1)'],
    fill=PatternFill("solid", fgColor="FFC7CE"), stopIfTrue=True))
sp.conditional_formatting.add(_sr, FormulaRule(
    formula=[f'$N{SP1}="반환완료"'], fill=PatternFill("solid", fgColor="E2EFDA"), stopIfTrue=True))
sp.conditional_formatting.add(_sr, FormulaRule(
    formula=[f'$N{SP1}="폐기"'], fill=PatternFill("solid", fgColor="D9D9D9"), stopIfTrue=True))
sp.conditional_formatting.add(_sr, FormulaRule(
    formula=[f'$N{SP1}="분석중"'], fill=PatternFill("solid", fgColor="FFF2CC"), stopIfTrue=True))

SA = R("지원 대장", "A", SP1, SP2); SB = R("지원 대장", "B", SP1, SP2)
SF = R("지원 대장", "F", SP1, SP2); SH = R("지원 대장", "H", SP1, SP2)
SI = R("지원 대장", "I", SP1, SP2); SJ = R("지원 대장", "J", SP1, SP2)
SM = R("지원 대장", "M", SP1, SP2); SN = R("지원 대장", "N", SP1, SP2)
SP_ = R("지원 대장", "P", SP1, SP2); SQ = R("지원 대장", "Q", SP1, SP2)
SR_ = R("지원 대장", "R", SP1, SP2)
NSP = SP2 - SP1 + 1

R_TOT = SP2 + 1
put(sp, f"A{R_TOT}", "합계", LBL_F, fill=TOT_FILL)
put(sp, f"B{R_TOT}", f'={NSP}-COUNTBLANK({SA})&" 점 접수"', FX_F, fill=TOT_FILL, align="left")
for L in ("C","D","E","F","G","J","K","L","M","N","O","P","Q","T"):
    put(sp, f"{L}{R_TOT}", "", FX_F, fill=TOT_FILL)
put(sp, f"H{R_TOT}", f'=SUM({SH})', FX_F, GB, TOT_FILL)
put(sp, f"I{R_TOT}", f'=SUM({SI})', FX_F, NUM, TOT_FILL)
put(sp, f"R{R_TOT}", f'=IFERROR(SUM({SR_})/({NSP}-COUNTBLANK({SA})),0)', FX_F, DEC, TOT_FILL)
put(sp, f"S{R_TOT}", f'=IFERROR(SUMIFS({R("지원 대장","S",SP1,SP2)},{SP_},">0")/COUNTIFS({SP_},">0"),0)',
    FX_F, DEC, TOT_FILL)

# ── 2. 매체 종류별
R_MD = R_TOT + 2
section(sp, R_MD, "2.  매체 종류별 현황", span=7)
hdr(sp, R_MD + 1, ["매체 종류","접수 점수","수량 합","용량 합(GB)","보관중","반환완료","비중"])
NMD = MD2 - MD1 + 1
R_MDT = R_MD + 2 + NMD
for i in range(NMD):
    r, cr = R_MD + 2 + i, MD1 + i
    md = f'$A{r}'
    g = f'IF({md}="","",'
    put(sp, f"A{r}", f'=IF(코드설정!$V{cr}="","",코드설정!$V{cr})', FX_F, fill=SUB_FILL, align="left")
    put(sp, f"B{r}", f'={g}COUNTIF({SF},{md}))', LNK_F, NUM)
    put(sp, f"C{r}", f'={g}SUMIF({SF},{md},{SI}))', LNK_F, NUM)
    put(sp, f"D{r}", f'={g}SUMIF({SF},{md},{SH}))', LNK_F, GB)
    put(sp, f"E{r}", f'={g}SUMPRODUCT(({SF}={md})*({SQ}="")*({SA}<>"")))', FX_F, NUM)
    put(sp, f"F{r}", f'={g}SUMPRODUCT(({SF}={md})*({SQ}<>"")))', FX_F, NUM)
    put(sp, f"G{r}", f'=IFERROR(B{r}/$B${R_MDT},0)', FX_F, PCT)
put(sp, f"A{R_MDT}", "합계", LBL_F, fill=TOT_FILL)
for L in "BCDEF":
    put(sp, f"{L}{R_MDT}", f'=SUM({L}{R_MD+2}:{L}{R_MDT-1})', FX_F, GB if L == "D" else NUM, TOT_FILL)
put(sp, f"G{R_MDT}", f'=IFERROR(B{R_MDT}/$B${R_MDT},0)', FX_F, PCT, TOT_FILL)

# ── 3. 분석 상태별
R_AS = R_MDT + 2
section(sp, R_AS, "3.  분석 상태별 현황", span=5)
hdr(sp, R_AS + 1, ["분석 상태","점수","비중","평균 보관일","평균 처리일"])
NAS = AS2 - AS1 + 1
R_AST = R_AS + 2 + NAS
for i in range(NAS):
    r, cr = R_AS + 2 + i, AS1 + i
    st = f'$A{r}'
    g = f'IF({st}="","",'
    put(sp, f"A{r}", f'=IF(코드설정!$X{cr}="","",코드설정!$X{cr})', FX_F, fill=SUB_FILL, align="left")
    put(sp, f"B{r}", f'={g}COUNTIF({SN},{st}))', LNK_F, NUM)
    put(sp, f"C{r}", f'=IFERROR(B{r}/$B${R_AST},0)', FX_F, PCT)
    put(sp, f"D{r}", f'={g}IFERROR(SUMIF({SN},{st},{SR_})/B{r},0))', FX_F, DEC)
    put(sp, f"E{r}", f'={g}IFERROR(SUMIFS({R("지원 대장","S",SP1,SP2)},{SN},{st},{SP_},">0")/COUNTIFS({SN},{st},{SP_},">0"),0))',
        FX_F, DEC)
put(sp, f"A{R_AST}", "합계", LBL_F, fill=TOT_FILL)
put(sp, f"B{R_AST}", f'=SUM(B{R_AS+2}:B{R_AST-1})', FX_F, NUM, TOT_FILL)
put(sp, f"C{R_AST}", f'=IFERROR(B{R_AST}/$B${R_AST},0)', FX_F, PCT, TOT_FILL)
for L in "DE": put(sp, f"{L}{R_AST}", "", FX_F, fill=TOT_FILL)

# ── 4. 확인이 필요한 것
R_CK = R_AST + 2
section(sp, R_CK, "4.  확인이 필요한 것", span=5)
hdr(sp, R_CK + 1, ["점검 항목","건수","상태","설명"])
CHECKS = [("분석완료 후 30일 초과 미반환",
           f'=SUMPRODUCT(({SP_}>0)*({SQ}="")*(TODAY()-{SP_}>30))', "반환 절차를 진행하세요"),
          ("지원번호 중복",
           f'=SUMPRODUCT(({SA}<>"")*(COUNTIF({SA},{SA}&"")>1))', "같은 번호가 두 번 쓰였습니다"),
          ("미등록 사건번호",
           f'=SUMPRODUCT(({SB}<>"")*(COUNTIF({KM_A},{SB}&"")=0))', "[사건마스터]에 없는 사건번호입니다"),
          ("봉인번호 누락",
           f'=SUMPRODUCT(({SA}<>"")*({SJ}=""))', "증거능력 확보를 위해 반드시 기재하세요"),
          ("보관 위치 미기재 (미반환 건)",
           f'=SUMPRODUCT(({SA}<>"")*({SM}="")*({SQ}=""))', "보관 중인 증거물의 위치를 적으세요"),
          ("매체 종류 미기재",
           f'=SUMPRODUCT(({SA}<>"")*({SF}=""))', "매체 종류별 통계에서 빠집니다")]
for i, (lab, f_, desc) in enumerate(CHECKS):
    r = R_CK + 2 + i
    put(sp, f"A{r}", lab, TXT_F, align="left")
    put(sp, f"B{r}", f_, FX_F, NUM)
    put(sp, f"C{r}", f'=IF($B{r}=0,"이상 없음","확인 필요")', FX_F)
    put(sp, f"D{r}", desc, NOTE_F, align="left"); sp.merge_cells(f"D{r}:E{r}")
    sp.conditional_formatting.add(f"A{r}:E{r}", FormulaRule(
        formula=[f'$B{r}>0'], fill=PatternFill("solid", fgColor="FFC7CE"),
        font=Font(color="9C0006", bold=True), stopIfTrue=True))
    sp.conditional_formatting.add(f"A{r}:E{r}", FormulaRule(
        formula=[f'$B{r}=0'], fill=PatternFill("solid", fgColor="E2EFDA"), stopIfTrue=True))
R_END = R_CK + 2 + len(CHECKS)
sp[f"A{R_END+1}"] = "※ 보관중 = 반환일이 비어 있는 건입니다. 폐기한 증거물은 분석 상태를 '폐기'로 두고 반환일에 폐기일을 적으세요."
sp[f"A{R_END+1}"].font = NOTE_F
sp[f"A{R_END+2}"] = "※ 한 사건에 지원번호가 여러 개 붙을 수 있고, 사건번호 없는 지원도 그대로 관리됩니다."
sp[f"A{R_END+2}"].font = NOTE_F

sp.page_setup.orientation = "landscape"
sp.page_setup.paperSize = sp.PAPERSIZE_A4
sp.page_setup.fitToWidth = 1; sp.page_setup.fitToHeight = 0
sp.sheet_properties.pageSetUpPr.fitToPage = True
sp.page_margins.left = sp.page_margins.right = 0.35
sp.print_options.horizontalCentered = True
sp.print_title_rows = "5:5"
sp.oddHeader.left.text = '&"Arial"&9지원 대장'
sp.oddHeader.right.text = '&"Arial"&9인쇄 &D'
sp.oddFooter.center.text = '&"Arial"&9&P / &N'
sp.sheet_properties.tabColor = "375623"

ORDER = ["대시보드","사용안내","월력표","업무일지","일일실시사항","사건마스터","지원 대장",
         "사업마스터","사업추진일지","장비·매체 관리","진행중 보드",
         "사건 이력 조회","사업별 현황","부서 전체현황","사건별 현황","개인별 현황",
         "교육 이수 현황","입력검증","코드설정"]
assert sorted(ORDER) == sorted(wb.sheetnames), set(ORDER) ^ set(wb.sheetnames)
wb._sheets = [wb[n] for n in ORDER]
wb.active = 0
wb.save(OUT)
print(f"Stage 7 완료 — 지원 대장 ({NSP}점 / 매체 {NMD}종 / 상태 {NAS}종 / 점검 {len(CHECKS)}항목)")
print(f"  대장 {SP1}~{SP2}행 · 합계 {R_TOT} · 매체별 {R_MD} · 상태별 {R_AS} · 점검 {R_CK}")
