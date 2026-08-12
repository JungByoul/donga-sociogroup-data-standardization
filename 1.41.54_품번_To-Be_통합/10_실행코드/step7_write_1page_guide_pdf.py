# -*- coding: utf-8 -*-
r"""품번_To-Be 의견 카테고리 + 사용된 규칙 1페이지 요약 PDF (v6.0+).
1.41.33/40_송부용/260810_CAS_품번_정리결과_v1.0.pdf 와 같은 포맷(제목/문단/표 스타일)으로 맞춤.
"""
from pathlib import Path
from fpdf import FPDF
from fpdf.fonts import FontFace

OUT = Path(r"C:\Users\정별\1_Work\1.41_동아쏘시오그룹(2)_데이터_표준화\1.41.54_품번_To-Be_통합\30_보고서\260812_품번_To-Be_카테고리와규칙_안내_v6.0+.pdf")

FONT_REG = r"C:\Windows\Fonts\malgun.ttf"
FONT_BOLD = r"C:\Windows\Fonts\malgunbd.ttf"

ROWS = [
    ("o>유지", "원본 그대로 둠", "공백 정리 외에는 아무것도 안 바꿈"),
    ("x>수량단위\n제거", "뒤에 붙은 수량+단위를 잘라냄",
     "단위 화이트리스트(G·KG·MG·UG·NG·ML·UL·L·M·MM·EA·AMP·PAK·RXN·TAB·CAP·KT·VL·%), "
     "복합단위(AMP-EA·KG-K·G-F·SET-F·G-K·ML-R), 'NxM단위' 곱셈표기, 앞자리 없는 소수, "
     "하이픈 뒤 공백, 단위 뒤 괄호 브랜드명, 선두 숫자+인치기호, 한글 단위 '-숫자중'"),
    ("x>안내라벨\n제거", "안내 문구를 지움",
     "견적번호 · 부품 번호 · USP · 시리얼 번호 · 일련번호 · Serial number · 모델명 · 끝의 '외'"),
    ("x>불용기호\n제거", "장식성 기호를 지움", "선두 '#' · 전체가 '[...]'로 감싸진 경우 · Agilent 계열 끝 'UI'"),
    ("x>부가정보\n제거", "부가 설명·브랜드명을 지움",
     "끝 콤마+텍스트(마지막 1개) · 끝 '_Unstained' · 개별 확인된 부속품 설명·기타 표기"),
    ("x>상세규격\n이동", "규격/치수로 판단해 품번_To-Be는 비우고 원문은 분리텍스트로 이동",
     "밸브·글러브류 규격문자열 · 치수표기 · 용기규격(숫자+LT·LC, 서술형 용기 옵션)"),
    ("x>의미없음", "코드로 볼 수 없는 서술형 문구라서 품번_To-Be를 비움",
     "개별 확인된 값(약품명+설명, 색상·라인명 등)"),
    ("x>값없음 등\n(7종)", "이전 '4차판정' 작업에서 이미 '품번 아님'으로 판정된 경우",
     "값없음·설명문키워드·CAS의심·품목명동일·숫자없음·단위값·구조식의심 - "
     "1.41.44 폴더 판정 결과를 행 고유번호(key)로 매칭해 그대로 반영"),
]


def main():
    pdf = FPDF(format="A4")
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_margins(18, 18, 18)

    pdf.add_font("Malgun", "", FONT_REG)
    pdf.add_font("Malgun", "B", FONT_BOLD)

    pdf.set_font("Malgun", "B", 22)
    pdf.cell(0, 11, "품번 To-Be 의견 카테고리 & 규칙", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(3)

    pdf.set_font("Malgun", "", 11)
    pdf.cell(0, 6, "S-TEPS 입고실적 최근 3개년 · 품번 To-Be 정제 작업(v2.0~v6.0)", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(4)

    pdf.set_font("Malgun", "", 11)
    pdf.multi_cell(
        0, 6.5,
        "품번_To-Be_의견 컬럼에 표시되는 각 항목의 뜻과, 실제로 어떤 방법(규칙)으로 그렇게 "
        "판단했는지를 정리했습니다.",
    )
    pdf.ln(6)

    pdf.set_font("Malgun", "B", 13)
    pdf.cell(0, 7, "분류 기준", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(2)

    pdf.set_font("Malgun", "", 9.6)
    header_style = FontFace(emphasis="B", fill_color=(235, 235, 235))
    with pdf.table(
        col_widths=(20, 40, 114),
        text_align=("CENTER", "LEFT", "LEFT"),
        line_height=5.6,
        v_align="MIDDLE",
        headings_style=header_style,
        borders_layout="ALL",
    ) as table:
        row = table.row()
        for h in ["의견 표시", "뜻", "사용된 규칙"]:
            row.cell(h)
        for tag, mean, rule in ROWS:
            row = table.row()
            row.cell(tag)
            row.cell(mean)
            row.cell(rule)

    pdf.ln(6)
    pdf.set_font("Malgun", "", 10)
    pdf.multi_cell(
        0, 5.5,
        "예외 3가지: 대한과학·Agilent는 하이픈 뒤 숫자가 수량이 아니라 서로 다른 실제 제품이라 "
        "'수량단위 제거' 규칙 자체를 적용하지 않음. Thermo Fisher는 'KOLAS'로 시작하는 값만 같은 "
        "이유로 제외함. 이 3가지는 상위 20개 회사에서만 적용되고 다른 회사에는 적용하지 않음.",
    )

    pdf.ln(8)
    pdf.set_font("Malgun", "", 9)
    pdf.set_text_color(120, 120, 120)
    pdf.cell(0, 5, "2026-08-12")
    pdf.set_text_color(0, 0, 0)

    OUT.parent.mkdir(parents=True, exist_ok=True)
    pdf.output(str(OUT))
    print(f"[저장] {OUT}  (페이지수: {pdf.pages_count})")


if __name__ == "__main__":
    main()
