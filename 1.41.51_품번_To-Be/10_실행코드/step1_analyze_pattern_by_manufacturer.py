# -*- coding: utf-8 -*-
r"""
[분석 전용 — 파일 변경 없음]

1.41.50_v0.5 파일에서 품번이 있는 행(14,013건) 중 제조사(정리) 상위 20개를 뽑아,
제조사별로 품번의 "형태 지문"(문자->A, 숫자->9, 그 외 기호는 그대로) 분포를 계산한다.

추가로 "핵심코드-수량단위" 형태(예: 09735-250G)인지 판별하는 정규식을 걸어서,
그 제조사 안에서 같은 핵심코드가 서로 다른 수량 접미사로 여러 번 나타나는 사례가 있는지
확인한다 — 있으면 접미사가 진짜 "수량 구분자"라는 근거가 되고, 없으면(혹은 접미사가 코드
구조 자체의 일부로 보이면) 접미사를 함부로 잘라내면 안 된다는 근거가 된다.

이 스크립트는 규칙을 확정하지 않고, 사용자가 규칙을 정할 수 있도록 근거 데이터만 출력한다.
"""
from collections import Counter, defaultdict
import re

import pandas as pd

SRC = r"C:\Users\정별\1_Work\1.41_동아쏘시오그룹(2)_데이터_표준화\1.41.50_작업중_파일_공유받음(0811)\260428_S-TEPS_입고실적만 ◆_최근3개년_uniq_v.0.5(0811_16시).xlsx"
SHEET = "Steps_중복제거_32359"
HEADER_ROW_IDX = 3  # pandas header=3 -> 엑셀 4행

COL_PN = "품번"
COL_MFR = "제조사\n정리"

TOP_N = 20

_SUFFIX_RE = re.compile(r"^(?P<core>.+?)-(?P<qty>\d+(\.\d+)?)\s*(?P<unit>[A-Za-z%μ]*)$")


def fingerprint(v: str) -> str:
    out = []
    for ch in v:
        if ch.isalpha():
            out.append("A")
        elif ch.isdigit():
            out.append("9")
        else:
            out.append(ch)
    return "".join(out)


def main():
    df = pd.read_excel(SRC, sheet_name=SHEET, header=HEADER_ROW_IDX)
    has_pn = df[COL_PN].notna()
    sub = df[has_pn].copy()
    sub[COL_PN] = sub[COL_PN].astype(str).str.strip()

    top_mfrs = sub[COL_MFR].value_counts().head(TOP_N)
    total_pn_rows = len(sub)

    print(f"[전체] 품번 보유 행: {total_pn_rows}건")
    print(f"[상위 {TOP_N}개 제조사 합계] {top_mfrs.sum()}건 ({top_mfrs.sum()/total_pn_rows:.1%})")
    print()

    for mfr, cnt in top_mfrs.items():
        vals = sub.loc[sub[COL_MFR] == mfr, COL_PN].tolist()
        fps = Counter(fingerprint(v) for v in vals)
        top_fps = fps.most_common(3)
        top_fp_coverage = sum(c for _, c in top_fps) / len(vals)

        # 수량접미사 패턴 매칭 + 같은 core가 여러 수량으로 재등장하는지 확인
        core_to_qtys = defaultdict(set)
        n_suffix_match = 0
        for v in vals:
            m = _SUFFIX_RE.match(v)
            if m:
                n_suffix_match += 1
                core_to_qtys[m.group("core")].add(m.group("qty") + m.group("unit"))
        cores_with_multi_qty = {c: qs for c, qs in core_to_qtys.items() if len(qs) > 1}

        print("=" * 70)
        print(f"[{mfr}] 총 {cnt}건")
        print(f"  형태지문 상위3: {top_fps}  (상위3 커버율 {top_fp_coverage:.1%})")
        print(f"  '핵심코드-수량단위' 패턴 매칭: {n_suffix_match}건 ({n_suffix_match/len(vals):.1%})")
        print(f"  그 중 같은 핵심코드가 다른 수량으로 재등장: {len(cores_with_multi_qty)}개 코드")
        if cores_with_multi_qty:
            sample_core = next(iter(cores_with_multi_qty))
            print(f"    예시 core={sample_core!r} -> 수량들={cores_with_multi_qty[sample_core]}")
        print(f"  샘플 품번 5건: {vals[:5]}")


if __name__ == "__main__":
    main()
