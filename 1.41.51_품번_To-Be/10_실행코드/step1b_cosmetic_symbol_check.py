# -*- coding: utf-8 -*-
r"""
[분석 전용 — 파일 변경 없음]

1차 분석에서 나온 힌트(#, [ ] 같은 장식 기호, Agilent/Corning의 '수량접미사'로 보였던 게
사실은 오탐일 가능성) 검증용 보조 스크립트.
"""
import re
from collections import Counter

import pandas as pd

SRC = r"C:\Users\정별\1_Work\1.41_동아쏘시오그룹(2)_데이터_표준화\1.41.50_작업중_파일_공유받음(0811)\260428_S-TEPS_입고실적만 ◆_최근3개년_uniq_v.0.5(0811_16시).xlsx"
SHEET = "Steps_중복제거_32359"
COL_PN = "품번"
COL_MFR = "제조사\n정리"

TARGET_MFRS = ["Sigma", "Sigma-Aldrich", "대한과학", "Cell Signaling Technology", "Sartorius", "Agilent", "Corning"]


def main():
    df = pd.read_excel(SRC, sheet_name=SHEET, header=3)
    sub = df[df[COL_PN].notna()].copy()
    sub[COL_PN] = sub[COL_PN].astype(str).str.strip()

    for mfr in TARGET_MFRS:
        vals = sub.loc[sub[COL_MFR] == mfr, COL_PN].tolist()
        valset = set(vals)
        n_hash = sum(1 for v in vals if v.startswith("#"))
        n_hash_dup = sum(1 for v in vals if v.startswith("#") and v[1:].strip() in valset)
        n_bracket = sum(1 for v in vals if v.startswith("[") and v.endswith("]"))
        n_bracket_dup = sum(1 for v in vals if v.startswith("[") and v.endswith("]") and v[1:-1] in valset)
        print(f"[{mfr}] 총 {len(vals)}건 / '#'시작 {n_hash}건(제거시 기존값과 중복 {n_hash_dup}건) / '[...]'감싸짐 {n_bracket}건(제거시 중복 {n_bracket_dup}건)")

    print()
    print("=== Agilent 'core=122' 원본 예시 (수량접미사 오탐 의심) ===")
    vals = sub.loc[sub[COL_MFR] == "Agilent", COL_PN].tolist()
    hits = [v for v in vals if re.match(r"^122-\d", v)]
    print(hits[:20])

    print()
    print("=== Corning 'core=COP' 원본 예시 ===")
    vals = sub.loc[sub[COL_MFR] == "Corning", COL_PN].tolist()
    hits = [v for v in vals if v.upper().startswith("COP")]
    print(hits[:20])


if __name__ == "__main__":
    main()
