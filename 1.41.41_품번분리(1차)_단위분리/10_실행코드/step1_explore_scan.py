# -*- coding: utf-8 -*-
r"""
O열(품번) 1차 판정 - 탐색용 스캔 (읽기 전용, 파일 저장 없음)
값없음/CAS의심은 확실한 규칙으로 걸러내고, 비품번 후보는 목록으로 뽑아서
사용자 확인 후 최종 규칙을 확정하기 위한 스크립트.
"""
import re
from pathlib import Path

import openpyxl

SRC = Path(r"C:\Users\정별\1_Work\1.41_동아쏘시오그룹(2)_데이터_표준화\1.41.33_CAS_품번_분리_최종(Q열)\40_송부용\260810_S-TEPS_입고실적만 ◆_최근3개년_uniq_최종.xlsx")

SHEET = "Steps_중복제거_32359"
DATA_START_ROW = 5
COL_O_품번 = 15
COL_Q_CAS원본 = 17
COL_CASNO_수정 = 19  # S: CAS No.(수정)

_CAS_RE = re.compile(r"\b(\d{2,7}-\d{2}-\d)\b")


def cas_checksum_valid(cas: str) -> bool:
    digits = cas.replace("-", "")
    body, check = digits[:-1], int(digits[-1])
    total = sum(int(d) * (i + 1) for i, d in enumerate(reversed(body)))
    return total % 10 == check


PLACEHOLDER_EXACT = {"-", "해당없음", "품번따로없음", "없음", "n/a", "na"}
DESC_KEYWORDS = [
    "mouse", "female", "male", "system", "server", "license", "standard",
    "core", "ref.", "type ", "week", "cage", "centrifuge", "triple", "quad",
    "agilent", "sciex", "windows", "주령", "시험 중",
]


def main():
    wb = openpyxl.load_workbook(SRC, data_only=False, read_only=True)
    ws = wb[SHEET]

    none_cnt = placeholder_cnt = cas_cnt = desc_cnt = item_cnt = 0
    cas_rows = []
    desc_rows = []

    for r, row in enumerate(ws.iter_rows(min_row=DATA_START_ROW), start=DATA_START_ROW):
        v = row[COL_O_품번 - 1].value
        if v in (None, ""):
            none_cnt += 1
            continue
        s = str(v).strip()
        s_lower = s.lower()

        if s_lower in PLACEHOLDER_EXACT or len(s) <= 2:
            placeholder_cnt += 1
            continue

        m = _CAS_RE.search(s)
        if m and cas_checksum_valid(m.group(1)):
            cas_cnt += 1
            q_val = row[COL_Q_CAS원본 - 1].value
            cas_final = row[COL_CASNO_수정 - 1].value
            cas_rows.append((r, s, q_val, cas_final))
            continue

        if "http" in s_lower or any(k in s_lower for k in DESC_KEYWORDS):
            desc_cnt += 1
            desc_rows.append((r, s))
            continue

        item_cnt += 1

    total = none_cnt + placeholder_cnt + cas_cnt + desc_cnt + item_cnt
    print(f"[분포] 공란: {none_cnt} / 값없음플레이스홀더: {placeholder_cnt} / CAS의심: {cas_cnt} / 비품번후보(키워드): {desc_cnt} / 품번(잠정): {item_cnt} / 합계: {total}")
    print()
    print("=== CAS의심 목록 (O열 값, Q열 원본, CAS No.(수정)) ===")
    for r, s, q, casf in cas_rows:
        print(f"  {r} | O={s} | Q={q} | CASNo수정={casf}")
    print()
    print(f"=== 비품번 후보(키워드 매칭) {len(desc_rows)}건 ===")
    for r, s in desc_rows:
        print(f"  {r} | {s}")


if __name__ == "__main__":
    main()
