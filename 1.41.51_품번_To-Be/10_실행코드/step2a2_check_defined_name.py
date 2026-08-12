# -*- coding: utf-8 -*-
import openpyxl

SRC = r"C:\Users\정별\1_Work\1.41_동아쏘시오그룹(2)_데이터_표준화\1.41.50_작업중_파일_공유받음(0811)\260428_S-TEPS_입고실적만 ◆_최근3개년_uniq_v.0.5(0811_16시).xlsx"
wb = openpyxl.load_workbook(SRC, data_only=False)
for name, defn in wb.defined_names.items():
    print(name, "->", defn.attr_text if hasattr(defn, "attr_text") else defn)
