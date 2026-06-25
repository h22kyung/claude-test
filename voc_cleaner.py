#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""VOC Excel 정리 자동화 스크립트.

매일 다운로드하는 VOC Excel 파일을 자동으로 정리한다.
  1) 세부시스템 / 답변자이름 필터, row상태=접수 제외
  2) tat(hr)=0 이면서 row상태=이관 인 행 제거
  3) 지정한 12개 컬럼만 지정 순서로 남기고 나머지 삭제
  4) 가/나/다/라 컬럼을 추가해 voc id별 소요시간을 답변자 컬럼에 기록

사용법:
    python voc_cleaner.py 입력.xlsx [출력.xlsx]
"""

import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

# ─────────────────────────────────────────────────────────────────────────────
# CONFIG ── 실제 값에 맞게 이 부분만 수정하세요.
# ─────────────────────────────────────────────────────────────────────────────
SUBSYSTEM_FILTER = ["a", "b", "c"]            # 세부시스템: 이 값만 남김
RESPONDER_FILTER = ["가", "나", "다", "라"]    # 답변자이름: 이 값만 남김 (= 추가 컬럼명)
MAIN_SHEET_NAME = 0                            # 메인 데이터 시트 (시트명 또는 0=첫 시트)
TIME_SHEET_NAME = "소요시간"                    # 소요시간 표가 있는 시트명

# 최종 출력 컬럼(순서 고정)
FINAL_COLUMNS = [
    "voc id",
    "voc유형",
    "제목",
    "row상태",
    "voc등록자",
    "voc등록자id",
    "voc등록자부서",
    "voc등록일",
    "현재상태",
    "처리결과",
    "답변자이름",
    "답변일",
]

# 필터/제거에 사용하는 컬럼명
COL_SUBSYSTEM = "세부시스템"
COL_RESPONDER = "답변자이름"
COL_ROW_STATE = "row상태"
COL_TAT = "tat(hr)"
COL_VOC_ID = "voc id"
COL_DURATION = "소요시간"

EXCLUDE_ROW_STATE = "접수"   # row상태가 이 값이면 제외 (요구사항 1)
DROP_ROW_STATE = "이관"      # tat=0 + 이 값이면 제거 (요구사항 2)
# ─────────────────────────────────────────────────────────────────────────────


def _norm(name: str) -> str:
    """컬럼명 비교용 정규화: 공백 제거 + 소문자."""
    return str(name).replace(" ", "").lower()


def _resolve_col(df: pd.DataFrame, wanted: str) -> str:
    """원하는 컬럼명을 실제 DataFrame 컬럼명으로 해석(공백/대소문자 무시)."""
    target = _norm(wanted)
    for col in df.columns:
        if _norm(col) == target:
            return col
    raise KeyError(
        f"필요한 컬럼 '{wanted}' 을(를) 찾을 수 없습니다. "
        f"엑셀의 실제 컬럼: {list(df.columns)}"
    )


def load_sheets(input_path: Path):
    """메인 시트와 소요시간 시트를 읽어 반환."""
    xls = pd.ExcelFile(input_path)

    main_df = pd.read_excel(xls, sheet_name=MAIN_SHEET_NAME)

    # 소요시간 시트 이름을 공백 무시로 찾는다.
    time_sheet = None
    for name in xls.sheet_names:
        if _norm(name) == _norm(TIME_SHEET_NAME):
            time_sheet = name
            break
    if time_sheet is None:
        raise KeyError(
            f"소요시간 시트 '{TIME_SHEET_NAME}' 을(를) 찾을 수 없습니다. "
            f"엑셀의 실제 시트: {xls.sheet_names}"
        )
    time_df = pd.read_excel(xls, sheet_name=time_sheet)
    return main_df, time_df


def clean(main_df: pd.DataFrame, time_df: pd.DataFrame) -> pd.DataFrame:
    """요구사항 1~4를 순서대로 적용한 결과 DataFrame을 반환."""
    df = main_df.copy()

    # 실제 컬럼명 해석
    c_sub = _resolve_col(df, COL_SUBSYSTEM)
    c_resp = _resolve_col(df, COL_RESPONDER)
    c_state = _resolve_col(df, COL_ROW_STATE)
    c_tat = _resolve_col(df, COL_TAT)
    c_vocid = _resolve_col(df, COL_VOC_ID)

    # 문자열 비교 시 앞뒤 공백 제거를 위한 헬퍼
    def s(col):
        return df[col].astype(str).str.strip()

    # ── 요구사항 1: 필터 ──────────────────────────────────────────────
    mask = (
        s(c_sub).isin([str(x).strip() for x in SUBSYSTEM_FILTER])
        & s(c_resp).isin([str(x).strip() for x in RESPONDER_FILTER])
        & (s(c_state) != EXCLUDE_ROW_STATE)
    )
    df = df[mask]

    # ── 요구사항 2: tat=0 & row상태=이관 제거 ─────────────────────────
    tat_num = pd.to_numeric(df[c_tat], errors="coerce")
    drop_mask = (tat_num == 0) & (s(c_state) == DROP_ROW_STATE)
    df = df[~drop_mask]

    # ── 요구사항 3: 컬럼 선별·재정렬 ──────────────────────────────────
    final_actual = [_resolve_col(df, c) for c in FINAL_COLUMNS]
    result = df[final_actual].copy()
    result.columns = FINAL_COLUMNS  # 표준 컬럼명으로 통일

    # ── 요구사항 4: 가/나/다/라 컬럼 추가 + 소요시간 매핑 ─────────────
    t_vocid = _resolve_col(time_df, COL_VOC_ID)
    t_dur = _resolve_col(time_df, COL_DURATION)
    # voc id -> 소요시간 (문자열 키로 통일)
    duration_map = {
        str(k).strip(): v
        for k, v in zip(time_df[t_vocid], time_df[t_dur])
    }

    for responder in RESPONDER_FILTER:
        # object dtype로 생성해야 숫자/문자 소요시간을 모두 담을 수 있다.
        result[responder] = pd.Series([""] * len(result), index=result.index, dtype=object)

    voc_ids = result["voc id"].astype(str).str.strip()
    responders = result["답변자이름"].astype(str).str.strip()
    for idx, voc_id, responder in zip(result.index, voc_ids, responders):
        if voc_id in duration_map and responder in RESPONDER_FILTER:
            result.at[idx, responder] = duration_map[voc_id]

    return result


def build_output_path(input_path: Path, output_arg: str | None) -> Path:
    if output_arg:
        return Path(output_arg)
    today = datetime.now().strftime("%Y%m%d")
    return input_path.with_name(f"{input_path.stem}_정리_{today}.xlsx")


def main(argv: list[str]) -> int:
    if len(argv) < 2:
        print("사용법: python voc_cleaner.py 입력.xlsx [출력.xlsx]")
        return 1

    input_path = Path(argv[1])
    if not input_path.exists():
        print(f"입력 파일을 찾을 수 없습니다: {input_path}")
        return 1

    output_path = build_output_path(input_path, argv[2] if len(argv) > 2 else None)

    try:
        main_df, time_df = load_sheets(input_path)
        result = clean(main_df, time_df)
    except KeyError as e:
        print(f"오류: {e}")
        return 1

    result.to_excel(output_path, index=False, engine="openpyxl")
    print(f"완료: {len(result)}행 → {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
