import streamlit as st
import pandas as pd
import io
import json
import base64
from datetime import datetime
import plotly.graph_objects as go
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google.oauth2 import service_account
import anthropic

# ============================================================
# 폴더 ID 설정
# ============================================================
FOLDERS = {
    "01_훈련생관리": "17TDAhLp_ohNh0Io2WHVFSeraXSJ-rubk",
    "02_회계관리":   "1gpFnrVNFWEHIvXjoktx-otTxGXPeTaoz",
    "03_성과관리":   "18wBjjN3xKt24cBpINVKVpzwqDtu2af-i",
    "04_수요관리":   "1sHHkE9HYDj7t1BxMPqqm8msw11F-eQBI",
    "05_만족도관리": "1b7oiNCbLaiGz_1Ilx0hjbc3j3GqpOkXP",
    "06_협약관리":   "1A5j2uHlVRLEUozEb6molm5jHIvdi_DWI",
}

# 2026 연간 목표 고정값 (지침서 기준)
TARGET_2026 = 1160

# ============================================================
# Google Drive 연결
# ============================================================
@st.cache_resource
def get_drive_service():
    creds = service_account.Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]),
        scopes=["https://www.googleapis.com/auth/drive.readonly"]
    )
    return build("drive", "v3", credentials=creds)

def list_files(service, folder_id):
    try:
        res = service.files().list(
            q=f"'{folder_id}' in parents and trashed=false",
            fields="files(id,name,mimeType,modifiedTime)",
            orderBy="modifiedTime desc"
        ).execute()
        return res.get("files", [])
    except Exception:
        return []

def get_file_bytes(service, file_id, mime_type):
    """Drive 파일을 바이트로 다운로드"""
    buf = io.BytesIO()
    if mime_type == "application/vnd.google-apps.spreadsheet":
        req = service.files().export_media(
            fileId=file_id,
            mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        req = service.files().get_media(fileId=file_id)
    dl = MediaIoBaseDownload(buf, req)
    done = False
    while not done:
        _, done = dl.next_chunk()
    buf.seek(0)
    return buf.read()

# ============================================================
# [핵심] pandas 기반 엑셀 파싱 함수
# 모든 수치는 pandas로만 연산 — 하드코딩/AI 추출 금지
# ============================================================

def parse_schedule_sheet(file_bytes):
    """
    1번 시트 '2026훈련일정' 파싱
    - skiprows=3 후 row 0을 헤더로 재설정
    - 훈련비과정: no가 숫자인 행만 필터
    - 수료인원, 목표인원 합산
    - 반환: {"actual": int, "target_row": int, "rows": DataFrame}
    """
    try:
        # 시트명 확인 (Google Sheets export 시 이름 변경 가능성 대비)
        xl = pd.ExcelFile(io.BytesIO(file_bytes))
        sheet_name_sched = next(
            (s for s in xl.sheet_names if "훈련일정" in s or "일정" in s),
            xl.sheet_names[0]
        )
        df = pd.read_excel(
            io.BytesIO(file_bytes),
            sheet_name=sheet_name_sched,
            skiprows=3,
            header=None
        )
        # row 0이 실제 컬럼 헤더
        df.columns = [None, 'no', '훈련유형', '훈련과정명', '회차', '일시',
                      '정원', '목표인원', '수료인원', '비고',
                      'x1', '훈련유형별목표', '계획인원', '상하반기', '상하반기계획']
        df = df.iloc[1:].reset_index(drop=True)

        # no가 숫자인 행만 (헤더/합계/지원비과정 섹션 제외)
        df_train = df[pd.to_numeric(df['no'], errors='coerce').notna()].copy()
        df_train['수료인원_n'] = pd.to_numeric(df_train['수료인원'], errors='coerce').fillna(0)
        df_train['목표인원_n'] = pd.to_numeric(df_train['목표인원'], errors='coerce').fillna(0)

        actual = int(df_train['수료인원_n'].sum())
        target_row = int(df_train['목표인원_n'].sum())

        # 훈련비 / 지원비 수료인원 분리 (지원비 헤더 행 기준으로 분할)
        idx_supp = df.index[df['훈련유형별목표'] == '지원비'].tolist()
        if idx_supp:
            split_idx = idx_supp[0]
            actual_train = int(df_train[df_train.index < split_idx]['수료인원_n'].sum())
            actual_supp  = int(df_train[df_train.index > split_idx]['수료인원_n'].sum())
        else:
            actual_train = actual
            actual_supp  = 0

        # 훈련비 연간 계획인원
        plan_mask = df['훈련유형별목표'] == '훈련비'
        plan_train = int(pd.to_numeric(df.loc[plan_mask, '계획인원'], errors='coerce').fillna(0).sum())

        # 폐강률: 목표인원 있고 수료인원 0 또는 NaN인 과정
        df_raw = df[pd.to_numeric(df['no'], errors='coerce').notna()].copy()
        df_raw['목표인원_n'] = pd.to_numeric(df_raw['목표인원'], errors='coerce')
        df_raw['수료인원_raw'] = pd.to_numeric(df_raw['수료인원'], errors='coerce')
        total_courses = df_raw[df_raw['목표인원_n'] > 0]
        closed_courses = total_courses[
            total_courses['수료인원_raw'].isna() | (total_courses['수료인원_raw'] == 0)
        ]
        total_cnt  = len(total_courses)
        closed_cnt = len(closed_courses)
        close_rate = round(closed_cnt / total_cnt * 100, 1) if total_cnt > 0 else 0

        return {
            "actual": actual,
            "actual_train": actual_train,
            "actual_supp": actual_supp,
            "target_row": target_row,
            "plan_train": plan_train,
            "total_courses": total_cnt,
            "closed_courses": closed_cnt,
            "close_rate": close_rate,
            "closed_list": closed_courses[['훈련과정명', '회차', '목표인원_n']].reset_index(drop=True),
            "rows": df_train[['no', '훈련유형', '훈련과정명', '회차', '일시', '수료인원_n', '목표인원_n']]
        }
    except Exception as e:
        return {"actual": 0, "target_row": 0, "plan_train": 880, "rows": pd.DataFrame(), "error": str(e)}


def parse_completion_sheet(file_bytes):
    """
    2번 시트 '2026수료현황' 파싱
    - skiprows=3 후 row 0을 헤더로 재설정
    - 총계 행에서 기업별 수료인원 추출
    - 만족도 4.5점 이하 과정 필터링
    - 반환: dict
    """
    try:
        xl2 = pd.ExcelFile(io.BytesIO(file_bytes))
        sheet_name_comp = next(
            (s for s in xl2.sheet_names if "수료현황" in s or "수료" in s),
            xl2.sheet_names[1]
        )
        df = pd.read_excel(
            io.BytesIO(file_bytes),
            sheet_name=sheet_name_comp,
            skiprows=3,
            header=None
        )
        headers = [None, 'no', '훈련유형', '훈련과정명', '회차', '기간', '정원',
                   '목표인원', '수료인원', 'GMTCK', '르노', '테너지', '불스원',
                   'GMK', '그 외', '업체명', '미수료자', '만족도',
                   'x1', 'x2', 'x3', 'x4', 'x5', 'x6', 'x7']
        df.columns = headers[:len(df.columns)]
        df = df.iloc[1:].reset_index(drop=True)

        # 총계 행 추출 (no == '총계')
        total_row = df[df['no'] == '총계'].iloc[0] if not df[df['no'] == '총계'].empty else None

        # 기업별 수료인원
        companies = {}
        if total_row is not None:
            for col in ['GMTCK', '르노', '테너지', '불스원', 'GMK', '그 외']:
                companies[col] = int(pd.to_numeric(total_row.get(col, 0), errors='coerce') or 0)
        total_actual = sum(companies.values())

        # 과정별 만족도: no가 숫자인 행 + 수료인원 있는 행
        df_sat = df[pd.to_numeric(df['no'], errors='coerce').notna() |
                    (pd.to_numeric(df['수료인원'], errors='coerce').notna() &
                     pd.to_numeric(df['만족도'], errors='coerce').notna())].copy()
        df_sat['만족도_n'] = pd.to_numeric(df_sat['만족도'], errors='coerce')
        df_sat['수료인원_n'] = pd.to_numeric(df_sat['수료인원'], errors='coerce')
        df_sat = df_sat.dropna(subset=['만족도_n'])

        # 만족도 4.5점 이하 과정
        low_sat = df_sat[df_sat['만족도_n'] <= 4.5][['훈련과정명', '회차', '만족도_n', '수료인원_n']].copy()
        low_sat = low_sat.dropna(subset=['훈련과정명'])

        # 평균 만족도
        avg_sat = round(df_sat['만족도_n'].mean(), 2) if not df_sat.empty else 0.0

        return {
            "total_actual": total_actual,
            "companies": companies,
            "avg_sat": avg_sat,
            "sat_table": df_sat[['훈련과정명', '회차', '만족도_n', '수료인원_n']].dropna(subset=['훈련과정명']),
            "low_sat": low_sat,
        }
    except Exception as e:
        return {
            "total_actual": 0, "companies": {}, "avg_sat": 0.0,
            "sat_table": pd.DataFrame(), "low_sat": pd.DataFrame(), "error": str(e)
        }


def parse_yearly_sheet(file_bytes):
    """
    '연도별훈련실적' 시트 파싱
    - 컬럼: 연도 / 계획(훈련비,지원비) / 실적(훈련비,지원비)
    - 반환: DataFrame
    """
    try:
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name="연도별훈련실적", header=None)
        # row 0 = 1단 헤더(연도/계획/실적), row 1 = 2단 헤더(훈련비/지원비)
        df.columns = ['연도', '계획_훈련비', '계획_지원비', '실적_훈련비', '실적_지원비']
        df = df.iloc[2:].reset_index(drop=True)  # 헤더 2행 제거
        df['연도'] = pd.to_numeric(df['연도'], errors='coerce')
        for col in ['계획_훈련비', '계획_지원비', '실적_훈련비', '실적_지원비']:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
        df = df.dropna(subset=['연도'])
        df['연도'] = df['연도'].astype(int)
        return df
    except Exception as e:
        return pd.DataFrame(columns=['연도', '계획_훈련비', '계획_지원비', '실적_훈련비', '실적_지원비'])


def parse_infra_budget(file_bytes):
    """
    '2026년 인프라지원금 집행표' 파싱
    1. 총 예산액: '2026년 인프라지원금 예산' 시트, col 0 == '총계' 행의 col 2
    2. 총 집행액 + 세부내역: '월별집행계' 시트, col 0(집행일자) + col 10(집행액) 실집행 행 합산
    3. 월별 집계, 항목별 집계, 세부 내역 반환
    """
    try:
        # ── 1. 총 예산액 ──
        df_b = pd.read_excel(io.BytesIO(file_bytes),
                             sheet_name="2026년 인프라지원금 예산", header=None)
        total_budget = int(pd.to_numeric(
            df_b.loc[df_b[0].astype(str).str.strip() == '총계', 2],
            errors='coerce'
        ).dropna().iloc[0])

        # ── 2. 실집행 내역 (월별집행계 시트) ──
        df_e = pd.read_excel(io.BytesIO(file_bytes),
                             sheet_name="월별집행계", header=None)
        df_e['집행일자'] = pd.to_datetime(df_e[0], format='mixed', errors='coerce')
        df_e['집행액']  = pd.to_numeric(df_e[10], errors='coerce').fillna(0)
        df_e['집행용도'] = df_e[9].fillna('').astype(str).str.strip()
        df_e['보조세목명'] = df_e[12].fillna('').astype(str).str.strip()

        # 실집행 행: 집행일자 있고 집행액 > 0, 합계/통계 행 제외
        exclude = df_e[1].astype(str).str.contains('합계|통계|청구가능|총 집행', na=False)
        df_detail = df_e[
            df_e['집행일자'].notna() & (df_e['집행액'] > 0) & ~exclude
        ][['집행일자', '보조세목명', '집행용도', '집행액']].copy()
        df_detail = df_detail.reset_index(drop=True)

        total_exec = int(df_detail['집행액'].sum())

        # ── 3. 월별 집계 ──
        df_detail['월'] = df_detail['집행일자'].dt.month
        monthly = (
            df_detail.groupby('월')['집행액']
            .sum()
            .reset_index()
            .rename(columns={'집행액': '집행액'})
        )
        monthly['월'] = monthly['월'].apply(lambda m: f"{m}월")

        # ── 4. 항목별 집계 (보조세목명 기준) ──
        by_category = (
            df_detail.groupby('보조세목명')['집행액']
            .sum()
            .reset_index()
            .sort_values('집행액', ascending=False)
        )

        # ── 5. ERP 예산 시트: 인프라지원금/훈련비/총계 예산·집행 ──
        df_erp = pd.read_excel(io.BytesIO(file_bytes),
                               sheet_name="2026년 ERP 예산", header=None)
        def _get_erp(col0_val, col1_val):
            mask = (df_erp[0].astype(str).str.strip() == col0_val) &                    (df_erp[1].astype(str).str.strip() == col1_val)
            row = df_erp[mask]
            if row.empty:
                return 0, 0
            budget = int(pd.to_numeric(row.iloc[0][3], errors='coerce') or 0)
            exec_  = int(pd.to_numeric(row.iloc[0][8], errors='coerce') or 0)
            return budget, exec_

        infra_b, infra_e = _get_erp('nan', '계')   # 인프라지원금 계 행
        train_b, train_e = _get_erp('nan', '계')    # 아래에서 재탐색
        # 훈련비 계 행은 col0='nan', col1='계' 중 두 번째 (idx=49)
        # col1 == '계' 인 두 행: idx=39(인프라), idx=49(훈련비)
        calc_rows = df_erp[df_erp[1].astype(str).str.strip() == '계']
        if len(calc_rows) >= 2:
            infra_b = int(pd.to_numeric(calc_rows.iloc[0][3], errors='coerce') or 0)
            infra_e = int(pd.to_numeric(calc_rows.iloc[0][8], errors='coerce') or 0)
            train_b = int(pd.to_numeric(calc_rows.iloc[1][3], errors='coerce') or 0)
            train_e = int(pd.to_numeric(calc_rows.iloc[1][8], errors='coerce') or 0)

        total_row = df_erp[df_erp[0].astype(str).str.strip() == '총계']
        grand_b = int(pd.to_numeric(total_row.iloc[0][3], errors='coerce') or 0) if not total_row.empty else 0
        grand_e = int(pd.to_numeric(total_row.iloc[0][8], errors='coerce') or 0) if not total_row.empty else 0

        return {
            "total_budget": total_budget,
            "total_exec": total_exec,
            "exec_rate": round(total_exec / total_budget * 100, 2) if total_budget > 0 else 0,
            "monthly": monthly,
            "by_category": by_category,
            "detail": df_detail,
            # ERP 예산 기반 예산/집행
            "infra_budget": infra_b, "infra_exec": infra_e,
            "infra_rate": round(infra_e / infra_b * 100, 2) if infra_b > 0 else 0,
            "train_budget": train_b, "train_exec": train_e,
            "train_rate": round(train_e / train_b * 100, 2) if train_b > 0 else 0,
            "grand_budget": grand_b, "grand_exec": grand_e,
            "grand_rate": round(grand_e / grand_b * 100, 2) if grand_b > 0 else 0,
        }
    except Exception as e:
        return {
            "total_budget": 0, "total_exec": 0, "exec_rate": 0,
            "monthly": pd.DataFrame(), "by_category": pd.DataFrame(),
            "detail": pd.DataFrame(),
            "infra_budget": 0, "infra_exec": 0, "infra_rate": 0,
            "train_budget": 0, "train_exec": 0, "train_rate": 0,
            "grand_budget": 0, "grand_exec": 0, "grand_rate": 0,
            "error": str(e)
        }


# ============================================================
# 파일 이름 매칭 헬퍼
# ============================================================
XLSX_MIMES = {
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "application/vnd.ms-excel",
    "application/vnd.google-apps.spreadsheet",
}

def is_xlsx(mime):
    return mime in XLSX_MIMES

def is_target_file(fname, keywords):
    return any(k in fname for k in keywords)


# ============================================================
# Drive 파일 수집 + pandas 파싱
# ============================================================
@st.cache_data(ttl=300)
def load_dashboard_data(_service):
    """
    Drive 폴더에서 파일을 수집하고 pandas로 파싱.
    AI는 회계 텍스트 분석에만 제한적으로 사용.
    """
    result = {
        "schedule": None,
        "completion": None,
        "yearly": None,
        "infra": None,       # parse_infra_budget 결과 (pandas 직접 파싱)
    }

    # 모든 폴더를 통합 탐색 — 파일명 키워드로 용도 구분
    all_folders = list(FOLDERS.values())
    all_files = []
    for folder_id in all_folders:
        all_files.extend(list_files(_service, folder_id))

    for f in all_files:
        mime = f["mimeType"]
        # 모든 형식 허용 (디버깅용 — 나중에 복원)
        name = f["name"]
        try:
            # 파일명 직접 매칭
            # "1. 2026 목표실적 및 일정(업데이트 중).xlsx"
            if result["schedule"] is None and "2026" in name and ("목표" in name or "훈련일정" in name):
                raw = get_file_bytes(_service, f["id"], f["mimeType"])
                result["schedule"]   = parse_schedule_sheet(raw)
                result["completion"] = parse_completion_sheet(raw)

            # 02_회계관리: "년도별훈련실적.xlsx"
            elif result["yearly"] is None and "년도별" in name and "훈련실적" in name:
                raw = get_file_bytes(_service, f["id"], f["mimeType"])
                result["yearly"] = parse_yearly_sheet(raw)

            # 02_회계관리: "2. 2026년 인프라지원금 집행표_김하얀.xlsx"
            elif result["infra"] is None and "인프라지원금" in name and "집행표" in name:
                raw = get_file_bytes(_service, f["id"], f["mimeType"])
                result["infra"] = parse_infra_budget(raw)

        except Exception:
            continue

    return result


# ============================================================
# Claude AI — 회계 분석만 담당
# ============================================================
def get_budget_from_claude(budget_text):
    """회계 파일 텍스트 → 예산/집행액만 추출"""
    if not budget_text:
        return {"total": 0, "executed": 0}
    try:
        client = anthropic.Anthropic(api_key=st.secrets["anthropic"]["api_key"])
        msg = client.messages.create(
            model="claude-sonnet-4-5",
            max_tokens=300,
            messages=[{"role": "user", "content": f"""{budget_text}

위 회계 데이터에서 2026년 총예산액과 집행액만 추출해 JSON으로 응답하세요.
마크다운 없이 JSON만:
{{"total": 숫자, "executed": 숫자}}"""}]
        )
        raw = msg.content[0].text.strip().replace("```json", "").replace("```", "").strip()
        return json.loads(raw)
    except Exception:
        return {"total": 0, "executed": 0}


# ============================================================
# HTML 다크 테이블 렌더링 헬퍼
# ============================================================
def dark_table(df: pd.DataFrame, accent: str = "#22D3EE", height: int = None):
    """DataFrame을 다크 글래스모피즘 HTML 테이블로 렌더링"""
    rows_html = ""
    for i, (_, row) in enumerate(df.iterrows()):
        bg = "rgba(255,255,255,0.02)" if i % 2 == 0 else "transparent"
        cells = ""
        for col in df.columns:
            val = row[col]
            # 숫자 우측 정렬
            try:
                float(val)
                align = "right"
                if isinstance(val, (int, float)) and abs(val) >= 1000:
                    val = f"{int(val):,}"
            except (ValueError, TypeError):
                align = "left"
            cells += f'<td style="padding:10px 14px;text-align:{align};color:#E5E9F0;border-bottom:1px solid rgba(255,255,255,0.06);font-size:13px">{val}</td>'
        rows_html += f'<tr style="background:{bg}">{cells}</tr>'

    headers = "".join([
        f'<th style="padding:10px 14px;text-align:left;color:#8A93A6;font-size:12px;font-weight:700;border-bottom:1px solid {accent}55;white-space:nowrap">{col}</th>'
        for col in df.columns
    ])

    height_style = f"max-height:{height}px;overflow-y:auto;" if height else ""
    html = f"""
    <div style="background:rgba(255,255,255,0.04);border-radius:16px;border:1px solid rgba(255,255,255,0.08);{height_style}overflow-x:auto;margin-bottom:1rem">
      <table style="width:100%;border-collapse:collapse">
        <thead><tr style="background:rgba(255,255,255,0.06)">{headers}</tr></thead>
        <tbody>{rows_html}</tbody>
      </table>
    </div>"""
    st.markdown(html, unsafe_allow_html=True)
def progress_bar(rate: float, color: str = "#34D399", height: int = 10):
    """카드 내부용 커스텀 진행바. 텍스트는 강조색, 흰색 사용 안 함."""
    pct = min(float(rate), 100)
    return (
        f'<div style="margin-top:12px">'
        f'<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:5px">'
        f'<span style="font-size:11px;color:#8A93A6">집행 진행도</span>'
        f'<span style="font-size:13px;font-weight:700;color:{color};'
        f'text-shadow:0 0 8px {color}88">{pct}%</span>'
        f'</div>'
        f'<div style="background:rgba(255,255,255,0.07);height:{height}px;'
        f'border-radius:{height}px;overflow:hidden">'
        f'<div style="width:{pct}%;height:100%;border-radius:{height}px;'
        f'background:linear-gradient(90deg,{color}99,{color});'
        f'box-shadow:0 0 10px {color}66"></div>'
        f'</div></div>'
    )



st.set_page_config(page_title="KPI 현황판", page_icon="📊", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;500;700;800&display=swap');

/* ── 배경 ── */
html, body, [data-testid="stAppViewContainer"] {
    background: radial-gradient(ellipse at 60% 20%, #0E1B33 0%, #070C18 100%) !important;
    font-family: 'Noto Sans KR', sans-serif !important;
}
[data-testid="stHeader"] { background: transparent !important; }
[data-testid="stSidebar"] {
    background: #0A1424 !important;
    border-right: 1px solid rgba(255,255,255,0.06) !important;
}
section[data-testid="stMain"] > div { background: transparent !important; }

/* ── 전역 텍스트 ── */
html, body, p, div, span, label, [data-testid] {
    font-family: 'Noto Sans KR', sans-serif !important;
    color: #E5E9F0 !important;
}

/* ── 페이지 타이틀 ── */
.page-title {
    font-size: 42px; font-weight: 800; color: #fff !important;
    text-align: center; letter-spacing: -0.5px;
    margin: 0.5rem 0 0.2rem;
}
.page-sub {
    text-align: center; color: #8A93A6 !important;
    font-size: 14px; margin-bottom: 1.2rem;
}

/* ── 섹션 제목 ── */
.sec {
    font-size: 17px; font-weight: 700; color: #fff !important;
    margin: 1.5rem 0 0.75rem;
}

/* ── 카드 ── */
.card {
    background: rgba(255,255,255,0.04);
    border-radius: 16px; padding: 20px;
    border: 1px solid rgba(255,255,255,0.08);
    position: relative; height: 100%; box-sizing: border-box;
}
.card-label { font-size: 12px; color: #8A93A6 !important; margin-bottom: 8px; }
.card-value { font-size: 34px; font-weight: 700; line-height: 1.1; }
.card-sub   { font-size: 12px; color: #8A93A6 !important; margin-top: 8px; }
.card-icon  { position: absolute; top: 16px; right: 16px; font-size: 18px; }

.card.cyan   { border-color: rgba(34,211,238,0.30);  box-shadow: 0 0 24px rgba(34,211,238,0.10); }
.card.cyan   .card-value { color: #22D3EE !important; }
.card.purple { border-color: rgba(192,132,252,0.30); box-shadow: 0 0 24px rgba(192,132,252,0.10); }
.card.purple .card-value { color: #C084FC !important; }
.card.green  { border-color: rgba(52,211,153,0.30);  box-shadow: 0 0 24px rgba(52,211,153,0.10); }
.card.green  .card-value { color: #34D399 !important; }
.card.amber  { border-color: rgba(251,191,36,0.30);  box-shadow: 0 0 24px rgba(251,191,36,0.10); }
.card.amber  .card-value { color: #FBBF24 !important; }
.card.muted  .card-value { color: #E5E9F0 !important; }

/* ── 배지 ── */
.badge {
    display: inline-block; font-size: 11px; font-weight: 600;
    padding: 2px 8px; border-radius: 6px; margin-top: 6px;
}
.badge.cyan  { background: rgba(34,211,238,0.12);  color: #22D3EE !important; }
.badge.green { background: rgba(52,211,153,0.12);  color: #34D399 !important; }
.badge.amber { background: rgba(251,191,36,0.12);  color: #FBBF24 !important; }
.badge.red   { background: rgba(239,68,68,0.12);   color: #F87171 !important; }

/* ── 만족도 경고 ── */
.warn-card {
    background: rgba(251,191,36,0.08);
    border: 1px solid rgba(251,191,36,0.35);
    border-radius: 12px; padding: 12px 16px; margin-bottom: 8px;
    color: #FBBF24 !important; font-size: 14px;
}

/* ── 탭 ── */
[data-testid="stTabs"] > div:first-child {
    border-bottom: 1px solid rgba(255,255,255,0.08) !important;
    background: transparent !important; gap: 0 !important;
}
button[data-baseweb="tab"] {
    background: transparent !important;
    color: #8A93A6 !important;
    font-size: 14px !important; font-weight: 500 !important;
    border: none !important; padding: 10px 18px !important;
}
button[data-baseweb="tab"][aria-selected="true"] {
    color: #fff !important; font-weight: 700 !important;
    border-bottom: 2px solid #22D3EE !important;
}

/* ── 테이블 ── */
[data-testid="stDataFrame"] iframe { border-radius: 12px !important; }

/* ── Progress ── */
[data-testid="stProgress"] > div > div {
    background: linear-gradient(90deg, #22D3EE, #34D399) !important;
    border-radius: 4px !important;
}
[data-testid="stProgress"] > div {
    background: rgba(255,255,255,0.08) !important; border-radius: 4px !important;
}

/* ── 버튼 ── */
button[kind="secondary"], button[kind="primary"] {
    background: rgba(34,211,238,0.10) !important;
    border: 1px solid rgba(34,211,238,0.35) !important;
    color: #22D3EE !important; border-radius: 10px !important;
    font-weight: 600 !important;
}

/* ── 구분선 ── */
hr { border-color: rgba(255,255,255,0.08) !important; }

/* ── caption ── */
[data-testid="stCaptionContainer"] p { color: #8A93A6 !important; }

/* ── info/success/error 박스 ── */
[data-testid="stAlert"] {
    border-radius: 10px !important;
    background: rgba(255,255,255,0.04) !important;
}

/* ── 탭 패널 배경 ── */
[data-testid="stTabsContent"] {
    background: transparent !important;
    padding-top: 1rem !important;
}

/* ── Plotly 차트 배경 투명 ── */
.js-plotly-plot, .plotly { background: transparent !important; }

/* ── spinner ── */
[data-testid="stSpinner"] { color: #22D3EE !important; }

/* ── selectbox / multiselect ── */
[data-testid="stSelectbox"] > div,
[data-testid="stMultiSelect"] > div {
    background: rgba(255,255,255,0.05) !important;
    border: 1px solid rgba(255,255,255,0.12) !important;
    border-radius: 10px !important;
    color: #E5E9F0 !important;
}
</style>
""", unsafe_allow_html=True)

# 헤더
st.markdown('<div class="page-title">2026 Training KPI Dashboard</div>', unsafe_allow_html=True)
st.markdown('<div class="page-sub">산업전환공동훈련센터 · 인하공업전문대학</div>', unsafe_allow_html=True)

_, btn_col = st.columns([6, 1])
with btn_col:
    if st.button("🔄 업데이트", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

st.divider()

# ── 데이터 로드 ──
with st.spinner("Google Drive 파일 분석 중..."):
    try:
        svc = get_drive_service()
        data = load_dashboard_data(svc)
    except Exception as e:
        st.error(f"Drive 연결 오류: {e}")
        st.stop()


sched = data.get("schedule") or {}
comp  = data.get("completion") or {}
yearly_df = data.get("yearly")

# 핵심 수치 (모두 pandas 연산 결과)
actual_2026  = comp.get("total_actual", 0)   # 2026수료현황 총계 행 합산
rate_2026    = round(actual_2026 / TARGET_2026 * 100, 1) if TARGET_2026 > 0 else 0
avg_sat      = comp.get("avg_sat", 0.0)
low_sat_df   = comp.get("low_sat", pd.DataFrame())
companies    = comp.get("companies", {})

# 2026년 훈련비/지원비 실적 — yearly_df에서 추출
actual_train_2026 = 0
actual_supp_2026  = 0
if yearly_df is not None and not yearly_df.empty:
    row_2026 = yearly_df[yearly_df['연도'] == 2026]
    if not row_2026.empty:
        actual_train_2026 = int(row_2026.iloc[0]['실적_훈련비'])
        actual_supp_2026  = int(row_2026.iloc[0]['실적_지원비'])
# actual_2026도 yearly_df 기준으로 재계산
actual_2026 = actual_train_2026 + actual_supp_2026 if (actual_train_2026 + actual_supp_2026) > 0 else actual_2026
rate_2026   = round(actual_2026 / TARGET_2026 * 100, 1) if TARGET_2026 > 0 else 0

# ── 탭 구성 ──
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📋 2026 핵심 KPI 및 일정",
    "🏭 협약기업 실적 및 만족도",
    "📈 연도별 성과 추이",
    "💰 인프라 예산 및 회계관리",
    "🔬 2025 훈련개발 분석"
])

# ============================================================
# Tab 1: 2026 핵심 KPI 및 일정
# ============================================================
with tab1:
    st.markdown('<div class="sec">종합 지표</div>', unsafe_allow_html=True)

    close_rate    = sched.get("close_rate", 0)
    closed_cnt    = sched.get("closed_courses", 0)
    total_cnt     = sched.get("total_courses", 0)
    closed_list   = sched.get("closed_list", pd.DataFrame())

    infra_kpi     = data.get("infra") or {}
    b_total       = infra_kpi.get("total_budget", 0)
    b_exec        = infra_kpi.get("total_exec", 0)
    b_rate        = infra_kpi.get("exec_rate", 0)
    b_remain      = b_total - b_exec

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        rate_badge = "green" if rate_2026 >= 90 else ("cyan" if rate_2026 >= 70 else "red")
        rate_text  = "순항 중" if rate_2026 >= 70 else "주의 필요"
        st.markdown(
            f'<div class="card cyan"><div class="card-icon">🎯</div>'
            f'<div class="card-label">목표 달성률 · 연간 목표 {TARGET_2026:,}명</div>'
            f'<div class="card-value">{rate_2026}%</div>'
            f'<div class="card-sub"><span class="badge {rate_badge}">{rate_text}</span></div></div>',
            unsafe_allow_html=True
        )
    with col2:
        st.markdown(
            f'<div class="card muted"><div class="card-icon">👥</div>'
            f'<div class="card-label">누적 수료인원 / 연간 목표</div>'
            f'<div class="card-value" style="font-size:26px">{actual_2026:,}'
            f'<span style="font-size:15px;color:#8A93A6"> / {TARGET_2026:,}명</span></div>'
            f'<div class="card-sub">'
            f'훈련비 <strong style="color:#22D3EE">{actual_train_2026:,}명</strong>'
            f'&nbsp;·&nbsp;지원비 <strong style="color:#C084FC">{actual_supp_2026:,}명</strong>'
            f'</div></div>',
            unsafe_allow_html=True
        )
    with col3:
        sc = "green" if avg_sat >= 4.5 else "amber"
        sc_text = "기준 충족" if avg_sat >= 4.5 else "개선 필요"
        st.markdown(
            f'<div class="card {sc}"><div class="card-icon">⭐</div>'
            f'<div class="card-label">평균 만족도</div>'
            f'<div class="card-value">{avg_sat}<span style="font-size:16px;color:#8A93A6"> / 5.0</span></div>'
            f'<div class="card-sub"><span class="badge {sc}">{sc_text}</span></div></div>',
            unsafe_allow_html=True
        )
    with col4:
        cc = "green" if close_rate < 10 else "amber"
        st.markdown(
            f'<div class="card {cc}"><div class="card-icon">📋</div>'
            f'<div class="card-label">과정 폐강률</div>'
            f'<div class="card-value">{close_rate}%</div>'
            f'<div class="card-sub">{closed_cnt}개 폐강 / 전체 {total_cnt}개</div></div>',
            unsafe_allow_html=True
        )

    st.markdown("<div style='margin-top:1rem'></div>", unsafe_allow_html=True)
    st.markdown('<div class="sec">인프라지원금 집행 현황</div>', unsafe_allow_html=True)
    b1, b2, b3 = st.columns(3)
    with b1:
        st.markdown(
            f'<div class="card muted"><div class="card-icon">💰</div>'
            f'<div class="card-label">총 예산액</div>'
            f'<div class="card-value" style="font-size:22px">{b_total:,}'
            f'<span style="font-size:13px;color:#8A93A6">원</span></div></div>',
            unsafe_allow_html=True
        )
    with b2:
        st.markdown(
            f'<div class="card muted"><div class="card-icon">📤</div>'
            f'<div class="card-label">총 집행액</div>'
            f'<div class="card-value" style="font-size:22px">{b_exec:,}'
            f'<span style="font-size:13px;color:#8A93A6">원</span></div>'
            f'<div class="card-sub">잔액 {b_remain:,}원</div></div>',
            unsafe_allow_html=True
        )
    with b3:
        bc = "green" if b_rate >= 50 else "amber"
        bc_color = "#34D399" if b_rate >= 50 else "#FBBF24"
        st.markdown(
            f'<div class="card {bc}"><div class="card-icon">📊</div>'
            f'<div class="card-label">집행률</div>'
            f'<div class="card-value">{b_rate}%</div>'
            f'{progress_bar(b_rate, bc_color)}'
            f'</div>',
            unsafe_allow_html=True
        )

    # 만족도 경고
    if not low_sat_df.empty:
        st.markdown('<div class="sec">⚠️ 만족도 경고 (4.5점 이하)</div>', unsafe_allow_html=True)
        for _, row in low_sat_df.iterrows():
            st.markdown(
                f'<div class="warn-card">⚠️ <strong>{row.get("훈련과정명","")}</strong> '
                f'{row.get("회차","")}회차 &nbsp;—&nbsp; '
                f'만족도 <strong>{row.get("만족도_n","")}점</strong> · 개선 필요</div>',
                unsafe_allow_html=True
            )

    # 폐강/미진행 과정 리스트업
    if not closed_list.empty:
        st.markdown('<div class="sec">📋 미진행(폐강) 과정 목록</div>', unsafe_allow_html=True)
        closed_disp = closed_list.rename(columns={'훈련과정명': '과정명', '회차': '회차', '목표인원_n': '목표인원'})
        closed_disp['목표인원'] = closed_disp['목표인원'].astype(int)
        dark_table(closed_disp, accent="#22D3EE")

    # 훈련일정 테이블
    st.markdown('<div class="sec">2026 훈련과정별 진행 현황</div>', unsafe_allow_html=True)
    rows_df = sched.get("rows", pd.DataFrame())
    if not rows_df.empty:
        display_df = rows_df.rename(columns={
            'no': 'No', '훈련유형': '유형', '훈련과정명': '과정명',
            '회차': '회차', '일시': '훈련 기간',
            '수료인원_n': '수료인원', '목표인원_n': '목표인원'
        })
        display_df['수료인원'] = display_df['수료인원'].astype(int)
        display_df['목표인원'] = display_df['목표인원'].astype(int)
        dark_table(display_df, accent="#22D3EE", height=400)

        plan_train = sched.get("plan_train", 880)
        actual_train = sched.get("actual", 0)
        st.caption(f"훈련비과정 기준 — 계획: {plan_train:,}명 | 현재 수료: {actual_train:,}명 | 잔여: {plan_train - actual_train:,}명")
    else:
        st.info("훈련일정 데이터를 불러오지 못했습니다.")

    if "error" in sched:
        st.error(f"파싱 오류: {sched['error']}")

# ============================================================
# Tab 2: 협약기업 실적 및 만족도
# ============================================================
with tab2:
    st.markdown('<div class="sec">협약기업별 수료 인원</div>', unsafe_allow_html=True)

    if companies:
        comp_df = pd.DataFrame([
            {"기업": k, "수료인원": v, "비율(%)": round(v / actual_2026 * 100, 1) if actual_2026 > 0 else 0}
            for k, v in companies.items() if v > 0
        ]).sort_values("수료인원", ascending=False)
        dark_table(comp_df, accent="#34D399")

        # Plotly 바차트
        fig = go.Figure(go.Bar(
            x=comp_df["기업"], y=comp_df["수료인원"],
            marker_color="#34D399", marker_line_width=0,
            text=comp_df["수료인원"], textposition="outside",
            textfont=dict(color="#E5E9F0")
        ))
        fig.update_layout(
            template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)", margin=dict(t=20, b=20, l=20, r=20),
            yaxis=dict(gridcolor="rgba(255,255,255,0.08)"),
            xaxis=dict(gridcolor="rgba(0,0,0,0)"),
            showlegend=False, height=300
        )
        st.plotly_chart(fig, use_container_width=True)

    st.markdown('<div class="sec">과정별 만족도 현황</div>', unsafe_allow_html=True)
    sat_df = comp.get("sat_table", pd.DataFrame())
    if not sat_df.empty:
        sat_df = sat_df.copy()
        sat_df['평가'] = sat_df['만족도_n'].apply(
            lambda x: "🟢 최우수" if x >= 4.8 else ("🟡 우수" if x > 4.5 else "🔴 개선필요")
        )
        sat_df = sat_df.rename(columns={
            '훈련과정명': '과정명', '회차': '회차',
            '만족도_n': '만족도', '수료인원_n': '수료인원'
        })
        sat_df['수료인원'] = pd.to_numeric(sat_df['수료인원'], errors='coerce').fillna(0).astype(int)
        dark_table(sat_df, accent="#C084FC")

    if not low_sat_df.empty:
        st.markdown('<div class="sec">🔴 만족도 4.5점 이하 과정 목록</div>', unsafe_allow_html=True)
        for _, row in low_sat_df.iterrows():
            st.markdown(
                f'<div class="warn-card">⚠️ <strong>{row.get("훈련과정명","")}</strong> '
                f'{row.get("회차","")}회차 &nbsp;—&nbsp; '
                f'만족도 <strong>{row.get("만족도_n","")}점</strong> '
                f'(수료인원: {int(row.get("수료인원_n", 0))}명)</div>',
                unsafe_allow_html=True
            )
    else:
        st.success("✅ 만족도 4.5점 이하 과정 없음")

    if "error" in comp:
        st.error(f"파싱 오류: {comp['error']}")

# ============================================================
# Tab 3: 연도별 성과 추이
# ============================================================
with tab3:
    st.markdown('<div class="sec">연도별 계획 대비 실적 추이</div>', unsafe_allow_html=True)

    if yearly_df is not None and not yearly_df.empty:
        display_yearly = yearly_df.copy()
        display_yearly['계획합계'] = display_yearly['계획_훈련비'] + display_yearly['계획_지원비']
        display_yearly['실적합계'] = display_yearly['실적_훈련비'] + display_yearly['실적_지원비']
        display_yearly['달성률(%)'] = display_yearly.apply(
            lambda r: round(r['실적합계'] / r['계획합계'] * 100, 1) if r['계획합계'] > 0 else 0, axis=1
        )
        display_yearly['비고'] = display_yearly['연도'].apply(
            lambda y: "▶ 진행중" if y == 2026 else ""
        )
        dark_table(display_yearly.rename(columns={
            '연도': '연도', '계획_훈련비': '계획(훈련비)', '계획_지원비': '계획(지원비)',
            '실적_훈련비': '실적(훈련비)', '실적_지원비': '실적(지원비)'
        }), accent="#C084FC")

        # Plotly 라인차트
        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=display_yearly['연도'], y=display_yearly['계획합계'],
            name="계획", mode="lines+markers+text",
            line=dict(color="#C084FC", width=2),
            marker=dict(size=7), text=display_yearly['계획합계'].apply(lambda v: f"{v:,}"),
            textposition="top center", textfont=dict(color="#C084FC", size=11),
            fill="tozeroy", fillcolor="rgba(192,132,252,0.08)"
        ))
        fig.add_trace(go.Scatter(
            x=display_yearly['연도'], y=display_yearly['실적합계'],
            name="실적", mode="lines+markers+text",
            line=dict(color="#22D3EE", width=2),
            marker=dict(size=7), text=display_yearly['실적합계'].apply(lambda v: f"{v:,}"),
            textposition="bottom center", textfont=dict(color="#22D3EE", size=11),
            fill="tozeroy", fillcolor="rgba(34,211,238,0.08)"
        ))
        fig.update_layout(
            template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)", height=380,
            margin=dict(t=20, b=20, l=20, r=20),
            yaxis=dict(gridcolor="rgba(255,255,255,0.08)"),
            xaxis=dict(gridcolor="rgba(0,0,0,0)", dtick=1),
            legend=dict(font=dict(color="#E5E9F0"))
        )
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("연도별 실적 데이터를 불러오지 못했습니다.")


# ============================================================
# Tab 4: 인프라 예산 및 회계관리
# ============================================================
with tab4:
    infra = data.get("infra") or {}

    if "error" in infra:
        st.error(f"회계 파일 파싱 오류: {infra['error']}")

    infra_b   = infra.get("infra_budget", 0)
    infra_e   = infra.get("infra_exec", 0)
    infra_r   = infra.get("infra_rate", 0)
    train_b   = infra.get("train_budget", 0)
    train_e   = infra.get("train_exec", 0)
    train_r   = infra.get("train_rate", 0)
    grand_b   = infra.get("grand_budget", 0)
    grand_e   = infra.get("grand_exec", 0)
    grand_r   = infra.get("grand_rate", 0)

    st.markdown('<div class="sec">2026 예산 총괄 (인프라지원금 + 훈련비)</div>', unsafe_allow_html=True)
    if grand_b > 0:
        g1, g2, g3 = st.columns(3)
        with g1:
            st.markdown(
                f'<div class="card muted"><div class="card-icon">💰</div>'
                f'<div class="card-label">총 예산 합계</div>'
                f'<div class="card-value" style="font-size:20px">{grand_b:,}<span style="font-size:13px;color:#8A93A6">원</span></div></div>',
                unsafe_allow_html=True)
        with g2:
            st.markdown(
                f'<div class="card muted"><div class="card-icon">📤</div>'
                f'<div class="card-label">총 집행 합계</div>'
                f'<div class="card-value" style="font-size:20px">{grand_e:,}<span style="font-size:13px;color:#8A93A6">원</span></div></div>',
                unsafe_allow_html=True)
        with g3:
            st.markdown(
                f'<div class="card green"><div class="card-icon">📊</div>'
                f'<div class="card-label">전체 집행률</div>'
                f'<div class="card-value">{grand_r}%</div>'
                f'<div class="card-sub">잔액 {grand_b-grand_e:,}원</div>'
                f'{progress_bar(grand_r, "#34D399")}'
                f'</div>',
                unsafe_allow_html=True)

    st.markdown('<div class="sec">① 인프라지원금 집행 현황</div>', unsafe_allow_html=True)
    if infra_b > 0:
        i1, i2, i3 = st.columns(3)
        with i1:
            st.markdown(f'<div class="card muted"><div class="card-label">예산액</div><div class="card-value" style="font-size:20px">{infra_b:,}<span style="font-size:13px;color:#8A93A6">원</span></div></div>', unsafe_allow_html=True)
        with i2:
            st.markdown(f'<div class="card muted"><div class="card-label">집행액</div><div class="card-value" style="font-size:20px">{infra_e:,}<span style="font-size:13px;color:#8A93A6">원</span></div></div>', unsafe_allow_html=True)
        with i3:
            st.markdown(
                f'<div class="card green"><div class="card-label">집행률</div>'
                f'<div class="card-value">{infra_r}%</div>'
                f'<div class="card-sub">잔액 {infra_b-infra_e:,}원</div>'
                f'{progress_bar(infra_r, "#34D399")}'
                f'</div>',
                unsafe_allow_html=True)
    else:
        st.info("인프라지원금 데이터를 불러오지 못했습니다.")

    st.markdown('<div class="sec">② 훈련비 집행 현황</div>', unsafe_allow_html=True)
    if train_b > 0:
        t1, t2, t3 = st.columns(3)
        with t1:
            st.markdown(f'<div class="card muted"><div class="card-label">예산액</div><div class="card-value" style="font-size:20px">{train_b:,}<span style="font-size:13px;color:#8A93A6">원</span></div></div>', unsafe_allow_html=True)
        with t2:
            st.markdown(f'<div class="card muted"><div class="card-label">집행액</div><div class="card-value" style="font-size:20px">{train_e:,}<span style="font-size:13px;color:#8A93A6">원</span></div></div>', unsafe_allow_html=True)
        with t3:
            st.markdown(
                f'<div class="card green"><div class="card-label">집행률</div>'
                f'<div class="card-value">{train_r}%</div>'
                f'<div class="card-sub">잔액 {train_b-train_e:,}원</div>'
                f'{progress_bar(train_r, "#34D399")}'
                f'</div>',
                unsafe_allow_html=True)
    else:
        st.info("훈련비 데이터를 불러오지 못했습니다.")

    # 월별 집행 Plotly 바차트
    monthly = infra.get("monthly", pd.DataFrame())
    if not monthly.empty:
        st.markdown('<div class="sec">월별 집행 현황</div>', unsafe_allow_html=True)
        fig = go.Figure(go.Bar(
            x=monthly["월"], y=monthly["집행액"],
            marker_color="#22D3EE", marker_line_width=0,
            text=monthly["집행액"].apply(lambda v: f"{int(v):,}"),
            textposition="outside", textfont=dict(color="#E5E9F0")
        ))
        fig.update_layout(
            template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)", height=300,
            margin=dict(t=20, b=20, l=20, r=20),
            yaxis=dict(gridcolor="rgba(255,255,255,0.08)"),
            xaxis=dict(gridcolor="rgba(0,0,0,0)")
        )
        st.plotly_chart(fig, use_container_width=True)

    # 항목별 집행 Plotly 바차트
    by_cat = infra.get("by_category", pd.DataFrame())
    if not by_cat.empty:
        st.markdown('<div class="sec">항목별(보조세목) 집행 현황</div>', unsafe_allow_html=True)
        fig2 = go.Figure(go.Bar(
            x=by_cat["보조세목명"], y=by_cat["집행액"],
            marker_color="#C084FC", marker_line_width=0,
            text=by_cat["집행액"].apply(lambda v: f"{int(v):,}"),
            textposition="outside", textfont=dict(color="#E5E9F0")
        ))
        fig2.update_layout(
            template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)", height=300,
            margin=dict(t=20, b=20, l=20, r=20),
            yaxis=dict(gridcolor="rgba(255,255,255,0.08)"),
            xaxis=dict(gridcolor="rgba(0,0,0,0)")
        )
        st.plotly_chart(fig2, use_container_width=True)

    # 세부 집행 내역 테이블
    detail = infra.get("detail", pd.DataFrame())
    if not detail.empty:
        st.markdown('<div class="sec">세부 집행 내역</div>', unsafe_allow_html=True)
        disp = detail[['집행일자', '보조세목명', '집행용도', '집행액']].copy()
        disp['집행일자'] = disp['집행일자'].dt.strftime('%Y-%m-%d')
        disp['집행액'] = disp['집행액'].astype(int)
        dark_table(disp, accent="#34D399", height=400)
        st.caption(f"총 {len(disp)}건 | 합계: {int(disp['집행액'].sum()):,}원")

# ── AI 상세 분석 ──
st.divider()
if st.button("🤖 AI 상세 분석 실행", type="secondary"):
    with st.spinner("Claude가 분석 중..."):
        try:
            yearly_summary = (
                yearly_df[['연도', '계획_훈련비', '계획_지원비', '실적_훈련비', '실적_지원비']]
                .to_dict(orient="records")
            ) if yearly_df is not None and not yearly_df.empty else []

            client = anthropic.Anthropic(api_key=st.secrets["anthropic"]["api_key"])
            msg = client.messages.create(
                model="claude-sonnet-4-5",
                max_tokens=1000,
                messages=[{"role": "user", "content": f"""
KPI 데이터:
- 2026 목표: {TARGET_2026:,}명 | 현재 수료: {actual_2026:,}명 | 달성률: {rate_2026}%
- 평균 만족도: {avg_sat}점 | 4.5점 이하 과정: {len(low_sat_df)}건
- 협약기업 수료: {json.dumps(companies, ensure_ascii=False)}
- 연도별 실적: {json.dumps(yearly_summary, ensure_ascii=False)}

4가지 항목 각 2~3문장으로 분석:
1. 🔴 핵심 위험 지표
2. 🟢 긍정 지표
3. ⚡ 즉시 조치 사항
4. 📋 하반기 전략 제안
"""}]
            )
            st.markdown(msg.content[0].text)
        except Exception as e:
            st.error(f"AI 분석 오류: {e}")

st.caption(f"마지막 업데이트: {datetime.now().strftime('%Y-%m-%d %H:%M')}")

# ============================================================
# Tab 5: 2025 훈련개발 분석
# ============================================================
with tab5:
    st.markdown('<div class="sec">2025 성과 분석 대시보드</div>', unsafe_allow_html=True)

    # ── 고정 데이터 (PDF/PPT 기준) ──
    sat_trend   = {"연도": [2022, 2023, 2024, 2025], "만족도": [4.09, 4.37, 4.42, 4.62]}
    comp_rate   = {"연도": [2022, 2023, 2024, 2025], "참여율": [22.0, 9.4, 6.5, 77.8]}
    grad_trend  = {"연도": [2022, 2023, 2024, 2025], "수료인원": [321, 887, 1089, 1699]}
    aidx_trend  = {"연도": [2022, 2023, 2024, 2025], "AI·DX수료": [39, 181, 297, 630]}
    train_hours = {"연도": [2022, 2023, 2024, 2025], "훈련시간": [320, 302, 370, 653]}

    # ── 상단 KPI 3개 카드 ──
    k1, k2, k3 = st.columns(3)
    with k1:
        st.markdown(
            '<div class="card cyan"><div class="card-icon">🔍</div>'
            '<div class="card-label">수요조사 · AI·DX 수요 (11개 협약기업)</div>'
            '<div class="card-value">61%</div>'
            '<div class="card-sub"><span class="badge cyan">AI·DX 최우선 수요 확인</span></div></div>',
            unsafe_allow_html=True
        )
    with k2:
        st.markdown(
            '<div class="card purple"><div class="card-icon">⭐</div>'
            '<div class="card-label">종합 만족도 · 전년 대비 +0.2점</div>'
            '<div class="card-value">4.62점</div>'
            '<div class="card-sub"><span class="badge green">+13% ↑ (3개년 최고점)</span></div></div>',
            unsafe_allow_html=True
        )
    with k3:
        st.markdown(
            '<div class="card green"><div class="card-icon">🏢</div>'
            '<div class="card-label">조직진단 · 전환단계 (18개 기업 전수 진단)</div>'
            '<div class="card-value">44.4%</div>'
            '<div class="card-sub"><span class="badge green">선제대응 22.2% · 정착 33.3%</span></div></div>',
            unsafe_allow_html=True
        )

    st.markdown("<div style='margin-top:1.5rem'></div>", unsafe_allow_html=True)

    # ── 만족도 추이 + 수료인원 추이 ──
    row1_left, row1_right = st.columns(2)

    with row1_left:
        st.markdown('<div class="sec">만족도 성장 추이 (2022–2025)</div>', unsafe_allow_html=True)
        fig_sat = go.Figure()
        fig_sat.add_trace(go.Scatter(
            x=sat_trend["연도"], y=sat_trend["만족도"],
            mode="lines+markers+text",
            line=dict(color="#C084FC", width=3),
            marker=dict(size=8, color="#C084FC"),
            fill="tozeroy", fillcolor="rgba(192,132,252,0.10)",
            text=[str(v) for v in sat_trend["만족도"]],
            textposition="top center",
            textfont=dict(color="#C084FC", size=12)
        ))
        fig_sat.update_layout(
            template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)", height=260,
            margin=dict(t=10, b=30, l=20, r=20),
            yaxis=dict(gridcolor="rgba(255,255,255,0.08)", range=[3.8, 5.0]),
            xaxis=dict(gridcolor="rgba(0,0,0,0)", dtick=1),
            showlegend=False,
            annotations=[dict(
                x=2025, y=4.62, text="<b>최고점 경신</b>",
                showarrow=True, arrowhead=2, arrowcolor="#C084FC",
                font=dict(color="#C084FC", size=11), ax=30, ay=-30
            )]
        )
        st.plotly_chart(fig_sat, use_container_width=True)

    with row1_right:
        st.markdown('<div class="sec">수료 인원 추이 (2022–2025)</div>', unsafe_allow_html=True)
        colors = ["rgba(34,211,238,0.4)", "rgba(34,211,238,0.6)",
                  "rgba(34,211,238,0.8)", "#22D3EE"]
        fig_grad = go.Figure(go.Bar(
            x=[str(y) for y in grad_trend["연도"]],
            y=grad_trend["수료인원"],
            marker_color=colors, marker_line_width=0,
            text=grad_trend["수료인원"],
            textposition="outside",
            textfont=dict(color="#E5E9F0", size=12)
        ))
        fig_grad.update_layout(
            template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)", height=260,
            margin=dict(t=10, b=30, l=20, r=20),
            yaxis=dict(gridcolor="rgba(255,255,255,0.08)"),
            xaxis=dict(gridcolor="rgba(0,0,0,0)"),
            showlegend=False
        )
        st.plotly_chart(fig_grad, use_container_width=True)

    # ── 협약기업 수요 분포 + 조직진단 3단계 ──
    row2_left, row2_right = st.columns(2)

    with row2_left:
        st.markdown('<div class="sec">협약기업 수요 분포</div>', unsafe_allow_html=True)
        demand_labels = ["AI·DX (61%)\n11개사", "EV전환 (56%)\n10개사", "기타"]
        demand_vals   = [61, 56, 20]
        demand_colors = ["#22D3EE", "#34D399", "rgba(138,147,166,0.4)"]
        fig_demand = go.Figure(go.Bar(
            x=demand_labels, y=demand_vals,
            marker_color=demand_colors, marker_line_width=0,
            text=[f"{v}%" for v in demand_vals],
            textposition="outside",
            textfont=dict(color="#E5E9F0", size=13)
        ))
        fig_demand.update_layout(
            template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)", height=260,
            margin=dict(t=10, b=30, l=20, r=20),
            yaxis=dict(gridcolor="rgba(255,255,255,0.08)", range=[0, 85]),
            xaxis=dict(gridcolor="rgba(0,0,0,0)"),
            showlegend=False
        )
        st.plotly_chart(fig_demand, use_container_width=True)

    with row2_right:
        st.markdown('<div class="sec">조직진단 3단계 분포 (18개 기업)</div>', unsafe_allow_html=True)
        stage_labels = ["선제대응\n(4개사)", "전환단계\n(8개사)", "정착단계\n(6개사)"]
        stage_vals   = [22.2, 44.4, 33.3]
        stage_colors = ["rgba(251,191,36,0.7)", "#22D3EE", "#34D399"]
        fig_stage = go.Figure(go.Bar(
            x=stage_labels, y=stage_vals,
            marker_color=stage_colors, marker_line_width=0,
            text=[f"{v}%" for v in stage_vals],
            textposition="outside",
            textfont=dict(color="#E5E9F0", size=13)
        ))
        fig_stage.update_layout(
            template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)", height=260,
            margin=dict(t=10, b=30, l=20, r=20),
            yaxis=dict(gridcolor="rgba(255,255,255,0.08)", range=[0, 65]),
            xaxis=dict(gridcolor="rgba(0,0,0,0)"),
            showlegend=False
        )
        st.plotly_chart(fig_stage, use_container_width=True)

    # ── AI·DX 수료 추이 + 훈련시간 추이 ──
    row3_left, row3_right = st.columns(2)

    with row3_left:
        st.markdown('<div class="sec">AI·DX 훈련과정 수료인원 추이</div>', unsafe_allow_html=True)
        fig_aidx = go.Figure()
        fig_aidx.add_trace(go.Scatter(
            x=aidx_trend["연도"], y=aidx_trend["AI·DX수료"],
            mode="lines+markers+text",
            line=dict(color="#34D399", width=3),
            marker=dict(size=8, color="#34D399"),
            fill="tozeroy", fillcolor="rgba(52,211,153,0.10)",
            text=[str(v) for v in aidx_trend["AI·DX수료"]],
            textposition="top center",
            textfont=dict(color="#34D399", size=12)
        ))
        fig_aidx.update_layout(
            template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)", height=240,
            margin=dict(t=10, b=30, l=20, r=20),
            yaxis=dict(gridcolor="rgba(255,255,255,0.08)"),
            xaxis=dict(gridcolor="rgba(0,0,0,0)", dtick=1),
            showlegend=False
        )
        st.plotly_chart(fig_aidx, use_container_width=True)

    with row3_right:
        st.markdown('<div class="sec">전문직무 훈련시간 추이</div>', unsafe_allow_html=True)
        fig_hours = go.Figure()
        fig_hours.add_trace(go.Scatter(
            x=train_hours["연도"], y=train_hours["훈련시간"],
            mode="lines+markers+text",
            line=dict(color="#FBBF24", width=3),
            marker=dict(size=8, color="#FBBF24"),
            fill="tozeroy", fillcolor="rgba(251,191,36,0.08)",
            text=[f"{v}h" for v in train_hours["훈련시간"]],
            textposition="top center",
            textfont=dict(color="#FBBF24", size=12)
        ))
        fig_hours.update_layout(
            template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)", height=240,
            margin=dict(t=10, b=30, l=20, r=20),
            yaxis=dict(gridcolor="rgba(255,255,255,0.08)"),
            xaxis=dict(gridcolor="rgba(0,0,0,0)", dtick=1),
            showlegend=False
        )
        st.plotly_chart(fig_hours, use_container_width=True)

    # ── 협약기업 참여율 반등 ──
    st.markdown('<div class="sec">협약기업 참여율 추이</div>', unsafe_allow_html=True)
    part_colors = ["rgba(239,68,68,0.6)", "rgba(239,68,68,0.5)",
                   "rgba(239,68,68,0.4)", "#22D3EE"]
    fig_part = go.Figure()
    fig_part.add_trace(go.Scatter(
        x=comp_rate["연도"], y=comp_rate["참여율"],
        mode="lines+markers+text",
        line=dict(color="#22D3EE", width=3),
        marker=dict(size=10, color=part_colors),
        text=[f"{v}%" for v in comp_rate["참여율"]],
        textposition=["bottom center", "bottom center", "bottom center", "top center"],
        textfont=dict(color="#E5E9F0", size=13)
    ))
    fig_part.update_layout(
        template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)", height=220,
        margin=dict(t=10, b=30, l=20, r=20),
        yaxis=dict(gridcolor="rgba(255,255,255,0.08)", range=[0, 95]),
        xaxis=dict(gridcolor="rgba(0,0,0,0)", dtick=1),
        showlegend=False,
        annotations=[dict(
            x=2025, y=77.8,
            text="<b>6.5% → 77.8% 반등 🚀</b>",
            showarrow=True, arrowhead=2, arrowcolor="#22D3EE",
            font=dict(color="#22D3EE", size=12), ax=-80, ay=-30
        )]
    )
    st.plotly_chart(fig_part, use_container_width=True)

    # ── On-Demand 실증성과 요약 테이블 ──
    st.markdown('<div class="sec">On-Demand 실증 성과 (수요 → 대응 → 결과)</div>', unsafe_allow_html=True)
    ondemand_data = pd.DataFrame([
        {"협약기업": "르노코리아", "수요 배경": "글로벌 고전압 안전교육 / 국내 적합기관 부재",
         "맞춤 대응": "Orange Training 3단계 인증 구축",
         "주요 성과": "글로벌 최초 외부교육기관 선정 · 131명 수료 · 진단 94→98점"},
        {"협약기업": "GMTCK", "수요 배경": "AI·DX 전환 가속 / 협력사 기술격차 해소",
         "맞춤 대응": "AI Week 전사 행사 연계 세미나 공동 기획",
         "주요 성과": "협력사 참여 21→55개사(+162%) · 만족도 8.53/10"},
        {"협약기업": "불스원·ESG", "수요 배경": "ESG 제품 검증 필요 / 소상공인 교육 사각지대",
         "맞춤 대응": "매연 27.3% 저감 실증 · 주말 방문교육",
         "주요 성과": "신제품 매출 17억원 · 8만개 판매 · 소상공인 44명 해소"},
    ])
    dark_table(ondemand_data, accent="#34D399")
st.caption(f"마지막 업데이트: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
