import re
from io import BytesIO

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="後方数値データ分析", layout="wide")
st.title("📊 後方数値データ分析ダッシュボード")

st.markdown(
    """
    <style>
    section[data-testid="stSidebar"] div[data-testid="stVerticalBlock"] {
        gap: 0.55rem;
    }
    section[data-testid="stSidebar"] label p {
        font-size: 0.88rem;
        font-weight: 600;
    }
    div[data-baseweb="select"] > div {
        min-height: 42px;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 0.4rem;
        flex-wrap: wrap;
    }
    .stTabs [data-baseweb="tab"] {
        height: 42px;
        padding-left: 0.8rem;
        padding-right: 0.8rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

CORE_METRICS = [
    "申込件数", "承認件数", "成約件数", "承認率", "成約率",
    "コスト", "申込CPA", "承認CPA",
    "取扱金額_当月", "取扱金額_翌月", "取扱金額_翌々月", "投下倍率",
]
RATE_METRICS = {"承認率", "成約率", "申込CPA", "承認CPA", "投下倍率"}

# Excelのふりがな情報などが見出しに連結された場合も吸収する
COLUMN_ALIASES = {
    "媒体名": ["媒体名", "媒体名バイタイメイ"],
    "媒体": ["媒体", "媒体バイタイ"],
    "成約フラグ": ["成約フラグ", "成約フラグセイヤク"],
    "単価": ["単価", "単価タンカ"],
    "借入フラグ": ["借入フラグ", "借入フラグカリイレ"],
    "大項目": ["大項目", "大項目ダイコウモク"],
    "中項目": ["中項目", "中項目チュウコウモク"],
    "小項目": ["小項目", "小項目ショウコウモク"],
    "小項目2": ["小項目2", "小項目2ショウコウモク"],
    "代理店コード": ["代理店コード", "代理店コードダイリテン"],
    "配信日": ["配信日", "配信日ハイシンビ"],
}


def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).replace("\n", " ").strip() for c in df.columns]
    rename = {}
    for canonical, candidates in COLUMN_ALIASES.items():
        if canonical in df.columns:
            continue
        for col in df.columns:
            normalized = re.sub(r"\s+", "", str(col))
            if normalized in candidates:
                rename[col] = canonical
                break
    return df.rename(columns=rename)


def make_unique_columns(df: pd.DataFrame) -> pd.DataFrame:
    seen = {}
    cols = []
    for col in df.columns:
        col = str(col)
        seen[col] = seen.get(col, 0) + 1
        cols.append(col if seen[col] == 1 else f"{col}_{seen[col]}")
    df = df.copy()
    df.columns = cols
    return df


@st.cache_data(show_spinner=False)
def get_sheet_names(file_bytes: bytes) -> list[str]:
    return pd.ExcelFile(BytesIO(file_bytes), engine="openpyxl").sheet_names


@st.cache_data(show_spinner="Excelを読み込んでいます…")
def read_sheet(file_bytes: bytes, sheet_name: str) -> pd.DataFrame:
    return make_unique_columns(clean_columns(pd.read_excel(
        BytesIO(file_bytes), sheet_name=sheet_name, engine="openpyxl"
    )))


def safe_div(numerator, denominator):
    if isinstance(denominator, pd.Series):
        return numerator.div(denominator.replace(0, np.nan))
    return np.nan if denominator in (0, None) or pd.isna(denominator) else numerator / denominator


def normalize_month(value) -> str | None:
    if pd.isna(value):
        return None
    text = str(value).strip()
    dt = pd.to_datetime(text, errors="coerce")
    if pd.isna(dt):
        m = re.search(r"(20\d{2})\D+(\d{1,2})", text)
        return f"{m.group(1)}-{int(m.group(2)):02d}" if m else None
    return dt.strftime("%Y-%m")


def prepare_cost_lookup(cost_df: pd.DataFrame) -> pd.DataFrame:
    cost_df = clean_columns(cost_df)
    if cost_df.empty or len(cost_df.columns) < 5:
        return pd.DataFrame(columns=["小項目2", "申込月", "補完元コスト"])

    key_col = "項目4" if "項目4" in cost_df.columns else cost_df.columns[3]
    month_cols = [c for c in cost_df.columns if normalize_month(c)]
    if not month_cols:
        return pd.DataFrame(columns=["小項目2", "申込月", "補完元コスト"])

    long_df = cost_df[[key_col] + month_cols].melt(
        id_vars=[key_col], var_name="月", value_name="補完元コスト"
    )
    long_df = long_df.rename(columns={key_col: "小項目2"})
    long_df["小項目2"] = long_df["小項目2"].astype(str).str.strip()
    long_df["申込月"] = long_df["月"].map(normalize_month)
    long_df["補完元コスト"] = pd.to_numeric(long_df["補完元コスト"], errors="coerce")
    return (
        long_df.dropna(subset=["申込月", "補完元コスト"])
        .groupby(["小項目2", "申込月"], as_index=False)["補完元コスト"].sum()
    )


def preprocess(raw: pd.DataFrame, master: pd.DataFrame, cost_df: pd.DataFrame | None) -> pd.DataFrame:
    df = clean_columns(raw)
    master = clean_columns(master)

    required = ["申込日", "媒体コード"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"後方数値データに必須列がありません: {', '.join(missing)}")
    if "媒体コード" not in master.columns:
        raise ValueError("媒体コードマスタに『媒体コード』列がありません。")

    for col in ["申込日", "承認日", "成約日"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    df["申込月"] = df["申込日"].dt.strftime("%Y-%m")

    # 後方データに既に付与済みの列があっても、マスタの最新分類を優先
    master_cols = [
        "媒体コード", "テンプレコード", "メニュー名", "大項目", "中項目",
        "小項目", "小項目2", "代理店コード", "配信日", "ポイント",
        "キャンペーン識別", "メニューコード", "コストフラグ",
    ]
    master_cols = [c for c in master_cols if c in master.columns]
    df["媒体コード"] = df["媒体コード"].astype(str).str.strip()
    master["媒体コード"] = master["媒体コード"].astype(str).str.strip()
    master_join = master[master_cols].drop_duplicates(subset=["媒体コード"], keep="last")
    overlap = [c for c in master_cols if c != "媒体コード" and c in df.columns]
    df = df.drop(columns=overlap, errors="ignore").merge(master_join, on="媒体コード", how="left")

    df["申込フラグ"] = 1
    df["承認フラグ"] = df.get("承認日", pd.Series(index=df.index, dtype="datetime64[ns]")).notna().astype(int)
    if "成約フラグ" in df.columns:
        existing = pd.to_numeric(df["成約フラグ"], errors="coerce")
        date_flag = df.get("成約日", pd.Series(index=df.index, dtype="datetime64[ns]")).notna().astype(int)
        df["成約フラグ"] = existing.fillna(date_flag).astype(int)
    else:
        df["成約フラグ"] = df.get("成約日", pd.Series(index=df.index, dtype="datetime64[ns]")).notna().astype(int)

    numeric_candidates = [
        "取扱金額_申込当月", "取扱金額_申込翌月末", "取扱金額_申込翌々月末",
        "コスト", "単価", "年齢", "年収", "同借希望額", "子供数",
        "住宅ローン返済月額", "他社借入件数", "他社借入残高", "承認枠", "成約枠",
    ]
    for col in numeric_candidates:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # AK列コストがエラー/空欄の行だけ、コストデータの媒体×月総額を申込行数で按分
    df["コスト_元データ"] = pd.to_numeric(df.get("コスト", np.nan), errors="coerce")
    df["コスト補完フラグ"] = 0
    if cost_df is not None and not cost_df.empty and "小項目2" in df.columns:
        lookup = prepare_cost_lookup(cost_df)
        if not lookup.empty:
            df["小項目2"] = df["小項目2"].astype(str).str.strip()
            counts = df.groupby(["小項目2", "申込月"], dropna=False).size().rename("媒体月申込件数").reset_index()
            df = df.merge(counts, on=["小項目2", "申込月"], how="left")
            df = df.merge(lookup, on=["小項目2", "申込月"], how="left")
            fallback = safe_div(df["補完元コスト"], df["媒体月申込件数"])
            invalid = df["コスト_元データ"].isna()
            df["コスト"] = df["コスト_元データ"].where(~invalid, fallback)
            df.loc[invalid & df["コスト"].notna(), "コスト補完フラグ"] = 1
        else:
            df["コスト"] = df["コスト_元データ"]
    else:
        df["コスト"] = df["コスト_元データ"]

    if "年齢" in df.columns:
        df["年齢グループ"] = pd.cut(
            df["年齢"], [-np.inf, 29, 39, 49, 59, np.inf],
            labels=["20代以下", "30代", "40代", "50代", "60代以上"]
        )
    if "年収" in df.columns:
        df["年収グループ"] = pd.cut(
            df["年収"], [-np.inf, 299, 499, 699, 999, np.inf],
            labels=["300万未満", "300-500万", "500-700万", "700-1000万", "1000万以上"]
        )
    return df


def aggregate(df: pd.DataFrame, group_cols: list[str]) -> pd.DataFrame:
    group_cols = list(dict.fromkeys([c for c in group_cols if c in df.columns]))
    work = df.copy()
    for col in ["取扱金額_申込当月", "取扱金額_申込翌月末", "取扱金額_申込翌々月末", "コスト"]:
        if col not in work.columns:
            work[col] = 0.0

    agg_map = {
        "申込件数": ("申込フラグ", "sum"),
        "承認件数": ("承認フラグ", "sum"),
        "成約件数": ("成約フラグ", "sum"),
        "コスト": ("コスト", "sum"),
        "取扱金額_当月": ("取扱金額_申込当月", "sum"),
        "取扱金額_翌月": ("取扱金額_申込翌月末", "sum"),
        "取扱金額_翌々月": ("取扱金額_申込翌々月末", "sum"),
    }
    result = work.groupby(group_cols, dropna=False, observed=True).agg(**agg_map).reset_index() if group_cols else pd.DataFrame({
        name: [work[src].agg(func)] for name, (src, func) in agg_map.items()
    })
    result["承認率"] = safe_div(result["承認件数"], result["申込件数"])
    result["成約率"] = safe_div(result["成約件数"], result["申込件数"])
    result["申込CPA"] = safe_div(result["コスト"], result["申込件数"])
    result["承認CPA"] = safe_div(result["コスト"], result["承認件数"])
    result["投下倍率"] = safe_div(result["取扱金額_翌月"], result["コスト"])
    return result


def add_extra_numeric_sums(df: pd.DataFrame, result: pd.DataFrame, group_cols: list[str], selected: list[str]) -> pd.DataFrame:
    # 集計軸が数値型として読み込まれていても、表示項目側へ重複追加しない。
    # 例：空欄が多い「大項目」がfloat型になった場合、reset_index時に列名衝突していた。
    group_cols = list(dict.fromkeys(c for c in group_cols if c in df.columns))
    selected = list(dict.fromkeys(selected))
    extra = [
        c for c in selected
        if c not in CORE_METRICS
        and c not in group_cols
        and c not in result.columns
        and c in df.columns
        and pd.api.types.is_numeric_dtype(df[c])
    ]
    if not extra:
        return result
    if group_cols:
        sums = (
            df.groupby(group_cols, dropna=False, observed=True)[extra]
            .sum(min_count=1)
            .reset_index()
        )
        return result.merge(sums, on=group_cols, how="left", validate="one_to_one")
    for col in extra:
        result[col] = df[col].sum(min_count=1)
    return result


def format_table(df: pd.DataFrame):
    formats = {c: "{:.1%}" for c in ["承認率", "成約率"] if c in df.columns}
    formats.update({c: "{:,.0f}" for c in ["コスト", "申込CPA", "承認CPA", "取扱金額_当月", "取扱金額_翌月", "取扱金額_翌々月"] if c in df.columns})
    formats.update({c: "{:.2f}" for c in ["投下倍率"] if c in df.columns})
    return df.style.format(formats, na_rep="-")


def to_excel(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="分析結果")
    return output.getvalue()


def metric_options(df: pd.DataFrame) -> list[str]:
    # 分類・識別系の列は、Excel上で全空欄だと数値型(float)判定されることがある。
    # それらを表示指標に混ぜず、実際の数値項目だけを候補にする。
    excluded = {
        "申込フラグ", "承認フラグ", "成約フラグ", "コスト_元データ",
        "コスト補完フラグ", "媒体月申込件数", "補完元コスト",
        "媒体コード", "テンプレコード", "メニュー名", "大項目", "中項目",
        "小項目", "小項目2", "代理店コード", "配信日", "ポイント",
        "キャンペーン識別", "メニューコード", "コストフラグ", "媒体名", "媒体",
        "申込月",
    }
    extras = [
        c for c in df.select_dtypes(include="number").columns
        if c not in excluded and c not in CORE_METRICS
    ]
    return list(dict.fromkeys(CORE_METRICS + extras))


def selection_filter(df: pd.DataFrame, column: str, label: str, key: str) -> pd.DataFrame:
    """候補が存在する分類列だけをサイドバーへ表示する。"""
    if column not in df.columns:
        return df

    values = df[column].dropna().astype(str).str.strip()
    choices = sorted(v for v in values.unique().tolist() if v and v.casefold() != "nan")
    if not choices:
        st.sidebar.caption(f"{label}: 現在選択可能な値なし")
        return df

    selected = st.sidebar.multiselect(
        label,
        choices,
        key=key,
        placeholder="すべて",
    )
    return df[df[column].astype(str).isin(selected)] if selected else df


def render_kpis(df: pd.DataFrame):
    kpi = aggregate(df, [])
    row = kpi.iloc[0]

    st.markdown(
        """
        <style>
        div[data-testid="stMetric"] {
            background: #ffffff;
            border: 1px solid rgba(49, 51, 63, 0.18);
            border-radius: 12px;
            padding: 14px 16px;
            min-height: 112px;
            box-shadow: 0 2px 7px rgba(0, 0, 0, 0.05);
        }
        div[data-testid="stMetricLabel"] p {
            font-size: 0.95rem !important;
            font-weight: 700 !important;
        }
        div[data-testid="stMetricValue"] {
            font-size: clamp(1.45rem, 2.1vw, 2.15rem) !important;
            line-height: 1.15 !important;
            overflow: visible !important;
            white-space: nowrap !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    # 7枚を一列に詰め込むと金額が省略されるため、4枚＋3枚の二段表示にする。
    first = st.columns(4)
    first[0].metric("申込件数", f"{row['申込件数']:,.0f}")
    first[1].metric("承認件数", f"{row['承認件数']:,.0f}")
    first[2].metric("承認率", f"{row['承認率']:.1%}" if pd.notna(row["承認率"]) else "-")
    first[3].metric("コスト", f"¥{row['コスト']:,.0f}" if pd.notna(row["コスト"]) else "-")

    second = st.columns(3)
    second[0].metric("申込CPA", f"¥{row['申込CPA']:,.0f}" if pd.notna(row["申込CPA"]) else "-")
    second[1].metric("承認CPA", f"¥{row['承認CPA']:,.0f}" if pd.notna(row["承認CPA"]) else "-")
    second[2].metric("投下倍率", f"{row['投下倍率']:.2f}" if pd.notna(row["投下倍率"]) else "-")


def allocation_simulator(df: pd.DataFrame):
    st.subheader("💸 媒体アロケーション影響試算")
    st.caption("過去実績の媒体別効率が移管後も続く前提で、元媒体から減る成果と先媒体で増える成果の差分を試算します。")
    if not {"大項目", "小項目", "コスト"}.issubset(df.columns):
        st.info("大項目・小項目・コストが揃うと利用できます。")
        return

    def valid_options(series: pd.Series) -> list[str]:
        values = series.dropna().astype(str).str.strip()
        return sorted(v for v in values.unique().tolist() if v and v.casefold() != "nan")

    bigs = valid_options(df["大項目"])
    if not bigs:
        st.info("現在のフィルタ条件では、選択可能な大項目がありません。")
        return

    # 依存する選択肢が変わった際、古いsession_state値が残って選択不能になるのを防ぐ。
    if st.session_state.get("alloc_from_big") not in bigs:
        st.session_state["alloc_from_big"] = bigs[0]
    if st.session_state.get("alloc_to_big") not in bigs:
        st.session_state["alloc_to_big"] = bigs[1] if len(bigs) > 1 else bigs[0]

    c1, c2, c3 = st.columns(3)
    from_big = c1.selectbox("アロケーション元：大項目", bigs, key="alloc_from_big")
    from_smalls = valid_options(df.loc[df["大項目"].astype(str).str.strip() == from_big, "小項目"])
    if not from_smalls:
        c2.warning("元媒体の小項目がありません。")
        return
    if st.session_state.get("alloc_from_small") not in from_smalls:
        st.session_state["alloc_from_small"] = from_smalls[0]
    from_small = c2.selectbox("アロケーション元：小項目", from_smalls, key="alloc_from_small")
    amount = c3.number_input(
        "移管金額（円）", min_value=0.0, value=500000.0,
        step=10000.0, format="%.0f", key="alloc_amount"
    )

    d1, d2 = st.columns(2)
    to_big = d1.selectbox("アロケーション先：大項目", bigs, key="alloc_to_big")
    to_smalls = valid_options(df.loc[df["大項目"].astype(str).str.strip() == to_big, "小項目"])
    if not to_smalls:
        d2.warning("移管先の小項目がありません。")
        return
    preferred_to = next((v for v in to_smalls if not (to_big == from_big and v == from_small)), to_smalls[0])
    if st.session_state.get("alloc_to_small") not in to_smalls:
        st.session_state["alloc_to_small"] = preferred_to
    to_small = d2.selectbox("アロケーション先：小項目", to_smalls, key="alloc_to_small")

    same_media = from_big == to_big and from_small == to_small
    if same_media:
        st.warning("アロケーション元と先が同じです。異なる媒体を選択してください。")

    if st.button("影響を試算", type="primary", disabled=amount <= 0 or same_media):
        source = df[
            (df["大項目"].astype(str).str.strip() == from_big)
            & (df["小項目"].astype(str).str.strip() == from_small)
        ]
        dest = df[
            (df["大項目"].astype(str).str.strip() == to_big)
            & (df["小項目"].astype(str).str.strip() == to_small)
        ]
        if source.empty or dest.empty:
            st.error("元または先媒体の対象データがありません。")
            return

        s = aggregate(source, []).iloc[0]
        d = aggregate(dest, []).iloc[0]
        if s["コスト"] <= 0 or d["コスト"] <= 0:
            st.error("元または先媒体の実績コストが0のため試算できません。")
            return

        def per_cost(row, metric):
            return row[metric] / row["コスト"] if row["コスト"] else np.nan

        before = aggregate(pd.concat([source, dest]), []).iloc[0]
        source_cost_after = max(s["コスト"] - amount, 0)
        effective_amount = min(amount, s["コスト"])
        dest_cost_after = d["コスト"] + effective_amount

        projected = {}
        for metric in ["申込件数", "承認件数", "成約件数", "取扱金額_翌月"]:
            projected[metric] = per_cost(s, metric) * source_cost_after + per_cost(d, metric) * dest_cost_after
        projected["コスト"] = source_cost_after + dest_cost_after
        projected["承認率"] = safe_div(projected["承認件数"], projected["申込件数"])
        projected["申込CPA"] = safe_div(projected["コスト"], projected["申込件数"])
        projected["承認CPA"] = safe_div(projected["コスト"], projected["承認件数"])
        projected["投下倍率"] = safe_div(projected["取扱金額_翌月"], projected["コスト"])

        rows = []
        for metric in ["投下倍率", "取扱金額_翌月", "申込CPA", "承認CPA", "申込件数", "承認率"]:
            b = before[metric]
            a = projected[metric]
            rows.append({"指標": metric, "移管前": b, "移管後試算": a, "差分": a - b, "変化率": safe_div(a - b, b)})
        impact = pd.DataFrame(rows)
        st.dataframe(format_table(impact), use_container_width=True, hide_index=True)
        if effective_amount < amount:
            st.warning(f"元媒体の実績コストを超えるため、試算上は {effective_amount:,.0f}円で計算しました。")


uploaded_file = st.file_uploader("後方数値データ・媒体コードマスタ・コストデータを含むExcel", type=["xlsx"])

if uploaded_file is None:
    st.info("Excelファイルをアップロードすると、シート選択が表示されます。")
    st.stop()

file_bytes = uploaded_file.getvalue()
sheets = get_sheet_names(file_bytes)

st.sidebar.header("📄 読み込み設定")
def_index = next((i for i, s in enumerate(sheets) if "後方" in s), 0)
master_index = next((i for i, s in enumerate(sheets) if "媒体コード" in s or "マスタ" in s), min(1, len(sheets)-1))
cost_index = next((i for i, s in enumerate(sheets) if "コスト" in s), 0)

data_sheet = st.sidebar.selectbox("後方数値データのシート", sheets, index=def_index)
master_sheet = st.sidebar.selectbox("媒体コードマスタのシート", sheets, index=master_index)
use_cost_sheet = st.sidebar.checkbox("コストデータシートでエラーを補完", value=True)
cost_sheet = st.sidebar.selectbox("コストデータのシート", sheets, index=cost_index, disabled=not use_cost_sheet)

try:
    raw_df = read_sheet(file_bytes, data_sheet)
    master_df = read_sheet(file_bytes, master_sheet)
    cost_df = read_sheet(file_bytes, cost_sheet) if use_cost_sheet else None
    base_df = preprocess(raw_df, master_df, cost_df)
except Exception as exc:
    st.error(f"読み込み・前処理でエラーが発生しました: {exc}")
    st.stop()

st.sidebar.header("🔍 全タブ共通フィルタ")
df = base_df.copy()
valid_dates = df["申込日"].dropna()
if not valid_dates.empty:
    date_range = st.sidebar.date_input("申込日（期間）", [valid_dates.min().date(), valid_dates.max().date()])
    if len(date_range) == 2:
        df = df[(df["申込日"] >= pd.Timestamp(date_range[0])) & (df["申込日"] < pd.Timestamp(date_range[1]) + pd.Timedelta(days=1))]

df = selection_filter(df, "大項目", "大項目", "filter_big")
df = selection_filter(df, "中項目", "中項目", "filter_middle")
df = selection_filter(df, "小項目", "小項目", "filter_small")
st.sidebar.caption(f"対象件数: {len(df):,} / {len(base_df):,}")
if "コスト補完フラグ" in df.columns:
    st.sidebar.caption(f"コスト補完行: {int(df['コスト補完フラグ'].sum()):,}")

if df.empty:
    st.warning("フィルタ条件に該当するデータがありません。")
    st.stop()

render_kpis(df)

# 分析軸候補：日付・識別・属性列。内部計算列は除外
internal_cols = {"申込フラグ", "承認フラグ", "成約フラグ", "コスト_元データ", "コスト補完フラグ", "媒体月申込件数", "補完元コスト"}
dimension_cols = [
    c for c in df.columns
    if c not in internal_cols and not pd.api.types.is_numeric_dtype(df[c]) and c not in ["承認日", "成約日"]
]
for preferred in ["申込月", "大項目", "中項目", "小項目", "小項目2", "媒体名", "キャンペーン識別", "メニューコード", "年齢グループ", "年収グループ"]:
    if preferred in df.columns and preferred not in dimension_cols:
        dimension_cols.insert(0, preferred)
dimension_cols = list(dict.fromkeys(dimension_cols))
all_metrics = metric_options(df)

tab_media, tab_win, tab_cross, tab_seg, tab_mail = st.tabs([
    "📊 媒体分析", "🏆 勝ちパターン", "🔄 クロス分析", "👤 セグメント別", "✉️ メルマガ"
])

with tab_media:
    st.subheader("媒体分析")
    default_groups = [c for c in ["大項目", "中項目", "小項目", "キャンペーン識別"] if c in dimension_cols]
    groups = st.multiselect("集計軸", dimension_cols, default=default_groups[:3], key="media_groups")
    default_metrics = [m for m in CORE_METRICS if m in all_metrics]
    metric_actions = st.columns([1, 1, 6])
    if metric_actions[0].button("全選択", key="media_select_all"):
        st.session_state["media_metrics"] = all_metrics
        st.rerun()
    if metric_actions[1].button("基本項目", key="media_select_core"):
        st.session_state["media_metrics"] = default_metrics
        st.rerun()
    selected_metrics = st.multiselect("表示項目", all_metrics, default=default_metrics, key="media_metrics")
    res = aggregate(df, groups)
    res = add_extra_numeric_sums(df, res, groups, selected_metrics)
    display_cols = list(dict.fromkeys(groups + [m for m in selected_metrics if m in res.columns]))
    if not display_cols:
        st.info("集計軸または表示項目を1つ以上選択してください。")
    else:
        media_output = res[display_cols]
        st.dataframe(format_table(media_output), use_container_width=True, hide_index=True)
        st.download_button("📥 媒体分析をダウンロード", to_excel(media_output), "媒体分析.xlsx")

    if "申込日" in df.columns and groups:
        trend_group = st.selectbox("日次推移の系列", groups, index=groups.index("中項目") if "中項目" in groups else 0, key="media_trend_group")
        trend_metric = st.selectbox("日次推移の指標", [m for m in selected_metrics if m in CORE_METRICS] or ["申込件数"], key="media_trend_metric")
        daily_df = df.copy(); daily_df["日付"] = daily_df["申込日"].dt.date
        daily = aggregate(daily_df, ["日付", trend_group])
        pivot = daily.pivot_table(index="日付", columns=trend_group, values=trend_metric, aggfunc="sum")
        st.line_chart(pivot)

with tab_win:
    st.subheader("勝ちパターン分析")
    big_options = sorted(df["大項目"].dropna().astype(str).unique()) if "大項目" in df.columns else []
    selected_big = st.selectbox("分析する大項目", ["すべて"] + big_options)
    win_df = df if selected_big == "すべて" else df[df["大項目"].astype(str) == selected_big]
    pattern_options = [c for c in ["小項目", "媒体名", "キャンペーン識別", "年齢グループ", "年収グループ", "性別", "利用目的", "都道府県"] if c in dimension_cols]
    pattern_defaults = [c for c in ["小項目", "利用目的", "年齢グループ"] if c in pattern_options]
    pattern_cols = st.multiselect("勝ちパターンの組み合わせ", pattern_options, default=pattern_defaults, key="win_pattern_cols")
    target_metric = st.selectbox("最大化する指標", ["承認率", "申込件数", "投下倍率", "取扱金額_翌月", "申込CPA", "承認CPA", "成約件数", "成約率"])
    min_count = st.number_input("最低申込件数", min_value=1, value=30, step=10)
    if pattern_cols:
        win = aggregate(win_df, pattern_cols)
        win = win[win["申込件数"] >= min_count]
        ascending = target_metric in {"申込CPA", "承認CPA"}
        win = win.sort_values(target_metric, ascending=ascending, na_position="last")
        st.dataframe(format_table(win.head(30)), use_container_width=True, hide_index=True)
        if not win.empty:
            top = win.iloc[0]
            combo = " × ".join(str(top[c]) for c in pattern_cols)
            direction = "最小" if ascending else "最大"
            st.success(f"{target_metric}が{direction}の勝ちパターン：{combo}")
        st.download_button("📥 勝ちパターンをダウンロード", to_excel(win), "勝ちパターン.xlsx")
    allocation_simulator(win_df)

with tab_cross:
    st.subheader("クロス分析")
    x = st.selectbox("X軸", dimension_cols, index=dimension_cols.index("年齢グループ") if "年齢グループ" in dimension_cols else 0, key="cross_x")
    y = st.selectbox("Y軸", dimension_cols, index=dimension_cols.index("利用目的") if "利用目的" in dimension_cols else min(1, len(dimension_cols)-1), key="cross_y")
    metric = st.selectbox("指標", [m for m in CORE_METRICS if m in all_metrics], key="cross_metric")
    if x != y:
        cross = aggregate(df, [x, y])
        st.dataframe(format_table(cross), use_container_width=True, hide_index=True)
        pivot = cross.pivot_table(index=x, columns=y, values=metric, aggfunc="sum")

        st.subheader("クロス集計表")
        st.caption("matplotlibに依存しない表示です。行・列をクリックして並べ替えできます。")
        if metric in {"承認率", "成約率"}:
            pivot_display = pivot.map(lambda v: f"{v:.1%}" if pd.notna(v) else "-")
        elif metric == "投下倍率":
            pivot_display = pivot.map(lambda v: f"{v:.2f}" if pd.notna(v) else "-")
        else:
            pivot_display = pivot.map(lambda v: f"{v:,.0f}" if pd.notna(v) else "-")
        st.dataframe(pivot_display, use_container_width=True)

        st.subheader("上位組み合わせ")
        sort_ascending = metric in {"申込CPA", "承認CPA"}
        top_n = st.slider("表示件数", min_value=5, max_value=50, value=15, step=5, key="cross_top_n")
        ranking_cols = [x, y, metric]
        ranking = cross[ranking_cols].sort_values(metric, ascending=sort_ascending, na_position="last").head(top_n)
        st.dataframe(format_table(ranking), use_container_width=True, hide_index=True)

        st.download_button("📥 クロス分析をダウンロード", to_excel(cross), "クロス分析.xlsx")

with tab_seg:
    st.subheader("セグメント別分析")
    seg = st.selectbox("分析項目", dimension_cols, index=dimension_cols.index("小項目") if "小項目" in dimension_cols else 0, key="seg_col")
    chart_metric = st.selectbox("縦棒グラフの指標", [m for m in CORE_METRICS if m in all_metrics], key="seg_metric")
    seg_res = aggregate(df, [seg])
    toggle = seg_res[[seg]].copy()
    toggle.insert(0, "表示", True)
    edited = st.data_editor(
        toggle,
        hide_index=True,
        use_container_width=True,
        disabled=[seg],
        column_config={"表示": st.column_config.CheckboxColumn("グラフ表示")},
        key="seg_toggle",
    )
    selected_values = edited.loc[edited["表示"], seg].tolist()
    chart_df = seg_res[seg_res[seg].isin(selected_values)].set_index(seg)[[chart_metric]]
    st.bar_chart(chart_df)
    st.dataframe(format_table(seg_res), use_container_width=True, hide_index=True)
    st.download_button("📥 セグメント分析をダウンロード", to_excel(seg_res), "セグメント分析.xlsx")

with tab_mail:
    st.subheader("メルマガ分析（中項目 = Mail）")
    if "中項目" not in df.columns:
        st.info("中項目列がありません。")
    else:
        mail_df = df[df["中項目"].astype(str).str.strip().str.casefold() == "mail"]
        if mail_df.empty:
            st.info("現在の共通フィルタ条件では、中項目『Mail』のデータがありません。")
        elif "小項目" not in mail_df.columns:
            st.info("小項目列がないため、小項目別のメルマガ分析を表示できません。")
        else:
            small_options = sorted(
                v for v in mail_df["小項目"].dropna().astype(str).str.strip().unique().tolist()
                if v and v.casefold() != "nan"
            )
            selected_smalls = st.multiselect(
                "表示する小項目（複数選択可）",
                small_options,
                default=small_options[: min(5, len(small_options))],
                key="mail_smalls",
            )
            if not selected_smalls:
                st.info("表示する小項目を1つ以上選択してください。")
            else:
                target_small_df = mail_df[mail_df["小項目"].astype(str).str.strip().isin(selected_smalls)]
                campaign_col = "キャンペーン識別" if "キャンペーン識別" in target_small_df.columns else "小項目2"
                if campaign_col not in target_small_df.columns:
                    st.info("キャンペーン識別に使用できる列がありません。")
                else:
                    campaigns = sorted(
                        v for v in target_small_df[campaign_col].dropna().astype(str).str.strip().unique().tolist()
                        if v and v.casefold() != "nan"
                    )
                    selected_campaigns = st.multiselect(
                        "キャンペーン",
                        campaigns,
                        default=campaigns[: min(10, len(campaigns))],
                        key="mail_campaigns",
                    )
                    target = (
                        target_small_df[target_small_df[campaign_col].astype(str).str.strip().isin(selected_campaigns)]
                        if selected_campaigns else target_small_df
                    )

                    mail_metrics = st.multiselect(
                        "表示項目",
                        CORE_METRICS,
                        default=CORE_METRICS,
                        key="mail_metrics",
                    )
                    g1, g2 = st.columns(2)
                    metric1 = g1.selectbox("比較グラフ1の指標", CORE_METRICS, index=0, key="mail_metric1")
                    metric2 = g2.selectbox(
                        "比較グラフ2の指標",
                        CORE_METRICS,
                        index=CORE_METRICS.index("承認率"),
                        key="mail_metric2",
                    )

                    # 全選択小項目をまとめた一覧。小項目ごとの行として確認・出力できる。
                    all_mail_res = aggregate(target, ["小項目", campaign_col])
                    display_mail_cols = ["小項目", campaign_col] + [m for m in mail_metrics if m in all_mail_res.columns]
                    st.subheader("小項目別一覧")
                    st.dataframe(
                        format_table(all_mail_res[display_mail_cols]),
                        use_container_width=True,
                        hide_index=True,
                    )

                    st.subheader("小項目・キャンペーン比較グラフ")
                    st.caption(
                        "横軸を小項目、系列をキャンペーンとして、選択した小項目を同じグラフ内で比較します。"
                    )

                    # 小項目を横軸、キャンペーンを系列にした集合縦棒グラフ。
                    # Streamlitのbar_chartでは、DataFrameの各列が系列として横並び表示される。
                    chart_source = aggregate(target, ["小項目", campaign_col])

                    chart_left, chart_right = st.columns(2)
                    for chart_area, metric, chart_key in [
                        (chart_left, metric1, "mail_chart_1"),
                        (chart_right, metric2, "mail_chart_2"),
                    ]:
                        chart_area.markdown(f"#### {metric}")
                        if metric not in chart_source.columns:
                            chart_area.info(f"{metric}を集計できません。")
                            continue

                        metric_pivot = chart_source.pivot_table(
                            index="小項目",
                            columns=campaign_col,
                            values=metric,
                            aggfunc="sum",
                        )
                        metric_pivot = metric_pivot.reindex(selected_smalls)

                        # 選択順を維持し、全欠損の系列だけ除外する。
                        metric_pivot = metric_pivot.dropna(axis=1, how="all")
                        if metric_pivot.empty or len(metric_pivot.columns) == 0:
                            chart_area.info("表示できるデータがありません。")
                        else:
                            chart_area.bar_chart(metric_pivot, use_container_width=True)

                    st.download_button(
                        "📥 メルマガ分析をダウンロード",
                        to_excel(all_mail_res),
                        "メルマガ分析_小項目別.xlsx",
                    )

