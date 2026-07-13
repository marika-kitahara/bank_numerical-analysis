import io
from datetime import date

import pandas as pd
import plotly.graph_objects as go
import streamlit as st


# ----------------------------------------------------
# ページ設定
# ----------------------------------------------------
st.set_page_config(
    page_title="後方数値データ分析",
    layout="wide",
)
st.title("📊 後方数値データ分析ダッシュボード")


# ----------------------------------------------------
# 定数
# ----------------------------------------------------
MASTER_BOX_URL = "https://rak.box.com/s/ocqdd7wgeoqnymhe7lkifycmnckpyncr"

CATEGORY_ORDERS = {
    "年収帯": ["0-499", "500-999", "1000以上"],
    "借入希望額帯": [
        "0", "1-9", "10-19", "20-29", "30-39", "40-49",
        "50-59", "60-69", "70-79", "80-89", "90-99",
        "100-199", "200-299", "300以上",
    ],
    "住宅ローン帯": [
        "0", "1-9", "10-19", "20-29", "30-39", "40-49",
        "50-59", "60-69", "70-79", "80-89", "90-99", "100以上",
    ],
    "勤続年数帯": ["0", "1-3", "4-9", "10-20", "21以上"],
}

AMOUNT_COLS = [
    "取扱金額_申込当月",
    "取扱金額_申込翌月末",
    "取扱金額_申込翌々月末",
]

NUMERIC_COLS = [
    "年齢",
    "年収",
    "同借希望額",
    "住宅ローン返済月額",
    "勤続年数",
    "他社借入件数",
    *AMOUNT_COLS,
]

CHART_OPTIONS = {
    "性別": "性別",
    "年齢": "年齢",
    "年収": "年収帯",
    "都道府県": "都道府県",
    "利用目的": "利用目的",
    "同借希望額": "借入希望額帯",
    "家族構成": "家族構成",
    "子供数": "子供数",
    "住宅ローン返済月額": "住宅ローン帯",
    "勤務状況": "勤務状況",
    "勤続年数": "勤続年数帯",
    "他社借入件数": "他社借入件数",
    "媒体名": "媒体名",
    "承認区分": "承認区分",
}

PIVOT_CANDIDATES = [
    "性別",
    "年齢",
    "年収帯",
    "都道府県",
    "利用目的",
    "借入希望額帯",
    "家族構成",
    "子供数",
    "住宅ローン帯",
    "勤務状況",
    "勤続年数帯",
    "他社借入件数",
    "媒体名",
    "カテゴリ",
    "承認区分",
]


# ----------------------------------------------------
# 共通関数
# ----------------------------------------------------
def normalize_column_name(value) -> str:
    return (
        str(value)
        .strip()
        .replace("\u3000", "")
        .replace("\xa0", "")
    )


def normalize_code(value):
    if pd.isna(value):
        return pd.NA

    if isinstance(value, float) and value.is_integer():
        return str(int(value))

    text = str(value).strip()

    if text.endswith(".0"):
        text = text[:-2]

    return text if text else pd.NA


def group_age_10(value) -> str:
    if pd.isna(value):
        return "不明"

    try:
        number = int(float(value))
    except (TypeError, ValueError):
        return "不明"

    if number < 10:
        return "0-9"
    if number < 20:
        return "10-19"
    if number < 30:
        return "20-29"
    if number < 40:
        return "30-39"
    if number < 50:
        return "40-49"
    if number < 60:
        return "50-59"
    if number < 70:
        return "60-69"
    if number < 80:
        return "70-79"
    if number < 90:
        return "80-89"
    return "90以上"


def group_income(value) -> str:
    if pd.isna(value):
        return "不明"
    if value < 500:
        return "0-499"
    if value < 1000:
        return "500-999"
    return "1000以上"


def group_loan(value) -> str:
    if pd.isna(value):
        return "不明"
    if value == 0:
        return "0"
    if value < 10:
        return "1-9"
    if value < 20:
        return "10-19"
    if value < 30:
        return "20-29"
    if value < 40:
        return "30-39"
    if value < 50:
        return "40-49"
    if value < 60:
        return "50-59"
    if value < 70:
        return "60-69"
    if value < 80:
        return "70-79"
    if value < 90:
        return "80-89"
    if value < 100:
        return "90-99"
    if value < 200:
        return "100-199"
    if value < 300:
        return "200-299"
    return "300以上"


def group_mortgage(value) -> str:
    if pd.isna(value):
        return "不明"
    if value == 0:
        return "0"
    if value < 10:
        return "1-9"
    if value < 20:
        return "10-19"
    if value < 30:
        return "20-29"
    if value < 40:
        return "30-39"
    if value < 50:
        return "40-49"
    if value < 60:
        return "50-59"
    if value < 70:
        return "60-69"
    if value < 80:
        return "70-79"
    if value < 90:
        return "80-89"
    if value < 100:
        return "90-99"
    return "100以上"


def group_years(value) -> str:
    if pd.isna(value):
        return "不明"
    if value == 0:
        return "0"
    if value <= 3:
        return "1-3"
    if value <= 9:
        return "4-9"
    if value <= 20:
        return "10-20"
    return "21以上"


@st.cache_data(show_spinner="媒体コードマスタを読み込んでいます...")
def read_master(file_bytes: bytes) -> pd.DataFrame:
    master = pd.read_excel(io.BytesIO(file_bytes))
    master.columns = [normalize_column_name(col) for col in master.columns]

    master.rename(
        columns={
            "会社名": "媒体名",
            "メニューコード": "カテゴリ",
        },
        inplace=True,
    )

    required = ["媒体名", "カテゴリ"]
    missing = [col for col in required if col not in master.columns]

    if missing:
        raise ValueError(
            "媒体コードマスタに必要な列がありません："
            + "、".join(missing)
        )

    # 縦持ち形式
    if "媒体コード" in master.columns:
        master_long = master[
            ["媒体名", "カテゴリ", "媒体コード"]
        ].copy()

    # 横持ち形式
    else:
        id_vars = ["媒体名", "カテゴリ"]

        # コード列候補だけを対象にする。
        # まず列名に「コード」を含む列を優先。
        code_cols = [
            col for col in master.columns
            if col not in id_vars and "コード" in str(col)
        ]

        # 見つからない場合のみ、従来どおりID列以外を対象にする。
        if not code_cols:
            code_cols = [
                col for col in master.columns
                if col not in id_vars
            ]

        if not code_cols:
            raise ValueError(
                "媒体コードとして使用できる列が見つかりませんでした。"
            )

        master_long = master.melt(
            id_vars=id_vars,
            value_vars=code_cols,
            var_name="コード列",
            value_name="媒体コード",
        )

    master_long["媒体コード"] = master_long["媒体コード"].map(
        normalize_code
    )
    master_long["媒体名"] = (
        master_long["媒体名"]
        .astype("string")
        .str.strip()
    )
    master_long["カテゴリ"] = (
        master_long["カテゴリ"]
        .astype("string")
        .str.strip()
    )

    master_long = (
        master_long
        .dropna(subset=["媒体コード"])
        .drop_duplicates(subset=["媒体コード"], keep="first")
        .reset_index(drop=True)
    )

    return master_long


@st.cache_data(show_spinner="後方数値データを読み込んでいます...")
def read_data(file_bytes: bytes) -> pd.DataFrame:
    df = pd.read_excel(io.BytesIO(file_bytes))
    df.columns = [normalize_column_name(col) for col in df.columns]

    if "媒体コード" in df.columns:
        df["媒体コード"] = df["媒体コード"].map(normalize_code)

    if "性別" in df.columns:
        original_gender = df["性別"].astype("string")
        extracted_gender = original_gender.str.extract(
            r"_(男性|女性)",
            expand=False,
        )
        df["性別"] = extracted_gender.fillna(original_gender)

    for col in NUMERIC_COLS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    if "申込日" in df.columns:
        df["申込日"] = pd.to_datetime(
            df["申込日"],
            errors="coerce",
        )

    for col in AMOUNT_COLS:
        if col not in df.columns:
            df[col] = 0

    df["取扱高"] = df[AMOUNT_COLS].sum(axis=1)

    if "承認区分" in df.columns:
        df["承認区分"] = df["承認区分"].fillna("NULL")
    else:
        df["承認区分"] = "NULL"

    return df


def add_group_columns(df: pd.DataFrame) -> pd.DataFrame:
    result = df.copy()

    if "年齢" in result.columns:
        result["年齢"] = result["年齢"].map(group_age_10)

    if "年収" in result.columns:
        result["年収帯"] = result["年収"].map(group_income)

    if "同借希望額" in result.columns:
        result["借入希望額帯"] = result["同借希望額"].map(group_loan)

    if "住宅ローン返済月額" in result.columns:
        result["住宅ローン帯"] = result[
            "住宅ローン返済月額"
        ].map(group_mortgage)

    if "勤続年数" in result.columns:
        result["勤続年数帯"] = result["勤続年数"].map(group_years)

    return result


def filter_multiselect(
    df: pd.DataFrame,
    column: str,
    label: str,
) -> pd.DataFrame:
    if column not in df.columns:
        return df

    options = sorted(
        df[column]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )

    selected = st.sidebar.multiselect(
        label,
        ["ALL", *options],
        default=["ALL"],
        key=f"filter_{column}",
    )

    if "ALL" in selected:
        return df

    return df[df[column].astype(str).isin(selected)].copy()


def create_dual_axis_chart(
    df: pd.DataFrame,
    category_col: str,
    title: str,
) -> go.Figure:
    if (
        category_col not in df.columns
        or df[category_col].dropna().empty
    ):
        return go.Figure()

    if category_col in CATEGORY_ORDERS:
        order = CATEGORY_ORDERS[category_col]
        count_data = (
            df[category_col]
            .value_counts()
            .reindex(order)
            .fillna(0)
        )
        sum_data = (
            df.groupby(category_col, dropna=False)["取扱高"]
            .sum()
            .reindex(order)
            .fillna(0)
        )
    else:
        count_data = (
            df[category_col]
            .astype("string")
            .value_counts()
            .sort_index()
        )
        sum_data = (
            df.groupby(category_col, dropna=False)["取扱高"]
            .sum()
            .reindex(count_data.index)
            .fillna(0)
        )

    fig = go.Figure()

    fig.add_trace(
        go.Bar(
            x=count_data.index.astype(str),
            y=count_data.values,
            name="件数",
            offsetgroup=0,
            yaxis="y",
        )
    )

    fig.add_trace(
        go.Bar(
            x=sum_data.index.astype(str),
            y=sum_data.values,
            name="取扱高（円）",
            offsetgroup=1,
            yaxis="y2",
        )
    )

    fig.update_layout(
        title=f"{title}（件数＋取扱高）",
        xaxis={"title": category_col},
        yaxis={"title": "件数", "side": "left"},
        yaxis2={
            "title": "取扱高（円）",
            "overlaying": "y",
            "side": "right",
        },
        barmode="group",
        legend={
            "orientation": "h",
            "yanchor": "bottom",
            "y": 1.02,
            "xanchor": "right",
            "x": 1,
        },
    )

    return fig


def to_1d_str(series: pd.Series) -> pd.Series:
    return series.apply(
        lambda value: value[0]
        if isinstance(value, (list, tuple))
        else value
    ).astype(str)


# ----------------------------------------------------
# サイドバー：アップロード
# ----------------------------------------------------
st.sidebar.header("ファイルアップロード")

uploaded_master = st.sidebar.file_uploader(
    "媒体コードマスタをアップロード",
    type=["xlsx"],
    key="uploaded_master",
    help="BOXから最新版をダウンロードしてアップロードしてください",
)

st.sidebar.markdown(
    f"[📂 媒体コードマスタはこちら]({MASTER_BOX_URL})"
)

uploaded_data = st.sidebar.file_uploader(
    "後方数値データをアップロード",
    type=["xlsx"],
    key="uploaded_data",
)

if uploaded_master is None or uploaded_data is None:
    st.info(
        "媒体コードマスタと後方数値データを"
        "アップロードしてください。"
    )
    st.stop()


# ----------------------------------------------------
# 読み込み・突合
# ----------------------------------------------------
try:
    master_long = read_master(uploaded_master.getvalue())
except Exception as exc:
    st.error(f"媒体コードマスタの読み込みに失敗しました：{exc}")
    st.stop()

try:
    df = read_data(uploaded_data.getvalue())
except Exception as exc:
    st.error(f"後方数値データの読み込みに失敗しました：{exc}")
    st.stop()

if "媒体コード" in df.columns and not master_long.empty:
    try:
        merged_df = df.merge(
            master_long,
            on="媒体コード",
            how="left",
            validate="many_to_one",
        )
    except Exception as exc:
        st.error(f"媒体コードマスタとの突合に失敗しました：{exc}")
        st.stop()
else:
    merged_df = df.copy()

if "媒体名" not in merged_df.columns:
    merged_df["媒体名"] = pd.NA

if "カテゴリ" not in merged_df.columns:
    merged_df["カテゴリ"] = pd.NA


# ----------------------------------------------------
# フィルタ
# ----------------------------------------------------
st.sidebar.header("フィルタ設定")

filtered_df = merged_df.copy()

if "申込日" in filtered_df.columns:
    date_series = filtered_df["申込日"].dropna()

    if date_series.empty:
        today = date.today()
        default_start = today
        default_end = today
    else:
        default_start = date_series.min().date()
        default_end = date_series.max().date()

    date_range = st.sidebar.date_input(
        "申込日範囲",
        value=(default_start, default_end),
    )

    if isinstance(date_range, (list, tuple)) and len(date_range) == 2:
        start_date, end_date = date_range
    else:
        start_date = date_range
        end_date = date_range

    filtered_df = filtered_df[
        filtered_df["申込日"].between(
            pd.to_datetime(start_date),
            pd.to_datetime(end_date),
            inclusive="both",
        )
    ].copy()

filtered_df = filter_multiselect(
    filtered_df,
    "カテゴリ",
    "カテゴリを選択",
)
filtered_df = filter_multiselect(
    filtered_df,
    "媒体名",
    "媒体名を選択",
)
filtered_df = filter_multiselect(
    filtered_df,
    "承認区分",
    "承認区分を選択",
)
filtered_df = filter_multiselect(
    filtered_df,
    "性別",
    "性別を選択",
)

filtered_df = add_group_columns(filtered_df)

st.metric("件数", f"{len(filtered_df):,}件")


# ----------------------------------------------------
# データ一覧・CSV
# ----------------------------------------------------
st.subheader("📋 フィルタ後データ一覧")

preview_limit = st.number_input(
    "画面に表示する最大行数",
    min_value=100,
    max_value=5000,
    value=500,
    step=100,
    help=(
        "画面表示だけを制限します。"
        "ダウンロードCSVにはフィルタ後の全件が含まれます。"
    ),
)

display_cols = []

for col in ["媒体コード", "媒体名", "カテゴリ"]:
    if col in filtered_df.columns:
        display_cols.append(col)

display_cols.extend(
    col for col in filtered_df.columns
    if col not in display_cols
)

st.dataframe(
    filtered_df[display_cols].head(int(preview_limit)),
    width="stretch",
)

csv_bytes = filtered_df.to_csv(
    index=False,
).encode("utf-8-sig")

st.download_button(
    "フィルタ後データCSVをダウンロード",
    data=csv_bytes,
    file_name="filtered_data.csv",
    mime="text/csv",
)


# ----------------------------------------------------
# 承認率一覧
# ----------------------------------------------------
if "媒体名" in filtered_df.columns:
    st.subheader("📌 媒体別 承認率一覧（降順）")

    approval_summary = (
        filtered_df.assign(
            承認フラグ=filtered_df["承認区分"].eq("承認").astype(int)
        )
        .groupby("媒体名", dropna=False)
        .agg(
            件数=("承認区分", "size"),
            承認件数=("承認フラグ", "sum"),
        )
        .reset_index()
    )

    approval_summary["承認率(%)"] = (
        approval_summary["承認件数"]
        .div(approval_summary["件数"].replace(0, pd.NA))
        .mul(100)
        .fillna(0)
        .round(2)
    )

    approval_summary = approval_summary.sort_values(
        "承認率(%)",
        ascending=False,
    )

    st.dataframe(
        approval_summary,
        width="stretch",
    )

    st.download_button(
        "承認率一覧CSVをダウンロード",
        data=approval_summary.to_csv(
            index=False,
        ).encode("utf-8-sig"),
        file_name="approval_summary.csv",
        mime="text/csv",
    )


# ----------------------------------------------------
# グラフ
# 一度に全グラフを生成せず、選択した1つだけ描画する
# ----------------------------------------------------
st.subheader("📈 項目別インタラクティブグラフ")

available_chart_labels = [
    label
    for label, column in CHART_OPTIONS.items()
    if column in filtered_df.columns
    and not filtered_df[column].dropna().empty
]

if available_chart_labels:
    selected_chart_label = st.selectbox(
        "表示するグラフ",
        available_chart_labels,
    )
    selected_chart_col = CHART_OPTIONS[selected_chart_label]

    chart = create_dual_axis_chart(
        filtered_df,
        selected_chart_col,
        selected_chart_label,
    )

    st.plotly_chart(
        chart,
        width="stretch",
    )
else:
    st.info("表示できるグラフ項目がありません。")


# ----------------------------------------------------
# クロス集計
# 実行ボタンを押したときだけ作る
# ----------------------------------------------------
st.subheader("🧮 クロス集計（ピボット）")

pivot_base = filtered_df.loc[
    :,
    ~pd.Index(filtered_df.columns).duplicated(),
].copy()

available_pivot_cols = [
    col for col in PIVOT_CANDIDATES
    if col in pivot_base.columns
]

if not available_pivot_cols:
    st.info("ピボット可能な項目がありません。")
else:
    row_dim = st.selectbox(
        "行（Row）",
        available_pivot_cols,
    )

    col_dim = st.selectbox(
        "列（Column）",
        ["（なし）", *available_pivot_cols],
    )

    value_metric = st.selectbox(
        "値（Value）",
        ["件数", "取扱高合計"],
    )

    show_percent = st.checkbox(
        "行方向の構成比（%）を表示",
        value=False,
    )

    if st.button("クロス集計を作成"):
        pivot_base[row_dim] = to_1d_str(pivot_base[row_dim])

        if col_dim != "（なし）":
            pivot_base[col_dim] = to_1d_str(pivot_base[col_dim])

        if value_metric == "件数":
            if col_dim == "（なし）":
                result = (
                    pivot_base
                    .groupby(row_dim, dropna=False)
                    .size()
                    .reset_index(name="件数")
                )

                if show_percent:
                    total = result["件数"].sum()
                    result["構成比(%)"] = (
                        result["件数"] / total * 100
                    ).round(2) if total else 0.0

            else:
                result = pd.crosstab(
                    pivot_base[row_dim],
                    pivot_base[col_dim],
                    dropna=False,
                )

                if show_percent:
                    row_total = result.sum(axis=1).replace(0, pd.NA)
                    result = (
                        result.div(row_total, axis=0)
                        .mul(100)
                        .round(2)
                        .fillna(0)
                    )

        else:
            pivot_base["取扱高合計"] = pd.to_numeric(
                pivot_base["取扱高"],
                errors="coerce",
            ).fillna(0)

            if col_dim == "（なし）":
                result = (
                    pivot_base
                    .groupby(row_dim, dropna=False)["取扱高合計"]
                    .sum()
                    .reset_index()
                )

                if show_percent:
                    total = result["取扱高合計"].sum()
                    result["構成比(%)"] = (
                        result["取扱高合計"] / total * 100
                    ).round(2) if total else 0.0

            else:
                result = pd.pivot_table(
                    pivot_base,
                    index=row_dim,
                    columns=col_dim,
                    values="取扱高合計",
                    aggfunc="sum",
                    fill_value=0,
                    dropna=False,
                )

                if show_percent:
                    row_total = result.sum(axis=1).replace(0, pd.NA)
                    result = (
                        result.div(row_total, axis=0)
                        .mul(100)
                        .round(2)
                        .fillna(0)
                    )

        st.dataframe(
            result,
            width="stretch",
        )

        download_df = (
            result.reset_index()
            if isinstance(result.index, pd.Index)
            else result
        )

        st.download_button(
            "クロス集計CSVをダウンロード",
            data=download_df.to_csv(
                index=False,
            ).encode("utf-8-sig"),
            file_name="pivot.csv",
            mime="text/csv",
        )
