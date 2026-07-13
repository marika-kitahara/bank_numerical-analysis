
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import date

# ----------------------------------------------------
# ページ設定
# ----------------------------------------------------
st.set_page_config(page_title="後方数値データ分析", layout="wide")
st.title("📊 後方数値データ分析ダッシュボード")

# ----------------------------------------------------
# カテゴリ順序定義
# ----------------------------------------------------
category_orders = {
    "年収帯": ['0-499', '500-999', '1000以上'],
    "借入希望額帯": ['0', '1-9', '10-19', '20-29', '30-39', '40-49', '50-59', '60-69', '70-79', '80-89', '90-99', '100-199', '200-299', '300以上'],
    "住宅ローン帯": ['0', '1-9', '10-19', '20-29', '30-39', '40-49', '50-59', '60-69', '70-79', '80-89', '90-99', '100以上'],
    "勤続年数帯": ['0', '1-3', '4-9', '10-20', '21以上']
}

# サイドバー：ファイルアップロード
st.sidebar.header("ファイルアップロード")
uploaded_data = st.sidebar.file_uploader("後方数値データをアップロード", type=["xlsx"])

# マスタファイル読み込み（GitHub固定）
master_path = "媒体コードマスタ.xlsx"
master = pd.read_excel(master_path)
master.columns = [str(c).strip().replace('\u3000', '').replace('\xa0', '') for c in master.columns]
master.rename(columns={"会社名": "媒体名"}, inplace=True)

id_vars = [col for col in master.columns if col in ["媒体名", "カテゴリ"]]
code_cols = [col for col in master.columns if col not in id_vars]
master_long = master.melt(id_vars=id_vars, value_vars=code_cols,
                          var_name="コード列", value_name="媒体コード").dropna(subset=["媒体コード"])

# 後方数値データ読み込み
if uploaded_data is not None:
    # 後方数値データ読み込み
    df = pd.read_excel(uploaded_data)
    df.columns = [str(c).strip().replace('\u3000', '').replace('\xa0', '') for c in df.columns]

    # 性別整形（例：'xxx_男性' → '男性'）
    if '性別' in df.columns:
        df['性別'] = df['性別'].astype(str).str.extract(r'_(男性|女性)', expand=False).fillna(df['性別'])

    # 数値列変換（存在チェック付き）
    numeric_cols = [
        '年齢', '年収', '同借希望額', '住宅ローン返済月額', '勤続年数',
        '他社借入件数', '取扱金額_申込当月', '取扱金額_申込翌月末', '取扱金額_申込翌々月末'
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')

    # 申込日 → datetime
    if '申込日' in df.columns:
        df['申込日'] = pd.to_datetime(df['申込日'], errors='coerce')

    # 取扱高計算（不足列は0で補完）
    amount_cols = ['取扱金額_申込当月', '取扱金額_申込翌月末', '取扱金額_申込翌々月末']
    for c in amount_cols:
        if c not in df.columns:
            df[c] = 0
    df['取扱高'] = df[amount_cols].sum(axis=1)

    # 承認区分のNULL処理
    if '承認区分' in df.columns:
        df['承認区分'] = df['承認区分'].fillna('NULL')
    else:
        df['承認区分'] = 'NULL'

    # マスタと突合（媒体コードがある前提）
    if '媒体コード' in df.columns and not master_long.empty:
        merged_df = df.merge(master_long, on="媒体コード", how="left")
    else:
        merged_df = df.copy()
        if '媒体名' not in merged_df.columns:
            merged_df['媒体名'] = pd.NA
        if 'カテゴリ' not in merged_df.columns:
            merged_df['カテゴリ'] = pd.NA

    # -------------------------
    # ✅ フィルタUI（日付・カテゴリなど）
    # -------------------------
    st.sidebar.header("フィルタ設定")

    # 日付範囲のデフォルト（NaT除去）
    if '申込日' in merged_df.columns:
        date_series = merged_df['申込日'].dropna()
        if not date_series.empty:
            default_start = date_series.min().date()
            default_end = date_series.max().date()
        else:
            today = date.today()
            default_start, default_end = today, today
        date_range = st.sidebar.date_input("申込日範囲", (default_start, default_end))
        # 単一選択にも対応
        if isinstance(date_range, tuple) and len(date_range) == 2:
            start_date, end_date = date_range
        else:
            start_date = date_range
            end_date = date_range

        filtered_df = merged_df[
            (merged_df['申込日'] >= pd.to_datetime(start_date)) &
            (merged_df['申込日'] <= pd.to_datetime(end_date))
        ].copy()
    else:
        st.sidebar.info("データに『申込日』列がないため、日付フィルタは無効です。")
        filtered_df = merged_df.copy()

    # カテゴリフィルタ
    category_options = ["ALL"] + sorted(filtered_df["カテゴリ"].dropna().unique().tolist())
    selected_categories = st.sidebar.multiselect("カテゴリを選択", category_options, default=["ALL"])
    if "ALL" not in selected_categories:
        filtered_df = filtered_df[filtered_df["カテゴリ"].isin(selected_categories)]

    # 媒体名フィルタ
    company_options = ["ALL"] + sorted(filtered_df["媒体名"].dropna().unique().tolist())
    selected_companies = st.sidebar.multiselect("媒体名を選択", company_options, default=["ALL"])
    if "ALL" not in selected_companies:
        filtered_df = filtered_df[filtered_df["媒体名"].isin(selected_companies)]

    # 承認区分フィルタ
    approval_options = ["ALL"] + sorted(filtered_df["承認区分"].dropna().unique().tolist())
    selected_approval = st.sidebar.multiselect("承認区分を選択", approval_options, default=["ALL"])
    if "ALL" not in selected_approval:
        filtered_df = filtered_df[filtered_df["承認区分"].isin(selected_approval)]

    # 性別フィルタ
    gender_options = ["ALL"] + sorted(filtered_df["性別"].dropna().unique().tolist())
    selected_genders = st.sidebar.multiselect("性別を選択", gender_options, default=["ALL"])
    if "ALL" not in selected_genders:
        filtered_df = filtered_df[filtered_df["性別"].isin(selected_genders)]

    st.write(f"件数: {len(filtered_df):,}件")

    # -------------------------
    # ✅ データ整形（年齢・年収帯など）
    # -------------------------
    def group_age_10(x):
        if pd.isna(x): return "不明"
        try:
            xi = int(float(x))
        except Exception:
            return "不明"
        if xi < 10: return "0-9"
        elif xi < 20: return "10-19"
        elif xi < 30: return "20-29"
        elif xi < 40: return "30-39"
        elif xi < 50: return "40-49"
        elif xi < 60: return "50-59"
        elif xi < 70: return "60-69"
        elif xi < 80: return "70-79"
        elif xi < 90: return "80-89"
        else: return "90以上"

    def group_income(x):
        if pd.isna(x): return "不明"
        return "0-499" if x < 500 else ("500-999" if x < 1000 else "1000以上")

    def group_loan(x):
        if pd.isna(x): return "不明"
        if x == 0: return "0"
        elif x < 10: return "1-9"
        elif x < 20: return "10-19"
        elif x < 30: return "20-29"
        elif x < 40: return "30-39"
        elif x < 50: return "40-49"
        elif x < 60: return "50-59"
        elif x < 70: return "60-69"
        elif x < 80: return "70-79"
        elif x < 90: return "80-89"
        elif x < 100: return "90-99"
        elif x < 200: return "100-199"
        elif x < 300: return "200-299"
        else: return "300以上"

    def group_mortgage(x):
        if pd.isna(x): return "不明"
        if x == 0: return "0"
        elif x < 10: return "1-9"
        elif x < 20: return "10-19"
        elif x < 30: return "20-29"
        elif x < 40: return "30-39"
        elif x < 50: return "40-49"
        elif x < 60: return "50-59"
        elif x < 70: return "60-69"
        elif x < 80: return "70-79"
        elif x < 90: return "80-89"
        elif x < 100: return "90-99"
        else: return "100以上"

    def group_years(x):
        if pd.isna(x): return "不明"
        if x == 0: return "0"
        elif x <= 3: return "1-3"
        elif x <= 9: return "4-9"
        elif x <= 20: return "10-20"
        else: return "21以上"

    if '年齢' in filtered_df.columns:
        filtered_df['年齢'] = filtered_df['年齢'].apply(group_age_10)
    if '年収' in filtered_df.columns:
        filtered_df['年収帯'] = filtered_df['年収'].apply(group_income)
    if '同借希望額' in filtered_df.columns:
        filtered_df['借入希望額帯'] = filtered_df['同借希望額'].apply(group_loan)
    if '住宅ローン返済月額' in filtered_df.columns:
        filtered_df['住宅ローン帯'] = filtered_df['住宅ローン返済月額'].apply(group_mortgage)
    if '勤続年数' in filtered_df.columns:
        filtered_df['勤続年数帯'] = filtered_df['勤続年数'].apply(group_years)

    # -------------------------
    # ✅ フィルタ後データテーブル＋CSV
    # -------------------------
    st.subheader("📋 フィルタ後データ一覧")
    display_cols = []
    if "媒体コード" in filtered_df.columns:
        display_cols.append("媒体コード")
    if "媒体名" in filtered_df.columns:
        display_cols.append("媒体名")
    # 先頭に媒体コード/媒体名を置いて残りを続ける
    display_cols += [col for col in filtered_df.columns if col not in display_cols]
    st.dataframe(filtered_df[display_cols], use_container_width=True)

    csv = filtered_df.to_csv(index=False).encode('utf-8-sig')
    st.download_button(
        label="フィルタ後データCSVをダウンロード",
        data=csv,
        file_name="filtered_data.csv",
        mime="text/csv"
    )

    # -------------------------
    # ✅ 承認率一覧＋CSVエクスポート
    # -------------------------
    if "媒体名" in filtered_df.columns:
        st.subheader("📌 媒体別 承認率一覧（降順）")
        approval_summary = (
            filtered_df.groupby("媒体名", dropna=False)
            .apply(lambda x: pd.Series({
                "件数": len(x),
                "承認件数": (x["承認区分"] == "承認").sum(),
                "承認率(%)": round(((x["承認区分"] == "承認").sum() / len(x) * 100), 2) if len(x) else 0.0
            }))
            .reset_index()
            .sort_values(by="承認率(%)", ascending=False)
        )
        st.dataframe(approval_summary, use_container_width=True)

        csv_approval = approval_summary.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label="承認率一覧CSVをダウンロード",
            data=csv_approval,
            file_name="approval_summary.csv",
            mime="text/csv"
        )

    # -------------------------
    # ✅ グラフ表示（件数＋取扱高のみ）
    # -------------------------
    st.subheader("📈 項目別インタラクティブグラフ")
    chart_cols = [
        ("性別", "性別"),
        ("年齢", "年齢"),
        ("年収", "年収帯"),
        ("都道府県", "都道府県"),
        ("利用目的", "利用目的"),
        ("同借希望額", "借入希望額帯"),
        ("家族構成", "家族構成"),
        ("子供数", "子供数"),
        ("住宅ローン返済月額", "住宅ローン帯"),
        ("勤務状況", "勤務状況"),
        ("勤続年数", "勤続年数帯"),
        ("他社借入件数", "他社借入件数"),
        ("媒体名", "媒体名"),
        ("承認区分", "承認区分")
    ]

    def create_dual_axis_grouped_chart(df, category_col, title):
        # カラム存在＆非空チェック
        if category_col not in df.columns or df[category_col].dropna().shape[0] == 0:
            return go.Figure()

        # カテゴリ順序対応
        if category_col in category_orders:
            ordered_categories = category_orders[category_col]
            count_data = df[category_col].value_counts().reindex(ordered_categories).fillna(0)
            sum_data = df.groupby(category_col)['取扱高'].sum().reindex(ordered_categories).fillna(0)
        else:
            count_data = df[category_col].value_counts().sort_index()
            sum_data = df.groupby(category_col)['取扱高'].sum().reindex(count_data.index).fillna(0)

        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=count_data.index,
            y=count_data.values,
            name="件数",
            marker_color="skyblue",
            offsetgroup=0,
            yaxis="y"
        ))
        fig.add_trace(go.Bar(
            x=sum_data.index,
            y=sum_data.values,
            name="取扱高（円）",
            marker_color="orange",
            offsetgroup=1,
            yaxis="y2"
        ))
        fig.update_layout(
            title=f"{title}（件数＋取扱高）",
            xaxis=dict(title=category_col),
            yaxis=dict(title="件数", side="left"),
            yaxis2=dict(title="取扱高（円）", overlaying="y", side="right"),
            barmode="group",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        return fig

    for title, col in chart_cols:
        if col in filtered_df.columns and filtered_df[col].dropna().shape[0] > 0:
            fig = create_dual_axis_grouped_chart(filtered_df, col, title)
            st.plotly_chart(fig, use_container_width=True)

    # -------------------------
    # ✅ クロス集計（ピボット）
    # -------------------------
    st.subheader("🧮 クロス集計（ピボット）")

    df_pivot_base = filtered_df.copy()
    # 重複列名の除去
    df_pivot_base = df_pivot_base.loc[:, ~pd.Index(df_pivot_base.columns).duplicated()]
    # MultiIndexのフラット化（保険）
    df_pivot_base.columns = [
        "_".join(map(str, c)) if isinstance(c, tuple) else str(c)
        for c in df_pivot_base.columns
    ]

    # 取扱高が無い場合でも動くように補完
    if "取扱高" not in df_pivot_base.columns:
        df_pivot_base["取扱高"] = 0

    # リスト/タプルが紛れているセルを 1次元化し、カテゴリを文字列化
    def to_1d_str(s: pd.Series) -> pd.Series:
        return s.apply(lambda v: v[0] if isinstance(v, (list, tuple)) else v).astype(str)

    pivot_candidates = [
        "性別", "年齢", "年収帯", "都道府県", "利用目的", "借入希望額帯",
        "家族構成", "子供数", "住宅ローン帯", "勤務状況", "勤続年数帯",
        "他社借入件数", "媒体名", "承認区分"
    ]
    available = [c for c in pivot_candidates if c in df_pivot_base.columns]

    if not available:
        st.warning("ピボット可能な項目が見つかりません。")
    else:
        row_dim = st.selectbox("行（Row）", available, index=0)
        col_dim = st.selectbox("列（Column）", ["（なし）"] + available, index=1 if len(available) > 1 else 0)
        value_metric = st.selectbox("値（Value）", ["件数", "取扱高合計"], index=0)
        show_percent = st.checkbox("行方向の構成比（%）を表示", value=False)

        # 選択列を 1次元の文字列に揃える
        df_pivot_base[row_dim] = to_1d_str(df_pivot_base[row_dim])
        if col_dim != "（なし）":
            df_pivot_base[col_dim] = to_1d_str(df_pivot_base[col_dim])

        try:
            if value_metric == "件数":
                if col_dim == "（なし）":
                    # 単純集計（行のみ）
                    result = (
                        df_pivot_base.groupby(row_dim, dropna=False)
                        .size()
                        .reset_index(name="件数")
                        .sort_values(by=row_dim)
                    )
                    if show_percent:
                        total = result["件数"].sum()
                        result["構成比(%)"] = (result["件数"] / total * 100).round(2) if total else 0.0
                    st.dataframe(result, use_container_width=True)
                    csv_bytes = result.to_csv(index=False).encode("utf-8-sig")
                else:
                    # 行 × 列のクロス集計（件数）
                    ct = pd.crosstab(
                        df_pivot_base[row_dim],
                        df_pivot_base[col_dim],
                        dropna=False
                    )
                    if show_percent:
                        denom = ct.sum(axis=1).replace(0, pd.NA)
                        ct_percent = (ct.div(denom, axis=0) * 100).round(2).fillna(0)
                        st.write("行方向の構成比（%）")
                        st.dataframe(ct_percent, use_container_width=True)
                        csv_bytes = ct_percent.reset_index().to_csv(index=False).encode("utf-8-sig")
                    else:
                        st.dataframe(ct, use_container_width=True)
                        csv_bytes = ct.reset_index().to_csv(index=False).encode("utf-8-sig")

            else:  # 取扱高合計
                df_pivot_base["取扱高合計"] = pd.to_numeric(df_pivot_base["取扱高"], errors="coerce").fillna(0)
                if col_dim == "（なし）":
                    result = (
                        df_pivot_base.groupby(row_dim, dropna=False)["取扱高合計"]
                        .sum()
                        .reset_index()
                        .sort_values(by=row_dim)
                    )
                    if show_percent:
                        total = result["取扱高合計"].sum()
                        result["構成比(%)"] = (result["取扱高合計"] / total * 100).round(2) if total else 0.0
                    st.dataframe(result, use_container_width=True)
                    csv_bytes = result.to_csv(index=False).encode("utf-8-sig")
                else:
                    pv = pd.pivot_table(
                        df_pivot_base,
                        index=[row_dim],
                        columns=[col_dim],
                        values="取扱高合計",
                        aggfunc="sum",
                        fill_value=0,
                        dropna=False,
                        sort=True
                    )
                    if show_percent:
                        row_sum = pv.sum(axis=1).replace(0, pd.NA)
                        pv_percent = (pv.div(row_sum, axis=0) * 100).round(2).fillna(0)
                        st.write("行方向の構成比（%）")
                        st.dataframe(pv_percent, use_container_width=True)
                        csv_bytes = pv_percent.reset_index().to_csv(index=False).encode("utf-8-sig")
                    else:
                        st.dataframe(pv, use_container_width=True)
                        csv_bytes = pv.reset_index().to_csv(index=False).encode("utf-8-sig")

            st.download_button("クロス集計CSVをダウンロード", csv_bytes, "pivot.csv", "text/csv")

        except Exception as e:
            with st.expander("🔎 デバッグ情報（開いて確認）"):
                st.write("エラー:", str(e))
                st.write("列一覧:", df_pivot_base.columns.tolist())
                dup_counts = pd.Series(df_pivot_base.columns).value_counts()
                st.write("重複列名（出現回数）:", dup_counts[dup_counts > 1] if (dup_counts > 1).any() else "なし")
                st.write("選択 Row/Column:", row_dim, col_dim)
            st.error("クロス集計でエラーが発生しました。")

else:
    # アップロードが未実施の案内
    st.info("Excelファイル（後方数値データ）をアップロードしてください。")


