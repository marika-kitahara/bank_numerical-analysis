import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(layout="wide")
st.title("📊 後方数値データ分析ダッシュボード")

MASTER_BOX_URL = "https://rak.box.com/s/ocqdd7wgeoqnymhe7lkifycmnckpyncr"

uploaded_master = st.file_uploader(
    "媒体コードマスタアップロード",
    type=["xlsx"],
    help="BOXから最新版をダウンロードしてアップロードしてください"
)

st.markdown(
    f"[📂 媒体コードマスタはこちら]({MASTER_BOX_URL})"
)

uploaded_file = st.file_uploader(
    "後方数値データアップロード",
    type=["xlsx"]
)

# ======================
# キャッシュ
# ======================
@st.cache_data
def load_data(file, master_file):
    df = pd.read_excel(file)
    master = pd.read_excel(master_file)
    return df, master

@st.cache_data
def preprocess(df, master):

    df.columns = df.columns.str.strip()
    master.columns = master.columns.str.strip()

    for col in ["申込日","承認日","成約日"]:
        df[col] = pd.to_datetime(df[col], errors="coerce")

    df["申込月"] = df["申込日"].dt.to_period("M").astype(str)

    df = df.merge(
        master[["媒体コード","媒体名","キャンペーン識別","メニューコード"]],
        on="媒体コード",
        how="left"
    )

    df["承認フラグ"] = df["承認日"].notna().astype(int)
    df["成約フラグ"] = df["成約日"].notna().astype(int)

    for col in ["媒体名","メニューコード","性別","都道府県","利用目的"]:
        if col in df.columns:
            df[col] = df[col].astype("category")

    # ======================
    #  セグメントグループ化
    # ======================

    # 年齢グループ
    if "年齢" in df.columns:
        df["年齢グループ"] = pd.cut(
            df["年齢"],
            bins=[0, 29, 39, 49, 59, 100],
            labels=["20代以下","30代","40代","50代","60代以上"]
        )

    # 年収グループ
    if "年収" in df.columns:
        df["年収グループ"] = pd.cut(
            df["年収"],
            bins=[0, 300, 500, 700, 1000, 99999],
            labels=["300万未満","300-500万","500-700万","700-1000万","1000万以上"]
        )

    return df

@st.cache_data
def aggregate(df, cols):

    cols = list(dict.fromkeys(cols))

    agg = df.groupby(cols, as_index=False).agg(
        全体件数=("媒体コード","count"),
        承認件数=("承認フラグ","sum"),
        成約件数=("成約フラグ","sum"),
        取扱金額_当月=("取扱金額_申込当月","sum"),
        取扱金額_翌月=("取扱金額_申込翌月末","sum"),
        取扱金額_翌々月=("取扱金額_申込翌々月末","sum"),
    )

    agg["承認率"] = agg["承認件数"] / agg["全体件数"]
    agg["成約率"] = agg["成約件数"] / agg["全体件数"]

    return agg

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# ======================
# メイン
# ======================
if uploaded_master and uploaded_file:

    raw_df, master = load_data(uploaded_file, uploaded_master)
    df = preprocess(raw_df, master)

    # ======================
    # フィルタ
    # ======================
    st.sidebar.header("🔍 フィルタ")

    min_d = df["申込日"].min()
    max_d = df["申込日"].max()

    dr = st.sidebar.date_input("申込日（期間）",[min_d, max_d])

    if len(dr)==2:
        df_period = df[
            (df["申込日"] >= pd.to_datetime(dr[0])) &
            (df["申込日"] <= pd.to_datetime(dr[1]))
        ]
    else:
        df_period = df.copy()

    menu_list = sorted(df_period["メニューコード"].dropna().unique())
    selected_menu = st.sidebar.multiselect("メニューコード", menu_list)

    if selected_menu:
        df_media_base = df_period[df_period["メニューコード"].isin(selected_menu)]
    else:
        df_media_base = df_period

    media_list = sorted(df_media_base["媒体名"].dropna().unique())
    selected_media = st.sidebar.multiselect("媒体名", media_list)

    df = df_period.copy()

    if selected_menu:
        df = df[df["メニューコード"].isin(selected_menu)]
    if selected_media:
        df = df[df["媒体名"].isin(selected_media)]

    st.sidebar.caption(f"対象件数: {len(df):,}")

    # ======================
    # KPI
    # ======================
    total = len(df)
    approved = df["承認フラグ"].sum()
    contract = df["成約フラグ"].sum()

    c1,c2,c3,c4,c5 = st.columns(5)
    c1.metric("全体件数", f"{total:,}")
    c2.metric("承認件数", f"{approved:,}")
    c3.metric("承認率", f"{approved/total:.1%}" if total else "0%")
    c4.metric("成約件数", f"{contract:,}")
    c5.metric("成約率", f"{contract/total:.1%}" if total else "0%")

    cross_cols = [
        "申込月","媒体名","キャンペーン識別","メニューコード",
        "年齢","年齢グループ",
        "年収","年収グループ",
        "年齢","性別","年収","都道府県","利用目的","同借希望額",
        "家族構成","子供数","住宅ローン返済月額",
        "勤務状況","勤務先種別","勤務先業種","勤続年数",
        "職種","役職","収入形態","他社借入件数","他社借入残高",
        "当行返済フラグ","承認枠","承認区分"
    ]

    # ======================
    # タブ
    # ======================
    tab1,tab2,tab3 = st.tabs([
        "📊 媒体分析",
        "🔄 クロス分析",
        "👤 セグメント"
    ])

    # ======================
    # 媒体分析
    # ======================
    with tab1:
        res = aggregate(df, ["媒体名","キャンペーン識別"])
        st.dataframe(res)
        st.download_button("📥 ダウンロード", to_excel(res), "媒体分析.xlsx")

        st.subheader("📈 日次推移（媒体別）")

        daily = df.groupby(["申込日","媒体名"]).size().reset_index(name="件数")
        pivot = daily.pivot(index="申込日", columns="媒体名", values="件数").fillna(0)

        pivot.index = pd.to_datetime(pivot.index).strftime("%Y/%m/%d")

        st.line_chart(pivot)

    # ======================
    # クロス分析
    # ======================
    with tab2:

        x = st.selectbox("X軸", cross_cols, index=cross_cols.index("媒体名"))
        y = st.selectbox("Y軸", cross_cols, index=cross_cols.index("利用目的"))

        metric_type = st.radio("指標",["全体件数","承認率","成約率"], horizontal=True)

        if x != y:
            cross = aggregate(df,[x,y])
            st.dataframe(cross)

            st.download_button("📥 ダウンロード", to_excel(cross), "クロス分析.xlsx")

            st.subheader("🔥 ヒートマップ")
            pivot = cross.pivot_table(
                index=x, columns=y, values=metric_type, aggfunc="sum"
            )
            st.dataframe(pivot.style.background_gradient(cmap="Blues"))

            st.subheader(f"📈 日次推移（{x}別）")

            daily = df.groupby(["申込日", x]).size().reset_index(name="件数")
            pivot2 = daily.pivot(index="申込日", columns=x, values="件数").fillna(0)

            pivot2.index = pd.to_datetime(pivot2.index).strftime("%Y/%m/%d")

            st.line_chart(pivot2)

    # ======================
    # セグメント
    # ======================
    with tab3:

        seg = st.selectbox("分析項目", cross_cols)

        res = aggregate(df,[seg])
        st.dataframe(res)

        st.download_button("📥 ダウンロード", to_excel(res), "セグメント.xlsx")

        st.subheader(f"📈 日次推移（{seg}別）")

        daily = df.groupby(["申込日", seg]).size().reset_index(name="件数")
        pivot = daily.pivot(index="申込日", columns=seg, values="件数").fillna(0)

        pivot.index = pd.to_datetime(pivot.index).strftime("%Y/%m/%d")

        st.line_chart(pivot)

    # ======================
    # 🏆 勝ちパターン抽出
    # ======================
    st.subheader("🏆 勝ちパターン")
    
    pattern_dict = {
        "媒体×年齢グループ": ["媒体名","年齢グループ"],
        "年収×利用目的": ["年収グループ","利用目的"],
        "媒体×年収×利用目的": ["媒体名","年収グループ","利用目的"]
    }

    pattern_name = st.selectbox("分析パターン", list(pattern_dict.keys()))

    group_cols = pattern_dict[pattern_name]

    win = aggregate(df, group_cols)

    # ✅ 母数フィルタ（超重要）
    win = win[win["全体件数"] > 30]
    
    win = win.sort_values("承認率", ascending=False)

    st.dataframe(win.head(10), use_container_width=True)

    st.download_button(
        "📥 ダウンロード",
        to_excel(win),
        "勝ちパターン.xlsx"
    )
    st.subheader("🤖 自動分析コメント")

    top = win.head(1)

    if len(top) > 0:
        row = top.iloc[0]

        combo = " × ".join([str(row[col]) for col in group_cols])

        comment = f"""
        最も承認率が高いのは  
        **{combo}**  

        承認率：{row['承認率']:.1%}  
        件数：{row['全体件数']}件  

        """

        st.info(comment)

else:
    st.info("媒体コードマスタと後方数値データをアップロードしてください。")
