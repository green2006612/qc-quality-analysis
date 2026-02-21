import os
import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO

# ─────────────────────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────────────────────
SPEC_COLS = list("ABCDEFGHIJKL")

DATA_FILE = "0217_excel_python test.xlsx"
MAP_FILE  = "0218_color mapping.xlsx"

# 4 categories: key → display label + background color
CATEGORY_COLORS = {
    "standard": "#FFFFFF",   # 符合中央標準 - white/transparent
    "stricter": "#C6EFCE",   # 比中央更嚴格 - green
    "looser":   "#FFD7D7",   # 比中央寬鬆   - pink
    "na":       "#D9D9D9",   # 不採納(NA)   - grey
    "unknown":  "#FFFACD",   # 未知/其他    - yellow
}

CATEGORY_LABELS = {
    "standard": "符合中央標準",
    "stricter": "比中央更嚴格",
    "looser":   "比中央寬鬆",
    "na":       "不採納(NA)",
    "unknown":  "未知",
}

# Map Chinese text (or prefix) → category key
# Handles both A-L cell values and color-mapping Result values
LABEL_TO_CAT = {
    "符合中央標準": "standard",
    "比中央更嚴格": "stricter",
    "比中央寬鬆":   "looser",
    "不採納(NA)":   "na",
    "不採納此規格種類": "na",
    "不採納":       "na",
}

# ─────────────────────────────────────────────────────────────
# Data Loading
# ─────────────────────────────────────────────────────────────
@st.cache_data
def load_data():
    """Load the 0217 sheet and rename columns to standard names."""
    df = pd.read_excel(DATA_FILE, sheet_name="0217")
    df = df.iloc[:, :17]  # keep first 17 columns only
    df.columns = ["Product", "Factory", "Customer", "Location", "Method"] + SPEC_COLS
    df = df.dropna(subset=["Product"])
    df = df[df["Product"].astype(str).str.startswith("Product")].copy()
    # Ensure A-L values are strings for matching
    for col in SPEC_COLS:
        df[col] = df[col].fillna("").astype(str).str.strip()
    df["Method"] = df["Method"].astype(str).str.strip()
    return df


@st.cache_data
def load_mapping():
    """Load the color-mapping rules file."""
    cm = pd.read_excel(MAP_FILE)
    cm = cm.iloc[:, :6]
    cm.columns = ["Method", "Condition", "Spec", "Code", "Result", "WWQA"]
    cm = cm.dropna(subset=["Method", "Code"])
    cm["Code"]   = cm["Code"].astype(str).str.strip()
    cm["Method"] = cm["Method"].astype(str).str.strip()
    cm["Spec"]   = cm["Spec"].astype(str).str.strip()
    cm["Result"] = cm["Result"].astype(str).str.strip()
    return cm


def build_lookup(cm_df: pd.DataFrame) -> dict:
    """Build a fast (Method, Spec, Code) → category key lookup dict."""
    lookup = {}
    for _, row in cm_df.iterrows():
        result = row["Result"]
        cat = "unknown"
        for label, c in LABEL_TO_CAT.items():
            if result.startswith(label):
                cat = c
                break
        lookup[(row["Method"], row["Spec"], row["Code"])] = cat
    return lookup


def classify_value(val: str, method: str, spec: str, lookup: dict) -> str:
    """
    Classify a single cell into one of 4 category keys.
    Priority:
      1. Already a known Chinese label → return category directly
      2. Found in color-mapping lookup → return mapped category
      3. Fall back to 'unknown'
    """
    val = str(val).strip()
    # Step 1 – direct label match
    for label, cat in LABEL_TO_CAT.items():
        if val.startswith(label):
            return cat
    # Step 2 – lookup by (Method, Spec, Code)
    return lookup.get((method, spec, val), "unknown")


@st.cache_data
def classify_all(df: pd.DataFrame, cm_df: pd.DataFrame) -> pd.DataFrame:
    """Add {col}_cat columns to df for all 12 spec columns."""
    lookup = build_lookup(cm_df)
    df2 = df.copy()
    for col in SPEC_COLS:
        df2[f"{col}_cat"] = df2.apply(
            lambda r: classify_value(r[col], r["Method"], col, lookup), axis=1
        )
    return df2


# ─────────────────────────────────────────────────────────────
# Styling helpers
# ─────────────────────────────────────────────────────────────
def apply_spec_colors(df_styles: pd.DataFrame) -> pd.DataFrame:
    """
    Pandas Styler apply() function: colour each A-L cell by its label.
    Accepts a display DataFrame where A-L cells contain CATEGORY_LABELS values.
    Returns a same-shape DataFrame of CSS strings.
    """
    rev_map = {v: k for k, v in CATEGORY_LABELS.items()}
    styles = pd.DataFrame("", index=df_styles.index, columns=df_styles.columns)
    for col in SPEC_COLS:
        if col in df_styles.columns:
            for idx in df_styles.index:
                label = df_styles.at[idx, col]
                cat   = rev_map.get(label, "unknown")
                styles.at[idx, col] = f"background-color: {CATEGORY_COLORS[cat]}"
    return styles


def color_single_col(val: str) -> str:
    """Styler map() for a single column containing CATEGORY_LABELS strings."""
    for label, cat in LABEL_TO_CAT.items():
        full_label = CATEGORY_LABELS.get(cat, "")
        if val == full_label:
            return f"background-color: {CATEGORY_COLORS[cat]}"
    return ""


def build_display_df(filtered_df: pd.DataFrame, id_cols: list) -> pd.DataFrame:
    """
    Build a display DataFrame replacing _cat codes with human-readable labels.
    id_cols: list of columns to include before A-L (e.g. ['Product','Factory',...])
    """
    rows = []
    for _, r in filtered_df.iterrows():
        row = {c: r[c] for c in id_cols if c in r}
        for col in SPEC_COLS:
            cat = r.get(f"{col}_cat", "unknown")
            row[col] = CATEGORY_LABELS.get(cat, cat)
        rows.append(row)
    return pd.DataFrame(rows)


# ─────────────────────────────────────────────────────────────
# Page setup
# ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="QC 品質管理分析",
    layout="wide",
    page_icon="📊",
)
st.title("📊 QC 規格品質分析系統")

# ─────────────────────────────────────────────────────────────
# Load data
# ─────────────────────────────────────────────────────────────
try:
    raw_df = load_data()
    cm_df  = load_mapping()
    df_c   = classify_all(raw_df, cm_df)
    st.sidebar.success(f"✅ 已載入 {len(raw_df)} 筆資料")
except FileNotFoundError as e:
    st.error(f"❌ 找不到檔案: {e}")
    st.info(f"請確認 Excel 檔案放在：`{os.path.abspath('.')}`")
    st.stop()
except Exception as e:
    st.error(f"❌ 資料載入失敗: {e}")
    st.stop()

# ─────────────────────────────────────────────────────────────
# Sidebar – global filters (apply to Tab 1 & Tab 2 only)
# ─────────────────────────────────────────────────────────────
st.sidebar.header("🔧 全域篩選（Tab 1 / Tab 2）")
all_factories = sorted(df_c["Factory"].dropna().unique())
all_methods   = sorted(df_c["Method"].dropna().unique())

sel_factory = st.sidebar.multiselect("🏭 工廠", all_factories, default=all_factories)
sel_method  = st.sidebar.multiselect("🔬 驗收方式", all_methods, default=all_methods)

filtered = df_c[
    df_c["Factory"].isin(sel_factory) &
    df_c["Method"].isin(sel_method)
].copy()

# ─────────────────────────────────────────────────────────────
# Legend component (reused across tabs)
# ─────────────────────────────────────────────────────────────
def show_legend():
    item_map = [
        ("standard", "符合中央標準 ✓", "White"),
        ("stricter",  "比中央更嚴格 ▲", "Green"),
        ("looser",    "比中央寬鬆   ▼", "Pink"),
        ("na",        "不採納(NA)  —",  "Grey"),
    ]
    cols = st.columns(4)
    for i, (cat, label, _) in enumerate(item_map):
        cols[i].markdown(
            f"<div style='background:{CATEGORY_COLORS[cat]};"
            f"border:1px solid #bbb;padding:6px 10px;border-radius:6px;"
            f"text-align:center;font-size:0.88rem'>{label}</div>",
            unsafe_allow_html=True,
        )
    st.markdown("")


# ─────────────────────────────────────────────────────────────
# Tabs
# ─────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs([
    "📋 產品×規格 顏色表",
    "📈 XYZ vs A-L 分布",
    "🔍 條件篩選名單",
    "↔️ 跨工廠比較",
])

# ══════════════════════════════════════════════════════════════
# TAB 1 – Product × A-L colour heatmap table
# ══════════════════════════════════════════════════════════════
with tab1:
    st.subheader("產品 × 規格(A-L) 顏色狀態總表")
    show_legend()

    if filtered.empty:
        st.warning("⚠️ 目前篩選條件下無資料，請調整側邊欄的篩選選項。")
    else:
        disp_df = build_display_df(
            filtered,
            id_cols=["Product", "Factory", "Customer", "驗收方式"],
        )
        # Fix: '驗收方式' column may not exist if renamed; use 'Method'
        disp_df = build_display_df(
            filtered,
            id_cols=["Product", "Factory", "Customer", "Method"],
        )
        disp_df = disp_df.rename(columns={"Method": "驗收方式"})

        st.dataframe(
            disp_df.style.apply(apply_spec_colors, axis=None),
            use_container_width=True,
            height=min(48 + len(disp_df) * 35, 550),
        )

        # Summary stats
        st.markdown("#### 📊 類別數量統計")
        cat_counts = {"類別": [], "數量": [], "佔比": []}
        total_cells = len(filtered) * len(SPEC_COLS)
        for cat, label in CATEGORY_LABELS.items():
            count = sum(
                (filtered[f"{col}_cat"] == cat).sum() for col in SPEC_COLS
            )
            cat_counts["類別"].append(label)
            cat_counts["數量"].append(count)
            cat_counts["佔比"].append(f"{count/total_cells*100:.1f}%")
        st.dataframe(pd.DataFrame(cat_counts), use_container_width=True, hide_index=True)

        # Download
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            disp_df.to_excel(writer, sheet_name="顏色總表", index=False)
        st.download_button(
            "📥 下載此表格 (Excel)",
            data=buf.getvalue(),
            file_name="product_spec_color_table.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ══════════════════════════════════════════════════════════════
# TAB 2 – XYZ vs A-L distribution (stacked bar + table)
# ══════════════════════════════════════════════════════════════
with tab2:
    st.subheader("驗收方式 (X/Y/Z) × 規格 (A-L) — 四類分布")
    show_legend()

    if filtered.empty:
        st.warning("⚠️ 目前篩選條件下無資料。")
    else:
        # Build long-form data
        dist_rows = []
        for _, r in filtered.iterrows():
            for col in SPEC_COLS:
                cat = r[f"{col}_cat"]
                dist_rows.append({
                    "Method": r["Method"],
                    "Spec":   col,
                    "cat":    cat,
                    "Label":  CATEGORY_LABELS.get(cat, cat),
                })
        dist_df = pd.DataFrame(dist_rows)
        agg = dist_df.groupby(["Method", "Spec", "Label"]).size().reset_index(name="Count")

        color_map = {CATEGORY_LABELS[k]: v for k, v in CATEGORY_COLORS.items()}
        label_order = list(CATEGORY_LABELS.values())

        fig = px.bar(
            agg,
            x="Spec", y="Count", color="Label",
            facet_col="Method",
            barmode="stack",
            color_discrete_map=color_map,
            category_orders={"Label": label_order},
            title="各規格欄位 × 驗收方式 的四類分布（堆疊長條圖）",
            labels={"Count": "數量", "Spec": "規格", "Label": "類別"},
        )
        fig.update_layout(height=420, legend_title_text="類別")
        fig.update_xaxes(tickangle=0)
        st.plotly_chart(fig, use_container_width=True)

        # Numeric pivot table
        st.markdown("#### 📋 數字明細表（Method × Spec × 類別）")
        pivot = agg.pivot_table(
            index=["Method", "Spec"],
            columns="Label",
            values="Count",
            fill_value=0,
            aggfunc="sum",
        )
        # Reorder columns if possible
        ordered_cols = [CATEGORY_LABELS[k] for k in ["standard","stricter","looser","na","unknown"] if CATEGORY_LABELS[k] in pivot.columns]
        pivot = pivot[ordered_cols]
        st.dataframe(pivot, use_container_width=True)

# ══════════════════════════════════════════════════════════════
# TAB 3 – A-X-Color → Factory + Customer drill-down
# ══════════════════════════════════════════════════════════════
with tab3:
    st.subheader("🔍 條件篩選：規格 + 驗收方式 + 顏色 → 工廠 & 客戶名單")
    st.caption("例：規格 A ＋ Method-X ＋ 比中央寬鬆(粉色) → 哪些工廠？哪些客戶？")

    c1, c2, c3 = st.columns(3)
    t3_spec   = c1.selectbox("📐 規格 (A-L)",  SPEC_COLS, key="t3s")
    t3_method = c2.selectbox(
        "🔬 驗收方式",
        sorted(df_c["Method"].dropna().unique()),
        key="t3m",
    )
    t3_cat = c3.selectbox(
        "🎨 顏色類別",
        options=["standard", "stricter", "looser", "na"],
        format_func=lambda x: {
            "standard": "⬜ 符合中央標準",
            "stricter": "🟢 比中央更嚴格",
            "looser":   "🩷 比中央寬鬆",
            "na":       "⬜ 不採納(NA)",
        }[x],
        key="t3c",
    )

    result_df = df_c[
        (df_c["Method"] == t3_method) &
        (df_c[f"{t3_spec}_cat"] == t3_cat)
    ].copy()

    if len(result_df) > 0:
        chosen_color = CATEGORY_COLORS[t3_cat]
        st.markdown(
            f"<div style='background:{chosen_color};border:1px solid #ccc;"
            f"padding:10px;border-radius:8px;font-size:1rem'>"
            f"✅ 規格 <b>{t3_spec}</b> ＋ <b>{t3_method}</b> ＋ "
            f"<b>{CATEGORY_LABELS[t3_cat]}</b>：共 <b>{len(result_df)}</b> 筆</div>",
            unsafe_allow_html=True,
        )
        st.markdown("")

        col_f, col_c = st.columns(2)
        with col_f:
            st.markdown("#### 🏭 工廠名單")
            fac_cnt = (
                result_df["Factory"].value_counts()
                .reset_index()
                .rename(columns={"Factory": "工廠", "count": "筆數"})
            )
            # Compatibility: pandas value_counts().reset_index() column naming varies
            fac_cnt.columns = ["工廠", "筆數"]
            st.dataframe(fac_cnt, use_container_width=True, hide_index=True)

        with col_c:
            st.markdown("#### 👤 客戶名單")
            cus_cnt = (
                result_df["Customer"].value_counts()
                .reset_index()
                .rename(columns={"Customer": "客戶", "count": "筆數"})
            )
            cus_cnt.columns = ["客戶", "筆數"]
            st.dataframe(cus_cnt, use_container_width=True, hide_index=True)

        st.markdown("#### 📋 詳細明細")
        detail = result_df[["Product", "Factory", "Customer", "Location", "Method", t3_spec]].copy()
        detail.insert(detail.columns.get_loc(t3_spec) + 1, "類別", CATEGORY_LABELS[t3_cat])
        st.dataframe(detail.reset_index(drop=True), use_container_width=True, hide_index=True)

    else:
        st.warning(f"⚠️ 找不到符合 **{t3_spec} - {t3_method} - {CATEGORY_LABELS[t3_cat]}** 的資料")

# ══════════════════════════════════════════════════════════════
# TAB 4 – Cross-factory product comparison
# ══════════════════════════════════════════════════════════════
with tab4:
    st.subheader("特定產品編號 → 各工廠規格顏色對比")

    all_products = sorted(df_c["Product"].dropna().unique())
    t4_product = st.selectbox("🔍 選擇產品編號", all_products, key="t4p")

    p_df = df_c[df_c["Product"] == t4_product].copy()

    st.info(f"產品 **{t4_product}** 共有 **{len(p_df)}** 筆生產記錄")

    if p_df.empty:
        st.warning("⚠️ 此產品無資料")
    else:
        show_legend()

        # Full coloured comparison table
        st.markdown("#### 📋 各工廠 × 規格(A-L) 顏色總覽")
        p_disp = build_display_df(p_df, id_cols=["Factory", "Customer", "Method"])
        p_disp = p_disp.rename(columns={"Method": "驗收方式"})

        st.dataframe(
            p_disp.style.apply(apply_spec_colors, axis=None),
            use_container_width=True,
        )

        st.markdown("---")
        st.markdown("#### 🎯 單一規格欄跨工廠顏色比較")

        t4_spec = st.selectbox("選擇規格欄位", SPEC_COLS, key="t4spec")

        spec_rows = []
        for _, r in p_df.iterrows():
            cat = r[f"{t4_spec}_cat"]
            spec_rows.append({
                "工廠":       r["Factory"],
                "客戶":       r["Customer"],
                "驗收方式":   r["Method"],
                f"{t4_spec} 原始值": r[t4_spec],
                f"{t4_spec} 類別":   CATEGORY_LABELS.get(cat, cat),
            })
        spec_df = pd.DataFrame(spec_rows)

        cat_col = f"{t4_spec} 類別"
        st.dataframe(
            spec_df.style.map(color_single_col, subset=[cat_col]),
            use_container_width=True,
            hide_index=True,
        )

        # Further filter by colour within this product
        st.markdown(f"**篩選：{t4_product} ＋ 規格 {t4_spec} ＋ 特定顏色 → 工廠 & 客戶**")
        t4_cat_filter = st.selectbox(
            "選擇顏色類別",
            options=["standard", "stricter", "looser", "na"],
            format_func=lambda x: {
                "standard": "⬜ 符合中央標準",
                "stricter": "🟢 比中央更嚴格",
                "looser":   "🩷 比中央寬鬆",
                "na":       "⬜ 不採納(NA)",
            }[x],
            key="t4cf",
        )

        filtered_p = p_df[p_df[f"{t4_spec}_cat"] == t4_cat_filter]
        if len(filtered_p) > 0:
            col_fa, col_cu = st.columns(2)
            with col_fa:
                st.markdown(f"**🏭 工廠（{CATEGORY_LABELS[t4_cat_filter]}）**")
                st.dataframe(
                    filtered_p[["Factory", "Customer", "Method"]].reset_index(drop=True),
                    use_container_width=True,
                    hide_index=True,
                )
            with col_cu:
                st.markdown("**👤 客戶統計**")
                cu_count = (
                    filtered_p["Customer"].value_counts()
                    .reset_index()
                )
                cu_count.columns = ["客戶", "次數"]
                st.dataframe(cu_count, use_container_width=True, hide_index=True)
        else:
            st.warning(
                f"⚠️ {t4_product} × {t4_spec} 中沒有 {CATEGORY_LABELS[t4_cat_filter]} 的工廠"
            )