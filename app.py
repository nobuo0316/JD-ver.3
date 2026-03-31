import io
from datetime import datetime
from typing import Optional

import pandas as pd
import requests
import streamlit as st

st.set_page_config(
    page_title="Job Matrix Manager - HQ & Grade Edition",
    page_icon="🌱",
    layout="wide",
)

st.markdown("""
<style>
.block-container {padding-top: 1.2rem; padding-bottom: 2rem;}
div[data-testid="stMetric"] {
    border: 1px solid rgba(49, 51, 63, 0.18);
    padding: 0.8rem;
    border-radius: 0.8rem;
    background: rgba(250,250,250,0.7);
}
.small-note {font-size: 0.92rem; color: #666;}
.card-box {
    white-space: pre-wrap;
    border: 1px solid rgba(49, 51, 63, 0.12);
    border-radius: 0.75rem;
    padding: 0.9rem;
    background: #fafafa;
    line-height: 1.65;
}
</style>
""", unsafe_allow_html=True)

st.title("🌱 Job Matrix Manager")
st.caption("本部 → 課 の職務記述書と、グレード別責任範囲を分けて管理する版です。")

try:
    SUPABASE_URL = st.secrets["SUPABASE_URL"].rstrip("/")
    SUPABASE_KEY = st.secrets["SUPABASE_KEY"]
except Exception:
    st.error("Streamlit secrets に SUPABASE_URL / SUPABASE_KEY が設定されていません。")
    st.stop()


def supabase_headers(prefer: Optional[str] = None) -> dict:
    headers = {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}",
        "Content-Type": "application/json",
    }
    if prefer:
        headers["Prefer"] = prefer
    return headers


def rest_get(path: str, params: Optional[dict] = None):
    response = requests.get(
        f"{SUPABASE_URL}/rest/v1/{path}",
        headers=supabase_headers(),
        params=params,
        timeout=30,
    )
    response.raise_for_status()
    return response.json()


def rest_post(path: str, payload: dict):
    response = requests.post(
        f"{SUPABASE_URL}/rest/v1/{path}",
        headers=supabase_headers(prefer="return=representation"),
        json=payload,
        timeout=30,
    )
    response.raise_for_status()
    return response.json()


GRADE_ORDER = ["G6", "G5B", "G5A", "G4", "G3", "G2"]

GRADE_NAMES = {
    "G6": "平社員",
    "G5B": "シニアスタッフ",
    "G5A": "上位シニアスタッフ",
    "G4": "スーパーバイザー",
    "G3": "課長",
    "G2": "次長",
}

HQ_STRUCTURE = {
    "管理本部": ["人事課", "総務課", "情報管理課", "経理課", "財務課"],
    "生産本部": ["エリア運営課", "設備・運搬課", "ナーサリー課", "マーケティング課", "プロセス管理課"],
    "コミュニケーション本部": ["インサイドコミュニケーション課", "アウトサイドコミュニケーション課"],
}

DEPARTMENT_DESCRIPTION_MASTER = {
    "人事課": "採用、労務管理、従業員情報管理、教育支援など、人材に関わる業務全般を担う。",
    "総務課": "庶務、文書管理、許認可・車両・備品・社内環境整備など、拠点運営を支える総務業務全般を担う。",
    "情報管理課": "財務情報、エリア情報、車両情報、災害事故情報などの整理・更新・活用支援を担う。",
    "経理課": "日常経理、会計記録、証憑管理、支払処理、月次締め対応など経理実務全般を担う。",
    "財務課": "資金繰り、予算実績管理、銀行対応、オンラインバンキング管理など財務業務全般を担う。",
    "エリア運営課": "各運営エリア・ロットの現場運営、進捗確認、人員・作業調整、現場課題対応を担う。",
    "設備・運搬課": "在庫管理、車両メンテナンス、調達・運搬など設備・物流運用全般を担う。",
    "ナーサリー課": "育苗、苗木品質管理、生産計画に基づく苗の準備などナーサリー業務全般を担う。",
    "マーケティング課": "折衝、契約管理、エリアチェック、植付計画に関する調整・資料整備・推進を担う。",
    "プロセス管理課": "工程・手順の標準化、メンテナンスプロセス、植付プロセスの改善と定着を担う。",
    "インサイドコミュニケーション課": "社内向けの情報共有、周知、社内連携の強化など内部コミュニケーションを担う。",
    "アウトサイドコミュニケーション課": "対外発信、渉外対応、地域・関係者との関係構築・連携強化を担う。",
}

DEPARTMENT_EVALUATION_MASTER = {
    "人事課": "正確性、対応期限遵守、情報管理、関係部署連携、採用・労務運用の安定性",
    "総務課": "正確性、対応スピード、段取り、社内サポート品質、管理精度",
    "情報管理課": "情報の正確性、更新頻度、活用しやすさ、管理ルール遵守、関係者連携",
    "経理課": "正確性、締め期限遵守、証憑管理、報告品質、改善への寄与",
    "財務課": "資金管理精度、期限遵守、分析力、金融機関対応、経営連携",
    "エリア運営課": "進捗管理、安全・品質管理、現場対応力、課題共有、目標達成度",
    "設備・運搬課": "安全性、稼働安定性、保守精度、運搬効率、問題発生時の対応力",
    "ナーサリー課": "品質維持、生産計画遵守、記録精度、改善提案、安全管理",
    "マーケティング課": "分析力、資料品質、関係者調整、推進力、目標達成度",
    "プロセス管理課": "標準化の浸透、改善効果、品質・安全への寄与、再現性、進捗管理",
    "インサイドコミュニケーション課": "情報浸透度、分かりやすさ、対応スピード、社内連携、施策推進力",
    "アウトサイドコミュニケーション課": "正確性、対外対応力、関係構築力、調整力、目標達成度",
}

GRADE_MASTER_DEFAULT = [
    {"grade": "G6", "grade_name": GRADE_NAMES["G6"], "positioning": "上位者の指示・方針のもとで担当業務を確実に遂行し、組織の一員として基本的な役割と責任を果たすランク。", "scope": "個人の担当業務の遂行責任", "keywords": "正確性 / 手順遵守 / 誠実さ / 協調性", "behavior_points": "担当業務に責任を持って取り組み、マナー・期限・約束を守り、周囲と協力して業務を遂行する。"},
    {"grade": "G5B", "grade_name": GRADE_NAMES["G5B"], "positioning": "G6からのステップアップ段階として、一定の自律性をもって業務を遂行するランク。", "scope": "担当領域の自律的な遂行責任", "keywords": "主体性 / 問題解決 / 柔軟対応 / 協力", "behavior_points": "必要な情報を取捨選択し、状況に応じた方法で実行しながら、組織成果に向けて自ら進んで協力する。"},
    {"grade": "G5A", "grade_name": GRADE_NAMES["G5A"], "positioning": "G5ランクの上位段階として、求められる職能基準を安定的に満たしているランク。", "scope": "複数業務の安定遂行と改善への責任", "keywords": "安定性 / 主体性 / 問題解決 / 周囲支援", "behavior_points": "G5Bと同じ基準を安定して発揮し、自分の役割を果たしながら周囲支援と改善提案にも取り組む。"},
    {"grade": "G4", "grade_name": GRADE_NAMES["G4"], "positioning": "組織運営に関与し、課規模の目標達成および人材・業務の管理責任を担うランク。", "scope": "課規模の成果責任 / 人・業務の管理責任", "keywords": "計画立案 / 人の管理 / 進捗管理 / 課題解決 / 浸透", "behavior_points": "月間目標に必要な施策を立案・展開し、メンバー管理、進捗管理、問題解決、方針浸透を担う。"},
    {"grade": "G3", "grade_name": GRADE_NAMES["G3"], "positioning": "組織運営に関与し、事業部規模の目標達成および人材・業務の管理責任を担うランク。", "scope": "事業部規模の成果責任 / 組織運営責任", "keywords": "年計画 / 部門統括 / 人材管理 / 横断連携 / 意思決定", "behavior_points": "年次計画の立案・展開、部門全体の管理、関係部署との調整、部下育成、課題解決を統括する。"},
    {"grade": "G2", "grade_name": GRADE_NAMES["G2"], "positioning": "組織運営に関与し、支店規模の目標達成および人材・業務の管理責任を担うランク。", "scope": "支店規模の成果責任 / 経営連携責任", "keywords": "中期計画 / 経営視点 / 全体最適 / 組織統括 / 横断推進", "behavior_points": "中期計画を踏まえた施策決定と展開、組織全体の管理、重要課題解決、経営と現場の接続を担う。"},
]

DEPARTMENT_COLUMNS = ["hq", "department", "job_title", "section_label", "role_summary", "evaluation_points"]
GRADE_COLUMNS = ["grade", "grade_name", "positioning", "scope", "keywords", "behavior_points"]


def generate_department_master() -> pd.DataFrame:
    rows = []
    for hq, departments in HQ_STRUCTURE.items():
        for department in departments:
            rows.append({
                "hq": hq,
                "department": department,
                "job_title": f"{department}担当",
                "section_label": department,
                "role_summary": DEPARTMENT_DESCRIPTION_MASTER.get(department, ""),
                "evaluation_points": DEPARTMENT_EVALUATION_MASTER.get(department, ""),
            })
    return pd.DataFrame(rows)


def generate_grade_master() -> pd.DataFrame:
    return pd.DataFrame(GRADE_MASTER_DEFAULT)


def safe_str(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def dataframe_to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")


def minimal_department_template_csv_bytes() -> bytes:
    return pd.DataFrame(columns=DEPARTMENT_COLUMNS).to_csv(index=False).encode("utf-8-sig")


def minimal_grade_template_csv_bytes() -> bytes:
    return pd.DataFrame(columns=GRADE_COLUMNS).to_csv(index=False).encode("utf-8-sig")


def save_snapshot(dept_df: pd.DataFrame, grade_df: pd.DataFrame, memo: str) -> None:
    payload = {
        "department_data": dept_df.to_dict(orient="records"),
        "grade_data": grade_df.to_dict(orient="records"),
        "memo": memo,
        "created_at": datetime.now().isoformat(),
    }
    rest_post("job_matrix_hq", payload)


def load_history() -> list:
    return rest_get("job_matrix_hq", params={"select": "*", "order": "created_at.desc"})


def validate_import_keys(base_df: pd.DataFrame, upload_df: pd.DataFrame, key_cols: list[str]) -> pd.DataFrame:
    base_keys = set(base_df[key_cols].astype(str).agg("||".join, axis=1))
    temp = upload_df.copy()
    temp["_merge_key"] = temp[key_cols].astype(str).agg("||".join, axis=1)
    invalid = temp[~temp["_merge_key"].isin(base_keys)].copy()
    return invalid.drop(columns=["_merge_key"])


def merge_partial_update(base_df: pd.DataFrame, upload_df: pd.DataFrame, key_cols: list[str], update_cols: list[str]):
    for col in key_cols:
        if col not in upload_df.columns:
            raise ValueError(f"CSVに必須列 '{col}' がありません。")

    base = base_df.copy()
    incoming = upload_df.copy()

    for col in key_cols:
        base[col] = base[col].astype(str).str.strip()
        incoming[col] = incoming[col].astype(str).str.strip()

    incoming = incoming.drop_duplicates(subset=key_cols, keep="last")
    base["_merge_key"] = base[key_cols].astype(str).agg("||".join, axis=1)
    incoming["_merge_key"] = incoming[key_cols].astype(str).agg("||".join, axis=1)
    incoming_map = incoming.set_index("_merge_key").to_dict(orient="index")

    changes = []
    for idx, row in base.iterrows():
        key = row["_merge_key"]
        if key not in incoming_map:
            continue
        source = incoming_map[key]
        for col in update_cols:
            if col not in incoming.columns:
                continue
            new_val = safe_str(source.get(col))
            old_val = safe_str(row.get(col))
            if new_val and new_val != old_val:
                base.at[idx, col] = new_val
                changes.append({"key": key, "column": col, "before": old_val, "after": new_val})

    base = base.drop(columns=["_merge_key"])
    return base, pd.DataFrame(changes)


def update_master_from_editor(master_df: pd.DataFrame, edited_df: pd.DataFrame, key_cols: list[str], update_cols: list[str]) -> pd.DataFrame:
    current = master_df.copy()
    edited = edited_df.copy()
    edited_map = edited.set_index(key_cols).to_dict(orient="index")

    for idx, row in current.iterrows():
        key = tuple(row[col] for col in key_cols)
        if key in edited_map:
            for col in update_cols:
                current.at[idx, col] = edited_map[key].get(col, row[col])
    return current


if "department_df" not in st.session_state:
    st.session_state.department_df = generate_department_master()
if "grade_df" not in st.session_state:
    st.session_state.grade_df = generate_grade_master()
if "last_change_log_department" not in st.session_state:
    st.session_state.last_change_log_department = pd.DataFrame()
if "last_change_log_grade" not in st.session_state:
    st.session_state.last_change_log_grade = pd.DataFrame()
if "connection_ok" not in st.session_state:
    try:
        _ = load_history()
        st.session_state.connection_ok = True
    except Exception as e:
        st.session_state.connection_ok = False
        st.session_state.connection_error = repr(e)

if not st.session_state.connection_ok:
    st.error(f"Supabase REST 接続失敗: {st.session_state.get('connection_error', 'Unknown error')}")
    st.info("Secrets、Supabaseテーブル、RLSポリシーを確認してください。")
    st.stop()

department_df = st.session_state.department_df.copy()
grade_df = st.session_state.grade_df.copy()

m1, m2, m3, m4 = st.columns(4)
m1.metric("本部数", len(HQ_STRUCTURE))
m2.metric("課数", len(department_df))
m3.metric("グレード数", len(grade_df))
m4.metric("保存方式", "本部+課 / グレード別")

st.markdown("## 🔎 フィルター")
c1, c2, c3 = st.columns([1.2, 1.2, 1.2])
with c1:
    hq_filter = st.multiselect("本部", options=list(HQ_STRUCTURE.keys()), default=[])
with c2:
    department_filter = st.multiselect("課", options=department_df["department"].tolist(), default=[])
with c3:
    keyword = st.text_input("キーワード検索", placeholder="説明、評価の観点、責任範囲など")

filtered_department_df = department_df.copy()
if hq_filter:
    filtered_department_df = filtered_department_df[filtered_department_df["hq"].isin(hq_filter)]
if department_filter:
    filtered_department_df = filtered_department_df[filtered_department_df["department"].isin(department_filter)]
if keyword:
    kw = keyword.lower()
    filtered_department_df = filtered_department_df[
        filtered_department_df.apply(lambda row: any(kw in str(v).lower() for v in row.values), axis=1)
    ]

filtered_grade_df = grade_df.copy()
if keyword:
    kw = keyword.lower()
    filtered_grade_df = filtered_grade_df[
        filtered_grade_df.apply(lambda row: any(kw in str(v).lower() for v in row.values), axis=1)
    ]

tab1, tab2, tab3, tab4, tab5 = st.tabs(["📊 本部・課一覧", "✏️ 編集", "📥📤 CSV", "📂 履歴", "🏷️ グレード定義"])


with tab1:
    st.subheader("本部 → 課 一覧")
    for hq in HQ_STRUCTURE.keys():
        sub = filtered_department_df[filtered_department_df["hq"] == hq].copy()
        if sub.empty:
            continue

        with st.expander(f"🏢 {hq}", expanded=True):
            for _, row in sub.iterrows():
                st.markdown(f"### {row['department']}")
                st.write(f"**職種**: {row['job_title']}")
                st.write(f"**係 / 表示名**: {row['section_label']}")
                st.write(f"**課の説明**: {row['role_summary']}")
                st.write(f"**評価の観点**: {row['evaluation_points']}")
                st.markdown("---")

    st.markdown("### 本部別サマリー")
    summary_rows = []
    for hq in HQ_STRUCTURE.keys():
        sub = department_df[department_df["hq"] == hq]
        summary_rows.append({"hq": hq, "department_count": len(sub)})
    st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)


with tab2:
    st.subheader("直接編集")
    edit_target = st.radio("編集対象", ["課マスタ", "グレードマスタ"], horizontal=True)

    if edit_target == "課マスタ":
        edited_department_df = st.data_editor(
            filtered_department_df,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            column_config={
                "hq": st.column_config.TextColumn("本部", disabled=True),
                "department": st.column_config.TextColumn("課", disabled=True),
                "job_title": st.column_config.TextColumn("職種"),
                "section_label": st.column_config.TextColumn("係 / 表示名"),
                "role_summary": st.column_config.TextColumn("課の説明", width="large"),
                "evaluation_points": st.column_config.TextColumn("評価の観点", width="large"),
            },
        )
        dept_memo = st.text_input("保存メモ（課マスタ）", placeholder="例: 本部構造に合わせて課の説明を更新")
        if st.button("💾 課マスタをSupabaseへ保存", type="primary"):
            new_department_df = update_master_from_editor(
                st.session_state.department_df,
                edited_department_df,
                key_cols=["hq", "department"],
                update_cols=["job_title", "section_label", "role_summary", "evaluation_points"],
            )
            st.session_state.department_df = new_department_df
            save_snapshot(st.session_state.department_df, st.session_state.grade_df, dept_memo or "department edit save")
            st.success("課マスタを保存しました。")

    else:
        edited_grade_df = st.data_editor(
            filtered_grade_df,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            column_config={
                "grade": st.column_config.TextColumn("グレード", disabled=True),
                "grade_name": st.column_config.TextColumn("グレード名"),
                "positioning": st.column_config.TextColumn("位置づけ", width="large"),
                "scope": st.column_config.TextColumn("責任範囲", width="medium"),
                "keywords": st.column_config.TextColumn("キーワード", width="medium"),
                "behavior_points": st.column_config.TextColumn("分かりやすい説明", width="large"),
            },
        )
        grade_memo = st.text_input("保存メモ（グレードマスタ）", placeholder="例: PDFに合わせて責任範囲を整理")
        if st.button("💾 グレードマスタをSupabaseへ保存", type="primary"):
            new_grade_df = update_master_from_editor(
                st.session_state.grade_df,
                edited_grade_df,
                key_cols=["grade"],
                update_cols=["grade_name", "positioning", "scope", "keywords", "behavior_points"],
            )
            st.session_state.grade_df = new_grade_df
            save_snapshot(st.session_state.department_df, st.session_state.grade_df, grade_memo or "grade edit save")
            st.success("グレードマスタを保存しました。")


with tab3:
    st.subheader("CSV 入出力")

    st.markdown("### 課マスタ")
    c1, c2 = st.columns(2)
    with c1:
        st.download_button("課マスタCSV出力", data=dataframe_to_csv_bytes(st.session_state.department_df), file_name="department_master_export.csv", mime="text/csv")
    with c2:
        st.download_button("課マスタ空テンプレートCSV出力", data=minimal_department_template_csv_bytes(), file_name="department_master_template.csv", mime="text/csv")

    uploaded_department_csv = st.file_uploader("課マスタ更新用CSV", type=["csv"], key="department_csv")
    if uploaded_department_csv is not None:
        try:
            import_department_df = pd.read_csv(uploaded_department_csv)
            st.dataframe(import_department_df, use_container_width=True, hide_index=True)
            if "hq" in import_department_df.columns and "department" in import_department_df.columns:
                invalid_department_df = validate_import_keys(st.session_state.department_df, import_department_df, key_cols=["hq", "department"])
                if not invalid_department_df.empty:
                    st.warning("課マスタに存在しない hq + department が含まれています。これらは更新されません。")
                    st.dataframe(invalid_department_df, use_container_width=True, hide_index=True)

            dept_csv_memo = st.text_input("課マスタCSV更新メモ", placeholder="例: 各課の説明を統一", key="dept_csv_memo")
            if st.button("課マスタCSVで部分更新", type="primary"):
                new_department_df, change_log = merge_partial_update(
                    st.session_state.department_df,
                    import_department_df,
                    key_cols=["hq", "department"],
                    update_cols=["job_title", "section_label", "role_summary", "evaluation_points"],
                )
                st.session_state.department_df = new_department_df
                st.session_state.last_change_log_department = change_log
                save_snapshot(st.session_state.department_df, st.session_state.grade_df, dept_csv_memo or "department csv update")
                st.success(f"課マスタ更新完了: {len(change_log)} 件")
        except Exception as e:
            st.error(f"課マスタCSV取込エラー: {e}")

    st.markdown("---")
    st.markdown("### グレードマスタ")
    c3, c4 = st.columns(2)
    with c3:
        st.download_button("グレードマスタCSV出力", data=dataframe_to_csv_bytes(st.session_state.grade_df), file_name="grade_master_export.csv", mime="text/csv")
    with c4:
        st.download_button("グレードマスタ空テンプレートCSV出力", data=minimal_grade_template_csv_bytes(), file_name="grade_master_template.csv", mime="text/csv")

    uploaded_grade_csv = st.file_uploader("グレードマスタ更新用CSV", type=["csv"], key="grade_csv")
    if uploaded_grade_csv is not None:
        try:
            import_grade_df = pd.read_csv(uploaded_grade_csv)
            st.dataframe(import_grade_df, use_container_width=True, hide_index=True)
            if "grade" in import_grade_df.columns:
                invalid_grade_df = validate_import_keys(st.session_state.grade_df, import_grade_df, key_cols=["grade"])
                if not invalid_grade_df.empty:
                    st.warning("グレードマスタに存在しない grade が含まれています。これらは更新されません。")
                    st.dataframe(invalid_grade_df, use_container_width=True, hide_index=True)

            grade_csv_memo = st.text_input("グレードマスタCSV更新メモ", placeholder="例: PDFに沿って文言調整", key="grade_csv_memo")
            if st.button("グレードマスタCSVで部分更新", type="primary"):
                new_grade_df, change_log = merge_partial_update(
                    st.session_state.grade_df,
                    import_grade_df,
                    key_cols=["grade"],
                    update_cols=["grade_name", "positioning", "scope", "keywords", "behavior_points"],
                )
                new_grade_df["grade_sort"] = new_grade_df["grade"].map({g: i for i, g in enumerate(GRADE_ORDER)})
                new_grade_df = new_grade_df.sort_values("grade_sort").drop(columns=["grade_sort"])
                st.session_state.grade_df = new_grade_df
                st.session_state.last_change_log_grade = change_log
                save_snapshot(st.session_state.department_df, st.session_state.grade_df, grade_csv_memo or "grade csv update")
                st.success(f"グレードマスタ更新完了: {len(change_log)} 件")
        except Exception as e:
            st.error(f"グレードマスタCSV取込エラー: {e}")

    st.markdown("---")
    st.markdown("### 最終更新ログ")
    if not st.session_state.last_change_log_department.empty:
        st.markdown("#### 課マスタ")
        st.dataframe(st.session_state.last_change_log_department, use_container_width=True, hide_index=True)
    if not st.session_state.last_change_log_grade.empty:
        st.markdown("#### グレードマスタ")
        st.dataframe(st.session_state.last_change_log_grade, use_container_width=True, hide_index=True)
    if st.session_state.last_change_log_department.empty and st.session_state.last_change_log_grade.empty:
        st.info("まだ更新ログはありません。")


with tab4:
    st.subheader("保存履歴")
    try:
        history = load_history()
    except Exception as e:
        st.error(f"履歴取得失敗: {e}")
        history = []

    if not history:
        st.info("保存履歴はまだありません。")
    else:
        for i, item in enumerate(history):
            created_at = item.get("created_at", "")
            memo = item.get("memo", "")
            hist_department_df = pd.DataFrame(item.get("department_data", []))
            hist_grade_df = pd.DataFrame(item.get("grade_data", []))

            with st.expander(f"{created_at} | {memo}", expanded=False):
                st.markdown("#### 課マスタ")
                if hist_department_df.empty:
                    st.info("課マスタ履歴なし")
                else:
                    st.dataframe(hist_department_df, use_container_width=True, hide_index=True)

                st.markdown("#### グレードマスタ")
                if hist_grade_df.empty:
                    st.info("グレードマスタ履歴なし")
                else:
                    hist_grade_df["grade_sort"] = hist_grade_df["grade"].map({g: j for j, g in enumerate(GRADE_ORDER)})
                    hist_grade_df = hist_grade_df.sort_values("grade_sort").drop(columns=["grade_sort"])
                    st.dataframe(hist_grade_df, use_container_width=True, hide_index=True)

                if st.button(f"この履歴を復元 #{i}", key=f"restore_{i}"):
                    if not hist_department_df.empty:
                        st.session_state.department_df = hist_department_df
                    if not hist_grade_df.empty:
                        st.session_state.grade_df = hist_grade_df
                    st.success("復元しました。必要なら再保存してください。")


with tab5:
    st.subheader("グレード定義")
    sorted_grade_df = st.session_state.grade_df.copy()
    sorted_grade_df["grade_sort"] = sorted_grade_df["grade"].map({g: i for i, g in enumerate(GRADE_ORDER)})
    sorted_grade_df = sorted_grade_df.sort_values("grade_sort").drop(columns=["grade_sort"])

    for _, row in sorted_grade_df.iterrows():
        with st.expander(f"{row['grade']} / {row['grade_name']}", expanded=True):
            st.write(f"**位置づけ**: {row['positioning']}")
            st.write(f"**責任範囲**: {row['scope']}")
            st.write(f"**キーワード**: {row['keywords']}")
            st.markdown(f"<div class='card-box'>{row['behavior_points']}</div>", unsafe_allow_html=True)

    st.markdown("---")
    if st.button("課マスタを初期状態に戻す"):
        st.session_state.department_df = generate_department_master()
        st.success("課マスタを初期状態に戻しました。")

    if st.button("グレードマスタを初期状態に戻す"):
        st.session_state.grade_df = generate_grade_master()
        st.success("グレードマスタを初期状態に戻しました。")

    st.markdown("---")
    st.markdown("### Supabaseテーブル想定")
    st.code(
        "table name: job_matrix_hq\n"
        "columns:\n"
        "- id (bigint, identity, primary key)\n"
        "- department_data (jsonb)\n"
        "- grade_data (jsonb)\n"
        "- memo (text)\n"
        "- created_at (timestamptz)\n",
        language="text",
    )

    st.markdown("### Streamlit Secrets 例")
    st.code('SUPABASE_URL = "https://xxxx.supabase.co"\nSUPABASE_KEY = "anon public key"', language="toml")

    st.markdown("### 補足")
    st.markdown("""
- 各課については **グレード別の詳細説明を廃止** し、**課ごとに1つの説明** に統一しています。  
- グレードの考え方は **別タブ** に分離しています。  
- CSVは **部分更新** です。空欄セルは既存値を保持します。  
- 保存は **Supabase REST API** を使っているため、`supabase-py` の proxy 問題を回避できます。  
- RLSをONにする場合は、`SELECT` と `INSERT` のポリシーを設定してください。  
""")
