
import io
from datetime import datetime
from typing import Optional

import pandas as pd
import requests
import streamlit as st

st.set_page_config(
    page_title="Job Matrix Manager V2",
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
    background: rgba(250,250,250,0.6);
}
.small-note {font-size: 0.9rem; color: #666;}
.task-box {
    white-space: pre-wrap;
    border: 1px solid rgba(49, 51, 63, 0.12);
    border-radius: 0.75rem;
    padding: 0.85rem;
    background: #fafafa;
    line-height: 1.65;
}
</style>
""", unsafe_allow_html=True)

st.title("🌱 Job Matrix Manager V2")
st.caption("部門 × グレードごとの仕事内容・責任範囲・KPI・具体業務を管理し、CSVで部分更新、Supabaseへ履歴保存できます。")

# =========================
# SECRETS / REST SETTINGS
# =========================
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


GRADES = {
    "G6": "平社員",
    "G5B": "シニアスタッフ",
    "G5A": "上位シニアスタッフ",
    "G4": "スーパーバイザー",
    "G3": "課長",
    "G2": "次長",
}

GRADE_ORDER = ["G6", "G5B", "G5A", "G4", "G3", "G2"]

DEPARTMENTS = [
    "人事課",
    "総務課",
    "経理課",
    "財務課",
    "情報管理課",
    "エリア運営課",
    "設備・運搬課",
    "ナーサリー課",
    "マーケティング課",
    "プロセス管理課",
    "インサイドコミュニケーション課",
    "アウトサイドコミュニケーション課",
]

UPDATE_COLUMNS = [
    "grade_name",
    "role_summary",
    "responsibility",
    "kpi",
    "task_details",
]

DEPARTMENT_DESCRIPTIONS = {
    "人事課": "採用・労務・人材配置・教育支援を担う。",
    "総務課": "拠点運営を支える庶務・文書・社内環境整備を担う。",
    "経理課": "日常経理、締め処理、会計記録を担う。",
    "財務課": "資金繰り、予算管理、金融機関対応を担う。",
    "情報管理課": "システム運用、データ整備、情報活用を担う。",
    "エリア運営課": "農地・現場の運営管理、作業調整、巡回確認を担う。",
    "設備・運搬課": "設備保守、資材運搬、車両・物流運用を担う。",
    "ナーサリー課": "苗木育成、品質管理、植栽前の育苗管理を担う。",
    "マーケティング課": "市場情報整理、外部向け資料、事業訴求支援を担う。",
    "プロセス管理課": "業務手順、工程、品質、安全の標準化を担う。",
    "インサイドコミュニケーション課": "社内周知、社内文化、組織浸透施策を担う。",
    "アウトサイドコミュニケーション課": "対外発信、渉外、地域・関係者連携を担う。",
}


def make_default_role_summary(department: str, grade: str, grade_name: str) -> str:
    data = {
        "人事課": {
            "G6": "採用・勤怠・人事データ入力などの定型業務を担当する。",
            "G5B": "採用実務や従業員管理業務を安定運用し、定型業務の精度向上を担う。",
            "G5A": "採用、労務、教育関連業務を横断的に支援し、改善提案も行う。",
            "G4": "人事オペレーションを監督し、メンバー指導と進捗管理を担う。",
            "G3": "人事課全体の計画立案、採用・配置・制度運用を統括する。",
            "G2": "人事戦略、組織体制整備、経営方針に基づく人材マネジメントを推進する。",
        },
        "総務課": {
            "G6": "備品管理、文書管理、来客対応などの総務実務を担当する。",
            "G5B": "庶務業務の安定運用と社内サポートを担う。",
            "G5A": "総務実務全般を理解し、業務改善や他部門連携を支援する。",
            "G4": "総務チームの進捗管理と日常運営の監督を担う。",
            "G3": "総務機能全体の運用管理とルール整備を統括する。",
            "G2": "拠点運営基盤の整備、ガバナンス強化、管理部門連携を推進する。",
        },
        "経理課": {
            "G6": "伝票入力、証憑整理、支払データ作成などの基礎経理業務を担当する。",
            "G5B": "日次・月次経理の実務を安定運用し、正確性を担保する。",
            "G5A": "経理実務全般に加え、締め処理や照合作業の改善も担う。",
            "G4": "経理業務の進捗管理、メンバー確認、締め処理管理を担う。",
            "G3": "月次・年次決算、会計管理、税務対応を統括する。",
            "G2": "財務・経理方針を踏まえた管理体制強化と経営報告を推進する。",
        },
        "財務課": {
            "G6": "資金関連データの整理や支払準備等の補助業務を担当する。",
            "G5B": "資金繰り関連資料の作成や銀行対応補助を担う。",
            "G5A": "財務管理資料作成、資金計画支援、数値分析を担う。",
            "G4": "資金管理業務の監督、実務進捗管理、財務報告補助を担う。",
            "G3": "資金繰り、予算実績管理、金融機関対応を統括する。",
            "G2": "財務戦略、投資判断支援、経営数値管理を推進する。",
        },
        "情報管理課": {
            "G6": "データ入力、ファイル管理、IT機器運用補助を担当する。",
            "G5B": "システム運用補助、情報整備、ユーザー対応を担う。",
            "G5A": "業務データ管理、IT活用支援、簡易改善提案を担う。",
            "G4": "情報管理実務の監督、システム運用管理、メンバー支援を担う。",
            "G3": "情報管理体制、業務システム運用、データ活用推進を統括する。",
            "G2": "情報戦略、デジタル化推進、情報統制基盤の整備を担う。",
        },
        "エリア運営課": {
            "G6": "現場オペレーション、巡回、作業記録などの基本業務を担当する。",
            "G5B": "担当エリアの運営実務を安定的に遂行し、現場課題を共有する。",
            "G5A": "現場の進捗確認、作業調整、改善提案を担う。",
            "G4": "現場チームの管理、日々の進捗監督、安全・品質確認を担う。",
            "G3": "複数エリアの運営管理、作業計画、人員配置を統括する。",
            "G2": "農地運営方針、拠点横断の運営最適化、事業計画達成を推進する。",
        },
        "設備・運搬課": {
            "G6": "設備点検補助、資材運搬、車両運用補助を担当する。",
            "G5B": "設備・車両の運用実務と日常点検を担う。",
            "G5A": "設備保全、運搬計画支援、トラブル初期対応を担う。",
            "G4": "設備保守・運搬業務の監督、安全管理、進捗管理を担う。",
            "G3": "設備管理計画、保全計画、運搬体制を統括する。",
            "G2": "設備投資方針、物流効率化、全体インフラ最適化を推進する。",
        },
        "ナーサリー課": {
            "G6": "苗の育成、灌水、除草、記録などの基礎栽培業務を担当する。",
            "G5B": "苗木育成業務を安定運用し、育成状況の記録管理を担う。",
            "G5A": "育苗計画支援、品質確認、作業改善提案を担う。",
            "G4": "ナーサリー現場の監督、人員配置、品質・安全管理を担う。",
            "G3": "育苗計画、数量管理、現場運営全体を統括する。",
            "G2": "育苗戦略、生産性向上、植栽計画との連携を推進する。",
        },
        "マーケティング課": {
            "G6": "市場情報整理、資料作成補助、基礎調査業務を担当する。",
            "G5B": "市場調査、資料更新、営業支援情報の整理を担う。",
            "G5A": "調査分析、社外資料作成、施策提案支援を担う。",
            "G4": "調査進捗管理、メンバー支援、施策実行管理を担う。",
            "G3": "市場分析、販売戦略支援、対外訴求計画を統括する。",
            "G2": "事業戦略と連動した市場開拓・発信方針を推進する。",
        },
        "プロセス管理課": {
            "G6": "業務記録、手順遵守、現場データ収集などを担当する。",
            "G5B": "工程管理補助、手順管理、定型レポート作成を担う。",
            "G5A": "工程改善、標準化支援、異常対応補助を担う。",
            "G4": "工程進捗・品質・安全の監督と現場是正を担う。",
            "G3": "業務プロセス設計、改善推進、管理指標運用を統括する。",
            "G2": "全体プロセス最適化、標準化方針、継続改善を推進する。",
        },
        "インサイドコミュニケーション課": {
            "G6": "社内連絡、情報配信補助、文書整理を担当する。",
            "G5B": "社内周知、イベント運営補助、情報整理を担う。",
            "G5A": "社内施策支援、社内発信、エンゲージメント向上施策を担う。",
            "G4": "社内コミュニケーション施策の進捗管理と運営を担う。",
            "G3": "社内広報、社内文化醸成、情報浸透施策を統括する。",
            "G2": "組織活性化、経営メッセージ浸透、社内連携強化を推進する。",
        },
        "アウトサイドコミュニケーション課": {
            "G6": "対外資料整理、問い合わせ対応補助、情報更新を担当する。",
            "G5B": "社外連絡、情報発信補助、対外資料管理を担う。",
            "G5A": "社外発信支援、関係者調整、広報実務を担う。",
            "G4": "対外コミュニケーション施策の運営管理を担う。",
            "G3": "広報・渉外・地域連携対応を統括する。",
            "G2": "対外関係戦略、地域・行政・関係者との信頼構築を推進する。",
        },
    }
    if department in data and grade in data[department]:
        return data[department][grade]
    return f"{grade_name}として{department}の業務を担当する。"


def make_default_responsibility(grade: str) -> str:
    return {
        "G6": "担当業務の正確な遂行、報告、手順遵守",
        "G5B": "担当領域の安定運用、改善提案、後輩支援",
        "G5A": "複数業務の遂行、改善提案、関係部署連携",
        "G4": "チーム管理、進捗管理、品質・安全管理、人材育成",
        "G3": "課全体の運営、計画立案、課題解決、予算・人員管理",
        "G2": "部門戦略推進、横断管理、経営連携、組織最適化",
    }[grade]


def make_default_kpi(grade: str) -> str:
    return {
        "G6": "業務達成率、処理件数、正確性、期限遵守率",
        "G5B": "担当業務達成率、品質、改善件数、納期遵守率",
        "G5A": "成果達成率、改善提案数、関連部署連携品質",
        "G4": "チーム達成率、進捗遵守率、品質指標、人材育成状況",
        "G3": "課目標達成率、予算遵守率、改善成果、組織運営状況",
        "G2": "部門目標達成率、戦略進捗、横断課題解決、収益・効率改善",
    }[grade]


def make_default_task_details(department: str, grade: str) -> str:
    # 改行区切りで10件前後
    tasks = {
        ("人事課", "G6"): [
            "応募者情報の入力と更新",
            "面接日程の調整",
            "応募者への連絡対応",
            "求人媒体の掲載情報更新",
            "応募書類の整理・保管",
            "入社書類の準備",
            "勤怠データの確認補助",
            "従業員情報の台帳更新",
            "採用進捗表の更新",
            "社内向け人事資料の作成補助",
        ],
        ("人事課", "G4"): [
            "採用進捗の管理",
            "メンバーの業務割当と確認",
            "労務手続きの進捗確認",
            "教育計画の運用支援",
            "人事データの精度確認",
            "採用課題の整理",
            "現場部門との採用調整",
            "定型業務の標準化",
            "人事関連レポートの取りまとめ",
            "後輩育成と指導",
        ],
        ("経理課", "G6"): [
            "請求書・領収書の整理",
            "伝票入力",
            "支払データの準備",
            "証憑のファイリング",
            "現金・小口精算補助",
            "月次資料の作成補助",
            "残高照合の補助",
            "経費申請内容の確認補助",
            "取引データの一覧更新",
            "会計関連書類の保管",
        ],
        ("ナーサリー課", "G6"): [
            "苗木の水やり",
            "育苗エリアの清掃",
            "苗木の状態確認",
            "育成記録の入力",
            "資材の補充",
            "苗床の整備",
            "病害虫の初期確認",
            "除草作業",
            "苗木の移動補助",
            "日報の作成",
        ],
        ("エリア運営課", "G6"): [
            "担当エリアの巡回",
            "作業進捗の確認",
            "作業記録の入力",
            "現場写真の整理",
            "資材使用状況の確認",
            "現場メンバーへの伝達",
            "異常箇所の報告",
            "簡易測定・確認作業",
            "日報提出",
            "作業準備と後片付け",
        ],
        ("設備・運搬課", "G6"): [
            "設備の日常点検補助",
            "運搬車両の点検補助",
            "資材の積み下ろし",
            "保守記録の入力",
            "工具の整理",
            "簡易メンテナンス補助",
            "運搬スケジュールの確認",
            "消耗品の補充",
            "異常の一次報告",
            "作業エリアの安全確認",
        ],
    }
    selected = tasks.get((department, grade))
    if not selected:
        base_map = {
            "G6": [
                "日常定型業務の実施",
                "データ入力・更新",
                "書類整理・保管",
                "関係者への連絡補助",
                "進捗表の更新",
                "担当業務の日報作成",
                "ルールに沿った記録管理",
                "簡易チェック業務",
                "上長への報告",
                "改善提案の共有",
            ],
            "G5B": [
                "担当業務の安定運用",
                "進捗確認",
                "データ精度の確認",
                "関連書類のレビュー補助",
                "後輩への作業説明",
                "業務改善点の整理",
                "担当領域の課題共有",
                "定例レポート作成",
                "関係部署との連携",
                "上長への報告・相談",
            ],
            "G5A": [
                "複数業務の実行管理",
                "改善提案の作成",
                "関係部署との調整",
                "業務フローの見直し",
                "異常時の一次対応",
                "実績データの分析",
                "手順書更新の支援",
                "後輩育成",
                "定例報告資料の作成",
                "進捗・品質の確認",
            ],
            "G4": [
                "チームメンバーの業務割当",
                "日次進捗の確認",
                "品質状況の確認",
                "安全確認",
                "課題の整理と是正指示",
                "メンバー指導",
                "部署内レポートの取りまとめ",
                "業務フローの標準化",
                "他部署との調整",
                "上長への報告",
            ],
            "G3": [
                "課全体の計画立案",
                "人員配置の検討",
                "予算進捗の確認",
                "課題管理",
                "改善テーマの推進",
                "他部署との折衝",
                "レポートレビュー",
                "月次目標の確認",
                "部下育成方針の策定",
                "マネジメント報告",
            ],
            "G2": [
                "部門戦略の策定支援",
                "横断課題の推進",
                "重要案件の意思決定支援",
                "事業計画との整合確認",
                "KPIレビュー",
                "予算・リソース最適化",
                "拠点間連携の推進",
                "組織課題の整理",
                "管理体制の強化",
                "経営層への報告",
            ],
        }
        selected = base_map[grade]
    return "\n".join(selected)


def generate_default() -> pd.DataFrame:
    rows = []
    for dept in DEPARTMENTS:
        for grade in GRADE_ORDER:
            rows.append({
                "department": dept,
                "grade": grade,
                "grade_name": GRADES[grade],
                "role_summary": make_default_role_summary(dept, grade, GRADES[grade]),
                "responsibility": make_default_responsibility(grade),
                "kpi": make_default_kpi(grade),
                "task_details": make_default_task_details(dept, grade),
            })
    return pd.DataFrame(rows)


def ensure_task_details_column(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    if "task_details" not in out.columns:
        out["task_details"] = ""
    return out


def save_snapshot(df: pd.DataFrame, memo: str) -> None:
    payload = {
        "data": df.to_dict(orient="records"),
        "memo": memo,
        "created_at": datetime.now().isoformat(),
    }
    rest_post("job_matrix", payload)


def load_history() -> list:
    return rest_get("job_matrix", params={"select": "*", "order": "created_at.desc"})


def normalize_text(value):
    if pd.isna(value):
        return None
    text = str(value).strip()
    return text if text != "" else None


def dataframe_to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")


def dataframe_to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    export_df = ensure_task_details_column(df.copy())
    export_df["grade_sort"] = export_df["grade"].map({g: i for i, g in enumerate(GRADE_ORDER)})
    export_df = export_df.sort_values(["department", "grade_sort"]).drop(columns=["grade_sort"])
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="job_matrix")
    output.seek(0)
    return output.getvalue()


def minimal_template_csv_bytes() -> bytes:
    template_df = pd.DataFrame(columns=[
        "department", "grade", "grade_name", "role_summary", "responsibility", "kpi", "task_details"
    ])
    return template_df.to_csv(index=False).encode("utf-8-sig")


def count_tasks(text) -> int:
    if pd.isna(text) or str(text).strip() == "":
        return 0
    return len([line.strip() for line in str(text).splitlines() if line.strip()])


def preview_tasks(text: str) -> str:
    if pd.isna(text) or str(text).strip() == "":
        return "具体業務未入力"
    lines = [line.strip() for line in str(text).splitlines() if line.strip()]
    return "\n".join([f"{i+1}. {line}" for i, line in enumerate(lines)])


def validate_import_keys(base_df: pd.DataFrame, upload_df: pd.DataFrame) -> pd.DataFrame:
    base_keys = set(
        base_df["department"].astype(str).str.strip() + "||" + base_df["grade"].astype(str).str.strip()
    )
    up = upload_df.copy()
    up["_merge_key"] = up["department"].astype(str).str.strip() + "||" + up["grade"].astype(str).str.strip()
    invalid = up[~up["_merge_key"].isin(base_keys)].copy()
    return invalid.drop(columns=["_merge_key"])


def merge_partial_update(base_df: pd.DataFrame, upload_df: pd.DataFrame):
    required_keys = ["department", "grade"]
    for col in required_keys:
        if col not in upload_df.columns:
            raise ValueError(f"CSVに必須列 '{col}' がありません。")

    updated_df = ensure_task_details_column(base_df.copy())
    import_df = upload_df.copy()
    if "task_details" not in import_df.columns:
        import_df["task_details"] = None

    for col in required_keys:
        updated_df[col] = updated_df[col].astype(str).str.strip()
        import_df[col] = import_df[col].astype(str).str.strip()

    import_df = import_df.drop_duplicates(subset=required_keys, keep="last")

    updated_df["_merge_key"] = updated_df["department"] + "||" + updated_df["grade"]
    import_df["_merge_key"] = import_df["department"] + "||" + import_df["grade"]
    import_map = import_df.set_index("_merge_key").to_dict(orient="index")

    changes = []
    for idx, row in updated_df.iterrows():
        key = row["_merge_key"]
        if key not in import_map:
            continue
        source = import_map[key]
        for col in UPDATE_COLUMNS:
            if col not in import_df.columns:
                continue
            old_val = normalize_text(row.get(col))
            new_val = normalize_text(source.get(col))
            if new_val is not None and new_val != old_val:
                updated_df.at[idx, col] = new_val
                changes.append({
                    "department": row["department"],
                    "grade": row["grade"],
                    "column": col,
                    "before": old_val,
                    "after": new_val,
                })

    updated_df = updated_df.drop(columns=["_merge_key"])
    log_df = pd.DataFrame(changes)
    return updated_df, log_df


def update_master_from_editor(master_df: pd.DataFrame, edited_df: pd.DataFrame) -> pd.DataFrame:
    current = ensure_task_details_column(master_df.copy())
    edited_df = ensure_task_details_column(edited_df.copy())
    edited_map = edited_df.set_index(["department", "grade"]).to_dict(orient="index")
    for idx, row in current.iterrows():
        key = (row["department"], row["grade"])
        if key in edited_map:
            for col in UPDATE_COLUMNS:
                current.at[idx, col] = edited_map[key].get(col, row[col])
    return current


def department_summary(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    df = ensure_task_details_column(df.copy())
    for dept in DEPARTMENTS:
        sub = df[df["department"] == dept]
        rows.append({
            "department": dept,
            "rows": len(sub),
            "filled_role_summary": int(sub["role_summary"].astype(str).str.strip().ne("").sum()),
            "filled_task_details": int(sub["task_details"].astype(str).str.strip().ne("").sum()),
            "task_count_total": int(sub["task_details"].apply(count_tasks).sum()),
        })
    return pd.DataFrame(rows)


# =========================
# SESSION
# =========================
if "df" not in st.session_state:
    st.session_state.df = generate_default()

st.session_state.df = ensure_task_details_column(st.session_state.df)

if "last_change_log" not in st.session_state:
    st.session_state.last_change_log = pd.DataFrame()

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

master_df = ensure_task_details_column(st.session_state.df.copy())
summary_df = department_summary(master_df)

m1, m2, m3, m4 = st.columns(4)
m1.metric("部門数", len(DEPARTMENTS))
m2.metric("グレード数", len(GRADE_ORDER))
m3.metric("管理行数", len(master_df))
m4.metric("総タスク数", int(master_df["task_details"].apply(count_tasks).sum()))

st.markdown("## 🔎 フィルター")
c1, c2, c3 = st.columns([1.2, 1, 1.2])
with c1:
    dept_filter = st.multiselect("部門", options=DEPARTMENTS, default=[])
with c2:
    grade_filter = st.multiselect("グレード", options=GRADE_ORDER, default=[])
with c3:
    keyword = st.text_input("キーワード検索", placeholder="仕事内容・責任範囲・KPI・具体業務など")

filtered_df = master_df.copy()
if dept_filter:
    filtered_df = filtered_df[filtered_df["department"].isin(dept_filter)]
if grade_filter:
    filtered_df = filtered_df[filtered_df["grade"].isin(grade_filter)]
if keyword:
    kw = keyword.lower()
    filtered_df = filtered_df[
        filtered_df.apply(lambda row: any(kw in str(v).lower() for v in row.values), axis=1)
    ]

filtered_df["task_count"] = filtered_df["task_details"].apply(count_tasks)
filtered_df["grade_sort"] = filtered_df["grade"].map({g: i for i, g in enumerate(GRADE_ORDER)})
filtered_df = filtered_df.sort_values(["department", "grade_sort"]).drop(columns=["grade_sort"])
st.caption(f"表示件数: {len(filtered_df)} / 全 {len(master_df)} 件")

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 一覧",
    "✏️ 編集",
    "📥📤 CSV/Excel",
    "📂 履歴",
    "⚙️ 設定",
])

with tab1:
    st.subheader("一覧ビュー")
    view_mode = st.radio("表示形式", ["部門ごとの展開表示", "全件テーブル"], horizontal=True)

    if view_mode == "部門ごとの展開表示":
        for dept in DEPARTMENTS:
            dept_df = filtered_df[filtered_df["department"] == dept].copy()
            if dept_df.empty:
                continue
            with st.expander(f"🏢 {dept}", expanded=False):
                st.markdown(f"<div class='small-note'>{DEPARTMENT_DESCRIPTIONS.get(dept, '')}</div>", unsafe_allow_html=True)
                for _, row in dept_df.iterrows():
                    st.markdown(f"### {row['grade']} / {row['grade_name']}")
                    st.write(f"**仕事内容概要**: {row['role_summary']}")
                    st.write(f"**責任範囲**: {row['responsibility']}")
                    st.write(f"**KPI**: {row['kpi']}")
                    st.write(f"**具体業務数**: {count_tasks(row['task_details'])}")
                    st.markdown(
                        f"<div class='task-box'>{preview_tasks(row['task_details'])}</div>",
                        unsafe_allow_html=True
                    )
                    st.markdown("---")
    else:
        table_df = filtered_df.copy()
        st.dataframe(table_df, use_container_width=True, hide_index=True)

    st.markdown("### 部門別サマリー")
    st.dataframe(summary_df, use_container_width=True, hide_index=True)

with tab2:
    st.subheader("直接編集")
    st.write("`task_details` は改行区切りで入力してください。1行が1業務になります。")

    edited_df = st.data_editor(
        filtered_df,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        column_config={
            "department": st.column_config.TextColumn("部門", disabled=True),
            "grade": st.column_config.TextColumn("グレード", disabled=True),
            "grade_name": st.column_config.TextColumn("グレード名"),
            "role_summary": st.column_config.TextColumn("仕事内容概要", width="medium"),
            "responsibility": st.column_config.TextColumn("責任範囲", width="medium"),
            "kpi": st.column_config.TextColumn("KPI", width="medium"),
            "task_details": st.column_config.TextColumn("具体業務（改行区切り）", width="large"),
            "task_count": st.column_config.NumberColumn("業務数", disabled=True),
        }
    )

    memo = st.text_input("保存メモ", placeholder="例: G6人事の具体業務を詳細化")
    if st.button("💾 編集内容をSupabaseへ保存", type="primary"):
        new_master = update_master_from_editor(st.session_state.df, edited_df.drop(columns=["task_count"], errors="ignore"))
        st.session_state.df = ensure_task_details_column(new_master)
        save_snapshot(st.session_state.df, memo or "manual edit save")
        st.success("保存完了")

with tab3:
    st.subheader("CSV / Excel 入出力")

    dl1, dl2, dl3 = st.columns(3)
    with dl1:
        st.download_button(
            "現在データをCSV出力",
            data=dataframe_to_csv_bytes(st.session_state.df),
            file_name="job_matrix_v2_export.csv",
            mime="text/csv",
        )
    with dl2:
        st.download_button(
            "現在データをExcel出力",
            data=dataframe_to_excel_bytes(st.session_state.df),
            file_name="job_matrix_v2_export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with dl3:
        st.download_button(
            "空テンプレートCSV出力",
            data=minimal_template_csv_bytes(),
            file_name="job_matrix_v2_template.csv",
            mime="text/csv",
        )

    st.markdown("---")
    st.markdown("### CSV部分更新")
    st.write("`department + grade` が一致する行だけ対象です。")
    st.write("CSVで値が入っているセルだけ更新し、空欄は既存値を保持します。")
    st.write("`task_details` は改行区切りで複数業務を入れてください。")

    uploaded_csv = st.file_uploader("更新用CSVをアップロード", type=["csv"])
    if uploaded_csv is not None:
        try:
            import_df = pd.read_csv(uploaded_csv)
            st.markdown("#### アップロード内容")
            st.dataframe(import_df, use_container_width=True, hide_index=True)

            if "department" in import_df.columns and "grade" in import_df.columns:
                invalid_df = validate_import_keys(st.session_state.df, import_df)
                if not invalid_df.empty:
                    st.warning("マスタに存在しない department + grade が含まれています。これらは更新されません。")
                    st.dataframe(invalid_df, use_container_width=True, hide_index=True)
                else:
                    st.success("すべてのキーが既存マスタに存在しています。")

            csv_memo = st.text_input("CSV更新メモ", placeholder="例: 部門長レビューで具体業務を追加")
            if st.button("CSV内容で部分更新する", type="primary"):
                new_df, change_log = merge_partial_update(st.session_state.df, import_df)
                st.session_state.df = ensure_task_details_column(new_df)
                st.session_state.last_change_log = change_log
                save_snapshot(st.session_state.df, csv_memo or f"csv partial update ({len(change_log)} changes)")
                st.success(f"更新完了: {len(change_log)} 件")

        except Exception as e:
            st.error(f"CSV取込エラー: {e}")

    st.markdown("### 最終更新ログ")
    if not st.session_state.last_change_log.empty:
        st.dataframe(st.session_state.last_change_log, use_container_width=True, hide_index=True)
    else:
        st.info("まだCSV更新ログはありません。")

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
        st.write("履歴を開くと、その時点のスナップショットを確認できます。")
        for i, item in enumerate(history):
            created_at = item.get("created_at", "")
            memo = item.get("memo", "")
            hist_df = ensure_task_details_column(pd.DataFrame(item.get("data", [])))
            with st.expander(f"{created_at} | {memo}", expanded=False):
                if hist_df.empty:
                    st.warning("履歴データが空です。")
                else:
                    hist_df["task_count"] = hist_df["task_details"].apply(count_tasks)
                    hist_df["grade_sort"] = hist_df["grade"].map({g: j for j, g in enumerate(GRADE_ORDER)})
                    hist_df = hist_df.sort_values(["department", "grade_sort"]).drop(columns=["grade_sort"])
                    st.dataframe(hist_df, use_container_width=True, hide_index=True)
                    if st.button(f"この履歴を復元 #{i}", key=f"restore_{i}"):
                        st.session_state.df = ensure_task_details_column(hist_df.drop(columns=["task_count"], errors="ignore"))
                        st.success("復元しました。必要なら編集タブ等で再保存してください。")

with tab5:
    st.subheader("設定 / 初期化")
    st.write("必要に応じて初期データへ戻せます。")
    if st.button("初期データに戻す"):
        st.session_state.df = generate_default()
        st.session_state.last_change_log = pd.DataFrame()
        st.success("初期データへ戻しました。")

    st.markdown("---")
    st.markdown("### Streamlit Secrets 例")
    st.code(
        'SUPABASE_URL = "https://xxxx.supabase.co"\nSUPABASE_KEY = "anon public key"',
        language="toml",
    )

    st.markdown("### CSV例")
    st.code(
        "department,grade,task_details\n"
        '人事課,G6,"応募者情報の入力\\n面接日程の調整\\n応募者への連絡対応\\n求人媒体の掲載情報更新"',
        language="csv",
    )

    st.markdown("### 補足")
    st.markdown("""
- `task_details` は **改行区切り** です。1行が1業務です。  
- 一覧タブでは具体業務を自動で 1. 2. 3. の形で表示します。  
- 保存は **Supabase REST API** を使っているため、`supabase-py` の proxy 依存問題を回避できます。  
- テーブルはスナップショット保存方式です。履歴をそのまま復元できます。  
- RLS を ON にする場合は、`SELECT` と `INSERT` のポリシーを設定してください。
""")
