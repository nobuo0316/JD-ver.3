import hashlib
import json
from io import BytesIO
from typing import Dict, List, Optional
from urllib import error, parse, request

import pandas as pd
import streamlit as st

# Optional Excel engines
try:
    import openpyxl  # noqa: F401
    OPENPYXL_AVAILABLE = True
except Exception:
    OPENPYXL_AVAILABLE = False

try:
    import xlsxwriter  # noqa: F401
    XLSXWRITER_AVAILABLE = True
except Exception:
    XLSXWRITER_AVAILABLE = False

st.set_page_config(
    page_title="Wage Table Guide / 賃金テーブル説明ページ",
    page_icon="📘",
    layout="wide",
)

# =========================================================
# Master data
# =========================================================
GRADES: List[str] = ["G6", "G5B", "G5A", "G4", "G3", "G2"]
STEPS: List[int] = list(range(1, 50))
NEXT_GRADE: Dict[str, str] = {
    "G6": "G5B",
    "G5B": "G5A",
    "G5A": "G4",
    "G4": "G3",
    "G3": "G2",
    "G2": "",
}
DEFAULT_EMPLOYEE_COLUMNS = [
    "Employee ID",
    "Name",
    "Grade",
    "Step",
    "University Graduate",
    "Adjustment Allowance",
    "Other Allowance",
    "Active",
]
DEFAULT_PARAMS = {
    "G6": {"base": 18000.0, "ap": 500.0, "pp": 1500.0},
    "G5B": {"base": 22000.0, "ap": 600.0, "pp": 1800.0},
    "G5A": {"base": 26000.0, "ap": 700.0, "pp": 2200.0},
    "G4": {"base": 32000.0, "ap": 800.0, "pp": 2800.0},
    "G3": {"base": 40000.0, "ap": 1000.0, "pp": 3500.0},
    "G2": {"base": 52000.0, "ap": 1200.0, "pp": 0.0},
}

LANGUAGE_PACK = {
    "日本語": {
        "title": "賃金テーブル管理・説明ページ",
        "subtitle": "制度説明、図解、シミュレーション、名簿管理を1つにまとめたページです。",
        "login_title": "ログイン",
        "login_id": "ID",
        "login_password": "パスワード",
        "login_button": "ログイン",
        "logout_button": "ログアウト",
        "login_error": "IDまたはパスワードが違います。",
        "logged_in_as": "ログイン中",
        "tab_overview": "制度説明",
        "tab_diagram": "図で理解",
        "tab_table": "賃金テーブル",
        "tab_sim": "昇格シミュレーション",
        "tab_next_year": "来年昇格シミュレーション",
        "tab_allowance_export": "手当込みエクスポート",
        "tab_employee": "従業員名簿",
        "tab_admin": "管理設定",
        "sidebar_currency": "通貨記号",
        "sidebar_decimals": "小数表示",
        "sidebar_example": "GS例",
        "overview_heading": "制度概要",
        "overview_text1": "本ページでは、Grade（グレード）と Step（ステップ）に基づく賃金テーブルを、日本語・英語の両方で説明・管理できます。",
        "overview_text2": "各社員には必ず Grade と Step があり、これを合わせて GS と表記します。例：G5A の Step 4 は G5A-S4 です。",
        "grade_axis": "横軸は Grade（G6～G2）、縦軸は Step（1～49）です。",
        "rule_heading": "基本ルール",
        "rule_ap": "AP（Annual Pay Raise）：毎年、同一グレード内で昇給する額です。",
        "rule_pp": "PP（Pay by Promotion）：昇格時に加算される昇給額です。",
        "rule_allow": "主な手当は、調整手当と大卒手当です。今後、必要に応じて追加可能です。",
        "rule_promo": "昇格時は、『現在給与 + 現在グレードの AP + 現在グレードの PP』を下回らないことを条件に、次グレードでその金額以上となる最も近い Step に移動します。",
        "grade_table": "グレード一覧",
        "grade": "グレード",
        "position": "役職イメージ",
        "next_grade": "次グレード",
        "glabel_G6": "平社員",
        "glabel_G5B": "シニアスタッフ",
        "glabel_G5A": "上位シニアスタッフ",
        "glabel_G4": "スーパーバイザー",
        "glabel_G3": "課長",
        "glabel_G2": "次長",
        "ap_pp_heading": "昇給ルール（AP / PP一覧）",
        "ap_label": "AP（毎年昇給）",
        "pp_label": "PP（昇格昇給）",
        "case_study_heading": "イメージしやすい例",
        "case_new_grad_title": "例1：新卒社員を採用した場合",
        "case_new_grad_text": "新卒社員は、例として G6-S1 からスタートします。その後、毎年 AP によって Step が1つずつ上がります。昇格時は、PP を加えた最低必要額を満たす次グレードの Step に移動します。",
        "case_mid_title": "例2：スーパーバイザーレベルの中途社員を採用した場合",
        "case_mid_text": "スーパーバイザーレベルの中途社員は、例として G4-S3 からスタートします。この場合も毎年 AP によって Step が上がり、昇格時には PP を加えた条件に基づいて次グレードの Step が決まります。",
        "case_note": "実際のスタート位置は、経験、スキル、採用条件に応じて決定されます。ここでは制度理解のための例を示しています。",
        "career_25_heading": "25年間のキャリア推移（モデルケース）",
        "career_newgrad_title": "新卒採用の場合（モデルケース）",
        "career_newgrad_text": "新卒社員は一般的に G6-S1 からスタートします。毎年、AP により Step が1段階ずつ上昇し、一定の評価および役割拡大に応じて上位 Grade へ昇格します。想定例として、入社時 G6-S1、5年目 G6-S5、6年目 G5B へ昇格、10年目 G5A へ昇格、15年目 G4 へ昇格、20年目 G3 へ昇格、25年目に G2 に到達するケースが考えられます。",
        "career_mid_title": "中途採用（スーパーバイザー相当）の場合",
        "career_mid_text": "スーパーバイザーレベルの中途社員は、G4帯からのスタートを想定します。初期 Grade が高いため、早期にマネジメント層へ移行する可能性があります。想定例として、入社時 G4-S3、3年目 G4-S5、4年目 G3 へ昇格、10年目 G2 へ昇格するケースが考えられます。",
        "career_note": "上記は制度理解のためのモデルケースです。実際の昇格・昇給は、評価、役割、組織方針に基づいて決定されます。",
        "career_diagram_heading": "キャリア推移イメージ（25年）",
        "diagram_heading1": "① Grade と Step の関係",
        "diagram_heading2": "② AP と PP の考え方",
        "diagram_heading3": "③ 昇格時の移動イメージ",
        "diagram_help1": "下の図では、横方向が Grade、縦方向が Step です。社員は必ずどこか1つの GS に所属します。",
        "diagram_help2": "AP は同じ Grade 内での毎年の昇給、PP は次の Grade へ上がるときの追加昇給です。",
        "diagram_help3": "昇格時は、最低必要額を満たす次グレードの最も近い Step に移動します。",
        "simple_example_text": "例：G5A-S4 の社員が昇格する場合、まず『現在給与 + AP + PP』で最低必要額を算出し、その金額以上となる G4 の最初の Step を探します。",
        "wage_heading": "賃金テーブル",
        "wage_caption": "値は直接編集できます。説明用にも管理用にも利用できます。",
        "table_view_mode": "表示モード",
        "table_mode_raw": "数値のみ",
        "table_mode_with_label": "GSラベル付き",
        "download_csv": "CSVをダウンロード",
        "download_excel": "Excelをダウンロード",
        "download_note": "必要に応じて、このまま CSV / Excel で配布できます。",
        "sim_heading": "昇格シミュレーション",
        "current_grade": "現在グレード",
        "current_step": "現在ステップ",
        "current_salary": "現在給与",
        "min_required": "昇格時の最低必要額",
        "promoted_grade": "昇格後グレード",
        "promoted_step": "昇格後ステップ",
        "promoted_salary": "昇格後基本給",
        "adjust_allowance": "調整手当",
        "univ_allowance": "大卒手当",
        "is_univ": "大卒対象",
        "other_allowance": "その他手当",
        "final_salary": "最終支給額",
        "simulate": "シミュレーション実行",
        "no_next_grade": "これ以上の次グレードはありません。",
        "gs_before": "昇格前GS",
        "gs_after": "昇格後GS",
        "promotion_flow": "昇格の流れ",
        "step_search_result": "該当Stepの探索結果",
        "next_year_heading": "来年昇格シミュレーション",
        "apply_next_step": "昇格前に来年の Step アップを反映する",
        "next_year_current_step": "来年時点の現在ステップ",
        "next_year_current_salary": "来年時点の現在給与",
        "next_year_result": "来年昇格後の GS",
        "allowance_export_heading": "手当込み賃金テーブル",
        "allowance_export_text": "固定額の手当を加えた賃金テーブルを表示・エクスポートできます。",
        "adjustment_allowance_export": "調整手当（固定額）",
        "university_allowance_export": "大卒手当（固定額）",
        "other_allowance_export": "その他手当（固定額）",
        "include_adjustment_allowance": "調整手当を含める",
        "include_university_allowance": "大卒手当を含める",
        "include_other_allowance": "その他手当を含める",
        "export_with_allowances_csv": "手当込みCSVをダウンロード",
        "export_with_allowances_excel": "手当込みExcelをダウンロード",
        "employee_heading": "従業員名簿",
        "employee_text": "従業員名簿を CSV で読み込み、各従業員の基本給、手当、合計支給額、来年昇格後 GS を一覧化できます。",
        "employee_template_download": "従業員名簿テンプレートをダウンロード",
        "employee_upload": "従業員名簿CSVをアップロード",
        "employee_apply": "従業員名簿を反映",
        "employee_preview": "従業員名簿プレビュー",
        "employee_import_success": "従業員名簿を反映しました。",
        "employee_import_empty": "従業員名簿CSVを選択してください。",
        "employee_import_error": "従業員名簿CSVの形式が正しくありません。",
        "employee_export_csv": "従業員給与一覧CSVをダウンロード",
        "employee_export_excel": "従業員給与一覧Excelをダウンロード",
        "employee_id": "社員ID",
        "employee_name": "氏名",
        "employee_grade": "Grade",
        "employee_step": "Step",
        "employee_basic_pay": "基本給",
        "employee_university_flag": "大卒対象",
        "employee_adjustment_allowance": "調整手当",
        "employee_other_allowance": "その他手当",
        "employee_university_allowance": "大卒手当",
        "employee_total_allowance": "手当合計",
        "employee_total_pay": "合計支給額",
        "employee_next_year_gs": "来年昇格後GS",
        "employee_next_year_basic_pay": "来年昇格後基本給",
        "active_only": "在籍者のみ表示",
        "default_university_allowance": "名簿計算用の大卒手当（固定額）",
        "admin_heading": "管理設定",
        "admin_text": "ここで初期テーブルの自動生成ルールを変更できます。実際の賃金額が決まっている場合は、下のテーブルを直接編集してください。",
        "base_salary": "Step1基準額",
        "ap": "AP（毎年昇給額）",
        "pp": "PP（昇格昇給額）",
        "rebuild": "設定値でテーブル再生成",
        "reset": "初期値に戻す",
        "warning_rebuild": "再生成すると、現在の手動編集内容は上書きされます。",
        "success_rebuild": "テーブルを再生成して保存しました。",
        "success_reset": "初期値に戻して保存しました。",
        "csv_import_heading": "設定CSVインポート",
        "csv_import_text": "この画面の設定値（Step1基準額、AP、PP）をCSVで一括更新できます。反映後、賃金テーブルも自動で再生成され、Supabase に保存されます。",
        "csv_upload": "設定CSVファイルをアップロード",
        "csv_apply": "CSVを設定に反映",
        "csv_template_download": "設定CSVテンプレートをダウンロード",
        "csv_import_success": "CSVから設定値を更新し、賃金テーブルを再生成して保存しました。",
        "csv_import_error": "CSVの形式が正しくありません。Grade, Base, AP, PP 列が必要で、G6, G5B, G5A, G4, G3, G2 の6行が必要です。",
        "csv_import_empty": "CSVファイルを選択してください。",
        "csv_preview_heading": "CSVプレビュー",
        "currency_preview": "表示例",
        "excel_unavailable": "この環境ではExcel出力が使えません。CSVを使うか、requirements.txt に openpyxl または xlsxwriter を追加してください。",
        "supabase_status_ok": "Supabase接続: ON",
        "supabase_status_off": "Supabase接続: OFF（ローカル初期値を使用）",
        "supabase_save_error": "Supabase保存に失敗しました。",
        "admin_password": "管理用パスワード",
        "admin_unlock": "管理ロック解除",
        "admin_locked": "管理設定の変更にはパスワードが必要です。",
        "admin_unlocked": "管理ロックを解除しました。",
        "admin_password_error": "パスワードが違います。",
        "admin_only_upload_notice": "🔒 この操作は管理者のみ実行できます。一般ユーザーは閲覧のみ可能です。",
    },
    "English": {
        "title": "Wage Table Management & Explanation Page",
        "subtitle": "A page that combines system explanation, visual guides, simulation, and roster management.",
        "login_title": "Login",
        "login_id": "ID",
        "login_password": "Password",
        "login_button": "Login",
        "logout_button": "Logout",
        "login_error": "Incorrect ID or password.",
        "logged_in_as": "Logged in as",
        "tab_overview": "Overview",
        "tab_diagram": "Visual Guide",
        "tab_table": "Wage Table",
        "tab_sim": "Promotion Simulation",
        "tab_next_year": "Next-Year Promotion",
        "tab_allowance_export": "Allowance Export",
        "tab_employee": "Employee Roster",
        "tab_admin": "Admin Settings",
        "sidebar_currency": "Currency symbol",
        "sidebar_decimals": "Decimal places",
        "sidebar_example": "GS Example",
        "overview_heading": "System Overview",
        "overview_text1": "This page explains and manages the wage table based on Grade and Step in both English and Japanese.",
        "overview_text2": "Each employee always has a Grade and a Step, abbreviated as GS. Example: G5A Step 4 is written as G5A-S4.",
        "grade_axis": "The horizontal axis is Grade (G6 to G2), and the vertical axis is Step (1 to 49).",
        "rule_heading": "Basic Rules",
        "rule_ap": "AP (Annual Pay Raise): the amount of annual raise within the same grade.",
        "rule_pp": "PP (Pay by Promotion): the additional increase applied upon promotion.",
        "rule_allow": "Main allowances are Adjustment Allowance and University Graduate Allowance. Additional allowances can be added later if needed.",
        "rule_promo": "Upon promotion, the new salary must not be lower than Current Salary + current grade AP + current grade PP. The employee moves to the closest step in the next grade that meets or exceeds that amount.",
        "grade_table": "Grade Reference",
        "grade": "Grade",
        "position": "Position Image",
        "next_grade": "Next Grade",
        "glabel_G6": "Staff",
        "glabel_G5B": "Senior Staff",
        "glabel_G5A": "Upper Senior Staff",
        "glabel_G4": "Supervisor",
        "glabel_G3": "Manager",
        "glabel_G2": "Deputy General Manager",
        "ap_pp_heading": "Raise Rules (AP / PP)",
        "ap_label": "AP (Annual Raise)",
        "pp_label": "PP (Promotion Raise)",
        "case_study_heading": "Easy-to-Understand Examples",
        "case_new_grad_title": "Case 1: When a new graduate is hired",
        "case_new_grad_text": "A new graduate may start at G6-S1. After that, the Step increases by AP every year. At promotion, the employee moves to the first Step in the next Grade that satisfies the minimum required amount including PP.",
        "case_mid_title": "Case 2: When a supervisor-level mid-career employee is hired",
        "case_mid_text": "A supervisor-level mid-career hire may start at G4-S3. In the same way, the Step increases every year through AP. At promotion, the employee moves to the first Step in the next Grade that satisfies the minimum required amount including PP.",
        "case_note": "The actual starting position depends on experience, skills, and hiring conditions. These are examples provided for easier understanding of the system.",
        "career_25_heading": "25-Year Career Progression (Model Cases)",
        "career_newgrad_title": "New Graduate Hire (Model Case)",
        "career_newgrad_text": "A new graduate typically starts at G6-S1. Every year, the Step increases by one level through AP, and the employee may be promoted to a higher Grade based on evaluation and increased responsibility. A sample path could be: Start at G6-S1, Year 5 at G6-S5, Year 6 promoted to G5B, Year 10 promoted to G5A, Year 15 promoted to G4, Year 20 promoted to G3, and Year 25 potentially reaching G2.",
        "career_mid_title": "Mid-Career Hire at Supervisor Level",
        "career_mid_text": "A mid-career employee hired at supervisor level may start in the G4 range. Because the initial Grade is higher, transition to management levels may occur sooner. A sample path could be: Start at G4-S3, Year 3 at G4-S5, Year 4 promoted to G3, and Year 10 promoted to G2.",
        "career_note": "These are model cases for explanation purposes. Actual salary increases and promotions are determined based on evaluation, role, and company policy.",
        "career_diagram_heading": "Career Progression Image (25 Years)",
        "diagram_heading1": "1) Relationship between Grade and Step",
        "diagram_heading2": "2) How AP and PP work",
        "diagram_heading3": "3) Promotion movement image",
        "diagram_help1": "In the chart below, Grade runs horizontally and Step runs vertically. Every employee always belongs to one GS position.",
        "diagram_help2": "AP means annual raise within the same grade. PP means the additional raise given when moving to the next grade.",
        "diagram_help3": "At promotion, the employee moves to the closest step in the next grade that meets the minimum required amount.",
        "simple_example_text": "Example: when an employee at G5A-S4 is promoted, first calculate Current Salary + AP + PP, then find the first Step in G4 that is equal to or higher than that threshold.",
        "wage_heading": "Wage Table",
        "wage_caption": "Values can be edited directly. The table can be used both for explanation and administration.",
        "table_view_mode": "View Mode",
        "table_mode_raw": "Raw Numbers",
        "table_mode_with_label": "With GS Labels",
        "download_csv": "Download CSV",
        "download_excel": "Download Excel",
        "download_note": "You can distribute this as CSV or Excel as needed.",
        "sim_heading": "Promotion Simulation",
        "current_grade": "Current Grade",
        "current_step": "Current Step",
        "current_salary": "Current Salary",
        "min_required": "Minimum Required for Promotion",
        "promoted_grade": "Promoted Grade",
        "promoted_step": "Promoted Step",
        "promoted_salary": "Promoted Base Salary",
        "adjust_allowance": "Adjustment Allowance",
        "univ_allowance": "University Allowance",
        "is_univ": "University Graduate",
        "other_allowance": "Other Allowance",
        "final_salary": "Final Pay",
        "simulate": "Run Simulation",
        "no_next_grade": "No higher grade is available.",
        "gs_before": "Current GS",
        "gs_after": "New GS",
        "promotion_flow": "Promotion Flow",
        "step_search_result": "Step Search Result",
        "next_year_heading": "Next-Year Promotion Simulation",
        "apply_next_step": "Apply next year's step-up before promotion",
        "next_year_current_step": "Next-year current step",
        "next_year_current_salary": "Next-year current salary",
        "next_year_result": "Next-year promoted GS",
        "allowance_export_heading": "Allowance-Included Wage Table",
        "allowance_export_text": "You can view and export a wage table with fixed allowances added.",
        "adjustment_allowance_export": "Adjustment allowance (fixed)",
        "university_allowance_export": "University allowance (fixed)",
        "other_allowance_export": "Other allowance (fixed)",
        "include_adjustment_allowance": "Include adjustment allowance",
        "include_university_allowance": "Include university allowance",
        "include_other_allowance": "Include other allowance",
        "export_with_allowances_csv": "Download allowance CSV",
        "export_with_allowances_excel": "Download allowance Excel",
        "employee_heading": "Employee Roster",
        "employee_text": "Import an employee roster CSV and calculate base pay, allowances, total pay, and next-year promoted GS for each employee.",
        "employee_template_download": "Download employee roster template",
        "employee_upload": "Upload employee roster CSV",
        "employee_apply": "Apply employee roster",
        "employee_preview": "Employee roster preview",
        "employee_import_success": "Employee roster was applied.",
        "employee_import_empty": "Please choose an employee roster CSV.",
        "employee_import_error": "Employee roster CSV format is invalid.",
        "employee_export_csv": "Download employee payroll CSV",
        "employee_export_excel": "Download employee payroll Excel",
        "employee_id": "Employee ID",
        "employee_name": "Name",
        "employee_grade": "Grade",
        "employee_step": "Step",
        "employee_basic_pay": "Base Pay",
        "employee_university_flag": "University Graduate",
        "employee_adjustment_allowance": "Adjustment Allowance",
        "employee_other_allowance": "Other Allowance",
        "employee_university_allowance": "University Allowance",
        "employee_total_allowance": "Allowance Total",
        "employee_total_pay": "Total Pay",
        "employee_next_year_gs": "Next-Year Promoted GS",
        "employee_next_year_basic_pay": "Next-Year Promoted Base Pay",
        "active_only": "Show active employees only",
        "default_university_allowance": "Fixed university allowance for roster calculation",
        "admin_heading": "Admin Settings",
        "admin_text": "Here you can change the auto-generation rule for the initial table. If actual salary values are fixed, you can edit the wage table directly below.",
        "base_salary": "Base Salary at Step 1",
        "ap": "AP (Annual Raise)",
        "pp": "PP (Promotion Raise)",
        "rebuild": "Rebuild table from settings",
        "reset": "Reset to defaults",
        "warning_rebuild": "Rebuilding will overwrite current manual edits.",
        "success_rebuild": "The wage table was rebuilt and saved.",
        "success_reset": "Defaults were restored and saved.",
        "csv_import_heading": "Settings CSV Import",
        "csv_import_text": "You can bulk update the settings shown on this screen (Base at Step 1, AP, and PP) by CSV. After import, the wage table is regenerated and saved to Supabase.",
        "csv_upload": "Upload settings CSV file",
        "csv_apply": "Apply CSV to settings",
        "csv_template_download": "Download settings CSV template",
        "csv_import_success": "The settings were updated from CSV, the wage table was regenerated, and the changes were saved.",
        "csv_import_error": "The CSV format is invalid. The file must contain Grade, Base, AP, and PP columns, with 6 rows for G6, G5B, G5A, G4, G3, and G2.",
        "csv_import_empty": "Please choose a CSV file first.",
        "csv_preview_heading": "CSV Preview",
        "currency_preview": "Preview",
        "excel_unavailable": "Excel export is unavailable in this environment. Please use CSV export or add openpyxl / xlsxwriter to requirements.txt.",
        "supabase_status_ok": "Supabase connection: ON",
        "supabase_status_off": "Supabase connection: OFF (using local defaults)",
        "supabase_save_error": "Failed to save to Supabase.",
        "admin_password": "Admin password",
        "admin_unlock": "Unlock admin",
        "admin_locked": "A password is required to change admin settings.",
        "admin_unlocked": "Admin is unlocked.",
        "admin_password_error": "Incorrect password.",
        "admin_only_upload_notice": "🔒 This action is available to administrators only. General users can view data only.",
    },
}


def t(key: str) -> str:
    return LANGUAGE_PACK[st.session_state.lang][key]


def grade_label(grade: str) -> str:
    return LANGUAGE_PACK[st.session_state.lang][f"glabel_{grade}"]


def format_money(value: float) -> str:
    decimals = st.session_state.decimals
    symbol = st.session_state.currency_symbol
    try:
        return f"{symbol}{float(value):,.{decimals}f}"
    except Exception:
        return "-"


def hash_password(password: str) -> str:
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


def login_enabled() -> bool:
    return get_supabase_config() is not None


def build_wage_table(params: Dict[str, Dict[str, float]]) -> pd.DataFrame:
    data = {"Step": STEPS}
    for g in GRADES:
        base = params[g]["base"]
        ap = params[g]["ap"]
        data[g] = [base + (step - 1) * ap for step in STEPS]
    return pd.DataFrame(data)


def make_excel_file(df: pd.DataFrame) -> Optional[bytes]:
    output = BytesIO()
    try:
        if OPENPYXL_AVAILABLE:
            engine = "openpyxl"
        elif XLSXWRITER_AVAILABLE:
            engine = "xlsxwriter"
        else:
            return None
        with pd.ExcelWriter(output, engine=engine) as writer:
            df.to_excel(writer, index=False, sheet_name="Export")
        output.seek(0)
        return output.getvalue()
    except Exception:
        return None


def build_settings_csv_template() -> pd.DataFrame:
    return pd.DataFrame([
        {"Grade": g, "Base": DEFAULT_PARAMS[g]["base"], "AP": DEFAULT_PARAMS[g]["ap"], "PP": DEFAULT_PARAMS[g]["pp"]}
        for g in GRADES
    ])


def build_employee_csv_template() -> pd.DataFrame:
    return pd.DataFrame([
        {
            "Employee ID": "E001",
            "Name": "Sample Employee",
            "Grade": "G5A",
            "Step": 4,
            "University Graduate": 1,
            "Adjustment Allowance": 1000,
            "Other Allowance": 500,
            "Active": 1,
        }
    ])


def validate_imported_settings_csv(df: pd.DataFrame) -> Dict[str, Dict[str, float]]:
    required_cols = ["Grade", "Base", "AP", "PP"]
    if list(df.columns) != required_cols:
        raise ValueError(f"Columns must be exactly: {required_cols}")
    if len(df) != len(GRADES):
        raise ValueError(f"Row count must be {len(GRADES)}")
    if df["Grade"].tolist() != GRADES:
        raise ValueError(f"Grade order must be: {GRADES}")

    new_params: Dict[str, Dict[str, float]] = {}
    for i, row in df.iterrows():
        try:
            new_params[str(row["Grade"])] = {
                "base": float(row["Base"]),
                "ap": float(row["AP"]),
                "pp": float(row["PP"]),
            }
        except Exception as exc:
            raise ValueError(f"Row {i + 1}: Base/AP/PP must be numeric") from exc
    return new_params


def validate_employee_roster_csv(df: pd.DataFrame) -> pd.DataFrame:
    required_cols = DEFAULT_EMPLOYEE_COLUMNS
    if list(df.columns) != required_cols:
        raise ValueError(f"Columns must be exactly: {required_cols}")

    out = df.copy()
    out["Grade"] = out["Grade"].astype(str)
    if not out["Grade"].isin(GRADES).all():
        raise ValueError("Grade column contains invalid grade")

    out["Step"] = pd.to_numeric(out["Step"], errors="raise").astype(int)
    if not out["Step"].isin(STEPS).all():
        raise ValueError("Step column contains invalid step")

    for col in ["University Graduate", "Adjustment Allowance", "Other Allowance", "Active"]:
        out[col] = pd.to_numeric(out[col], errors="raise")

    out["University Graduate"] = out["University Graduate"].astype(int)
    out["Active"] = out["Active"].astype(int)
    return out


def params_to_rows(params: Dict[str, Dict[str, float]]) -> List[Dict[str, float]]:
    return [
        {"grade": g, "base": float(params[g]["base"]), "ap": float(params[g]["ap"]), "pp": float(params[g]["pp"])}
        for g in GRADES
    ]


def rows_to_params(rows: List[Dict[str, float]]) -> Dict[str, Dict[str, float]]:
    rows_by_grade = {row["grade"]: row for row in rows}
    params: Dict[str, Dict[str, float]] = {}
    for g in GRADES:
        params[g] = {
            "base": float(rows_by_grade[g]["base"]),
            "ap": float(rows_by_grade[g]["ap"]),
            "pp": float(rows_by_grade[g]["pp"]),
        }
    return params


def get_supabase_config() -> Optional[Dict[str, str]]:
    try:
        url = st.secrets["JM_SUPABASE_URL"]
        service_role_key = st.secrets["JM_SUPABASE_SERVICE_ROLE_KEY"]
        table = st.secrets.get("JM_SETTINGS_TABLE", "wage_settings")
        users_table = st.secrets.get("JM_USERS_TABLE", "job_matrix_users")
        return {
            "url": str(url).rstrip("/"),
            "key": str(service_role_key),
            "table": str(table),
            "users_table": str(users_table),
        }
    except Exception:
        return None


def supabase_request(method: str, path: str, body: Optional[object] = None, query: Optional[Dict[str, str]] = None):
    config = get_supabase_config()
    if config is None:
        raise RuntimeError("Supabase is not configured")

    url = f"{config['url']}/rest/v1/{path}"
    if query:
        url = f"{url}?{parse.urlencode(query)}"

    headers = {
        "apikey": config["key"],
        "Authorization": f"Bearer {config['key']}",
        "Content-Type": "application/json",
    }
    if method == "POST":
        headers["Prefer"] = "resolution=merge-duplicates,return=representation"
    elif method in ("PATCH",):
        headers["Prefer"] = "return=representation"

    data = None
    if body is not None:
        data = json.dumps(body).encode("utf-8")

    req = request.Request(url=url, data=data, headers=headers, method=method)
    with request.urlopen(req, timeout=20) as response:
        text = response.read().decode("utf-8")
        return json.loads(text) if text else None


def load_settings_from_supabase() -> Dict[str, Dict[str, float]]:
    config = get_supabase_config()
    if config is None:
        return {k: v.copy() for k, v in DEFAULT_PARAMS.items()}
    try:
        result = supabase_request(
            method="GET",
            path=config["table"],
            query={"select": "grade,base,ap,pp", "order": "grade.asc"},
        )
        if not result:
            return {k: v.copy() for k, v in DEFAULT_PARAMS.items()}
        grades_found = [row["grade"] for row in result]
        if sorted(grades_found) != sorted(GRADES):
            return {k: v.copy() for k, v in DEFAULT_PARAMS.items()}
        return rows_to_params(result)
    except Exception:
        return {k: v.copy() for k, v in DEFAULT_PARAMS.items()}


def save_settings_to_supabase(params: Dict[str, Dict[str, float]]) -> None:
    config = get_supabase_config()
    if config is None:
        return
    try:
        supabase_request(method="POST", path=config["table"], body=params_to_rows(params), query={"on_conflict": "grade"})
    except error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="ignore")
        raise RuntimeError(detail) from exc
    except Exception as exc:
        raise RuntimeError(str(exc)) from exc


def get_login_users_from_supabase() -> List[Dict[str, object]]:


    config = get_supabase_config()
    if config is None:
        return [fallback_admin]
    try:
        result = supabase_request(
            method="GET",
            path=config["users_table"],
            query={"select": "*", "limit": "500"},
        )
        if isinstance(result, list):
            return result if isinstance(result, list) else []
        return [fallback_admin]
        
    except Exception:
        return [fallback_admin]


def _candidate_login_values(user: Dict[str, object]) -> List[str]:
    keys = ["username", "login_id", "user_id", "email", "id"]
    values: List[str] = []
    for key in keys:
        value = user.get(key)
        if value is None:
            continue
        text = str(value).strip()
        if text:
            values.append(text)
    return values


def get_user_by_login_id(login_id: str) -> Optional[Dict[str, object]]:
    users = get_login_users_from_supabase()
    for user in users:
        if str(user.get("username")).strip() == login_id:
            return user
    return None


def verify_password(password: str, user: Dict[str, object]) -> bool:
    raw = str(password or "")
    stored = str(user.get("password_hash") or "").strip()

    if not stored:
        return False

    input_hash = hashlib.sha256(raw.encode("utf-8")).hexdigest()

    return stored.lower() == input_hash.lower()


def normalize_role(role_value: object) -> str:
    role = str(role_value or "viewer").strip().lower()
    return role if role in ("admin", "viewer") else "viewer"


def is_user_active(user: Dict[str, object]) -> bool:
    value = user.get("is_active", True)
    if value is None:
        return True
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in ("true", "1", "yes", "y")


def authenticate_user(login_id: str, password: str) -> Optional[Dict[str, object]]:
    user = get_user_by_login_id(login_id)
    if not user:
        return None
    if not is_user_active(user):
        return None
    if not verify_password(password, user):
        return None
    user["role"] = normalize_role(user.get("role"))
    return user


def is_admin() -> bool:
    return st.session_state.get("user_role", "viewer") == "admin"


def is_viewer() -> bool:
    return st.session_state.get("user_role", "viewer") in ("viewer", "admin")


def get_current_salary(df: pd.DataFrame, grade: str, step: int) -> float:
    row = df.loc[df["Step"] == step]
    return float(row.iloc[0][grade])


def find_promotion_result(df: pd.DataFrame, params: Dict[str, Dict[str, float]], current_grade: str, current_step: int) -> Optional[Dict[str, float]]:
    next_grade = NEXT_GRADE[current_grade]
    if not next_grade:
        return None
    current_salary = get_current_salary(df, current_grade, current_step)
    minimum_required = current_salary + params[current_grade]["ap"] + params[current_grade]["pp"]
    next_rows = df[["Step", next_grade]].copy()
    eligible = next_rows[next_rows[next_grade] >= minimum_required]
    if eligible.empty:
        target_step = int(next_rows.iloc[-1]["Step"])
        target_salary = float(next_rows.iloc[-1][next_grade])
    else:
        target_step = int(eligible.iloc[0]["Step"])
        target_salary = float(eligible.iloc[0][next_grade])
    return {
        "current_salary": current_salary,
        "minimum_required": minimum_required,
        "target_step": target_step,
        "target_salary": target_salary,
        "target_grade": next_grade,
    }


def find_next_year_promotion_result(df: pd.DataFrame, params: Dict[str, Dict[str, float]], current_grade: str, current_step: int, apply_next_step: bool = True) -> Optional[Dict[str, float]]:
    simulated_step = min(current_step + 1, max(STEPS)) if apply_next_step else current_step
    base_result = find_promotion_result(df, params, current_grade, simulated_step)
    if base_result is None:
        return None
    return {
        "next_year_current_step": simulated_step,
        "next_year_current_salary": base_result["current_salary"],
        "minimum_required": base_result["minimum_required"],
        "target_grade": base_result["target_grade"],
        "target_step": base_result["target_step"],
        "target_salary": base_result["target_salary"],
    }


def build_allowance_export_table(df: pd.DataFrame, include_adjustment: bool, adjustment_amount: float, include_university: bool, university_amount: float, include_other: bool, other_amount: float) -> pd.DataFrame:
    out = df.copy()
    allowance_total = 0.0
    if include_adjustment:
        allowance_total += adjustment_amount
    if include_university:
        allowance_total += university_amount
    if include_other:
        allowance_total += other_amount
    for g in GRADES:
        out[g] = pd.to_numeric(out[g], errors="coerce")
    out["Allowance Total"] = allowance_total
    for g in GRADES:
        out[f"{g} Total"] = out[g] + allowance_total
    return out


def build_employee_payroll(roster_df: pd.DataFrame, wage_df: pd.DataFrame, params: Dict[str, Dict[str, float]], university_allowance_amount: float, apply_next_step: bool = True) -> pd.DataFrame:
    rows = []
    for _, row in roster_df.iterrows():
        grade = str(row["Grade"])
        step = int(row["Step"])
        basic_pay = get_current_salary(wage_df, grade, step)
        adjustment_allowance = float(row["Adjustment Allowance"])
        other_allowance = float(row["Other Allowance"])
        university_flag = int(row["University Graduate"])
        university_allowance = float(university_allowance_amount if university_flag == 1 else 0.0)
        total_allowance = adjustment_allowance + other_allowance + university_allowance
        total_pay = basic_pay + total_allowance
        next_year = find_next_year_promotion_result(wage_df, params, grade, step, apply_next_step=apply_next_step)
        next_year_gs = "-"
        next_year_basic = None
        if next_year is not None:
            next_year_gs = f"{next_year['target_grade']}-S{next_year['target_step']}"
            next_year_basic = float(next_year["target_salary"])
        rows.append({
            t("employee_id"): str(row["Employee ID"]),
            t("employee_name"): str(row["Name"]),
            t("employee_grade"): grade,
            t("employee_step"): step,
            t("employee_basic_pay"): basic_pay,
            t("employee_university_flag"): university_flag,
            t("employee_adjustment_allowance"): adjustment_allowance,
            t("employee_other_allowance"): other_allowance,
            t("employee_university_allowance"): university_allowance,
            t("employee_total_allowance"): total_allowance,
            t("employee_total_pay"): total_pay,
            t("employee_next_year_gs"): next_year_gs,
            t("employee_next_year_basic_pay"): next_year_basic,
            "_active": int(row["Active"]),
        })
    return pd.DataFrame(rows)


def display_table_with_formats(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for g in GRADES:
        out[g] = out[g].apply(format_money)
    return out


def display_table_with_gs(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, row in df.iterrows():
        step = int(row["Step"])
        record = {"Step": step}
        for g in GRADES:
            record[g] = f"{g}-S{step} / {format_money(float(row[g]))}"
        rows.append(record)
    return pd.DataFrame(rows)


def grade_step_grid(selected_grade: str = "G5A", selected_step: int = 4) -> str:
    lines = []
    lines.append("digraph G {")
    lines.append("rankdir=TB;")
    lines.append('graph [nodesep="0.25", ranksep="0.35"];')
    lines.append('node [shape="box", style="rounded,filled", fillcolor="white", width="1.0", height="0.5", fontsize="10"];')
    headers = []
    for g in GRADES:
        header = f"header_{g}"
        lines.append(f'{header} [label="{g}", shape="plaintext", fontsize="12"];')
        headers.append(header)
    lines.append("{ rank=same; " + "; ".join(headers) + "; }")
    for s in range(1, 6):
        row_nodes = []
        for g in GRADES:
            node_name = f"{g}_{s}"
            label = f"{g}-S{s}"
            if g == selected_grade and s == selected_step:
                lines.append(f'{node_name} [label="{label}", fillcolor="lightblue"];')
            else:
                lines.append(f'{node_name} [label="{label}"];')
            row_nodes.append(node_name)
        lines.append("{ rank=same; " + "; ".join(row_nodes) + "; }")
    for g in GRADES:
        chain = [f"header_{g}"] + [f"{g}_{s}" for s in range(1, 6)]
        for i in range(len(chain) - 1):
            lines.append(f'{chain[i]} -> {chain[i + 1]} [style="invis", weight=10];')
    lines.append("}")
    return "\n".join(lines)


def raise_diagram() -> str:
    return """
    digraph G {
      rankdir=LR;
      node [shape="box", style="rounded,filled", fillcolor="white"];
      A [label="Current Salary\n現在給与"];
      B [label="+ AP\nAnnual Raise"];
      C [label="+ PP\nPromotion Raise"];
      D [label="Minimum Required\n最低必要額"];
      E [label="Find first Step in next Grade\n次グレードで該当Stepを探す", fillcolor="lightyellow"];
      A -> B -> C -> D -> E;
    }
    """


def promotion_diagram(current_grade: str, current_step: int, next_grade: str, target_step: int) -> str:
    return f"""
    digraph G {{
      rankdir=LR;
      node [shape="box", style="rounded,filled", fillcolor="white"];
      A [label="{current_grade}-S{current_step}\nCurrent GS", fillcolor="lightblue"];
      B [label="AP + PP\nadded"];
      C [label="Search next grade\n{next_grade}"];
      D [label="{next_grade}-S{target_step}\nNew GS", fillcolor="lightgreen"];
      A -> B -> C -> D;
    }}
    """


def save_and_rebuild(params: Dict[str, Dict[str, float]]) -> None:
    st.session_state.params = params
    st.session_state.wage_df = build_wage_table(params)
    save_settings_to_supabase(params)


# =========================================================
# Session state
# =========================================================
if "lang" not in st.session_state:
    st.session_state.lang = "日本語"
if "currency_symbol" not in st.session_state:
    st.session_state.currency_symbol = "₱"
if "decimals" not in st.session_state:
    st.session_state.decimals = 0
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "login_user" not in st.session_state:
    st.session_state.login_user = ""
if "display_name" not in st.session_state:
    st.session_state.display_name = ""
if "user_role" not in st.session_state:
    st.session_state.user_role = "viewer"
if "params" not in st.session_state:
    st.session_state.params = load_settings_from_supabase()
if "wage_df" not in st.session_state:
    st.session_state.wage_df = build_wage_table(st.session_state.params)
if "employee_roster_df" not in st.session_state:
    st.session_state.employee_roster_df = pd.DataFrame(columns=DEFAULT_EMPLOYEE_COLUMNS)

# =========================================================
# Simple visual styling
# =========================================================
st.markdown(
    """
    <style>
    .main-card {
        background: linear-gradient(135deg, #f8fbff 0%, #eef6ff 100%);
        border: 1px solid #d7e7ff;
        border-radius: 14px;
        padding: 14px 16px;
        margin-bottom: 14px;
    }
    .info-card {
        background: #f7fafc;
        border-left: 5px solid #4a90e2;
        border-radius: 10px;
        padding: 12px 14px;
        margin-bottom: 12px;
    }
    .warn-card {
        background: #fff8e8;
        border-left: 5px solid #f0ad4e;
        border-radius: 10px;
        padding: 12px 14px;
        margin-bottom: 12px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# =========================================================
# Login gate
# =========================================================
if login_enabled() and not st.session_state.logged_in:
    st.title(t("login_title"))
    login_id = st.text_input(t("login_id"))
    login_password = st.text_input(t("login_password"), type="password")
    if st.button(t("login_button"), use_container_width=True):
        user = authenticate_user(login_id.strip(), login_password)
        if user is not None:
            resolved_login = ""
            for k in ("username", "login_id", "user_id", "email", "id"):
                v = str(user.get(k) or "").strip()
                if v:
                    resolved_login = v
                    break
            st.session_state.logged_in = True
            st.session_state.login_user = resolved_login or login_id.strip()
            st.session_state.display_name = str(user.get("display_name") or st.session_state.login_user).strip()
            st.session_state.user_role = normalize_role(user.get("role"))
            st.rerun()
        else:
            st.error(t("login_error"))
            with st.expander("Login debug"):
                cfg = get_supabase_config()
                st.write({
                    "supabase_configured": cfg is not None,
                    "users_table": cfg.get("users_table") if cfg else None,
                    "user_count_preview": len(get_login_users_from_supabase()) if cfg else 0,
                    "login_id_entered": login_id.strip(),
                })
    st.stop()
elif not login_enabled():
    st.warning("JM_* secrets are not configured. The app is running in local viewer mode.")

if not is_admin():
    st.session_state.employee_roster_df = pd.DataFrame(columns=DEFAULT_EMPLOYEE_COLUMNS)

# =========================================================
# Sidebar
# =========================================================
st.sidebar.title("Wage Table")
st.session_state.lang = st.sidebar.radio(
    "Language / 言語",
    ["日本語", "English"],
    index=0 if st.session_state.lang == "日本語" else 1,
)
st.session_state.currency_symbol = st.sidebar.text_input(t("sidebar_currency"), value=st.session_state.currency_symbol)
st.session_state.decimals = st.sidebar.selectbox(t("sidebar_decimals"), [0, 1, 2], index=[0, 1, 2].index(st.session_state.decimals))

if login_enabled() and st.session_state.logged_in:
    role_label = st.session_state.user_role or "viewer"
    display_user = st.session_state.display_name or st.session_state.login_user
    st.sidebar.caption(f"{t('logged_in_as')}: {display_user} ({role_label})")
    if st.sidebar.button(t("logout_button"), use_container_width=True):
        st.session_state.logged_in = False
        st.session_state.login_user = ""
        st.session_state.display_name = ""
        st.session_state.user_role = "viewer"
        st.session_state.employee_roster_df = pd.DataFrame(columns=DEFAULT_EMPLOYEE_COLUMNS)
        st.rerun()

if get_supabase_config() is not None:
    st.sidebar.success(t("supabase_status_ok"))
else:
    st.sidebar.info(t("supabase_status_off"))

st.sidebar.caption(f"{t('currency_preview')}: {format_money(12345.67)}")
st.sidebar.markdown("---")
st.sidebar.write(t("sidebar_example"))
example_grade = st.sidebar.selectbox("Grade", GRADES, index=2, key="example_grade")
example_step = st.sidebar.selectbox("Step", STEPS, index=3, key="example_step")
st.sidebar.info(f"GS = {example_grade}-S{example_step}")

# =========================================================
# Main header
# =========================================================
st.markdown(f"<div class='main-card'><h1 style='margin-bottom:4px;'>{t('title')}</h1><div>{t('subtitle')}</div></div>", unsafe_allow_html=True)

m1, m2, m3 = st.columns(3)
with m1:
    st.metric("Grades", len(GRADES))
with m2:
    st.metric("Steps", len(STEPS))
with m3:
    st.metric("GS Patterns", len(GRADES) * len(STEPS))

tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
    t("tab_overview"),
    t("tab_diagram"),
    t("tab_table"),
    t("tab_sim"),
    t("tab_next_year"),
    t("tab_allowance_export"),
    t("tab_employee"),
    t("tab_admin"),
])

# =========================================================
# Tab 1: Overview
# =========================================================
with tab1:
    st.subheader(t("overview_heading"))
    st.write(t("overview_text1"))
    st.write(t("overview_text2"))
    st.write(t("grade_axis"))
    st.subheader(t("rule_heading"))
    st.markdown(f"- {t('rule_ap')}")
    st.markdown(f"- {t('rule_pp')}")
    st.markdown(f"- {t('rule_allow')}")
    st.info(t("rule_promo"))

    ref_df = pd.DataFrame({
        t("grade"): GRADES,
        t("position"): [grade_label(g) for g in GRADES],
        t("next_grade"): [NEXT_GRADE[g] if NEXT_GRADE[g] else "-" for g in GRADES],
    })
    st.subheader(t("grade_table"))
    st.dataframe(ref_df, use_container_width=True, hide_index=True)

    st.subheader(t("ap_pp_heading"))
    ap_pp_df = pd.DataFrame([
        {t("grade"): g, t("ap_label"): format_money(st.session_state.params[g]["ap"]), t("pp_label"): format_money(st.session_state.params[g]["pp"])}
        for g in GRADES
    ])
    st.dataframe(ap_pp_df, use_container_width=True, hide_index=True)

    st.subheader(t("case_study_heading"))
    with st.container(border=True):
        st.markdown(f"**{t('case_new_grad_title')}**")
        st.write(t("case_new_grad_text"))
    with st.container(border=True):
        st.markdown(f"**{t('case_mid_title')}**")
        st.write(t("case_mid_text"))
    st.caption(t("case_note"))

    st.subheader(t("career_25_heading"))
    with st.container(border=True):
        st.markdown(f"**{t('career_newgrad_title')}**")
        st.write(t("career_newgrad_text"))
    with st.container(border=True):
        st.markdown(f"**{t('career_mid_title')}**")
        st.write(t("career_mid_text"))
    st.info(t("career_note"))

    st.subheader(t("career_diagram_heading"))
    if st.session_state.lang == "日本語":
        label_start = "入社"
        label_5 = "5年"
        label_10 = "10年"
        label_15 = "15年"
        label_20 = "20年"
        label_25 = "25年"
        label_promo = "昇格"
    else:
        label_start = "Start"
        label_5 = "Year 5"
        label_10 = "Year 10"
        label_15 = "Year 15"
        label_20 = "Year 20"
        label_25 = "Year 25"
        label_promo = "Promotion"

    st.graphviz_chart(f"""
    digraph G {{
        rankdir=LR;
        node [shape="box", style="rounded,filled", fillcolor="white"];
        N1 [label="G6-S1\\n{label_start}"];
        N2 [label="G6-S5\\n{label_5}"];
        N3 [label="G5B\\n{label_promo}"];
        N4 [label="G5A\\n{label_10}"];
        N5 [label="G4\\n{label_15}"];
        N6 [label="G3\\n{label_20}"];
        N7 [label="G2\\n{label_25}", fillcolor="lightgreen"];
        N1 -> N2 -> N3 -> N4 -> N5 -> N6 -> N7;
    }}
    """)

# =========================================================
# Tab 2: Diagrams
# =========================================================
with tab2:
    st.subheader(t("diagram_heading1"))
    st.write(t("diagram_help1"))
    st.graphviz_chart(grade_step_grid(example_grade, min(example_step, 5)))
    st.subheader(t("diagram_heading2"))
    st.write(t("diagram_help2"))
    st.graphviz_chart(raise_diagram())
    st.subheader(t("diagram_heading3"))
    st.write(t("diagram_help3"))
    sample_result = find_promotion_result(st.session_state.wage_df, st.session_state.params, "G5A", 4)
    if sample_result:
        st.graphviz_chart(promotion_diagram("G5A", 4, sample_result["target_grade"], int(sample_result["target_step"])))
    st.info(t("simple_example_text"))

# =========================================================
# Tab 3: Wage table
# =========================================================
with tab3:
    st.subheader(t("wage_heading"))
    st.caption(t("wage_caption"))
    edited_df = st.data_editor(
        st.session_state.wage_df,
        use_container_width=True,
        num_rows="fixed",
        hide_index=True,
        column_config={
            "Step": st.column_config.NumberColumn("Step", disabled=True),
            **{g: st.column_config.NumberColumn(g, format="%.2f") for g in GRADES},
        },
        key="wage_table_editor",
    )
    st.session_state.wage_df = edited_df
    view_mode = st.radio(t("table_view_mode"), [t("table_mode_raw"), t("table_mode_with_label")], horizontal=True)
    if view_mode == t("table_mode_raw"):
        st.dataframe(display_table_with_formats(st.session_state.wage_df), use_container_width=True, hide_index=True)
    else:
        st.dataframe(display_table_with_gs(st.session_state.wage_df), use_container_width=True, hide_index=True)
    csv_bytes = st.session_state.wage_df.to_csv(index=False).encode("utf-8-sig")
    excel_bytes = make_excel_file(st.session_state.wage_df)
    d1, d2 = st.columns(2)
    with d1:
        st.download_button(t("download_csv"), data=csv_bytes, file_name="wage_table.csv", mime="text/csv", use_container_width=True)
    with d2:
        if excel_bytes is not None:
            st.download_button(t("download_excel"), data=excel_bytes, file_name="wage_table.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        else:
            st.info(t("excel_unavailable"))
    st.caption(t("download_note"))

# =========================================================
# Tab 4: Promotion simulation
# =========================================================
with tab4:
    st.subheader(t("sim_heading"))
    c1, c2, c3 = st.columns(3)
    with c1:
        current_grade = st.selectbox(t("current_grade"), GRADES[:-1], index=0)
    with c2:
        current_step = st.selectbox(t("current_step"), STEPS, index=0)
    with c3:
        is_univ = st.checkbox(t("is_univ"), value=False)
    a1, a2, a3 = st.columns(3)
    with a1:
        adjustment_allowance = st.number_input(t("adjust_allowance"), min_value=0.0, value=0.0, step=100.0)
    with a2:
        university_allowance = st.number_input(t("univ_allowance"), min_value=0.0, value=0.0, step=100.0)
    with a3:
        other_allowance = st.number_input(t("other_allowance"), min_value=0.0, value=0.0, step=100.0)
    if st.button(t("simulate"), use_container_width=True, key="simulate_current"):
        result = find_promotion_result(st.session_state.wage_df, st.session_state.params, current_grade, int(current_step))
        if result is None:
            st.warning(t("no_next_grade"))
        else:
            total_allowance = adjustment_allowance + other_allowance + (university_allowance if is_univ else 0.0)
            final_salary = result["target_salary"] + total_allowance
            r1, r2, r3 = st.columns(3)
            with r1:
                st.metric(t("gs_before"), f"{current_grade}-S{current_step}")
            with r2:
                st.metric(t("gs_after"), f"{result['target_grade']}-S{result['target_step']}")
            with r3:
                st.metric(t("final_salary"), format_money(final_salary))
            st.subheader(t("promotion_flow"))
            st.graphviz_chart(promotion_diagram(current_grade, int(current_step), result["target_grade"], int(result["target_step"])))
            search_df = st.session_state.wage_df[["Step", result["target_grade"]]].copy()
            search_df["Eligible"] = search_df[result["target_grade"]] >= result["minimum_required"]
            st.subheader(t("step_search_result"))
            st.dataframe(search_df, use_container_width=True, hide_index=True)

# =========================================================
# Tab 5: Next year simulation
# =========================================================
with tab5:
    st.subheader(t("next_year_heading"))
    ny1, ny2, ny3 = st.columns(3)
    with ny1:
        next_year_grade = st.selectbox(t("current_grade"), GRADES[:-1], index=0, key="next_year_grade")
    with ny2:
        next_year_step = st.selectbox(t("current_step"), STEPS, index=0, key="next_year_step")
    with ny3:
        apply_next_step_flag = st.checkbox(t("apply_next_step"), value=True)
    if st.button(t("simulate"), use_container_width=True, key="simulate_next_year"):
        next_year_result = find_next_year_promotion_result(st.session_state.wage_df, st.session_state.params, next_year_grade, int(next_year_step), apply_next_step=apply_next_step_flag)
        if next_year_result is None:
            st.warning(t("no_next_grade"))
        else:
            nyr1, nyr2, nyr3 = st.columns(3)
            with nyr1:
                st.metric(t("next_year_current_step"), f"S{next_year_result['next_year_current_step']}")
            with nyr2:
                st.metric(t("next_year_current_salary"), format_money(next_year_result["next_year_current_salary"]))
            with nyr3:
                st.metric(t("next_year_result"), f"{next_year_result['target_grade']}-S{next_year_result['target_step']}")
            st.graphviz_chart(promotion_diagram(next_year_grade, int(next_year_result["next_year_current_step"]), next_year_result["target_grade"], int(next_year_result["target_step"])))

# =========================================================
# Tab 6: Allowance export
# =========================================================
with tab6:
    st.subheader(t("allowance_export_heading"))
    st.write(t("allowance_export_text"))
    ex1, ex2, ex3 = st.columns(3)
    with ex1:
        include_adjustment = st.checkbox(t("include_adjustment_allowance"), value=True)
        adjustment_amount = st.number_input(t("adjustment_allowance_export"), min_value=0.0, value=0.0, step=100.0)
    with ex2:
        include_university = st.checkbox(t("include_university_allowance"), value=False)
        university_amount = st.number_input(t("university_allowance_export"), min_value=0.0, value=0.0, step=100.0)
    with ex3:
        include_other = st.checkbox(t("include_other_allowance"), value=False)
        other_amount = st.number_input(t("other_allowance_export"), min_value=0.0, value=0.0, step=100.0)
    allowance_export_df = build_allowance_export_table(st.session_state.wage_df, include_adjustment, adjustment_amount, include_university, university_amount, include_other, other_amount)
    st.dataframe(allowance_export_df, use_container_width=True, hide_index=True)
    allowance_csv = allowance_export_df.to_csv(index=False).encode("utf-8-sig")
    allowance_excel = make_excel_file(allowance_export_df)
    exd1, exd2 = st.columns(2)
    with exd1:
        st.download_button(t("export_with_allowances_csv"), data=allowance_csv, file_name="wage_table_with_allowances.csv", mime="text/csv", use_container_width=True)
    with exd2:
        if allowance_excel is not None:
            st.download_button(t("export_with_allowances_excel"), data=allowance_excel, file_name="wage_table_with_allowances.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        else:
            st.info(t("excel_unavailable"))

# =========================================================
# Tab 7: Employee roster
# =========================================================
with tab7:
    st.subheader(t("employee_heading"))
    st.markdown(f"<div class='info-card'>{t('employee_text')}</div>", unsafe_allow_html=True)
    template_emp_csv = build_employee_csv_template().to_csv(index=False).encode("utf-8-sig")
    st.download_button(t("employee_template_download"), data=template_emp_csv, file_name="employee_roster_template.csv", mime="text/csv")

    uploaded_emp_csv = None
    if not is_admin():
        st.markdown(f"<div class='warn-card'>{t('admin_only_upload_notice')}</div>", unsafe_allow_html=True)
        st.file_uploader(t("employee_upload"), type=["csv"], key="employee_csv_upload_disabled", disabled=True)
        st.button(t("employee_apply"), use_container_width=True, disabled=True, key="employee_apply_disabled")
    else:
        uploaded_emp_csv = st.file_uploader(t("employee_upload"), type=["csv"], key="employee_csv_upload")
        if uploaded_emp_csv is not None:
            try:
                uploaded_emp_csv.seek(0)
                preview_emp_df = pd.read_csv(uploaded_emp_csv)
                st.markdown(f"**{t('employee_preview')}**")
                st.dataframe(preview_emp_df, use_container_width=True, hide_index=True)
            except Exception:
                st.error(t("employee_import_error"))
        if st.button(t("employee_apply"), use_container_width=True):
            if uploaded_emp_csv is None:
                st.warning(t("employee_import_empty"))
            else:
                try:
                    uploaded_emp_csv.seek(0)
                    imported_emp_df = pd.read_csv(uploaded_emp_csv)
                    st.session_state.employee_roster_df = validate_employee_roster_csv(imported_emp_df)
                    st.success(t("employee_import_success"))
                except Exception as e:
                    st.error(f"{t('employee_import_error')}\n\nDetail: {str(e)}")

    if not st.session_state.employee_roster_df.empty and is_admin():
        active_only = st.checkbox(t("active_only"), value=True)
        roster_university_allowance = st.number_input(t("default_university_allowance"), min_value=0.0, value=0.0, step=100.0)
        roster_apply_next_step = st.checkbox(t("apply_next_step"), value=True, key="roster_apply_next_step")
        payroll_df = build_employee_payroll(
            st.session_state.employee_roster_df,
            st.session_state.wage_df,
            st.session_state.params,
            university_allowance_amount=roster_university_allowance,
            apply_next_step=roster_apply_next_step,
        )
        if active_only:
            active_mask = st.session_state.employee_roster_df["Active"].astype(int).tolist()
            payroll_df = payroll_df[[a == 1 for a in active_mask]].reset_index(drop=True)
        display_payroll_df = payroll_df.copy()
        money_cols = [
            t("employee_basic_pay"),
            t("employee_adjustment_allowance"),
            t("employee_other_allowance"),
            t("employee_university_allowance"),
            t("employee_total_allowance"),
            t("employee_total_pay"),
            t("employee_next_year_basic_pay"),
        ]
        for col in money_cols:
            if col in display_payroll_df.columns:
                display_payroll_df[col] = display_payroll_df[col].apply(lambda x: format_money(x) if pd.notna(x) else "-")
        if "_active" in display_payroll_df.columns:
            display_payroll_df = display_payroll_df.drop(columns=["_active"])
        st.dataframe(display_payroll_df, use_container_width=True, hide_index=True)
        emp_csv = payroll_df.drop(columns=["_active"]).to_csv(index=False).encode("utf-8-sig")
        emp_excel = make_excel_file(payroll_df.drop(columns=["_active"]))
        e1, e2 = st.columns(2)
        with e1:
            st.download_button(t("employee_export_csv"), data=emp_csv, file_name="employee_payroll.csv", mime="text/csv", use_container_width=True)
        with e2:
            if emp_excel is not None:
                st.download_button(t("employee_export_excel"), data=emp_excel, file_name="employee_payroll.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            else:
                st.info(t("excel_unavailable"))
    elif not is_admin():
        st.info("Employee CSV data is hidden for viewers. Only admins can upload, view, and export employee-level data.")

# =========================================================
# Tab 8: Admin
# =========================================================
with tab8:
    st.subheader(t("admin_heading"))
    st.markdown(f"<div class='info-card'>{t('admin_text')}</div>", unsafe_allow_html=True)
    st.markdown(f"<div class='warn-card'>{t('warning_rebuild')}</div>", unsafe_allow_html=True)

    if st.session_state.lang == "日本語":
        st.info("表形式で Base / AP / PP をまとめて確認できます。管理者は編集・再生成、Viewer は閲覧のみです。")
        st.caption("AP・PP もこの表で直接変更できます。変更後に『再生成』を押すと賃金テーブルへ反映されます。")
        grade_col_label = "グレード"
        position_col_label = "役職"
    else:
        st.info("You can review Base / AP / PP in a spreadsheet-style grid. Admins can edit and rebuild, while viewers can only view.")
        st.caption("AP and PP can also be updated directly in this grid. Click Rebuild after editing to apply the changes to the wage table.")
        grade_col_label = "Grade"
        position_col_label = "Position"

    settings_editor_df = pd.DataFrame([
        {
            grade_col_label: g,
            position_col_label: grade_label(g),
            t("base_salary"): float(st.session_state.params[g]["base"]),
            t("ap"): float(st.session_state.params[g]["ap"]),
            t("pp"): float(st.session_state.params[g]["pp"]),
        }
        for g in GRADES
    ])

    edited_settings_df = st.data_editor(
        settings_editor_df,
        hide_index=True,
        use_container_width=True,
        num_rows="fixed",
        disabled=[grade_col_label, position_col_label] if is_admin() else True,
        column_config={
            grade_col_label: st.column_config.TextColumn(width="small"),
            position_col_label: st.column_config.TextColumn(width="medium"),
            t("base_salary"): st.column_config.NumberColumn(min_value=0.0, step=100.0, format="%.2f"),
            t("ap"): st.column_config.NumberColumn(min_value=0.0, step=50.0, format="%.2f"),
            t("pp"): st.column_config.NumberColumn(min_value=0.0, step=50.0, format="%.2f"),
        },
        key="admin_settings_grid",
    )

    if is_admin():
        tmp_params = {}
        for _, row in edited_settings_df.iterrows():
            g = str(row[grade_col_label]).strip()
            tmp_params[g] = {
                "base": float(row[t("base_salary")]),
                "ap": float(row[t("ap")]),
                "pp": float(row[t("pp")]),
            }

        b1, b2 = st.columns(2)
        with b1:
            if st.button(t("rebuild"), use_container_width=True):
                try:
                    save_and_rebuild(tmp_params)
                    st.success(t("success_rebuild"))
                    st.rerun()
                except Exception as e:
                    st.error(f"{t('supabase_save_error')}\n\nDetail: {str(e)}")
        with b2:
            if st.button(t("reset"), use_container_width=True):
                try:
                    reset_params = {k: v.copy() for k, v in DEFAULT_PARAMS.items()}
                    save_and_rebuild(reset_params)
                    st.success(t("success_reset"))
                    st.rerun()
                except Exception as e:
                    st.error(f"{t('supabase_save_error')}\n\nDetail: {str(e)}")

        st.markdown("---")
        st.subheader(t("csv_import_heading"))
        st.markdown(f"<div class='info-card'>{t('csv_import_text')}</div>", unsafe_allow_html=True)
        template_csv = build_settings_csv_template().to_csv(index=False).encode("utf-8-sig")
        st.download_button(t("csv_template_download"), data=template_csv, file_name="wage_table_settings_template.csv", mime="text/csv")

        uploaded_csv = st.file_uploader(t("csv_upload"), type=["csv"], key="settings_csv_upload")
        if uploaded_csv is not None:
            try:
                uploaded_csv.seek(0)
                preview_df = pd.read_csv(uploaded_csv)
                st.markdown(f"**{t('csv_preview_heading')}**")
                st.dataframe(preview_df, use_container_width=True, hide_index=True)
            except Exception:
                st.error(t("csv_import_error"))
        if st.button(t("csv_apply"), use_container_width=True):
            if uploaded_csv is None:
                st.warning(t("csv_import_empty"))
            else:
                try:
                    uploaded_csv.seek(0)
                    imported_df = pd.read_csv(uploaded_csv)
                    new_params = validate_imported_settings_csv(imported_df)
                    save_and_rebuild(new_params)
                    st.success(t("csv_import_success"))
                    st.rerun()
                except Exception as e:
                    st.error(f"{t('csv_import_error')}\n\nDetail: {str(e)}")
    else:
        st.info("Viewer mode: settings can be viewed here, but only admins can rebuild tables or import CSV files.")

st.markdown("---")
st.caption("Created for bilingual wage table explanation, visual guidance, promotion simulation, employee roster import/export, CSV import, and Supabase persistence in Streamlit.")
