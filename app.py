import io
from datetime import datetime
from typing import Optional

import pandas as pd
import requests
import streamlit as st

st.set_page_config(
    page_title="Job Matrix Manager",
    layout="wide",
)

st.markdown(
    """
<style>
.block-container {
    padding-top: 1rem;
    padding-bottom: 1.5rem;
    max-width: 1400px;
}
html, body, [class*="css"]  {
    font-family: Arial, Helvetica, sans-serif;
}
div[data-testid="stMetric"] {
    border: 1px solid #d9dde3;
    padding: 0.8rem;
    border-radius: 0.35rem;
    background: #ffffff;
    box-shadow: none;
}
div[data-testid="stMetric"] label,
div[data-testid="stMetricValue"] {
    color: #1f2937;
}
.small-note {
    font-size: 0.86rem;
    color: #4b5563;
}
.grade-box {
    white-space: pre-wrap;
    border: 1px solid #d9dde3;
    padding: 0.9rem 1rem;
    background: #ffffff;
    border-radius: 0.35rem;
    line-height: 1.7;
}
.handout-box {
    white-space: pre-wrap;
    border: 1px solid #d9dde3;
    padding: 1rem 1.1rem;
    background: #ffffff;
    border-radius: 0.35rem;
    line-height: 1.85;
}
.doc-section-title {
    font-weight: 700;
    margin-top: 0.35rem;
}
div[data-testid="stExpander"] {
    border: 1px solid #d9dde3;
    border-radius: 0.35rem;
    background: #ffffff;
}
h1, h2, h3 {
    letter-spacing: 0.01em;
}
</style>
""",
    unsafe_allow_html=True,
)

st.title("Job Matrix Manager")
st.caption(
    "部門・グレード別の業務概要、グレード定義、KPIを管理し、配布用のWord/PDFを自動出力できます。"
)

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
    "kpi",
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

DEFAULT_GRADE_OVERVIEW = {
    "G6": (
        "上位者の指示および方針に基づき、担当業務を正確かつ着実に遂行することが求められるグレードである。\n"
        "組織の一員として基本的な役割と責任を果たし、誠実な業務遂行姿勢、基本的な問題対応力および周囲との協働姿勢を備えることが求められる。"
    ),
    "G5B": (
        "G6からのステップアップ段階として、一定の自律性をもって業務を遂行し、期待される役割と責任を主体的に果たすことが求められるグレードである。\n"
        "状況に応じた適切な判断と行動により、担当業務の安定運用に貢献するとともに、周囲との協力を通じて組織成果の向上に寄与することが期待される。"
    ),
    "G5A": (
        "G5ランクの上位段階として、求められる職能基準を安定的に満たし、継続的に成果を発揮することが求められるグレードである。\n"
        "主体的な行動と的確な判断により担当領域を着実に遂行し、業務品質および効率の向上を通じて組織へ継続的に貢献することが期待される。"
    ),
    "G4": (
        "課単位の目標達成に責任を持ち、業務および人材の管理を通じて組織運営に貢献することが求められるグレードである。\n"
        "計画立案、進捗管理、メンバー指導および関係者調整を通じて、課全体の成果最大化を担うことが期待される。"
    ),
    "G3": (
        "事業部単位の目標達成に責任を持ち、組織運営および人材マネジメントを統括することが求められるグレードである。\n"
        "年度方針に基づく計画立案と実行を通じて、事業部全体の成果創出と組織運営の安定化を担うことが期待される。"
    ),
    "G2": (
        "支店または拠点全体の目標達成に責任を持ち、経営方針に基づく戦略の立案および実行を担うグレードである。\n"
        "中期的な視点で組織運営を統括し、持続的な成長と全体最適の実現に向けて、組織全体を牽引することが期待される。"
    ),
}


DEPARTMENT_ROLE_TOPICS = {
    "人事課": "採用、労務管理、人材配置、教育支援等の人事関連業務",
    "総務課": "総務機能に関する庶務業務、文書管理、来客対応等の業務",
    "経理課": "伝票処理、証憑管理、支払対応、会計記録整備等の経理業務",
    "財務課": "資金管理、予算管理、支払準備、金融機関対応支援等の財務業務",
    "情報管理課": "データ整備、ファイル管理、システム運用支援、情報管理に関する業務",
    "エリア運営課": "現場運営、巡回確認、作業調整、進捗把握等のエリア運営業務",
    "設備・運搬課": "設備保守、車両運用、資材運搬、安全確認等の設備・運搬業務",
    "ナーサリー課": "苗木育成、灌水、除草、記録管理、品質確認等の育苗業務",
    "マーケティング課": "市場調査、資料作成、対外向け情報整理、事業訴求支援等の業務",
    "プロセス管理課": "業務手順の整備、工程管理、品質・安全確認、改善推進等の業務",
    "インサイドコミュニケーション課": "社内周知、社内施策運営、情報発信、組織浸透支援等の業務",
    "アウトサイドコミュニケーション課": "対外発信、渉外対応、関係者調整、地域連携支援等の業務",
}


def make_default_role_summary(department: str, grade: str, grade_name: str) -> str:
    topic = DEPARTMENT_ROLE_TOPICS.get(department, f"{department}に関する業務")

    grade_templates = {
        "G6": (
            "{topic}を担当します。"
            "定められた手順および上位者の指示に基づき、日常業務を正確かつ着実に遂行し、関係部署と連携しながら円滑な業務運営を支えます。"
            "基礎実務の安定運用を通じて、組織運営の土台づくりに貢献します。"
        ),
        "G5B": (
            "{topic}を担当します。"
            "一定の自律性をもって担当業務を遂行し、状況に応じた判断と関係部署との連携を通じて、日常業務の安定運用を支えます。"
            "担当領域における品質維持と業務の確実な遂行を通じて、組織成果の向上に貢献します。"
        ),
        "G5A": (
            "{topic}を担当します。"
            "担当領域を横断的に捉え、状況に応じた判断のもとで業務を主体的に遂行し、関係部署と連携しながら円滑な運営を支えます。"
            "安定した成果の発揮と業務改善への取り組みを通じて、組織全体の業務品質向上に貢献します。"
        ),
        "G4": (
            "{topic}に関する業務運営を担当します。"
            "課としての目標達成に向けて、業務の進捗管理、人員配置、メンバー指導および関係者との調整を行い、日常運営を統括します。"
            "計画的な運営と継続的な改善を通じて、課全体の成果最大化に貢献します。"
        ),
        "G3": (
            "{topic}に関する事業部運営を担当します。"
            "事業部目標の達成に向けて、計画立案、業務管理、人材マネジメントおよび関係部署との調整を統括し、部門運営を推進します。"
            "重点課題への対応と組織運営の高度化を通じて、事業部全体の成果創出に貢献します。"
        ),
        "G2": (
            "{topic}に関する拠点全体の運営を担当します。"
            "経営方針に基づき、中期的な視点で方針立案、組織運営、資源配分および関係部門との連携を統括し、全体最適を推進します。"
            "重要課題への意思決定と体制整備を通じて、拠点全体の持続的な成長と価値創出に貢献します。"
        ),
    }
    return grade_templates.get(grade, "{topic}を担当します。日常業務を正確かつ効率的に遂行し、関係部署と連携しながら業務の円滑な運営を支えます。また、安定したオペレーションの実現に貢献します。").format(topic=topic)



def make_default_role_summary(department: str, grade: str, grade_name: str) -> str:
    data = {
        "人事課": {
            "G6": "採用・勤怠・人事データ入力などの定型実務を担当する。上位者の指示に沿って正確に処理し、必要な情報を遅れなく整備する。社内の基礎的な人事運営を支える役割を担う。",
            "G5B": "採用実務や従業員管理業務を一定の自律性をもって遂行する。日常業務を安定運用しながら、関係部署との連携や進捗管理にも主体的に関わる。担当領域の品質向上に継続して貢献する。",
            "G5A": "採用、労務、教育関連業務を横断的に支援する。状況に応じて優先順位を判断し、周囲を支えながら実務の改善提案も行う。安定した成果を継続的に出す中核人材として機能する。",
            "G4": "人事オペレーション全体を監督し、採用・労務・教育実務の進捗を管理する。メンバーへの指導や業務分担を行い、課としての成果達成を支える。現場部門との調整を通じて人事課題の解決を進める。",
            "G3": "人事課全体の計画立案、採用・配置・制度運用を統括する。組織運営の観点から優先課題を整理し、経営や現場と連携しながら施策を推進する。部門成果と人材基盤の両面を管理する役割を担う。",
            "G2": "経営方針に基づき人事戦略や組織体制整備を推進する。拠点や部門をまたぐ人材課題を整理し、制度・配置・育成の方向性を定める。全社視点で持続的な組織運営を支える役割を担う。",
        },
        "総務課": {
            "G6": "備品管理、文書管理、来客対応などの総務実務を担当する。日常の依頼や定型対応を正確に処理し、社内環境の維持を支える。拠点運営の土台となる基礎業務を担う。",
            "G5B": "庶務業務や社内サポート業務を自律的に遂行する。関連部門と連携しながら日常運営の安定化を図り、必要な対応を漏れなく進める。総務機能の継続的な品質維持に貢献する。",
            "G5A": "総務実務全般を理解し、複数業務を横断して支援する。現場運営上の課題を捉え、手順改善や調整対応も担う。周囲を支えながら総務体制の安定運用に寄与する。",
            "G4": "総務チームの日常運営を監督し、進捗管理やメンバー支援を行う。社内ルールの浸透や課題是正を推進し、業務品質を維持する。部署全体の円滑な運営を支える管理役を担う。",
            "G3": "総務機能全体の運用管理とルール整備を統括する。拠点運営の基盤強化に向けて優先課題を整理し、他部門と連携して改善施策を進める。組織運営を安定させる中核的な役割を担う。",
            "G2": "拠点運営基盤の整備、ガバナンス強化、管理部門連携を推進する。中長期的な視点で総務機能の方向性を定め、全体最適の観点から制度や運用を整える。組織全体の安定運営を牽引する。",
        },
        "経理課": {
            "G6": "伝票入力、証憑整理、支払データ作成などの基礎経理業務を担当する。定められたルールに沿って正確に処理し、会計記録の整備を支える。日々の経理運営を下支えする役割を担う。",
            "G5B": "日次・月次経理の実務を自律的に運用する。数値の正確性や処理期限を意識しながら、定型業務の安定化に貢献する。関連資料の整合確認や報告補助も担う。",
            "G5A": "経理実務全般に加え、締め処理や照合作業の改善も担う。複数データを横断的に確認し、必要に応じて対応方法を見直す。課内の業務品質向上に継続的に寄与する。",
            "G4": "経理業務の進捗管理やメンバー確認を行い、締め処理全体を監督する。課題や差異を把握して是正し、安定した月次運営を支える。チーム成果を意識した管理業務を担う。",
            "G3": "月次・年次決算、会計管理、税務対応を統括する。経理課全体の方針と優先順位を整理し、必要な是正や改善を主導する。会社の数値基盤を支える責任を担う。",
            "G2": "財務・経理方針を踏まえた管理体制強化と経営報告を推進する。拠点全体の数値管理水準を高め、重要課題への対応方針を定める。経営判断に資する会計・管理基盤を整備する。",
        },
        "財務課": {
            "G6": "資金関連データの整理や支払準備などの補助業務を担当する。必要資料を正確に整備し、資金管理実務が円滑に進むよう支援する。日々の財務運営を下支えする役割を担う。",
            "G5B": "資金繰り関連資料の作成や銀行対応補助を自律的に進める。日常の数値管理や確認業務を安定運用し、必要な情報を適切に共有する。財務実務の精度と継続性に貢献する。",
            "G5A": "財務管理資料作成、資金計画支援、数値分析を横断的に担う。状況に応じて必要情報を整理し、改善提案や対応方針の補助も行う。課内の意思決定を支える実務中核として機能する。",
            "G4": "資金管理業務を監督し、実務進捗や報告内容を管理する。課題を把握して関係者と調整し、財務運営の安定化を図る。チーム全体での成果達成を支える管理役を担う。",
            "G3": "資金繰り、予算実績管理、金融機関対応を統括する。財務課題を整理し、経営や関連部門と連携しながら必要な施策を進める。経営判断に直結する数値管理責任を担う。",
            "G2": "財務戦略、投資判断支援、経営数値管理を推進する。中長期視点で資金・収支・投資の方向性を整え、全体最適を意識して意思決定を支える。組織全体の財務健全性を牽引する。",
        },
        "情報管理課": {
            "G6": "データ入力、ファイル管理、IT機器運用補助を担当する。ルールに沿って情報を整理し、日常のシステム利用が滞りなく進むよう支援する。基礎的な情報管理体制を支える役割を担う。",
            "G5B": "システム運用補助、情報整備、ユーザー対応を自律的に進める。必要な問い合わせに適切に対応し、日常運用の安定化を図る。担当領域の情報品質向上に継続して貢献する。",
            "G5A": "業務データ管理、IT活用支援、簡易改善提案を横断的に担う。現場課題を把握しながら、より使いやすい運用や整理方法を検討する。情報活用の実効性を高める中核人材として機能する。",
            "G4": "情報管理実務を監督し、システム運用管理やメンバー支援を行う。課題発生時の整理と対応方針の共有を進め、業務継続性を確保する。チームとしての管理品質向上を担う。",
            "G3": "情報管理体制、業務システム運用、データ活用推進を統括する。部門課題を整理し、業務効率や統制の観点から改善を主導する。組織全体の情報基盤を支える役割を担う。",
            "G2": "情報戦略、デジタル化推進、情報統制基盤の整備を担う。拠点横断での運用ルールや投資優先順位を定め、全体最適の観点から体制を強化する。経営に資する情報環境の構築を推進する。",
        },
        "エリア運営課": {
            "G6": "現場オペレーション、巡回、作業記録などの基本業務を担当する。上位者の指示に沿って現場状況を正確に把握し、必要情報を記録・報告する。日々のエリア運営を支える役割を担う。",
            "G5B": "担当エリアの運営実務を安定的に遂行する。現場の進捗や異常を主体的に把握し、関係者と連携して必要対応を進める。担当範囲の安定運営に継続して貢献する。",
            "G5A": "現場の進捗確認、作業調整、改善提案を横断的に担う。複数条件を踏まえて優先順位を判断し、より円滑な運営方法を検討する。現場品質と生産性向上を支える中核として機能する。",
            "G4": "現場チームの管理、日々の進捗監督、安全・品質確認を担う。人員配置や作業指示を通じて課単位の成果達成を支える。現場課題の早期把握と是正を主導する。",
            "G3": "複数エリアの運営管理、作業計画、人員配置を統括する。全体の状況を俯瞰し、資源配分や進捗管理の方針を定める。エリア運営全体の安定性と成果責任を担う。",
            "G2": "農地運営方針、拠点横断の運営最適化、事業計画達成を推進する。中期視点で人員・設備・現場体制の方向性を整理し、全体最適の観点から意思決定を行う。事業の現場基盤を牽引する。",
        },
        "設備・運搬課": {
            "G6": "設備点検補助、資材運搬、車両運用補助を担当する。日常点検や運搬実務を安全に実施し、異常時は速やかに報告する。設備・物流の基礎運営を支える役割を担う。",
            "G5B": "設備・車両の運用実務と日常点検を自律的に進める。運行や保守に必要な情報を適切に管理し、安定稼働を支える。現場の安全性と継続性の維持に貢献する。",
            "G5A": "設備保全、運搬計画支援、トラブル初期対応を横断的に担う。状況に応じて優先順位を判断し、効率的な運用方法の改善も行う。設備・物流機能の品質向上を支える。",
            "G4": "設備保守・運搬業務の監督、安全管理、進捗管理を担う。メンバーへの指示や調整を通じて課としての成果達成を支える。設備や物流上の課題を早期に整理し、是正を進める。",
            "G3": "設備管理計画、保全計画、運搬体制を統括する。全体の稼働状況や課題を踏まえて優先施策を定め、関係部署と連携して改善を進める。安定した供給・運搬体制の責任を担う。",
            "G2": "設備投資方針、物流効率化、全体インフラ最適化を推進する。中長期視点で保全・更新・運搬体制の方向性を整理し、拠点全体の基盤強化を進める。組織運営を支えるインフラ戦略を担う。",
        },
        "ナーサリー課": {
            "G6": "苗の育成、灌水、除草、記録などの基礎栽培業務を担当する。定められた手順に沿って日々の育苗作業を丁寧に実施し、状態変化を記録・報告する。育苗現場の基礎運営を支える役割を担う。",
            "G5B": "苗木育成業務を安定運用し、育成状況の記録管理を担う。日々の状態を主体的に確認し、必要な対応を上位者や関係者と連携して進める。品質維持に継続して貢献する。",
            "G5A": "育苗計画支援、品質確認、作業改善提案を横断的に担う。現場状況に応じて優先順位を判断し、より良い育苗体制づくりを支援する。生産性と品質の両立を支える中核人材として機能する。",
            "G4": "ナーサリー現場の監督、人員配置、品質・安全管理を担う。進捗や育成状況を把握し、課としての目標達成に向けて現場運営を整える。メンバー指導や課題是正も主導する。",
            "G3": "育苗計画、数量管理、現場運営全体を統括する。植栽計画や他部門との連携を踏まえ、必要な人員・資材・体制を整える。ナーサリー課全体の成果責任を担う。",
            "G2": "育苗戦略、生産性向上、植栽計画との連携を推進する。中長期視点で体制・品質・数量の方向性を定め、拠点全体の育苗基盤を強化する。事業全体を支える生産戦略を担う。",
        },
        "マーケティング課": {
            "G6": "市場情報整理、資料作成補助、基礎調査業務を担当する。必要データを正確に集約し、社内外向け資料作成を支援する。営業や企画判断の基礎となる情報整備を担う。",
            "G5B": "市場調査、資料更新、営業支援情報の整理を自律的に進める。必要な情報を分かりやすく整理し、関係者が使いやすい形で共有する。日常的な市場把握の品質向上に貢献する。",
            "G5A": "調査分析、社外資料作成、施策提案支援を横断的に担う。市場動向を踏まえて改善提案や情報発信の工夫を行い、対外訴求力を高める。部門の実務中核として継続的な成果を出す。",
            "G4": "調査進捗管理、メンバー支援、施策実行管理を担う。必要な情報収集や発信計画を整理し、チームとしての成果達成を支える。課題を把握しながら実行力の高い運営を進める。",
            "G3": "市場分析、販売戦略支援、対外訴求計画を統括する。事業の方向性と市場動向を結びつけ、必要な施策を立案・推進する。対外コミュニケーションの成果責任を担う。",
            "G2": "事業戦略と連動した市場開拓・発信方針を推進する。中長期視点で重点市場や訴求方針を定め、拠点や関係部署を横断して実行基盤を整える。会社の対外価値向上を牽引する。",
        },
        "プロセス管理課": {
            "G6": "業務記録、手順遵守、現場データ収集などを担当する。定められたプロセスに沿って作業し、必要な情報を正確に記録・報告する。標準化された業務運営の基礎を支える役割を担う。",
            "G5B": "工程管理補助、手順管理、定型レポート作成を自律的に進める。日常運用の中で異常や改善点を把握し、必要な情報共有を行う。安定した工程運営の維持に貢献する。",
            "G5A": "工程改善、標準化支援、異常対応補助を横断的に担う。業務の流れを見ながら改善提案を行い、品質・安全・効率の向上を支援する。部門横断での改善実務にも関わる。",
            "G4": "工程進捗・品質・安全の監督と現場是正を担う。メンバーへの指示や確認を通じて、課単位での安定運営を実現する。問題発生時には原因整理と対応推進を行う。",
            "G3": "業務プロセス設計、改善推進、管理指標運用を統括する。全体の流れを俯瞰し、重点課題を整理して標準化や是正を主導する。継続改善の成果責任を担う。",
            "G2": "全体プロセス最適化、標準化方針、継続改善を推進する。中長期視点で業務構造を見直し、拠点横断で再現性の高い運営体制を整える。組織全体の生産性向上を牽引する。",
        },
        "インサイドコミュニケーション課": {
            "G6": "社内連絡、情報配信補助、文書整理を担当する。必要な情報を正確に整理し、社内へ分かりやすく伝達できるよう支援する。円滑な社内コミュニケーションの基礎運営を担う。",
            "G5B": "社内周知、イベント運営補助、情報整理を自律的に進める。関係者と連携しながら社内施策を滞りなく運用し、必要情報の浸透を支える。日常的な社内活性化に貢献する。",
            "G5A": "社内施策支援、社内発信、エンゲージメント向上施策を横断的に担う。状況に応じて伝え方や運用方法を工夫し、社内の理解促進を進める。組織文化醸成を支える中核人材として機能する。",
            "G4": "社内コミュニケーション施策の進捗管理と運営を担う。関係者調整やメンバー支援を通じて、課としての成果達成を支える。情報浸透や施策実行の品質を管理する。",
            "G3": "社内広報、社内文化醸成、情報浸透施策を統括する。組織課題や経営メッセージを踏まえ、必要な施策を立案・推進する。社内の一体感形成に責任を持つ。",
            "G2": "組織活性化、経営メッセージ浸透、社内連携強化を推進する。中長期視点で社内文化や発信方針の方向性を定め、全体最適の観点から施策を統括する。組織力向上を牽引する。",
        },
        "アウトサイドコミュニケーション課": {
            "G6": "対外資料整理、問い合わせ対応補助、情報更新を担当する。定められたルールに沿って情報を整備し、社外対応が円滑に進むよう支援する。対外発信の基礎運営を担う。",
            "G5B": "社外連絡、情報発信補助、対外資料管理を自律的に進める。必要な情報を適切に整理し、関係者と連携して円滑な対外対応を支える。日常的な社外コミュニケーション品質向上に貢献する。",
            "G5A": "社外発信支援、関係者調整、広報実務を横断的に担う。状況に応じて伝え方や段取りを工夫し、対外発信の効果を高める。組織の対外信頼構築を支える中核人材として機能する。",
            "G4": "対外コミュニケーション施策の運営管理を担う。メンバー支援や関係者調整を通じて、課単位での成果達成を支える。対外対応の品質と進捗を管理する役割を担う。",
            "G3": "広報・渉外・地域連携対応を統括する。会社方針や地域との関係性を踏まえて必要施策を推進し、対外信頼の維持向上を図る。対外コミュニケーション全体の責任を担う。",
            "G2": "対外関係戦略、地域・行政・関係者との信頼構築を推進する。中長期視点で対外方針を定め、全社的な連携の中で重要関係を強化する。組織の外部価値を高める役割を担う。",
        },
    }
    if department in data and grade in data[department]:
        return data[department][grade]
    return (
        f"{grade_name}として{department}の実務を担う。"
        "担当領域の業務を正確かつ継続的に遂行し、関係者と連携しながら成果達成に貢献する。"
        "必要に応じて改善や調整にも関わり、組織運営を支える役割を果たす。"
    )


def make_default_kpi(grade: str) -> str:
    return {
        "G6": "業務達成率、処理件数、正確性、期限遵守率",
        "G5B": "担当業務達成率、品質、改善件数、納期遵守率",
        "G5A": "成果達成率、改善提案数、関連部署連携品質",
        "G4": "チーム達成率、進捗遵守率、品質指標、人材育成状況",
        "G3": "課目標達成率、予算遵守率、改善成果、組織運営状況",
        "G2": "部門目標達成率、戦略進捗、横断課題解決、収益・効率改善",
    }[grade]


def generate_default() -> pd.DataFrame:
    rows = []
    for dept in DEPARTMENTS:
        for grade in GRADE_ORDER:
            rows.append(
                {
                    "department": dept,
                    "grade": grade,
                    "grade_name": GRADES[grade],
                    "role_summary": make_default_role_summary(dept, grade, GRADES[grade]),
                    "kpi": make_default_kpi(grade),
                }
            )
    return pd.DataFrame(rows)


def get_default_grade_master() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {"grade": grade, "grade_name": GRADES[grade], "grade_overview": DEFAULT_GRADE_OVERVIEW[grade]}
            for grade in GRADE_ORDER
        ]
    )


def ensure_main_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    defaults = {
        "department": "",
        "grade": "",
        "grade_name": "",
        "role_summary": "",
        "kpi": "",
    }
    for col, default in defaults.items():
        if col not in out.columns:
            out[col] = default
    return out[["department", "grade", "grade_name", "role_summary", "kpi"]]


def ensure_grade_master(df: Optional[pd.DataFrame]) -> pd.DataFrame:
    base = get_default_grade_master()
    if df is None or df.empty:
        return base
    out = df.copy()
    for col in ["grade", "grade_name", "grade_overview"]:
        if col not in out.columns:
            out[col] = ""
    out = out[["grade", "grade_name", "grade_overview"]].copy()
    out["grade"] = out["grade"].astype(str).str.strip()
    out = out.drop_duplicates(subset=["grade"], keep="last")
    merged = base[["grade"]].merge(out, on="grade", how="left", suffixes=("", "_uploaded"))
    merged["grade_name"] = merged["grade_name"].fillna(merged["grade"].map(GRADES))
    merged["grade_overview"] = merged["grade_overview"].fillna(merged["grade"].map(DEFAULT_GRADE_OVERVIEW))
    merged["grade_name"] = merged["grade_name"].replace("", pd.NA).fillna(merged["grade"].map(GRADES))
    merged["grade_overview"] = merged["grade_overview"].replace("", pd.NA).fillna(merged["grade"].map(DEFAULT_GRADE_OVERVIEW))
    merged["grade_sort"] = merged["grade"].map({g: i for i, g in enumerate(GRADE_ORDER)})
    return merged.sort_values("grade_sort").drop(columns=["grade_sort"])


def attach_grade_overview(main_df: pd.DataFrame, grade_master: pd.DataFrame) -> pd.DataFrame:
    df = ensure_main_columns(main_df)
    gm = ensure_grade_master(grade_master)
    merged = df.merge(gm[["grade", "grade_overview"]], on="grade", how="left")
    merged["grade_overview"] = merged["grade_overview"].fillna(merged["grade"].map(DEFAULT_GRADE_OVERVIEW))
    return merged


def save_snapshot(main_df: pd.DataFrame, grade_master_df: pd.DataFrame, memo: str) -> None:
    payload = {
        "data": ensure_main_columns(main_df).to_dict(orient="records"),
        "grade_master": ensure_grade_master(grade_master_df).to_dict(orient="records"),
        "memo": memo,
        "created_at": datetime.now().isoformat(),
    }
    rest_post("job_matrix", payload)


def load_history() -> list:
    return rest_get("job_matrix", params={"select": "*", "order": "created_at.desc"})


def load_latest_snapshot():
    history = load_history()
    if not history:
        return None, None

    latest = history[0]
    latest_df = ensure_main_columns(pd.DataFrame(latest.get("data", [])))
    latest_grade_master = ensure_grade_master(pd.DataFrame(latest.get("grade_master", [])))
    return latest_df, latest_grade_master


def normalize_text(value):
    if pd.isna(value):
        return None
    text = str(value).strip()
    return text if text != "" else None


def dataframe_to_csv_bytes(df: pd.DataFrame) -> bytes:
    export_df = ensure_main_columns(df.copy())
    return export_df.to_csv(index=False).encode("utf-8-sig")


def grade_master_to_csv_bytes(df: pd.DataFrame) -> bytes:
    export_df = ensure_grade_master(df.copy())
    return export_df.to_csv(index=False).encode("utf-8-sig")


def dataframe_to_excel_bytes(df: pd.DataFrame, grade_master_df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    export_df = attach_grade_overview(df.copy(), grade_master_df.copy())
    export_df["grade_sort"] = export_df["grade"].map({g: i for i, g in enumerate(GRADE_ORDER)})
    export_df = export_df.sort_values(["department", "grade_sort"]).drop(columns=["grade_sort"])
    grade_export = ensure_grade_master(grade_master_df.copy())

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="job_matrix")
        grade_export.to_excel(writer, index=False, sheet_name="grade_master")
    output.seek(0)
    return output.getvalue()


def minimal_template_csv_bytes() -> bytes:
    template_df = pd.DataFrame(columns=["department", "grade", "grade_name", "role_summary", "kpi"])
    return template_df.to_csv(index=False).encode("utf-8-sig")


def minimal_grade_template_csv_bytes() -> bytes:
    template_df = pd.DataFrame(columns=["grade", "grade_name", "grade_overview"])
    return template_df.to_csv(index=False).encode("utf-8-sig")


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

    updated_df = ensure_main_columns(base_df.copy())
    import_df = upload_df.copy()

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
                changes.append(
                    {
                        "department": row["department"],
                        "grade": row["grade"],
                        "column": col,
                        "before": old_val,
                        "after": new_val,
                    }
                )

    updated_df = updated_df.drop(columns=["_merge_key"])
    log_df = pd.DataFrame(changes)
    return updated_df, log_df


def merge_grade_master_update(base_df: pd.DataFrame, upload_df: pd.DataFrame):
    if "grade" not in upload_df.columns:
        raise ValueError("CSVに必須列 'grade' がありません。")

    updated_df = ensure_grade_master(base_df.copy())
    import_df = upload_df.copy()
    for col in ["grade", "grade_name", "grade_overview"]:
        if col not in import_df.columns:
            import_df[col] = None

    updated_df["grade"] = updated_df["grade"].astype(str).str.strip()
    import_df["grade"] = import_df["grade"].astype(str).str.strip()
    import_df = import_df.drop_duplicates(subset=["grade"], keep="last")
    import_map = import_df.set_index("grade").to_dict(orient="index")

    changes = []
    for idx, row in updated_df.iterrows():
        grade = row["grade"]
        if grade not in import_map:
            continue
        source = import_map[grade]
        for col in ["grade_name", "grade_overview"]:
            old_val = normalize_text(row.get(col))
            new_val = normalize_text(source.get(col))
            if new_val is not None and new_val != old_val:
                updated_df.at[idx, col] = new_val
                changes.append(
                    {
                        "grade": grade,
                        "column": col,
                        "before": old_val,
                        "after": new_val,
                    }
                )

    log_df = pd.DataFrame(changes)
    return updated_df, log_df


def update_master_from_editor(master_df: pd.DataFrame, edited_df: pd.DataFrame) -> pd.DataFrame:
    current = ensure_main_columns(master_df.copy())
    edited = ensure_main_columns(edited_df.copy())
    edited_map = edited.set_index(["department", "grade"]).to_dict(orient="index")
    for idx, row in current.iterrows():
        key = (row["department"], row["grade"])
        if key in edited_map:
            for col in UPDATE_COLUMNS:
                current.at[idx, col] = edited_map[key].get(col, row[col])
    return current


def department_summary(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    df = ensure_main_columns(df.copy())
    for dept in DEPARTMENTS:
        sub = df[df["department"] == dept]
        rows.append(
            {
                "department": dept,
                "rows": len(sub),
                "filled_role_summary": int(sub["role_summary"].astype(str).str.strip().ne("").sum()),
                "filled_kpi": int(sub["kpi"].astype(str).str.strip().ne("").sum()),
            }
        )
    return pd.DataFrame(rows)


def split_kpi_items(kpi_text: str) -> list[str]:
    if pd.isna(kpi_text) or str(kpi_text).strip() == "":
        return []
    text = str(kpi_text).replace("\n", "、")
    parts = [p.strip(" ・-\t") for p in text.replace(",", "、").split("、")]
    return [p for p in parts if p]


def make_handout_dict(row: pd.Series) -> dict:
    department = str(row.get("department", "")).strip()
    grade = str(row.get("grade", "")).strip()
    grade_name = str(row.get("grade_name", "")).strip()
    role_summary = str(row.get("role_summary", "")).strip()
    grade_overview = str(row.get("grade_overview", "")).strip()
    kpi_items = split_kpi_items(str(row.get("kpi", "")))

    return {
        "title": "配属・職務定義書",
        "department": department,
        "position": f"{grade}（{grade_name}）",
        "job_overview": (
            f"{role_summary}"
            if role_summary
            else f"担当領域の業務を遂行します。"
        ),
        "grade_definition": (
            f"本職位は、{grade_overview}"
            if grade_overview and not grade_overview.startswith("本職位は")
            else grade_overview
        ),
        "kpi_items": kpi_items,
        "note": "本定義は、業務内容・組織状況に応じて見直される場合があります。",
    }


def make_handout_text(row: pd.Series) -> str:
    data = make_handout_dict(row)
    kpi_text = "\n".join([f"・{item}" for item in data["kpi_items"]]) if data["kpi_items"] else "・別途設定"
    return (
        f"【{data['title']}】\n\n"
        f"■ 所属\n{data['department']}\n\n"
        f"■ 職位\n{data['position']}\n\n"
        f"■ 職務概要\n{data['job_overview']}\n\n"
        f"■ グレード定義\n{data['grade_definition']}\n\n"
        f"■ 評価指標（KPI）\n{kpi_text}\n\n"
        f"■ 備考\n{data['note']}"
    )


def dataframe_for_handout_export(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["grade_sort"] = out["grade"].map({g: i for i, g in enumerate(GRADE_ORDER)})
    return out.sort_values(["department", "grade_sort"]).drop(columns=["grade_sort"])


def create_handout_excel_bytes(export_df: pd.DataFrame) -> bytes:
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
        from openpyxl.utils import get_column_letter
        from openpyxl.worksheet.pagebreak import Break
    except Exception as e:
        raise RuntimeError(
            "Excel出力には openpyxl が必要です。requirements.txt に openpyxl を追加してください。"
        ) from e

    wb = Workbook()
    default_ws = wb.active
    wb.remove(default_ws)

    thin_side = Side(style="thin", color="C7D2E0")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    label_fill = PatternFill("solid", fgColor="EAF1FB")
    title_font = Font(name="Arial", size=14, bold=True)
    label_font = Font(name="Arial", size=10.5, bold=True)
    body_font = Font(name="Arial", size=10.5)
    center_align = Alignment(horizontal="center", vertical="center")
    top_wrap = Alignment(horizontal="left", vertical="top", wrap_text=True)

    records = export_df.to_dict(orient="records")
    for idx, record in enumerate(records, start=1):
        row = pd.Series(record)
        data = make_handout_dict(row)

        raw_sheet_name = f"{data['department']}_{row.get('grade', '')}"
        invalid_chars = ['\\', '/', '*', '?', ':', '[', ']']
        for ch in invalid_chars:
            raw_sheet_name = raw_sheet_name.replace(ch, "_")
        sheet_name = raw_sheet_name[:31] or f"Sheet{idx}"

        counter = 1
        base_sheet_name = sheet_name
        while sheet_name in wb.sheetnames:
            suffix = f"_{counter}"
            sheet_name = f"{base_sheet_name[:31-len(suffix)]}{suffix}"
            counter += 1

        ws = wb.create_sheet(title=sheet_name)
        ws.sheet_view.showGridLines = False
        ws.freeze_panes = "A1"
        ws.page_setup.orientation = "portrait"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        ws.page_margins.left = 0.35
        ws.page_margins.right = 0.35
        ws.page_margins.top = 0.45
        ws.page_margins.bottom = 0.45
        ws.print_options.horizontalCentered = False
        ws.sheet_properties.pageSetUpPr.fitToPage = True

        ws.column_dimensions["A"].width = 18
        ws.column_dimensions["B"].width = 55

        ws.merge_cells("A1:B1")
        ws["A1"] = data["title"]
        ws["A1"].font = title_font
        ws["A1"].alignment = center_align
        ws.row_dimensions[1].height = 24

        rows = [
            ("所属", data["department"]),
            ("職位", data["position"]),
            ("職務概要", data["job_overview"]),
            ("グレード定義", data["grade_definition"]),
            ("評価指標（KPI）", "\n".join([f"・{item}" for item in data["kpi_items"]]) if data["kpi_items"] else "・別途設定"),
            ("備考", data["note"]),
            ("署名", "____________________"),
            ("受領日", "____________________"),
        ]

        start_row = 3
        for offset, (label, value) in enumerate(rows):
            r = start_row + offset
            ws[f"A{r}"] = label
            ws[f"B{r}"] = value
            ws[f"A{r}"].font = label_font
            ws[f"B{r}"].font = body_font
            ws[f"A{r}"].fill = label_fill
            ws[f"A{r}"].border = border
            ws[f"B{r}"].border = border
            ws[f"A{r}"].alignment = center_align
            ws[f"B{r}"].alignment = top_wrap

        for row_no, height in {
            3: 22,
            4: 22,
            5: 70,
            6: 84,
            7: 52,
            8: 30,
            9: 22,
            10: 22,
        }.items():
            ws.row_dimensions[row_no].height = height

        ws["B12"] = f"作成日: {datetime.now().strftime('%Y-%m-%d')}"
        ws["B12"].font = Font(name="Arial", size=9, italic=True)
        ws["B12"].alignment = Alignment(horizontal="right")

        ws.print_area = "A1:B12"

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def create_handout_docx_bytes(export_df: pd.DataFrame) -> bytes:
    try:
        from docx import Document
        from docx.enum.section import WD_SECTION
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        from docx.shared import Pt
    except Exception as e:
        raise RuntimeError(
            "Word出力には python-docx が必要です。requirements.txt に python-docx を追加してください。"
        ) from e

    def set_cell_shading(cell, fill: str):
        tc_pr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:fill"), fill)
        tc_pr.append(shd)

    def set_cell_margins(cell, top=90, start=100, bottom=90, end=100):
        tc = cell._tc
        tc_pr = tc.get_or_add_tcPr()
        tc_mar = tc_pr.first_child_found_in("w:tcMar")
        if tc_mar is None:
            tc_mar = OxmlElement("w:tcMar")
            tc_pr.append(tc_mar)
        for margin_name, value in [("top", top), ("start", start), ("bottom", bottom), ("end", end)]:
            node = tc_mar.find(qn(f"w:{margin_name}"))
            if node is None:
                node = OxmlElement(f"w:{margin_name}")
                tc_mar.append(node)
            node.set(qn("w:w"), str(value))
            node.set(qn("w:type"), "dxa")

    doc = Document()
    section = doc.sections[0]
    section.top_margin = Pt(42)
    section.bottom_margin = Pt(42)
    section.left_margin = Pt(46)
    section.right_margin = Pt(46)

    styles = doc.styles
    styles["Normal"].font.name = "Arial"
    styles["Normal"].font.size = Pt(10.5)
    styles["Normal"]._element.rPr.rFonts.set(qn("w:eastAsia"), "MS Gothic")

    records = export_df.to_dict(orient="records")
    for idx, record in enumerate(records):
        row = pd.Series(record)
        data = make_handout_dict(row)

        title = doc.add_paragraph()
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = title.add_run(data["title"])
        run.bold = True
        run.font.size = Pt(15)

        meta = doc.add_paragraph()
        meta.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        meta_run = meta.add_run(datetime.now().strftime("作成日: %Y-%m-%d"))
        meta_run.font.size = Pt(9)

        table = doc.add_table(rows=0, cols=2)
        table.style = "Table Grid"
        table.autofit = False
        table.columns[0].width = Pt(110)
        table.columns[1].width = Pt(360)

        rows = [
            ("所属", data["department"]),
            ("職位", data["position"]),
            ("職務概要", data["job_overview"]),
            ("グレード定義", data["grade_definition"]),
            ("評価指標（KPI）", "\n".join([f"・{item}" for item in data["kpi_items"]]) if data["kpi_items"] else "・別途設定"),
            ("備考", data["note"]),
        ]

        for label, value in rows:
            cells = table.add_row().cells
            cells[0].text = label
            cells[1].text = value
            set_cell_shading(cells[0], "EAF1FB")
            for c in cells:
                set_cell_margins(c)
                for p in c.paragraphs:
                    for r in p.runs:
                        r.font.name = "Arial"
                        r._element.rPr.rFonts.set(qn("w:eastAsia"), "MS Gothic")
                        r.font.size = Pt(10.5)

        doc.add_paragraph("")
        sig = doc.add_paragraph()
        sig.add_run("署名: ____________________    ")
        sig.add_run("受領日: ____________________")

        if idx != len(records) - 1:
            doc.add_section(WD_SECTION.NEW_PAGE)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def create_handout_pdf_bytes(export_df: pd.DataFrame) -> bytes:
    try:
        from reportlab.lib import colors
        from reportlab.lib.enums import TA_CENTER, TA_LEFT
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
        from reportlab.lib.units import mm
        from reportlab.pdfbase.cidfonts import UnicodeCIDFont
        from reportlab.pdfbase.pdfmetrics import registerFont
        from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    except Exception as e:
        raise RuntimeError(
            "PDF出力には reportlab が必要です。requirements.txt に reportlab を追加してください。"
        ) from e

    registerFont(UnicodeCIDFont("HeiseiKakuGo-W5"))
    registerFont(UnicodeCIDFont("HeiseiMin-W3"))

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=16 * mm,
        rightMargin=16 * mm,
        topMargin=15 * mm,
        bottomMargin=15 * mm,
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "TitleJP",
        parent=styles["Title"],
        fontName="HeiseiKakuGo-W5",
        fontSize=14,
        leading=18,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#1F2937"),
        spaceAfter=8,
    )
    body_style = ParagraphStyle(
        "BodyJP",
        parent=styles["BodyText"],
        fontName="HeiseiMin-W3",
        fontSize=10,
        leading=15,
        alignment=TA_LEFT,
        textColor=colors.HexColor("#111827"),
    )
    label_style = ParagraphStyle(
        "LabelJP",
        parent=body_style,
        fontName="HeiseiKakuGo-W5",
    )

    story = []
    records = export_df.to_dict(orient="records")
    for idx, record in enumerate(records):
        row = pd.Series(record)
        data = make_handout_dict(row)
        story.append(Paragraph(data["title"], title_style))
        story.append(Spacer(1, 3 * mm))

        rows = [
            [Paragraph("所属", label_style), Paragraph(data["department"].replace("\n", "<br/>"), body_style)],
            [Paragraph("職位", label_style), Paragraph(data["position"], body_style)],
            [Paragraph("職務概要", label_style), Paragraph(data["job_overview"].replace("\n", "<br/>"), body_style)],
            [Paragraph("グレード定義", label_style), Paragraph(data["grade_definition"].replace("\n", "<br/>"), body_style)],
            [
                Paragraph("評価指標（KPI）", label_style),
                Paragraph("<br/>".join([f"・{item}" for item in data["kpi_items"]]) if data["kpi_items"] else "・別途設定", body_style),
            ],
            [Paragraph("備考", label_style), Paragraph(data["note"], body_style)],
        ]

        table = Table(rows, colWidths=[34 * mm, 140 * mm], repeatRows=0)
        table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#EAF1FB")),
                    ("BOX", (0, 0), (-1, -1), 0.75, colors.HexColor("#C7D2E0")),
                    ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#C7D2E0")),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 8),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                    ("TOPPADDING", (0, 0), (-1, -1), 7),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
                ]
            )
        )
        story.append(table)
        story.append(Spacer(1, 5 * mm))
        story.append(Paragraph("署名: ____________________　　受領日: ____________________", body_style))
        if idx != len(records) - 1:
            story.append(PageBreak())

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


# =========================
# SESSION
# =========================
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

if "df" not in st.session_state or "grade_master" not in st.session_state:
    latest_df, latest_grade_master = load_latest_snapshot()

    if latest_df is not None and not latest_df.empty:
        st.session_state.df = latest_df
    else:
        st.session_state.df = generate_default()

    if latest_grade_master is not None and not latest_grade_master.empty:
        st.session_state.grade_master = latest_grade_master
    else:
        st.session_state.grade_master = get_default_grade_master()

st.session_state.df = ensure_main_columns(st.session_state.df)
st.session_state.grade_master = ensure_grade_master(st.session_state.grade_master)

if "last_change_log" not in st.session_state:
    st.session_state.last_change_log = pd.DataFrame()
if "last_grade_change_log" not in st.session_state:
    st.session_state.last_grade_change_log = pd.DataFrame()

master_df = ensure_main_columns(st.session_state.df.copy())
grade_master_df = ensure_grade_master(st.session_state.grade_master.copy())
display_master_df = attach_grade_overview(master_df, grade_master_df)
summary_df = department_summary(master_df)

m1, m2, m3, m4 = st.columns(4)
m1.metric("部門数", len(DEPARTMENTS))
m2.metric("グレード数", len(GRADE_ORDER))
m3.metric("管理行数", len(master_df))
m4.metric("グレード定義数", len(grade_master_df))

st.markdown("## フィルター")
c1, c2, c3 = st.columns([1.2, 1, 1.2])
with c1:
    dept_filter = st.multiselect("部門", options=DEPARTMENTS, default=[])
with c2:
    grade_filter = st.multiselect("グレード", options=GRADE_ORDER, default=[])
with c3:
    keyword = st.text_input("キーワード検索", placeholder="職務概要・グレード定義・KPI など")

filtered_df = display_master_df.copy()
if dept_filter:
    filtered_df = filtered_df[filtered_df["department"].isin(dept_filter)]
if grade_filter:
    filtered_df = filtered_df[filtered_df["grade"].isin(grade_filter)]
if keyword:
    kw = keyword.lower()
    filtered_df = filtered_df[
        filtered_df.apply(lambda row: any(kw in str(v).lower() for v in row.values), axis=1)
    ]

filtered_df["grade_sort"] = filtered_df["grade"].map({g: i for i, g in enumerate(GRADE_ORDER)})
filtered_df = filtered_df.sort_values(["department", "grade_sort"]).drop(columns=["grade_sort"])
st.caption(f"表示件数: {len(filtered_df)} / 全 {len(master_df)} 件")

tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs(
    [
        "一覧",
        "編集",
        "配布用出力",
        "CSV / Excel",
        "履歴",
        "グレード定義",
        "設定",
    ]
)

with tab1:
    st.subheader("一覧ビュー")
    view_mode = st.radio("表示形式", ["部門ごとの展開表示", "全件テーブル"], horizontal=True)

    if view_mode == "部門ごとの展開表示":
        for dept in DEPARTMENTS:
            dept_df = filtered_df[filtered_df["department"] == dept].copy()
            if dept_df.empty:
                continue
            with st.expander(f"{dept}", expanded=False):
                st.markdown(
                    f"<div class='small-note'>{DEPARTMENT_DESCRIPTIONS.get(dept, '')}</div>",
                    unsafe_allow_html=True,
                )
                for _, row in dept_df.iterrows():
                    st.markdown(f"### {row['grade']} / {row['grade_name']}")
                    st.markdown(
                        f"<div class='handout-box'>{make_handout_text(row)}</div>",
                        unsafe_allow_html=True,
                    )
                    st.markdown("---")
    else:
        table_df = filtered_df.copy()
        table_df["配布用プレビュー"] = table_df.apply(make_handout_text, axis=1)
        st.dataframe(table_df, use_container_width=True, hide_index=True)

    st.markdown("### 部門別サマリー")
    st.dataframe(summary_df, use_container_width=True, hide_index=True)

with tab2:
    st.subheader("直接編集")
    st.write("`職務概要` は配布用の書き方を前提にした 2〜3文で入力してください。`グレード定義` は別タブの内容を自動参照します。")

    edited_df = st.data_editor(
        filtered_df,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        column_config={
            "department": st.column_config.TextColumn("所属部門", disabled=True),
            "grade": st.column_config.TextColumn("グレード", disabled=True),
            "grade_name": st.column_config.TextColumn("グレード名"),
            "role_summary": st.column_config.TextColumn("職務概要（2〜3文）", width="large"),
            "grade_overview": st.column_config.TextColumn("グレード定義", disabled=True, width="large"),
            "kpi": st.column_config.TextColumn("評価指標（KPI）", width="medium"),
        },
    )

    memo = st.text_input("保存メモ", placeholder="例: 総務課と財務課の配布文面を調整")
    if st.button("編集内容をSupabaseへ保存", type="primary"):
        new_master = update_master_from_editor(
            st.session_state.df,
            edited_df.drop(columns=["grade_overview"], errors="ignore"),
        )
        st.session_state.df = ensure_main_columns(new_master)
        save_snapshot(st.session_state.df, st.session_state.grade_master, memo or "manual edit save")
        st.success("保存完了")

with tab3:
    st.subheader("配布用出力")
    st.write("新人配属時にそのまま渡せる形式で、Word / PDF を出力します。")

    export_mode = st.radio("出力対象", ["1件だけ出力", "フィルター結果をまとめて出力"], horizontal=True)

    export_source_df = dataframe_for_handout_export(filtered_df)
    selected_export_df = export_source_df.copy()

    if export_mode == "1件だけ出力":
        options_df = dataframe_for_handout_export(display_master_df)
        option_map = {
            f"{r['department']} | {r['grade']} / {r['grade_name']}": idx
            for idx, r in options_df.reset_index(drop=True).iterrows()
        }
        selected_label = st.selectbox("出力する職位", list(option_map.keys()))
        selected_export_df = options_df.iloc[[option_map[selected_label]]].copy()
    else:
        st.caption("現在のフィルター結果を対象に一括出力します。")
        if selected_export_df.empty:
            st.warning("フィルター結果が0件です。部門・グレード条件を見直してください。")

    if not selected_export_df.empty:
        preview_row = selected_export_df.iloc[0]
        st.markdown("### プレビュー")
        st.markdown(
            f"<div class='handout-box'>{make_handout_text(preview_row)}</div>",
            unsafe_allow_html=True,
        )
        if len(selected_export_df) > 1:
            st.caption(f"このプレビューは {len(selected_export_df)} 件のうち先頭1件です。")

        base_name = (
            f"job_assignment_{preview_row['department']}_{preview_row['grade']}"
            if len(selected_export_df) == 1
            else f"job_assignment_bundle_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        )

        col_docx, col_pdf, col_xlsx = st.columns(3)
        with col_docx:
            try:
                docx_bytes = create_handout_docx_bytes(selected_export_df)
                st.download_button(
                    "Wordを出力",
                    data=docx_bytes,
                    file_name=f"{base_name}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
            except Exception as e:
                st.error(f"Word出力エラー: {e}")
        with col_pdf:
            try:
                pdf_bytes = create_handout_pdf_bytes(selected_export_df)
                st.download_button(
                    "PDFを出力",
                    data=pdf_bytes,
                    file_name=f"{base_name}.pdf",
                    mime="application/pdf",
                )
            except Exception as e:
                st.error(f"PDF出力エラー: {e}")
        with col_xlsx:
            try:
                excel_bytes = create_handout_excel_bytes(selected_export_df)
                st.download_button(
                    "配布用Excelを出力",
                    data=excel_bytes,
                    file_name=f"{base_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as e:
                st.error(f"Excel出力エラー: {e}")

with tab4:
    st.subheader("CSV / Excel 入出力")

    dl1, dl2, dl3, dl4, dl5 = st.columns(5)
    with dl1:
        st.download_button(
            "現在データをCSV出力",
            data=dataframe_to_csv_bytes(st.session_state.df),
            file_name="job_matrix_v4_export.csv",
            mime="text/csv",
        )
    with dl2:
        st.download_button(
            "現在データをExcel出力",
            data=dataframe_to_excel_bytes(st.session_state.df, st.session_state.grade_master),
            file_name="job_matrix_v4_export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with dl3:
        st.download_button(
            "グレード定義CSV出力",
            data=grade_master_to_csv_bytes(st.session_state.grade_master),
            file_name="grade_master_v4_export.csv",
            mime="text/csv",
        )
    with dl4:
        st.download_button(
            "業務マスタ空テンプレートCSV",
            data=minimal_template_csv_bytes(),
            file_name="job_matrix_v4_template.csv",
            mime="text/csv",
        )
    with dl5:
        st.download_button(
            "グレード定義テンプレートCSV",
            data=minimal_grade_template_csv_bytes(),
            file_name="grade_master_template.csv",
            mime="text/csv",
        )

    st.markdown("---")
    st.markdown("### 業務マスタCSV部分更新")
    st.write("`department + grade` が一致する行だけ対象です。")
    st.write("CSVで値が入っているセルだけ更新し、空欄は既存値を保持します。")
    st.write("更新対象列: `grade_name`, `role_summary`, `kpi`")

    uploaded_csv = st.file_uploader("業務マスタ更新用CSVをアップロード", type=["csv"], key="job_csv")
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

            csv_memo = st.text_input("業務マスタ更新メモ", placeholder="例: 部門長レビュー反映", key="job_csv_memo")
            if st.button("業務マスタをCSV内容で部分更新する", type="primary"):
                new_df, change_log = merge_partial_update(st.session_state.df, import_df)
                st.session_state.df = ensure_main_columns(new_df)
                st.session_state.last_change_log = change_log
                save_snapshot(
                    st.session_state.df,
                    st.session_state.grade_master,
                    csv_memo or f"job csv partial update ({len(change_log)} changes)",
                )
                st.success(f"更新完了: {len(change_log)} 件")
        except Exception as e:
            st.error(f"CSV取込エラー: {e}")

    st.markdown("---")
    st.markdown("### グレード定義CSV部分更新")
    st.write("`grade` が一致する行だけ対象です。")
    st.write("更新対象列: `grade_name`, `grade_overview`")

    uploaded_grade_csv = st.file_uploader("グレード定義更新用CSVをアップロード", type=["csv"], key="grade_csv")
    if uploaded_grade_csv is not None:
        try:
            import_grade_df = pd.read_csv(uploaded_grade_csv)
            st.markdown("#### グレード定義アップロード内容")
            st.dataframe(import_grade_df, use_container_width=True, hide_index=True)

            grade_csv_memo = st.text_input(
                "グレード定義更新メモ",
                placeholder="例: PDF文言に合わせて定義修正",
                key="grade_csv_memo",
            )
            if st.button("グレード定義をCSV内容で部分更新する", type="primary"):
                new_grade_df, grade_change_log = merge_grade_master_update(
                    st.session_state.grade_master,
                    import_grade_df,
                )
                st.session_state.grade_master = ensure_grade_master(new_grade_df)
                st.session_state.last_grade_change_log = grade_change_log
                save_snapshot(
                    st.session_state.df,
                    st.session_state.grade_master,
                    grade_csv_memo or f"grade csv partial update ({len(grade_change_log)} changes)",
                )
                st.success(f"更新完了: {len(grade_change_log)} 件")
        except Exception as e:
            st.error(f"グレード定義CSV取込エラー: {e}")

    st.markdown("### 最終更新ログ（業務マスタ）")
    if not st.session_state.last_change_log.empty:
        st.dataframe(st.session_state.last_change_log, use_container_width=True, hide_index=True)
    else:
        st.info("まだ業務マスタCSV更新ログはありません。")

    st.markdown("### 最終更新ログ（グレード定義）")
    if not st.session_state.last_grade_change_log.empty:
        st.dataframe(st.session_state.last_grade_change_log, use_container_width=True, hide_index=True)
    else:
        st.info("まだグレード定義CSV更新ログはありません。")

with tab5:
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
            hist_df = ensure_main_columns(pd.DataFrame(item.get("data", [])))
            hist_grade_master = ensure_grade_master(pd.DataFrame(item.get("grade_master", [])))
            hist_display_df = attach_grade_overview(hist_df, hist_grade_master)

            with st.expander(f"{created_at} | {memo}", expanded=False):
                if hist_df.empty:
                    st.warning("履歴データが空です。")
                else:
                    hist_display_df["grade_sort"] = hist_display_df["grade"].map(
                        {g: j for j, g in enumerate(GRADE_ORDER)}
                    )
                    hist_display_df = hist_display_df.sort_values(["department", "grade_sort"]).drop(
                        columns=["grade_sort"]
                    )
                    st.markdown("#### 業務マスタ")
                    st.dataframe(hist_display_df, use_container_width=True, hide_index=True)
                    st.markdown("#### グレード定義")
                    st.dataframe(hist_grade_master, use_container_width=True, hide_index=True)
                    if st.button(f"この履歴を復元 #{i}", key=f"restore_{i}"):
                        st.session_state.df = ensure_main_columns(hist_df)
                        st.session_state.grade_master = ensure_grade_master(hist_grade_master)
                        st.success("復元しました。必要なら編集タブ等で再保存してください。")

with tab6:
    st.subheader("グレード定義")
    st.write("ここで各グレードの説明欄を管理します。一覧・編集・配布用出力の `グレード定義` はこの内容を引用表示します。")

    grade_editor = st.data_editor(
        grade_master_df,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        column_config={
            "grade": st.column_config.TextColumn("グレード", disabled=True),
            "grade_name": st.column_config.TextColumn("グレード名"),
            "grade_overview": st.column_config.TextColumn("グレード定義", width="large"),
        },
    )

    st.markdown("### プレビュー")
    for _, row in ensure_grade_master(grade_editor).iterrows():
        st.markdown(f"#### {row['grade']} / {row['grade_name']}")
        st.markdown(f"<div class='grade-box'>{row['grade_overview']}</div>", unsafe_allow_html=True)

    grade_memo = st.text_input("グレード定義保存メモ", placeholder="例: PDF文言に寄せて修正", key="grade_tab_memo")
    if st.button("グレード定義を保存", type="primary"):
        st.session_state.grade_master = ensure_grade_master(grade_editor)
        save_snapshot(st.session_state.df, st.session_state.grade_master, grade_memo or "grade master update")
        st.success("グレード定義を保存しました。")

with tab7:
    st.subheader("設定 / 初期化")
    st.write("必要に応じて初期データへ戻せます。")
    if st.button("初期データに戻す"):
        st.session_state.df = generate_default()
        st.session_state.grade_master = get_default_grade_master()
        st.session_state.last_change_log = pd.DataFrame()
        st.session_state.last_grade_change_log = pd.DataFrame()
        st.success("初期データへ戻しました。")

    st.markdown("---")
    st.markdown("### Streamlit Secrets 例")
    st.code(
        'SUPABASE_URL = "https://xxxx.supabase.co"\nSUPABASE_KEY = "anon public key"',
        language="toml",
    )

    st.markdown("### requirements.txt 追記推奨")
    st.code(
        "python-docx\nreportlab\nopenpyxl\npandas\nrequests",
        language="text",
    )

    st.markdown("### 業務マスタCSV例")
    st.code(
        "department,grade,role_summary,kpi\n"
        '人事課,G6,"採用・勤怠・人事データ入力などの定型実務を担当する。上位者の指示に沿って正確に処理し、必要な情報を遅れなく整備する。社内の基礎的な人事運営を支える役割を担う。","業務達成率、処理件数、正確性、期限遵守率"',
        language="csv",
    )

    st.markdown("### グレード定義CSV例")
    st.code(
        "grade,grade_name,grade_overview\n"
        'G6,平社員,"上位者の指示・方針のもとで担当業務を確実に遂行し、組織の一員として基本的な役割と責任を果たすグレード。"',
        language="csv",
    )

    st.markdown("### 補足")
    st.markdown(
        """
- 一覧表示は、新人配布用の文面フォーマットでプレビューする設計に変更しています。  
- `職務概要`  `...` 形式でWord/PDFへ自動整形されます。  
- `グレード定義` は別タブで管理し、一覧・編集・配布用出力へ引用表示します。  
- 保存は **Supabase REST API** を使っているため、`supabase-py` の proxy 依存問題を回避できます。  
- テーブルはスナップショット保存方式です。履歴をそのまま復元できます。  
- RLS を ON にする場合は、`SELECT` と `INSERT` のポリシーを設定してください。
"""
    )
