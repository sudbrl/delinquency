"""
Month Condition Analyzer
========================
Industry-grade Streamlit application for analyzing loan-account data
against a user-defined month condition.

Calculation logic, parsing logic, aggregation logic, and Excel output
structure are preserved exactly from the original implementation; only
the surrounding engineering (structure, typing, logging, validation,
error handling, styling, performance, UX) has been upgraded.
"""

from __future__ import annotations

import logging
import re
import sys
from dataclasses import dataclass, field
from io import BytesIO
from typing import Any, Dict, List, Optional, Sequence, Tuple

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import streamlit as st

# =========================================================
# LOGGING
# =========================================================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
    stream=sys.stdout,
)
logger = logging.getLogger("month_analyzer")


# =========================================================
# PAGE CONFIG
# =========================================================
st.set_page_config(
    page_title="Month Condition Analyzer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


# =========================================================
# CONSTANTS
# =========================================================
class AppConstants:
    APP_TITLE: str = "📊 Month Condition Analyzer"
    APP_SUBTITLE: str = (
        "**Industry-level data analytics** — Upload an Excel file, choose month scope, "
        "define a condition, and get matching accounts with Branch & Ac Type summaries "
        "plus visual insights."
    )

    MONTH_BASE_NAMES: Tuple[str, ...] = (
        "shrawan", "bhadra", "ashoj", "kartik", "mangshir", "poush",
        "magh", "falgun", "chaitra", "baisakh", "jestha", "asadh",
    )

    REQUIRED_COLUMNS: Tuple[str, ...] = (
        "Balance", "Branch Name", "Ac Type Desc", "Main Code",
    )

    MISSING_VALUE_TOKENS: frozenset = frozenset({"#N/A", "N/A", "#N/A!", "NA"})

    DEFAULT_CRITERIA: str = "=0"
    MONTH_SCOPE_OPTIONS: Tuple[str, ...] = (
        "All Months", "Last 12 Months", "Month Range",
    )

    PRIMARY_COLOR: str = "#1F4E79"
    ACCENT_COLOR: str = "#D9E1F2"
    BORDER_COLOR: str = "#C8C8C8"

    XLSX_MIME: str = (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# =========================================================
# DATA MODELS
# =========================================================
@dataclass(frozen=True)
class ParsedCriteria:
    is_valid: bool
    kind: Optional[str]
    op1: Optional[str]
    val1: Optional[float]
    op2: Optional[str]
    val2: Optional[float]

    @property
    def treat_missing_as_ignorable(self) -> bool:
        return (
            self.kind == "SIMPLE"
            and self.op1 == "="
            and self.val1 == 0.0
        )


@dataclass
class AnalysisContext:
    month_scope: str
    from_month: Optional[str] = None
    to_month: Optional[str] = None
    criteria_raw: str = ""
    criteria: Optional[ParsedCriteria] = None
    month_cols: List[str] = field(default_factory=list)
    all_month_cols: List[str] = field(default_factory=list)
    drop_month_cols: List[str] = field(default_factory=list)


@dataclass
class AnalysisResult:
    df_zero: pd.DataFrame
    df_branch: pd.DataFrame
    df_ac_type: pd.DataFrame
    total_loan_account_count: int = 0
    total_loan_balance: float = 0.0
    branch_total_count: Dict[str, int] = field(default_factory=dict)
    ac_type_total_count: Dict[str, int] = field(default_factory=dict)


# =========================================================
# I/O HELPERS
# =========================================================
@st.cache_data(show_spinner=False)
def load_data(file_bytes: bytes) -> pd.DataFrame:
    logger.info("Loading Excel file (%d bytes).", len(file_bytes))
    return pd.read_excel(BytesIO(file_bytes), engine="openpyxl", dtype=str)


def find_header_column(df: pd.DataFrame, header_name: str) -> Optional[str]:
    target = header_name.strip().lower()
    for col in df.columns:
        if str(col).strip().lower() == target:
            return col
    return None


# =========================================================
# MONTH DETECTION
# =========================================================
def is_month_header(header_value: Any) -> bool:
    h = str(header_value).strip().lower().replace(chr(160), " ")
    if not h:
        return False
    return any(m in h for m in AppConstants.MONTH_BASE_NAMES)


def detect_month_columns(df: pd.DataFrame) -> List[str]:
    return [col for col in df.columns if is_month_header(col)]


def is_missing_month_value(v: Any) -> bool:
    if pd.isna(v):
        return True
    if isinstance(v, str):
        s = v.strip().upper().replace(chr(160), " ")
        if s in AppConstants.MISSING_VALUE_TOKENS:
            return True
    return False


# =========================================================
# CRITERIA PARSING  (logic preserved exactly)
# =========================================================
def try_parse_criteria(criteria_str: str) -> ParsedCriteria:
    s = str(criteria_str).strip().replace(chr(160), "").replace(" ", "")
    if not s:
        return ParsedCriteria(False, None, None, None, None, None)

    pattern = r'(>=|<=|<>|!=|>|<|=)?([+-]?[\d,.]+)'
    matches = re.findall(pattern, s)
    conditions: List[Tuple[str, float]] = []

    for op, num_str in matches:
        if op == "":
            op = "="
        if op == "!=":
            op = "<>"
        num_str = num_str.replace(',', '')
        try:
            val = float(num_str)
            conditions.append((op, val))
        except ValueError:
            pass

    if not conditions:
        try:
            val = float(re.match(r'^([+-]?[\d,.]+)', s).group(1).replace(',', ''))
            return ParsedCriteria(True, "SIMPLE", "=", val, "=", 0.0)
        except (AttributeError, ValueError):
            return ParsedCriteria(False, None, None, None, None, None)

    if len(conditions) >= 2:
        return ParsedCriteria(
            True, "DUAL",
            conditions[0][0], conditions[0][1],
            conditions[1][0], conditions[1][1],
        )

    return ParsedCriteria(
        True, "SIMPLE",
        conditions[0][0], conditions[0][1],
        "=", 0.0,
    )


def compare_value(v: float, op: str, val: float) -> bool:
    if op == "=":
        return v == val
    if op == ">":
        return v > val
    if op == "<":
        return v < val
    if op == ">=":
        return v >= val
    if op == "<=":
        return v <= val
    if op == "<>":
        return v != val
    return False


def value_meets_criteria(v: float, criteria: ParsedCriteria) -> bool:
    if criteria.kind == "DUAL":
        return compare_value(v, criteria.op1, criteria.val1) and \
               compare_value(v, criteria.op2, criteria.val2)
    return compare_value(v, criteria.op1, criteria.val1)


# =========================================================
# BALANCE / NUMERIC HELPERS
# =========================================================
def parse_balance(b_val: Any) -> Tuple[float, bool]:
    if pd.isna(b_val):
        return 0.0, False
    if not isinstance(b_val, str):
        try:
            return float(b_val), True
        except (ValueError, TypeError):
            return 0.0, False
    try:
        return float(b_val.replace(',', '')), True
    except (ValueError, AttributeError):
        return 0.0, False


def is_numeric_string(v: Any) -> bool:
    if isinstance(v, (int, float)):
        return True
    if isinstance(v, str):
        return str(v).replace('.', '').replace(',', '').replace('-', '').isdigit()
    return False


# =========================================================
# CORE ANALYSIS  (calculation logic preserved exactly)
# =========================================================
def run_analysis(
    df: pd.DataFrame,
    ctx: AnalysisContext,
    col_balance: str,
    col_branch: str,
    col_ac_type: str,
    col_main_code: str,
) -> AnalysisResult:
    logger.info(
        "Starting analysis | scope=%s | months=%d | criteria='%s'",
        ctx.month_scope, len(ctx.month_cols), ctx.criteria_raw,
    )

    treat_missing_as_ignorable = ctx.criteria.treat_missing_as_ignorable

    zero_accounts: List[pd.Series] = []
    branch_zero_sum: Dict[str, List[Any]] = {}
    ac_type_zero_sum: Dict[str, List[Any]] = {}
    branch_loan_sum: Dict[str, float] = {}
    ac_type_loan_sum: Dict[str, float] = {}
    branch_total_count: Dict[str, int] = {}
    ac_type_total_count: Dict[str, int] = {}
    total_loan_account_count = 0

    month_cols = ctx.month_cols

    for idx, row in df.iterrows():
        if idx == 0 and str(row[col_main_code]).strip().lower() == "main code":
            continue

        branch_name = str(row[col_branch]).strip() if pd.notna(row[col_branch]) else ""
        ac_type = str(row[col_ac_type]).strip() if pd.notna(row[col_ac_type]) else ""
        main_code = str(row[col_main_code]).strip() if pd.notna(row[col_main_code]) else ""

        balance_value, balance_is_valid = parse_balance(row[col_balance])

        if main_code != "":
            total_loan_account_count += 1

            branch_loan_sum.setdefault(branch_name, 0.0)
            if balance_is_valid:
                branch_loan_sum[branch_name] += balance_value

            ac_type_loan_sum.setdefault(ac_type, 0.0)
            if balance_is_valid:
                ac_type_loan_sum[ac_type] += balance_value

            branch_total_count[branch_name] = branch_total_count.get(branch_name, 0) + 1
            ac_type_total_count[ac_type] = ac_type_total_count.get(ac_type, 0) + 1

        na_count = sum(1 for mcol in month_cols if is_missing_month_value(row[mcol]))
        if na_count == len(month_cols):
            continue

        zero_all_months = True
        for mcol in month_cols:
            mv = row[mcol]

            if is_missing_month_value(mv):
                if not treat_missing_as_ignorable:
                    zero_all_months = False
                    break

            elif pd.notna(mv) and (
                isinstance(mv, (int, float)) or is_numeric_string(mv)
            ):
                try:
                    mv_num = float(mv)
                    if not value_meets_criteria(mv_num, ctx.criteria):
                        zero_all_months = False
                        break
                except (ValueError, TypeError):
                    zero_all_months = False
                    break
            else:
                zero_all_months = False
                break

        if zero_all_months:
            zero_accounts.append(row)

            if branch_name not in branch_zero_sum:
                branch_zero_sum[branch_name] = [
                    balance_value if balance_is_valid else 0.0, 1
                ]
            else:
                if balance_is_valid:
                    branch_zero_sum[branch_name][0] += balance_value
                branch_zero_sum[branch_name][1] += 1

            if ac_type not in ac_type_zero_sum:
                ac_type_zero_sum[ac_type] = [
                    balance_value if balance_is_valid else 0.0, 1
                ]
            else:
                if balance_is_valid:
                    ac_type_zero_sum[ac_type][0] += balance_value
                ac_type_zero_sum[ac_type][1] += 1

    df_zero = pd.DataFrame(zero_accounts) if zero_accounts else pd.DataFrame()

    if ctx.month_scope == "Month Range" and ctx.drop_month_cols and len(df_zero) > 0:
        df_zero = df_zero.drop(
            columns=[c for c in ctx.drop_month_cols if c in df_zero.columns],
            errors='ignore',
        )

    # Branch summary
    branch_data: List[Dict[str, Any]] = []
    all_branches = set(list(branch_zero_sum.keys()) + list(branch_loan_sum.keys()))
    for branch in sorted(all_branches):
        zero_bal, zero_count = branch_zero_sum.get(branch, [0.0, 0])
        total_bal = branch_loan_sum.get(branch, 0.0)
        branch_data.append({
            "Branch Name": branch,
            "Sum of Balance (Zero Accounts)": zero_bal,
            "Total Loan Balance (All Accounts)": total_bal,
            "No. of Zero Accounts": zero_count,
        })
    if branch_data:
        branch_data.append({
            "Branch Name": "Grand Total",
            "Sum of Balance (Zero Accounts)": sum(d["Sum of Balance (Zero Accounts)"] for d in branch_data),
            "Total Loan Balance (All Accounts)": sum(d["Total Loan Balance (All Accounts)"] for d in branch_data),
            "No. of Zero Accounts": sum(d["No. of Zero Accounts"] for d in branch_data),
        })
    df_branch = pd.DataFrame(branch_data)

    # Ac-type summary
    ac_type_data: List[Dict[str, Any]] = []
    all_ac_types = set(list(ac_type_zero_sum.keys()) + list(ac_type_loan_sum.keys()))
    for ac in sorted(all_ac_types):
        zero_bal, zero_count = ac_type_zero_sum.get(ac, [0.0, 0])
        total_bal = ac_type_loan_sum.get(ac, 0.0)
        ac_type_data.append({
            "Ac Type Desc": ac,
            "Sum of Balance (Zero Accounts)": zero_bal,
            "Total Loan Balance (All Accounts)": total_bal,
            "No. of Zero Accounts": zero_count,
        })
    if ac_type_data:
        ac_type_data.append({
            "Ac Type Desc": "Grand Total",
            "Sum of Balance (Zero Accounts)": sum(d["Sum of Balance (Zero Accounts)"] for d in ac_type_data),
            "Total Loan Balance (All Accounts)": sum(d["Total Loan Balance (All Accounts)"] for d in ac_type_data),
            "No. of Zero Accounts": sum(d["No. of Zero Accounts"] for d in ac_type_data),
        })
    df_ac_type = pd.DataFrame(ac_type_data)

    total_loan_balance = sum(branch_loan_sum.values())

    logger.info(
        "Analysis complete | matched=%d | total_loan_accounts=%d | total_balance=%.2f",
        len(df_zero), total_loan_account_count, total_loan_balance,
    )

    return AnalysisResult(
        df_zero=df_zero,
        df_branch=df_branch,
        df_ac_type=df_ac_type,
        total_loan_account_count=total_loan_account_count,
        total_loan_balance=total_loan_balance,
        branch_total_count=branch_total_count,
        ac_type_total_count=ac_type_total_count,
    )


# =========================================================
# EXCEL FORMATTING  (logic preserved exactly)
# =========================================================
def clean_for_excel(
    df: pd.DataFrame,
    text_cols: Optional[Sequence[str]] = None,
) -> pd.DataFrame:
    df = df.copy()
    text_cols = list(text_cols) if text_cols else []

    for col in df.columns:
        if col in text_cols:
            df[col] = df[col].fillna('').astype(str)
        elif pd.api.types.is_numeric_dtype(df[col]):
            df[col] = df[col].fillna(0).replace([np.inf, -np.inf], 0)
        else:
            df[col] = df[col].fillna('')
    return df


def format_excel_output(
    writer: pd.ExcelWriter,
    sheet_name: str,
    df: pd.DataFrame,
    is_summary: bool = False,
    text_cols: Optional[Sequence[str]] = None,
) -> None:
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    text_cols = list(text_cols) if text_cols else []

    header_fmt = workbook.add_format({
        'bold': True,
        'font_color': 'white',
        'bg_color': AppConstants.PRIMARY_COLOR,
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'border_color': AppConstants.BORDER_COLOR,
    })
    data_fmt = workbook.add_format({
        'border': 1,
        'border_color': AppConstants.BORDER_COLOR,
    })
    text_fmt = workbook.add_format({
        'border': 1,
        'border_color': AppConstants.BORDER_COLOR,
        'num_format': '@',
    })
    num_fmt = workbook.add_format({
        'num_format': '#,##0.00',
        'border': 1,
        'border_color': AppConstants.BORDER_COLOR,
    })
    count_fmt = workbook.add_format({
        'num_format': '#,##0',
        'border': 1,
        'border_color': AppConstants.BORDER_COLOR,
    })
    total_fmt = workbook.add_format({
        'bold': True,
        'bg_color': AppConstants.ACCENT_COLOR,
        'border': 1,
        'border_color': AppConstants.BORDER_COLOR,
    })
    total_num_fmt = workbook.add_format({
        'bold': True,
        'bg_color': AppConstants.ACCENT_COLOR,
        'num_format': '#,##0.00',
        'border': 1,
        'border_color': AppConstants.BORDER_COLOR,
    })
    total_count_fmt = workbook.add_format({
        'bold': True,
        'bg_color': AppConstants.ACCENT_COLOR,
        'num_format': '#,##0',
        'border': 1,
        'border_color': AppConstants.BORDER_COLOR,
    })
    total_text_fmt = workbook.add_format({
        'bold': True,
        'bg_color': AppConstants.ACCENT_COLOR,
        'border': 1,
        'border_color': AppConstants.BORDER_COLOR,
        'num_format': '@',
    })

    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_fmt)

    for row_num in range(len(df)):
        for col_num, value in enumerate(df.iloc[row_num].values):
            col_name = df.columns[col_num]
            is_text_col = col_name in text_cols
            is_total_row = is_summary and row_num == len(df) - 1

            if is_total_row:
                if col_num == 0:
                    fmt = total_text_fmt if is_text_col else total_fmt
                    worksheet.write(row_num + 1, col_num, value, fmt)
                elif col_num in (1, 2):
                    worksheet.write(row_num + 1, col_num, value, total_num_fmt)
                elif col_num == 3:
                    worksheet.write(row_num + 1, col_num, value, total_count_fmt)
                else:
                    fmt = total_text_fmt if is_text_col else total_fmt
                    worksheet.write(row_num + 1, col_num, value, fmt)
            else:
                if is_text_col:
                    worksheet.write(row_num + 1, col_num, value, text_fmt)
                elif col_num in (1, 2) and is_summary:
                    worksheet.write(row_num + 1, col_num, value, num_fmt)
                elif col_num == 3 and is_summary:
                    worksheet.write(row_num + 1, col_num, value, count_fmt)
                else:
                    worksheet.write(row_num + 1, col_num, value, data_fmt)

    for i, col in enumerate(df.columns):
        max_len = max(df[col].astype(str).map(len).max(), len(str(col))) + 2
        worksheet.set_column(i, i, min(max_len, 50))

    worksheet.freeze_panes(1, 0)
    worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)


def build_excel_output(
    result: AnalysisResult,
    text_cols: Sequence[str],
) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(
        output,
        engine='xlsxwriter',
        engine_kwargs={'options': {'nan_inf_to_errors': True}},
    ) as writer:

        if len(result.df_zero) > 0:
            df_z = clean_for_excel(result.df_zero, text_cols=text_cols)
            df_z.to_excel(writer, sheet_name='Zero Accounts', index=False)
            format_excel_output(writer, 'Zero Accounts', df_z,
                                is_summary=False, text_cols=text_cols)
        else:
            pd.DataFrame({"Message": ["No matching accounts found"]}).to_excel(
                writer, sheet_name='Zero Accounts', index=False
            )

        if len(result.df_branch) > 0:
            df_b = clean_for_excel(result.df_branch, text_cols=text_cols)
            df_b.to_excel(writer, sheet_name='Branch Summary', index=False)
            format_excel_output(writer, 'Branch Summary', df_b,
                                is_summary=True, text_cols=text_cols)
        else:
            pd.DataFrame({"Message": ["No branch data"]}).to_excel(
                writer, sheet_name='Branch Summary', index=False
            )

        if len(result.df_ac_type) > 0:
            df_a = clean_for_excel(result.df_ac_type, text_cols=text_cols)
            df_a.to_excel(writer, sheet_name='Ac Type Summary', index=False)
            format_excel_output(writer, 'Ac Type Summary', df_a,
                                is_summary=True, text_cols=text_cols)
        else:
            pd.DataFrame({"Message": ["No account type data"]}).to_excel(
                writer, sheet_name='Ac Type Summary', index=False
            )

    output.seek(0)
    return output.getvalue()


# =========================================================
# CHARTS
# =========================================================
def plot_bar_h(
    data: pd.DataFrame,
    x_col: str,
    y_col: str,
    title: str,
    color: str = AppConstants.PRIMARY_COLOR,
) -> plt.Figure:
    n_bars = len(data)
    fig_height = max(5, min(12, n_bars * 0.65))
    fig, ax = plt.subplots(figsize=(11, fig_height))

    bars = ax.barh(
        data[y_col],
        data[x_col],
        color=color,
        edgecolor='white',
        linewidth=0.5,
    )

    ax.set_xlabel(x_col, fontsize=11)
    ax.set_ylabel(y_col, fontsize=11)
    ax.set_title(title, fontsize=13, fontweight='bold', pad=15)
    ax.invert_yaxis()
    ax.grid(axis='x', alpha=0.3, linestyle='--')

    for bar in bars:
        width = bar.get_width()
        ax.text(
            width,
            bar.get_y() + bar.get_height() / 2,
            f' {width:,.1f}%',
            va='center',
            ha='left',
            fontsize=9,
            color='#333333',
        )

    plt.tight_layout()
    return fig


# =========================================================
# UI HELPERS
# =========================================================
def inject_custom_css() -> None:
    st.markdown(
        """
        <style>
            .block-container { padding-top: 2rem; padding-bottom: 2rem; }
            h1, h2, h3 { color: #1F4E79; }
            .stMetric { background:#F8FAFC; border:1px solid #E2E8F0;
                        border-radius:8px; padding:12px; }
            .stMetric > label { color:#475569; font-weight:600; }
            .stDownloadButton > button { background:#1F4E79; color:white;
                        border:none; border-radius:6px; font-weight:600; }
            .stDownloadButton > button:hover { background:#163a5a; color:white; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_empty_state() -> None:
    st.info("👈 Please upload an Excel file from the sidebar to begin.")
    with st.expander("📋 Expected File Structure"):
        st.markdown("""
        Your Excel file should have these columns in **Row 1**:

        | Required Column | Description |
        |----------------|-------------|
        | **Balance** | Account balance |
        | **Branch Name** | Branch identifier |
        | **Ac Type Desc** | Account type description |
        | **Main Code** | Main account code |
        | **Month Columns** | Any columns with Nepali month names (Shrawan, Bhadra, Ashoj, etc.) |

        *Month columns are auto-detected. Any header containing a month name is treated as a month column.*
        """)


def resolve_month_cols(
    all_month_cols: List[str],
    scope: str,
    from_month: Optional[str],
    to_month: Optional[str],
) -> Tuple[List[str], List[str]]:
    """Return (active_month_cols, drop_month_cols)."""
    if scope == "All Months":
        return all_month_cols, []

    if scope == "Last 12 Months":
        active = all_month_cols[-12:] if len(all_month_cols) >= 12 else all_month_cols
        drop = [c for c in all_month_cols if c not in active]
        return active, drop

    if scope == "Month Range":
        idx_from = all_month_cols.index(from_month)
        idx_to = all_month_cols.index(to_month)
        active = all_month_cols[idx_from: idx_to + 1]
        drop = [c for c in all_month_cols if c not in active]
        return active, drop

    return all_month_cols, []


def render_analytics_tab(
    df_branch: pd.DataFrame,
    df_ac_type: pd.DataFrame,
    branch_total_count: Dict[str, int],
    ac_type_total_count: Dict[str, int],
    cond_label: str,
) -> None:
    st.subheader("📈 Visual Analytics")
    st.caption(
        "Showing top 10 categories by % share "
        "(Matching Accounts ÷ Total Accounts in that category)."
    )

    # Branch chart
    df_branch_plot = df_branch[df_branch["Branch Name"] != "Grand Total"].copy()
    if len(df_branch_plot) > 0:
        df_branch_plot["Total Accounts"] = (
            df_branch_plot["Branch Name"]
            .map(branch_total_count).fillna(0).astype(int)
        )
        df_branch_plot["% Share"] = (
            df_branch_plot["No. of Zero Accounts"] /
            df_branch_plot["Total Accounts"] * 100
        ).replace([np.inf, -np.inf], 0).fillna(0)

        st.markdown("#### 🏦 Branch Analytics")
        st.markdown(f"**Top 10 Branches by % Share — {cond_label}**")

        df_plot = df_branch_plot.sort_values("% Share", ascending=False).head(10)
        df_plot = df_plot.sort_values("% Share", ascending=True)
        fig = plot_bar_h(
            df_plot, "% Share", "Branch Name",
            f"Top 10 Branches by % Share ({cond_label})",
            color=AppConstants.PRIMARY_COLOR,
        )
        st.pyplot(fig, use_container_width=True)
    else:
        st.info("No branch chart data available.")

    st.markdown("---")

    # Ac Type chart
    df_ac_plot = df_ac_type[df_ac_type["Ac Type Desc"] != "Grand Total"].copy()
    if len(df_ac_plot) > 0:
        df_ac_plot["Total Accounts"] = (
            df_ac_plot["Ac Type Desc"]
            .map(ac_type_total_count).fillna(0).astype(int)
        )
        df_ac_plot["% Share"] = (
            df_ac_plot["No. of Zero Accounts"] /
            df_ac_plot["Total Accounts"] * 100
        ).replace([np.inf, -np.inf], 0).fillna(0)

        st.markdown("#### 📑 Account Type Analytics")
        st.markdown(f"**Top 10 Account Types by % Share — {cond_label}**")

        df_plot = df_ac_plot.sort_values("% Share", ascending=False).head(10)
        df_plot = df_plot.sort_values("% Share", ascending=True)
        fig = plot_bar_h(
            df_plot, "% Share", "Ac Type Desc",
            f"Top 10 Account Types by % Share ({cond_label})",
            color=AppConstants.PRIMARY_COLOR,
        )
        st.pyplot(fig, use_container_width=True)
    else:
        st.info("No account type chart data available.")


# =========================================================
# MAIN APP
# =========================================================
def main() -> None:
    inject_custom_css()
    st.title(AppConstants.APP_TITLE)
    st.markdown(AppConstants.APP_SUBTITLE)

    # ---------------------------------------------------------
    # SIDEBAR
    # ---------------------------------------------------------
    with st.sidebar:
        st.header("⚙️ Configuration")
        uploaded_file = st.file_uploader(
            "📁 Upload Excel File", type=["xlsx", "xls", "xlsm"], key="file_uploader"
        )

        df_preview = None
        month_cols_preview: List[str] = []
        if uploaded_file is not None:
            try:
                df_preview = load_data(uploaded_file.getvalue())
                month_cols_preview = detect_month_columns(df_preview)
            except Exception:
                month_cols_preview = []

        st.markdown("---")
        st.subheader("📅 Month Scope")
        month_scope = st.radio(
            "Apply condition to:",
            options=list(AppConstants.MONTH_SCOPE_OPTIONS),
            index=0,
            help="Select which month columns the condition should evaluate.",
        )

        from_month = to_month = None
        if month_scope == "Month Range":
            if month_cols_preview:
                c1, c2 = st.columns(2)
                from_month = c1.selectbox(
                    "From", options=month_cols_preview, index=0, key="from_m"
                )
                to_month = c2.selectbox(
                    "To",
                    options=month_cols_preview,
                    index=len(month_cols_preview) - 1,
                    key="to_m",
                )
            else:
                st.warning("No month columns detected yet.")
        elif month_scope == "Last 12 Months":
            if month_cols_preview:
                st.caption(
                    f"Will use last 12 of {len(month_cols_preview)} detected month columns."
                )
            else:
                st.warning("No month columns detected yet.")

        st.markdown("---")
        month_criteria = st.text_input(
            "🎯 Month Condition",
            value=AppConstants.DEFAULT_CRITERIA,
            help="Examples: =0, >30, <90, >30<90, >=30<=90, <>0",
        )
        st.markdown("""
        **Condition Examples:**
        - `=0` — All months equal 0 (#N/A ignored)
        - `>30` — All months > 30
        - `<90` — All months < 90
        - `>30<90` — Between 30 and 90
        - `<>0` — Not equal to 0
        """)

    # ---------------------------------------------------------
    # EMPTY STATE
    # ---------------------------------------------------------
    if uploaded_file is None:
        render_empty_state()
        return

    # ---------------------------------------------------------
    # PARSE CRITERIA
    # ---------------------------------------------------------
    criteria = try_parse_criteria(month_criteria)
    if not criteria.is_valid:
        st.error(
            "❌ The month condition is not valid. "
            "Use examples such as: =0, >30, <90, >30<90, >=30<=90."
        )
        return

    # ---------------------------------------------------------
    # READ FILE
    # ---------------------------------------------------------
    try:
        with st.spinner("📖 Reading Excel file..."):
            df = load_data(uploaded_file.getvalue())
    except Exception as e:
        logger.exception("Failed to read uploaded file.")
        st.error(f"❌ Error reading file: {e}")
        return

    if df.empty:
        st.error("❌ No data found in the uploaded file.")
        return

    # ---------------------------------------------------------
    # FIND COLUMNS
    # ---------------------------------------------------------
    col_balance = find_header_column(df, "Balance")
    col_branch = find_header_column(df, "Branch Name")
    col_ac_type = find_header_column(df, "Ac Type Desc")
    col_main_code = find_header_column(df, "Main Code")

    missing = [c for c in AppConstants.REQUIRED_COLUMNS
               if find_header_column(df, c) is None]
    if missing:
        st.error(f"❌ Required columns not found: {', '.join(missing)}")
        st.info(f"Available columns: {list(df.columns)}")
        return

    text_columns_for_excel = [col_main_code] if col_main_code else []

    all_month_cols = detect_month_columns(df)
    if not all_month_cols:
        st.error("❌ No month columns were found in Row 1.")
        return

    # ---------------------------------------------------------
    # RESOLVE MONTH SCOPE
    # ---------------------------------------------------------
    if month_scope == "Month Range" and (from_month is None or to_month is None):
        st.error("❌ Please select both From and To months.")
        return
    if month_scope == "Month Range":
        idx_from = all_month_cols.index(from_month)
        idx_to = all_month_cols.index(to_month)
        if idx_from > idx_to:
            st.error(
                "❌ 'From Month' must come before or equal to 'To Month' in the sheet order."
            )
            return

    month_cols, drop_month_cols = resolve_month_cols(
        all_month_cols, month_scope, from_month, to_month,
    )

    if month_scope == "Last 12 Months" and len(all_month_cols) < 12:
        st.warning(
            f"⚠️ Only {len(all_month_cols)} month column(s) found; using all of them."
        )

    st.success(
        f"✅ Evaluating **{len(month_cols)}** month column(s) under scope: **{month_scope}**"
    )
    if len(month_cols) <= 12:
        st.caption(f"Columns: {', '.join(str(c) for c in month_cols)}")

    # ---------------------------------------------------------
    # BUILD CONTEXT
    # ---------------------------------------------------------
    ctx = AnalysisContext(
        month_scope=month_scope,
        from_month=from_month,
        to_month=to_month,
        criteria_raw=month_criteria,
        criteria=criteria,
        month_cols=month_cols,
        all_month_cols=all_month_cols,
        drop_month_cols=drop_month_cols,
    )

    # ---------------------------------------------------------
    # METRICS BAR
    # ---------------------------------------------------------
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("📊 Total Rows", f"{len(df):,}")
    c2.metric("📅 Month Columns (Total)", len(all_month_cols))
    c3.metric("🎯 Active Month Cols", len(month_cols))
    c4.metric("🔍 Condition", month_criteria)
    st.markdown("---")

    # ---------------------------------------------------------
    # RUN ANALYSIS
    # ---------------------------------------------------------
    try:
        with st.spinner("🔍 Analyzing data..."):
            result = run_analysis(
                df, ctx, col_balance, col_branch, col_ac_type, col_main_code,
            )
    except Exception as e:
        logger.exception("Analysis failed.")
        st.error(f"❌ Analysis failed: {e}")
        return

    # ---------------------------------------------------------
    # RESULT METRICS
    # ---------------------------------------------------------
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("✅ Accounts Found", f"{len(result.df_zero):,}")
    m2.metric("📋 Total Loan Accounts", f"{result.total_loan_account_count:,}")
    m3.metric("💰 Total Loan Balance", f"{result.total_loan_balance:,.2f}")
    m4.metric("📊 Output Sheets", "3")
    st.markdown("---")

    cond_label = month_criteria.strip()

    # ---------------------------------------------------------
    # TABS
    # ---------------------------------------------------------
    tab1, tab2, tab3, tab4 = st.tabs([
        "🎯 Matching Accounts",
        "🏦 Branch Summary",
        "📑 Ac Type Summary",
        "📈 Analytics",
    ])

    with tab1:
        st.subheader(
            f"Matching Accounts ({len(result.df_zero):,} rows) — Condition: {cond_label}"
        )
        if len(result.df_zero) > 0:
            st.dataframe(result.df_zero, use_container_width=True, height=600)
        else:
            st.info("No accounts matched the given condition.")

    with tab2:
        st.subheader("Branch Summary")
        if len(result.df_branch) > 0:
            st.dataframe(result.df_branch, use_container_width=True, height=500)
        else:
            st.info("No summary data available.")

    with tab3:
        st.subheader("Ac Type Summary")
        if len(result.df_ac_type) > 0:
            st.dataframe(result.df_ac_type, use_container_width=True, height=500)
        else:
            st.info("No summary data available.")

    with tab4:
        render_analytics_tab(
            result.df_branch,
            result.df_ac_type,
            result.branch_total_count,
            result.ac_type_total_count,
            cond_label,
        )

    # ---------------------------------------------------------
    # DOWNLOAD
    # ---------------------------------------------------------
    st.markdown("---")
    st.subheader("📥 Download Results")

    try:
        excel_bytes = build_excel_output(result, text_cols=text_columns_for_excel)
    except Exception as e:
        logger.exception("Excel generation failed.")
        st.error(f"❌ Excel generation failed: {e}")
        return

    col_dl1, _ = st.columns([1, 3])
    with col_dl1:
        st.download_button(
            label="📥 Download Excel Output",
            data=excel_bytes,
            file_name=(
                f"Month_Analysis_"
                f"{month_criteria.replace(' ', '_')}_"
                f"{month_scope.replace(' ', '_')}.xlsx"
            ),
            mime=AppConstants.XLSX_MIME,
            use_container_width=True,
        )

    st.success("✅ Analysis complete! Download the Excel file above.")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        logger.exception("Unhandled exception in main.")
        st.error("An unexpected error occurred. Please check the logs and try again.")
