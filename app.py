from __future__ import annotations

import os
import time
from pathlib import Path

import msal
import pandas as pd
import plotly.express as px
import streamlit as st

from src.revenue_loader import AppConfig, build_project_snapshot, load_master_projects, load_revenue_files, merge_with_master, summarize
from src.sharepoint_data_source import DataSourceError, SharePointRuntimeConfig, mirror_sharepoint_files


st.set_page_config(
    page_title="Revenue Analysis Dashboard",
    page_icon="chart_with_upwards_trend",
    layout="wide",
)


def _bootstrap_runtime_env() -> None:
    # Streamlit Cloud stores values in st.secrets; mirror RAS_* keys into env.
    try:
        secret_items = dict(st.secrets)
    except Exception:
        return

    for key, value in secret_items.items():
        if not str(key).startswith("RAS_"):
            continue
        if key in os.environ:
            continue
        if isinstance(value, (str, int, float, bool)):
            os.environ[str(key)] = str(value)


_bootstrap_runtime_env()


CURRENCY_COLUMNS = [
    "rev_month",
    "cost_month",
    "margin_month",
    "rev_ytd",
    "cost_ytd",
    "margin_ytd",
    "rev_jtd",
    "cost_jtd",
    "margin_jtd",
    "projected_margin",
]

PCT_COLUMNS = ["target_gm_pct", "gm_month_pct", "gm_ytd_pct", "gm_jtd_pct"]
METRIC_CHOICES = ["GM%", "Revenue", "Cost"]
WINDOW_CHOICES = ["Last 3 months", "Last 6 months", "Last 12 months", "All months"]
APP_CACHE_VERSION = "2026-05-22-1"
DATA_SOURCE_MODE = os.getenv("RAS_DATA_SOURCE", "local").strip().lower()
AUTH_MODE = os.getenv("RAS_AUTH_MODE", "sso").strip().lower()
AUTH_SESSION_MINUTES = int(os.getenv("RAS_AUTH_SESSION_MINUTES", "60"))
SSO_TENANT_ID = os.getenv("RAS_SSO_TENANT_ID", "organizations")
# Default to Microsoft's first-party Azure CLI public client so hosted apps can
# still enforce sign-in when custom app registration secrets are not set yet.
SSO_CLIENT_ID = os.getenv("RAS_SSO_CLIENT_ID", "04b07795-8ddb-461a-bbee-02f9e1bf7b46")
SSO_REDIRECT_URI = os.getenv("RAS_SSO_REDIRECT_URI", "https://revenue-analysis-rasi-0605.streamlit.app/").strip()
SSO_ALLOWED_DOMAIN = os.getenv("RAS_SSO_ALLOWED_DOMAIN", "rcgglobalservices.com").strip().lower()
SSO_SCOPES = [scope.strip() for scope in os.getenv("RAS_SSO_SCOPES", "User.Read").split(",") if scope.strip()]

FRIENDLY_HEADERS = {
    "month": "Month",
    "customer_name": "Customer Name",
    "customer_id": "Customer ID",
    "project_name": "Project Name",
    "project_id": "Project ID",
    "engagement_model": "Engagement Model",
    "start_date": "Start Date",
    "end_date": "End Date",
    "target_gm_pct": "Target GM%",
    "rev_month": "Revenue (Month)",
    "cost_month": "Cost (Month)",
    "margin_month": "Margin (Month)",
    "gm_month_pct": "GM% (Month)",
    "rev_ytd": "Revenue YTD",
    "cost_ytd": "Cost YTD",
    "margin_ytd": "Margin YTD",
    "gm_ytd_pct": "GM% YTD",
    "rev_jtd": "Revenue JTD",
    "cost_jtd": "Cost JTD",
    "margin_jtd": "Margin JTD",
    "gm_jtd_pct": "GM% JTD",
    "projected_margin": "Projected Margin",
    "risk_status": "Risk Status",
    "risk_flag": "Risk Flag",
    "cost_increased": "Cost Increased",
    "revenue_changed": "Revenue Changed",
    "revenue_month": "Revenue (Month)",
    "cost_month": "Cost (Month)",
    "revenue_ytd": "Revenue YTD",
    "cost_ytd": "Cost YTD",
    "revenue_jtd": "Revenue JTD",
    "cost_jtd": "Cost JTD",
}


def _find_default_paths() -> AppConfig:
    app_folder = Path(__file__).resolve().parent
    data_folder = app_folder.parent
    template_file = data_folder.parent / "Revenue Output Dashboard Sample.xlsx"
    return AppConfig(data_folder=data_folder, template_file=template_file)


def _is_authenticated() -> bool:
    auth_ok = bool(st.session_state.get("auth_ok"))
    expires_at = float(st.session_state.get("auth_expires_at", 0.0) or 0.0)
    return auth_ok and expires_at > time.time()


def _logout_user() -> None:
    for key in [
        "auth_ok",
        "auth_email",
        "auth_name",
        "auth_expires_at",
        "sso_device_flow",
        "sso_auth_code_flow",
    ]:
        if key in st.session_state:
            del st.session_state[key]


def _clear_auth_session() -> None:
    # Keep in-progress device flow intact between reruns.
    for key in ["auth_ok", "auth_email", "auth_name", "auth_expires_at"]:
        st.session_state.pop(key, None)

def _build_sso_app() -> msal.PublicClientApplication:
    authority = f"https://login.microsoftonline.com/{SSO_TENANT_ID}"
    return msal.PublicClientApplication(client_id=SSO_CLIENT_ID, authority=authority)


def _extract_claim(claims: dict, *keys: str) -> str:
    for key in keys:
        value = claims.get(key)
        if value:
            return str(value)
    return ""


def _render_sso_login_gate() -> bool:
    st.subheader("Sign-in Required")
    st.info("Please sign in with your organizational Microsoft account to access this dashboard.")

    if not SSO_CLIENT_ID:
        st.error("SSO is enabled but RAS_SSO_CLIENT_ID is not configured.")
        return False

    if not SSO_REDIRECT_URI:
        st.error("SSO is enabled but RAS_SSO_REDIRECT_URI is not configured.")
        return False

    app = _build_sso_app()
    query_params = dict(st.query_params)
    if "error" in query_params:
        st.error(query_params.get("error_description") or "Microsoft sign-in failed.")
        st.query_params.clear()
        st.session_state.pop("sso_auth_code_flow", None)
        return False

    if "code" in query_params:
        flow = st.session_state.get("sso_auth_code_flow")
        if not flow:
            st.warning("Your sign-in session expired. Please start sign-in again.")
            st.query_params.clear()
            return False
        try:
            result = app.acquire_token_by_auth_code_flow(flow, query_params)
        except ValueError:
            st.error("Sign-in validation failed. Please try again.")
            st.query_params.clear()
            st.session_state.pop("sso_auth_code_flow", None)
            return False
        st.query_params.clear()
        st.session_state.pop("sso_auth_code_flow", None)
    else:
        flow = st.session_state.get("sso_auth_code_flow")
        if not flow:
            flow = app.initiate_auth_code_flow(
                scopes=SSO_SCOPES,
                redirect_uri=SSO_REDIRECT_URI,
                prompt="select_account",
            )
            st.session_state["sso_auth_code_flow"] = flow

        auth_uri = flow.get("auth_uri")
        if not auth_uri:
            st.error("Unable to start Microsoft sign-in. Please verify app registration settings.")
            return False

        st.link_button("Sign in with Microsoft", auth_uri, type="primary", use_container_width=True)
        if st.button("Reset Sign-In Session", use_container_width=True):
            st.session_state.pop("sso_auth_code_flow", None)
            st.query_params.clear()
            st.rerun()
        return False

    if "access_token" not in result:
        description = result.get("error_description", "Sign-in was not completed.")
        st.error(f"Sign-in failed: {description}")
        return False

    claims = result.get("id_token_claims", {})
    email = _extract_claim(claims, "preferred_username", "email", "upn").strip().lower()
    name = _extract_claim(claims, "name", "preferred_username", "upn") or "User"

    if not email:
        st.error("SSO succeeded, but no organizational email claim was returned.")
        return False

    if SSO_ALLOWED_DOMAIN and not email.endswith(f"@{SSO_ALLOWED_DOMAIN}"):
        st.error("This account is not part of the allowed organizational domain.")
        _logout_user()
        return False

    st.session_state["auth_ok"] = True
    st.session_state["auth_email"] = email
    st.session_state["auth_name"] = name
    st.session_state["auth_expires_at"] = time.time() + (AUTH_SESSION_MINUTES * 60)
    st.session_state.pop("sso_device_flow", None)
    st.success(f"Signed in as {email}")
    st.rerun()
    return False


def _enforce_authentication() -> bool:
    if AUTH_MODE != "sso":
        st.error("Unsupported auth mode. Set RAS_AUTH_MODE to 'sso'.")
        return False

    if not SSO_ALLOWED_DOMAIN:
        st.error("RAS_SSO_ALLOWED_DOMAIN must be configured to enforce organization-only access.")
        return False

    if _is_authenticated():
        return True

    _clear_auth_session()
    return _render_sso_login_gate()


@st.cache_data(show_spinner=False)
def load_all_data(data_folder: str, template_file: str, cache_version: str):
    revenue_data = load_revenue_files(Path(data_folder))
    master_data = load_master_projects(Path(template_file))
    merged = merge_with_master(revenue_data, master_data)
    snapshot = build_project_snapshot(merged)
    return merged, snapshot, len(revenue_data), len(master_data)


@st.cache_data(show_spinner=False)
def resolve_sharepoint_mirror(cache_version: str):
    config = SharePointRuntimeConfig.from_env()
    return mirror_sharepoint_files(config)


def _format_for_display(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for col in CURRENCY_COLUMNS:
        if col in out.columns:
            out[col] = out[col].fillna(0.0)
    for col in PCT_COLUMNS:
        if col in out.columns:
            out[col] = out[col].fillna(0.0)

    numeric_columns = out.select_dtypes(include=["number"]).columns
    for col in numeric_columns:
        out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0.0).round(2)

    if "risk_flag" in out.columns:
        out["risk_flag"] = out["risk_flag"].map({True: "Yes", False: "No"}).fillna("No")
    if "cost_increased" in out.columns:
        out["cost_increased"] = out["cost_increased"].map({True: "Yes", False: "No"}).fillna("No")
    if "revenue_changed" in out.columns:
        out["revenue_changed"] = out["revenue_changed"].map({True: "Yes", False: "No"}).fillna("No")

    if "start_date" in out.columns:
        out["start_date"] = pd.to_datetime(out["start_date"], errors="coerce").dt.strftime("%m/%d/%Y")
    if "end_date" in out.columns:
        out["end_date"] = pd.to_datetime(out["end_date"], errors="coerce").dt.strftime("%m/%d/%Y")
    if "month" in out.columns:
        out["month"] = pd.to_datetime(out["month"], errors="coerce").dt.strftime("%b-%Y")
    return out.rename(columns={k: v for k, v in FRIENDLY_HEADERS.items() if k in out.columns})


def _render_kpis(data: pd.DataFrame):
    revenue = float(data["rev_month"].sum()) if "rev_month" in data else 0.0
    cost = float(data["cost_month"].sum()) if "cost_month" in data else 0.0
    margin = revenue - cost
    gm_pct = (margin / revenue * 100) if revenue else 0.0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Revenue", f"${revenue:,.2f}")
    c2.metric("Cost", f"${cost:,.2f}")
    c3.metric("Margin", f"${margin:,.2f}")
    c4.metric("GM%", f"{gm_pct:,.2f}%")


def _style_dashboard_row(row: pd.Series):
    # Zebra shading improves readability; risk rows remain highlighted.
    base = "#f8fafc" if int(row.name) % 2 == 0 else "#eef2f7"
    is_risk = str(row.get("Risk Status", row.get("risk_status", ""))).strip().lower() == "risk"
    if is_risk:
        base = "#ffefcc"
    return [f"background-color: {base}; color: #111827"] * len(row)


def _as_two_decimal_styler(df: pd.DataFrame):
    numeric_columns = df.select_dtypes(include=["number"]).columns
    format_map = {col: "{:.2f}" for col in numeric_columns}
    return df.style.format(format_map)


def _select_month_window(all_months: list[pd.Timestamp], window_choice: str) -> list[pd.Timestamp]:
    if not all_months:
        return []
    if window_choice == "All months":
        return all_months

    size_map = {
        "Last 3 months": 3,
        "Last 6 months": 6,
        "Last 12 months": 12,
    }
    size = size_map.get(window_choice, 3)
    return all_months[-size:]


def _build_month_metric_pivot(
    data: pd.DataFrame,
    index_columns: list[str],
    selected_months: list[pd.Timestamp],
    selected_metrics: list[str],
) -> pd.DataFrame:
    if data.empty or not selected_months or not selected_metrics:
        return pd.DataFrame(columns=index_columns)

    working = data.copy()
    working = working.loc[working["month"].isin(selected_months)].copy()
    if working.empty:
        return pd.DataFrame(columns=index_columns)

    grouped = (
        working.groupby(index_columns + ["month"], dropna=False)
        .agg(
            rev_month=("rev_month", "sum"),
            cost_month=("cost_month", "sum"),
        )
        .reset_index()
    )
    grouped["margin_month"] = grouped["rev_month"] - grouped["cost_month"]
    grouped["gm_month_pct"] = (
        grouped["margin_month"] / grouped["rev_month"].where(grouped["rev_month"] != 0) * 100
    ).fillna(0)
    grouped["month_label"] = pd.to_datetime(grouped["month"]).dt.strftime("%b-%Y")

    metric_map = {
        "GM%": ("gm_month_pct", "GM%"),
        "Revenue": ("rev_month", "Revenue"),
        "Cost": ("cost_month", "Cost"),
    }

    month_order = [pd.Timestamp(m).strftime("%b-%Y") for m in sorted(selected_months, reverse=True)]
    pivots = []
    for metric in selected_metrics:
        value_column, suffix = metric_map[metric]
        pivot = (
            grouped.pivot_table(
                index=index_columns,
                columns="month_label",
                values=value_column,
                aggfunc="sum",
                fill_value=0.0,
            )
            .reindex(columns=month_order, fill_value=0.0)
            .reset_index()
        )
        rename_cols = {m: f"{m} {suffix}" for m in month_order if m in pivot.columns}
        pivot = pivot.rename(columns=rename_cols)
        pivots.append(pivot)

    merged_pivot = pivots[0]
    for pivot in pivots[1:]:
        merged_pivot = merged_pivot.merge(pivot, on=index_columns, how="outer")
    return merged_pivot.fillna(0.0)


def main():
    st.title("Revenue Analysis Dashboard")
    st.caption("Auto-refreshes source files on each run. Blank and null values are treated as 0.")

    if not _enforce_authentication():
        return

    defaults = _find_default_paths()
    data_folder = str(defaults.data_folder)
    template_file = str(defaults.template_file)

    with st.sidebar:
        st.header("Access")
        st.write(f"Signed in as: {st.session_state.get('auth_email', '')}")
        if st.button("Logout", use_container_width=True):
            _logout_user()
            st.rerun()

        st.header("Data Source")
        st.caption("Source files and template are fixed by deployment configuration.")

        if st.button("Refresh Data", use_container_width=True):
            load_all_data.clear()
            resolve_sharepoint_mirror.clear()

    if DATA_SOURCE_MODE == "sharepoint":
        try:
            with st.spinner("Syncing monthly files from SharePoint..."):
                mirror_result = resolve_sharepoint_mirror(APP_CACHE_VERSION)
            data_folder = str(mirror_result.data_folder)
            template_file = str(mirror_result.template_file)
        except DataSourceError as exc:
            st.error(str(exc))
            st.stop()
        except Exception as exc:
            st.error(f"Unexpected SharePoint sync error: {exc}")
            st.stop()

    merged, snapshot, revenue_count, master_count = load_all_data(data_folder, template_file, APP_CACHE_VERSION)

    if merged.empty:
        if revenue_count == 0:
            st.warning("No monthly files found. Expected pattern: Solutions_Revenue_<Month>_<Year>*.xlsx")
        elif master_count == 0:
            st.warning(
                "Template master could not be loaded (or has no Customer ID + Project ID rows). "
                "Only accounts/projects from the template are allowed, so no output can be shown."
            )
        else:
            st.warning("No matching account/project rows found between monthly files and template master.")
        return

    merged["customer_name"] = merged["customer_name"].fillna("")
    merged = merged.sort_values(["customer_name", "project_name", "month"], na_position="last")

    tab_dashboard, tab_summary, tab_trends = st.tabs(["Dashboard", "Summary", "Trends"])

    with tab_dashboard:
        st.subheader("Project Dashboard")

        month_values = sorted([pd.Timestamp(m) for m in merged["month"].dropna().unique()])
        account_options = sorted(name for name in merged["customer_name"].dropna().unique() if name)

        control_col1, control_col2, control_col3, control_col4, control_col5 = st.columns(
            [1.5, 2.0, 2.0, 2.0, 1.0]
        )

        with control_col1:
            window_choice = st.selectbox(
                "Month Window",
                options=WINDOW_CHOICES,
                index=0,
            )

        with control_col2:
            selected_accounts = st.multiselect(
                "Account Filter",
                options=account_options,
                default=[],
                placeholder="All accounts",
            )

        account_filtered = merged.copy()
        if selected_accounts:
            account_filtered = account_filtered.loc[account_filtered["customer_name"].isin(selected_accounts)]

        project_options = sorted(name for name in account_filtered["project_name"].dropna().unique() if name)
        with control_col3:
            selected_projects = st.multiselect(
                "Project Filter",
                options=project_options,
                default=[],
                placeholder="All projects",
            )

        with control_col4:
            selected_metrics = st.multiselect(
                "Metrics",
                options=METRIC_CHOICES,
                default=["GM%"],
                placeholder="Select metrics",
            )

        with control_col5:
            compact_view = st.toggle(
                "Compact Columns",
                value=True,
                help="Default on to reduce horizontal scrolling and improve readability.",
            )

        if not selected_metrics:
            selected_metrics = ["GM%"]

        filtered_merged = account_filtered.copy()
        if selected_projects:
            filtered_merged = filtered_merged.loc[filtered_merged["project_name"].isin(selected_projects)]

        selected_months = _select_month_window(month_values, window_choice)
        filtered_window = filtered_merged.loc[filtered_merged["month"].isin(selected_months)].copy()

        if filtered_window.empty and selected_months:
            st.warning("No monthly rows available for selected filters and month window.")
            return

        _render_kpis(filtered_window)

        latest_snapshot = snapshot.copy()
        latest_snapshot = latest_snapshot[[
            "customer_name",
            "customer_id",
            "project_name",
            "project_id",
            "engagement_model",
            "start_date",
            "end_date",
            "target_gm_pct",
            "projected_margin",
            "cost_increased",
            "revenue_changed",
        ]]

        if selected_accounts:
            latest_snapshot = latest_snapshot.loc[latest_snapshot["customer_name"].isin(selected_accounts)]
        if selected_projects:
            latest_snapshot = latest_snapshot.loc[latest_snapshot["project_name"].isin(selected_projects)]

        project_month_pivot = _build_month_metric_pivot(
            filtered_window,
            index_columns=["customer_id", "project_id"],
            selected_months=selected_months,
            selected_metrics=selected_metrics,
        )

        dashboard = latest_snapshot.merge(project_month_pivot, on=["customer_id", "project_id"], how="left")
        dashboard = dashboard.fillna(0)

        dashboard["risk_flag"] = dashboard["cost_increased"].fillna(False) | dashboard["revenue_changed"].fillna(False)
        dashboard["risk_status"] = dashboard["risk_flag"].map({True: "Risk", False: "Normal"})

        dynamic_columns = [
            col for col in dashboard.columns if any(col.endswith(suffix) for suffix in ["GM%", "Revenue", "Cost"])
        ]

        compact_columns = [
            "customer_name",
            "customer_id",
            "project_name",
            "project_id",
            "engagement_model",
            "target_gm_pct",
            "projected_margin",
            "risk_status",
            "risk_flag",
        ]

        full_columns = [
            "customer_name",
            "customer_id",
            "project_name",
            "project_id",
            "engagement_model",
            "start_date",
            "end_date",
            "target_gm_pct",
            "projected_margin",
            "risk_status",
            "cost_increased",
            "revenue_changed",
            "risk_flag",
        ]

        dashboard = dashboard[(compact_columns if compact_view else full_columns) + dynamic_columns]

        dashboard = dashboard.sort_values(["customer_name", "project_name"], na_position="last").fillna(0)
        display_dashboard = _format_for_display(dashboard.reset_index(drop=True))

        styled = _as_two_decimal_styler(display_dashboard).apply(_style_dashboard_row, axis=1)
        st.dataframe(styled, use_container_width=True, height=380)

        st.caption("FP projects are highlighted when cost rises across months or revenue changes across months.")

    with tab_summary:
        st.subheader("Summary")
        summary_col1, summary_col2, summary_col3 = st.columns([1.4, 1.4, 2.2])
        with summary_col1:
            summary_level = st.radio(
                "Summary View",
                options=["Account", "Project"],
                index=0,
                horizontal=True,
                help="Default is Account summary.",
            )
        with summary_col2:
            summary_window = st.selectbox("Month Window", options=WINDOW_CHOICES, index=0, key="summary_window")
        with summary_col3:
            summary_metrics = st.multiselect(
                "Metrics",
                options=METRIC_CHOICES,
                default=["GM%"],
                key="summary_metrics",
            )

        if not summary_metrics:
            summary_metrics = ["GM%"]

        summary_months = _select_month_window(month_values, summary_window)
        summary_source = merged.loc[merged["month"].isin(summary_months)].copy()
        summary_index = ["customer_name"] if summary_level == "Account" else ["customer_name", "project_name"]
        summary_wide = _build_month_metric_pivot(summary_source, summary_index, summary_months, summary_metrics)
        summary_wide = summary_wide.sort_values(summary_index, na_position="last").fillna(0)
        display_summary = _format_for_display(summary_wide)
        st.dataframe(_as_two_decimal_styler(display_summary), use_container_width=True, height=560)

    with tab_trends:
        st.subheader("Monthly Trends")
        trend_col1, trend_col2, trend_col3 = st.columns(3)
        with trend_col1:
            metric = st.selectbox(
                "Metric",
                options=["Revenue", "Cost", "Margin"],
                index=0,
            )

        trend_account_options = sorted(name for name in merged["customer_name"].dropna().unique() if name)
        with trend_col2:
            trend_selected_accounts = st.multiselect(
                "Account Filter",
                options=trend_account_options,
                default=[],
                placeholder="All accounts",
                key="trend_account_filter",
            )

        trend_filtered = merged.copy()
        if trend_selected_accounts:
            trend_filtered = trend_filtered.loc[trend_filtered["customer_name"].isin(trend_selected_accounts)]

        trend_project_options = sorted(name for name in trend_filtered["project_name"].dropna().unique() if name)
        with trend_col3:
            trend_selected_projects = st.multiselect(
                "Project Filter",
                options=trend_project_options,
                default=[],
                placeholder="All projects",
                key="trend_project_filter",
            )

        if trend_selected_projects:
            trend_filtered = trend_filtered.loc[trend_filtered["project_name"].isin(trend_selected_projects)]

        field_map = {"Revenue": "rev_month", "Cost": "cost_month", "Margin": "margin_month"}
        y_col = field_map[metric]

        trend = (
            trend_filtered.groupby(["month", "customer_name"], dropna=False)[["rev_month", "cost_month"]]
            .sum()
            .reset_index()
        )
        trend["margin_month"] = trend["rev_month"] - trend["cost_month"]

        chart = px.line(
            trend.sort_values("month"),
            x="month",
            y=y_col,
            color="customer_name",
            markers=True,
            title=f"{metric} Trend by Customer",
        )
        chart.update_layout(xaxis_title="Month", yaxis_title=metric)
        st.plotly_chart(chart, use_container_width=True)

        display_trend = _format_for_display(trend[["month", "customer_name", "rev_month", "cost_month", "margin_month"]])
        st.dataframe(_as_two_decimal_styler(display_trend), use_container_width=True, height=320)


if __name__ == "__main__":
    main()
