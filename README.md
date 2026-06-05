# Revenue Analysis Dashboard

## What this app does
- Reads all monthly files matching `Solutions_Revenue_<Month>_<Year>*.xlsx` from the folder configured in the sidebar.
- Reads the template workbook (`Revenue Output Dashboard Sample.xlsx`) and uses the Dashboard sheet (or first matching master sheet) as customer/project master where available.
- Treats blanks/nulls as `0` for all financial values.
- Calculates:
  - `Margin = Revenue - Cost`
  - `GM% = ((Revenue - Cost) / Revenue) * 100`
- Supports:
  - Project Dashboard view
  - Account / Project summary (default: Account)
  - Revenue / Cost / Margin trends chart
  - FP project highlighting if cost increases or revenue changes over months
  - Projected margin based on observed monthly run-rate and project end date

## Folder assumptions
Default values auto-resolve as:
- Monthly files folder: parent of this app folder
- Template file: one level above monthly folder, named `Revenue Output Dashboard Sample.xlsx`

You can override both paths in the app sidebar.

## Run on Windows
1. Double-click `run_dashboard.bat`
2. Browser opens with the dashboard

## Run on macOS/Linux
1. Make launcher executable:
   - `chmod +x run_dashboard.sh`
2. Run:
   - `./run_dashboard.sh`

## SharePoint / multi-user note
- Place this app folder in the shared path.
- Each user can run it locally from that path.
- Shortcut behavior: launch directly to dashboard and data refreshes on launch.

## Security layer (optional SSO)
You can enable Microsoft organizational SSO directly in the app.

1. Set environment variables before launch:
  - `RAS_AUTH_MODE=sso`
  - `RAS_SSO_CLIENT_ID=<Azure App Registration Client ID>`
  - `RAS_SSO_TENANT_ID=<tenant id or organizations>`
  - `RAS_SSO_ALLOWED_DOMAIN=<yourcompany.com>`
  - `RAS_AUTH_SESSION_MINUTES=60`
  - Optional: `RAS_SSO_SCOPES=User.Read`
2. Launch the app.
3. Users must complete Microsoft device-code sign-in before any revenue data is loaded.

To disable auth for local testing:
- `RAS_AUTH_MODE=none`

## Data source modes
The app supports two source modes.

1. `RAS_DATA_SOURCE=local` (default)
  - Uses local folder and template file paths from the sidebar.
2. `RAS_DATA_SOURCE=sharepoint`
  - Pulls monthly files and template from SharePoint through Microsoft Graph.
  - Sidebar source fields become read-only and reflect environment configuration.

### Required variables for SharePoint mode
- `RAS_SP_TENANT_ID`
- `RAS_SP_CLIENT_ID`
- `RAS_SP_CLIENT_SECRET`
- `RAS_SP_SITE_ID`
- `RAS_SP_DRIVE_ID`
- `RAS_SP_FOLDER_PATH`
- `RAS_SP_TEMPLATE_FILE` (default: `Revenue Output Dashboard Sample.xlsx`)

The app mirrors SharePoint files into a local managed cache before parsing.

## GitHub organization bootstrap
For hosted rollout, create repository under organization `Myridius-India` and enable:
- branch protection on `main`
- required pull request review
- secret scanning and Dependabot

Automation helper:
- Run `scripts/bootstrap_github_repo.ps1 -Org Myridius-India -Repo revenue-analysis-dashboard -Push`
- Prerequisite: GitHub CLI (`gh`) installed and authenticated with org repo-create permission.

## Streamlit Cloud deployment (Git-based)
1. Go to Streamlit Community Cloud and create a new app from GitHub repo `Myridius-India/revenue-analysis-dashboard`.
2. Set entry file to `app.py`.
3. In app Secrets, add all `RAS_*` keys from `.env.example`.
4. Deploy and test sign-in + data refresh.

Note:
- The app mirrors `RAS_*` values from Streamlit Secrets into environment variables automatically at startup.

## Output behavior
- Sorted by `Customer Name` (A-Z)
- Account-level summary is selected by default
- All customer/project combinations from master are retained even if monthly data is missing
