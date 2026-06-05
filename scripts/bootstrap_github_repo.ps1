param(
    [string]$Org = "Myridius-India",
    [string]$Repo = "revenue-analysis-dashboard",
    [switch]$Push
)

$ErrorActionPreference = "Stop"

if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
    throw "GitHub CLI (gh) is not installed. Install gh and login first."
}

$null = gh auth status

$repoFull = "$Org/$Repo"

$exists = $false
try {
    gh repo view $repoFull | Out-Null
    $exists = $true
} catch {
    $exists = $false
}

if (-not $exists) {
    gh repo create $repoFull --private --description "Revenue analysis dashboard" --confirm
}

if (-not (Test-Path .git)) {
    git init
    git branch -M main
}

$remotes = git remote
if ($remotes -notcontains "origin") {
    git remote add origin "https://github.com/$repoFull.git"
}

git add .

if (-not (git diff --cached --quiet)) {
    git commit -m "feat: hosted SharePoint mode, SSO gate, CI/CD bootstrap"
}

if ($Push) {
    git push -u origin main
}

Write-Host "Repository bootstrap complete for $repoFull"
