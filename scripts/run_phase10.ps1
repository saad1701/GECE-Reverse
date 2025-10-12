# scripts\run_phase10.ps1
# Fail on first error
$ErrorActionPreference = "Stop"

# Jump to repo root (this file lives in scripts\)
Set-Location (Split-Path $MyInvocation.MyCommand.Path -Parent)
Set-Location ..

# 0) Kill stray NUL if present (Windows quirk)
if (Test-Path -LiteralPath '.\NUL') {
    Remove-Item -LiteralPath '.\NUL' -Force -ErrorAction SilentlyContinue
}

# 1) Rebuild (use py -3 on Windows)
py -3 scripts\generate_manifest.py
py -3 scripts\validate_manifest.py
# render_report.py in this repo autodetects inputs; don't pass long flags
py -3 scripts\render_report.py

# 2) Git checks without pager
git --no-pager status
git --no-pager diff --name-status

Write-Host "`nPhase 10 script completed." -ForegroundColor Green

