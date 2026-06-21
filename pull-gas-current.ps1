$ErrorActionPreference = "Stop"

$workspace = Split-Path -Parent $MyInvocation.MyCommand.Path
$env:npm_config_cache = Join-Path $workspace ".npm-cache"
$authPath = Join-Path $workspace ".clasp-auth"
$projectPath = Join-Path $workspace "gas-current"

Write-Host "Pulling current Google Apps Script project into: $projectPath"
npx.cmd @google/clasp pull --auth $authPath --project $projectPath
