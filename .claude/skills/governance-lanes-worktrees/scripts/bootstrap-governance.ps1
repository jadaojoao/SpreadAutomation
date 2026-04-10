param(
  [Parameter(Mandatory = $true)]
  [string]$RepoPath,

  [ValidateSet("dry-run", "apply")]
  [string]$Mode = "dry-run",

  [ValidateSet("preserve", "merge", "force")]
  [string]$OverwritePolicy = "preserve",

  [switch]$Force,

  [switch]$ConfigureGit,

  [ValidateSet("conservative", "strict", "disabled")]
  [string]$GitProfile = "conservative",

  [switch]$ApplyGlobalGitConfig,

  [string]$BaseBranch,

  [ValidateSet("en", "pt-BR")]
  [string]$Locale = "en"
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

if ($Force) {
  $OverwritePolicy = "force"
}

function Require-Command {
  param([string]$CommandName)

  if (-not (Get-Command $CommandName -ErrorAction SilentlyContinue)) {
    throw "Required command not found in PATH: $CommandName"
  }
}

function Ensure-Directory {
  param([string]$Path)

  if (-not (Test-Path $Path)) {
    New-Item -ItemType Directory -Path $Path -Force | Out-Null
  }
}

function Get-RelativeRepoPath {
  param(
    [string]$RepoRoot,
    [string]$AbsolutePath
  )

  if ($AbsolutePath.StartsWith($RepoRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
    return $AbsolutePath.Substring($RepoRoot.Length).TrimStart("\\", "/")
  }

  return $AbsolutePath
}

function Get-DefaultBranch {
  param(
    [string]$RepoRoot,
    [string]$FallbackBranch
  )

  if ($FallbackBranch) {
    return $FallbackBranch
  }

  $originHead = git -C $RepoRoot symbolic-ref refs/remotes/origin/HEAD 2>$null
  if ($LASTEXITCODE -eq 0 -and $originHead) {
    $parts = "$originHead".Trim() -split "/"
    if ($parts.Length -gt 0) {
      return $parts[-1]
    }
  }

  return "master"
}

function Backup-File {
  param(
    [string]$RepoRoot,
    [string]$AbsoluteTarget,
    [string]$BackupRoot
  )

  $relativeTarget = Get-RelativeRepoPath -RepoRoot $RepoRoot -AbsolutePath $AbsoluteTarget
  $backupPath = Join-Path $BackupRoot $relativeTarget
  Ensure-Directory (Split-Path $backupPath -Parent)
  Copy-Item -Path $AbsoluteTarget -Destination $backupPath -Force
  return $backupPath
}

function Restore-FromBackup {
  param(
    [string]$RepoRoot,
    [array]$CreatedTargets,
    [array]$OverwrittenTargets,
    [array]$MergeCandidates
  )

  foreach ($entry in $OverwrittenTargets) {
    $targetPath = Join-Path $RepoRoot $entry.Target
    Ensure-Directory (Split-Path $targetPath -Parent)
    Copy-Item -Path $entry.BackupPath -Destination $targetPath -Force
  }

  foreach ($target in $CreatedTargets) {
    $targetPath = Join-Path $RepoRoot $target
    if (Test-Path $targetPath) {
      Remove-Item -Path $targetPath -Force
    }
  }

  foreach ($candidate in $MergeCandidates) {
    $candidatePath = Join-Path $RepoRoot $candidate.Target
    if (Test-Path $candidatePath) {
      Remove-Item -Path $candidatePath -Force
    }
  }
}

function Set-GitSetting {
  param(
    [string]$RepoRoot,
    [string]$Scope,
    [string]$Key,
    [string]$Value
  )

  if ($Scope -eq "global") {
    git config --global $Key $Value | Out-Null
    return
  }

  git -C $RepoRoot config $Key $Value | Out-Null
}

function Apply-GitProfile {
  param(
    [string]$RepoRoot,
    [string]$Profile,
    [bool]$ApplyGlobal,
    [string]$DetectedBaseBranch
  )

  $applied = @()

  if ($Profile -eq "disabled") {
    return $applied
  }

  $localSettings = [ordered]@{
    "fetch.prune" = "true"
    "pull.rebase" = "true"
    "rebase.autostash" = "true"
    "push.default" = "simple"
    "core.safecrlf" = "true"
  }

  if ($env:OS -eq "Windows_NT") {
    $localSettings["core.autocrlf"] = "true"
  } else {
    $localSettings["core.autocrlf"] = "input"
  }

  if ($Profile -eq "strict") {
    $localSettings["merge.ff"] = "false"
  }

  foreach ($key in $localSettings.Keys) {
    $value = $localSettings[$key]
    Set-GitSetting -RepoRoot $RepoRoot -Scope "local" -Key $key -Value $value
    $applied += [PSCustomObject]@{
      Scope = "local"
      Key = $key
      Value = $value
    }
  }

  if ($ApplyGlobal) {
    $globalSettings = [ordered]@{
      "init.defaultBranch" = $DetectedBaseBranch
      "push.default" = "simple"
    }

    foreach ($key in $globalSettings.Keys) {
      $value = $globalSettings[$key]
      Set-GitSetting -RepoRoot $RepoRoot -Scope "global" -Key $key -Value $value
      $applied += [PSCustomObject]@{
        Scope = "global"
        Key = $key
        Value = $value
      }
    }
  }

  return $applied
}

Require-Command -CommandName "git"

if (-not (Test-Path $RepoPath)) {
  throw "RepoPath does not exist: $RepoPath"
}

$repoResolved = Resolve-Path $RepoPath
$repoRoot = $repoResolved.Path
$detectedBaseBranch = Get-DefaultBranch -RepoRoot $repoRoot -FallbackBranch $BaseBranch

$insideRepo = git -C $repoRoot rev-parse --is-inside-work-tree 2>$null
if ($LASTEXITCODE -ne 0 -or "$insideRepo".Trim() -ne "true") {
  throw "RepoPath is not a Git repository: $repoRoot"
}

$dirtyStatus = git -C $repoRoot status --porcelain
if ($dirtyStatus) {
  Write-Warning "Repository has uncommitted changes. Continue with caution."
}

$skillRoot = Split-Path -Parent $PSScriptRoot
$assetsRoot = Join-Path $skillRoot "assets"

$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backupRoot = Join-Path $repoRoot ".governance-bootstrap/backups/$timestamp"
$mergeCandidatesRoot = Join-Path $repoRoot ".governance-bootstrap/merge-candidates/$timestamp"
$manifestPath = Join-Path $backupRoot "manifest.json"

$mappings = @(
  @{ Source = "AGENTS.template.md"; Target = "AGENTS.md" },
  @{ Source = "parallel-lanes.template.md"; Target = "docs/governance/parallel-lanes.md" },
  @{ Source = "operators-runbook.template.md"; Target = "docs/governance/operators-runbook.md" },
  @{ Source = "rollback-recovery.template.md"; Target = "docs/governance/rollback-recovery.md" },
  @{ Source = "task.issue.template.yml"; Target = ".github/ISSUE_TEMPLATE/task.yml" },
  @{ Source = "epic.issue.template.yml"; Target = ".github/ISSUE_TEMPLATE/epic.yml" },
  @{ Source = "issue-config.yml"; Target = ".github/ISSUE_TEMPLATE/config.yml" },
  @{ Source = "pull_request.template.md"; Target = ".github/PULL_REQUEST_TEMPLATE.md" },
  @{ Source = "path-policy.template.json"; Target = ".github/guardrails/path-policy.json" },
  @{ Source = "pr-issue-guardrails.workflow.yml"; Target = ".github/workflows/pr-issue-guardrails.yml" },
  @{ Source = "governance.config.template.yaml"; Target = ".github/governance.config.yaml" }
)

Write-Output "Governance bootstrap mode: $Mode"
Write-Output "Repository: $repoRoot"
Write-Output "Locale: $Locale"
Write-Output "Overwrite policy: $OverwritePolicy"
Write-Output "Detected base branch: $detectedBaseBranch"
Write-Output "Git profile requested: $GitProfile"
Write-Output ""

$planned = @()
$createdTargets = @()
$overwrittenTargets = @()
$mergeCandidates = @()
$gitApplied = @()

foreach ($item in $mappings) {
  $sourcePath = Join-Path $assetsRoot $item.Source
  if (-not (Test-Path $sourcePath)) {
    throw "Missing asset: $sourcePath"
  }

  $targetPath = Join-Path $repoRoot $item.Target
  $targetExists = Test-Path $targetPath

  $action = "create"
  if ($targetExists) {
    switch ($OverwritePolicy) {
      "preserve" { $action = "skip" }
      "merge" { $action = "merge-candidate" }
      "force" { $action = "overwrite" }
    }
  }

  $planned += [PSCustomObject]@{
    Action = $action
    Source = $item.Source
    Target = $item.Target
  }
}

$planned | Format-Table -AutoSize | Out-String | Write-Output

if ($Mode -eq "dry-run") {
  Write-Output "Dry-run complete. No files were changed."
  return
}

Ensure-Directory $backupRoot
Ensure-Directory $mergeCandidatesRoot

try {
  foreach ($item in $mappings) {
    $sourcePath = Join-Path $assetsRoot $item.Source
    $targetPath = Join-Path $repoRoot $item.Target
    $targetExists = Test-Path $targetPath

    $action = "create"
    if ($targetExists) {
      switch ($OverwritePolicy) {
        "preserve" { $action = "skip" }
        "merge" { $action = "merge-candidate" }
        "force" { $action = "overwrite" }
      }
    }

    if ($action -eq "skip") {
      continue
    }

    if ($action -eq "merge-candidate") {
      $candidatePath = Join-Path $mergeCandidatesRoot $item.Target
      Ensure-Directory (Split-Path $candidatePath -Parent)
      Copy-Item -Path $sourcePath -Destination $candidatePath -Force
      $mergeCandidates += [PSCustomObject]@{
        Source = $item.Source
        Target = Get-RelativeRepoPath -RepoRoot $repoRoot -AbsolutePath $candidatePath
      }
      continue
    }

    Ensure-Directory (Split-Path $targetPath -Parent)

    if ($targetExists) {
      $backupPath = Backup-File -RepoRoot $repoRoot -AbsoluteTarget $targetPath -BackupRoot $backupRoot
      $overwrittenTargets += [PSCustomObject]@{
        Target = $item.Target
        BackupPath = $backupPath
      }
    } else {
      $createdTargets += $item.Target
    }

    Copy-Item -Path $sourcePath -Destination $targetPath -Force
  }

  if ($ConfigureGit) {
    $gitApplied = Apply-GitProfile -RepoRoot $repoRoot -Profile $GitProfile -ApplyGlobal $ApplyGlobalGitConfig.IsPresent -DetectedBaseBranch $detectedBaseBranch
  }

  $manifest = [PSCustomObject]@{
    timestamp = $timestamp
    mode = $Mode
    locale = $Locale
    overwrite_policy = $OverwritePolicy
    repo_root = $repoRoot
    base_branch = $detectedBaseBranch
    created = $createdTargets
    overwritten = $overwrittenTargets
    merge_candidates = $mergeCandidates
    git_applied = $gitApplied
  }

  $manifest | ConvertTo-Json -Depth 6 | Set-Content -Path $manifestPath -Encoding UTF8
}
catch {
  Write-Warning "Bootstrap failed. Attempting rollback from backup manifest state."
  Restore-FromBackup -RepoRoot $repoRoot -CreatedTargets $createdTargets -OverwrittenTargets $overwrittenTargets -MergeCandidates $mergeCandidates
  throw
}

Write-Output "Apply mode completed."
Write-Output "Backup snapshot: $backupRoot"
Write-Output "Backup manifest: $manifestPath"

if ($mergeCandidates.Count -gt 0) {
  Write-Output ""
  Write-Output "Merge candidates were generated instead of overwriting existing files:"
  $mergeCandidates | Format-Table -AutoSize | Out-String | Write-Output
}

if ($ConfigureGit) {
  Write-Output ""
  Write-Output "Git settings applied:"
  if ($gitApplied.Count -gt 0) {
    $gitApplied | Format-Table -AutoSize | Out-String | Write-Output
  } else {
    Write-Output "No git settings were applied."
  }
}

Write-Output "Review generated files and customize lane ownership, labels, and path policy before opening PRs."
