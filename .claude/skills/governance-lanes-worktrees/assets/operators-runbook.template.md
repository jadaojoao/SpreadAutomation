# Operators Runbook

Use this document as your operational playbook after governance bootstrap.

## Daily Routine
1. Check open tasks in your lane.
2. Check child tasks received and requested.
3. Check open PRs linked to those tasks.
4. Resolve pending consumption before opening new cross-lane work.

## Discovery-first Rule
Before any apply or enforcement changes, review:
- existing `.github` governance artifacts,
- governance docs and runbooks,
- repository structure ownership map.

## Task Execution Flow
1. Open task issue with required labels and metadata.
2. Create worktree and task branch.
3. Implement in dedicated worktree only.
4. Open one PR with linked issue closure line.
5. Complete PR after checks are green.
6. Confirm linked issue closure and cleanup worktree.

## Bootstrap Commands
- PowerShell dry-run:
  powershell -ExecutionPolicy Bypass -File scripts/bootstrap-governance.ps1 -RepoPath . -Mode dry-run
- PowerShell apply:
  powershell -ExecutionPolicy Bypass -File scripts/bootstrap-governance.ps1 -RepoPath . -Mode apply -OverwritePolicy preserve
- Bash dry-run:
  ./scripts/bootstrap-governance.sh --repo . --mode dry-run

## Helper Commands
- Worktree create:
  powershell -ExecutionPolicy Bypass -File scripts/worktree_create.ps1 -Issue 27 -Slug sample-task -Lane ops-quality
- Worktree status:
  powershell -ExecutionPolicy Bypass -File scripts/worktree_status.ps1
- PR complete:
  powershell -ExecutionPolicy Bypass -File scripts/pr_complete.ps1 -Pr 28
- Worktree remove:
  powershell -ExecutionPolicy Bypass -File scripts/worktree_remove.ps1 -Issue 27 -Slug sample-task -Lane ops-quality

## Safety
- Do not skip required checks.
- Do not force-remove worktree without explicit approval.
- Do not use destructive Git reset commands.
- Do not auto-change protected or global Git settings without explicit approval.

## Escalation
If guardrails or apply fail:
1. inspect rollback guide,
2. restore from backup manifest,
3. rerun dry-run,
4. reattempt apply only after conflicts are resolved.
