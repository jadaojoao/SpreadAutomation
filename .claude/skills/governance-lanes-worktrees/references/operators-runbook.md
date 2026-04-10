# Operators Runbook (V2)

This runbook describes day-to-day governance operation after bootstrap.

## Daily Start Checklist
1. Check open tasks in your lane.
2. Check child tasks received from other lanes.
3. Check child tasks requested by your lane.
4. Check open PRs linked to those issues.
5. Resolve `awaiting-consumption` before opening new cross-lane work.

## Task Execution Flow
1. Open or find a task issue.
2. Validate labels and required sections.
3. Create dedicated branch and worktree.
4. Implement and push validated checkpoints.
5. Open one PR with issue closure line.
6. Pass guardrails and required checks.
7. Complete merge only after checks are green.
8. Confirm issue closure, remove remote branch (if configured), remove worktree.

## Bootstrap and Validation Commands
### PowerShell dry-run
```powershell
powershell -ExecutionPolicy Bypass -File scripts/bootstrap-governance.ps1 -RepoPath . -Mode dry-run
```

### PowerShell apply (example)
```powershell
powershell -ExecutionPolicy Bypass -File scripts/bootstrap-governance.ps1 -RepoPath . -Mode apply -OverwritePolicy preserve
```

### Shell dry-run
```bash
./scripts/bootstrap-governance.sh --repo . --mode dry-run
```

## Worktree/PR Helper Commands
### Create task worktree
```powershell
powershell -ExecutionPolicy Bypass -File scripts/worktree_create.ps1 -Issue 27 -Slug improve-governance -Lane ops-quality
```

### Review active worktrees
```powershell
powershell -ExecutionPolicy Bypass -File scripts/worktree_status.ps1
```

### Complete pull request safely
```powershell
powershell -ExecutionPolicy Bypass -File scripts/pr_complete.ps1 -Pr 85
```

### Remove task worktree
```powershell
powershell -ExecutionPolicy Bypass -File scripts/worktree_remove.ps1 -Issue 27 -Slug improve-governance -Lane ops-quality
```

## Child Task Protocol
- Open formal child task in owning lane when write-set crosses lane ownership.
- Parent remains blocked while child is open or under review.
- Parent moves to awaiting-consumption after child merge and before requester validation.
- Only requester lane confirms consumption and unblocks parent.

## Safety Rules
- Do not bypass required checks.
- Do not force-remove worktrees without explicit confirmation.
- Do not auto-change protected/global Git settings without explicit approval.
- Do not use destructive reset commands.

## Escalation
If apply or guardrails fail:
1. consult rollback guide,
2. restore from backup manifest,
3. rerun dry-run,
4. reopen apply only after conflict resolution.
