# AGENTS.md

Operational contract for contributors and coding agents.

## Source of truth
- Backlog lives in GitHub Issues.
- Pull requests close task issues, not epics.

## Discovery-first startup
At the start of each executable chat:
- check open tasks in your lane,
- check child tasks received from other lanes,
- check child tasks requested by your lane,
- check open PRs linked to those issues,
- resolve awaiting-consumption before expanding cross-lane scope.

## Mandatory flow
1. Find or create an open task issue before changing versioned files.
2. Ensure required labels exist: kind, status, priority, area, risk, lane.
3. Ensure issue body includes owner, lane, workspace, write-set, and risk.
4. Work in a dedicated worktree and branch.
5. Open a single official PR with Closes #<issue-number> in the body.
6. Keep task status in-progress while PR is open or waiting review.
7. Complete PR only after required checks are green and required approvals are done.
8. Confirm linked issue closure and branch cleanup before declaring task complete.

## Branch and worktree
- Branch pattern: <lane>/<issue-number>-<slug>
- Worktree pattern: .claude/worktrees/<lane>/<issue-number>-<slug>/
- Rule: one task, one owner, one branch, one worktree, one PR.

## Parallel work protocol
- Interfaces are additive-only by default during parallel execution.
- Breaking contract changes require a dedicated contract-sensitive task.
- If write-set collision happens across tasks, declare dependency before continuing.

## Child tasks
- If work crosses lane ownership, open formal child tasks.
- Child tasks must include Task mae, Lane solicitante, and Criterio de consumo.
- Parent task remains blocked or awaiting-consumption until requester confirms consumption.

## Risk policy
- safe: isolated write-set
- shared: shared or critical paths, draft PR required
- contract-sensitive: public contract changes, draft PR and compatibility section required

## Completion
A task is complete only when PR is merged and linked issue is closed.

## AI completion gate
- Never report task completion while PR is still open.
- Never finalize a task while the linked issue remains open.
- If external review or approval is pending, report the pending gate explicitly and keep the task open.

## Safety constraints
- Required checks must stay green for merge completion.
- Never force-remove worktrees without explicit approval.
- Never use destructive Git reset commands.
