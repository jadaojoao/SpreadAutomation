# Operation Flow (V2)

This skill is autonomous for discovery and validation, and explicit for risk-bearing actions.

## Phase 0: Scope and Guardrails
1. Confirm repository path and platform adapter.
2. Confirm autonomy mode and execution intent (dry-run or apply).
3. Declare non-destructive policy for this run.

Output:
- run intent summary

## Phase 1: Mandatory Discovery
1. Scan repository tree and detect domain boundaries.
2. Inspect `.github` governance assets:
	- issue templates,
	- PR template,
	- workflows,
	- guardrail policy files.
3. Inspect operational docs and contracts (`AGENTS`, runbooks, README, governance docs).
4. Triages GitHub workload context (issues/PRs) when available:
	- open tasks in lane,
	- child tasks received/requested,
	- pending consumption states.
5. Detect conflicts between current repo and target governance contract.

Output:
- discovery report
- lane/path suggestion map
- conflict and portability risk list

Gate:
- if discovery is incomplete and config requires strict discovery, block apply mode.

## Phase 2: Interactive Tuning
1. Run interview batches:
	- lane design,
	- labels/status/risk model,
	- issue/PR contract,
	- path policy strictness,
	- Git profile and scope,
	- overwrite/rollback policy.
2. Confirm ambiguous decisions before normalization.

Output:
- user-approved tuning decisions

## Phase 3: Normalize Configuration
1. Merge defaults, discovered facts, and user choices.
2. Build normalized lane model.
3. Build contracts (issue, PR, child-task).
4. Build path policy and risk floors.
5. Build guardrail enforcement settings.
6. Build git configuration policy and confirmation gates.

Output:
- normalized governance config object

## Phase 4: Preflight and Safety Planning
1. Validate repository is a Git repository.
2. Validate required tools for selected mode/platform.
3. Validate overwrite strategy.
4. Build backup plan with manifest.
5. Build rollback plan.

Output:
- preflight report
- safety decision (ready to dry-run/apply)

## Phase 5: Generate Package (Dry-Run First)
1. Generate issue templates and PR template.
2. Generate policy files and guardrail workflow.
3. Generate operational docs (contract/runbook/recovery).
4. Generate helper scripts.
5. Produce create/overwrite/skip plan.

Output:
- deterministic generation plan

Gate:
- if mode is dry-run, stop after report unless user explicitly upgrades to apply.

## Phase 6: Validate Generated Contract
1. Validate lane uniqueness and path ownership consistency.
2. Validate risk ranking and critical-group floors.
3. Validate issue/PR section aliases and parser expectations.
4. Validate child-task protocol consistency.
5. Validate guardrail rules for duplicate PR, write-set, and path classification.
6. Validate Git policy constraints and protected setting gates.

Output:
- pass/fail report
- remediation list

## Phase 7: Apply (Explicit Approval Required)
Only after explicit user approval:
1. Write files according to overwrite policy.
2. Record backup manifest.
3. Optionally apply approved Git settings.
4. Emit applied-change summary.

Output:
- apply report

## Phase 8: Recovery (On Failure)
1. Restore overwritten targets from backup manifest.
2. Remove created targets from failed run.
3. Re-validate repository state.
4. Report recovery status and remediation path.

Output:
- rollback report

## Phase 9: Operate Day-to-Day
1. Run lane/child-task/PR triage.
2. Work one task per branch/worktree/PR.
3. Enforce write-set and risk contract in PR.
4. Merge only after required checks.
5. Confirm issue closure and cleanup.

## Conservative Safety Policy
- automatic: discovery, normalization, dry-run, non-destructive validation
- explicit confirmation required: apply, force overwrite, merge completion actions, global/protected Git changes
- forbidden: destructive reset operations
