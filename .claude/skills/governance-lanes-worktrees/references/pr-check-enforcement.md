# PR Check Enforcement (V2)

This document defines guardrail behavior expected from the generated workflow.

## Core Contract Rules
- Branch must match configured task pattern.
- PR body must include issue closure line matching branch issue.
- Linked issue must be open and labeled as task.
- Only one open PR per task issue is allowed.

## Label and Section Rules
- Exactly one label for each required namespace (`status`, `priority`, `area`, `risk`, `lane`).
- Required issue sections must exist.
- Section parsing must support configured aliases (for locale portability).
- Risk section value must match risk label.
- Lane section value must match lane label.
- Workspace section must match configured worktree pattern.

## Child Task Rules (when enabled)
- Child task must reference parent task.
- Child task must declare requester lane and consumption criteria.
- Requester lane must match parent lane.
- Parent issue must list child task reference.
- Parent issue must remain in configured blocked statuses while child is pending.

## Write-Set and Path Rules
- Every changed file must match declared write-set patterns.
- Every changed file must be classified by path policy.
- Lane allowlists must permit every changed path.
- Critical groups must enforce:
	- minimum risk floor,
	- draft requirement,
	- compatibility requirement when configured.
- Disallowed domain mixes must fail validation.

## Compatibility Rules
- If a changed file matches a group with compatibility requirement, PR body must include compatibility section.
- Accepted values come from config (`additive-only`, `breaking-approved`, or custom values).

## Drift and Safety Rules
- Guardrail should fail clearly when required policy files are missing or malformed.
- Guardrail messages should be locale-aware when configured.
- Overlapping critical groups for the same file should fail to prevent ambiguous policy evaluation.

## Completion Rules
- Required checks must pass before merge completion script proceeds.
- Linked issue must close after merge.
- Branch cleanup must be explicit and traceable.
