---
name: governance-lanes-worktrees
description: "Bootstrap and operate a discovery-first, fully configurable governance system for GitHub repositories: issue-first workflow, lanes, write-set and risk contracts, child-task protocol, worktrees, PR guardrails, and safe Git profiles. Use when users ask to adapt any repository to this governance model."
argument-hint: "Provide target repository path, preferred platform (v1 GitHub), and whether to run dry-run or apply."
user-invocable: true
disable-model-invocation: false
---

# Governance Lanes and Worktrees

## Purpose
Use this skill to adapt any repository to an enforceable governance model with:
- issue-first execution,
- lane ownership and risk boundaries,
- write-set contract validation,
- child task coordination between lanes,
- safe worktree and PR lifecycle,
- conservative, explicit Git configuration policy.

This skill is discovery-first. It does not apply governance before scanning the target repository.

## Mandatory Discovery Requirement
Before drafting configuration or writing files, this skill must explore the target repository and summarize findings.

Minimum required discovery scope:
1. GitHub backlog and review state relevant to governance adoption:
   - open task issues in the active lane,
   - child tasks received and requested,
   - open PRs linked to those issues,
   - pending `awaiting-consumption` states.
2. Repository governance assets:
   - `.github/ISSUE_TEMPLATE/**`,
   - `.github/PULL_REQUEST_TEMPLATE.md`,
   - `.github/workflows/**`,
   - `.github/guardrails/**`.
3. Repository operational documents:
   - root contract docs (`AGENTS.md`, equivalents),
   - runbooks and governance docs under `docs/**`,
   - onboarding docs (`README`, contribution docs).
4. Structure inventory for lane inference:
   - top-level folders,
   - critical runtime and contract paths,
   - existing script/tooling stack.

If discovery cannot be completed, the skill must stop before apply mode and report blockers.

## When To Use
Use this skill when users ask to:
- replicate this governance model in another repository,
- standardize branching/worktree/PR flow,
- enforce labels, sections, write-set, and risk gates,
- formalize cross-lane child-task protocol,
- bootstrap governance artifacts with safe defaults,
- configure Git with conservative policies and explicit confirmation gates.

## Inputs
Required or inferred inputs:
- target repository path,
- platform adapter (v1: GitHub),
- autonomy mode,
- bootstrap mode (`dry-run` or `apply`),
- overwrite policy (`preserve`, `merge`, `force`),
- locale (`en` default, `pt-BR` optional),
- lane model (preset or custom),
- branch/worktree patterns,
- label taxonomy and risk model,
- child-task strictness,
- guardrail strictness,
- Git profile and scope.

When inputs are omitted, defaults from [Config Schema](./references/config-schema.md) apply.

## Safe Defaults
- Platform: `github`
- Locale: `en`
- Autonomy mode: `conservative`
- Lane preset: `frontend`, `backend`, `ops-quality`
- Merge strategy: `squash`
- Bootstrap mode: `dry-run`
- Overwrite policy: `preserve`
- Path policy strictness: `standard`
- Child-task protocol: `enabled`
- Git profile: `conservative`
- Protected Git settings: never auto-changed without explicit confirmation

## Interactive Interview Contract
The skill must interview the user in short batches and adapt configuration accordingly.

Required interview phases:
1. Discovery confirmation:
   - validate what was inferred from repo scan.
2. Lane design:
   - preset vs custom lanes,
   - ownership patterns and exceptions.
3. Workflow contract:
   - labels, issue/PR sections, branch/worktree patterns.
4. Policy strictness:
   - critical groups, risk floors, disallowed domain mixes,
   - child-task protocol strictness.
5. Git policy:
   - profile and scope,
   - protected settings confirmation behavior.
6. Execution:
   - dry-run plan review,
   - explicit confirmation before apply.

## Autonomous Behavior Contract
The skill may execute automatically:
- repository discovery and conflict reporting,
- config normalization,
- dry-run generation plan,
- non-destructive validation.

The skill must require explicit confirmation for:
- apply mode writes,
- force overwrite,
- merge completion actions,
- remote branch deletion,
- global Git changes,
- protected Git settings.

Never perform destructive reset operations.

## Procedure
1. Discover target repository governance and structure.
2. Classify reusable vs conflicting existing artifacts.
3. Build normalized configuration object.
4. Run preflight and backup planning.
5. Generate templates, policy, workflow, and helper scripts.
6. Validate schema, lane/path/risk consistency, and guardrail contract.
7. Present dry-run report.
8. Apply only after explicit approval.
9. Run post-apply checks and report remediation if needed.

Detailed lifecycle: [Operation Flow](./references/operation-flow.md).

## Generated Artifacts
- [Governance Config Template](./assets/governance.config.template.yaml)
- [Path Policy Template](./assets/path-policy.template.json)
- [Task Issue Template](./assets/task.issue.template.yml)
- [Epic Issue Template](./assets/epic.issue.template.yml)
- [Issue Config Template](./assets/issue-config.yml)
- [PR Template](./assets/pull_request.template.md)
- [Guardrail Workflow Template](./assets/pr-issue-guardrails.workflow.yml)
- [Auto Merge Workflow Template](./assets/auto-merge.workflow.yml)
- [Post Merge Workflow Template](./assets/post-merge.workflow.yml)
- [Root Contract Template](./assets/AGENTS.template.md)
- [Parallel Lanes Template](./assets/parallel-lanes.template.md)
- [Operators Runbook Template](./assets/operators-runbook.template.md)
- [Rollback Recovery Template](./assets/rollback-recovery.template.md)
- [Bootstrap Script PowerShell](./scripts/bootstrap-governance.ps1)
- [Bootstrap Script Shell](./scripts/bootstrap-governance.sh)

## Merge Lifecycle Automation
The skill generates two complementary workflows that close the governance loop after PR validation:

**auto-merge.workflow.yml**
- Triggered by `workflow_run` on the guardrails check completing with `success`
- Finds the open PR for the commit, checks if it is not a draft
- If ready: squash-merges into the base branch and deletes the head branch
- If still draft: logs and skips without error

**post-merge.workflow.yml**
- Triggered by `pull_request` closed event when `merged == true`
- Extracts linked issue number from `Closes #N` / `Fixes #N` / `Resolves #N` in PR body
- Swaps `status:in-progress` → `status:completed` label on the linked issue
- Adds a completion comment referencing the merged PR number
- Closes the issue with `state_reason: completed`

Both workflows require `GITHUB_TOKEN` with `contents:write`, `pull-requests:write`, and `issues:write` permissions respectively.

**Path policy requirement**: The generated `path-policy.json` must include `.claude/**` in the `ops-quality` lane allowlist and in the `shared-governance` critical group patterns so skill files can be governed correctly.

## Quality Gate
A run is accepted only if:
- discovery was completed and summarized,
- config is valid and fully resolved,
- dry-run shows deterministic create/overwrite/skip plan,
- apply mode is safe and reversible,
- guardrails and policy are internally consistent,
- auto-merge and post-merge workflows are included in the output,
- generated docs explain day-to-day operation and recovery.

Use [Quality Checks](./references/quality-checks.md).

## Invocation Examples
- "Bootstrap governance for this repository in dry-run, using inferred lanes and conservative Git policy."
- "Adapt this monorepo with custom lanes mobile/platform/quality and strict path policy."
- "Scan the repo and propose governance conflicts before writing files."
- "Apply governance with English templates, PT-BR aliases, and preserve mode."

## References
- [Config Schema](./references/config-schema.md)
- [Operation Flow](./references/operation-flow.md)
- [Quality Checks](./references/quality-checks.md)
- [Tuning Questions](./references/tuning-questions.md)
- [PR Check Enforcement](./references/pr-check-enforcement.md)
- [Operators Runbook](./references/operators-runbook.md)
- [Rollback Recovery](./references/rollback-recovery.md)
