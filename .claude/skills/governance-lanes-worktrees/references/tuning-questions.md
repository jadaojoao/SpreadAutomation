# Tuning Questions (Interactive Contract)

Use this guide to interview the user in small batches and derive a fully resolved governance configuration.

## Batch 1: Repository and Discovery Scope
1. What repository should receive governance bootstrap?
2. Should discovery be strict (block apply if scan is incomplete)?
3. Which sources are mandatory in this org:
	- `.github` artifacts,
	- governance docs,
	- top-level folder map,
	- issue/PR triage?

## Batch 2: Lane Model
1. Use preset lanes or custom lane set?
2. For each lane, what paths are owned?
3. Should any paths be shared across lanes?
4. Which lane may touch shared governance paths?

## Batch 3: Labels and Status Lifecycle
1. Keep default label families (`kind`, `status`, `priority`, `area`, `risk`, `lane`) or customize?
2. Which status values represent blocked and awaiting-consumption?
3. Do you need custom area taxonomy for this project domain?

## Batch 4: Issue and PR Contract
1. Which issue body sections are mandatory?
2. Which aliases/locales should be accepted for section parsing?
3. Which closure keywords should be accepted (`Closes`, `Fixes`, `Resolves`)?
4. Is compatibility declaration required only for critical groups or broader scope?

## Batch 5: Worktree and Branching
1. Branch pattern (default `task/<issue-number>-<slug>`) or custom?
2. Worktree root and pattern (default `.claude/worktrees/<lane>/<issue-number>-<slug>/`) or custom?
3. Base branch mode: auto-detect or fixed value?

## Batch 6: Child Task Protocol
1. Enable child-task protocol?
2. Should parent task be required to list child tasks?
3. Must parent status be blocked/awaiting-consumption while child is active?
4. Must requester lane equal parent lane?

## Batch 7: Path Policy and Risk
1. Which critical groups are required?
2. Minimum risk for each critical group?
3. Which domain mixes must be blocked?
4. Should unclassified paths fail guardrails?

## Batch 8: Git Policy
1. Git profile mode: conservative, strict, or disabled?
2. Scope: repo-local only or include global/worktree settings?
3. Which settings are protected and require explicit confirmation?
4. Should lane-specific Git profiles be applied?

## Batch 9: Execution and Safety
1. Default mode: dry-run or apply?
2. Overwrite policy: preserve, merge, or force?
3. Require backup manifest and rollback on failure?
4. Require explicit confirmation for apply and force actions?

## Batch 10: Localization and Deliverables
1. Locale default (`en` or `pt-BR`)?
2. Should templates be English-first with optional PT-BR aliases?
3. Which docs are mandatory outputs:
	- root governance contract,
	- parallel lanes guide,
	- operators runbook,
	- rollback guide?

## Decision Rules
- If user does not answer a question, apply safe defaults.
- If decisions conflict (for example, strict guardrails with permissive path policy), flag and request explicit resolution.
- Always present a normalized summary before apply mode.
