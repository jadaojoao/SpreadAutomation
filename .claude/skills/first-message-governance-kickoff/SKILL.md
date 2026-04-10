---
name: first-message-governance-kickoff
description: "Kick off a new chat with governance-ready context loading and lane triage before implementation. Use for first message, start chat, initialize context, or when user invokes /first-message-governance-kickoff lane:<lane>. Reads mandatory operational docs, checks lane tasks/child tasks/PR state, and reports if execution is ready or blocked."
argument-hint: "Preferred: /first-message-governance-kickoff lane:frontend|backend|ops-quality. If no lane is provided, runs lane-agnostic kickoff."
---

# First Message Governance Kickoff

## Purpose
Use this skill at the start of a new chat to make the agent load the right operational context before coding.

This skill is startup-only and governance-first:
- it prepares context,
- it checks execution preconditions,
- it reports readiness,
- then hands off to implementation.

## Invocation Contract
Preferred invocation includes lane:
- `/first-message-governance-kickoff lane:frontend`
- `/first-message-governance-kickoff lane:backend`
- `/first-message-governance-kickoff lane:ops-quality`

Lane-agnostic invocation is allowed:
- `/first-message-governance-kickoff`

Lane parsing rules:
- accepted forms: `lane:frontend`, `lane:backend`, `lane:ops-quality`
- if no lane token is provided: run lane-agnostic mode
- if invalid lane token is provided: report invalid token and continue in lane-agnostic mode
- if multiple lanes are provided: stop and ask user to choose exactly one lane

## Mandatory Sources To Read First
Always read these before any implementation action:
- `AGENTS.md`
- `CLAUDE.md`
- `docs/governance/parallel-lanes.md`
- `docs/AGENTS.md`
- `README.md`

Also read path policy when governance enforcement is relevant:
- `.github/guardrails/path-policy.json`

## Area-Specific Required Reading
If the request already targets one of these areas, load required docs before edits:
- area `apps/web/**`: `docs/INTERFACE_MAP.md`
- area `apps/api/**`: `docs/INTERFACE_MAP.md` and `docs/V2_API_CONTRACT.md`
- area `src/**`: `docs/CONTEXT.md`
- area `docs/SITEMAP.MD`: `docs/INTERFACE_MAP.md`

## Startup Workflow

### Phase 1: Normalize Startup Mode
1. Parse lane token from the invocation.
2. Set mode:
   - `lane-aware` when lane is provided,
   - `lane-agnostic` when lane is missing.

### Phase 2: Load Operational Context
1. Read mandatory governance files.
2. Extract key rules:
   - issue-first requirement,
   - one task/owner/branch/worktree/PR,
   - child-task protocol,
   - critical-path and risk requirements,
   - completion criteria (checks green + merge confirmed).

### Phase 3: Triage Work State
If mode is `lane-aware`, check:
1. open task issues in the lane,
2. child tasks received from other lanes,
3. child tasks opened by this lane to other lanes,
4. open PRs linked to those issues,
5. pending `status:awaiting-consumption` deliveries,
6. **merged PRs with linked issue still open** — governance drift; report as `NEEDS_ISSUE_CLOSURE`,
7. **open PRs where guardrails passed and PR is not draft** — auto-merge may be pending; report state.

If mode is `lane-agnostic`, do a lightweight generic triage and mark lane checks as skipped.

### Phase 4: Readiness Decision
Return one status:
- `READY_TO_EXECUTE`
- `BLOCKED_AWAITING_CONSUMPTION`
- `BLOCKED_MISSING_TASK_ISSUE`
- `READY_WITH_TRIAGE_DEGRADED` (when remote issue/PR triage is unavailable)
- `NEEDS_ISSUE_CLOSURE` (merged PR found with linked issue still open — close issue before new work)

## Output Format
Return a concise startup briefing with:
1. `Mode`: lane-aware or lane-agnostic
2. `Lane`: explicit lane or `none`
3. `Mandatory docs loaded`: list
4. `Triage summary`: tasks, child tasks, PRs, awaiting-consumption, merged-but-open issues, auto-merge pending
5. `Blocking conditions`: explicit yes/no with reason
6. `Next allowed action`: exact next step before coding

## Hard Rules Enforced By This Skill
- never start changing versioned files before issue-first checks
- if lane-aware and awaiting-consumption exists, consume it before expanding scope
- do not ignore child-task dependencies
- do not proceed with lane-specific execution on ambiguous lane ownership
- do not treat docs/AGENTS.md as live backlog; use GitHub Issues as source of truth

## Question Policy
Keep startup questions minimal.
Ask only when necessary:
- lane ambiguous and request is clearly lane-owned
- triage unavailable and user must decide fallback confidence

If lane is not provided and request is not lane-specific, do not force a lane question.

## Examples
- `/first-message-governance-kickoff lane:backend`
- `/first-message-governance-kickoff lane:frontend`
- `/first-message-governance-kickoff lane:ops-quality`
- `/first-message-governance-kickoff` (generic kickoff)

## Completion Gate
This skill is complete only when it has:
- loaded mandatory sources,
- completed lane-aware or lane-agnostic triage,
- reported readiness/blocking status,
- provided a concrete next step for execution.
