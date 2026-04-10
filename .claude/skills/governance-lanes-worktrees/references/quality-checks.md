# Quality Checks (V2)

Use this checklist to confirm the governance skill is production-ready for arbitrary repositories.

## 1. Structural Integrity
- `SKILL.md` has valid YAML frontmatter.
- skill folder name and `name` field match.
- description contains trigger keywords for discovery.
- all references and assets resolve with relative paths.
- shell and PowerShell scripts are present when declared.

## 2. Discovery Compliance
- run blocks apply mode when required discovery is incomplete.
- discovery report includes:
	- `.github` governance inventory,
	- docs and contract inventory,
	- repository structure map,
	- conflicts and portability risks.
- lane/PR triage is attempted and fallback behavior is explicit when remote access is unavailable.

## 3. Configuration Completeness
- schema covers all runtime decisions.
- defaults exist for optional fields.
- lane model is explicit and non-empty.
- risk model is ordered and consistent with critical groups.
- issue and PR contracts include alias support for localization.
- overwrite policy and rollback policy are configurable.
- Git policy covers scopes, profiles, protected settings, and validation rules.

## 4. Artifact Coherence
- issue templates match contract fields.
- PR template matches compatibility and evidence requirements.
- path policy matches lane/risk model.
- guardrail workflow enforcement matches contract and policy.
- runbook and rollback docs match generated scripts and flow.
- generated artifact list in config matches actual files.

## 5. Operational Safety
- conservative mode is default.
- explicit confirmation is required for apply, force overwrite, merge completion, and protected/global Git changes.
- backup manifest is created before overwrite when apply mode runs.
- rollback restores previous state on failure.
- no destructive reset operations are used.

## 6. Behavior Checks
- dry-run produces deterministic create/overwrite/skip report.
- apply mode is idempotent on second run.
- conflict report includes actionable remediation.
- protected Git settings are never modified automatically.

## 7. Verification Matrix
- greenfield repository scenario validated.
- existing-governance repository scenario validated.
- mixed-domain or monorepo-like scenario validated.
- guardrail pass/fail fixtures validated for key rules.
- locale variants validated (`en` and optional `pt-BR`).
- script parity validated between PowerShell and shell bootstrap.

## Completion Standard
The skill is accepted when it can:
1. run mandatory discovery,
2. produce a safe dry-run plan,
3. apply with explicit confirmation and rollback safety,
4. enforce lane/risk/write-set/child-task contracts through guardrails,
5. generate operator documentation that is consistent with the generated contract.
