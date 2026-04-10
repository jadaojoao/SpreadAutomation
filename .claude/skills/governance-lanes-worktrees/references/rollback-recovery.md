# Rollback and Recovery (V2)

Use this guide when bootstrap apply or guardrails produce unexpected failures.

## Recovery Principles
- prefer additive corrections over destructive actions,
- keep issue and PR history intact,
- recover deterministically from backup manifest,
- rerun dry-run before any second apply attempt.

## Bootstrap Failure Recovery
1. Locate backup root:
   - `.governance-bootstrap/backups/<timestamp>/`
2. Open backup manifest (if generated) and inspect:
   - overwritten targets,
   - created targets,
   - generated checksum metadata.
3. Restore overwritten files from backup snapshot.
4. Remove created files from failed run.
5. Validate repository status.
6. Rerun bootstrap in dry-run and inspect conflicts.

## Guardrail Failure Recovery
1. Identify failing rule category:
   - labels,
   - section parsing,
   - write-set coverage,
   - path policy classification,
   - risk floor,
   - child-task consistency,
   - duplicate PR per task.
2. Fix configuration contract or issue/PR metadata.
3. If policy changed, update path policy and workflow artifacts together.
4. Re-run checks through PR synchronization.

## Parser and Locale Drift Recovery
If templates were localized but guardrails still parse old headings:
1. confirm section aliases in config,
2. regenerate workflow/template artifacts from same normalized config,
3. rerun dry-run before apply.

## Template-Only Recovery
When only forms/checklists are wrong:
1. restore affected issue or PR templates,
2. keep policy/workflow untouched if valid,
3. rerun guardrail contract validation.

## Git Configuration Recovery
1. inspect current config with origin:
   - `git config --list --show-origin`
2. restore expected values from policy or backup notes.
3. avoid global/protected changes unless explicitly approved.
4. verify pull/fetch/rebase behavior after restore.

## Safe-Mode Runbook
If confidence is low:
1. force dry-run mode,
2. set overwrite policy to preserve,
3. disable non-essential Git apply actions,
4. require explicit confirmation for every risky step.

## Post-Recovery Checklist
- repository clean or known-dirty with expected files,
- policy and workflow in sync,
- required checks passing,
- apply mode re-enabled only after successful dry-run.
