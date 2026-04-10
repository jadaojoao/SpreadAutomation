# Rollback and Recovery

## If Bootstrap Apply Fails
1. Locate backup snapshot at .governance-bootstrap/backups/<timestamp>/.
2. Open backup manifest and identify overwritten and created targets.
3. Restore overwritten files from the snapshot.
4. Remove files created by failed run when needed.
5. Re-run bootstrap in dry-run mode to inspect conflicts.

## If Guardrails Block Unexpectedly
1. Read failing workflow message.
2. Fix issue labels/metadata or path policy mapping.
3. Re-run checks by updating PR.

## If Locale or Parsing Drifts
1. Confirm section aliases in governance config.
2. Regenerate workflow and templates from the same config.
3. Re-run dry-run validation before apply.

## If Templates Are Wrong
1. Restore affected templates from backup.
2. Reapply updated templates only after dry-run review.

## If Git Settings Need Revert
1. Read current config with origin:
   git config --list --show-origin
2. Reapply expected values per governance config.
3. Validate fetch and PR operations.

## Safe Mode Recommendation
- Use dry-run mode.
- Keep conservative autonomy.
- Require explicit confirmation for merge and force remove.

## Post-Recovery Checklist
- Repository status is expected and understood.
- Governance config, policy, and workflow are synchronized.
- Guardrails pass on next PR synchronization.
