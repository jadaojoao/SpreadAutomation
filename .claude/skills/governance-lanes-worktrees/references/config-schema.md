# Governance Config Schema (V2)

This schema defines all configurable knobs for a portable, discovery-first governance bootstrap.

Versioning note:
- v2 remains GitHub-first.
- Structure is adapter-ready for future GitLab/Bitbucket support.

## Minimal Required Root Fields
- `version`
- `platform`
- `locale`
- `discovery`
- `autonomy`
- `lanes`
- `labels`
- `risk_model`
- `issue_contract`
- `pull_request_contract`
- `worktree`
- `child_tasks`
- `path_policy`
- `guardrails`
- `git_configuration`
- `bootstrap`
- `artifacts`

## Root Fields

### `version`
- type: string
- required: yes
- default: `2.0.0`

### `platform`
- type: string
- required: yes
- allowed: `github`
- default: `github`

### `locale`
- type: string
- required: yes
- allowed: `en`, `pt-BR`
- default: `en`

### `discovery`
- type: object
- required: yes
- fields:
  - `required`: bool (default: true)
  - `scan_github_issues`: bool
  - `scan_github_prs`: bool
  - `scan_github_artifacts`: bool
  - `scan_docs`: bool
  - `scan_repo_tree`: bool
  - `required_paths`: array of path globs
  - `block_apply_if_incomplete`: bool

### `autonomy`
- type: object
- required: yes
- fields:
  - `mode`: `conservative` | `aggressive` | `explain-only`
  - `require_confirm_for_apply`: bool
  - `require_confirm_for_merge`: bool
  - `require_confirm_for_force_remove`: bool
  - `allow_auto_branch_delete`: bool

## Governance Model

### `lanes`
- type: array
- required: yes
- min items: 1
- each item:
  - `id`: string (unique, machine-safe)
  - `display_name`: string
  - `description`: string
  - `ownership_patterns`: array of glob
  - `allow_shared_governance`: bool
  - `allow_critical_groups`: array of group ids

### `labels`
- type: object
- required: yes
- fields:
  - `kind_options`: array
  - `status_options`: array
  - `priority_options`: array
  - `area_options`: array
  - `risk_options`: array
  - `lane_prefix`: string (default: `lane:`)
  - `status_blocked`: string
  - `status_awaiting_consumption`: string

### `risk_model`
- type: object
- required: yes
- fields:
  - `ordered_levels`: array from low to high
  - `minimum_for_critical`: map group id -> risk level
  - `default_for_normal_changes`: risk level

## Contracts

### `issue_contract`
- type: object
- required: yes
- fields:
  - `required_labels`: array
  - `required_sections`: array
  - `section_aliases`: object map
  - `closes_keywords`: array (example: `Closes`, `Fixes`, `Resolves`)
  - `linked_issue_pattern`: regex string
  - `require_write_set`: bool
  - `require_workspace_path`: bool
  - `workspace_path_pattern`: string
  - `allow_epic_close_by_pr`: bool

### `pull_request_contract`
- type: object
- required: yes
- fields:
  - `required_sections`: array
  - `section_aliases`: object map
  - `require_compatibility_for_groups`: array of critical groups
  - `compatibility_values`: array (`additive-only`, `breaking-approved`, custom)
  - `default_merge_method`: `squash` | `merge` | `rebase`
  - `required_checks_mode`: `required` | `all` | `none`
  - `poll_seconds`: integer
  - `timeout_seconds`: integer
  - `requires_issue_closure_line`: bool

### `worktree`
- type: object
- required: yes
- fields:
  - `root`: string
  - `branch_pattern`: string (must include issue number and slug tokens)
  - `worktree_pattern`: string (must include lane, issue number, slug)
  - `base_branch_mode`: `auto-detect` | `fixed`
  - `base_branch`: string (used when fixed or fallback)
  - `remove_requires_merged`: bool

### `child_tasks`
- type: object
- required: yes
- fields:
  - `enabled`: bool
  - `require_parent_reference`: bool
  - `require_requester_lane`: bool
  - `require_consumption_criteria`: bool
  - `parent_required_status`: array
  - `requester_must_match_parent_lane`: bool
  - `require_parent_lists_child`: bool

## Policy and Enforcement

### `path_policy`
- type: object
- required: yes
- fields:
  - `policy_file`: path
  - `lane_allowlists`: object map lane -> patterns
  - `critical_groups`: array
    - each group:
      - `name`
      - `patterns`
      - `owner_lane`
      - `allowed_lanes`
      - `minimum_risk`
      - `require_draft`
      - `require_compatibility`
  - `disallowed_domain_mixes`: array with `left` and `right`

### `guardrails`
- type: object
- required: yes
- fields:
  - `workflow_file`: path
  - `fail_on_unclassified_paths`: bool
  - `enforce_one_pr_per_task`: bool
  - `enforce_branch_issue_match`: bool
  - `enforce_label_uniqueness`: bool
  - `enforce_write_set_coverage`: bool
  - `enforce_child_task_rules`: bool

## Git Configuration

### `git_configuration`
- type: object
- required: yes
- fields:
  - `enabled`: bool
  - `mode`: `conservative` | `strict` | `disabled`
  - `apply_scopes`:
    - `global`: bool
    - `repo_local`: bool
    - `worktree_local`: bool
  - `repo_local_defaults`: key/value
  - `global_defaults`: key/value
  - `profiles`: object map profile -> key/value
  - `profile_by_lane`: lane -> profile
  - `protected_settings`: array
  - `os_defaults`: windows/linux/macos maps
  - `validation_rules`: array of allow/deny rules
  - `require_confirmation_for`: array (`global`, `protected`, `strict_profile`)

## Bootstrap Behavior

### `bootstrap`
- type: object
- required: yes
- fields:
  - `default_mode`: `dry-run` | `apply`
  - `overwrite_policy`: `preserve` | `merge` | `force`
  - `backup_dir`: path
  - `create_backup_manifest`: bool
  - `rollback_on_failure`: bool
  - `preflight_required_tools`: array
  - `shell_support`: array (`powershell`, `bash`)

### `artifacts`
- type: object
- required: yes
- fields:
  - issue templates, PR template, workflow, path policy, docs, runbook, rollback docs
  - helper scripts for worktree create/status/remove and PR completion

## Localization

### `localization`
- type: object
- required: yes
- fields:
  - `default_locale`: string
  - `supported_locales`: array
  - `headings`: locale map for issue/PR section names
  - `workflow_errors`: locale map for guardrail failure text
  - `template_labels`: locale map for form labels/descriptions

## Validation Rules
- lane ids must be unique.
- label namespaces must not overlap.
- risk levels in `risk_model` must cover all critical groups.
- branch/worktree patterns must include required tokens.
- critical groups must define minimum risk and allowed lanes.
- no critical group overlap for the same path.
- disallowed domain mixes must have non-empty left/right sets.
- unclassified changed files must fail when strictness requires classification.
- protected Git settings cannot be auto-applied in conservative mode.
- global Git changes require explicit confirmation in conservative mode.
- locale aliases must map all required section names.

## Recommended Presets
- `preset_three_lanes`: frontend, backend, ops-quality
- `preset_two_lanes`: product, platform
- `preset_minimal`: core, docs

## Conservative Git Defaults
Repo-local defaults:
- `fetch.prune: true`
- `pull.rebase: true`
- `rebase.autostash: true`
- `core.safecrlf: true`
- `push.default: simple`

Global defaults (confirmation required):
- `init.defaultBranch: <detected-or-configured-base-branch>`
- `push.default: simple`

Protected settings (never automatic):
- `user.name`
- `user.email`
- `credential.helper`
- `commit.gpgsign`
- `core.sshCommand`
- `merge.ff`
