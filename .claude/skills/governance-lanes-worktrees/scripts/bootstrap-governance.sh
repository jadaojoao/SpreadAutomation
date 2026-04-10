#!/usr/bin/env bash
set -euo pipefail

MODE="dry-run"
OVERWRITE_POLICY="preserve"
CONFIGURE_GIT="false"
GIT_PROFILE="conservative"
APPLY_GLOBAL_GIT_CONFIG="false"
BASE_BRANCH=""
LOCALE="en"
REPO_PATH=""

die() {
	echo "ERROR: $*" >&2
	exit 1
}

ensure_dir() {
	local dir_path="$1"
	mkdir -p "$dir_path"
}

require_cmd() {
	local name="$1"
	if ! command -v "$name" >/dev/null 2>&1; then
		die "Required command not found in PATH: $name"
	fi
}

show_help() {
	cat <<EOF
Usage: bootstrap-governance.sh --repo <path> [options]

Options:
	--mode <dry-run|apply>                  Default: dry-run
	--overwrite-policy <preserve|merge|force>  Default: preserve
	--configure-git                         Apply Git profile settings
	--git-profile <conservative|strict|disabled> Default: conservative
	--apply-global-git-config               Apply global Git defaults (requires --configure-git)
	--base-branch <name>                    Override auto-detected base branch
	--locale <en|pt-BR>                     Default: en
	--help                                  Show this message

Examples:
	./scripts/bootstrap-governance.sh --repo . --mode dry-run
	./scripts/bootstrap-governance.sh --repo . --mode apply --overwrite-policy force --configure-git
EOF
}

while [[ $# -gt 0 ]]; do
	case "$1" in
		--repo)
			REPO_PATH="${2:-}"
			shift 2
			;;
		--mode)
			MODE="${2:-}"
			shift 2
			;;
		--overwrite-policy)
			OVERWRITE_POLICY="${2:-}"
			shift 2
			;;
		--configure-git)
			CONFIGURE_GIT="true"
			shift
			;;
		--git-profile)
			GIT_PROFILE="${2:-}"
			shift 2
			;;
		--apply-global-git-config)
			APPLY_GLOBAL_GIT_CONFIG="true"
			shift
			;;
		--base-branch)
			BASE_BRANCH="${2:-}"
			shift 2
			;;
		--locale)
			LOCALE="${2:-}"
			shift 2
			;;
		--help)
			show_help
			exit 0
			;;
		*)
			die "Unknown argument: $1"
			;;
	esac
done

[[ -n "$REPO_PATH" ]] || die "--repo is required"
[[ "$MODE" == "dry-run" || "$MODE" == "apply" ]] || die "Invalid --mode: $MODE"
[[ "$OVERWRITE_POLICY" == "preserve" || "$OVERWRITE_POLICY" == "merge" || "$OVERWRITE_POLICY" == "force" ]] || die "Invalid --overwrite-policy: $OVERWRITE_POLICY"
[[ "$GIT_PROFILE" == "conservative" || "$GIT_PROFILE" == "strict" || "$GIT_PROFILE" == "disabled" ]] || die "Invalid --git-profile: $GIT_PROFILE"
[[ "$LOCALE" == "en" || "$LOCALE" == "pt-BR" ]] || die "Invalid --locale: $LOCALE"

require_cmd git

REPO_ROOT="$(cd "$REPO_PATH" && pwd)"
if ! git -C "$REPO_ROOT" rev-parse --is-inside-work-tree >/dev/null 2>&1; then
	die "RepoPath is not a Git repository: $REPO_ROOT"
fi

if [[ -n "$(git -C "$REPO_ROOT" status --porcelain)" ]]; then
	echo "WARN: Repository has uncommitted changes. Continue with caution."
fi

if [[ -n "$BASE_BRANCH" ]]; then
	DETECTED_BASE_BRANCH="$BASE_BRANCH"
else
	DETECTED_BASE_BRANCH="$(git -C "$REPO_ROOT" symbolic-ref refs/remotes/origin/HEAD 2>/dev/null | awk -F'/' '{print $NF}' || true)"
	if [[ -z "$DETECTED_BASE_BRANCH" ]]; then
		DETECTED_BASE_BRANCH="master"
	fi
fi

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SKILL_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
ASSETS_ROOT="$SKILL_ROOT/assets"

TIMESTAMP="$(date +%Y%m%d-%H%M%S)"
BACKUP_ROOT="$REPO_ROOT/.governance-bootstrap/backups/$TIMESTAMP"
MERGE_ROOT="$REPO_ROOT/.governance-bootstrap/merge-candidates/$TIMESTAMP"
MANIFEST_PATH="$BACKUP_ROOT/manifest.txt"

MAPPINGS=(
	"AGENTS.template.md|AGENTS.md"
	"parallel-lanes.template.md|docs/governance/parallel-lanes.md"
	"operators-runbook.template.md|docs/governance/operators-runbook.md"
	"rollback-recovery.template.md|docs/governance/rollback-recovery.md"
	"task.issue.template.yml|.github/ISSUE_TEMPLATE/task.yml"
	"epic.issue.template.yml|.github/ISSUE_TEMPLATE/epic.yml"
	"issue-config.yml|.github/ISSUE_TEMPLATE/config.yml"
	"pull_request.template.md|.github/PULL_REQUEST_TEMPLATE.md"
	"path-policy.template.json|.github/guardrails/path-policy.json"
	"pr-issue-guardrails.workflow.yml|.github/workflows/pr-issue-guardrails.yml"
	"governance.config.template.yaml|.github/governance.config.yaml"
)

for pair in "${MAPPINGS[@]}"; do
	source_file="${pair%%|*}"
	[[ -f "$ASSETS_ROOT/$source_file" ]] || die "Missing asset: $ASSETS_ROOT/$source_file"
done

action_for_target() {
	local target_file="$1"
	if [[ ! -e "$target_file" ]]; then
		echo "create"
		return
	fi

	case "$OVERWRITE_POLICY" in
		preserve) echo "skip" ;;
		merge) echo "merge-candidate" ;;
		force) echo "overwrite" ;;
	esac
}

CREATED=()
OVERWRITTEN=()
MERGE_CANDIDATES=()
GIT_APPLIED=()

rollback_on_error() {
	if [[ "$MODE" != "apply" ]]; then
		return
	fi

	echo "WARN: Bootstrap failed. Attempting rollback."
	for row in "${OVERWRITTEN[@]}"; do
		target_rel="${row%%|*}"
		backup_abs="${row##*|}"
		target_abs="$REPO_ROOT/$target_rel"
		ensure_dir "$(dirname "$target_abs")"
		cp -f "$backup_abs" "$target_abs"
	done

	for target_rel in "${CREATED[@]}"; do
		target_abs="$REPO_ROOT/$target_rel"
		[[ -e "$target_abs" ]] && rm -f "$target_abs"
	done

	for target_rel in "${MERGE_CANDIDATES[@]}"; do
		candidate_abs="$REPO_ROOT/$target_rel"
		[[ -e "$candidate_abs" ]] && rm -f "$candidate_abs"
	done
}

trap rollback_on_error ERR

echo "Governance bootstrap mode: $MODE"
echo "Repository: $REPO_ROOT"
echo "Locale: $LOCALE"
echo "Overwrite policy: $OVERWRITE_POLICY"
echo "Detected base branch: $DETECTED_BASE_BRANCH"
echo "Git profile requested: $GIT_PROFILE"
echo

printf "%-16s %-36s %s\n" "Action" "Source" "Target"
printf "%-16s %-36s %s\n" "------" "------" "------"
for pair in "${MAPPINGS[@]}"; do
	source_file="${pair%%|*}"
	target_rel="${pair##*|}"
	target_abs="$REPO_ROOT/$target_rel"
	action="$(action_for_target "$target_abs")"
	printf "%-16s %-36s %s\n" "$action" "$source_file" "$target_rel"
done

if [[ "$MODE" == "dry-run" ]]; then
	echo
	echo "Dry-run complete. No files were changed."
	exit 0
fi

ensure_dir "$BACKUP_ROOT"
ensure_dir "$MERGE_ROOT"

for pair in "${MAPPINGS[@]}"; do
	source_file="${pair%%|*}"
	target_rel="${pair##*|}"
	source_abs="$ASSETS_ROOT/$source_file"
	target_abs="$REPO_ROOT/$target_rel"
	action="$(action_for_target "$target_abs")"

	case "$action" in
		skip)
			continue
			;;
		merge-candidate)
			candidate_abs="$MERGE_ROOT/$target_rel"
			ensure_dir "$(dirname "$candidate_abs")"
			cp -f "$source_abs" "$candidate_abs"
			MERGE_CANDIDATES+=("${candidate_abs#$REPO_ROOT/}")
			;;
		overwrite)
			backup_abs="$BACKUP_ROOT/$target_rel"
			ensure_dir "$(dirname "$backup_abs")"
			cp -f "$target_abs" "$backup_abs"
			OVERWRITTEN+=("$target_rel|$backup_abs")
			ensure_dir "$(dirname "$target_abs")"
			cp -f "$source_abs" "$target_abs"
			;;
		create)
			CREATED+=("$target_rel")
			ensure_dir "$(dirname "$target_abs")"
			cp -f "$source_abs" "$target_abs"
			;;
	esac
done

if [[ "$CONFIGURE_GIT" == "true" && "$GIT_PROFILE" != "disabled" ]]; then
	git -C "$REPO_ROOT" config fetch.prune true
	GIT_APPLIED+=("local fetch.prune=true")
	git -C "$REPO_ROOT" config pull.rebase true
	GIT_APPLIED+=("local pull.rebase=true")
	git -C "$REPO_ROOT" config rebase.autostash true
	GIT_APPLIED+=("local rebase.autostash=true")
	git -C "$REPO_ROOT" config push.default simple
	GIT_APPLIED+=("local push.default=simple")
	git -C "$REPO_ROOT" config core.safecrlf true
	GIT_APPLIED+=("local core.safecrlf=true")

	if [[ "$(uname -s)" == "Darwin" || "$(uname -s)" == "Linux" ]]; then
		git -C "$REPO_ROOT" config core.autocrlf input
		GIT_APPLIED+=("local core.autocrlf=input")
	else
		git -C "$REPO_ROOT" config core.autocrlf true
		GIT_APPLIED+=("local core.autocrlf=true")
	fi

	if [[ "$GIT_PROFILE" == "strict" ]]; then
		git -C "$REPO_ROOT" config merge.ff false
		GIT_APPLIED+=("local merge.ff=false")
	fi

	if [[ "$APPLY_GLOBAL_GIT_CONFIG" == "true" ]]; then
		git config --global init.defaultBranch "$DETECTED_BASE_BRANCH"
		GIT_APPLIED+=("global init.defaultBranch=$DETECTED_BASE_BRANCH")
		git config --global push.default simple
		GIT_APPLIED+=("global push.default=simple")
	fi
fi

{
	echo "timestamp: $TIMESTAMP"
	echo "mode: $MODE"
	echo "locale: $LOCALE"
	echo "overwrite_policy: $OVERWRITE_POLICY"
	echo "repo_root: $REPO_ROOT"
	echo "base_branch: $DETECTED_BASE_BRANCH"
	echo "created:"
	for row in "${CREATED[@]}"; do
		echo "  - $row"
	done
	echo "overwritten:"
	for row in "${OVERWRITTEN[@]}"; do
		echo "  - ${row%%|*}"
	done
	echo "merge_candidates:"
	for row in "${MERGE_CANDIDATES[@]}"; do
		echo "  - $row"
	done
	echo "git_applied:"
	for row in "${GIT_APPLIED[@]}"; do
		echo "  - $row"
	done
} > "$MANIFEST_PATH"

echo
echo "Apply mode completed."
echo "Backup snapshot: $BACKUP_ROOT"
echo "Backup manifest: $MANIFEST_PATH"

if [[ ${#MERGE_CANDIDATES[@]} -gt 0 ]]; then
	echo
	echo "Merge candidates generated under: $MERGE_ROOT"
fi

if [[ ${#GIT_APPLIED[@]} -gt 0 ]]; then
	echo
	echo "Git settings applied:"
	for row in "${GIT_APPLIED[@]}"; do
		echo "- $row"
	done
fi

echo "Review generated files and adjust ownership/labels/policy before opening PRs."
