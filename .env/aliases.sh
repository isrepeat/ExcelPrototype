_codex_env_dir="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

codexlog() {
  local log_dir
  log_dir="${_codex_env_dir}/Codex/logs"
  mkdir -p "$log_dir"
  script -q -f -c "codex" "${log_dir}/codex-$(date +%F_%H-%M).log"
}
