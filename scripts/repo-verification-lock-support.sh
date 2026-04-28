#!/usr/bin/env bash
# Shared single-run lock support for top-level GridGrind verification entrypoints.
#
# Expects the caller to define:
#   lock_dir
#   pid_file
#
# Optional caller overrides:
#   lock_scope_name
#   lock_scope_advice

lock_initialization_attempts=40
lock_initialization_sleep_seconds=0.05
lock_scope_name="${lock_scope_name:-GridGrind verification command}"
lock_scope_advice="${lock_scope_advice:-run one GridGrind verification command at a time}"
lock_is_reentrant=false
lock_owned_by_current_process=false

write_lock_pid() {
  printf '%s\n' "$$" > "${pid_file}"
}

read_lock_pid() {
  if [[ -f "${pid_file}" ]]; then
    <"${pid_file}" tr -d '[:space:]'
  fi
}

lock_pid_description() {
  local lock_pid=${1:-}
  local command_line=''
  [[ "${lock_pid}" =~ ^[0-9]+$ ]] || return 0
  command_line="$(ps -o command= -p "${lock_pid}" 2>/dev/null | head -1 || true)"
  if [[ -n "${command_line}" ]]; then
    printf '%s' "${command_line}" | sed -E 's/[[:space:]]+/ /g' | cut -c1-160
  fi
}

current_process_descends_from_pid() {
  local target_pid=${1:-}
  local candidate_pid=$$
  local parent_pid=''

  [[ "${target_pid}" =~ ^[0-9]+$ ]] || return 1
  while [[ "${candidate_pid}" =~ ^[0-9]+$ ]] && (( candidate_pid > 1 )); do
    if [[ "${candidate_pid}" == "${target_pid}" ]]; then
      return 0
    fi
    parent_pid="$(ps -o ppid= -p "${candidate_pid}" 2>/dev/null | tr -d '[:space:]' || true)"
    [[ -n "${parent_pid}" ]] || return 1
    candidate_pid="${parent_pid}"
  done
  return 1
}

report_lock_conflict() {
  local lock_pid=${1:-}
  local lock_description=''
  if [[ -n "${lock_pid}" ]]; then
    lock_description="$(lock_pid_description "${lock_pid}")"
    if [[ -n "${lock_description}" ]]; then
      printf 'another %s is already running with PID %s (%s); %s\n' \
        "${lock_scope_name}" \
        "${lock_pid}" \
        "${lock_description}" \
        "${lock_scope_advice}" >&2
      return
    fi
    printf 'another %s is already running with PID %s; %s\n' \
      "${lock_scope_name}" \
      "${lock_pid}" \
      "${lock_scope_advice}" >&2
    return
  fi
  printf 'another %s is already starting; %s\n' \
    "${lock_scope_name}" \
    "${lock_scope_advice}" >&2
}

reclaim_stale_lock() {
  rm -rf "${lock_dir}"
  mkdir -p "${lock_dir}"
  write_lock_pid
  lock_is_reentrant=false
  lock_owned_by_current_process=true
}

acquire_lock() {
  local attempt=0
  local lock_pid=''

  mkdir -p "$(dirname -- "${lock_dir}")"
  if mkdir "${lock_dir}" 2>/dev/null; then
    write_lock_pid
    lock_is_reentrant=false
    lock_owned_by_current_process=true
    return 0
  fi

  for ((attempt = 0; attempt < lock_initialization_attempts; attempt++)); do
    if [[ ! -d "${lock_dir}" ]]; then
      if mkdir "${lock_dir}" 2>/dev/null; then
        write_lock_pid
        lock_is_reentrant=false
        lock_owned_by_current_process=true
        return 0
      fi
    fi

    lock_pid="$(read_lock_pid || true)"
    if [[ -n "${lock_pid}" ]]; then
      if kill -0 "${lock_pid}" 2>/dev/null; then
        if current_process_descends_from_pid "${lock_pid}"; then
          lock_is_reentrant=true
          lock_owned_by_current_process=false
          return 0
        fi
        report_lock_conflict "${lock_pid}"
        exit 1
      fi
      reclaim_stale_lock
      return 0
    fi

    sleep "${lock_initialization_sleep_seconds}"
  done

  if [[ -d "${lock_dir}" ]] && [[ ! -f "${pid_file}" ]]; then
    reclaim_stale_lock
    return 0
  fi

  lock_pid="$(read_lock_pid || true)"
  if [[ -n "${lock_pid}" ]] && kill -0 "${lock_pid}" 2>/dev/null; then
    if current_process_descends_from_pid "${lock_pid}"; then
      lock_is_reentrant=true
      lock_owned_by_current_process=false
      return 0
    fi
    report_lock_conflict "${lock_pid}"
    exit 1
  fi

  report_lock_conflict ''
  exit 1
}

cleanup_lock() {
  if [[ "${lock_is_reentrant}" == true ]]; then
    return 0
  fi
  if [[ "${lock_owned_by_current_process}" == true ]]; then
    rm -rf "${lock_dir}"
  fi
}
