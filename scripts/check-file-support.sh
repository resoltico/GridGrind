#!/usr/bin/env bash
# Shared portable file helpers for root-check progress monitoring.

file_size_bytes() {
    local file_path=$1
    local size=''

    if [[ ! -f "${file_path}" ]]; then
        printf '0'
        return 0
    fi

    if size="$(stat -f '%z' "${file_path}" 2>/dev/null)"; then
        printf '%s' "${size}"
        return 0
    fi
    if size="$(stat -c '%s' "${file_path}" 2>/dev/null)"; then
        printf '%s' "${size}"
        return 0
    fi

    size="$(wc -c < "${file_path}" | tr -d '[:space:]')"
    printf '%s' "${size:-0}"
}
