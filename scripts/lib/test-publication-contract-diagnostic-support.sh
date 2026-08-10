#!/usr/bin/env bash
# Independent-JVM checks for deterministic packaged CLI diagnostic payloads.

verify_packaged_diagnostic_byte_stability() {
    local packaged_cli_jar=$1
    local diagnostic_test_root=$2
    local diagnostic_request_path="${diagnostic_test_root}/deterministic-diagnostic-request.json"

    cat > "${diagnostic_request_path}" <<'JSON'
{
  "protocolVersion": "V2",
  "source": { "type": "NEW" },
  "persistence": { "type": "OVERWRITE" },
  "execution": {
    "mode": { "type": "STREAMING_WRITE" },
    "journal": { "level": "SUMMARY" },
    "calculation": {
      "strategy": { "type": "EVALUATE_ALL" },
      "markRecalculateOnOpen": false
    }
  },
  "steps": []
}
JSON

    capture_packaged_diagnostic() {
        local expected_exit_code=$1
        local output_path=$2
        local error_path=$3
        shift 3
        local actual_exit_code

        set +e
        java -jar "${packaged_cli_jar}" "$@" --execution-root "${diagnostic_test_root}" \
            < "${diagnostic_request_path}" > "${output_path}" 2> "${error_path}"
        actual_exit_code=$?
        set -e
        [[ ${actual_exit_code} -eq ${expected_exit_code} ]] || die \
            "packaged CLI diagnostic exited ${actual_exit_code}, expected ${expected_exit_code}"
        [[ -s "${output_path}" ]] || die "packaged CLI diagnostic did not write its primary payload"
        [[ ! -s "${error_path}" ]] || die "packaged CLI diagnostic wrote unexpected stderr"
    }

    capture_packaged_diagnostic 2 \
        "${diagnostic_test_root}/command-error-first.json" \
        "${diagnostic_test_root}/command-error-first.stderr"
    capture_packaged_diagnostic 2 \
        "${diagnostic_test_root}/command-error-second.json" \
        "${diagnostic_test_root}/command-error-second.stderr"
    cmp -s \
        "${diagnostic_test_root}/command-error-first.json" \
        "${diagnostic_test_root}/command-error-second.json" || die \
        "packaged CommandError diagnostics are not byte-stable across independent JVMs"
    grep -Fq 'OVERWRITE persistence requires an EXISTING source' \
        "${diagnostic_test_root}/command-error-first.json" || die \
        "packaged CommandError diagnostic did not retain the persistence validation problem"
    grep -Fq 'STREAMING_WRITE' "${diagnostic_test_root}/command-error-first.json" || die \
        "packaged CommandError diagnostic did not retain the execution-mode validation problem"

    capture_packaged_diagnostic 1 \
        "${diagnostic_test_root}/doctor-report-first.json" \
        "${diagnostic_test_root}/doctor-report-first.stderr" \
        --doctor-request
    capture_packaged_diagnostic 1 \
        "${diagnostic_test_root}/doctor-report-second.json" \
        "${diagnostic_test_root}/doctor-report-second.stderr" \
        --doctor-request
    cmp -s \
        "${diagnostic_test_root}/doctor-report-first.json" \
        "${diagnostic_test_root}/doctor-report-second.json" || die \
        "packaged doctor diagnostics are not byte-stable across independent JVMs"
    grep -Fq '"valid":false' "${diagnostic_test_root}/doctor-report-first.json" || die \
        "packaged doctor diagnostic no longer reports validation findings in its own payload"
}
