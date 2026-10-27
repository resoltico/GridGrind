#!/usr/bin/env bash
# Deterministic checks and delivery artifacts share one Gradle task graph.
set -euo pipefail
source "$(cd -P -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)/scripts/lib/common.sh"
require_java_26
exec "${GRIDGRIND_ROOT}/gradlew" --project-dir "${GRIDGRIND_ROOT}" \
    check coverage :cli:shadowDistZip :cli:shadowDistTar "$@"
