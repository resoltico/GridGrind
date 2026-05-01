#!/usr/bin/env bash
# Refresh checkout-rooted example fixtures and the generated package-security workbook asset.

set -euo pipefail

resolve_script_dir() {
    local source_path="${BASH_SOURCE[0]}"
    while [[ -h "${source_path}" ]]; do
        local source_dir
        source_dir="$(cd -P -- "$(dirname -- "${source_path}")" && pwd)"
        source_path="$(readlink "${source_path}")"
        if [[ "${source_path}" != /* ]]; then
            source_path="${source_dir}/${source_path}"
        fi
    done
    cd -P -- "$(dirname -- "${source_path}")" && pwd
}

readonly script_dir="$(resolve_script_dir)"
readonly repo_root="$(cd "${script_dir}/.." && pwd)"
readonly gradlew="${repo_root}/gradlew"
readonly jar_path="${repo_root}/cli/build/libs/gridgrind.jar"
readonly scratch_root="${repo_root}/tmp/sync-generated-examples"
readonly request_path="${scratch_root}/package-security-asset-request.json"
readonly response_path="${scratch_root}/package-security-asset-response.json"
readonly verify_response_path="${scratch_root}/package-security-example-response.json"
readonly asset_directory="${repo_root}/examples/package-security-assets"
readonly asset_path="${asset_directory}/gridgrind-package-security.xlsx"

mkdir -p "${scratch_root}" "${asset_directory}"

"${gradlew}" \
    :contract:writeRepositoryExamples \
    :cli:shadowJar \
    "$@"

cat > "${request_path}" <<EOF
{
  "protocolVersion": "V1",
  "planId": "generate-package-security-asset",
  "source": {
    "type": "NEW"
  },
  "persistence": {
    "type": "SAVE_AS",
    "path": "${asset_path}",
    "security": {
      "encryption": {
        "password": "GridGrind-2026",
        "mode": "AGILE"
      }
    }
  },
  "execution": {
    "mode": {
      "readMode": "FULL_XSSF",
      "writeMode": "FULL_XSSF"
    },
    "journal": {
      "level": "NORMAL"
    },
    "calculation": {
      "strategy": {
        "type": "DO_NOT_CALCULATE"
      },
      "markRecalculateOnOpen": false
    }
  },
  "formulaEnvironment": {
    "externalWorkbooks": [],
    "missingWorkbookPolicy": "ERROR",
    "udfToolpacks": []
  },
  "steps": [
    {
      "stepId": "ensure-secure",
      "target": {
        "type": "SHEET_BY_NAME",
        "name": "Secure"
      },
      "action": {
        "type": "ENSURE_SHEET"
      }
    },
    {
      "stepId": "seed-secure-cells",
      "target": {
        "type": "RANGE_BY_RANGE",
        "sheetName": "Secure",
        "range": "A1:B3"
      },
      "action": {
        "type": "SET_RANGE",
        "rows": [
          [
            {
              "type": "TEXT",
              "source": {
                "type": "INLINE",
                "text": "Field"
              }
            },
            {
              "type": "TEXT",
              "source": {
                "type": "INLINE",
                "text": "Value"
              }
            }
          ],
          [
            {
              "type": "TEXT",
              "source": {
                "type": "INLINE",
                "text": "Owner"
              }
            },
            {
              "type": "TEXT",
              "source": {
                "type": "INLINE",
                "text": "GridGrind"
              }
            }
          ],
          [
            {
              "type": "TEXT",
              "source": {
                "type": "INLINE",
                "text": "Status"
              }
            },
            {
              "type": "TEXT",
              "source": {
                "type": "INLINE",
                "text": "Encrypted"
              }
            }
          ]
        ]
      }
    }
  ]
}
EOF

java -jar "${jar_path}" \
    --request "${request_path}" \
    --response "${response_path}"

java -jar "${jar_path}" \
    --request "${repo_root}/examples/package-security-inspect-request.json" \
    --response "${verify_response_path}"

printf 'Refreshed example fixtures and package-security asset under %s\n' "${repo_root}/examples"
