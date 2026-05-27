#!/usr/bin/env bash
# Refresh checkout-rooted example fixtures plus generated workbook assets used by discovery flows.

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
readonly cli_shadow_jar_support="${repo_root}/scripts/lib/cli-shadow-jar-support.sh"
readonly scratch_root="${repo_root}/tmp/sync-generated-examples"
readonly request_path="${scratch_root}/package-security-asset-request.json"
readonly response_path="${scratch_root}/package-security-asset-response.json"
readonly verify_response_path="${scratch_root}/package-security-example-response.json"
readonly asset_directory="${repo_root}/examples/package-security-assets"
readonly asset_path="${asset_directory}/gridgrind-package-security.xlsx"
readonly task_asset_directory="${repo_root}/examples/task-starter-assets"
readonly task_asset_path="${task_asset_directory}/workbook-ops-source.xlsx"
readonly task_request_path="${scratch_root}/task-starter-workbook-request.json"
readonly task_response_path="${scratch_root}/task-starter-workbook-response.json"

# shellcheck source=/dev/null
source "${cli_shadow_jar_support}"

mkdir -p "${scratch_root}" "${asset_directory}" "${task_asset_directory}"

"${gradlew}" \
    :cli:writeRepositoryExamples \
    :cli:shadowJar \
    "$@"

readonly jar_path="$(ensure_cli_shadow_jar "${repo_root}")"

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
    "mode": {"type": "FULL_XSSF"},
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

cat > "${task_request_path}" <<EOF
{
  "protocolVersion": "V1",
  "planId": "generate-task-starter-workbook-asset",
  "source": {
    "type": "NEW"
  },
  "persistence": {
    "type": "SAVE_AS",
    "path": "${task_asset_path}"
  },
  "execution": {
    "mode": {"type": "FULL_XSSF"},
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
      "stepId": "ensure-template",
      "target": {
        "type": "SHEET_BY_NAME",
        "name": "Template"
      },
      "action": {
        "type": "ENSURE_SHEET"
      }
    },
    {
      "stepId": "seed-template-range",
      "target": {
        "type": "RANGE_BY_RANGE",
        "sheetName": "Template",
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
                "text": "Owner"
              }
            },
            {
              "type": "TEXT",
              "source": {
                "type": "INLINE",
                "text": "Status"
              }
            }
          ],
          [
            {
              "type": "TEXT",
              "source": {
                "type": "INLINE",
                "text": "Ada"
              }
            },
            {
              "type": "TEXT",
              "source": {
                "type": "INLINE",
                "text": "Ready"
              }
            }
          ],
          [
            {
              "type": "TEXT",
              "source": {
                "type": "INLINE",
                "text": "Lin"
              }
            },
            {
              "type": "TEXT",
              "source": {
                "type": "INLINE",
                "text": "Review"
              }
            }
          ]
        ]
      }
    },
    {
      "stepId": "comment-template-header",
      "target": {
        "type": "CELL_BY_ADDRESS",
        "sheetName": "Template",
        "address": "A1"
      },
      "action": {
        "type": "SET_COMMENT",
        "comment": {
          "text": {
            "type": "INLINE",
            "text": "Template owner column"
          },
          "author": "GridGrind",
          "visible": false
        }
      }
    },
    {
      "stepId": "set-template-signature-line",
      "target": {
        "type": "SHEET_BY_NAME",
        "name": "Template"
      },
      "action": {
        "type": "SET_SIGNATURE_LINE",
        "signatureLine": {
          "name": "TemplateSignature",
          "anchor": {
            "type": "TWO_CELL",
            "from": {
              "columnIndex": 3,
              "rowIndex": 1,
              "dx": 0,
              "dy": 0
            },
            "to": {
              "columnIndex": 6,
              "rowIndex": 5,
              "dx": 0,
              "dy": 0
            },
            "behavior": "MOVE_AND_RESIZE"
          },
          "allowComments": false,
          "signingInstructions": "Review the workbook before signing.",
          "suggestedSigner": "Ada Lovelace",
          "suggestedSigner2": "Operations",
          "suggestedSignerEmail": "ada@example.com",
          "invalidStamp": "invalid",
          "plainSignature": {
            "format": "PNG",
            "source": {
              "type": "INLINE_BASE64",
              "base64Data": "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII="
            }
          }
        }
      }
    },
    {
      "stepId": "ensure-summary",
      "target": {
        "type": "SHEET_BY_NAME",
        "name": "Summary"
      },
      "action": {
        "type": "ENSURE_SHEET"
      }
    },
    {
      "stepId": "set-summary-label",
      "target": {
        "type": "CELL_BY_ADDRESS",
        "sheetName": "Summary",
        "address": "A1"
      },
      "action": {
        "type": "SET_CELL",
        "value": {
          "type": "TEXT",
          "source": {
            "type": "INLINE",
            "text": "Template status"
          }
        }
      }
    },
    {
      "stepId": "set-summary-formula",
      "target": {
        "type": "CELL_BY_ADDRESS",
        "sheetName": "Summary",
        "address": "B1"
      },
      "action": {
        "type": "SET_CELL",
        "value": {
          "type": "FORMULA",
          "source": {
            "type": "INLINE",
            "text": "'Template'!B2"
          }
        }
      }
    }
  ]
}
EOF

java -jar "${jar_path}" \
    --request "${task_request_path}" \
    --response "${task_response_path}"

java -jar "${jar_path}" \
    --request "${repo_root}/examples/package-security-inspect-request.json" \
    --response "${verify_response_path}"

printf 'Refreshed example fixtures and workbook assets under %s\n' "${repo_root}/examples"
