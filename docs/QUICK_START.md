# Quick start

Use the pinned Eclipse Temurin 27 JDK and checked-in Gradle wrapper. The
[local setup](DEVELOPER.md#toolchain-and-dependencies) and
[contributor container](DEVELOPER.md#contributor-container) use the same JDK as the runtime image:

    ./gradlew :cli:shadowJar
    java -jar cli/build/libs/gridgrind.jar help

Materialize a published recipe into a new workspace:

    java -jar cli/build/libs/gridgrind.jar recipe materialize BUDGET --workspace budget-workspace
    java -jar cli/build/libs/gridgrind.jar template grant --response budget-workspace/grant.json
    java -jar cli/build/libs/gridgrind.jar recipe show BUDGET --details

Edit grant.json to permit the recipe's exact operation IDs, read resources, workbook target
scope and SAVE_AS output path/collision policy. Set the required `maxInputSnapshotBytes`
allowance for file and standard-input acquisition; the template starts at zero. It authorizes
no operations until edited.
Then diagnose and execute:

    java -jar cli/build/libs/gridgrind.jar doctor --request budget-workspace/budget-request.json --grant budget-workspace/grant.json
    java -jar cli/build/libs/gridgrind.jar run --request budget-workspace/budget-request.json --grant budget-workspace/grant.json --response budget-workspace/result.json

Request-owned relative paths resolve from the request file's directory. Host response/grant
filenames resolve from the process working directory. Relative FILE read-authority and SAVE_AS
publication-authority paths inside the grant also use that process directory. For the BUDGET
commands above, the request's `generated-workbooks/gridgrind-budget.xlsx` refers to
`budget-workspace/generated-workbooks/gridgrind-budget.xlsx` from the shell's working directory;
use that latter path or its absolute equivalent in `publicationAuthority.path`. The grant file's
own directory does not establish an authority root. Stdin requires an explicit request root:

    java -jar cli/build/libs/gridgrind.jar run --request - --execution-root budget-workspace --grant budget-workspace/grant.json < budget-workspace/budget-request.json

The request document and STANDARD_INPUT payloads cannot consume the same stream. Response
publication creates a new file; choose a fresh response path on each run.

## Container

Published versions are at `ghcr.io/resoltico/gridgrind:X.Y.Z`; use an explicit version for
reproducible runs. The current release workflow supplies matching container sources under
`X.Y.Z-sources` and in the [GitHub release](https://github.com/resoltico/GridGrind/releases).
The source tag is a data-only image; see [source retrieval](RELEASING.md#container-source-delivery).

The application and contributor images support AMD64 and 64-bit ARM, including
[Raspberry Pi](DEVELOPER.md#raspberry-pi). Build the local image and mount the caller-owned
workspace at its identical absolute path.
This keeps absolute grant paths valid inside the container. Match the host uid/gid so output
files remain caller-owned:

    docker buildx build --load -t gridgrind-local .
    workspace="$(pwd)"
    docker run --rm --user "$(id -u):$(id -g)" -v "$workspace:$workspace" -w "$workspace" gridgrind-local help

For a prepared request and grant in the mounted directory:

    docker run --rm --user "$(id -u):$(id -g)" -v "$workspace:$workspace" -w "$workspace" gridgrind-local run --request budget-workspace/budget-request.json --grant budget-workspace/grant.json --response budget-workspace/container-result.json

## Minimal plan

This complete plan can be inspected by doctor without publication:

```json
{
  "protocolVersion": 1,
  "source": {
    "type": "NEW"
  },
  "persistence": {
    "type": "NONE"
  },
  "steps": [
    {
      "stepId": "ensure-budget",
      "target": {
        "type": "SHEET_BY_NAME",
        "name": "Budget"
      },
      "action": {
        "type": "ENSURE_SHEET"
      }
    }
  ]
}
```

Use [catalog search/show and schemas](CLI.md) for exact supported shapes, the
[generated reference](reference/README.md) for operation metadata, and the
[execution model](EXECUTION_MODEL.md) for authority and failure semantics.
