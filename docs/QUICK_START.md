# Quick start

Use a Java 26 JDK and the checked-in Gradle wrapper:

    ./gradlew :cli:shadowJar
    java -jar cli/build/libs/gridgrind.jar help

Materialize a published recipe into a new workspace:

    java -jar cli/build/libs/gridgrind.jar recipe materialize BUDGET --workspace budget-workspace
    java -jar cli/build/libs/gridgrind.jar template grant --response budget-workspace/grant.json
    java -jar cli/build/libs/gridgrind.jar recipe show BUDGET --details

Edit grant.json to permit the recipe's exact operation IDs, read resources, workbook target
scope and SAVE_AS output path/collision policy. The template is deliberately fail-closed.
Then diagnose and execute:

    java -jar cli/build/libs/gridgrind.jar doctor --request budget-workspace/budget-request.json --grant budget-workspace/grant.json
    java -jar cli/build/libs/gridgrind.jar run --request budget-workspace/budget-request.json --grant budget-workspace/grant.json --response budget-workspace/result.json

Request-owned relative paths resolve from the request file's directory. Host response/grant
paths resolve from the host invocation. Stdin requires an explicit root:

    java -jar cli/build/libs/gridgrind.jar run --request - --execution-root budget-workspace --grant budget-workspace/grant.json < budget-workspace/budget-request.json

The request document and STANDARD_INPUT payloads cannot consume the same stream. Response
publication creates a new file; choose a fresh response path on each run.

## Container

Build the local image and mount the caller-owned workspace at its identical absolute path.
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
