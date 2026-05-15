package dev.erst.gridgrind.cli.discovery;

/** Public metadata for one generated built-in example workbook plan. */
public record ShippedExampleEntry(
    String id,
    String suggestedRequestPath,
    String summary,
    ExampleWorkspaceMode workspaceMode,
    java.util.List<String> requiredPaths) {
  public ShippedExampleEntry {
    id = CliDiscoveryValidation.requireNonBlank(id, "id");
    suggestedRequestPath =
        CliDiscoveryValidation.requireNonBlank(suggestedRequestPath, "suggestedRequestPath");
    summary = CliDiscoveryValidation.requireNonBlank(summary, "summary");
    java.util.Objects.requireNonNull(workspaceMode, "workspaceMode must not be null");
    requiredPaths = CliDiscoveryValidation.copyStringsAllowEmpty(requiredPaths, "requiredPaths");
    if (!suggestedRequestPath.startsWith("examples/")) {
      throw new IllegalArgumentException("suggestedRequestPath must start with examples/");
    }
    if (!suggestedRequestPath.endsWith(".json")) {
      throw new IllegalArgumentException("suggestedRequestPath must end with .json");
    }
  }
}
