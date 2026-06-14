package dev.erst.gridgrind.cli.discovery;

/** Public metadata for one generated built-in example workbook plan. */
public record ShippedExampleEntry(
    String id,
    String requestFileName,
    String summary,
    ExampleWorkspaceMode workspaceMode,
    java.util.List<String> requiredWorkspacePaths) {
  public ShippedExampleEntry {
    id = CliDiscoveryValidation.requireNonBlank(id, "id");
    requestFileName = CliDiscoveryValidation.requireNonBlank(requestFileName, "requestFileName");
    summary = CliDiscoveryValidation.requireNonBlank(summary, "summary");
    java.util.Objects.requireNonNull(workspaceMode, "workspaceMode must not be null");
    requiredWorkspacePaths =
        CliDiscoveryValidation.copyStringsAllowEmpty(
            requiredWorkspacePaths, "requiredWorkspacePaths");
    if (requestFileName.contains("/") || requestFileName.contains("\\")) {
      throw new IllegalArgumentException(
          "requestFileName must be one portable file name, not a repository path");
    }
    if (!requestFileName.endsWith(".json")) {
      throw new IllegalArgumentException("requestFileName must end with .json");
    }
  }
}
