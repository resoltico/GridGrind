package dev.erst.gridgrind.cli.discovery;

/** Public metadata for one generated built-in example workbook plan. */
public record ShippedExampleEntry(
    String id,
    String fileName,
    String summary,
    ExampleWorkspaceMode workspaceMode,
    java.util.List<String> requiredPaths) {
  public ShippedExampleEntry {
    id = CliDiscoveryValidation.requireNonBlank(id, "id");
    fileName = CliDiscoveryValidation.requireNonBlank(fileName, "fileName");
    summary = CliDiscoveryValidation.requireNonBlank(summary, "summary");
    java.util.Objects.requireNonNull(workspaceMode, "workspaceMode must not be null");
    requiredPaths = CliDiscoveryValidation.copyOptionalStrings(requiredPaths);
    if (!fileName.endsWith(".json")) {
      throw new IllegalArgumentException("fileName must end with .json");
    }
  }
}
