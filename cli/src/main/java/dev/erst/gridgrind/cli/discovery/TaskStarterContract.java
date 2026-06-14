package dev.erst.gridgrind.cli.discovery;

import java.util.List;
import java.util.Objects;

/** Portability contract for one printed task starter request emitted by the CLI. */
public record TaskStarterContract(
    String requestFileName,
    ExampleWorkspaceMode workspaceMode,
    List<String> requiredWorkspacePaths) {
  public TaskStarterContract {
    requestFileName = CliDiscoveryValidation.requireNonBlank(requestFileName, "requestFileName");
    Objects.requireNonNull(workspaceMode, "workspaceMode must not be null");
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
    if (workspaceMode == ExampleWorkspaceMode.SELF_CONTAINED && !requiredWorkspacePaths.isEmpty()) {
      throw new IllegalArgumentException(
          "SELF_CONTAINED task starters must not publish requiredWorkspacePaths");
    }
    if (workspaceMode == ExampleWorkspaceMode.REQUIRES_EXAMPLE_ASSETS
        && requiredWorkspacePaths.isEmpty()) {
      throw new IllegalArgumentException(
          "REQUIRES_EXAMPLE_ASSETS task starters must publish requiredWorkspacePaths");
    }
  }

  /** Creates one self-contained task starter contract with one portable request file name. */
  public static TaskStarterContract selfContained(String requestFileName) {
    return new TaskStarterContract(requestFileName, ExampleWorkspaceMode.SELF_CONTAINED, List.of());
  }

  /** Creates one asset-backed task starter contract with workspace-relative asset requirements. */
  public static TaskStarterContract assetBacked(
      String requestFileName, String... requiredWorkspacePaths) {
    return new TaskStarterContract(
        requestFileName,
        ExampleWorkspaceMode.REQUIRES_EXAMPLE_ASSETS,
        List.of(requiredWorkspacePaths));
  }
}
