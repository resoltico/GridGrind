package dev.erst.gridgrind.cli.discovery;

import java.util.List;
import java.util.Objects;

/** Portability contract for one printed task starter request emitted by the CLI. */
public record TaskStarterContract(
    String suggestedRequestPath, ExampleWorkspaceMode workspaceMode, List<String> requiredPaths) {
  public TaskStarterContract {
    suggestedRequestPath =
        CliDiscoveryValidation.requireNonBlank(suggestedRequestPath, "suggestedRequestPath");
    Objects.requireNonNull(workspaceMode, "workspaceMode must not be null");
    requiredPaths = CliDiscoveryValidation.copyStringsAllowEmpty(requiredPaths, "requiredPaths");
    if (!suggestedRequestPath.startsWith("tasks/")) {
      throw new IllegalArgumentException("suggestedRequestPath must start with tasks/");
    }
    if (!suggestedRequestPath.endsWith(".json")) {
      throw new IllegalArgumentException("suggestedRequestPath must end with .json");
    }
    if (workspaceMode == ExampleWorkspaceMode.SELF_CONTAINED && !requiredPaths.isEmpty()) {
      throw new IllegalArgumentException(
          "SELF_CONTAINED task starters must not publish requiredPaths");
    }
    if (workspaceMode == ExampleWorkspaceMode.REQUIRES_EXAMPLE_ASSETS && requiredPaths.isEmpty()) {
      throw new IllegalArgumentException(
          "REQUIRES_EXAMPLE_ASSETS task starters must publish requiredPaths");
    }
  }

  /** Creates one self-contained task starter contract rooted at the task starters directory. */
  public static TaskStarterContract selfContained(String suggestedRequestPath) {
    return new TaskStarterContract(
        suggestedRequestPath, ExampleWorkspaceMode.SELF_CONTAINED, List.of());
  }

  /** Creates one asset-backed task starter contract rooted at the task starters directory. */
  public static TaskStarterContract assetBacked(
      String suggestedRequestPath, String... requiredPaths) {
    return new TaskStarterContract(
        suggestedRequestPath, ExampleWorkspaceMode.REQUIRES_EXAMPLE_ASSETS, List.of(requiredPaths));
  }
}
