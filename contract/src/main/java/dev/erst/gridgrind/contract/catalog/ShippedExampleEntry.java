package dev.erst.gridgrind.contract.catalog;

import java.util.ArrayList;
import java.util.List;

/** Public metadata for one generated built-in example workbook plan. */
public record ShippedExampleEntry(
    String id,
    String fileName,
    String summary,
    WorkspaceMode workspaceMode,
    List<String> requiredPaths) {
  /** Describes the minimum workspace shape needed for one built-in example to run. */
  public enum WorkspaceMode {
    BLANK_WORKSPACE,
    REPOSITORY_ASSETS
  }

  public ShippedExampleEntry {
    id = CatalogRecordValidation.requireNonBlank(id, "id");
    fileName = CatalogRecordValidation.requireNonBlank(fileName, "fileName");
    summary = CatalogRecordValidation.requireNonBlank(summary, "summary");
    java.util.Objects.requireNonNull(workspaceMode, "workspaceMode must not be null");
    requiredPaths = copyRequiredPaths(requiredPaths);
    if (!fileName.endsWith(".json")) {
      throw new IllegalArgumentException("fileName must end with .json");
    }
    if (workspaceMode == WorkspaceMode.BLANK_WORKSPACE && !requiredPaths.isEmpty()) {
      throw new IllegalArgumentException(
          "requiredPaths must be empty when workspaceMode is BLANK_WORKSPACE");
    }
    if (workspaceMode == WorkspaceMode.REPOSITORY_ASSETS && requiredPaths.isEmpty()) {
      throw new IllegalArgumentException(
          "requiredPaths must not be empty when workspaceMode is REPOSITORY_ASSETS");
    }
  }

  private static List<String> copyRequiredPaths(List<String> requiredPaths) {
    if (requiredPaths == null) {
      return List.of();
    }
    List<String> copy = new ArrayList<>(requiredPaths.size());
    for (String requiredPath : requiredPaths) {
      String normalized = CatalogRecordValidation.requireNonBlank(requiredPath, "requiredPath");
      if (normalized.endsWith("/")) {
        throw new IllegalArgumentException("requiredPath must not end with /");
      }
      copy.add(normalized);
    }
    return List.copyOf(copy);
  }
}
