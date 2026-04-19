package dev.erst.gridgrind.contract.catalog;

import java.util.List;
import java.util.Objects;

/** Contract-owned CLI help metadata rendered by thin downstream transports. */
public record CliSurface(
    List<String> executionLines,
    List<String> limitLines,
    List<String> requestLines,
    List<String> fileWorkflowLines,
    List<CoordinateSystemEntry> coordinateSystems,
    List<String> discoveryLines,
    String standardInputRequiresRequestMessage) {
  public CliSurface {
    executionLines = CatalogRecordValidation.copyStrings(executionLines, "executionLines");
    limitLines = CatalogRecordValidation.copyStrings(limitLines, "limitLines");
    requestLines = CatalogRecordValidation.copyStrings(requestLines, "requestLines");
    fileWorkflowLines = CatalogRecordValidation.copyStrings(fileWorkflowLines, "fileWorkflowLines");
    Objects.requireNonNull(coordinateSystems, "coordinateSystems must not be null");
    coordinateSystems =
        coordinateSystems.stream()
            .map(entry -> Objects.requireNonNull(entry, "coordinateSystems must not contain nulls"))
            .toList();
    discoveryLines = CatalogRecordValidation.copyStrings(discoveryLines, "discoveryLines");
    standardInputRequiresRequestMessage =
        CatalogRecordValidation.requireNonBlank(
            standardInputRequiresRequestMessage, "standardInputRequiresRequestMessage");
  }

  /** One coordinate-system row published in CLI help and the machine-readable catalog. */
  public record CoordinateSystemEntry(String pattern, String convention) {
    public CoordinateSystemEntry {
      pattern = CatalogRecordValidation.requireNonBlank(pattern, "pattern");
      convention = CatalogRecordValidation.requireNonBlank(convention, "convention");
    }
  }
}
