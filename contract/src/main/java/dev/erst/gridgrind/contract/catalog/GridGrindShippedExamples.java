package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;

/** Contract-owned generated example workbook plans published through the CLI and repository. */
public final class GridGrindShippedExamples {
  /** One built-in example request emitted by the CLI and mirrored under `examples/`. */
  public record ShippedExample(String id, String fileName, String summary, WorkbookPlan plan) {
    public ShippedExample {
      id = CatalogRecordValidation.requireNonBlank(id, "id");
      if (!id.equals(id.toUpperCase(Locale.ROOT))) {
        throw new IllegalArgumentException("id must use upper-case discovery tokens");
      }
      fileName = CatalogRecordValidation.requireNonBlank(fileName, "fileName");
      summary = CatalogRecordValidation.requireNonBlank(summary, "summary");
      Objects.requireNonNull(plan, "plan must not be null");
      if (!fileName.endsWith(".json")) {
        throw new IllegalArgumentException("fileName must end with .json");
      }
    }
  }

  record ExampleRequirements(
      ShippedExampleEntry.WorkspaceMode workspaceMode, List<String> requiredPaths) {
    ExampleRequirements {
      Objects.requireNonNull(workspaceMode, "workspaceMode must not be null");
      requiredPaths = requiredPaths == null ? List.of() : List.copyOf(requiredPaths);
    }
  }

  private static final List<ShippedExample> EXAMPLES = buildExamples(ExamplePathLayout.BUILT_IN);
  private static final List<ShippedExample> REPOSITORY_EXAMPLES =
      buildExamples(ExamplePathLayout.REPOSITORY);
  private static final Map<String, ExampleRequirements> EXAMPLE_REQUIREMENTS =
      buildExampleRequirements();

  private GridGrindShippedExamples() {}

  /** Returns the ordered list of built-in examples. */
  public static List<ShippedExample> examples() {
    return EXAMPLES;
  }

  /**
   * Returns the checked-in example fixtures rooted for in-repository execution from `examples/`.
   */
  public static List<ShippedExample> repositoryExamples() {
    return REPOSITORY_EXAMPLES;
  }

  /** Returns built-in examples that can execute from a blank artifact workspace. */
  public static List<ShippedExample> selfContainedExamples() {
    return EXAMPLES.stream()
        .filter(
            example ->
                catalogEntry(example).workspaceMode()
                    == ShippedExampleEntry.WorkspaceMode.BLANK_WORKSPACE)
        .toList();
  }

  /** Returns built-in examples that require copied repository asset directories. */
  public static List<ShippedExample> repositoryAssetBackedExamples() {
    return EXAMPLES.stream()
        .filter(
            example ->
                catalogEntry(example).workspaceMode()
                    == ShippedExampleEntry.WorkspaceMode.REPOSITORY_ASSETS)
        .toList();
  }

  /** Returns public catalog metadata for the built-in example set. */
  public static List<ShippedExampleEntry> catalogEntries() {
    return EXAMPLES.stream().map(GridGrindShippedExamples::catalogEntry).toList();
  }

  /** Finds one built-in example by its stable upper-case id. */
  public static Optional<ShippedExample> find(String id) {
    Objects.requireNonNull(id, "id must not be null");
    return EXAMPLES.stream().filter(example -> example.id().equals(id)).findFirst();
  }

  static ShippedExampleEntry catalogEntry(ShippedExample example) {
    ExampleRequirements requirements = EXAMPLE_REQUIREMENTS.get(example.id());
    if (requirements == null) {
      throw new IllegalStateException("Missing shipped-example requirements for " + example.id());
    }
    return new ShippedExampleEntry(
        example.id(),
        example.fileName(),
        example.summary(),
        requirements.workspaceMode(),
        requirements.requiredPaths());
  }

  private static Map<String, ExampleRequirements> buildExampleRequirements() {
    return Map.ofEntries(
        entry("BUDGET", ShippedExampleEntry.WorkspaceMode.BLANK_WORKSPACE),
        entry("WORKBOOK_HEALTH", ShippedExampleEntry.WorkspaceMode.BLANK_WORKSPACE),
        entry("SHEET_MAINTENANCE", ShippedExampleEntry.WorkspaceMode.BLANK_WORKSPACE),
        entry("ASSERTION", ShippedExampleEntry.WorkspaceMode.BLANK_WORKSPACE),
        entry("ARRAY_FORMULA", ShippedExampleEntry.WorkspaceMode.BLANK_WORKSPACE),
        entry(
            "CUSTOM_XML",
            ShippedExampleEntry.WorkspaceMode.REPOSITORY_ASSETS,
            "custom-xml-assets/custom-xml-mapping.xlsx",
            "custom-xml-assets/custom-xml-update.xml"),
        entry(
            "SOURCE_BACKED_INPUT",
            ShippedExampleEntry.WorkspaceMode.REPOSITORY_ASSETS,
            "source-backed-input-assets/title.txt",
            "source-backed-input-assets/total-formula.txt",
            "source-backed-input-assets/payload.bin"),
        entry("SIGNATURE_LINE", ShippedExampleEntry.WorkspaceMode.BLANK_WORKSPACE),
        entry("LARGE_FILE_MODES", ShippedExampleEntry.WorkspaceMode.BLANK_WORKSPACE),
        entry("CHART", ShippedExampleEntry.WorkspaceMode.BLANK_WORKSPACE),
        entry("PIVOT", ShippedExampleEntry.WorkspaceMode.BLANK_WORKSPACE),
        entry(
            "PACKAGE_SECURITY_INSPECTION",
            ShippedExampleEntry.WorkspaceMode.REPOSITORY_ASSETS,
            "package-security-assets/gridgrind-package-security.xlsx"),
        entry("FILE_HYPERLINK_HEALTH", ShippedExampleEntry.WorkspaceMode.BLANK_WORKSPACE),
        entry("INTROSPECTION_ANALYSIS", ShippedExampleEntry.WorkspaceMode.BLANK_WORKSPACE));
  }

  private static Map.Entry<String, ExampleRequirements> entry(
      String id, ShippedExampleEntry.WorkspaceMode workspaceMode, String... requiredPaths) {
    return Map.entry(id, new ExampleRequirements(workspaceMode, List.of(requiredPaths)));
  }

  private static List<ShippedExample> buildExamples(ExamplePathLayout paths) {
    return List.of(
        WorkbookWorkflowExamples.budgetExample(paths),
        WorkbookWorkflowExamples.workbookHealthExample(paths),
        WorkbookWorkflowExamples.sheetMaintenanceExample(paths),
        WorkbookWorkflowExamples.assertionExample(paths),
        WorkbookAssetExamples.arrayFormulaExample(paths),
        WorkbookAssetExamples.customXmlExample(paths),
        WorkbookAssetExamples.sourceBackedInputExample(paths),
        WorkbookAssetExamples.signatureLineExample(paths),
        WorkbookWorkflowExamples.largeFileModesExample(paths),
        WorkbookAssetExamples.chartExample(paths),
        WorkbookAssetExamples.pivotExample(paths),
        WorkbookAssetExamples.packageSecurityInspectionExample(paths),
        WorkbookWorkflowExamples.fileHyperlinkHealthExample(paths),
        WorkbookWorkflowExamples.introspectionAnalysisExample(paths));
  }
}
