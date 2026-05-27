package dev.erst.gridgrind.cli.examples;

import dev.erst.gridgrind.cli.discovery.ExampleWorkspaceMode;
import dev.erst.gridgrind.cli.discovery.ShippedExampleCatalog;
import dev.erst.gridgrind.cli.discovery.ShippedExampleEntry;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;

/** CLI-owned generated example workbook plans used by the CLI and repository fixtures. */
public final class GridGrindShippedExamples {
  /** One built-in example request emitted by the CLI and mirrored under `examples/`. */
  public record ShippedExample(
      String id, String requestFileName, String summary, WorkbookPlan plan) {
    public ShippedExample {
      id = requireNonBlank(id, "id");
      if (!id.equals(id.toUpperCase(Locale.ROOT))) {
        throw new IllegalArgumentException("id must use upper-case discovery tokens");
      }
      requestFileName = requireNonBlank(requestFileName, "requestFileName");
      summary = requireNonBlank(summary, "summary");
      Objects.requireNonNull(plan, "plan must not be null");
      if (!requestFileName.endsWith(".json")) {
        throw new IllegalArgumentException("requestFileName must end with .json");
      }
    }
  }

  /** Indicates whether one built-in example is portable or requires repository assets. */
  public record ExampleRequirements(
      ExampleWorkspaceMode workspaceMode, List<String> requiredPaths) {
    public ExampleRequirements {
      Objects.requireNonNull(workspaceMode, "workspaceMode must not be null");
      requiredPaths =
          List.copyOf(Objects.requireNonNull(requiredPaths, "requiredPaths must not be null"));
    }
  }

  private static final List<ShippedExample> EXAMPLES = buildExamples(ExamplePathLayout.BUILT_IN);
  private static final List<ShippedExample> REPOSITORY_EXAMPLES =
      buildExamples(ExamplePathLayout.REPOSITORY);
  private static final Map<String, ExampleRequirements> EXAMPLE_REQUIREMENTS =
      buildExampleRequirements();
  private static final ShippedExampleCatalog CATALOG =
      new ShippedExampleCatalog(
          dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion.current(), catalogEntries());

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
                requirementsFor(example).workspaceMode() == ExampleWorkspaceMode.SELF_CONTAINED)
        .toList();
  }

  /** Returns built-in examples that require copied repository asset directories. */
  public static List<ShippedExample> repositoryAssetBackedExamples() {
    return EXAMPLES.stream()
        .filter(
            example ->
                requirementsFor(example).workspaceMode()
                    == ExampleWorkspaceMode.REQUIRES_EXAMPLE_ASSETS)
        .toList();
  }

  /** Returns public catalog metadata for the built-in example set. */
  public static List<ShippedExampleEntry> catalogEntries() {
    return EXAMPLES.stream().map(GridGrindShippedExamples::catalogPublicEntry).toList();
  }

  /** Returns the machine-readable example catalog for CLI discovery. */
  public static ShippedExampleCatalog catalog() {
    return CATALOG;
  }

  /** Finds one built-in example by its stable upper-case id. */
  public static Optional<ShippedExample> find(String id) {
    Objects.requireNonNull(id, "id must not be null");
    return EXAMPLES.stream().filter(example -> example.id().equals(id)).findFirst();
  }

  /** Returns the portability contract for one stable built-in example id. */
  public static Optional<ExampleWorkspaceMode> workspaceModeFor(String id) {
    Objects.requireNonNull(id, "id must not be null");
    return find(id)
        .map(GridGrindShippedExamples::requirementsFor)
        .map(ExampleRequirements::workspaceMode);
  }

  /** Returns the portability requirements for one concrete built-in example entry. */
  public static ExampleRequirements requirementsFor(ShippedExample example) {
    Objects.requireNonNull(example, "example must not be null");
    ExampleRequirements requirements = EXAMPLE_REQUIREMENTS.get(example.id());
    if (requirements == null) {
      throw new IllegalStateException("Missing shipped-example requirements for " + example.id());
    }
    return requirements;
  }

  private static ShippedExampleEntry catalogPublicEntry(ShippedExample example) {
    Objects.requireNonNull(example, "example must not be null");
    ExampleRequirements requirements = requirementsFor(example);
    return new ShippedExampleEntry(
        example.id(),
        "examples/" + example.requestFileName(),
        example.summary(),
        requirements.workspaceMode(),
        requirements.requiredPaths());
  }

  private static Map<String, ExampleRequirements> buildExampleRequirements() {
    return Map.ofEntries(
        entry("BUDGET", ExampleWorkspaceMode.SELF_CONTAINED),
        entry("WORKBOOK_HEALTH", ExampleWorkspaceMode.SELF_CONTAINED),
        entry("SHEET_MAINTENANCE", ExampleWorkspaceMode.SELF_CONTAINED),
        entry("ASSERTION", ExampleWorkspaceMode.SELF_CONTAINED),
        entry("ARRAY_FORMULA", ExampleWorkspaceMode.SELF_CONTAINED),
        entry(
            "CUSTOM_XML",
            ExampleWorkspaceMode.REQUIRES_EXAMPLE_ASSETS,
            "custom-xml-assets/custom-xml-mapping.xlsx",
            "custom-xml-assets/custom-xml-update.xml"),
        entry(
            "SOURCE_BACKED_INPUT",
            ExampleWorkspaceMode.REQUIRES_EXAMPLE_ASSETS,
            "source-backed-input-assets/title.txt",
            "source-backed-input-assets/total-formula.txt",
            "source-backed-input-assets/payload.bin"),
        entry("SIGNATURE_LINE", ExampleWorkspaceMode.SELF_CONTAINED),
        entry("LARGE_FILE_MODES", ExampleWorkspaceMode.SELF_CONTAINED),
        entry("CHART", ExampleWorkspaceMode.SELF_CONTAINED),
        entry("PIVOT", ExampleWorkspaceMode.SELF_CONTAINED),
        entry(
            "PACKAGE_SECURITY_INSPECTION",
            ExampleWorkspaceMode.REQUIRES_EXAMPLE_ASSETS,
            "package-security-assets/gridgrind-package-security.xlsx"),
        entry("FILE_HYPERLINK_HEALTH", ExampleWorkspaceMode.SELF_CONTAINED),
        entry("INTROSPECTION_ANALYSIS", ExampleWorkspaceMode.SELF_CONTAINED));
  }

  private static Map.Entry<String, ExampleRequirements> entry(
      String id, ExampleWorkspaceMode workspaceMode, String... requiredPaths) {
    return Map.entry(id, new ExampleRequirements(workspaceMode, List.of(requiredPaths)));
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  private static List<ShippedExample> buildExamples(ExamplePathLayout paths) {
    return List.of(
        WorkbookAuthoringExamples.budgetExample(paths),
        WorkbookAuditExamples.workbookHealthExample(paths),
        WorkbookAuthoringExamples.sheetMaintenanceExample(paths),
        WorkbookAuthoringExamples.assertionExample(paths),
        WorkbookAuthoringExamples.arrayFormulaExample(paths),
        WorkbookIntegrationExamples.customXmlExample(paths),
        WorkbookIntegrationExamples.sourceBackedInputExample(paths),
        WorkbookVisualizationExamples.signatureLineExample(paths),
        WorkbookAuditExamples.largeFileModesExample(paths),
        WorkbookVisualizationExamples.chartExample(paths),
        WorkbookVisualizationExamples.pivotExample(paths),
        WorkbookIntegrationExamples.packageSecurityInspectionExample(paths),
        WorkbookAuditExamples.fileHyperlinkHealthExample(paths),
        WorkbookAuditExamples.introspectionAnalysisExample(paths));
  }
}
