package dev.erst.gridgrind.cli.examples;

import dev.erst.gridgrind.cli.discovery.RecipeAdvisory;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.List;

/** Canonical published recipe list for the CLI recipe registry. */
final class GridGrindRecipeDefinitions {
  private static final List<GridGrindRecipeDefinition> DEFINITIONS =
      List.of(
          example(
              "PIVOT",
              "pivot-request.json",
              "Pivot authoring from a contiguous range with pivot readback and health analysis.",
              RecipeAdvisory.SELF_CONTAINED,
              List.of("pivot", "report", "analysis", "authoring"),
              WorkbookVisualizationExamples.pivotPlan(ExamplePathLayout.BUILT_IN),
              WorkbookVisualizationExamples.pivotPlan(ExamplePathLayout.REPOSITORY)),
          example(
              "CHART",
              "chart-request.json",
              "Supported chart authoring with named-range-backed series and factual chart readback.",
              RecipeAdvisory.SELF_CONTAINED,
              List.of("dashboard", "chart", "visual", "authoring"),
              WorkbookVisualizationExamples.chartPlan(ExamplePathLayout.BUILT_IN),
              WorkbookVisualizationExamples.chartPlan(ExamplePathLayout.REPOSITORY)),
          example(
              "CUSTOM_XML",
              "custom-xml-request.json",
              "Repo-asset-backed existing-workbook custom-XML mapping discovery, XML export, and file-backed XML import.",
              RecipeAdvisory.REQUIRES_EXAMPLE_ASSETS,
              List.of(
                  "custom-xml-assets/custom-xml-mapping.xlsx",
                  "custom-xml-assets/custom-xml-update.xml"),
              List.of("custom-xml", "mapping", "integration", "readback"),
              WorkbookIntegrationExamples.customXmlPlan(ExamplePathLayout.BUILT_IN),
              WorkbookIntegrationExamples.customXmlPlan(ExamplePathLayout.REPOSITORY)),
          example(
              "SHEET_MAINTENANCE",
              "sheet-maintenance-request.json",
              "Copy-sheet maintenance walkthrough with comment reread and workbook findings.",
              RecipeAdvisory.SELF_CONTAINED,
              List.of("maintenance", "sheet", "authoring", "copy", "cleanup"),
              WorkbookAuthoringExamples.sheetMaintenancePlan(ExamplePathLayout.BUILT_IN),
              WorkbookAuthoringExamples.sheetMaintenancePlan(ExamplePathLayout.REPOSITORY)),
          example(
              "WORKBOOK_HEALTH",
              "workbook-health-request.json",
              "Compact no-save workbook-health pass with targeted formula and aggregate findings.",
              RecipeAdvisory.SELF_CONTAINED,
              List.of("audit", "analysis", "health", "inspection", "formula"),
              WorkbookAuditExamples.workbookHealthPlan(ExamplePathLayout.BUILT_IN),
              WorkbookAuditExamples.workbookHealthPlan(ExamplePathLayout.REPOSITORY)),
          example(
              "ARRAY_FORMULA",
              "array-formula-request.json",
              "Array-formula authoring with factual group readback and group clearing.",
              RecipeAdvisory.SELF_CONTAINED,
              List.of("formula", "array-formula", "authoring", "inspection"),
              WorkbookAuthoringExamples.arrayFormulaPlan(ExamplePathLayout.BUILT_IN),
              WorkbookAuthoringExamples.arrayFormulaPlan(ExamplePathLayout.REPOSITORY)),
          example(
              "BUDGET",
              "budget-request.json",
              "Selector-first budget sheet with styling, formula totals, readback, and schema inspection.",
              RecipeAdvisory.SELF_CONTAINED,
              List.of("budget", "authoring", "table", "formula", "inspection"),
              WorkbookAuthoringExamples.budgetPlan(ExamplePathLayout.BUILT_IN),
              WorkbookAuthoringExamples.budgetPlan(ExamplePathLayout.REPOSITORY)),
          example(
              "FILE_HYPERLINK_HEALTH",
              "file-hyperlink-health-request.json",
              "File and document hyperlink authoring with explicit hyperlink-health analysis.",
              RecipeAdvisory.SELF_CONTAINED,
              List.of("hyperlink", "health", "analysis", "inspection"),
              WorkbookAuditExamples.fileHyperlinkHealthPlan(ExamplePathLayout.BUILT_IN),
              WorkbookAuditExamples.fileHyperlinkHealthPlan(ExamplePathLayout.REPOSITORY)),
          example(
              "LARGE_FILE_MODES",
              "large-file-modes-request.json",
              "Low-memory STREAMING_WRITE plan with append-only rows and recalc-on-open flagging.",
              RecipeAdvisory.SELF_CONTAINED,
              List.of("execution-mode", "streaming", "large-files", "analysis"),
              WorkbookAuditExamples.largeFileModesPlan(ExamplePathLayout.BUILT_IN),
              WorkbookAuditExamples.largeFileModesPlan(ExamplePathLayout.REPOSITORY)),
          example(
              "INTROSPECTION_ANALYSIS",
              "introspection-analysis-request.json",
              "Batch factual reads plus formula, hyperlink, named-range, and aggregate workbook analysis.",
              RecipeAdvisory.SELF_CONTAINED,
              List.of("introspection", "schema", "analysis", "inspection"),
              WorkbookAuditExamples.introspectionAnalysisPlan(ExamplePathLayout.BUILT_IN),
              WorkbookAuditExamples.introspectionAnalysisPlan(ExamplePathLayout.REPOSITORY)),
          example(
              "PACKAGE_SECURITY_INSPECTION",
              "package-security-inspect-request.json",
              "Repo-asset-backed encrypted package open plus factual package-security and cell inspection.",
              RecipeAdvisory.REQUIRES_EXAMPLE_ASSETS,
              List.of("package-security-assets/gridgrind-package-security.xlsx"),
              List.of("package-security", "ooxml", "inspection", "encrypted"),
              WorkbookIntegrationExamples.packageSecurityInspectionPlan(ExamplePathLayout.BUILT_IN),
              WorkbookIntegrationExamples.packageSecurityInspectionPlan(
                  ExamplePathLayout.REPOSITORY)),
          example(
              "SIGNATURE_LINE",
              "signature-line-request.json",
              "Signature-line authoring with drawing-object readback and authored anchor replacement.",
              RecipeAdvisory.SELF_CONTAINED,
              List.of("drawing", "signature", "visual", "authoring"),
              WorkbookVisualizationExamples.signatureLinePlan(ExamplePathLayout.BUILT_IN),
              WorkbookVisualizationExamples.signatureLinePlan(ExamplePathLayout.REPOSITORY)),
          example(
              "ASSERTION",
              "assertion-request.json",
              "Mutate then verify with first-class assertions, verbose journaling, and factual readback.",
              RecipeAdvisory.SELF_CONTAINED,
              List.of("assertion", "verification", "fact-readback", "quality"),
              WorkbookAuthoringExamples.assertionPlan(ExamplePathLayout.BUILT_IN),
              WorkbookAuthoringExamples.assertionPlan(ExamplePathLayout.REPOSITORY)),
          example(
              "SOURCE_BACKED_INPUT",
              "source-backed-input-request.json",
              "Repo-asset-backed file text, formula, and binary payload authoring without large inline literals.",
              RecipeAdvisory.REQUIRES_EXAMPLE_ASSETS,
              List.of(
                  "source-backed-input-assets/title.txt",
                  "source-backed-input-assets/total-formula.txt",
                  "source-backed-input-assets/payload.bin"),
              List.of("source-backed", "payloads", "integration", "binary", "text"),
              WorkbookIntegrationExamples.sourceBackedInputPlan(ExamplePathLayout.BUILT_IN),
              WorkbookIntegrationExamples.sourceBackedInputPlan(ExamplePathLayout.REPOSITORY)),
          TabularReportTaskRecipe.definition(),
          DashboardTaskRecipe.definition(),
          DataEntryWorkflowTaskRecipe.definition(),
          PivotReportTaskRecipe.definition(),
          AuditExistingWorkbookTaskRecipe.definition(),
          CustomXmlWorkflowTaskRecipe.definition(),
          DrawingAndSignatureWorkflowTaskRecipe.definition(),
          WorkbookMaintenanceTaskRecipe.definition());

  private GridGrindRecipeDefinitions() {}

  static List<GridGrindRecipeDefinition> definitions() {
    return DEFINITIONS;
  }

  private static GridGrindExampleRecipeDefinition example(
      String id,
      String requestFileName,
      String summary,
      RecipeAdvisory advisory,
      List<String> intentTags,
      WorkbookPlan builtInPlan,
      WorkbookPlan repositoryPlan) {
    return new GridGrindExampleRecipeDefinition(
        id, requestFileName, summary, advisory, List.of(), intentTags, builtInPlan, repositoryPlan);
  }

  private static GridGrindExampleRecipeDefinition example(
      String id,
      String requestFileName,
      String summary,
      RecipeAdvisory advisory,
      List<String> requiredWorkspacePaths,
      List<String> intentTags,
      WorkbookPlan builtInPlan,
      WorkbookPlan repositoryPlan) {
    return new GridGrindExampleRecipeDefinition(
        id,
        requestFileName,
        summary,
        advisory,
        requiredWorkspacePaths,
        intentTags,
        builtInPlan,
        repositoryPlan);
  }
}
