package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.List;
import java.util.Locale;
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

  private static final List<ShippedExample> EXAMPLES =
      List.of(
          WorkbookWorkflowExamples.budgetExample(),
          WorkbookWorkflowExamples.workbookHealthExample(),
          WorkbookWorkflowExamples.assertionExample(),
          WorkbookAssetExamples.sourceBackedInputExample(),
          WorkbookWorkflowExamples.largeFileModesExample(),
          WorkbookAssetExamples.chartExample(),
          WorkbookAssetExamples.pivotExample(),
          WorkbookAssetExamples.packageSecurityInspectionExample(),
          WorkbookWorkflowExamples.fileHyperlinkHealthExample(),
          WorkbookWorkflowExamples.introspectionAnalysisExample());

  private GridGrindShippedExamples() {}

  /** Returns the ordered list of built-in examples. */
  public static List<ShippedExample> examples() {
    return EXAMPLES;
  }

  /** Returns public catalog metadata for the built-in example set. */
  public static List<ShippedExampleEntry> catalogEntries() {
    return EXAMPLES.stream()
        .map(
            example -> new ShippedExampleEntry(example.id(), example.fileName(), example.summary()))
        .toList();
  }

  /** Finds one built-in example by its stable upper-case id. */
  public static Optional<ShippedExample> find(String id) {
    Objects.requireNonNull(id, "id must not be null");
    return EXAMPLES.stream().filter(example -> example.id().equals(id)).findFirst();
  }
}
