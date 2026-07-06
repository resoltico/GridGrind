package dev.erst.gridgrind.cli.examples;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.DrawingMutationAction;
import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.CustomXmlImportInput;
import dev.erst.gridgrind.contract.dto.CustomXmlMappingLocator;
import dev.erst.gridgrind.contract.dto.EmbeddedObjectInput;
import dev.erst.gridgrind.contract.dto.OoxmlOpenSecurityInput;
import dev.erst.gridgrind.contract.dto.PictureDataInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.SheetIntrospectionQuery;
import dev.erst.gridgrind.contract.query.WorkbookAssetIntrospectionQuery;
import dev.erst.gridgrind.contract.query.WorkbookIntrospectionQuery;
import dev.erst.gridgrind.contract.selector.DrawingObjectSelector;
import dev.erst.gridgrind.contract.source.BinarySourceInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;

/** Generated examples for repository-backed assets, external payloads, and package workflows. */
final class WorkbookIntegrationExamples {
  private static final String ONE_PIXEL_PNG_BASE64 =
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII=";

  private WorkbookIntegrationExamples() {}

  static WorkbookPlan sourceBackedInputPlan(ExamplePathLayout paths) {
    return ExampleWorkbookPlans.defaultExecutionPlan(
        "source-backed-input-workflow",
        new WorkbookPlan.WorkbookSource.New(),
        new WorkbookPlan.WorkbookPersistence.None(),
        ExampleSteps.step(
            "ensure-inputs",
            ExampleSelectors.sheet("Inputs"),
            new WorkbookMutationAction.EnsureSheet()),
        ExampleSteps.step(
            "seed-values",
            ExampleSelectors.range("Inputs", "B2:B3"),
            new CellMutationAction.SetRange(
                ExampleCellValues.rows(
                    ExampleCellValues.row(ExampleCellValues.number(12.0d)),
                    ExampleCellValues.row(ExampleCellValues.number(18.0d))))),
        ExampleSteps.step(
            "set-title-from-file",
            ExampleSelectors.cell("Inputs", "A1"),
            new CellMutationAction.SetCell(
                new CellInput.Text(
                    TextSourceInput.utf8File(
                        paths.asset("source-backed-input-assets/title.txt"))))),
        ExampleSteps.step(
            "set-total-formula-from-file",
            ExampleSelectors.cell("Inputs", "B4"),
            new CellMutationAction.SetCell(
                new CellInput.Formula(
                    TextSourceInput.utf8File(
                        paths.asset("source-backed-input-assets/total-formula.txt"))))),
        ExampleSteps.step(
            "attach-payload-from-file",
            ExampleSelectors.sheet("Inputs"),
            new DrawingMutationAction.SetEmbeddedObject(
                new EmbeddedObjectInput(
                    "InputsPayload",
                    "Inputs payload",
                    "inputs-payload.txt",
                    "open",
                    BinarySourceInput.file(paths.asset("source-backed-input-assets/payload.bin")),
                    new PictureDataInput(
                        ExcelPictureFormat.PNG,
                        BinarySourceInput.inlineBase64(ONE_PIXEL_PNG_BASE64)),
                    ExampleDrawingAnchors.anchor(3, 0, 5, 4)))),
        ExampleSteps.read(
            "read-cells",
            ExampleSelectors.cells("Inputs", "A1", "B4"),
            new SheetIntrospectionQuery.GetCells()),
        ExampleSteps.read(
            "read-payload",
            new DrawingObjectSelector.ByName("Inputs", "InputsPayload"),
            new WorkbookAssetIntrospectionQuery.GetDrawingObjectPayload()));
  }

  static WorkbookPlan customXmlPlan(ExamplePathLayout paths) {
    return ExampleWorkbookPlans.defaultExecutionPlan(
        "custom-xml-workflow",
        new WorkbookPlan.WorkbookSource.ExistingFile(
            paths.asset("custom-xml-assets/custom-xml-mapping.xlsx")),
        new WorkbookPlan.WorkbookPersistence.None(),
        ExampleSteps.read(
            "read-custom-xml-mappings",
            ExampleSelectors.workbook(),
            new WorkbookIntrospectionQuery.GetCustomXmlMappings()),
        ExampleSteps.read(
            "export-custom-xml-before-import",
            ExampleSelectors.workbook(),
            new WorkbookIntrospectionQuery.ExportCustomXmlMapping(
                new CustomXmlMappingLocator(1L, "CORSO_mapping"), true, "UTF-8")),
        ExampleSteps.step(
            "import-custom-xml",
            ExampleSelectors.workbook(),
            new StructuredMutationAction.ImportCustomXmlMapping(
                new CustomXmlImportInput(
                    new CustomXmlMappingLocator(1L, "CORSO_mapping"),
                    TextSourceInput.utf8File(
                        paths.asset("custom-xml-assets/custom-xml-update.xml"))))),
        ExampleSteps.read(
            "read-imported-cells",
            ExampleSelectors.cells("Foglio1", "A1", "B1", "C1", "D1"),
            new SheetIntrospectionQuery.GetCells()),
        ExampleSteps.read(
            "export-custom-xml-after-import",
            ExampleSelectors.workbook(),
            new WorkbookIntrospectionQuery.ExportCustomXmlMapping(
                new CustomXmlMappingLocator(1L, "CORSO_mapping"), true, "UTF-8")));
  }

  static WorkbookPlan packageSecurityInspectionPlan(ExamplePathLayout paths) {
    return ExampleWorkbookPlans.defaultExecutionPlan(
        "package-security-inspection-workflow",
        new WorkbookPlan.WorkbookSource.ExistingFile(
            paths.asset("package-security-assets/gridgrind-package-security.xlsx"),
            new OoxmlOpenSecurityInput(java.util.Optional.of("GridGrind-2026"))),
        new WorkbookPlan.WorkbookPersistence.None(),
        ExampleSteps.read(
            "security",
            ExampleSelectors.workbook(),
            new WorkbookIntrospectionQuery.GetPackageSecurity()),
        ExampleSteps.read(
            "secure-cells",
            ExampleSelectors.cells("Secure", "A1", "A2", "B2", "A3", "B3"),
            new SheetIntrospectionQuery.GetCells()));
  }
}
