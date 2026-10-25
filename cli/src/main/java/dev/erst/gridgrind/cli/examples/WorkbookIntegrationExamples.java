package dev.erst.gridgrind.cli.examples;

import dev.erst.gridgrind.core.model.customxml.CustomXmlImportInput;
import dev.erst.gridgrind.core.model.customxml.CustomXmlMappingLocator;
import dev.erst.gridgrind.core.model.drawing.EmbeddedObjectInput;
import dev.erst.gridgrind.core.model.drawing.PictureDataInput;
import dev.erst.gridgrind.core.model.security.OoxmlOpenSecurityInput;
import dev.erst.gridgrind.core.model.value.CellInput;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.DrawingMutationAction;
import dev.erst.gridgrind.core.plan.DrawingObjectSelector;
import dev.erst.gridgrind.core.plan.SheetIntrospectionQuery;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookAssetIntrospectionQuery;
import dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.source.BinarySourceInput;
import dev.erst.gridgrind.core.source.TextSourceInput;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import dev.erst.gridgrind.core.vocabulary.PlanId;

/** Generated examples for repository-backed assets, external payloads, and package workflows. */
final class WorkbookIntegrationExamples {
  private WorkbookIntegrationExamples() {}

  static WorkbookPlan sourceBackedInputPlan(ExamplePathLayout paths) {
    return ExampleWorkbookPlans.defaultExecutionPlan(
        new PlanId("source-backed-input-workflow"),
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
                        BinarySourceInput.inlineBase64(
                            TaskStarterRecipeSupport.onePixelPngBase64())),
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
        new PlanId("custom-xml-workflow"),
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
        new PlanId("package-security-inspection-workflow"),
        new WorkbookPlan.WorkbookSource.ExistingFile(
            paths.asset("package-security-assets/gridgrind-package-security.xlsx"),
            new OoxmlOpenSecurityInput(
                java.util.Optional.of(
                    new dev.erst.gridgrind.core.source.SecretReference("source-open-password")))),
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
