package dev.erst.gridgrind.core.discovery;

import dev.erst.gridgrind.core.model.customxml.CustomXmlImportSpec;
import dev.erst.gridgrind.core.model.customxml.CustomXmlMappingLocator;
import dev.erst.gridgrind.core.model.drawing.EmbeddedObjectSpec;
import dev.erst.gridgrind.core.model.drawing.PictureDataSpec;
import dev.erst.gridgrind.core.model.security.OoxmlOpenSecuritySpec;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.DrawingMutationAction;
import dev.erst.gridgrind.core.plan.DrawingObjectSelector;
import dev.erst.gridgrind.core.plan.SheetIntrospectionQuery;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookAssetIntrospectionQuery;
import dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.source.BinarySourceSpec;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import dev.erst.gridgrind.core.vocabulary.PlanId;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;

/** Generated examples for repository-backed assets, external payloads, and package workflows. */
final class WorkbookIntegrationExamples {
  private WorkbookIntegrationExamples() {}

  static WorkbookPlan sourceBackedInputPlan(ExamplePathLayout paths) {
    return ExampleWorkbookPlans.defaultExecutionPlan(
        new PlanId("source-backed-input-workflow"),
        new WorkbookPlan.WorkbookSource.New(),
        new WorkbookPlan.WorkbookPersistence.None(),
        ExampleSteps.step(
            new StepId("ensure-inputs"),
            ExampleSelectors.sheet("Inputs"),
            new WorkbookMutationAction.EnsureSheet()),
        ExampleSteps.step(
            new StepId("seed-values"),
            ExampleSelectors.range("Inputs", "B2:B3"),
            new CellMutationAction.SetRange(
                ExampleCellValues.rows(
                    ExampleCellValues.row(ExampleCellValues.number(12.0d)),
                    ExampleCellValues.row(ExampleCellValues.number(18.0d))))),
        ExampleSteps.step(
            new StepId("set-title-from-file"),
            ExampleSelectors.cell("Inputs", "A1"),
            new CellMutationAction.SetCell(
                new CellValueSpec.Text(
                    TextSourceSpec.utf8File(paths.asset("source-backed-input-assets/title.txt"))))),
        ExampleSteps.step(
            new StepId("set-total-formula-from-file"),
            ExampleSelectors.cell("Inputs", "B4"),
            new CellMutationAction.SetCell(
                new CellValueSpec.Formula(
                    TextSourceSpec.utf8File(
                        paths.asset("source-backed-input-assets/total-formula.txt"))))),
        ExampleSteps.step(
            new StepId("attach-payload-from-file"),
            ExampleSelectors.sheet("Inputs"),
            new DrawingMutationAction.SetEmbeddedObject(
                new EmbeddedObjectSpec(
                    "InputsPayload",
                    "Inputs payload",
                    "inputs-payload.txt",
                    "open",
                    BinarySourceSpec.file(paths.asset("source-backed-input-assets/payload.bin")),
                    new PictureDataSpec(
                        ExcelPictureFormat.PNG,
                        BinarySourceSpec.inlineBase64(TaskStarters.onePixelPngBase64())),
                    ExampleDrawingAnchors.anchor(3, 0, 5, 4)))),
        ExampleSteps.read(
            new StepId("read-cells"),
            ExampleSelectors.cells("Inputs", "A1", "B4"),
            new SheetIntrospectionQuery.GetCells()),
        ExampleSteps.read(
            new StepId("read-payload"),
            new DrawingObjectSelector.ByName(new SheetName("Inputs"), "InputsPayload"),
            new WorkbookAssetIntrospectionQuery.GetDrawingObjectPayload()));
  }

  static WorkbookPlan customXmlPlan(ExamplePathLayout paths) {
    return ExampleWorkbookPlans.defaultExecutionPlan(
        new PlanId("custom-xml-workflow"),
        new WorkbookPlan.WorkbookSource.ExistingFile(
            paths.asset("custom-xml-assets/custom-xml-mapping.xlsx")),
        new WorkbookPlan.WorkbookPersistence.None(),
        ExampleSteps.read(
            new StepId("read-custom-xml-mappings"),
            ExampleSelectors.workbook(),
            new WorkbookIntrospectionQuery.GetCustomXmlMappings()),
        ExampleSteps.read(
            new StepId("export-custom-xml-before-import"),
            ExampleSelectors.workbook(),
            new WorkbookIntrospectionQuery.ExportCustomXmlMapping(
                new CustomXmlMappingLocator(
                    java.util.Optional.of(1L), java.util.Optional.of("CORSO_mapping")),
                true,
                "UTF-8")),
        ExampleSteps.step(
            new StepId("import-custom-xml"),
            ExampleSelectors.workbook(),
            new StructuredMutationAction.ImportCustomXmlMapping(
                new CustomXmlImportSpec(
                    new CustomXmlMappingLocator(
                        java.util.Optional.of(1L), java.util.Optional.of("CORSO_mapping")),
                    TextSourceSpec.utf8File(
                        paths.asset("custom-xml-assets/custom-xml-update.xml"))))),
        ExampleSteps.read(
            new StepId("read-imported-cells"),
            ExampleSelectors.cells("Foglio1", "A1", "B1", "C1", "D1"),
            new SheetIntrospectionQuery.GetCells()),
        ExampleSteps.read(
            new StepId("export-custom-xml-after-import"),
            ExampleSelectors.workbook(),
            new WorkbookIntrospectionQuery.ExportCustomXmlMapping(
                new CustomXmlMappingLocator(
                    java.util.Optional.of(1L), java.util.Optional.of("CORSO_mapping")),
                true,
                "UTF-8")));
  }

  static WorkbookPlan packageSecurityInspectionPlan(ExamplePathLayout paths) {
    return ExampleWorkbookPlans.defaultExecutionPlan(
        new PlanId("package-security-inspection-workflow"),
        new WorkbookPlan.WorkbookSource.ExistingFile(
            paths.asset("package-security-assets/gridgrind-package-security.xlsx"),
            new OoxmlOpenSecuritySpec(
                java.util.Optional.of(
                    new dev.erst.gridgrind.core.source.SecretReference("source-open-password")))),
        new WorkbookPlan.WorkbookPersistence.None(),
        ExampleSteps.read(
            new StepId("security"),
            ExampleSelectors.workbook(),
            new WorkbookIntrospectionQuery.GetPackageSecurity()),
        ExampleSteps.read(
            new StepId("secure-cells"),
            ExampleSelectors.cells("Secure", "A1", "A2", "B2", "A3", "B3"),
            new SheetIntrospectionQuery.GetCells()));
  }
}
