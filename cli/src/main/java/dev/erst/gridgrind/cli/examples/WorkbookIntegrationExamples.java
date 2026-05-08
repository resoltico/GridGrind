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

  static GridGrindShippedExamples.ShippedExample sourceBackedInputExample(ExamplePathLayout paths) {
    return ExamplePlanSupport.example(
        "SOURCE_BACKED_INPUT",
        "source-backed-input-request.json",
        "Repo-asset-backed file text, formula, and binary payload authoring without large inline literals.",
        ExamplePlanSupport.defaultExecutionPlan(
            "source-backed-input-workflow",
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExamplePlanSupport.step(
                "ensure-inputs",
                ExamplePlanSupport.sheet("Inputs"),
                new WorkbookMutationAction.EnsureSheet()),
            ExamplePlanSupport.step(
                "seed-values",
                ExamplePlanSupport.range("Inputs", "B2:B3"),
                new CellMutationAction.SetRange(
                    ExamplePlanSupport.rows(
                        ExamplePlanSupport.row(ExamplePlanSupport.number(12.0d)),
                        ExamplePlanSupport.row(ExamplePlanSupport.number(18.0d))))),
            ExamplePlanSupport.step(
                "set-title-from-file",
                ExamplePlanSupport.cell("Inputs", "A1"),
                new CellMutationAction.SetCell(
                    new CellInput.Text(
                        TextSourceInput.utf8File(
                            paths.asset("source-backed-input-assets/title.txt"))))),
            ExamplePlanSupport.step(
                "set-total-formula-from-file",
                ExamplePlanSupport.cell("Inputs", "B4"),
                new CellMutationAction.SetCell(
                    new CellInput.Formula(
                        TextSourceInput.utf8File(
                            paths.asset("source-backed-input-assets/total-formula.txt"))))),
            ExamplePlanSupport.step(
                "attach-payload-from-file",
                ExamplePlanSupport.sheet("Inputs"),
                new DrawingMutationAction.SetEmbeddedObject(
                    new EmbeddedObjectInput(
                        "InputsPayload",
                        "Inputs payload",
                        "inputs-payload.txt",
                        "open",
                        BinarySourceInput.file(
                            paths.asset("source-backed-input-assets/payload.bin")),
                        new PictureDataInput(
                            ExcelPictureFormat.PNG,
                            BinarySourceInput.inlineBase64(ONE_PIXEL_PNG_BASE64)),
                        ExamplePlanSupport.anchor(3, 0, 5, 4)))),
            ExamplePlanSupport.read(
                "read-cells",
                ExamplePlanSupport.cells("Inputs", "A1", "B4"),
                new SheetIntrospectionQuery.GetCells()),
            ExamplePlanSupport.read(
                "read-payload",
                new DrawingObjectSelector.ByName("Inputs", "InputsPayload"),
                new WorkbookAssetIntrospectionQuery.GetDrawingObjectPayload())));
  }

  static GridGrindShippedExamples.ShippedExample customXmlExample(ExamplePathLayout paths) {
    return ExamplePlanSupport.example(
        "CUSTOM_XML",
        "custom-xml-request.json",
        "Repo-asset-backed existing-workbook custom-XML mapping discovery, XML export, and file-backed XML import.",
        ExamplePlanSupport.defaultExecutionPlan(
            "custom-xml-workflow",
            new WorkbookPlan.WorkbookSource.ExistingFile(
                paths.asset("custom-xml-assets/custom-xml-mapping.xlsx")),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExamplePlanSupport.read(
                "read-custom-xml-mappings",
                ExamplePlanSupport.workbook(),
                new WorkbookIntrospectionQuery.GetCustomXmlMappings()),
            ExamplePlanSupport.read(
                "export-custom-xml-before-import",
                ExamplePlanSupport.workbook(),
                new WorkbookIntrospectionQuery.ExportCustomXmlMapping(
                    new CustomXmlMappingLocator(1L, "CORSO_mapping"), true, "UTF-8")),
            ExamplePlanSupport.step(
                "import-custom-xml",
                ExamplePlanSupport.workbook(),
                new StructuredMutationAction.ImportCustomXmlMapping(
                    new CustomXmlImportInput(
                        new CustomXmlMappingLocator(1L, "CORSO_mapping"),
                        TextSourceInput.utf8File(
                            paths.asset("custom-xml-assets/custom-xml-update.xml"))))),
            ExamplePlanSupport.read(
                "read-imported-cells",
                ExamplePlanSupport.cells("Foglio1", "A1", "B1", "C1", "D1"),
                new SheetIntrospectionQuery.GetCells()),
            ExamplePlanSupport.read(
                "export-custom-xml-after-import",
                ExamplePlanSupport.workbook(),
                new WorkbookIntrospectionQuery.ExportCustomXmlMapping(
                    new CustomXmlMappingLocator(1L, "CORSO_mapping"), true, "UTF-8"))));
  }

  static GridGrindShippedExamples.ShippedExample packageSecurityInspectionExample(
      ExamplePathLayout paths) {
    return ExamplePlanSupport.example(
        "PACKAGE_SECURITY_INSPECTION",
        "package-security-inspect-request.json",
        "Repo-asset-backed encrypted package open plus factual package-security and cell inspection.",
        ExamplePlanSupport.defaultExecutionPlan(
            "package-security-inspection-workflow",
            new WorkbookPlan.WorkbookSource.ExistingFile(
                paths.asset("package-security-assets/gridgrind-package-security.xlsx"),
                new OoxmlOpenSecurityInput(java.util.Optional.of("GridGrind-2026"))),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExamplePlanSupport.read(
                "security",
                ExamplePlanSupport.workbook(),
                new WorkbookIntrospectionQuery.GetPackageSecurity()),
            ExamplePlanSupport.read(
                "secure-cells",
                ExamplePlanSupport.cells("Secure", "A1", "A2", "B2", "A3", "B3"),
                new SheetIntrospectionQuery.GetCells())));
  }
}
