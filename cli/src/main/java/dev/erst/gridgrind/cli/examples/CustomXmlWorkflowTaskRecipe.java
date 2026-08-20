package dev.erst.gridgrind.cli.examples;

import static dev.erst.gridgrind.cli.examples.GridGrindTaskRecipeSupport.assetBackedStarter;
import static dev.erst.gridgrind.cli.examples.GridGrindTaskRecipeSupport.discovery;
import static dev.erst.gridgrind.cli.examples.GridGrindTaskRecipeSupport.intent;
import static dev.erst.gridgrind.cli.examples.GridGrindTaskRecipeSupport.narrative;
import static dev.erst.gridgrind.cli.examples.GridGrindTaskRecipeSupport.phase;
import static dev.erst.gridgrind.cli.examples.GridGrindTaskRecipeSupport.profile;
import static dev.erst.gridgrind.cli.examples.GridGrindTaskRecipeSupport.ref;
import static dev.erst.gridgrind.cli.examples.GridGrindTaskRecipeSupport.signals;
import static dev.erst.gridgrind.cli.examples.GridGrindTaskRecipeSupport.task;
import static dev.erst.gridgrind.cli.examples.GridGrindTaskRecipeSupport.workflow;

import dev.erst.gridgrind.cli.discovery.TaskArtifactKind;
import dev.erst.gridgrind.cli.discovery.TaskAssetMode;
import dev.erst.gridgrind.cli.discovery.TaskGoalKind;
import dev.erst.gridgrind.cli.discovery.TaskInputKind;
import dev.erst.gridgrind.cli.discovery.TaskMutationMode;
import dev.erst.gridgrind.cli.discovery.TaskPersistenceMode;
import dev.erst.gridgrind.cli.discovery.TaskPhasePurpose;
import dev.erst.gridgrind.cli.discovery.TaskSourceMode;
import dev.erst.gridgrind.cli.discovery.TaskVerificationKind;
import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.dto.CustomXmlImportInput;
import dev.erst.gridgrind.contract.dto.CustomXmlMappingLocator;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.SheetIntrospectionQuery;
import dev.erst.gridgrind.contract.query.WorkbookIntrospectionQuery;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import java.util.List;

/** Published task recipe for custom-XML import and export workflows. */
final class CustomXmlWorkflowTaskRecipe {
  private CustomXmlWorkflowTaskRecipe() {}

  static GridGrindTaskRecipeDefinition definition() {
    return GridGrindTaskRecipeSupport.definition(
        task(
            "CUSTOM_XML_WORKFLOW",
            assetBackedStarter(
                "CUSTOM_XML_WORKFLOW",
                "custom-xml-assets/custom-xml-mapping.xlsx",
                "custom-xml-assets/custom-xml-update.xml"),
            List.of("office", "xml", "mapping", "import", "export"),
            discovery(
                List.of("custom xml", "mapping", "xml import", "xml export", "mapped workbook"),
                intent(
                    List.of(
                        TaskGoalKind.INSPECT,
                        TaskGoalKind.EXPORT,
                        TaskGoalKind.IMPORT,
                        TaskGoalKind.VERIFY),
                    List.of(
                        TaskArtifactKind.WORKBOOK,
                        TaskArtifactKind.CUSTOM_XML_MAPPING,
                        TaskArtifactKind.XML_PAYLOAD))),
            narrative(
                "Inspect, export, and import existing workbook custom-XML mappings on mapped .xlsx files.",
                List.of(
                    "Mapped workbook metadata is discovered before any XML payload is pushed in.",
                    "Existing mappings can be exported for inspection or diffing.",
                    "XML payloads can be imported back through the existing workbook mapping."),
                List.of(
                    "Mapped source workbook path.",
                    "Existing mapping id and name.",
                    "XML payload source when the workbook should be updated."),
                List.of(
                    "Schema validation during export.",
                    "SAVE_AS persistence when the imported workbook should be kept separately.",
                    "Post-import factual rereads or exports for verification.")),
            profile(
                TaskSourceMode.EXISTING_WORKBOOK,
                TaskPersistenceMode.SAVE_AS,
                TaskMutationMode.MUTATING,
                TaskAssetMode.REQUIRES_EXTERNAL_PAYLOADS),
            signals(
                List.of(
                    TaskInputKind.SOURCE_WORKBOOK_PATH,
                    TaskInputKind.MAPPING_LOCATOR,
                    TaskInputKind.XML_PAYLOAD,
                    TaskInputKind.PERSISTENCE_TARGET_PATH),
                List.of(TaskVerificationKind.FACT_READBACK, TaskVerificationKind.EXPORT_REREAD)),
            workflow(
                List.of(
                    phase(
                        TaskPhasePurpose.INSPECT,
                        "Discover Existing Mappings",
                        "Read the workbook mapping metadata before attempting export or import.",
                        List.of(
                            ref("sourceTypes", "EXISTING"),
                            ref("inspectionQueryTypes", "GET_CUSTOM_XML_MAPPINGS"),
                            ref("inspectionQueryTypes", "EXPORT_CUSTOM_XML_MAPPING")),
                        List.of("Map discovery tells you which locator to target.")),
                    phase(
                        TaskPhasePurpose.IMPORT,
                        "Import Updated XML",
                        "Push one XML payload through the chosen existing mapping.",
                        List.of(
                            ref("mutationActionTypes", "IMPORT_CUSTOM_XML_MAPPING"),
                            ref("persistenceTypes", "SAVE_AS")),
                        List.of("Use SAVE_AS when you want to preserve the original workbook.")),
                    phase(
                        TaskPhasePurpose.VERIFY,
                        "Verify The Result",
                        "Re-export or reread workbook facts after the import.",
                        List.of(ref("inspectionQueryTypes", "EXPORT_CUSTOM_XML_MAPPING")),
                        List.of(
                            "Post-import export is the fastest proof that the mapping changed."))),
                List.of(
                    "IMPORT_CUSTOM_XML_MAPPING requires an existing mapping; discovery comes first.",
                    "Large XML payloads belong in UTF8_FILE or STANDARD_INPUT sources instead of huge inline JSON.",
                    "Mapped workbook structure is workbook-specific, so replace placeholder mapping locators before execution."))),
        starterPlan());
  }

  private static WorkbookPlan starterPlan() {
    String taskId = "CUSTOM_XML_WORKFLOW";
    String mappingWorkbook = "custom-xml-assets/custom-xml-mapping.xlsx";
    String updateXml = "custom-xml-assets/custom-xml-update.xml";
    return ExampleWorkbookPlans.defaultExecutionPlan(
        TaskStarterRecipeSupport.taskPlanId(taskId),
        new WorkbookPlan.WorkbookSource.ExistingFile(mappingWorkbook),
        ExampleWorkbookPlans.saveAsExisting(TaskStarterRecipeSupport.taskWorkbookPath(taskId)),
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
                    TextSourceInput.utf8File(updateXml)))),
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
}
