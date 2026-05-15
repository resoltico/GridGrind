package dev.erst.gridgrind.cli.discovery;

import static dev.erst.gridgrind.cli.discovery.GridGrindTaskEntrySupport.discovery;
import static dev.erst.gridgrind.cli.discovery.GridGrindTaskEntrySupport.intent;
import static dev.erst.gridgrind.cli.discovery.GridGrindTaskEntrySupport.narrative;
import static dev.erst.gridgrind.cli.discovery.GridGrindTaskEntrySupport.phase;
import static dev.erst.gridgrind.cli.discovery.GridGrindTaskEntrySupport.profile;
import static dev.erst.gridgrind.cli.discovery.GridGrindTaskEntrySupport.ref;
import static dev.erst.gridgrind.cli.discovery.GridGrindTaskEntrySupport.signals;
import static dev.erst.gridgrind.cli.discovery.GridGrindTaskEntrySupport.task;
import static dev.erst.gridgrind.cli.discovery.GridGrindTaskEntrySupport.workflow;

import java.util.List;

/** Task descriptor for custom-XML import/export workflows. */
final class CustomXmlWorkflowTaskDefinition {
  private CustomXmlWorkflowTaskDefinition() {}

  static TaskEntry entry() {
    return task(
        "CUSTOM_XML_WORKFLOW",
        discovery(
            List.of("custom xml", "mapping", "xml import", "xml export", "mapped workbook"),
            List.of("office", "xml", "mapping", "import", "export"),
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
            "Inspect, export, and import existing workbook custom-XML mappings on mapped .xlsx"
                + " files.",
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
                    List.of("Post-import export is the fastest proof that the mapping changed."))),
            List.of(
                "IMPORT_CUSTOM_XML_MAPPING requires an existing mapping; discovery comes first.",
                "Large XML payloads belong in UTF8_FILE or STANDARD_INPUT sources instead of"
                    + " huge inline JSON.",
                "Mapped workbook structure is workbook-specific, so replace placeholder mapping"
                    + " locators before execution.")));
  }
}
