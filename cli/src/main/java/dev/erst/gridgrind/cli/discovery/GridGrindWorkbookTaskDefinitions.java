package dev.erst.gridgrind.cli.discovery;

import static dev.erst.gridgrind.cli.discovery.GridGrindTaskEntrySupport.phase;
import static dev.erst.gridgrind.cli.discovery.GridGrindTaskEntrySupport.profile;
import static dev.erst.gridgrind.cli.discovery.GridGrindTaskEntrySupport.ref;
import static dev.erst.gridgrind.cli.discovery.GridGrindTaskEntrySupport.signals;
import static dev.erst.gridgrind.cli.discovery.GridGrindTaskEntrySupport.task;
import static dev.erst.gridgrind.cli.discovery.GridGrindTaskEntrySupport.workflow;

import java.util.List;

/** CLI-owned task descriptors for maintenance, XML, audit, and drawing workflows. */
final class GridGrindWorkbookTaskDefinitions {
  private GridGrindWorkbookTaskDefinitions() {}

  static List<TaskEntry> entries() {
    return List.of(
        auditExistingWorkbook(),
        customXmlWorkflow(),
        drawingAndSignatureWorkflow(),
        workbookMaintenance());
  }

  private static TaskEntry auditExistingWorkbook() {
    return task(
        "AUDIT_EXISTING_WORKBOOK",
        "Inspect an existing workbook, surface health findings, and avoid mutation unless a"
            + " follow-up workflow explicitly asks for it.",
        profile(
            TaskSourceMode.EXISTING_WORKBOOK,
            TaskPersistenceMode.NONE,
            TaskMutationMode.READ_ONLY,
            TaskAssetMode.SELF_CONTAINED),
        signals(
            List.of(TaskInputKind.SOURCE_WORKBOOK_PATH, TaskInputKind.TARGET_SHEET_NAMES),
            List.of(TaskVerificationKind.FACT_READBACK, TaskVerificationKind.HEALTH_ANALYSIS)),
        List.of("office", "audit", "analysis", "readback", "safety"),
        List.of(
            "Existing workbook structure is surfaced as facts instead of guesses.",
            "Health analyses aggregate into one workbook-level findings view.",
            "The default audit posture stays read-only and non-persistent."),
        List.of(
            "Source workbook path.",
            "Target sheets or workbook areas that matter to the audit.",
            "Any expected invariants that should later become assertions."),
        List.of(
            "Package security inspection for protected OOXML files.",
            "Targeted formula-surface inspection before broad workbook findings.",
            "Follow-up assertion plans after the first audit pass."),
        workflow(
            List.of(
                phase(
                    TaskPhasePurpose.INSPECT,
                    "Open And Inspect The Package",
                    "Start with workbook-level facts and security state before diving into sheet"
                        + " logic.",
                    List.of(
                        ref("sourceTypes", "EXISTING"),
                        ref("persistenceTypes", "NONE"),
                        ref("inspectionQueryTypes", "GET_WORKBOOK_SUMMARY"),
                        ref("inspectionQueryTypes", "GET_PACKAGE_SECURITY")),
                    List.of("A no-persistence audit pass reduces accidental mutation.")),
                phase(
                    TaskPhasePurpose.ANALYZE,
                    "Analyze The Workbook",
                    "Read formula surfaces and aggregate workbook findings into one operator"
                        + " view.",
                    List.of(
                        ref("inspectionQueryTypes", "GET_FORMULA_SURFACE"),
                        ref("inspectionQueryTypes", "ANALYZE_WORKBOOK_FINDINGS")),
                    List.of(
                        "Use targeted factual reads first when a later finding needs detail."))),
            List.of(
                "EVENT_READ is intentionally limited; many audit plans require FULL_XSSF.",
                "Loaded formulas that Apache POI cannot evaluate surface as UNSUPPORTED_FORMULA"
                    + " instead of silently recalculating.",
                "If the audit needs strict policy checks, promote the findings into ASSERTION"
                    + " steps in a follow-up plan.")));
  }

  private static TaskEntry customXmlWorkflow() {
    return task(
        "CUSTOM_XML_WORKFLOW",
        "Inspect, export, and import existing workbook custom-XML mappings on mapped .xlsx"
            + " files.",
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
        List.of("office", "xml", "mapping", "import", "export"),
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
            "Post-import factual rereads or exports for verification."),
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

  private static TaskEntry drawingAndSignatureWorkflow() {
    return task(
        "DRAWING_AND_SIGNATURE_WORKFLOW",
        "Author or inspect drawing-backed workbook content such as signature lines, pictures,"
            + " shapes, and embedded objects.",
        profile(
            TaskSourceMode.NEW_WORKBOOK,
            TaskPersistenceMode.SAVE_AS,
            TaskMutationMode.MUTATING,
            TaskAssetMode.SELF_CONTAINED),
        signals(
            List.of(
                TaskInputKind.TARGET_SHEET_NAMES,
                TaskInputKind.TARGET_OBJECT_NAMES,
                TaskInputKind.DRAWING_ANCHORS,
                TaskInputKind.PERSISTENCE_TARGET_PATH),
            List.of(TaskVerificationKind.FACT_READBACK)),
        List.of("office", "drawing", "signature", "picture", "embedded-object"),
        List.of(
            "Drawing-backed workbook objects are authored by name instead of fragile XML"
                + " surgery.",
            "Anchors can be moved authoritatively after creation.",
            "Drawing-object readback confirms what the workbook now contains."),
        List.of(
            "Target sheet and object names.",
            "Anchor positions.",
            "Persistence target when the workbook should be saved."),
        List.of(
            "Signature-line preview images.",
            "Follow-up drawing-object payload extraction or deletion.",
            "Picture, shape, or embedded-object variants alongside signature lines."),
        workflow(
            List.of(
                phase(
                    TaskPhasePurpose.PREPARE,
                    "Prepare The Drawing Surface",
                    "Create the target sheet before any named drawing object is authored.",
                    List.of(ref("sourceTypes", "NEW"), ref("mutationActionTypes", "ENSURE_SHEET")),
                    List.of("A stable sheet name gives drawing objects a durable home.")),
                phase(
                    TaskPhasePurpose.AUTHOR,
                    "Author Drawing Objects",
                    "Create one named drawing object and move it authoritatively when needed.",
                    List.of(
                        ref("mutationActionTypes", "SET_SIGNATURE_LINE"),
                        ref("mutationActionTypes", "SET_DRAWING_OBJECT_ANCHOR"),
                        ref("persistenceTypes", "SAVE_AS")),
                    List.of("Named drawing objects are easier to inspect and move later.")),
                phase(
                    TaskPhasePurpose.VERIFY,
                    "Inspect The Result",
                    "Read back worksheet drawing metadata after authoring or anchor changes.",
                    List.of(ref("inspectionQueryTypes", "GET_DRAWING_OBJECTS")),
                    List.of("Use factual reread to confirm the visible workbook surface."))),
            List.of(
                "Signature lines are VML-backed drawing objects with workbook-side constraints.",
                "Invalid drawing payloads are rejected without leaving partial artifacts behind.",
                "Image-heavy workflows should prefer file-backed binary sources over huge inline"
                    + " base64 payloads.")));
  }

  private static TaskEntry workbookMaintenance() {
    return task(
        "WORKBOOK_MAINTENANCE",
        "Safely copy or normalize existing workbook sheets while verifying comments, drawings,"
            + " and other cloned workbook state after the change.",
        profile(
            TaskSourceMode.EXISTING_WORKBOOK,
            TaskPersistenceMode.SAVE_AS,
            TaskMutationMode.MUTATING,
            TaskAssetMode.SELF_CONTAINED),
        signals(
            List.of(
                TaskInputKind.SOURCE_WORKBOOK_PATH,
                TaskInputKind.TARGET_SHEET_NAMES,
                TaskInputKind.PERSISTENCE_TARGET_PATH),
            List.of(TaskVerificationKind.FACT_READBACK, TaskVerificationKind.HEALTH_ANALYSIS)),
        List.of("office", "maintenance", "repair", "copy", "comments"),
        List.of(
            "Workbook maintenance flows can inspect the source before making structural changes.",
            "Copy-sheet work can be followed immediately by comment and drawing readback.",
            "The updated workbook can be saved separately while you verify the result."),
        List.of(
            "Existing workbook path.",
            "Source sheet names and destination copy names.",
            "Persistence target when the maintained workbook should be kept."),
        List.of(
            "Workbook-level findings after the maintenance pass.",
            "Comment rereads on copied sheets.",
            "Drawing-object rereads on copied sheets when they matter."),
        workflow(
            List.of(
                phase(
                    TaskPhasePurpose.INSPECT,
                    "Inspect The Current Workbook",
                    "Read source workbook facts before changing sheet structure.",
                    List.of(
                        ref("sourceTypes", "EXISTING"),
                        ref("inspectionQueryTypes", "GET_WORKBOOK_SUMMARY"),
                        ref("inspectionQueryTypes", "GET_COMMENTS")),
                    List.of("Factual inspection is the baseline for any maintenance pass.")),
                phase(
                    TaskPhasePurpose.AUTHOR,
                    "Apply Structural Maintenance",
                    "Copy one existing sheet into a new workbook-visible destination sheet.",
                    List.of(
                        ref("mutationActionTypes", "COPY_SHEET"),
                        ref("persistenceTypes", "SAVE_AS")),
                    List.of("Copy-sheet work is safest when the destination name is explicit.")),
                phase(
                    TaskPhasePurpose.VERIFY,
                    "Verify The Copied Surface",
                    "Read back copied comments, drawings, and workbook findings immediately.",
                    List.of(
                        ref("inspectionQueryTypes", "GET_COMMENTS"),
                        ref("inspectionQueryTypes", "GET_DRAWING_OBJECTS"),
                        ref("inspectionQueryTypes", "ANALYZE_WORKBOOK_FINDINGS")),
                    List.of("Verification closes the loop on maintenance changes."))),
            List.of(
                "Replace placeholder sheet names before execution; copy-sheet requires an"
                    + " existing source sheet.",
                "Maintenance flows are safest with SAVE_AS so the source workbook stays intact"
                    + " until verification passes.",
                "Use targeted rereads after COPY_SHEET instead of assuming copied comments or"
                    + " drawings stayed coherent.")));
  }
}
