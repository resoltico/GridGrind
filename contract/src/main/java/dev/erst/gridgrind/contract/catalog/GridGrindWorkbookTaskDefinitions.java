package dev.erst.gridgrind.contract.catalog;

import static dev.erst.gridgrind.contract.catalog.GridGrindTaskDefinitionSupport.definition;
import static dev.erst.gridgrind.contract.catalog.GridGrindTaskDefinitionSupport.phase;
import static dev.erst.gridgrind.contract.catalog.GridGrindTaskDefinitionSupport.protocolLookupNote;
import static dev.erst.gridgrind.contract.catalog.GridGrindTaskDefinitionSupport.rebasePlan;
import static dev.erst.gridgrind.contract.catalog.GridGrindTaskDefinitionSupport.ref;
import static dev.erst.gridgrind.contract.catalog.GridGrindTaskDefinitionSupport.task;

import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.DrawingObjectSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import java.util.List;

/** Contract-owned task definitions for maintenance, XML, audit, and drawing workflows. */
final class GridGrindWorkbookTaskDefinitions {
  private GridGrindWorkbookTaskDefinitions() {}

  static List<TaskDefinition> definitions() {
    return List.of(
        auditExistingWorkbook(),
        customXmlWorkflow(),
        drawingAndSignatureWorkflow(),
        workbookMaintenance());
  }

  private static TaskDefinition auditExistingWorkbook() {
    TaskEntry task =
        task(
            "AUDIT_EXISTING_WORKBOOK",
            "Inspect an existing workbook, surface health findings, and avoid mutation unless a"
                + " follow-up workflow explicitly asks for it.",
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
            List.of(
                phase(
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
                    "Analyze The Workbook",
                    "Read formula surfaces and aggregate workbook findings into one operator"
                        + " view.",
                    List.of(
                        ref("inspectionQueryTypes", "GET_FORMULA_SURFACE"),
                        ref("inspectionQueryTypes", "ANALYZE_WORKBOOK_FINDINGS")),
                    List.of(
                        "Use targeted factual reads first when a later finding needs detail."))),
            List.of(
                "EVENT_READ is intentionally limited; many audit plans still require FULL_XSSF.",
                "Loaded formulas that Apache POI cannot evaluate surface as UNSUPPORTED_FORMULA"
                    + " instead of silently recalculating.",
                "If the audit needs strict policy checks, promote the findings into ASSERTION"
                    + " steps in a follow-up plan."));
    return definition(
        task,
        List.of(
            "audit existing workbook",
            "health check workbook",
            "inspect workbook findings",
            "read workbook facts",
            "package security audit"),
        ExamplePlanSupport.defaultExecutionPlan(
            "audit-existing-workbook-starter",
            new WorkbookPlan.WorkbookSource.ExistingFile("audit-input.xlsx"),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExamplePlanSupport.read(
                "workbook",
                ExamplePlanSupport.workbook(),
                new InspectionQuery.GetWorkbookSummary()),
            ExamplePlanSupport.read(
                "package-security",
                ExamplePlanSupport.workbook(),
                new InspectionQuery.GetPackageSecurity()),
            ExamplePlanSupport.read(
                "formula-surface",
                new SheetSelector.All(),
                new InspectionQuery.GetFormulaSurface()),
            ExamplePlanSupport.read(
                "workbook-findings",
                ExamplePlanSupport.workbook(),
                new InspectionQuery.AnalyzeWorkbookFindings())),
        List.of(
            "Replace audit-input.xlsx with the workbook you want to inspect.",
            "The starter plan stays non-destructive by default, so you can run discovery first and"
                + " only add mutations later if the audit calls for them.",
            protocolLookupNote(
                "assertion shapes",
                List.of(
                    "assertionTypes:EXPECT_NAMED_RANGE_PRESENT",
                    "assertionTypes:EXPECT_CELL_VALUE"),
                List.of("assertion"))));
  }

  private static TaskDefinition customXmlWorkflow() {
    TaskEntry task =
        task(
            "CUSTOM_XML_WORKFLOW",
            "Inspect, export, and import existing workbook custom-XML mappings on mapped .xlsx"
                + " files.",
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
            List.of(
                phase(
                    "Discover Existing Mappings",
                    "Read the workbook mapping metadata before attempting export or import.",
                    List.of(
                        ref("sourceTypes", "EXISTING"),
                        ref("inspectionQueryTypes", "GET_CUSTOM_XML_MAPPINGS"),
                        ref("inspectionQueryTypes", "EXPORT_CUSTOM_XML_MAPPING")),
                    List.of("Map discovery tells you which locator to target.")),
                phase(
                    "Import Updated XML",
                    "Push one XML payload through the chosen existing mapping.",
                    List.of(
                        ref("mutationActionTypes", "IMPORT_CUSTOM_XML_MAPPING"),
                        ref("persistenceTypes", "SAVE_AS")),
                    List.of("Use SAVE_AS when you want to preserve the original workbook.")),
                phase(
                    "Verify The Result",
                    "Re-export or reread workbook facts after the import.",
                    List.of(ref("inspectionQueryTypes", "EXPORT_CUSTOM_XML_MAPPING")),
                    List.of("Post-import export is the fastest proof that the mapping changed."))),
            List.of(
                "IMPORT_CUSTOM_XML_MAPPING requires an existing mapping; discovery comes first.",
                "Large XML payloads belong in UTF8_FILE or STANDARD_INPUT sources instead of huge"
                    + " inline JSON.",
                "Mapped workbook structure is workbook-specific, so replace the placeholder"
                    + " mapping locator before execution."));
    return definition(
        task,
        List.of(
            "custom xml mapping",
            "xml import export",
            "mapped workbook",
            "xml mapped table",
            "import xml into xlsx"),
        ExamplePlanSupport.defaultExecutionPlan(
            "custom-xml-workflow-starter",
            new WorkbookPlan.WorkbookSource.ExistingFile("mapped-workbook.xlsx"),
            ExamplePlanSupport.saveAs("custom-xml-output.xlsx"),
            ExamplePlanSupport.read(
                "discover-mappings",
                ExamplePlanSupport.workbook(),
                new InspectionQuery.GetCustomXmlMappings())),
        List.of(
            "Replace mapped-workbook.xlsx with the mapped workbook you want to inspect.",
            "Run discover-mappings first, then add EXPORT_CUSTOM_XML_MAPPING and"
                + " IMPORT_CUSTOM_XML_MAPPING steps with the exact locator returned by that"
                + " discovery output.",
            "The starter plan uses SAVE_AS so XML imports do not have to overwrite the source"
                + " workbook.",
            protocolLookupNote(
                "mapping discovery, export, and import shapes",
                List.of(
                    "inspectionQueryTypes:GET_CUSTOM_XML_MAPPINGS",
                    "mutationActionTypes:IMPORT_CUSTOM_XML_MAPPING",
                    "inspectionQueryTypes:EXPORT_CUSTOM_XML_MAPPING"),
                List.of("custom xml"))));
  }

  private static TaskDefinition drawingAndSignatureWorkflow() {
    TaskEntry task =
        task(
            "DRAWING_AND_SIGNATURE_WORKFLOW",
            "Author or inspect drawing-backed workbook content such as signature lines, pictures,"
                + " shapes, and embedded objects.",
            List.of("office", "drawing", "signature", "picture", "embedded-object"),
            List.of(
                "Drawing-backed workbook objects are authored by name instead of by fragile XML"
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
            List.of(
                phase(
                    "Prepare The Drawing Surface",
                    "Create the target sheet before any named drawing object is authored.",
                    List.of(ref("sourceTypes", "NEW"), ref("mutationActionTypes", "ENSURE_SHEET")),
                    List.of("A stable sheet name gives drawing objects a durable home.")),
                phase(
                    "Author Drawing Objects",
                    "Create one named drawing object and move it authoritatively when needed.",
                    List.of(
                        ref("mutationActionTypes", "SET_SIGNATURE_LINE"),
                        ref("mutationActionTypes", "SET_DRAWING_OBJECT_ANCHOR"),
                        ref("persistenceTypes", "SAVE_AS")),
                    List.of("Named drawing objects are easier to inspect and move later.")),
                phase(
                    "Inspect The Result",
                    "Read back worksheet drawing metadata after authoring or anchor changes.",
                    List.of(ref("inspectionQueryTypes", "GET_DRAWING_OBJECTS")),
                    List.of("Use factual reread to confirm the visible workbook surface."))),
            List.of(
                "Signature lines are VML-backed drawing objects with workbook-side constraints.",
                "Invalid drawing payloads are rejected without leaving partial artifacts behind.",
                "Image-heavy workflows should prefer file-backed binary sources over huge inline"
                    + " base64 payloads."));
    return definition(
        task,
        List.of(
            "signature line",
            "drawing objects",
            "picture workflow",
            "embedded object",
            "move drawing anchor"),
        rebasePlan(
            WorkbookAssetExamples.signatureLineExample(ExamplePathLayout.BUILT_IN).plan(),
            "drawing-and-signature-starter",
            new WorkbookPlan.WorkbookSource.New(),
            ExamplePlanSupport.saveAs("signature-workflow.xlsx")),
        List.of(
            "The starter plan is runnable and demonstrates named signature-line authoring plus an"
                + " explicit anchor move.",
            "Replace the signer metadata, anchor coordinates, and output filename with your own"
                + " workbook flow.",
            protocolLookupNote(
                "drawing-authoring shapes",
                List.of(
                    "mutationActionTypes:SET_SIGNATURE_LINE",
                    "mutationActionTypes:SET_DRAWING_OBJECT_ANCHOR"),
                List.of("drawing", "signature line"))));
  }

  private static TaskDefinition workbookMaintenance() {
    TaskEntry task =
        task(
            "WORKBOOK_MAINTENANCE",
            "Safely copy or normalize existing workbook sheets while verifying comments,"
                + " drawings, and other cloned workbook state after the change.",
            List.of("office", "maintenance", "repair", "copy", "comments"),
            List.of(
                "Workbook maintenance flows can inspect the source before making structural"
                    + " changes.",
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
            List.of(
                phase(
                    "Inspect The Current Workbook",
                    "Read source workbook facts before changing sheet structure.",
                    List.of(
                        ref("sourceTypes", "EXISTING"),
                        ref("inspectionQueryTypes", "GET_WORKBOOK_SUMMARY"),
                        ref("inspectionQueryTypes", "GET_COMMENTS")),
                    List.of("Factual inspection is the baseline for any maintenance pass.")),
                phase(
                    "Apply Structural Maintenance",
                    "Copy one existing sheet into a new workbook-visible destination sheet.",
                    List.of(
                        ref("mutationActionTypes", "COPY_SHEET"),
                        ref("persistenceTypes", "SAVE_AS")),
                    List.of("Copy-sheet work is safest when the destination name is explicit.")),
                phase(
                    "Verify The Copied Surface",
                    "Read back copied comments, drawings, and workbook findings immediately.",
                    List.of(
                        ref("inspectionQueryTypes", "GET_COMMENTS"),
                        ref("inspectionQueryTypes", "GET_DRAWING_OBJECTS"),
                        ref("inspectionQueryTypes", "ANALYZE_WORKBOOK_FINDINGS")),
                    List.of("Verification closes the loop on maintenance changes."))),
            List.of(
                "Replace the placeholder sheet names before execution; copy-sheet requires an"
                    + " existing source sheet.",
                "Maintenance flows are safest with SAVE_AS so the source workbook stays intact"
                    + " until verification passes.",
                "Use targeted rereads after COPY_SHEET instead of assuming copied comments or"
                    + " drawings stayed coherent."));
    return definition(
        task,
        List.of(
            "repair broken workbook comments",
            "copy sheets safely",
            "sheet maintenance",
            "comment repair",
            "clone worksheet with comments"),
        ExamplePlanSupport.defaultExecutionPlan(
            "workbook-maintenance-starter",
            new WorkbookPlan.WorkbookSource.ExistingFile("maintenance-input.xlsx"),
            ExamplePlanSupport.saveAs("maintenance-output.xlsx"),
            ExamplePlanSupport.read(
                "workbook",
                ExamplePlanSupport.workbook(),
                new InspectionQuery.GetWorkbookSummary()),
            ExamplePlanSupport.read(
                "comments-before-copy",
                new CellSelector.AllUsedInSheet("Template"),
                new InspectionQuery.GetComments()),
            ExamplePlanSupport.step(
                "copy-template",
                new SheetSelector.ByName("Template"),
                new WorkbookMutationAction.CopySheet("Template Copy")),
            ExamplePlanSupport.read(
                "comments-after-copy",
                new CellSelector.AllUsedInSheet("Template Copy"),
                new InspectionQuery.GetComments()),
            ExamplePlanSupport.read(
                "drawings-after-copy",
                new DrawingObjectSelector.AllOnSheet("Template Copy"),
                new InspectionQuery.GetDrawingObjects()),
            ExamplePlanSupport.read(
                "workbook-findings",
                ExamplePlanSupport.workbook(),
                new InspectionQuery.AnalyzeWorkbookFindings())),
        List.of(
            "Replace maintenance-input.xlsx and the placeholder Template sheet names with the"
                + " workbook and sheet you actually need to maintain.",
            "The starter plan uses SAVE_AS so you can verify copied comments and drawings before"
                + " deciding whether to overwrite the source workbook.",
            protocolLookupNote(
                "maintenance and verification shapes",
                List.of("mutationActionTypes:COPY_SHEET", "inspectionQueryTypes:GET_COMMENTS"),
                List.of("copy sheet", "comments"))));
  }
}
