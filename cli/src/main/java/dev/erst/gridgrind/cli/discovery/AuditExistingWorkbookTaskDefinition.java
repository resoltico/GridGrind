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

/** Task descriptor for read-only workbook audit workflows. */
final class AuditExistingWorkbookTaskDefinition {
  private AuditExistingWorkbookTaskDefinition() {}

  static TaskEntry entry() {
    return task(
        "AUDIT_EXISTING_WORKBOOK",
        discovery(
            List.of("audit", "inspect", "health", "findings", "diagnose", "read only"),
            List.of("office", "audit", "analysis", "readback", "safety"),
            intent(
                List.of(TaskGoalKind.INSPECT, TaskGoalKind.ANALYZE, TaskGoalKind.VERIFY),
                List.of(
                    TaskArtifactKind.WORKBOOK,
                    TaskArtifactKind.PACKAGE_SECURITY,
                    TaskArtifactKind.FORMULA_SURFACE))),
        narrative(
            "Inspect an existing workbook, surface health findings, and avoid mutation unless a"
                + " follow-up workflow explicitly asks for it.",
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
                "Follow-up assertion plans after the first audit pass.")),
        profile(
            TaskSourceMode.EXISTING_WORKBOOK,
            TaskPersistenceMode.NONE,
            TaskMutationMode.READ_ONLY,
            TaskAssetMode.SELF_CONTAINED),
        signals(
            List.of(TaskInputKind.SOURCE_WORKBOOK_PATH, TaskInputKind.TARGET_SHEET_NAMES),
            List.of(TaskVerificationKind.FACT_READBACK, TaskVerificationKind.HEALTH_ANALYSIS)),
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
}
