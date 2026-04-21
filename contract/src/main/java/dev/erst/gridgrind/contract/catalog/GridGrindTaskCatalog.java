package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import java.util.List;
import java.util.Optional;

/** High-level task/intention catalog layered on top of the exact protocol surface. */
public final class GridGrindTaskCatalog {
  private static final TaskCatalog CATALOG = buildCatalog();

  private GridGrindTaskCatalog() {}

  /** Returns the full task catalog. */
  public static TaskCatalog catalog() {
    return CATALOG;
  }

  /** Returns one task entry by its stable id, or empty when unknown. */
  public static Optional<TaskEntry> entryFor(String id) {
    String lookup = CatalogRecordValidation.requireNonBlank(id, "id").trim();
    return CATALOG.tasks().stream().filter(task -> task.id().equals(lookup)).findFirst();
  }

  private static TaskCatalog buildCatalog() {
    TaskCatalog catalog =
        new TaskCatalog(
            GridGrindProtocolVersion.current(),
            List.of(
                task(
                    "TABULAR_REPORT",
                    "Create a structured worksheet report with typed cells, table semantics, and"
                        + " factual or analytical readback.",
                    List.of("office", "reporting", "tables", "formulas", "save"),
                    List.of(
                        "Sheet structure is created intentionally instead of ad hoc cell drift.",
                        "Rows can be modeled as a table so filtering, totals, and readback stay"
                            + " authoritative.",
                        "Critical facts can be asserted or inspected before the workbook is"
                            + " persisted."),
                    List.of(
                        "Sheet names and header layout.",
                        "Typed row values or source-backed payload files for larger authored"
                            + " content.",
                        "Persistence target when the result must be saved."),
                    List.of(
                        "Totals rows and formula-backed summaries.",
                        "Cell styling or print-layout refinements.",
                        "Assertions on key balances or counts."),
                    List.of(
                        phase(
                            "Lay Out The Workbook",
                            "Create sheets, headers, and fixed structure before data rows arrive.",
                            List.of(
                                ref("sourceTypes", "NEW"),
                                ref("mutationActionTypes", "ENSURE_SHEET"),
                                ref("mutationActionTypes", "SET_CELL"),
                                ref("mutationActionTypes", "SET_RANGE")),
                            List.of(
                                "Use one intentional sheet skeleton instead of scattered cell"
                                    + " writes.")),
                        phase(
                            "Model The Table",
                            "Move the sheet from loose cells into tabular semantics.",
                            List.of(
                                ref("mutationActionTypes", "SET_TABLE"),
                                ref("mutationActionTypes", "APPEND_ROW"),
                                ref("persistenceTypes", "SAVE_AS")),
                            List.of(
                                "Prefer one workbook-global table definition per logical data"
                                    + " region.")),
                        phase(
                            "Verify And Inspect",
                            "Read back or assert the facts that make the report trustworthy.",
                            List.of(
                                ref("assertionTypes", "EXPECT_CELL_VALUE"),
                                ref("inspectionQueryTypes", "GET_CELLS"),
                                ref("inspectionQueryTypes", "ANALYZE_WORKBOOK_FINDINGS")),
                            List.of(
                                "Use assertions for invariants and inspections for factual"
                                    + " payloads."))),
                    List.of(
                        "Large authored literals belong in UTF8_FILE, FILE, or STANDARD_INPUT"
                            + " sources instead of huge inline JSON.",
                        "Table headers must remain nonblank and unique.",
                        "Formula authoring is scalar-only; newer Excel constructs may still be"
                            + " parser-limited.")),
                task(
                    "DASHBOARD",
                    "Assemble an executive dashboard from tables, named ranges, and supported"
                        + " simple charts.",
                    List.of("office", "dashboard", "charts", "summary", "kpi"),
                    List.of(
                        "Summary sheets and KPI surfaces are intentionally structured.",
                        "Reusable named surfaces back formulas or charts instead of fragile"
                            + " copied ranges.",
                        "Charts are authored and then read back through the same contract"
                            + " surface."),
                    List.of(
                        "Metric definitions and source ranges.",
                        "Chart names and anchors.",
                        "Target persistence path when the dashboard must be saved."),
                    List.of(
                        "Assertions that required dashboard entities exist.",
                        "Workbook-health analysis after authoring.",
                        "Named-range-backed chart series for reusable models."),
                    List.of(
                        phase(
                            "Assemble Summary Sheets",
                            "Create the dashboard canvas and key text or formula cells first.",
                            List.of(
                                ref("sourceTypes", "NEW"),
                                ref("mutationActionTypes", "ENSURE_SHEET"),
                                ref("mutationActionTypes", "SET_CELL"),
                                ref("mutationActionTypes", "SET_RANGE")),
                            List.of(
                                "Keep summary layout intentional so later chart anchors are"
                                    + " stable.")),
                        phase(
                            "Define Reusable Model Surfaces",
                            "Create tables or named ranges that charts and formulas can depend on.",
                            List.of(
                                ref("mutationActionTypes", "SET_TABLE"),
                                ref("mutationActionTypes", "SET_NAMED_RANGE")),
                            List.of(
                                "Named surfaces reduce accidental drift when the dashboard"
                                    + " evolves.")),
                        phase(
                            "Author And Inspect Visuals",
                            "Create supported charts and verify that the expected entities exist.",
                            List.of(
                                ref("mutationActionTypes", "SET_CHART"),
                                ref("inspectionQueryTypes", "GET_CHARTS"),
                                ref("assertionTypes", "EXPECT_PRESENT"),
                                ref("persistenceTypes", "SAVE_AS")),
                            List.of(
                                "Use factual chart readback to confirm what the workbook now"
                                    + " contains."))),
                    List.of(
                        "SET_CHART is intentionally limited to BAR, LINE, and PIE in the current"
                            + " authoritative mutation contract.",
                        "Chart title and series FORMULA titles must resolve to one cell.",
                        "Unsupported loaded chart detail is preserved on unrelated edits but is not"
                            + " available for authoritative mutation.")),
                task(
                    "AUDIT_EXISTING_WORKBOOK",
                    "Inspect an existing workbook, surface health findings, and avoid mutation"
                        + " unless a follow-up workflow explicitly asks for it.",
                    List.of("office", "audit", "analysis", "readback", "safety"),
                    List.of(
                        "Existing workbook structure is surfaced as facts instead of guesses.",
                        "Health analyses aggregate into one workbook-level findings view.",
                        "The default audit posture can stay read-only and non-persistent."),
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
                            "Start with workbook-level facts and security state before diving into"
                                + " detailed sheet logic.",
                            List.of(
                                ref("sourceTypes", "EXISTING"),
                                ref("persistenceTypes", "NONE"),
                                ref("inspectionQueryTypes", "GET_WORKBOOK_SUMMARY"),
                                ref("inspectionQueryTypes", "GET_PACKAGE_SECURITY")),
                            List.of(
                                "A no-persistence audit pass reduces the chance of accidental"
                                    + " mutation.")),
                        phase(
                            "Analyze The Workbook",
                            "Read formula surfaces and aggregate workbook findings into one"
                                + " operator-facing view.",
                            List.of(
                                ref("inspectionQueryTypes", "GET_FORMULA_SURFACE"),
                                ref("inspectionQueryTypes", "ANALYZE_WORKBOOK_FINDINGS")),
                            List.of(
                                "Use targeted factual reads first when a later aggregate finding"
                                    + " needs explanation."))),
                    List.of(
                        "EVENT_READ is intentionally limited; many audit plans still require the"
                            + " default FULL_XSSF path.",
                        "Loaded formulas that Apache POI cannot evaluate surface as"
                            + " UNSUPPORTED_FORMULA rather than silently recalculating.",
                        "If the audit needs strict policy checks, promote the findings into"
                            + " ASSERTION steps in a follow-up plan.")),
                task(
                    "DATA_ENTRY_WORKFLOW",
                    "Build an intake-style sheet for operators, then load rows and validate key"
                        + " fields before persistence.",
                    List.of("office", "data-entry", "validation", "append-only", "workflow"),
                    List.of(
                        "One intake sheet or workbook is structured before row ingestion begins.",
                        "Business rules can be expressed through validation and typed writes.",
                        "Key fields are read back before the workbook is saved."),
                    List.of(
                        "Schema or header layout.",
                        "Row payload source.",
                        "Persistence target when the workbook should leave memory."),
                    List.of(
                        "Append-only large-file plans.",
                        "Validation rules on business-critical columns.",
                        "Post-write factual readback for spot checks."),
                    List.of(
                        phase(
                            "Prepare The Intake Surface",
                            "Create the intake sheet and attach the first layer of operator"
                                + " guidance or validation.",
                            List.of(
                                ref("sourceTypes", "NEW"),
                                ref("mutationActionTypes", "ENSURE_SHEET"),
                                ref("mutationActionTypes", "SET_CELL"),
                                ref("mutationActionTypes", "SET_DATA_VALIDATION")),
                            List.of("Stabilize the schema before appending many rows.")),
                        phase(
                            "Load Rows",
                            "Append data in a way that keeps the workflow readable and"
                                + " deterministic.",
                            List.of(
                                ref("mutationActionTypes", "APPEND_ROW"),
                                ref("persistenceTypes", "SAVE_AS")),
                            List.of(
                                "For large append-only workflows, evaluate whether the execution"
                                    + " mode limits fit STREAMING_WRITE.")),
                        phase(
                            "Spot-Check The Result",
                            "Read back critical cells instead of trusting write success alone.",
                            List.of(ref("inspectionQueryTypes", "GET_CELLS")),
                            List.of(
                                "Read back the fields that operators would inspect manually if"
                                    + " this were not automated."))),
                    List.of(
                        "STREAMING_WRITE is append-only and requires a NEW source with constrained"
                            + " mutation support.",
                        "STANDARD_INPUT-authored values require --request because stdin cannot"
                            + " carry both the request JSON and authored payload content.",
                        "Data validation rules can enforce operator guidance, but they do not"
                            + " replace factual readback."))));
    validateCapabilityReferences(catalog);
    return catalog;
  }

  /** Validates that every task capability reference resolves against the protocol catalog. */
  static void validateCapabilityReferences(TaskCatalog catalog) {
    for (TaskEntry task : catalog.tasks()) {
      for (TaskPhase phase : task.phases()) {
        for (TaskCapabilityRef capabilityRef : phase.capabilityRefs()) {
          if (GridGrindProtocolCatalog.entryFor(capabilityRef.qualifiedId()).isEmpty()) {
            throw new IllegalStateException(
                "Task "
                    + task.id()
                    + " references unknown protocol capability "
                    + capabilityRef.qualifiedId());
          }
        }
      }
    }
  }

  private static TaskEntry task(
      String id,
      String summary,
      List<String> intentTags,
      List<String> outcomes,
      List<String> requiredInputs,
      List<String> optionalFeatures,
      List<TaskPhase> phases,
      List<String> commonPitfalls) {
    return new TaskEntry(
        id,
        summary,
        intentTags,
        outcomes,
        requiredInputs,
        optionalFeatures,
        phases,
        commonPitfalls);
  }

  private static TaskPhase phase(
      String label, String objective, List<TaskCapabilityRef> capabilityRefs, List<String> notes) {
    return new TaskPhase(label, objective, capabilityRefs, notes);
  }

  private static TaskCapabilityRef ref(String group, String id) {
    return new TaskCapabilityRef(group, id);
  }
}
