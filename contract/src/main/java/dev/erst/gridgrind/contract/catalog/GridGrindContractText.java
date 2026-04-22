package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

/**
 * Core-owned public contract wording shared by thin downstream surfaces such as the CLI help,
 * protocol catalog summaries, and request-validation messages.
 */
public final class GridGrindContractText {
  private static final long REQUEST_DOCUMENT_LIMIT_BYTES = 16L * 1024 * 1024;
  private static final List<Class<? extends MutationAction>>
      STREAMING_WRITE_MUTATION_ACTION_CLASSES =
          List.of(MutationAction.EnsureSheet.class, MutationAction.AppendRow.class);
  private static final List<Class<? extends InspectionQuery>> EVENT_READ_INSPECTION_QUERY_CLASSES =
      List.of(InspectionQuery.GetWorkbookSummary.class, InspectionQuery.GetSheetSummary.class);
  private static final List<String> WORKBOOK_ANALYSIS_FAMILIES =
      List.of(
          "formula health",
          "data-validation health",
          "conditional-formatting health",
          "autofilter health",
          "table health",
          "pivot-table health",
          "hyperlink health",
          "named-range health");

  private GridGrindContractText() {}

  /** Ordered protocol mutation-action classes accepted by `STREAMING_WRITE`. */
  public static Set<Class<? extends MutationAction>> streamingWriteMutationActionClasses() {
    return Set.copyOf(STREAMING_WRITE_MUTATION_ACTION_CLASSES);
  }

  /** Ordered protocol inspection-query classes accepted by `EVENT_READ`. */
  public static Set<Class<? extends InspectionQuery>> eventReadInspectionQueryClasses() {
    return Set.copyOf(EVENT_READ_INSPECTION_QUERY_CLASSES);
  }

  /** Human-readable mutation-action id list accepted by `STREAMING_WRITE`. */
  public static String streamingWriteMutationActionTypePhrase() {
    return GridGrindExecutionModeMetadata.streamingWrite().allowedActionPhrase();
  }

  /** Human-readable inspection-query id list accepted by `EVENT_READ`. */
  public static String eventReadInspectionQueryTypePhrase() {
    return GridGrindExecutionModeMetadata.eventRead().allowedQueryPhrase();
  }

  /** Human-readable aggregate analysis-family list used by workbook-health discovery surfaces. */
  public static String workbookAnalysisFamilyPhrase() {
    return humanJoin(WORKBOOK_ANALYSIS_FAMILIES);
  }

  /** One stable description of request-authored formula boundaries. */
  public static String formulaAuthoringLimitSummary() {
    return "request-authored formulas are scalar only; array-formula braces such as"
        + " {=SUM(A1:A2*B1:B2)} are rejected as INVALID_FORMULA, and LAMBDA/LET are currently"
        + " rejected as INVALID_FORMULA because Apache POI cannot parse them."
        + " Other newer constructs may fail the same way.";
  }

  /** One stable description of loaded-formula evaluation boundaries. */
  public static String loadedFormulaSupportSummary() {
    return "formulas that Apache POI parses but cannot evaluate surface as"
        + " UNSUPPORTED_FORMULA.";
  }

  /** One stable catalog summary for `GET_SHEET_LAYOUT`. */
  public static String sheetLayoutReadSummary() {
    return "Return one sheet's layout object with pane, zoomPercent, presentation,"
        + " and per-row or per-column metadata."
        + " Row and column entries include hidden, outlineLevel, and collapsed"
        + " state where Excel persists it."
        + " Readback is factual and does not clamp malformed positive persisted"
        + " row heights, column widths, or default row height values.";
  }

  /** One stable catalog summary for `GET_FORMULA_SURFACE`. */
  public static String formulaSurfaceReadSummary() {
    return "Return analysis.totalFormulaCellCount plus per-sheet formula usage groups."
        + " analysis.sheets[*] includes sheetName, formulaCellCount,"
        + " distinctFormulaCount, and grouped formulas with occurrenceCount"
        + " and addresses.";
  }

  /** One stable catalog summary for `GET_NAMED_RANGE_SURFACE`. */
  public static String namedRangeSurfaceReadSummary() {
    return "Return analysis.workbookScopedCount, sheetScopedCount,"
        + " rangeBackedCount, formulaBackedCount, and namedRanges."
        + " Each namedRanges entry reports name, scope, refersToFormula,"
        + " and backing kind.";
  }

  /** One stable catalog summary for `ANALYZE_FORMULA_HEALTH`. */
  public static String formulaHealthReadSummary() {
    return "Return analysis.checkedFormulaCellCount, a severity summary,"
        + " and findings for formula errors, volatile usage,"
        + " or evaluation failures.";
  }

  /** One stable catalog summary for `ANALYZE_NAMED_RANGE_HEALTH`. */
  public static String namedRangeHealthReadSummary() {
    return "Return analysis.checkedNamedRangeCount, a severity summary,"
        + " and named-range findings such as broken references,"
        + " unresolved targets, or scope shadowing.";
  }

  /** One stable catalog summary for `ANALYZE_WORKBOOK_FINDINGS`. */
  public static String workbookFindingsReadSummary() {
    return "Return analysis.summary plus one flat analysis.findings list after running"
        + " all analysis families ("
        + workbookAnalysisFamilyPhrase()
        + ") across the entire workbook and aggregate findings in a single response."
        + " This is the primary workbook-health check and pairs naturally with"
        + " persistence.type=NONE when no save is required.";
  }

  /** One stable catalog summary for `ExecutionModeInput`. */
  public static String executionModeInputSummary() {
    GridGrindExecutionModeMetadata.EventReadMode eventRead =
        GridGrindExecutionModeMetadata.eventRead();
    GridGrindExecutionModeMetadata.StreamingWriteMode streamingWrite =
        GridGrindExecutionModeMetadata.streamingWrite();
    return "Execution-mode settings that select low-memory read and write"
        + " execution families."
        + " readMode defaults to FULL_XSSF when omitted."
        + " writeMode defaults to FULL_XSSF when omitted."
        + " "
        + eventRead.mode().name()
        + " supports "
        + eventRead.allowedQueryPhrase()
        + " only and requires execution.calculation.strategy="
        + eventRead.requiredCalculationStrategyId()
        + " with markRecalculateOnOpen="
        + eventRead.markRecalculateOnOpenAllowed()
        + " (LIM-019)."
        + " "
        + streamingWrite.mode().name()
        + " supports "
        + streamingWrite.allowedActionPhrase()
        + " on "
        + streamingWrite.requiredSourceTypeId()
        + " workbooks only;"
        + " execution.calculation may only keep strategy="
        + streamingWrite.requiredCalculationStrategyId()
        + " and optionally set markRecalculateOnOpen="
        + streamingWrite.markRecalculateOnOpenAllowed()
        + " (LIM-020).";
  }

  /** One stable catalog summary for `ExecutionPolicyInput`. */
  public static String executionPolicyInputSummary() {
    return "Optional request execution policy covering execution.mode, execution.journal,"
        + " and execution.calculation."
        + " Omit it to accept the default FULL_XSSF read/write path, NORMAL journal detail,"
        + " and DO_NOT_CALCULATE calculation policy.";
  }

  /** One stable catalog summary for `ExecutionJournalInput`. */
  public static String executionJournalInputSummary() {
    return "Optional execution-journal settings."
        + " level defaults to NORMAL when omitted."
        + " SUMMARY returns compact target summaries,"
        + " NORMAL returns expanded target summaries,"
        + " and VERBOSE also enables fine-grained event emission for CLI stderr rendering.";
  }

  /** One stable catalog summary for `CalculationPolicyInput`. */
  public static String calculationPolicyInputSummary() {
    return "Optional explicit formula-calculation policy."
        + " strategy defaults to DO_NOT_CALCULATE when omitted."
        + " markRecalculateOnOpen defaults to false when omitted."
        + " Use EVALUATE_ALL or EVALUATE_TARGETS for immediate server-side evaluation,"
        + " CLEAR_CACHES_ONLY to strip persisted caches,"
        + " or markRecalculateOnOpen=true when Excel-compatible clients should recalculate later.";
  }

  /** One stable catalog summary for `CalculationStrategyInput`. */
  public static String calculationStrategyInputSummary() {
    return "One explicit formula-calculation strategy."
        + " DO_NOT_CALCULATE leaves formulas untouched,"
        + " EVALUATE_ALL evaluates every formula cell,"
        + " EVALUATE_TARGETS evaluates one explicit formula-cell set,"
        + " and CLEAR_CACHES_ONLY strips persisted formula caches without evaluating.";
  }

  /** One stable discovery line for `ANALYZE_WORKBOOK_FINDINGS`. */
  public static String workbookFindingsDiscoverySummary() {
    return "ANALYZE_WORKBOOK_FINDINGS aggregates " + workbookAnalysisFamilyPhrase() + ".";
  }

  /** One stable help and runtime message for stdin-backed authored values. */
  public static String standardInputRequiresRequestMessage() {
    return "STANDARD_INPUT-authored values require --request so stdin is available for input"
        + " content instead of the request JSON";
  }

  /** Maximum accepted JSON request document size in bytes. */
  public static long requestDocumentLimitBytes() {
    return REQUEST_DOCUMENT_LIMIT_BYTES;
  }

  /** Human-readable summary of the canonical JSON request document limit. */
  public static String requestDocumentLimitSummary() {
    return "request JSON must not exceed 16 MiB ("
        + REQUEST_DOCUMENT_LIMIT_BYTES
        + " bytes); use UTF8_FILE, FILE, or STANDARD_INPUT sources for large authored payloads.";
  }

  /** One stable product-owned message for oversized JSON request payloads. */
  public static String requestDocumentTooLargeMessage() {
    return "Request JSON exceeds the maximum size of 16 MiB ("
        + REQUEST_DOCUMENT_LIMIT_BYTES
        + " bytes); move large authored payloads into UTF8_FILE, FILE, or STANDARD_INPUT"
        + " sources.";
  }

  /** One stable step-kind explanation shared by help and discovery surfaces. */
  public static String stepKindSummary() {
    return "Use MUTATION steps for workbook changes, ASSERTION steps for first-class"
        + " verification, and INSPECTION steps for factual or analytical reads."
        + " Step kind is inferred from exactly one of action, assertion, or query;"
        + " do not send step.type.";
  }

  /** Stable mutation-action discriminator lookup by protocol subtype class. */
  public static String mutationActionTypeName(Class<? extends MutationAction> mutationActionClass) {
    return GridGrindProtocolTypeNames.mutationActionTypeName(mutationActionClass);
  }

  /** Stable inspection-query discriminator lookup by protocol subtype class. */
  public static String inspectionQueryTypeName(
      Class<? extends InspectionQuery> inspectionQueryClass) {
    return GridGrindProtocolTypeNames.inspectionQueryTypeName(inspectionQueryClass);
  }

  static Map<Class<?>, String> typeNamesByClass(Class<?> rootType) {
    return GridGrindProtocolTypeNames.typeNamesByClass(rootType);
  }

  static String humanJoin(List<String> values) {
    List<String> parts = values.stream().filter(value -> !value.isBlank()).toList();
    if (parts.isEmpty()) {
      throw new IllegalArgumentException("values must not be empty");
    }
    if (parts.size() == 1) {
      return parts.getFirst();
    }
    if (parts.size() == 2) {
      return parts.getFirst() + " and " + parts.getLast();
    }
    return parts.subList(0, parts.size() - 1).stream().collect(Collectors.joining(", "))
        + ", and "
        + parts.getLast();
  }
}
