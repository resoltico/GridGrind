package dev.erst.gridgrind.contract.catalog;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.stream.Collectors;

/**
 * Core-owned public contract wording shared by thin downstream surfaces such as the CLI help,
 * protocol catalog summaries, and request-validation messages.
 */
public final class GridGrindContractText {
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
  private static final Map<Class<?>, String> MUTATION_ACTION_TYPE_NAMES =
      typeNamesByClass(MutationAction.class);
  private static final Map<Class<?>, String> INSPECTION_QUERY_TYPE_NAMES =
      typeNamesByClass(InspectionQuery.class);

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
    return humanJoin(
        STREAMING_WRITE_MUTATION_ACTION_CLASSES.stream()
            .map(GridGrindContractText::mutationActionTypeName)
            .toList());
  }

  /** Human-readable inspection-query id list accepted by `EVENT_READ`. */
  public static String eventReadInspectionQueryTypePhrase() {
    return humanJoin(
        EVENT_READ_INSPECTION_QUERY_CLASSES.stream()
            .map(GridGrindContractText::inspectionQueryTypeName)
            .toList());
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
        + " state where Excel persists it.";
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
    return "Execution-mode settings that select low-memory read and write"
        + " execution families."
        + " readMode defaults to FULL_XSSF when omitted."
        + " writeMode defaults to FULL_XSSF when omitted."
        + " EVENT_READ supports "
        + eventReadInspectionQueryTypePhrase()
        + " only and requires execution.calculation.strategy=DO_NOT_CALCULATE with"
        + " markRecalculateOnOpen=false (LIM-019)."
        + " STREAMING_WRITE supports "
        + streamingWriteMutationActionTypePhrase()
        + " on NEW workbooks only;"
        + " execution.calculation may only keep strategy=DO_NOT_CALCULATE and optionally set"
        + " markRecalculateOnOpen=true (LIM-020).";
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

  /** One stable step-kind explanation shared by help and discovery surfaces. */
  public static String stepKindSummary() {
    return "Use MUTATION steps for workbook changes, ASSERTION steps for first-class"
        + " verification, and INSPECTION steps for factual or analytical reads.";
  }

  /** Stable mutation-action discriminator lookup by protocol subtype class. */
  public static String mutationActionTypeName(Class<? extends MutationAction> mutationActionClass) {
    return requiredTypeName(MUTATION_ACTION_TYPE_NAMES, mutationActionClass, "MutationAction");
  }

  /** Stable inspection-query discriminator lookup by protocol subtype class. */
  public static String inspectionQueryTypeName(
      Class<? extends InspectionQuery> inspectionQueryClass) {
    return requiredTypeName(INSPECTION_QUERY_TYPE_NAMES, inspectionQueryClass, "InspectionQuery");
  }

  private static String requiredTypeName(
      Map<Class<?>, String> typeNamesByClass, Class<?> typeClass, String rootTypeName) {
    Objects.requireNonNull(typeClass, "typeClass must not be null");
    String typeName = typeNamesByClass.get(typeClass);
    if (typeName == null) {
      throw new IllegalArgumentException(
          "No discriminator name registered for " + rootTypeName + " subtype " + typeClass);
    }
    return typeName;
  }

  static Map<Class<?>, String> typeNamesByClass(Class<?> rootType) {
    JsonSubTypes jsonSubTypes = rootType.getAnnotation(JsonSubTypes.class);
    if (jsonSubTypes == null) {
      throw new IllegalArgumentException(rootType + " is missing @JsonSubTypes");
    }
    return Arrays.stream(jsonSubTypes.value())
        .collect(Collectors.toUnmodifiableMap(JsonSubTypes.Type::value, JsonSubTypes.Type::name));
  }

  private static String humanJoin(List<String> values) {
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
