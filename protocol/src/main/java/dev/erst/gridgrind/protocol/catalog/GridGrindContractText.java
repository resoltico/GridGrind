package dev.erst.gridgrind.protocol.catalog;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import dev.erst.gridgrind.protocol.operation.WorkbookOperation;
import dev.erst.gridgrind.protocol.read.WorkbookReadOperation;
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
  private static final List<Class<? extends WorkbookOperation>> STREAMING_WRITE_OPERATION_CLASSES =
      List.of(
          WorkbookOperation.EnsureSheet.class,
          WorkbookOperation.AppendRow.class,
          WorkbookOperation.ForceFormulaRecalculationOnOpen.class);
  private static final List<Class<? extends WorkbookReadOperation>> EVENT_READ_OPERATION_CLASSES =
      List.of(
          WorkbookReadOperation.GetWorkbookSummary.class,
          WorkbookReadOperation.GetSheetSummary.class);
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
  private static final Map<Class<?>, String> OPERATION_TYPE_NAMES =
      typeNamesByClass(WorkbookOperation.class);
  private static final Map<Class<?>, String> READ_TYPE_NAMES =
      typeNamesByClass(WorkbookReadOperation.class);

  private GridGrindContractText() {}

  /** Ordered protocol operation classes accepted by `STREAMING_WRITE`. */
  public static Set<Class<? extends WorkbookOperation>> streamingWriteOperationClasses() {
    return Set.copyOf(STREAMING_WRITE_OPERATION_CLASSES);
  }

  /** Ordered protocol read classes accepted by `EVENT_READ`. */
  public static Set<Class<? extends WorkbookReadOperation>> eventReadOperationClasses() {
    return Set.copyOf(EVENT_READ_OPERATION_CLASSES);
  }

  /** Human-readable operation-id list accepted by `STREAMING_WRITE`. */
  public static String streamingWriteOperationTypePhrase() {
    return humanJoin(
        STREAMING_WRITE_OPERATION_CLASSES.stream()
            .map(GridGrindContractText::operationTypeName)
            .toList());
  }

  /** Human-readable read-id list accepted by `EVENT_READ`. */
  public static String eventReadReadTypePhrase() {
    return humanJoin(
        EVENT_READ_OPERATION_CLASSES.stream().map(GridGrindContractText::readTypeName).toList());
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
    return "Optional top-level request settings that select low-memory read and write"
        + " execution families."
        + " readMode defaults to FULL_XSSF when omitted."
        + " writeMode defaults to FULL_XSSF when omitted."
        + " EVENT_READ supports "
        + eventReadReadTypePhrase()
        + " only (LIM-019)."
        + " STREAMING_WRITE supports "
        + streamingWriteOperationTypePhrase()
        + " on NEW workbooks only (LIM-020).";
  }

  /** One stable discovery line for `ANALYZE_WORKBOOK_FINDINGS`. */
  public static String workbookFindingsDiscoverySummary() {
    return "ANALYZE_WORKBOOK_FINDINGS aggregates " + workbookAnalysisFamilyPhrase() + ".";
  }

  /** Stable operation discriminator lookup by protocol subtype class. */
  public static String operationTypeName(Class<? extends WorkbookOperation> operationClass) {
    return requiredTypeName(OPERATION_TYPE_NAMES, operationClass, "WorkbookOperation");
  }

  /** Stable read discriminator lookup by protocol subtype class. */
  public static String readTypeName(Class<? extends WorkbookReadOperation> readClass) {
    return requiredTypeName(READ_TYPE_NAMES, readClass, "WorkbookReadOperation");
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

  private static Map<Class<?>, String> typeNamesByClass(Class<?> rootType) {
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
