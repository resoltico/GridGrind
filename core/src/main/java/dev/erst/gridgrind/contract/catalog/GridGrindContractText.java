package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.query.SheetIntrospectionQuery;
import dev.erst.gridgrind.contract.query.WorkbookIntrospectionQuery;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

/**
 * Core-owned public contract wording shared by thin downstream surfaces such as the CLI help,
 * protocol catalog summaries, and request-validation messages.
 */
public final class GridGrindContractText {
  private static final List<Class<? extends MutationAction>>
      STREAMING_WRITE_MUTATION_ACTION_CLASSES =
          List.of(WorkbookMutationAction.EnsureSheet.class, CellMutationAction.AppendRow.class);
  private static final List<Class<? extends InspectionQuery>> EVENT_READ_INSPECTION_QUERY_CLASSES =
      List.of(
          WorkbookIntrospectionQuery.GetWorkbookSummary.class,
          SheetIntrospectionQuery.GetSheetSummary.class);

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

  /** One stable catalog summary for `ExecutionModeInput`. */
  public static String executionModeInputSummary() {
    GridGrindExecutionModeMetadata.FullXssfMode fullXssf =
        GridGrindExecutionModeMetadata.fullXssf();
    GridGrindExecutionModeMetadata.EventReadMode eventRead =
        GridGrindExecutionModeMetadata.eventRead();
    GridGrindExecutionModeMetadata.StreamingWriteMode streamingWrite =
        GridGrindExecutionModeMetadata.streamingWrite();
    return "Execution-mode settings that select one runtime family by discriminator."
        + " "
        + fullXssf.modeId()
        + " keeps the standard full workbook read/write path."
        + " "
        + eventRead.modeId()
        + " supports "
        + eventRead.allowedQueryPhrase()
        + " only and requires execution.calculation.strategy="
        + eventRead.requiredCalculationStrategyId()
        + " with markRecalculateOnOpen="
        + eventRead.markRecalculateOnOpenAllowed()
        + " (LIM-019)."
        + " "
        + streamingWrite.modeId()
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
        + " execution.calculation, and execution.assertionMode."
        + " Omit execution when the standard FULL_XSSF / SUMMARY / DO_NOT_CALCULATE / FAIL_FAST policy"
        + " is intended, and omit any nested execution field that should keep its own default.";
  }

  /** One stable catalog summary for `FormulaEnvironmentInput`. */
  public static String formulaEnvironmentInputSummary() {
    return "Optional request-scoped formula-evaluation environment covering external workbook"
        + " bindings, missing-workbook policy, and template-backed UDF toolpacks."
        + " Omit formulaEnvironment when the default evaluator is intended.";
  }

  /** One stable catalog summary for `ExecutionJournalInput`. */
  public static String executionJournalInputSummary() {
    return "Explicit execution-journal settings."
        + " Use ExecutionJournalInput.defaults() for SUMMARY detail."
        + " SUMMARY returns compact, stable target summaries without timing telemetry,"
        + " NORMAL returns expanded target summaries with timing telemetry,"
        + " and VERBOSE also enables fine-grained event emission for CLI stderr rendering.";
  }

  /** One stable catalog summary for `CalculationPolicyInput`. */
  public static String calculationPolicyInputSummary() {
    return "Explicit formula-calculation policy."
        + " Use CalculationPolicyInput.defaults() for DO_NOT_CALCULATE with"
        + " markRecalculateOnOpen=false."
        + " Use DEFERRED_CALCULATION to inspect capabilities without immediate evaluation,"
        + " Use EVALUATE_ALL or EVALUATE_TARGETS for immediate server-side evaluation,"
        + " CLEAR_CACHES_ONLY to strip persisted caches,"
        + " or markRecalculateOnOpen=true when Excel-compatible clients should recalculate later.";
  }

  /** One stable catalog summary for `CalculationStrategyInput`. */
  public static String calculationStrategyInputSummary() {
    return "One explicit formula-calculation strategy."
        + " DO_NOT_CALCULATE leaves formulas untouched,"
        + " DEFERRED_CALCULATION reports formula capabilities but leaves evaluation to an"
        + " Excel-compatible client,"
        + " EVALUATE_ALL evaluates every formula cell when all are immediately evaluable and"
        + " otherwise returns PARTIAL without fabricating computed values,"
        + " EVALUATE_TARGETS applies the same lenient rule to one explicit formula-cell set,"
        + " REQUIRE_EVALUATION fails when any formula cannot be evaluated immediately,"
        + " and CLEAR_CACHES_ONLY strips persisted formula caches without evaluating.";
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
