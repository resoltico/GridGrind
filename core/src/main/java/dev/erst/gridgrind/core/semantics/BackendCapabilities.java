package dev.erst.gridgrind.core.semantics;

import dev.erst.gridgrind.core.catalog.GridGrindContractText;
import dev.erst.gridgrind.core.catalog.GridGrindProtocolTypeNames;
import dev.erst.gridgrind.core.model.formula.CalculationPolicySpec;
import dev.erst.gridgrind.core.model.formula.CalculationStrategySpec;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.ExecutionModeSpec;
import dev.erst.gridgrind.core.plan.InspectionQuery;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.MutationAction;
import dev.erst.gridgrind.core.plan.SheetIntrospectionQuery;
import dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import java.util.List;
import java.util.Objects;

/**
 * Structured execution-mode contract metadata shared by help, catalog summaries, validation, and
 * runtime error messages.
 */
@dev.erst.gridgrind.core.limits.EnforcesLimit({
  dev.erst.gridgrind.core.limits.Limit.ID_EVENT_READ_EXECUTION_MODE_SCOPE,
  dev.erst.gridgrind.core.limits.Limit.ID_STREAMING_WRITE_EXECUTION_MODE_SCOPE
})
public final class BackendCapabilities {
  private static final FullXssfMode FULL_XSSF = new FullXssfMode(ExecutionModeSpec.FullXssf.class);
  private static final EventReadMode EVENT_READ =
      new EventReadMode(
          ExecutionModeSpec.EventRead.class,
          List.of(
              WorkbookIntrospectionQuery.GetWorkbookSummary.class,
              SheetIntrospectionQuery.GetSheetSummary.class),
          CalculationStrategySpec.DoNotCalculate.class,
          false);
  private static final StreamingWriteMode STREAMING_WRITE =
      new StreamingWriteMode(
          ExecutionModeSpec.StreamingWrite.class,
          WorkbookPlan.WorkbookSource.New.class,
          List.of(WorkbookMutationAction.EnsureSheet.class, CellMutationAction.AppendRow.class),
          CalculationStrategySpec.DoNotCalculate.class,
          true);

  private BackendCapabilities() {}

  /** The pure capability declaration exposed by an execution backend. */
  public sealed interface Mode permits FullXssfMode, EventReadMode, StreamingWriteMode {
    /** Returns the authored execution-mode record represented by this capability declaration. */
    Class<? extends ExecutionModeSpec> modeClass();

    /** Returns the discriminator derived from the mode record metadata. */
    String modeId();

    /** Describes supported operations and constraints for catalog/help rendering. */
    String catalogSummary();

    /** Reports whether this backend can take the admitted plan on its execution path. */
    boolean accepts(WorkbookPlan request);
  }

  /** Returns the canonical backend declaration for a mode record, if this is a mode type. */
  public static java.util.Optional<Mode> forModeClass(Class<?> type) {
    return java.util.stream.Stream.<Mode>of(FULL_XSSF, EVENT_READ, STREAMING_WRITE)
        .filter(mode -> mode.modeClass().equals(type))
        .findFirst();
  }

  /** Returns the standard FULL_XSSF contract metadata. */
  public static FullXssfMode fullXssf() {
    return FULL_XSSF;
  }

  /** Returns the structured EVENT_READ contract metadata. */
  public static EventReadMode eventRead() {
    return EVENT_READ;
  }

  /** Returns the structured STREAMING_WRITE contract metadata. */
  public static StreamingWriteMode streamingWrite() {
    return STREAMING_WRITE;
  }

  /** Structured standard full-memory execution contract for `execution.mode.type=FULL_XSSF`. */
  public record FullXssfMode(Class<? extends ExecutionModeSpec> modeClass) implements Mode {
    @Override
    public boolean accepts(WorkbookPlan request) {
      return modeClass.isInstance(request.effectiveExecutionMode())
          || (EVENT_READ.modeClass().isInstance(request.effectiveExecutionMode())
              && !EVENT_READ.accepts(request));
    }

    public FullXssfMode {
      Objects.requireNonNull(modeClass, "modeClass must not be null");
    }

    /** Returns the canonical mode id. */
    @Override
    public String modeId() {
      return GridGrindProtocolTypeNames.executionModeTypeName(modeClass);
    }

    /** Returns the shared summary text used by help and catalog surfaces. */
    @Override
    public String catalogSummary() {
      return "Standard full workbook read and write execution with no special low-memory"
          + " restrictions.";
    }
  }

  /** Structured summary-read contract for `execution.mode.type=EVENT_READ`. */
  public record EventReadMode(
      Class<? extends ExecutionModeSpec> modeClass,
      List<Class<? extends InspectionQuery>> allowedQueries,
      Class<? extends CalculationStrategySpec> requiredCalculationStrategy,
      boolean markRecalculateOnOpenAllowed)
      implements Mode {
    public EventReadMode {
      Objects.requireNonNull(modeClass, "modeClass must not be null");
      allowedQueries = List.copyOf(allowedQueries);
      if (allowedQueries.isEmpty()) {
        throw new IllegalArgumentException("allowedQueries must not be empty");
      }
      for (Class<? extends InspectionQuery> allowedQuery : allowedQueries) {
        Objects.requireNonNull(allowedQuery, "allowedQueries must not contain nulls");
      }
      Objects.requireNonNull(
          requiredCalculationStrategy, "requiredCalculationStrategy must not be null");
    }

    /** Calculation constraints shared by routing, validation and catalog. */
    public boolean allowsCalculation(CalculationPolicySpec policy) {
      return requiredCalculationStrategy.isInstance(policy.strategy())
          && (markRecalculateOnOpenAllowed || !policy.markRecalculateOnOpen());
    }

    /** Direct source-read eligibility; other EVENT_READ requests use XSSF materialization. */
    @Override
    public boolean accepts(WorkbookPlan request) {
      return modeClass.isInstance(request.effectiveExecutionMode())
          && allowsCalculation(request.calculationPolicy())
          && request.steps().stream().allMatch(InspectionStep.class::isInstance)
          && request.persistence() instanceof WorkbookPlan.WorkbookPersistence.None
          && request.source() instanceof WorkbookPlan.WorkbookSource.ExistingFile;
    }

    /** Returns the canonical query ids accepted by EVENT_READ. */
    public List<String> allowedQueryIds() {
      return allowedQueries.stream()
          .map(GridGrindProtocolTypeNames::inspectionQueryTypeName)
          .toList();
    }

    /** Returns a human-readable phrase of the canonical query ids accepted by EVENT_READ. */
    public String allowedQueryPhrase() {
      return GridGrindContractText.humanJoin(allowedQueryIds());
    }

    /** Returns the canonical strategy id required by EVENT_READ. */
    public String requiredCalculationStrategyId() {
      return GridGrindProtocolTypeNames.calculationStrategyTypeName(requiredCalculationStrategy);
    }

    /** Returns the canonical mode id. */
    @Override
    public String modeId() {
      return GridGrindProtocolTypeNames.executionModeTypeName(modeClass);
    }

    /** Returns the canonical validation message for incompatible calculation policy. */
    public String calculationFailureMessage() {
      return "execution.mode.type="
          + modeId()
          + " requires execution.calculation.strategy="
          + requiredCalculationStrategyId()
          + " and markRecalculateOnOpen="
          + markRecalculateOnOpenAllowed;
    }

    /** Returns the canonical validation message for a non-inspection step in EVENT_READ. */
    public String unsupportedStepMessage(String stepKind) {
      Objects.requireNonNull(stepKind, "stepKind must not be null");
      return "execution.mode.type="
          + modeId()
          + " supports inspection steps only; unsupported step kind: "
          + stepKind;
    }

    /** Returns the canonical validation message for an unsupported query id in EVENT_READ. */
    public String unsupportedQueryMessage(String queryType) {
      Objects.requireNonNull(queryType, "queryType must not be null");
      return "execution.mode.type="
          + modeId()
          + " supports "
          + allowedQueryPhrase()
          + " only; unsupported inspection query type: "
          + queryType;
    }

    /** Returns the shared summary text used by help and catalog surfaces. */
    @Override
    public String catalogSummary() {
      return allowedQueryPhrase()
          + " only, with execution.calculation.strategy="
          + requiredCalculationStrategyId()
          + " and markRecalculateOnOpen="
          + markRecalculateOnOpenAllowed
          + ".";
    }
  }

  /** Structured append-write contract for `execution.mode.type=STREAMING_WRITE`. */
  public record StreamingWriteMode(
      Class<? extends ExecutionModeSpec> modeClass,
      Class<? extends WorkbookPlan.WorkbookSource> requiredSourceType,
      List<Class<? extends MutationAction>> allowedActions,
      Class<? extends CalculationStrategySpec> requiredCalculationStrategy,
      boolean markRecalculateOnOpenAllowed)
      implements Mode {
    public StreamingWriteMode {
      Objects.requireNonNull(modeClass, "modeClass must not be null");
      Objects.requireNonNull(requiredSourceType, "requiredSourceType must not be null");
      allowedActions = List.copyOf(allowedActions);
      if (allowedActions.isEmpty()) {
        throw new IllegalArgumentException("allowedActions must not be empty");
      }
      for (Class<? extends MutationAction> allowedAction : allowedActions) {
        Objects.requireNonNull(allowedAction, "allowedActions must not contain nulls");
      }
      Objects.requireNonNull(
          requiredCalculationStrategy, "requiredCalculationStrategy must not be null");
    }

    /** Checks the calculation strategy and recalculate-on-open capability. */
    public boolean allowsCalculation(CalculationPolicySpec policy) {
      return requiredCalculationStrategy.isInstance(policy.strategy())
          && (markRecalculateOnOpenAllowed || !policy.markRecalculateOnOpen());
    }

    @Override
    public boolean accepts(WorkbookPlan request) {
      return modeClass.isInstance(request.effectiveExecutionMode());
    }

    /** Returns the canonical mutation ids accepted by STREAMING_WRITE. */
    public List<String> allowedActionIds() {
      return allowedActions.stream()
          .map(GridGrindProtocolTypeNames::mutationActionTypeName)
          .toList();
    }

    /**
     * Returns a human-readable phrase of the canonical mutation ids accepted by STREAMING_WRITE.
     */
    public String allowedActionPhrase() {
      return GridGrindContractText.humanJoin(allowedActionIds());
    }

    /** Returns the canonical source id required by STREAMING_WRITE. */
    public String requiredSourceTypeId() {
      return GridGrindProtocolTypeNames.workbookSourceTypeName(requiredSourceType);
    }

    /** Returns the canonical strategy id required by STREAMING_WRITE. */
    public String requiredCalculationStrategyId() {
      return GridGrindProtocolTypeNames.calculationStrategyTypeName(requiredCalculationStrategy);
    }

    /** Returns the canonical mode id. */
    @Override
    public String modeId() {
      return GridGrindProtocolTypeNames.executionModeTypeName(modeClass);
    }

    /** Returns the canonical validation message for incompatible calculation policy. */
    public String calculationFailureMessage() {
      return "execution.mode.type="
          + modeId()
          + " requires execution.calculation.strategy="
          + requiredCalculationStrategyId()
          + " because low-memory streaming writes do not support immediate server-side"
          + " evaluation or cache clearing";
    }

    /** Returns the canonical validation message for an unsupported source type. */
    public String invalidSourceMessage() {
      return "execution.mode.type="
          + modeId()
          + " requires source.type="
          + requiredSourceTypeId()
          + " because low-memory streaming writes do not author in-place edits on "
          + GridGrindProtocolTypeNames.workbookSourceTypeName(
              WorkbookPlan.WorkbookSource.ExistingFile.class)
          + " workbooks";
    }

    /** Returns the canonical validation message for an unsupported mutation id. */
    public String unsupportedActionMessage(String actionType) {
      Objects.requireNonNull(actionType, "actionType must not be null");
      return "execution.mode.type="
          + modeId()
          + " supports "
          + allowedActionPhrase()
          + " only; unsupported mutation action type: "
          + actionType;
    }

    /** Returns the canonical validation message for a missing sheet materialization step. */
    public String missingEnsureSheetBeforeAppendMessage() {
      return "execution.mode.type="
          + modeId()
          + " requires "
          + GridGrindProtocolTypeNames.mutationActionTypeName(
              WorkbookMutationAction.EnsureSheet.class)
          + " before "
          + GridGrindProtocolTypeNames.mutationActionTypeName(CellMutationAction.AppendRow.class)
          + " so the streaming writer has a"
          + " materialized sheet target";
    }

    /** Returns the canonical validation message for a missing sheet before one assertion step. */
    public String missingEnsureSheetBeforeAssertionMessage() {
      return "execution.mode.type="
          + modeId()
          + " requires "
          + GridGrindProtocolTypeNames.mutationActionTypeName(
              WorkbookMutationAction.EnsureSheet.class)
          + " before any assertion step so the streaming workbook can"
          + " materialize a sheet";
    }

    /** Returns the canonical validation message for a missing sheet before one inspection step. */
    public String missingEnsureSheetBeforeInspectionMessage() {
      return "execution.mode.type="
          + modeId()
          + " requires "
          + GridGrindProtocolTypeNames.mutationActionTypeName(
              WorkbookMutationAction.EnsureSheet.class)
          + " before any inspection step so the streaming workbook can"
          + " materialize a sheet";
    }

    /** Returns the canonical validation message for a streaming plan with no sheet creation. */
    public String missingEnsureSheetMutationMessage() {
      return "execution.mode.type="
          + modeId()
          + " requires at least one "
          + GridGrindProtocolTypeNames.mutationActionTypeName(
              WorkbookMutationAction.EnsureSheet.class)
          + " mutation";
    }

    /** Returns the shared summary text used by help and catalog surfaces. */
    @Override
    public String catalogSummary() {
      return "source.type must be "
          + requiredSourceTypeId()
          + "; mutation actions limited to "
          + allowedActionPhrase()
          + "; execution.calculation must keep strategy="
          + requiredCalculationStrategyId()
          + " and may set markRecalculateOnOpen="
          + markRecalculateOnOpenAllowed
          + ".";
    }
  }
}
