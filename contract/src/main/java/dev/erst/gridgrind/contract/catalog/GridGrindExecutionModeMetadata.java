package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.dto.CalculationStrategyInput;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import java.util.List;
import java.util.Objects;

/**
 * Structured execution-mode contract metadata shared by help, catalog summaries, validation, and
 * runtime error messages.
 */
public final class GridGrindExecutionModeMetadata {
  private static final EventReadMode EVENT_READ =
      new EventReadMode(
          ExecutionModeInput.ReadMode.EVENT_READ,
          List.of(InspectionQuery.GetWorkbookSummary.class, InspectionQuery.GetSheetSummary.class),
          CalculationStrategyInput.DoNotCalculate.class,
          false);
  private static final StreamingWriteMode STREAMING_WRITE =
      new StreamingWriteMode(
          ExecutionModeInput.WriteMode.STREAMING_WRITE,
          WorkbookPlan.WorkbookSource.New.class,
          List.of(WorkbookMutationAction.EnsureSheet.class, CellMutationAction.AppendRow.class),
          CalculationStrategyInput.DoNotCalculate.class,
          true);

  private GridGrindExecutionModeMetadata() {}

  /** Returns the structured EVENT_READ contract metadata. */
  public static EventReadMode eventRead() {
    return EVENT_READ;
  }

  /** Returns the structured STREAMING_WRITE contract metadata. */
  public static StreamingWriteMode streamingWrite() {
    return STREAMING_WRITE;
  }

  /** Structured low-memory summary-read contract for `execution.mode.readMode=EVENT_READ`. */
  public record EventReadMode(
      ExecutionModeInput.ReadMode mode,
      List<Class<? extends InspectionQuery>> allowedQueries,
      Class<? extends CalculationStrategyInput> requiredCalculationStrategy,
      boolean markRecalculateOnOpenAllowed) {
    public EventReadMode {
      Objects.requireNonNull(mode, "mode must not be null");
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

    /** Returns the canonical validation message for incompatible calculation policy. */
    public String calculationFailureMessage() {
      return "execution.mode.readMode="
          + mode.name()
          + " requires execution.calculation.strategy="
          + requiredCalculationStrategyId()
          + " and markRecalculateOnOpen="
          + markRecalculateOnOpenAllowed;
    }

    /** Returns the canonical validation message for a non-inspection step in EVENT_READ. */
    public String unsupportedStepMessage(String stepKind) {
      Objects.requireNonNull(stepKind, "stepKind must not be null");
      return "execution.mode.readMode="
          + mode.name()
          + " supports inspection steps only; unsupported step kind: "
          + stepKind;
    }

    /** Returns the canonical validation message for an unsupported query id in EVENT_READ. */
    public String unsupportedQueryMessage(String queryType) {
      Objects.requireNonNull(queryType, "queryType must not be null");
      return "execution.mode.readMode="
          + mode.name()
          + " supports "
          + allowedQueryPhrase()
          + " only; unsupported inspection query type: "
          + queryType;
    }

    /** Returns the shared summary text used by help and catalog surfaces. */
    public String catalogSummary() {
      return allowedQueryPhrase()
          + " only, with execution.calculation.strategy="
          + requiredCalculationStrategyId()
          + " and markRecalculateOnOpen="
          + markRecalculateOnOpenAllowed
          + ".";
    }
  }

  /** Structured low-memory append-write contract for `execution.mode.writeMode=STREAMING_WRITE`. */
  public record StreamingWriteMode(
      ExecutionModeInput.WriteMode mode,
      Class<? extends WorkbookPlan.WorkbookSource> requiredSourceType,
      List<Class<? extends MutationAction>> allowedActions,
      Class<? extends CalculationStrategyInput> requiredCalculationStrategy,
      boolean markRecalculateOnOpenAllowed) {
    public StreamingWriteMode {
      Objects.requireNonNull(mode, "mode must not be null");
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

    /** Returns the canonical validation message for incompatible calculation policy. */
    public String calculationFailureMessage() {
      return "execution.mode.writeMode="
          + mode.name()
          + " requires execution.calculation.strategy="
          + requiredCalculationStrategyId()
          + " because low-memory streaming writes do not support immediate server-side"
          + " evaluation or cache clearing";
    }

    /** Returns the canonical validation message for an unsupported source type. */
    public String invalidSourceMessage() {
      return "execution.mode.writeMode="
          + mode.name()
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
      return "execution.mode.writeMode="
          + mode.name()
          + " supports "
          + allowedActionPhrase()
          + " only; unsupported mutation action type: "
          + actionType;
    }

    /** Returns the canonical validation message for a missing sheet materialization step. */
    public String missingEnsureSheetBeforeAppendMessage() {
      return "execution.mode.writeMode="
          + mode.name()
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
      return "execution.mode.writeMode="
          + mode.name()
          + " requires "
          + GridGrindProtocolTypeNames.mutationActionTypeName(
              WorkbookMutationAction.EnsureSheet.class)
          + " before any assertion step so the streaming workbook can"
          + " materialize a sheet";
    }

    /** Returns the canonical validation message for a missing sheet before one inspection step. */
    public String missingEnsureSheetBeforeInspectionMessage() {
      return "execution.mode.writeMode="
          + mode.name()
          + " requires "
          + GridGrindProtocolTypeNames.mutationActionTypeName(
              WorkbookMutationAction.EnsureSheet.class)
          + " before any inspection step so the streaming workbook can"
          + " materialize a sheet";
    }

    /** Returns the canonical validation message for a streaming plan with no sheet creation. */
    public String missingEnsureSheetMutationMessage() {
      return "execution.mode.writeMode="
          + mode.name()
          + " requires at least one "
          + GridGrindProtocolTypeNames.mutationActionTypeName(
              WorkbookMutationAction.EnsureSheet.class)
          + " mutation";
    }

    /** Returns the shared summary text used by help and catalog surfaces. */
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
