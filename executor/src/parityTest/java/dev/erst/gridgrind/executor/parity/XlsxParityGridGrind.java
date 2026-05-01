package dev.erst.gridgrind.executor.parity;

import dev.erst.gridgrind.contract.catalog.Catalog;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog;
import dev.erst.gridgrind.contract.catalog.TypeEntry;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.GridGrindResponsePersistence;
import dev.erst.gridgrind.contract.dto.OoxmlOpenSecurityInput;
import dev.erst.gridgrind.contract.dto.OoxmlPersistenceSecurityInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.executor.DefaultGridGrindRequestExecutor;
import dev.erst.gridgrind.executor.ExecutionInputBindings;
import dev.erst.gridgrind.executor.ExecutionJournalSink;
import dev.erst.gridgrind.executor.parity.ParityPlanSupport.PendingMutation;
import java.nio.file.Path;
import java.util.List;

/** Small parity-test helper for executing GridGrind requests and reading the protocol catalog. */
final class XlsxParityGridGrind {
  private XlsxParityGridGrind() {}

  static Catalog catalog() {
    return GridGrindProtocolCatalog.catalog();
  }

  static boolean hasSourceType(String id) {
    return hasType(catalog().sourceTypes(), id);
  }

  static boolean hasPersistenceType(String id) {
    return hasType(catalog().persistenceTypes(), id);
  }

  static boolean hasMutationActionType(String id) {
    return hasType(catalog().mutationActionTypes(), id);
  }

  static boolean hasCalculationStrategyType(String id) {
    return catalog().nestedTypes().stream()
        .filter(group -> "calculationStrategyTypes".equals(group.group()))
        .flatMap(group -> group.types().stream())
        .map(TypeEntry::id)
        .anyMatch(id::equals);
  }

  static boolean hasInspectionQueryType(String id) {
    return hasType(catalog().inspectionQueryTypes(), id);
  }

  static GridGrindResponse executeReadWorkbook(Path workbookPath, InspectionStep... inspections) {
    return executeReadWorkbook(workbookPath, (OoxmlOpenSecurityInput) null, inspections);
  }

  static GridGrindResponse executeReadWorkbook(
      Path workbookPath, OoxmlOpenSecurityInput sourceSecurity, InspectionStep... inspections) {
    return executeReadWorkbook(workbookPath, sourceSecurity, null, inspections);
  }

  static GridGrindResponse executeReadWorkbook(
      Path workbookPath, ExecutionModeInput executionMode, InspectionStep... inspections) {
    return executeReadWorkbook(workbookPath, null, executionMode, inspections);
  }

  static GridGrindResponse executeReadWorkbook(
      Path workbookPath,
      OoxmlOpenSecurityInput sourceSecurity,
      ExecutionModeInput executionMode,
      InspectionStep... inspections) {
    return executeReadWorkbook(workbookPath, sourceSecurity, executionMode, null, inspections);
  }

  static GridGrindResponse executeReadWorkbook(
      Path workbookPath,
      FormulaEnvironmentInput formulaEnvironment,
      InspectionStep... inspections) {
    return executeReadWorkbook(
        workbookPath,
        (OoxmlOpenSecurityInput) null,
        (ExecutionModeInput) null,
        formulaEnvironment,
        inspections);
  }

  static GridGrindResponse executeReadWorkbook(
      Path workbookPath,
      ExecutionModeInput executionMode,
      FormulaEnvironmentInput formulaEnvironment,
      InspectionStep... inspections) {
    return executeReadWorkbook(workbookPath, null, executionMode, formulaEnvironment, inspections);
  }

  static GridGrindResponse executeReadWorkbook(
      Path workbookPath,
      OoxmlOpenSecurityInput sourceSecurity,
      ExecutionModeInput executionMode,
      FormulaEnvironmentInput formulaEnvironment,
      InspectionStep... inspections) {
    return executeReadWorkbook(
        workbookPath,
        sourceSecurity,
        executionMode == null
            ? ExecutionPolicyInput.defaults()
            : ExecutionPolicyInput.mode(executionMode),
        formulaEnvironment,
        inspections);
  }

  static GridGrindResponse executeReadWorkbook(
      Path workbookPath,
      OoxmlOpenSecurityInput sourceSecurity,
      ExecutionPolicyInput execution,
      FormulaEnvironmentInput formulaEnvironment,
      InspectionStep... inspections) {
    return execute(
        WorkbookPlan.standard(
            existingWorkbookSource(workbookPath, sourceSecurity),
            new WorkbookPlan.WorkbookPersistence.None(),
            execution == null ? ExecutionPolicyInput.defaults() : execution,
            formulaEnvironment == null ? FormulaEnvironmentInput.empty() : formulaEnvironment,
            List.of(inspections)));
  }

  private static GridGrindResponse execute(WorkbookPlan request) {
    return new DefaultGridGrindRequestExecutor()
        .execute(request, ExecutionInputBindings.processDefault(), ExecutionJournalSink.NOOP);
  }

  private static WorkbookPlan.WorkbookSource.ExistingFile existingWorkbookSource(
      Path workbookPath, OoxmlOpenSecurityInput sourceSecurity) {
    return sourceSecurity == null
        ? new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString())
        : new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString(), sourceSecurity);
  }

  private static WorkbookPlan.WorkbookPersistence.SaveAs saveAsPersistence(
      Path saveAsPath, OoxmlPersistenceSecurityInput persistenceSecurity) {
    return persistenceSecurity == null
        ? new WorkbookPlan.WorkbookPersistence.SaveAs(saveAsPath.toString())
        : new WorkbookPlan.WorkbookPersistence.SaveAs(saveAsPath.toString(), persistenceSecurity);
  }

  static GridGrindResponse.Success readWorkbook(Path workbookPath, InspectionStep... inspections) {
    return success(executeReadWorkbook(workbookPath, (OoxmlOpenSecurityInput) null, inspections));
  }

  static GridGrindResponse executeMutateWorkbook(
      Path workbookPath,
      Path saveAsPath,
      List<PendingMutation> mutations,
      InspectionStep... inspections) {
    return executeMutateWorkbook(
        workbookPath, null, saveAsPath, null, null, mutations, inspections);
  }

  static GridGrindResponse executeMutateWorkbook(
      Path workbookPath,
      OoxmlOpenSecurityInput sourceSecurity,
      Path saveAsPath,
      OoxmlPersistenceSecurityInput persistenceSecurity,
      FormulaEnvironmentInput formulaEnvironment,
      List<PendingMutation> mutations,
      InspectionStep... inspections) {
    return executeMutateWorkbook(
        workbookPath,
        sourceSecurity,
        saveAsPath,
        persistenceSecurity,
        null,
        formulaEnvironment,
        mutations,
        inspections);
  }

  static GridGrindResponse executeMutateWorkbook(
      Path workbookPath,
      OoxmlOpenSecurityInput sourceSecurity,
      Path saveAsPath,
      OoxmlPersistenceSecurityInput persistenceSecurity,
      ExecutionPolicyInput execution,
      FormulaEnvironmentInput formulaEnvironment,
      List<PendingMutation> mutations,
      InspectionStep... inspections) {
    return execute(
        ParityPlanSupport.request(
            existingWorkbookSource(workbookPath, sourceSecurity),
            saveAsPersistence(saveAsPath, persistenceSecurity),
            execution,
            formulaEnvironment,
            mutations,
            List.of(inspections)));
  }

  static GridGrindResponse.Success readWorkbook(
      Path workbookPath, OoxmlOpenSecurityInput sourceSecurity, InspectionStep... inspections) {
    return success(executeReadWorkbook(workbookPath, sourceSecurity, inspections));
  }

  static GridGrindResponse.Success readWorkbook(
      Path workbookPath, ExecutionModeInput executionMode, InspectionStep... inspections) {
    return success(executeReadWorkbook(workbookPath, executionMode, inspections));
  }

  static GridGrindResponse.Success readWorkbook(
      Path workbookPath,
      FormulaEnvironmentInput formulaEnvironment,
      InspectionStep... inspections) {
    return success(executeReadWorkbook(workbookPath, formulaEnvironment, inspections));
  }

  static GridGrindResponse.Success readWorkbook(
      Path workbookPath,
      ExecutionModeInput executionMode,
      FormulaEnvironmentInput formulaEnvironment,
      InspectionStep... inspections) {
    return success(
        executeReadWorkbook(workbookPath, executionMode, formulaEnvironment, inspections));
  }

  static GridGrindResponse.Success mutateWorkbook(
      Path workbookPath,
      Path saveAsPath,
      List<PendingMutation> mutations,
      InspectionStep... inspections) {
    return mutateWorkbook(workbookPath, null, saveAsPath, null, null, mutations, inspections);
  }

  static GridGrindResponse.Success mutateWorkbook(
      Path workbookPath,
      Path saveAsPath,
      FormulaEnvironmentInput formulaEnvironment,
      List<PendingMutation> mutations,
      InspectionStep... inspections) {
    return mutateWorkbook(
        workbookPath, null, saveAsPath, null, formulaEnvironment, mutations, inspections);
  }

  static GridGrindResponse.Success mutateWorkbook(
      Path workbookPath,
      OoxmlOpenSecurityInput sourceSecurity,
      Path saveAsPath,
      OoxmlPersistenceSecurityInput persistenceSecurity,
      FormulaEnvironmentInput formulaEnvironment,
      List<PendingMutation> mutations,
      InspectionStep... inspections) {
    return mutateWorkbook(
        workbookPath,
        sourceSecurity,
        saveAsPath,
        persistenceSecurity,
        null,
        formulaEnvironment,
        mutations,
        inspections);
  }

  static GridGrindResponse.Success mutateWorkbook(
      Path workbookPath,
      OoxmlOpenSecurityInput sourceSecurity,
      Path saveAsPath,
      OoxmlPersistenceSecurityInput persistenceSecurity,
      ExecutionPolicyInput execution,
      FormulaEnvironmentInput formulaEnvironment,
      List<PendingMutation> mutations,
      InspectionStep... inspections) {
    return success(
        executeMutateWorkbook(
            workbookPath,
            sourceSecurity,
            saveAsPath,
            persistenceSecurity,
            execution,
            formulaEnvironment,
            mutations,
            inspections));
  }

  static GridGrindResponse.Success mutateWorkbook(
      Path workbookPath,
      Path saveAsPath,
      ExecutionPolicyInput execution,
      FormulaEnvironmentInput formulaEnvironment,
      List<PendingMutation> mutations,
      InspectionStep... inspections) {
    return mutateWorkbook(
        workbookPath,
        null,
        saveAsPath,
        null,
        execution,
        formulaEnvironment,
        mutations,
        inspections);
  }

  static GridGrindResponse.Success writeNewWorkbook(
      Path saveAsPath,
      ExecutionModeInput executionMode,
      List<PendingMutation> mutations,
      InspectionStep... inspections) {
    return writeNewWorkbook(saveAsPath, null, executionMode, mutations, inspections);
  }

  static GridGrindResponse.Success writeNewWorkbook(
      Path saveAsPath,
      OoxmlPersistenceSecurityInput persistenceSecurity,
      ExecutionModeInput executionMode,
      List<PendingMutation> mutations,
      InspectionStep... inspections) {
    return writeNewWorkbook(
        saveAsPath,
        persistenceSecurity,
        executionMode == null
            ? ExecutionPolicyInput.defaults()
            : ExecutionPolicyInput.mode(executionMode),
        mutations,
        inspections);
  }

  static GridGrindResponse.Success writeNewWorkbook(
      Path saveAsPath,
      OoxmlPersistenceSecurityInput persistenceSecurity,
      ExecutionPolicyInput execution,
      List<PendingMutation> mutations,
      InspectionStep... inspections) {
    return success(
        execute(
            WorkbookPlan.standard(
                new WorkbookPlan.WorkbookSource.New(),
                saveAsPersistence(saveAsPath, persistenceSecurity),
                execution,
                FormulaEnvironmentInput.empty(),
                ParityPlanSupport.steps(mutations, List.of(inspections)))));
  }

  static GridGrindResponse.Success overwriteWorkbook(
      Path workbookPath, List<PendingMutation> mutations, InspectionStep... inspections) {
    return overwriteWorkbook(workbookPath, null, mutations, inspections);
  }

  static GridGrindResponse.Success overwriteWorkbook(
      Path workbookPath,
      FormulaEnvironmentInput formulaEnvironment,
      List<PendingMutation> mutations,
      InspectionStep... inspections) {
    return overwriteWorkbook(workbookPath, null, formulaEnvironment, mutations, inspections);
  }

  static GridGrindResponse.Success overwriteWorkbook(
      Path workbookPath,
      ExecutionPolicyInput execution,
      FormulaEnvironmentInput formulaEnvironment,
      List<PendingMutation> mutations,
      InspectionStep... inspections) {
    return success(
        execute(
            ParityPlanSupport.request(
                new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                new WorkbookPlan.WorkbookPersistence.OverwriteSource(),
                execution,
                formulaEnvironment,
                mutations,
                List.of(inspections))));
  }

  static GridGrindResponse.Failure mutateWorkbookExpectingFailure(
      Path workbookPath, List<PendingMutation> mutations, InspectionStep... inspections) {
    return mutateWorkbookExpectingFailure(workbookPath, null, mutations, inspections);
  }

  static GridGrindResponse.Failure mutateWorkbookExpectingFailure(
      Path workbookPath,
      FormulaEnvironmentInput formulaEnvironment,
      List<PendingMutation> mutations,
      InspectionStep... inspections) {
    return mutateWorkbookExpectingFailure(
        workbookPath, null, formulaEnvironment, mutations, inspections);
  }

  static GridGrindResponse.Failure mutateWorkbookExpectingFailure(
      Path workbookPath,
      OoxmlOpenSecurityInput sourceSecurity,
      FormulaEnvironmentInput formulaEnvironment,
      List<PendingMutation> mutations,
      InspectionStep... inspections) {
    return mutateWorkbookExpectingFailure(
        workbookPath, sourceSecurity, null, formulaEnvironment, mutations, inspections);
  }

  static GridGrindResponse.Failure mutateWorkbookExpectingFailure(
      Path workbookPath,
      OoxmlOpenSecurityInput sourceSecurity,
      ExecutionPolicyInput execution,
      FormulaEnvironmentInput formulaEnvironment,
      List<PendingMutation> mutations,
      InspectionStep... inspections) {
    GridGrindResponse response =
        execute(
            ParityPlanSupport.request(
                existingWorkbookSource(workbookPath, sourceSecurity),
                new WorkbookPlan.WorkbookPersistence.None(),
                execution,
                formulaEnvironment,
                mutations,
                List.of(inspections)));
    if (response instanceof GridGrindResponse.Failure failure) {
      return failure;
    }
    throw new AssertionError("Expected GridGrind failure but request succeeded: " + response);
  }

  static String savedPath(GridGrindResponse.Success success) {
    return switch (success.persistence()) {
      case GridGrindResponsePersistence.PersistenceOutcome.SavedAs savedAs ->
          savedAs.executionPath();
      case GridGrindResponsePersistence.PersistenceOutcome.Overwritten overwritten ->
          overwritten.executionPath();
      case GridGrindResponsePersistence.PersistenceOutcome.NotSaved _ ->
          throw new AssertionError("Expected the workbook to be persisted");
    };
  }

  static <T extends InspectionResult> T read(
      GridGrindResponse.Success success, String stepId, Class<T> type) {
    return type.cast(
        success.inspections().stream()
            .filter(result -> result.stepId().equals(stepId))
            .findFirst()
            .filter(type::isInstance)
            .orElseThrow(
                () -> new AssertionError("Missing read result " + stepId + " as " + type)));
  }

  private static GridGrindResponse.Success success(GridGrindResponse response) {
    if (response instanceof GridGrindResponse.Success success) {
      return success;
    }
    throw new AssertionError("Expected GridGrind success but got " + response);
  }

  private static boolean hasType(List<TypeEntry> entries, String id) {
    return entries.stream().map(TypeEntry::id).anyMatch(id::equals);
  }
}
