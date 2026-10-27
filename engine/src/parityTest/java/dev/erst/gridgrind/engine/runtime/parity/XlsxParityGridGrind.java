package dev.erst.gridgrind.engine.runtime.parity;

import dev.erst.gridgrind.core.catalog.Catalog;
import dev.erst.gridgrind.core.catalog.GridGrindProtocolCatalog;
import dev.erst.gridgrind.core.catalog.TypeEntry;
import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.model.security.OoxmlOpenSecuritySpec;
import dev.erst.gridgrind.core.model.security.OoxmlPersistenceSecuritySpec;
import dev.erst.gridgrind.core.plan.ExecutionModeSpec;
import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
import dev.erst.gridgrind.core.plan.InspectionResult;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.core.result.WorkbookResultPersistence;
import dev.erst.gridgrind.engine.runtime.DefaultGridGrindRequestExecutor;
import dev.erst.gridgrind.engine.runtime.ExecutionProgressSink;
import dev.erst.gridgrind.engine.runtime.parity.ParityPlanSupport.PendingMutation;
import java.nio.file.Path;
import java.util.List;
import java.util.Objects;

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
        .filter(group -> "calculationStrategySpecTypes".equals(group.group()))
        .flatMap(group -> group.types().stream())
        .map(TypeEntry::id)
        .anyMatch(id::equals);
  }

  static boolean hasInspectionQueryType(String id) {
    return hasType(catalog().inspectionQueryTypes(), id);
  }

  static WorkbookResult executeReadWorkbook(Path workbookPath, InspectionStep... inspections) {
    return executeReadWorkbook(workbookPath, (OoxmlOpenSecuritySpec) null, inspections);
  }

  static WorkbookResult executeReadWorkbook(
      Path workbookPath, OoxmlOpenSecuritySpec sourceSecurity, InspectionStep... inspections) {
    return executeReadWorkbook(workbookPath, sourceSecurity, null, inspections);
  }

  static WorkbookResult executeReadWorkbook(
      Path workbookPath, ExecutionModeSpec executionMode, InspectionStep... inspections) {
    return executeReadWorkbook(workbookPath, null, executionMode, inspections);
  }

  static WorkbookResult executeReadWorkbook(
      Path workbookPath,
      OoxmlOpenSecuritySpec sourceSecurity,
      ExecutionModeSpec executionMode,
      InspectionStep... inspections) {
    return executeReadWorkbook(workbookPath, sourceSecurity, executionMode, null, inspections);
  }

  static WorkbookResult executeReadWorkbook(
      Path workbookPath, FormulaEnvironmentSpec formulaEnvironment, InspectionStep... inspections) {
    return executeReadWorkbook(
        workbookPath,
        (OoxmlOpenSecuritySpec) null,
        (ExecutionModeSpec) null,
        formulaEnvironment,
        inspections);
  }

  static WorkbookResult executeReadWorkbook(
      Path workbookPath,
      ExecutionModeSpec executionMode,
      FormulaEnvironmentSpec formulaEnvironment,
      InspectionStep... inspections) {
    return executeReadWorkbook(workbookPath, null, executionMode, formulaEnvironment, inspections);
  }

  static WorkbookResult executeReadWorkbook(
      Path workbookPath,
      OoxmlOpenSecuritySpec sourceSecurity,
      ExecutionModeSpec executionMode,
      FormulaEnvironmentSpec formulaEnvironment,
      InspectionStep... inspections) {
    return executeReadWorkbook(
        workbookPath,
        sourceSecurity,
        executionMode == null
            ? ExecutionPolicySpec.defaults()
            : ExecutionPolicySpec.mode(executionMode),
        formulaEnvironment,
        inspections);
  }

  static WorkbookResult executeReadWorkbook(
      Path workbookPath,
      OoxmlOpenSecuritySpec sourceSecurity,
      ExecutionPolicySpec execution,
      FormulaEnvironmentSpec formulaEnvironment,
      InspectionStep... inspections) {
    return execute(
        XlsxParitySupport.executionRootFor(workbookPath),
        WorkbookPlan.standard(
            existingWorkbookSource(workbookPath, sourceSecurity),
            new WorkbookPlan.WorkbookPersistence.None(),
            execution == null ? ExecutionPolicySpec.defaults() : execution,
            formulaEnvironment == null ? FormulaEnvironmentSpec.empty() : formulaEnvironment,
            List.of(inspections)));
  }

  private static WorkbookResult execute(Path executionRoot, WorkbookPlan request) {
    return new DefaultGridGrindRequestExecutor()
        .execute(
            request,
            XlsxParitySupport.bindings(executionRoot, request),
            ExecutionProgressSink.NOOP);
  }

  private static WorkbookPlan.WorkbookSource.ExistingFile existingWorkbookSource(
      Path workbookPath, OoxmlOpenSecuritySpec sourceSecurity) {
    return sourceSecurity == null
        ? new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString())
        : new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString(), sourceSecurity);
  }

  private static WorkbookPlan.WorkbookPersistence.SaveAs saveAsPersistence(
      Path saveAsPath, OoxmlPersistenceSecuritySpec persistenceSecurity) {
    return persistenceSecurity == null
        ? new WorkbookPlan.WorkbookPersistence.SaveAs(
            saveAsPath.toString(), WorkbookPlan.WorkbookPersistence.IfExists.REJECT)
        : new WorkbookPlan.WorkbookPersistence.SaveAs(
            saveAsPath.toString(),
            WorkbookPlan.WorkbookPersistence.IfExists.REJECT,
            persistenceSecurity);
  }

  static WorkbookResult.Success readWorkbook(Path workbookPath, InspectionStep... inspections) {
    return success(executeReadWorkbook(workbookPath, (OoxmlOpenSecuritySpec) null, inspections));
  }

  static WorkbookResult executeMutateWorkbook(
      Path workbookPath,
      Path saveAsPath,
      List<PendingMutation> mutations,
      InspectionStep... inspections) {
    return executeMutateWorkbook(
        workbookPath,
        null,
        saveAsPath,
        OoxmlPersistenceSecuritySpec.none(),
        null,
        mutations,
        inspections);
  }

  static WorkbookResult executeMutateWorkbook(
      Path workbookPath,
      OoxmlOpenSecuritySpec sourceSecurity,
      Path saveAsPath,
      OoxmlPersistenceSecuritySpec persistenceSecurity,
      FormulaEnvironmentSpec formulaEnvironment,
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

  static WorkbookResult executeMutateWorkbook(
      Path workbookPath,
      OoxmlOpenSecuritySpec sourceSecurity,
      Path saveAsPath,
      OoxmlPersistenceSecuritySpec persistenceSecurity,
      ExecutionPolicySpec execution,
      FormulaEnvironmentSpec formulaEnvironment,
      List<PendingMutation> mutations,
      InspectionStep... inspections) {
    return execute(
        XlsxParitySupport.executionRootFor(saveAsPath),
        ParityPlanSupport.request(
            existingWorkbookSource(workbookPath, sourceSecurity),
            saveAsPersistence(
                saveAsPath,
                Objects.requireNonNull(
                    persistenceSecurity,
                    "EXISTING-source parity writes require an explicit security policy")),
            execution,
            formulaEnvironment,
            mutations,
            List.of(inspections)));
  }

  static WorkbookResult.Success readWorkbook(
      Path workbookPath, OoxmlOpenSecuritySpec sourceSecurity, InspectionStep... inspections) {
    return success(executeReadWorkbook(workbookPath, sourceSecurity, inspections));
  }

  static WorkbookResult.Success readWorkbook(
      Path workbookPath, ExecutionModeSpec executionMode, InspectionStep... inspections) {
    return success(executeReadWorkbook(workbookPath, executionMode, inspections));
  }

  static WorkbookResult.Success readWorkbook(
      Path workbookPath, FormulaEnvironmentSpec formulaEnvironment, InspectionStep... inspections) {
    return success(executeReadWorkbook(workbookPath, formulaEnvironment, inspections));
  }

  static WorkbookResult.Success readWorkbook(
      Path workbookPath,
      ExecutionModeSpec executionMode,
      FormulaEnvironmentSpec formulaEnvironment,
      InspectionStep... inspections) {
    return success(
        executeReadWorkbook(workbookPath, executionMode, formulaEnvironment, inspections));
  }

  static WorkbookResult.Success mutateWorkbook(
      Path workbookPath,
      Path saveAsPath,
      List<PendingMutation> mutations,
      InspectionStep... inspections) {
    return mutateWorkbook(
        workbookPath,
        null,
        saveAsPath,
        OoxmlPersistenceSecuritySpec.none(),
        null,
        mutations,
        inspections);
  }

  static WorkbookResult.Success mutateWorkbook(
      Path workbookPath,
      Path saveAsPath,
      FormulaEnvironmentSpec formulaEnvironment,
      List<PendingMutation> mutations,
      InspectionStep... inspections) {
    return mutateWorkbook(
        workbookPath,
        null,
        saveAsPath,
        OoxmlPersistenceSecuritySpec.none(),
        formulaEnvironment,
        mutations,
        inspections);
  }

  static WorkbookResult.Success mutateWorkbook(
      Path workbookPath,
      OoxmlOpenSecuritySpec sourceSecurity,
      Path saveAsPath,
      OoxmlPersistenceSecuritySpec persistenceSecurity,
      FormulaEnvironmentSpec formulaEnvironment,
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

  static WorkbookResult.Success mutateWorkbook(
      Path workbookPath,
      OoxmlOpenSecuritySpec sourceSecurity,
      Path saveAsPath,
      OoxmlPersistenceSecuritySpec persistenceSecurity,
      ExecutionPolicySpec execution,
      FormulaEnvironmentSpec formulaEnvironment,
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

  static WorkbookResult.Success mutateWorkbook(
      Path workbookPath,
      Path saveAsPath,
      ExecutionPolicySpec execution,
      FormulaEnvironmentSpec formulaEnvironment,
      List<PendingMutation> mutations,
      InspectionStep... inspections) {
    return mutateWorkbook(
        workbookPath,
        null,
        saveAsPath,
        OoxmlPersistenceSecuritySpec.none(),
        execution,
        formulaEnvironment,
        mutations,
        inspections);
  }

  static WorkbookResult.Success writeNewWorkbook(
      Path saveAsPath,
      ExecutionModeSpec executionMode,
      List<PendingMutation> mutations,
      InspectionStep... inspections) {
    return writeNewWorkbook(saveAsPath, null, executionMode, mutations, inspections);
  }

  static WorkbookResult.Success writeNewWorkbook(
      Path saveAsPath,
      OoxmlPersistenceSecuritySpec persistenceSecurity,
      ExecutionModeSpec executionMode,
      List<PendingMutation> mutations,
      InspectionStep... inspections) {
    return writeNewWorkbook(
        saveAsPath,
        persistenceSecurity,
        executionMode == null
            ? ExecutionPolicySpec.defaults()
            : ExecutionPolicySpec.mode(executionMode),
        mutations,
        inspections);
  }

  static WorkbookResult.Success writeNewWorkbook(
      Path saveAsPath,
      OoxmlPersistenceSecuritySpec persistenceSecurity,
      ExecutionPolicySpec execution,
      List<PendingMutation> mutations,
      InspectionStep... inspections) {
    return success(
        execute(
            XlsxParitySupport.executionRootFor(saveAsPath),
            WorkbookPlan.standard(
                new WorkbookPlan.WorkbookSource.New(),
                saveAsPersistence(saveAsPath, persistenceSecurity),
                execution,
                FormulaEnvironmentSpec.empty(),
                ParityPlanSupport.steps(mutations, List.of(inspections)))));
  }

  static WorkbookResult.Success overwriteWorkbook(
      Path workbookPath, List<PendingMutation> mutations, InspectionStep... inspections) {
    return overwriteWorkbook(workbookPath, null, mutations, inspections);
  }

  static WorkbookResult.Success overwriteWorkbook(
      Path workbookPath,
      FormulaEnvironmentSpec formulaEnvironment,
      List<PendingMutation> mutations,
      InspectionStep... inspections) {
    return overwriteWorkbook(workbookPath, null, formulaEnvironment, mutations, inspections);
  }

  static WorkbookResult.Success overwriteWorkbook(
      Path workbookPath,
      ExecutionPolicySpec execution,
      FormulaEnvironmentSpec formulaEnvironment,
      List<PendingMutation> mutations,
      InspectionStep... inspections) {
    return success(
        execute(
            XlsxParitySupport.executionRootFor(workbookPath),
            ParityPlanSupport.request(
                new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                new WorkbookPlan.WorkbookPersistence.Overwrite(OoxmlPersistenceSecuritySpec.none()),
                execution,
                formulaEnvironment,
                mutations,
                List.of(inspections))));
  }

  static WorkbookResult.Failure mutateWorkbookExpectingFailure(
      Path workbookPath, List<PendingMutation> mutations, InspectionStep... inspections) {
    return mutateWorkbookExpectingFailure(workbookPath, null, mutations, inspections);
  }

  static WorkbookResult.Failure mutateWorkbookExpectingFailure(
      Path workbookPath,
      FormulaEnvironmentSpec formulaEnvironment,
      List<PendingMutation> mutations,
      InspectionStep... inspections) {
    return mutateWorkbookExpectingFailure(
        workbookPath, null, formulaEnvironment, mutations, inspections);
  }

  static WorkbookResult.Failure mutateWorkbookExpectingFailure(
      Path workbookPath,
      OoxmlOpenSecuritySpec sourceSecurity,
      FormulaEnvironmentSpec formulaEnvironment,
      List<PendingMutation> mutations,
      InspectionStep... inspections) {
    return mutateWorkbookExpectingFailure(
        workbookPath, sourceSecurity, null, formulaEnvironment, mutations, inspections);
  }

  static WorkbookResult.Failure mutateWorkbookExpectingFailure(
      Path workbookPath,
      OoxmlOpenSecuritySpec sourceSecurity,
      ExecutionPolicySpec execution,
      FormulaEnvironmentSpec formulaEnvironment,
      List<PendingMutation> mutations,
      InspectionStep... inspections) {
    WorkbookResult response =
        execute(
            XlsxParitySupport.executionRootFor(workbookPath),
            ParityPlanSupport.request(
                existingWorkbookSource(workbookPath, sourceSecurity),
                new WorkbookPlan.WorkbookPersistence.None(),
                execution,
                formulaEnvironment,
                mutations,
                List.of(inspections)));
    if (response instanceof WorkbookResult.Failure failure) {
      return failure;
    }
    throw new AssertionError("Expected GridGrind failure but request succeeded: " + response);
  }

  static String savedPath(WorkbookResult.Success success) {
    return switch (success.persistence()) {
      case WorkbookResultPersistence.PersistenceOutcome.SavedAs savedAs ->
          writtenExecutionPath(savedAs.publication());
      case WorkbookResultPersistence.PersistenceOutcome.Overwritten overwritten ->
          writtenExecutionPath(overwritten.publication());
      case WorkbookResultPersistence.PersistenceOutcome.NotSaved _ ->
          throw new AssertionError("Expected the workbook to be persisted");
    };
  }

  private static String writtenExecutionPath(
      WorkbookResultPersistence.PublicationOutcome publication) {
    return switch (publication) {
      case WorkbookResultPersistence.PublicationOutcome.Published published ->
          published.executionPath();
      case WorkbookResultPersistence.PublicationOutcome.NotAttempted _,
          WorkbookResultPersistence.PublicationOutcome.NotPublished _,
          WorkbookResultPersistence.PublicationOutcome.Uncertain _ ->
          throw new AssertionError("Expected the workbook to be published");
    };
  }

  static <T extends InspectionResult> T read(
      WorkbookResult.Success success, String stepId, Class<T> type) {
    return type.cast(
        success.inspections().stream()
            .filter(result -> result.stepId().value().equals(stepId))
            .findFirst()
            .filter(type::isInstance)
            .orElseThrow(
                () -> new AssertionError("Missing read result " + stepId + " as " + type)));
  }

  private static WorkbookResult.Success success(WorkbookResult response) {
    if (response instanceof WorkbookResult.Success success) {
      return success;
    }
    throw new AssertionError("Expected GridGrind success but got " + response);
  }

  private static boolean hasType(List<TypeEntry> entries, String id) {
    return entries.stream().map(TypeEntry::id).anyMatch(id::equals);
  }
}
