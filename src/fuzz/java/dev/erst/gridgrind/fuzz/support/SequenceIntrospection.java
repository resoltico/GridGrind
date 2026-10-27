package dev.erst.gridgrind.fuzz.support;

import dev.erst.gridgrind.core.plan.AssertionStep;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.ResolvedStep;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.engine.backend.xssf.*;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/** Summarizes generated workflow and command shapes for reports, replay, and telemetry. */
public final class SequenceIntrospection {
  private SequenceIntrospection() {}

  /** Counts protocol mutation kinds by stable wire-facing action name. */
  public static Map<String, Long> mutationKinds(List<MutationStep> mutations) {
    Objects.requireNonNull(mutations, "mutations must not be null");
    LinkedHashMap<String, Long> counts = new LinkedHashMap<>();
    mutations.forEach(mutation -> increment(counts, mutationKind(mutation)));
    return Map.copyOf(counts);
  }

  /** Counts protocol inspection kinds by stable wire-facing query name. */
  public static Map<String, Long> inspectionKinds(List<InspectionStep> inspections) {
    Objects.requireNonNull(inspections, "inspections must not be null");
    LinkedHashMap<String, Long> counts = new LinkedHashMap<>();
    inspections.forEach(inspection -> increment(counts, inspectionKind(inspection)));
    return Map.copyOf(counts);
  }

  /** Counts protocol assertion kinds by stable wire-facing assertion name. */
  public static Map<String, Long> assertionKinds(List<AssertionStep> assertions) {
    Objects.requireNonNull(assertions, "assertions must not be null");
    LinkedHashMap<String, Long> counts = new LinkedHashMap<>();
    assertions.forEach(assertion -> increment(counts, assertionKind(assertion)));
    return Map.copyOf(counts);
  }

  /** Counts workbook command kinds by stable engine-facing name. */
  public static Map<String, Long> commandKinds(List<ResolvedStep> commands) {
    Objects.requireNonNull(commands, "commands must not be null");
    LinkedHashMap<String, Long> counts = new LinkedHashMap<>();
    commands.forEach(command -> increment(counts, commandKind(command)));
    return Map.copyOf(counts);
  }

  /** Counts style-attribute labels observed in protocol style mutations. */
  public static Map<String, Long> styleKinds(List<MutationStep> mutations) {
    Objects.requireNonNull(mutations, "mutations must not be null");
    LinkedHashMap<String, Long> counts = new LinkedHashMap<>();
    mutations.forEach(
        mutation -> {
          if (mutation.action()
              instanceof dev.erst.gridgrind.core.plan.CellMutationAction.ApplyStyle applyStyle) {
            StyleKindIntrospection.styleKinds(applyStyle.style())
                .forEach((key, value) -> counts.merge(key, value, Long::sum));
          }
        });
    return Map.copyOf(counts);
  }

  /** Counts style-attribute labels observed in engine style commands. */
  public static Map<String, Long> styleKindsFromCommands(List<ResolvedStep> commands) {
    Objects.requireNonNull(commands, "commands must not be null");
    LinkedHashMap<String, Long> counts = new LinkedHashMap<>();
    commands.forEach(
        command -> {
          if (command.action()
              instanceof dev.erst.gridgrind.core.plan.CellMutationAction.ApplyStyle applyStyle) {
            StyleKindIntrospection.styleKinds(applyStyle.style())
                .forEach((key, value) -> counts.merge(key, value, Long::sum));
          }
        });
    return Map.copyOf(counts);
  }

  /** Returns the response family label used in telemetry and summaries. */
  public static String responseKind(WorkbookResult response) {
    Objects.requireNonNull(response, "response must not be null");
    return switch (response) {
      case WorkbookResult.Success _ -> "SUCCESS";
      case WorkbookResult.Failure _ -> "FAILURE";
    };
  }

  /** Counts ordered inspection steps requested by one authored workflow. */
  public static int inspectionCount(WorkbookPlan request) {
    Objects.requireNonNull(request, "request must not be null");
    return request.stepPartition().inspections().size();
  }

  /** Counts ordered assertion steps requested by one authored workflow. */
  public static int assertionCount(WorkbookPlan request) {
    Objects.requireNonNull(request, "request must not be null");
    return request.stepPartition().assertions().size();
  }

  /** Returns the stable workbook source-type label for one request. */
  public static String sourceKind(WorkbookPlan request) {
    Objects.requireNonNull(request, "request must not be null");
    return switch (request.source()) {
      case WorkbookPlan.WorkbookSource.New _ -> "NEW";
      case WorkbookPlan.WorkbookSource.ExistingFile _ -> "EXISTING";
    };
  }

  /** Returns the stable workbook persistence-type label for one request. */
  public static String persistenceKind(WorkbookPlan request) {
    Objects.requireNonNull(request, "request must not be null");
    return switch (request.persistence()) {
      case WorkbookPlan.WorkbookPersistence.None _ -> "NONE";
      case WorkbookPlan.WorkbookPersistence.Overwrite _ -> "OVERWRITE";
      case WorkbookPlan.WorkbookPersistence.SaveAs _ -> "SAVE_AS";
    };
  }

  /** Returns the stable wire-facing name for one protocol mutation step. */
  static String mutationKind(MutationStep mutation) {
    Objects.requireNonNull(mutation, "mutation must not be null");
    return mutation.action().actionType();
  }

  /** Returns the stable engine-facing name for one workbook command variant. */
  static String commandKind(ResolvedStep command) {
    Objects.requireNonNull(command, "command must not be null");
    return dev.erst.gridgrind.core.catalog.GridGrindProtocolTypeNames.mutationActionTypeName(
        command.action().authoredType());
  }

  /** Returns the stable wire-facing name for one protocol inspection step. */
  static String inspectionKind(InspectionStep inspection) {
    Objects.requireNonNull(inspection, "inspection must not be null");
    return inspection.query().queryType();
  }

  /** Returns the stable wire-facing name for one protocol assertion step. */
  static String assertionKind(AssertionStep assertion) {
    Objects.requireNonNull(assertion, "assertion must not be null");
    return assertion.assertion().assertionType();
  }

  private static void increment(Map<String, Long> counts, String key) {
    counts.merge(key, 1L, Long::sum);
  }
}
