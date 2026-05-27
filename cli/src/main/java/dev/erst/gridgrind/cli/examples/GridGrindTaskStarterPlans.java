package dev.erst.gridgrind.cli.examples;

import dev.erst.gridgrind.cli.discovery.TaskStarterContract;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;

/** Executable task starter scenarios published by the CLI for official task ids. */
public final class GridGrindTaskStarterPlans {
  private static final Map<String, TaskStarterPlan> STARTERS = buildStarterMap();

  private GridGrindTaskStarterPlans() {}

  /** Returns the public starter contract for one official stable task id. */
  public static TaskStarterContract contractFor(String taskId) {
    Objects.requireNonNull(taskId, "taskId must not be null");
    TaskStarterPlan starter = STARTERS.get(taskId);
    if (starter == null) {
      throw new IllegalStateException("Missing task starter contract for " + taskId);
    }
    return starter.contract();
  }

  /** Returns the executable starter request for one official stable task id. */
  public static WorkbookPlan planFor(String taskId) {
    Objects.requireNonNull(taskId, "taskId must not be null");
    TaskStarterPlan starter = STARTERS.get(taskId);
    if (starter == null) {
      throw new IllegalStateException("Missing task starter plan for " + taskId);
    }
    return starter.plan();
  }

  /** Returns the starter request when one official task starter is defined for the id. */
  public static Optional<WorkbookPlan> findPlan(String taskId) {
    Objects.requireNonNull(taskId, "taskId must not be null");
    return Optional.ofNullable(STARTERS.get(taskId)).map(TaskStarterPlan::plan);
  }

  private static Map<String, TaskStarterPlan> buildStarterMap() {
    List<TaskStarterPlan> plans =
        java.util.stream.Stream.of(
                GridGrindTaskStarterReportPlans.starters(),
                GridGrindTaskStarterWorkflowPlans.starters(),
                GridGrindTaskStarterAssetPlans.starters())
            .flatMap(List::stream)
            .toList();
    return toStarterMap(plans);
  }

  static Map<String, TaskStarterPlan> toStarterMap(List<TaskStarterPlan> plans) {
    Objects.requireNonNull(plans, "plans must not be null");
    return java.util.Collections.unmodifiableMap(
        plans.stream()
            .collect(
                java.util.stream.Collectors.toMap(
                    TaskStarterPlan::taskId,
                    java.util.function.Function.identity(),
                    GridGrindTaskStarterPlans::duplicateTaskStarter,
                    java.util.LinkedHashMap::new)));
  }

  private static TaskStarterPlan duplicateTaskStarter(TaskStarterPlan left, TaskStarterPlan right) {
    throw new IllegalStateException(
        "Duplicate task starter plan for " + left.taskId() + " and " + right.taskId());
  }
}
