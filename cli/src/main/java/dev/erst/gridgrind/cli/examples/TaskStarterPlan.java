package dev.erst.gridgrind.cli.examples;

import dev.erst.gridgrind.cli.discovery.TaskStarterContract;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.Objects;

/** Executable starter scenario plus its published CLI contract for one stable task id. */
record TaskStarterPlan(String taskId, TaskStarterContract contract, WorkbookPlan plan) {
  TaskStarterPlan {
    Objects.requireNonNull(taskId, "taskId must not be null");
    if (taskId.isBlank()) {
      throw new IllegalArgumentException("taskId must not be blank");
    }
    Objects.requireNonNull(contract, "contract must not be null");
    Objects.requireNonNull(plan, "plan must not be null");
  }
}
