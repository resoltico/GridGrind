package dev.erst.gridgrind.cli.examples;

import dev.erst.gridgrind.cli.discovery.ExampleWorkspaceMode;
import dev.erst.gridgrind.cli.discovery.TaskStarterContract;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Objects;

/** Shared paths and starter factories for published CLI task plans. */
final class TaskStarterPlanSupport {
  private static final String ONE_PIXEL_PNG_BASE64 =
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII=";

  private TaskStarterPlanSupport() {}

  static TaskStarterPlan selfContainedStarter(String taskId, WorkbookPlan plan) {
    return new TaskStarterPlan(
        taskId, TaskStarterContract.selfContained(taskRequestPath(taskId)), plan);
  }

  static TaskStarterPlan assetBackedStarter(
      String taskId, WorkbookPlan plan, String firstRequiredPath, String... otherRequiredPaths) {
    Objects.requireNonNull(firstRequiredPath, "firstRequiredPath must not be null");
    List<String> requiredPaths = new ArrayList<>(otherRequiredPaths.length + 1);
    requiredPaths.add(firstRequiredPath);
    requiredPaths.addAll(List.of(otherRequiredPaths));
    return new TaskStarterPlan(
        taskId,
        new TaskStarterContract(
            taskRequestPath(taskId),
            ExampleWorkspaceMode.REQUIRES_EXAMPLE_ASSETS,
            List.copyOf(requiredPaths)),
        plan);
  }

  static String taskPlanId(String taskId) {
    return slug(taskId) + "-starter";
  }

  static String taskWorkbookPath(String taskId) {
    return "generated-workbooks/" + slug(taskId) + ".xlsx";
  }

  static String taskStarterAsset(String fileName) {
    return "task-starter-assets/" + fileName;
  }

  static String onePixelPngBase64() {
    return ONE_PIXEL_PNG_BASE64;
  }

  private static String taskRequestPath(String taskId) {
    return "tasks/" + slug(taskId) + "-request.json";
  }

  private static String slug(String taskId) {
    return taskId.toLowerCase(Locale.ROOT).replace('_', '-');
  }
}
