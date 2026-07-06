package dev.erst.gridgrind.cli.examples;

import java.util.Locale;
import java.util.Objects;

/** Shared path and asset conventions for published task-starter recipes. */
final class TaskStarterRecipeSupport {
  private static final String ONE_PIXEL_PNG_BASE64 =
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII=";

  private TaskStarterRecipeSupport() {}

  static String taskRequestFileName(String taskId) {
    return slug(taskId) + "-request.json";
  }

  static String taskPlanId(String taskId) {
    return slug(taskId) + "-starter";
  }

  static String taskWorkbookPath(String taskId) {
    return "generated-workbooks/" + slug(taskId) + ".xlsx";
  }

  static String taskStarterAsset(String fileName) {
    Objects.requireNonNull(fileName, "fileName must not be null");
    if (fileName.isBlank()) {
      throw new IllegalArgumentException("fileName must not be blank");
    }
    return "task-starter-assets/" + fileName;
  }

  static String onePixelPngBase64() {
    return ONE_PIXEL_PNG_BASE64;
  }

  private static String slug(String taskId) {
    Objects.requireNonNull(taskId, "taskId must not be null");
    if (taskId.isBlank()) {
      throw new IllegalArgumentException("taskId must not be blank");
    }
    return taskId.toLowerCase(Locale.ROOT).replace('_', '-');
  }
}
