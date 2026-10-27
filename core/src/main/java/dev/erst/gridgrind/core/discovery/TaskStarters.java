package dev.erst.gridgrind.core.discovery;

import dev.erst.gridgrind.core.vocabulary.PlanId;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.Locale;

/** Shared path and asset conventions for published task-starter recipes. */
final class TaskStarters {
  private static final String ONE_PIXEL_PNG_BASE64 =
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVQI12P4//8/AAX+Av7czFnnAAAAAElFTkSuQmCC";

  private TaskStarters() {}

  static String taskRequestFileName(String taskId) {
    return slug(taskId) + "-request.json";
  }

  static PlanId taskPlanId(String taskId) {
    return new PlanId(slug(taskId) + "-starter");
  }

  static String taskWorkbookPath(String taskId) {
    return "generated-workbooks/" + slug(taskId) + ".xlsx";
  }

  static String taskStarterAsset(String fileName) {
    TextValidation.requireNonBlank(fileName, "fileName");
    return "task-starter-assets/" + fileName;
  }

  static String onePixelPngBase64() {
    return ONE_PIXEL_PNG_BASE64;
  }

  private static String slug(String taskId) {
    TextValidation.requireNonBlank(taskId, "taskId");
    return taskId.toLowerCase(Locale.ROOT).replace('_', '-');
  }
}
