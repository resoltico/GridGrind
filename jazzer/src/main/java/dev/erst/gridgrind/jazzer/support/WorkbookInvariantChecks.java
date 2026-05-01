package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import java.util.Base64;

/** Validates protocol responses and workbook state without depending on JUnit assertions. */
public final class WorkbookInvariantChecks {
  private WorkbookInvariantChecks() {}

  /** Requires the response shape to satisfy the protocol invariants the fuzzers rely on. */
  public static void requireResponseShape(GridGrindResponse response) {
    WorkbookInvariantResponseChecks.requireResponseShape(response);
  }

  /**
   * Requires the response shape to agree with the request's source, reads, and persistence
   * contract.
   */
  public static void requireWorkflowOutcomeShape(WorkbookPlan request, GridGrindResponse response) {
    WorkbookInvariantResponseChecks.requireWorkflowOutcomeShape(request, response);
  }

  /** Requires the open workbook to satisfy the structural invariants the fuzzers rely on. */
  public static void requireWorkbookShape(ExcelWorkbook workbook) {
    WorkbookInvariantWorkbookChecks.requireWorkbookShape(workbook);
  }

  static void require(boolean condition, String message) {
    if (!condition) {
      throw new IllegalStateException(message);
    }
  }

  static void requireNonBlank(String value, String fieldName) {
    require(value != null, fieldName + " must not be null");
    require(!value.isBlank(), fieldName + " must not be blank");
  }

  static void requireBase64(String value, String fieldName) {
    require(value != null, fieldName + " must not be null");
    try {
      Base64.getDecoder().decode(value);
    } catch (IllegalArgumentException exception) {
      throw new IllegalStateException(fieldName + " must be valid base64", exception);
    }
  }

  static void requireNonBlank(TextSourceInput value, String fieldName) {
    require(value != null, fieldName + " must not be null");
    if (value instanceof TextSourceInput.Inline inline) {
      require(!inline.text().isBlank(), fieldName + " must not be blank");
    }
  }
}
