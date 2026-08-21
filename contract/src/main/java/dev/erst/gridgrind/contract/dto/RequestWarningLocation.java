package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.Objects;

/** The authored request surface to which one non-fatal warning applies. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = RequestWarningLocation.Step.class, name = "STEP"),
  @JsonSubTypes.Type(value = RequestWarningLocation.RequestPath.class, name = "REQUEST_PATH"),
  @JsonSubTypes.Type(value = RequestWarningLocation.FormulaCell.class, name = "FORMULA_CELL")
})
public sealed interface RequestWarningLocation
    permits RequestWarningLocation.Step,
        RequestWarningLocation.RequestPath,
        RequestWarningLocation.FormulaCell {
  /** Returns the step index for ordered step warnings, or {@code -1} for request-level warnings. */
  default int orderingStepIndex() {
    return switch (this) {
      case Step step -> step.stepIndex();
      case RequestPath _ -> -1;
      case FormulaCell _ -> -1;
    };
  }

  /** One warning attached to a concrete authored workbook step. */
  record Step(int stepIndex, String stepId, String stepType) implements RequestWarningLocation {
    public Step {
      if (stepIndex < 0) {
        throw new IllegalArgumentException("stepIndex must not be negative");
      }
      stepId = requireNonBlank(stepId, "stepId");
      stepType = requireNonBlank(stepType, "stepType");
    }
  }

  /** One warning attached to a request-owned filesystem path. */
  record RequestPath(String path, String pathRole) implements RequestWarningLocation {
    public RequestPath {
      path = requireNonBlank(path, "path");
      pathRole = requireNonBlank(pathRole, "pathRole");
    }
  }

  /** One warning attached to a formula cell identified during calculation capability analysis. */
  record FormulaCell(String sheetName, String address, String formula)
      implements RequestWarningLocation {
    public FormulaCell {
      sheetName = requireNonBlank(sheetName, "sheetName");
      address = requireNonBlank(address, "address");
      formula = requireNonBlank(formula, "formula");
    }
  }

  private static String requireNonBlank(String value, String name) {
    Objects.requireNonNull(value, name + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(name + " must not be blank");
    }
    return value;
  }
}
