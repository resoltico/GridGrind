package dev.erst.gridgrind.core.problem;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.core.vocabulary.TextValidation;

/** The authored request surface to which one non-fatal warning applies. */
@Union(discriminator = "type", catalog = true, displayOrder = 47)
public sealed interface RequestWarningLocation
    permits dev.erst.gridgrind.core.problem.RequestWarningLocation.Step,
        dev.erst.gridgrind.core.problem.RequestWarningLocation.RequestPath,
        dev.erst.gridgrind.core.problem.RequestWarningLocation.RequestByteOffset,
        dev.erst.gridgrind.core.problem.RequestWarningLocation.FormulaCell {
  /** Returns the step index for ordered step warnings, or {@code -1} for request-level warnings. */
  default int orderingStepIndex() {
    return this instanceof Step step ? step.stepIndex() : -1;
  }

  /** One warning attached to a concrete authored workbook step. */
  @Variant(id = "STEP", summary = "Warning located at one authored workbook step.")
  record Step(int stepIndex, StepId stepId, String stepType) implements RequestWarningLocation {
    public Step {
      if (stepIndex < 0) {
        throw new IllegalArgumentException("stepIndex must not be negative");
      }

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      stepType = TextValidation.requireNonBlank(stepType, "stepType");
    }
  }

  /** One warning attached to a request-owned filesystem path. */
  @Variant(id = "REQUEST_PATH", summary = "Warning located at one request-owned filesystem path.")
  record RequestPath(String path, String pathRole) implements RequestWarningLocation {
    public RequestPath {
      path = TextValidation.requireNonBlank(path, "path");
      pathRole = TextValidation.requireNonBlank(pathRole, "pathRole");
    }
  }

  /** One warning attached to an exact byte position in the authored request stream. */
  @Variant(
      id = "REQUEST_BYTE_OFFSET",
      summary = "Warning located at one authored request byte offset.")
  record RequestByteOffset(long byteOffset) implements RequestWarningLocation {
    public RequestByteOffset {
      if (byteOffset < 0) {
        throw new IllegalArgumentException("byteOffset must not be negative");
      }
    }
  }

  /** One warning attached to a formula cell identified during calculation capability analysis. */
  @Variant(
      id = "FORMULA_CELL",
      summary = "Warning located at one formula cell assessed by calculation capability analysis.")
  record FormulaCell(SheetLabel sheetName, String address, String formula)
      implements RequestWarningLocation {
    public FormulaCell {
      java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
      address = TextValidation.requireNonBlank(address, "address");
      formula = TextValidation.requireNonBlank(formula, "formula");
    }
  }
}
