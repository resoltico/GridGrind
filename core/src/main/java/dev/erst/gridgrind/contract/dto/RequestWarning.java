package dev.erst.gridgrind.contract.dto;

import java.util.Objects;

/** One non-fatal request warning with a truthful step or request-path location. */
public record RequestWarning(
    GridGrindWarningCode code, RequestWarningLocation location, String message) {
  public RequestWarning {
    Objects.requireNonNull(code, "code must not be null");
    Objects.requireNonNull(location, "location must not be null");
    Objects.requireNonNull(message, "message must not be null");
    if (message.isBlank()) {
      throw new IllegalArgumentException("message must not be blank");
    }
  }

  /** Creates a warning attached to one authored workbook step. */
  public RequestWarning(
      GridGrindWarningCode code, int stepIndex, String stepId, String stepType, String message) {
    this(code, new RequestWarningLocation.Step(stepIndex, stepId, stepType), message);
  }

  /** Creates the portability warning for one contained absolute request-owned path. */
  public static RequestWarning nonPortableAbsolutePath(String path, String pathRole) {
    return new RequestWarning(
        GridGrindWarningCode.NON_PORTABLE_ABSOLUTE_PATH,
        new RequestWarningLocation.RequestPath(path, pathRole),
        "Absolute request-owned paths are portable only when the same execution-root layout exists: "
            + path);
  }

  /** Creates the warning for one accepted leading UTF-8 byte-order mark. */
  public static RequestWarning utf8BomIgnored() {
    return new RequestWarning(
        GridGrindWarningCode.UTF8_BOM_IGNORED,
        new RequestWarningLocation.RequestByteOffset(0),
        "Ignored one leading UTF-8 byte-order mark.");
  }
}
