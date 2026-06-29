package dev.erst.gridgrind.contract.dto;

import java.util.Objects;
import java.util.Optional;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/** Shared request-problem path and resolution helpers used by CLI and doctor/report surfaces. */
public final class GridGrindRequestProblemSupport {
  private static final Pattern MISSING_REQUIRED_FIELD_PATTERN =
      Pattern.compile("^Missing required field '([^']+)'$");
  private static final Pattern EXPLICIT_NULL_FIELD_PATTERN =
      Pattern.compile(
          "^Field '([^']+)' must be omitted when absent; explicit null is not accepted\\.$");
  private static final Pattern UNKNOWN_FIELD_PATTERN = Pattern.compile("^Unknown field '([^']+)'$");
  private static final Pattern UNKNOWN_TYPE_VALUE_PATTERN =
      Pattern.compile("^Unknown type value '([^']+)'");
  private static final Pattern UNSUPPORTED_VALUE_PATTERN =
      Pattern.compile("^Unsupported value '([^']+)' for field '([^']+)'");
  private static final Pattern DUPLICATE_STEP_ID_PATTERN =
      Pattern.compile("^steps must not contain duplicate stepId values: (.+)$");
  private static final Pattern NON_BLANK_FIELD_PATTERN =
      Pattern.compile("^([A-Za-z0-9.\\[\\]_]+) must not be blank$");

  private GridGrindRequestProblemSupport() {}

  /** Returns the canonical public wording for one missing required request field. */
  public static String missingRequiredFieldMessage(String jsonPath) {
    return "Missing required field '" + requireJsonPath(jsonPath) + "'";
  }

  /** Returns the canonical public wording for one explicit-null request field. */
  public static String explicitNullFieldMessage(String jsonPath) {
    return "Field '"
        + requireJsonPath(jsonPath)
        + "' must be omitted when absent; explicit null is not accepted.";
  }

  /** Returns whether one public request message reports one missing required field. */
  public static boolean isMissingRequiredFieldMessage(String message) {
    return MISSING_REQUIRED_FIELD_PATTERN
        .matcher(Objects.requireNonNullElse(message, "").trim())
        .matches();
  }

  /** Returns whether one public request message reports one explicit-null field. */
  public static boolean isExplicitNullFieldMessage(String message) {
    return EXPLICIT_NULL_FIELD_PATTERN
        .matcher(Objects.requireNonNullElse(message, "").trim())
        .matches();
  }

  /** Extracts one actionable JSON path when the public request error wording carries it. */
  public static Optional<String> jsonPathFromMessage(String message) {
    String normalized = Objects.requireNonNullElse(message, "").trim();
    if (normalized.isEmpty()) {
      return Optional.empty();
    }
    Matcher missingRequired = MISSING_REQUIRED_FIELD_PATTERN.matcher(normalized);
    if (missingRequired.matches()) {
      return Optional.of(missingRequired.group(1));
    }
    Matcher explicitNullField = EXPLICIT_NULL_FIELD_PATTERN.matcher(normalized);
    if (explicitNullField.matches()) {
      return Optional.of(explicitNullField.group(1));
    }
    Matcher unknownField = UNKNOWN_FIELD_PATTERN.matcher(normalized);
    if (unknownField.matches()) {
      return Optional.of(unknownField.group(1));
    }
    Matcher unsupportedValue = UNSUPPORTED_VALUE_PATTERN.matcher(normalized);
    if (unsupportedValue.matches()) {
      return Optional.of(unsupportedValue.group(2));
    }
    Matcher nonBlankField = NON_BLANK_FIELD_PATTERN.matcher(normalized);
    if (nonBlankField.matches()) {
      return Optional.of(nonBlankField.group(1));
    }
    return Optional.empty();
  }

  /** Returns one request-problem-specific remediation sentence when the message supports it. */
  public static Optional<String> specificResolution(
      GridGrindProblemCode code, String message, ProblemContext context) {
    Objects.requireNonNull(code, "code must not be null");
    Objects.requireNonNull(context, "context must not be null");
    return specificResolution(code, message, requestJsonPath(context));
  }

  /** Returns one request-problem-specific remediation sentence when the message supports it. */
  public static Optional<String> specificResolution(
      GridGrindProblemCode code, String message, Optional<String> jsonPath) {
    Objects.requireNonNull(code, "code must not be null");
    Objects.requireNonNull(jsonPath, "jsonPath must not be null");
    String normalizedMessage = Objects.requireNonNullElse(message, "").trim();
    if (normalizedMessage.isEmpty()) {
      return Optional.empty();
    }
    Matcher duplicateStepId = DUPLICATE_STEP_ID_PATTERN.matcher(normalizedMessage);
    if (duplicateStepId.matches()) {
      return Optional.of(
          "Make every stepId unique. Rename or remove the duplicate value '"
              + duplicateStepId.group(1)
              + "'.");
    }
    Matcher missingRequired = MISSING_REQUIRED_FIELD_PATTERN.matcher(normalizedMessage);
    if (missingRequired.matches()) {
      return Optional.of(missingFieldResolution(missingRequired.group(1)));
    }
    Matcher explicitNullField = EXPLICIT_NULL_FIELD_PATTERN.matcher(normalizedMessage);
    if (explicitNullField.matches()) {
      String field = explicitNullField.group(1);
      return Optional.of(
          "Remove field '"
              + field
              + "' entirely when it is absent; explicit null is not part of the request"
              + " contract.");
    }
    Matcher unknownField = UNKNOWN_FIELD_PATTERN.matcher(normalizedMessage);
    if (unknownField.matches()) {
      String field = unknownField.group(1);
      return Optional.of(
          "Remove or rename unexpected field '" + field + "' so the request matches the protocol.");
    }
    Matcher unsupportedValue = UNSUPPORTED_VALUE_PATTERN.matcher(normalizedMessage);
    if (unsupportedValue.matches()) {
      String field = unsupportedValue.group(2);
      return Optional.of(
          "Replace field '"
              + field
              + "' with one supported value. Use --print-protocol-catalog --lookup or --search"
              + " when you need the allowed values.");
    }
    Matcher nonBlankField = NON_BLANK_FIELD_PATTERN.matcher(normalizedMessage);
    if (nonBlankField.matches()) {
      String field = nonBlankField.group(1);
      return Optional.of("Provide a non-blank value for field '" + field + "'.");
    }
    if (code == GridGrindProblemCode.INVALID_REQUEST_SHAPE && jsonPath.isPresent()) {
      if (UNKNOWN_TYPE_VALUE_PATTERN.matcher(normalizedMessage).find()) {
        String field = jsonPath.orElseThrow();
        return Optional.of(
            "Replace field '"
                + field
                + "' with one supported type value. Use --print-protocol-catalog --lookup or"
                + " --search when you need the allowed values.");
      }
      return Optional.of(
          "Fix field '" + jsonPath.orElseThrow() + "' so it matches the published request shape.");
    }
    return Optional.empty();
  }

  /** Returns whether one public request message describes a structural shape mismatch. */
  public static boolean looksLikeRequestShapeViolation(String message) {
    String normalizedMessage = Objects.requireNonNullElse(message, "").trim();
    if (normalizedMessage.isEmpty()) {
      return false;
    }
    return MISSING_REQUIRED_FIELD_PATTERN.matcher(normalizedMessage).matches()
        || EXPLICIT_NULL_FIELD_PATTERN.matcher(normalizedMessage).matches()
        || UNKNOWN_FIELD_PATTERN.matcher(normalizedMessage).matches()
        || UNKNOWN_TYPE_VALUE_PATTERN.matcher(normalizedMessage).find()
        || UNSUPPORTED_VALUE_PATTERN.matcher(normalizedMessage).find();
  }

  private static Optional<String> requestJsonPath(ProblemContext context) {
    if (context instanceof ProblemContext.ReadRequest readRequest) {
      return readRequest.jsonPath();
    }
    return Optional.empty();
  }

  private static String missingFieldResolution(String jsonPath) {
    if ("protocolVersion".equals(jsonPath)) {
      return "Add protocolVersion: \"V1\" at the request root.";
    }
    if (jsonPath.endsWith(".type")) {
      return "Add the required type discriminator at '" + jsonPath + "'.";
    }
    return "Add required field '" + jsonPath + "' to the request payload.";
  }

  private static String requireJsonPath(String jsonPath) {
    Objects.requireNonNull(jsonPath, "jsonPath must not be null");
    if (jsonPath.isBlank()) {
      throw new IllegalArgumentException("jsonPath must not be blank");
    }
    return jsonPath;
  }
}
