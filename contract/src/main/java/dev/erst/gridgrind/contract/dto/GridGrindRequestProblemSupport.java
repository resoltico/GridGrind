package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.contract.json.ActionableInvariantMessage;
import dev.erst.gridgrind.contract.json.ActionableShapeMessage;
import dev.erst.gridgrind.contract.json.DuplicateStepId;
import dev.erst.gridgrind.contract.json.ExplicitNullField;
import dev.erst.gridgrind.contract.json.FieldValidationProblem;
import dev.erst.gridgrind.contract.json.MessageInvariant;
import dev.erst.gridgrind.contract.json.MessageShape;
import dev.erst.gridgrind.contract.json.MissingRequiredField;
import dev.erst.gridgrind.contract.json.MissingTypeDiscriminator;
import dev.erst.gridgrind.contract.json.NonXlsxPath;
import dev.erst.gridgrind.contract.json.RequestProblemDescriptor;
import dev.erst.gridgrind.contract.json.UnknownField;
import dev.erst.gridgrind.contract.json.UnknownTypeValue;
import dev.erst.gridgrind.contract.json.UnsupportedValue;
import java.util.Objects;
import java.util.Optional;

/** Shared request-problem path and resolution helpers used by CLI and doctor/report surfaces. */
public final class GridGrindRequestProblemSupport {
  private GridGrindRequestProblemSupport() {}

  /** Returns the canonical public wording for one missing required request field. */
  public static String missingRequiredFieldMessage(String jsonPath) {
    return message(new MissingRequiredField(jsonPath));
  }

  /** Returns the canonical public wording for one explicit-null request field. */
  public static String explicitNullFieldMessage(String jsonPath) {
    return message(new ExplicitNullField(jsonPath));
  }

  /** Returns the canonical public wording for one typed request problem. */
  public static String message(RequestProblemDescriptor descriptor) {
    Objects.requireNonNull(descriptor, "descriptor must not be null");
    return switch (descriptor) {
      case MissingRequiredField missingRequiredField ->
          "Missing required field '" + missingRequiredField.jsonPathValue() + "'";
      case MissingTypeDiscriminator missingTypeDiscriminator ->
          "Missing required field '" + missingTypeDiscriminator.jsonPathValue() + "'";
      case UnknownField unknownField -> "Unknown field '" + unknownField.jsonPathValue() + "'";
      case UnknownTypeValue unknownTypeValue -> renderUnknownTypeValueMessage(unknownTypeValue);
      case UnsupportedValue unsupportedValue -> renderUnsupportedValueMessage(unsupportedValue);
      case ExplicitNullField explicitNullField ->
          "Field '"
              + explicitNullField.jsonPathValue()
              + "' must be omitted when absent; explicit null is not accepted.";
      case ActionableShapeMessage actionableShapeMessage -> actionableShapeMessage.message();
      case MessageShape messageShape -> messageShape.message();
      case DuplicateStepId duplicateStepId ->
          "steps must not contain duplicate stepId values: " + duplicateStepId.value();
      case NonXlsxPath nonXlsxPath ->
          "path must end in .xlsx (got: '" + nonXlsxPath.actualExtension() + "')";
      case FieldValidationProblem fieldValidationProblem -> fieldValidationProblem.message();
      case ActionableInvariantMessage actionableInvariantMessage ->
          actionableInvariantMessage.message();
      case MessageInvariant messageInvariant -> messageInvariant.message();
    };
  }

  /** Returns one request-problem-specific remediation sentence for the typed request problem. */
  public static String resolution(RequestProblemDescriptor descriptor, ProblemContext context) {
    Objects.requireNonNull(descriptor, "descriptor must not be null");
    Objects.requireNonNull(context, "context must not be null");
    Optional<String> requestJsonPath = requestJsonPath(context);
    return switch (descriptor) {
      case MissingRequiredField missingRequiredField ->
          missingFieldResolution(
              preciseRequestPath(missingRequiredField.jsonPathValue(), requestJsonPath));
      case MissingTypeDiscriminator missingTypeDiscriminator ->
          "Add the required type discriminator at '"
              + preciseRequestPath(missingTypeDiscriminator.jsonPathValue(), requestJsonPath)
              + "'.";
      case UnknownField unknownField ->
          "Remove or rename unexpected field '"
              + preciseRequestPath(unknownField.jsonPathValue(), requestJsonPath)
              + "' so the request matches the protocol.";
      case UnknownTypeValue unknownTypeValue ->
          unknownTypeValue
              .jsonPath()
              .or(() -> requestJsonPath)
              .map(
                  jsonPath ->
                      "Replace field '"
                          + jsonPath
                          + "' with one supported type value. Use"
                          + " --print-protocol-catalog --lookup or --search when you need the"
                          + " allowed values.")
              .orElse("Replace the unknown type value with one supported by the protocol.");
      case UnsupportedValue unsupportedValue ->
          unsupportedValue
              .jsonPath()
              .map(
                  jsonPath ->
                      "Replace field '"
                          + jsonPath
                          + "' with one supported value. Use --print-protocol-catalog --lookup"
                          + " or --search when you need the allowed values.")
              .orElse("Replace the unsupported value with one allowed by the protocol.");
      case ExplicitNullField explicitNullField ->
          "Remove field '"
              + preciseRequestPath(explicitNullField.jsonPathValue(), requestJsonPath)
              + "' entirely when it is absent; explicit null is not part of the request"
              + " contract.";
      case ActionableShapeMessage actionableShapeMessage ->
          actionableShapeMessage.resolutionValue();
      case MessageShape messageShape ->
          messageShape
              .jsonPath()
              .or(() -> requestJsonPath)
              .map(
                  jsonPath ->
                      "Fix field '" + jsonPath + "' so it matches the published request shape.")
              .orElse("Fix the request payload so it matches the published request shape.");
      case DuplicateStepId duplicateStepId ->
          "Make every stepId unique. Rename or remove the duplicate value '"
              + duplicateStepId.value()
              + "'.";
      case NonXlsxPath nonXlsxPath ->
          nonXlsxPath
              .jsonPath()
              .map(
                  jsonPath ->
                      "Provide a path ending in .xlsx for field '"
                          + preciseRequestPath(jsonPath, requestJsonPath)
                          + "'.")
              .orElse("Provide a workbook path ending in .xlsx.");
      case FieldValidationProblem fieldValidationProblem -> fieldValidationProblem.resolution();
      case ActionableInvariantMessage actionableInvariantMessage ->
          actionableInvariantMessage.resolutionValue();
      case MessageInvariant messageInvariant ->
          messageInvariant
              .jsonPath()
              .or(() -> requestJsonPath)
              .map(jsonPath -> "Fix field '" + jsonPath + "' so it satisfies the request contract.")
              .orElse("Fix the request data so it satisfies the request contract.");
    };
  }

  private static Optional<String> requestJsonPath(ProblemContext context) {
    if (context instanceof ProblemContext.ReadRequest readRequest) {
      return readRequest.jsonPath();
    }
    return Optional.empty();
  }

  private static String missingFieldResolution(String jsonPath) {
    if ("protocolVersion".equals(jsonPath)) {
      return "Add protocolVersion: \"V2\" at the request root.";
    }
    if (jsonPath.endsWith(".type")) {
      return "Add the required type discriminator at '" + jsonPath + "'.";
    }
    return "Add required field '" + jsonPath + "' to the request payload.";
  }

  private static String preciseRequestPath(
      String descriptorJsonPath, Optional<String> requestJsonPath) {
    if (requestJsonPath.isEmpty()) {
      return descriptorJsonPath;
    }
    String contextJsonPath = requestJsonPath.orElseThrow();
    if (contextJsonPath.equals(descriptorJsonPath)
        || contextJsonPath.endsWith("." + descriptorJsonPath)
        || (descriptorJsonPath.startsWith("[") && contextJsonPath.endsWith(descriptorJsonPath))) {
      return contextJsonPath;
    }
    return descriptorJsonPath;
  }

  private static String renderUnknownTypeValueMessage(UnknownTypeValue unknownTypeValue) {
    StringBuilder message =
        new StringBuilder(64)
            .append("Unknown type value '")
            .append(unknownTypeValue.typeId())
            .append('\'');
    unknownTypeValue
        .specificGuidance()
        .ifPresent(guidance -> message.append("; ").append(guidance));
    if (!unknownTypeValue.similarValues().isEmpty()) {
      message
          .append("; similar valid values: ")
          .append(String.join(", ", unknownTypeValue.similarValues()));
    }
    return message.toString();
  }

  private static String renderUnsupportedValueMessage(UnsupportedValue unsupportedValue) {
    String allowedValues = String.join(", ", unsupportedValue.allowedValues());
    return unsupportedValue
        .jsonPath()
        .map(
            jsonPath ->
                "Unsupported value '"
                    + unsupportedValue.value()
                    + "' for field '"
                    + jsonPath
                    + "'; expected one of: "
                    + allowedValues)
        .orElse(
            "Unsupported value '"
                + unsupportedValue.value()
                + "'; expected one of: "
                + allowedValues);
  }
}
