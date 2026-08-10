package dev.erst.gridgrind.contract.json;

import java.io.IOException;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import org.jspecify.annotations.Nullable;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.node.ArrayNode;
import tools.jackson.databind.node.ObjectNode;

/** Protects diagnostics for request fields declared secret by their exact owner path. */
public final class RequestDiagnosticRedactor {
  private static final String REDACTED_VALUE = "[REDACTED]";
  private static final String SENSITIVE_VALUE_MESSAGE = "Sensitive request value is invalid.";
  private final Set<String> secretFieldPaths;

  private RequestDiagnosticRedactor(Set<String> secretFieldPaths) {
    this.secretFieldPaths = Set.copyOf(secretFieldPaths);
  }

  static RequestDiagnosticRedactor empty() {
    return new RequestDiagnosticRedactor(Set.of());
  }

  /**
   * Redacts only problem objects whose structured request context identifies a secret field.
   *
   * <p>Whole-document lexical substitution is intentionally forbidden: a short password can be a
   * substring of ordinary workbook data, filenames, and diagnostic prose. Request binding instead
   * produces a sensitive-safe message before output, and this method provides a second, path-owned
   * output boundary for future diagnostic carriers.
   */
  public byte[] redactSerializedJson(byte[] payload, boolean pretty) throws IOException {
    Objects.requireNonNull(payload, "payload must not be null");
    if (secretFieldPaths.isEmpty()) {
      return payload;
    }
    JsonNode root = GridGrindJsonMapperSupport.WIRE_JSON_MAPPER.readTree(payload);
    if (!redactProblemNodes(root)) {
      return payload;
    }
    return GridGrindJsonCodecSupport.writeBytes(
        pretty
            ? GridGrindJsonMapperSupport.PRETTY_WIRE_JSON_MAPPER
            : GridGrindJsonMapperSupport.WIRE_JSON_MAPPER,
        root);
  }

  static String safeValidationFailureMessage(
      @Nullable String message, Class<?> requestType, Optional<String> jsonPath) {
    Objects.requireNonNull(requestType, "requestType must not be null");
    return forRequestType(requestType).safeBindingFailureMessage(message, jsonPath);
  }

  String safeBindingFailureMessage(@Nullable String message, Optional<String> jsonPath) {
    Objects.requireNonNull(jsonPath, "jsonPath must not be null");
    if (jsonPath
        .filter(path -> RequestSecretFieldPaths.contains(secretFieldPaths, path))
        .isPresent()) {
      return SENSITIVE_VALUE_MESSAGE;
    }
    return message == null || message.isBlank()
        ? "Request value violates the request contract."
        : message;
  }

  static RequestDiagnosticRedactor forRequestType(Class<?> requestType) {
    return new RequestDiagnosticRedactor(RequestSecretFieldPaths.forRequestType(requestType));
  }

  private boolean redactProblemNodes(JsonNode node) {
    return switch (node) {
      case ObjectNode object -> redactObject(object);
      case ArrayNode array -> redactArray(array);
      default -> false;
    };
  }

  private boolean redactObject(ObjectNode object) {
    boolean changed = redactProblemIfSensitive(object);
    for (var property : object.properties()) {
      changed |= redactProblemNodes(property.getValue());
    }
    return changed;
  }

  private boolean redactArray(ArrayNode array) {
    boolean changed = false;
    for (JsonNode element : array) {
      changed |= redactProblemNodes(element);
    }
    return changed;
  }

  private boolean redactProblemIfSensitive(ObjectNode object) {
    Optional<String> requestPath = requestPath(object.path("context"));
    if (!isSecretRequestPath(requestPath)) {
      return false;
    }
    boolean changed = redactTextualProperty(object, "message");
    changed |= redactTextualProperty(object, "resolution");
    JsonNode causes = object.path("causes");
    if (causes instanceof ArrayNode causeArray) {
      for (JsonNode cause : causeArray) {
        if (cause instanceof ObjectNode causeObject) {
          changed |= redactTextualProperty(causeObject, "message");
        }
      }
    }
    return changed;
  }

  private static Optional<String> requestPath(JsonNode context) {
    JsonNode json = context.path("json");
    JsonNode jsonPath = json.path("jsonPath");
    return jsonPath.isTextual() ? Optional.of(jsonPath.textValue()) : Optional.empty();
  }

  private boolean isSecretRequestPath(Optional<String> requestPath) {
    return requestPath
        .filter(path -> RequestSecretFieldPaths.contains(secretFieldPaths, path))
        .isPresent();
  }

  private static boolean redactTextualProperty(ObjectNode object, String propertyName) {
    JsonNode value = object.get(propertyName);
    if (value == null || !value.isTextual()) {
      return false;
    }
    object.put(propertyName, REDACTED_VALUE);
    return true;
  }
}
