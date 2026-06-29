package dev.erst.gridgrind.contract.json;

import dev.erst.gridgrind.contract.dto.GridGrindRequestProblemSupport;
import java.io.IOException;
import java.io.InputStream;
import java.util.Objects;
import java.util.Optional;
import java.util.function.Function;
import tools.jackson.core.JacksonException;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.json.JsonMapper;

/** Owns tree decode mechanics and explicit-null rejection for the protocol JSON seam. */
final class GridGrindJsonCodecSupport {
  private GridGrindJsonCodecSupport() {}

  static <T> T readValue(
      InputStream inputStream,
      JsonMapper mapper,
      Class<T> targetType,
      Function<JacksonException, IllegalArgumentException> failureMapper)
      throws IOException {
    return decodeValue(
        readTree(inputStream, mapper, failureMapper), mapper, targetType, failureMapper);
  }

  static <T> T readValue(
      byte[] bytes,
      JsonMapper mapper,
      Class<T> targetType,
      Function<JacksonException, IllegalArgumentException> failureMapper)
      throws IOException {
    return decodeValue(readTree(bytes, mapper, failureMapper), mapper, targetType, failureMapper);
  }

  static <T> T decodeTree(
      JsonNode node,
      JsonMapper mapper,
      Class<T> targetType,
      Function<JacksonException, IllegalArgumentException> failureMapper)
      throws IOException {
    return decodeValue(node, mapper, targetType, failureMapper);
  }

  static byte[] writeBytes(JsonMapper mapper, Object value) throws IOException {
    return mapper.writeValueAsBytes(value);
  }

  static JsonNode readTree(
      InputStream inputStream,
      JsonMapper mapper,
      Function<JacksonException, IllegalArgumentException> failureMapper)
      throws IOException {
    try {
      return requirePresentDocument(mapper.readTree(inputStream));
    } catch (JacksonException exception) {
      throw failureMapper.apply(exception);
    }
  }

  static JsonNode readTree(
      byte[] bytes,
      JsonMapper mapper,
      Function<JacksonException, IllegalArgumentException> failureMapper)
      throws IOException {
    try {
      return requirePresentDocument(mapper.readTree(bytes));
    } catch (JacksonException exception) {
      throw failureMapper.apply(exception);
    }
  }

  private static <T> T decodeValue(
      JsonNode node,
      JsonMapper mapper,
      Class<T> targetType,
      Function<JacksonException, IllegalArgumentException> failureMapper)
      throws IOException {
    requireNonNullRoot(node);
    try {
      rejectExplicitNullMembers(node, "");
      GridGrindJsonStepPayloadShapeSupport.rejectInvalidStepPayloadShapes(node);
      return mapper.treeToValue(node, targetType);
    } catch (JacksonException exception) {
      throw normalizeMissingRequiredFailure(node, failureMapper.apply(exception));
    }
  }

  private static IllegalArgumentException normalizeMissingRequiredFailure(
      JsonNode node, IllegalArgumentException failure) {
    if (!(failure instanceof InvalidRequestShapeException shapeFailure)) {
      return failure;
    }
    String message = Objects.requireNonNullElse(shapeFailure.getMessage(), "");
    if (!GridGrindRequestProblemSupport.isExplicitNullFieldMessage(message)) {
      return failure;
    }
    String messagePath = GridGrindRequestProblemSupport.jsonPathFromMessage(message).orElseThrow();
    if (GridGrindJsonPathSupport.pathExists(node, messagePath)) {
      return failure;
    }
    String missingRequiredMessage =
        GridGrindRequestProblemSupport.missingRequiredFieldMessage(messagePath);
    return new InvalidRequestShapeException(
        missingRequiredMessage,
        shapeFailure.jsonPath().or(() -> Optional.of(messagePath)),
        shapeFailure.jsonLine(),
        shapeFailure.jsonColumn(),
        shapeFailure.getCause());
  }

  private static void rejectExplicitNullMembers(JsonNode node, String path) {
    if (node.isObject()) {
      for (var entry : node.properties()) {
        String childPath = path.isEmpty() ? entry.getKey() : path + "." + entry.getKey();
        if (entry.getValue().isNull()) {
          throw new InvalidRequestShapeException(
              GridGrindRequestProblemSupport.explicitNullFieldMessage(childPath),
              Optional.of(childPath),
              Optional.empty(),
              Optional.empty(),
              null);
        }
        rejectExplicitNullMembers(entry.getValue(), childPath);
      }
      return;
    }
    if (node.isArray()) {
      for (int index = 0; index < node.size(); index++) {
        JsonNode child = node.get(index);
        String childPath = path + "[" + index + "]";
        if (child.isNull()) {
          throw new InvalidRequestShapeException(
              GridGrindRequestProblemSupport.explicitNullFieldMessage(childPath),
              Optional.of(childPath),
              Optional.empty(),
              Optional.empty(),
              null);
        }
        rejectExplicitNullMembers(child, childPath);
      }
    }
  }

  private static void requireNonNullRoot(JsonNode node) {
    if (node.isNull()) {
      throw new InvalidRequestShapeException(
          "JSON payload must not be null",
          Optional.empty(),
          Optional.empty(),
          Optional.empty(),
          null);
    }
  }

  private static JsonNode requirePresentDocument(JsonNode node) {
    return Optional.ofNullable(node)
        .filter(candidate -> !candidate.isMissingNode())
        .orElseThrow(GridGrindJsonCodecSupport::invalidJsonPayloadException);
  }

  private static InvalidJsonException invalidJsonPayloadException() {
    return new InvalidJsonException(
        "Invalid JSON payload", Optional.empty(), Optional.empty(), Optional.empty(), null);
  }
}
