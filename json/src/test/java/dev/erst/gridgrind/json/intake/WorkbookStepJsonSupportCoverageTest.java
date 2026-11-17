package dev.erst.gridgrind.json.intake;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.plan.WorkbookStep;
import dev.erst.gridgrind.core.problem.InvalidRequestShapeException;
import dev.erst.gridgrind.core.problem.MessageShape;
import dev.erst.gridgrind.core.problem.MissingTypeDiscriminator;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.Optional;
import org.junit.jupiter.api.Test;
import tools.jackson.core.JacksonException;
import tools.jackson.core.JsonParser;
import tools.jackson.core.ObjectReadContext;
import tools.jackson.core.json.JsonFactory;
import tools.jackson.databind.DeserializationContext;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.ValueDeserializer;
import tools.jackson.databind.annotation.JsonDeserialize;
import tools.jackson.databind.exc.MismatchedInputException;
import tools.jackson.databind.json.JsonMapper;

/** Covers branch-heavy step deserialization helpers that own precise nested path derivation. */
class WorkbookStepJsonSupportCoverageTest {
  @Test
  void qualifiedFieldNameHandlesExistingPathsAndStructuredProblemSources() throws IOException {
    JsonMapper mapper = dev.erst.gridgrind.json.codec.GridGrindJsonMapperSupport.JSON_MAPPER;
    JsonNode emptyObject = mapper.createObjectNode();
    MismatchedInputException alreadyQualified =
        (MismatchedInputException)
            MismatchedInputException.from(parser("{}"), WorkbookStep.class, "bad")
                .prependPath(new Object(), "name")
                .prependPath(new Object(), "target");
    MismatchedInputException indexedPath =
        (MismatchedInputException)
            MismatchedInputException.from(parser("{}"), WorkbookStep.class, "bad")
                .prependPath(new Object(), 0);
    MismatchedInputException fieldOnlyPath =
        (MismatchedInputException)
            MismatchedInputException.from(parser("{}"), WorkbookStep.class, "bad")
                .prependPath(new Object(), "target");
    MismatchedInputException qualifiedIndexedPath =
        (MismatchedInputException)
            MismatchedInputException.from(parser("{}"), WorkbookStep.class, "bad")
                .prependPath(new Object(), 0)
                .prependPath(new Object(), "target");
    MismatchedInputException inferredStructurally =
        MismatchedInputException.from(parser("{}"), NamedRecord.class, "bad");
    InvalidRequestShapeException typedProblem =
        new InvalidRequestShapeException(
            new MissingTypeDiscriminator("type"),
            java.util.Optional.of("type"),
            java.util.Optional.empty(),
            java.util.Optional.empty(),
            null);

    assertEquals(
        "target.name",
        WorkbookStepJsonFailurePathSupport.qualifiedFieldName(
            "target", null, Object.class, alreadyQualified));
    assertEquals(
        "target[0]",
        WorkbookStepJsonFailurePathSupport.qualifiedFieldName(
            "target", null, Object.class, indexedPath));
    assertEquals(
        "target",
        WorkbookStepJsonFailurePathSupport.qualifiedFieldName(
            "target", null, Object.class, fieldOnlyPath));
    assertEquals(
        "target[0]",
        WorkbookStepJsonFailurePathSupport.qualifiedFieldName(
            "target", null, Object.class, qualifiedIndexedPath));
    assertEquals(
        "assertion.type",
        WorkbookStepJsonFailurePathSupport.qualifiedFieldName(
            "assertion", null, Object.class, typedProblem));
    assertEquals(
        "target.name",
        WorkbookStepJsonFailurePathSupport.qualifiedFieldName(
            "target", emptyObject, NamedRecord.class, inferredStructurally));
    assertEquals(
        "query",
        WorkbookStepJsonFailurePathSupport.qualifiedFieldName(
            "query", null, Object.class, new IllegalArgumentException(" ")));
    assertEquals(
        "query",
        WorkbookStepJsonFailurePathSupport.qualifiedFieldName(
            "query",
            null,
            Object.class,
            new IllegalArgumentException("not a field-specific failure")));
    assertEquals(
        "query",
        WorkbookStepJsonFailurePathSupport.qualifiedFieldName(
            "query", null, NamedRecord.class, inferredStructurally));
    assertEquals(
        "query",
        WorkbookStepJsonFailurePathSupport.qualifiedFieldName(
            "query",
            null,
            Object.class,
            new InvalidRequestShapeException(
                new MessageShape("bad", Optional.empty()),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                null)));
  }

  @Test
  void fieldGrammarPreservesJacksonIdentityAndQualifiedPaths() throws IOException {
    JsonMapper mapper = dev.erst.gridgrind.json.codec.GridGrindJsonMapperSupport.JSON_MAPPER;
    JsonNode emptyObject = mapper.createObjectNode();
    try (JsonParser missingParser = parser("{}");
        JsonParser nestedParser = parser("{}")) {
      MismatchedInputException missing =
          MismatchedInputException.from(missingParser, NamedRecord.class, "bad");
      MismatchedInputException nested =
          MismatchedInputException.from(nestedParser, WorkbookStep.class, "bad");
      nested.prependPath(new Object(), "name");
      assertSame(
          missing,
          WorkbookStepJsonFailurePathSupport.wrapJacksonFailure(
              "target", emptyObject, NamedRecord.class, missing));
      assertSame(
          nested,
          WorkbookStepJsonFailurePathSupport.wrapJacksonFailure(
              "target", emptyObject, Object.class, nested));
      assertEquals("target.name", renderPath(missing));
      assertEquals("target.name", renderPath(nested));
    }
  }

  @Test
  void directWorkbookStepDeserializationStillRejectsMissingPayloadWithoutTopLevelPrecheck() {
    MismatchedInputException failure =
        assertThrows(
            MismatchedInputException.class,
            () ->
                dev.erst.gridgrind.json.codec.GridGrindJsonMapperSupport.JSON_MAPPER.readValue(
                    "{\"stepId\":\"only-target\",\"target\":{}}", WorkbookStep.class));

    assertEquals(
        "Each step must contain exactly one of 'action', 'assertion', or 'query'",
        failure.getOriginalMessage());
  }

  @Test
  void deserializeFieldPreservesAnUnclassifiedCustomDeserializerFailure() throws IOException {
    JsonMapper mapper = dev.erst.gridgrind.json.codec.GridGrindJsonMapperSupport.JSON_MAPPER;
    var node = mapper.readTree("{\"anything\":true}");
    try (JsonParser parser = mapper.createParser("{}")) {
      IllegalArgumentException failure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  WorkbookStepJsonDeserializer.deserializeField(
                      node, parser, AlwaysInvalidPayload.class, "action"));
      assertEquals("zoomPercent must be between 10 and 400", failure.getMessage());
      assertFalse(failure instanceof dev.erst.gridgrind.core.problem.RequestProblemSource);
    }
  }

  private static JsonParser parser(String json) throws IOException {
    return new JsonFactory()
        .createParser(
            ObjectReadContext.empty(),
            new ByteArrayInputStream(json.getBytes(StandardCharsets.UTF_8)));
  }

  private static String renderPath(JacksonException exception) {
    StringBuilder rendered = new StringBuilder();
    for (JacksonException.Reference reference : exception.getPath()) {
      if (reference.getPropertyName() != null) {
        if (!rendered.isEmpty()) {
          rendered.append('.');
        }
        rendered.append(reference.getPropertyName());
      } else {
        rendered.append('[').append(reference.getIndex()).append(']');
      }
    }
    return rendered.toString();
  }

  private record NamedRecord(String name) {}

  /** Synthetic payload type whose custom deserializer throws a bare IllegalArgumentException. */
  @JsonDeserialize(using = AlwaysInvalidPayloadDeserializer.class)
  private static final class AlwaysInvalidPayload {}

  /** Minimal unowned deserializer whose raw failure must not acquire an input-refusal meaning. */
  private static final class AlwaysInvalidPayloadDeserializer
      extends ValueDeserializer<AlwaysInvalidPayload> {
    @Override
    public AlwaysInvalidPayload deserialize(JsonParser parser, DeserializationContext context) {
      throw new IllegalArgumentException("zoomPercent must be between 10 and 400");
    }
  }
}
