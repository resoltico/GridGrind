package dev.erst.gridgrind.contract.step;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.json.InvalidRequestShapeException;
import dev.erst.gridgrind.contract.json.MessageShape;
import dev.erst.gridgrind.contract.json.MissingTypeDiscriminator;
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
    JsonMapper mapper = JsonMapper.builder().build();
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
  void wrapHelpersPreserveCauseMessageFallbackAndQualifiedPaths() throws IOException {
    JsonMapper mapper = JsonMapper.builder().build();
    JsonNode emptyObject = mapper.createObjectNode();
    IllegalArgumentException nullMessage = new IllegalArgumentException();
    IllegalArgumentException fieldMessage = new IllegalArgumentException("Field 'type' must exist");
    MismatchedInputException fieldlessJackson =
        MismatchedInputException.from(parser("{}"), NamedRecord.class, "bad");
    MismatchedInputException nestedJackson =
        (MismatchedInputException)
            MismatchedInputException.from(parser("{}"), WorkbookStep.class, "bad")
                .prependPath(new Object(), "name");

    MismatchedInputException nullMessageFailure =
        assertInstanceOf(
            MismatchedInputException.class,
            WorkbookStepJsonFailurePathSupport.wrapIllegalArgumentFailure(
                parser("{}"), "target", nullMessage));
    MismatchedInputException fieldMessageFailure =
        assertInstanceOf(
            MismatchedInputException.class,
            WorkbookStepJsonFailurePathSupport.wrapIllegalArgumentFailure(
                parser("{}"), "target", fieldMessage));
    JacksonException inferredJacksonFailure =
        WorkbookStepJsonFailurePathSupport.wrapJacksonFailure(
            "target", emptyObject, NamedRecord.class, fieldlessJackson);
    JacksonException nestedJacksonFailure =
        WorkbookStepJsonFailurePathSupport.wrapJacksonFailure(
            "target", emptyObject, Object.class, nestedJackson);

    assertEquals("Invalid request shape", nullMessageFailure.getOriginalMessage());
    assertSame(nullMessage, nullMessageFailure.getCause());
    assertEquals("target", renderPath(nullMessageFailure));

    assertEquals("Field 'type' must exist", fieldMessageFailure.getOriginalMessage());
    assertEquals("target", renderPath(fieldMessageFailure));

    assertEquals("target.name", renderPath(inferredJacksonFailure));
    assertEquals("target.name", renderPath(nestedJacksonFailure));
  }

  @Test
  void directWorkbookStepDeserializationStillRejectsMissingPayloadWithoutTopLevelPrecheck() {
    MismatchedInputException failure =
        assertThrows(
            MismatchedInputException.class,
            () ->
                JsonMapper.builder()
                    .build()
                    .readValue("{\"stepId\":\"only-target\",\"target\":{}}", WorkbookStep.class));

    assertEquals(
        "Each step must contain exactly one of 'action', 'assertion', or 'query'",
        failure.getOriginalMessage());
  }

  @Test
  void deserializeFieldWrapsBareIllegalArgumentFailuresAgainstTheOwningField() throws IOException {
    JsonMapper mapper = JsonMapper.builder().build();
    var node = mapper.readTree("{\"anything\":true}");
    try (JsonParser parser = mapper.createParser("{}")) {
      MismatchedInputException failure =
          assertThrows(
              MismatchedInputException.class,
              () ->
                  WorkbookStepJsonDeserializer.deserializeField(
                      node, parser, AlwaysInvalidPayload.class, "action"));

      assertEquals("zoomPercent must be between 10 and 400", failure.getOriginalMessage());
      assertEquals("action", renderPath(failure));
      assertEquals("zoomPercent must be between 10 and 400", failure.getCause().getMessage());
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

  /** Minimal deserializer used to exercise the explicit IllegalArgumentException wrapper path. */
  private static final class AlwaysInvalidPayloadDeserializer
      extends ValueDeserializer<AlwaysInvalidPayload> {
    @Override
    public AlwaysInvalidPayload deserialize(JsonParser parser, DeserializationContext context) {
      throw new IllegalArgumentException("zoomPercent must be between 10 and 400");
    }
  }
}
