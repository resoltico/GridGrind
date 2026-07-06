package dev.erst.gridgrind.contract.step;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.json.InvalidRequestException;
import dev.erst.gridgrind.contract.json.InvalidRequestShapeException;
import dev.erst.gridgrind.contract.json.MessageInvariant;
import dev.erst.gridgrind.contract.json.MessageShape;
import dev.erst.gridgrind.contract.json.MissingTypeDiscriminator;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.time.DateTimeException;
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

    InvalidRequestException nullMessageFailure =
        assertInstanceOf(
            InvalidRequestException.class,
            WorkbookStepJsonFailurePathSupport.wrapIllegalArgumentFailure("target", nullMessage));
    InvalidRequestException fieldMessageFailure =
        assertInstanceOf(
            InvalidRequestException.class,
            WorkbookStepJsonFailurePathSupport.wrapIllegalArgumentFailure("target", fieldMessage));
    JacksonException inferredJacksonFailure =
        WorkbookStepJsonFailurePathSupport.wrapJacksonFailure(
            "target", emptyObject, NamedRecord.class, fieldlessJackson);
    JacksonException nestedJacksonFailure =
        WorkbookStepJsonFailurePathSupport.wrapJacksonFailure(
            "target", emptyObject, Object.class, nestedJackson);

    assertEquals("Invalid request", nullMessageFailure.getMessage());
    assertSame(nullMessage, nullMessageFailure.getCause());
    assertEquals(Optional.of("target"), nullMessageFailure.jsonPath());

    assertEquals("Field 'type' must exist", fieldMessageFailure.getMessage());
    assertEquals(Optional.of("target"), fieldMessageFailure.jsonPath());

    assertEquals("target.name", renderPath(inferredJacksonFailure));
    assertEquals("target.name", renderPath(nestedJacksonFailure));
  }

  @Test
  void wrapValidationJacksonFailureHandlesShapeAndDateTimeValidationCauses() throws IOException {
    JsonMapper mapper = JsonMapper.builder().build();
    JsonNode emptyObject = mapper.createObjectNode();
    MismatchedInputException shapeFailure =
        WorkbookStepJsonFailurePathSupport.inputMismatch(
            parser("{}"), new MissingTypeDiscriminator("type"));
    InvalidRequestException invalidRequestCause =
        new InvalidRequestException(
            new MessageInvariant("owned failure", Optional.of("name")),
            Optional.of("name"),
            Optional.empty(),
            Optional.empty(),
            null);
    MismatchedInputException invalidRequestFailure =
        MismatchedInputException.from(parser("{}"), WorkbookStep.class, "bad");
    invalidRequestFailure.initCause(new RuntimeException(invalidRequestCause));

    MismatchedInputException illegalArgumentFailure =
        MismatchedInputException.from(parser("{}"), WorkbookStep.class, "bad");
    illegalArgumentFailure.initCause(
        new RuntimeException(new IllegalArgumentException("unsupported literal")));

    MismatchedInputException dateTimeFailure =
        MismatchedInputException.from(parser("{}"), WorkbookStep.class, "bad");
    dateTimeFailure.initCause(new RuntimeException(new DateTimeException("unsupported date")));

    assertTrue(
        WorkbookStepJsonFailurePathSupport.wrapValidationJacksonFailure(
                "action", emptyObject, NamedRecord.class, shapeFailure)
            .isEmpty());

    InvalidRequestException wrappedInvalidRequestFailure =
        WorkbookStepJsonFailurePathSupport.wrapValidationJacksonFailure(
                "action", emptyObject, NamedRecord.class, invalidRequestFailure)
            .orElseThrow();
    assertEquals("owned failure", wrappedInvalidRequestFailure.getMessage());
    assertEquals(Optional.of("action.name"), wrappedInvalidRequestFailure.jsonPath());
    assertSame(invalidRequestFailure, wrappedInvalidRequestFailure.getCause());

    InvalidRequestException wrappedIllegalArgumentFailure =
        WorkbookStepJsonFailurePathSupport.wrapValidationJacksonFailure(
                "action", emptyObject, NamedRecord.class, illegalArgumentFailure)
            .orElseThrow();
    assertEquals("unsupported literal", wrappedIllegalArgumentFailure.getMessage());
    assertEquals(Optional.of("action"), wrappedIllegalArgumentFailure.jsonPath());
    assertSame(illegalArgumentFailure, wrappedIllegalArgumentFailure.getCause());

    InvalidRequestException wrappedDateTimeFailure =
        WorkbookStepJsonFailurePathSupport.wrapValidationJacksonFailure(
                "action", emptyObject, NamedRecord.class, dateTimeFailure)
            .orElseThrow();
    assertEquals("unsupported date", wrappedDateTimeFailure.getMessage());
    assertEquals(Optional.of("action"), wrappedDateTimeFailure.jsonPath());
    assertSame(dateTimeFailure, wrappedDateTimeFailure.getCause());
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
      InvalidRequestException failure =
          assertThrows(
              InvalidRequestException.class,
              () ->
                  WorkbookStepJsonDeserializer.deserializeField(
                      node, parser, AlwaysInvalidPayload.class, "action"));

      assertEquals("zoomPercent must be between 10 and 400", failure.getMessage());
      assertEquals(Optional.of("action"), failure.jsonPath());
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
