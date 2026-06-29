package dev.erst.gridgrind.contract.step;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import org.junit.jupiter.api.Test;
import tools.jackson.core.JacksonException;
import tools.jackson.core.JsonParser;
import tools.jackson.core.ObjectReadContext;
import tools.jackson.core.json.JsonFactory;
import tools.jackson.databind.DeserializationContext;
import tools.jackson.databind.ValueDeserializer;
import tools.jackson.databind.annotation.JsonDeserialize;
import tools.jackson.databind.exc.MismatchedInputException;
import tools.jackson.databind.json.JsonMapper;

/** Covers branch-heavy step deserialization helpers that own precise nested path derivation. */
class WorkbookStepJsonSupportCoverageTest {
  @Test
  void qualifiedFieldNameHandlesExistingPathsIndexesAndMessageInference() throws IOException {
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
    MismatchedInputException inferredFromJacksonMessage =
        MismatchedInputException.from(
            parser("{}"), WorkbookStep.class, "Missing required creator property 'name'");

    assertEquals(
        "target.name",
        WorkbookStepJsonFailurePathSupport.qualifiedFieldName("target", alreadyQualified));
    assertEquals(
        "target[0]", WorkbookStepJsonFailurePathSupport.qualifiedFieldName("target", indexedPath));
    assertEquals(
        "target", WorkbookStepJsonFailurePathSupport.qualifiedFieldName("target", fieldOnlyPath));
    assertEquals(
        "target[0]",
        WorkbookStepJsonFailurePathSupport.qualifiedFieldName("target", qualifiedIndexedPath));
    assertEquals(
        "action.zoomPercent",
        WorkbookStepJsonFailurePathSupport.qualifiedFieldName(
            "action", new IllegalArgumentException("zoomPercent must be between 10 and 400")));
    assertEquals(
        "assertion.type",
        WorkbookStepJsonFailurePathSupport.qualifiedFieldName(
            "assertion", new IllegalArgumentException("Field 'type' must be a string")));
    assertEquals(
        "assertion.type",
        WorkbookStepJsonFailurePathSupport.qualifiedFieldName(
            "assertion", new IllegalArgumentException("missing required creator property 'type'")));
    assertEquals(
        "assertion.type",
        WorkbookStepJsonFailurePathSupport.qualifiedFieldName(
            "assertion", new IllegalArgumentException("missing type id property 'type'")));
    assertEquals(
        "target.name",
        WorkbookStepJsonFailurePathSupport.qualifiedFieldName(
            "target", inferredFromJacksonMessage));
    assertEquals(
        "query",
        WorkbookStepJsonFailurePathSupport.qualifiedFieldName(
            "query", new IllegalArgumentException(" ")));
    assertEquals(
        "query",
        WorkbookStepJsonFailurePathSupport.qualifiedFieldName(
            "query", new IllegalArgumentException("not a field-specific failure")));
  }

  @Test
  void wrapHelpersPreserveCauseMessageFallbackAndQualifiedPaths() throws IOException {
    IllegalArgumentException nullMessage = new IllegalArgumentException();
    IllegalArgumentException fieldMessage = new IllegalArgumentException("Field 'type' must exist");
    MismatchedInputException fieldlessJackson =
        MismatchedInputException.from(
            parser("{}"), WorkbookStep.class, "Missing required creator property 'name'");
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
        WorkbookStepJsonFailurePathSupport.wrapJacksonFailure("target", fieldlessJackson);
    JacksonException nestedJacksonFailure =
        WorkbookStepJsonFailurePathSupport.wrapJacksonFailure("target", nestedJackson);

    assertEquals("Invalid request shape", nullMessageFailure.getOriginalMessage());
    assertSame(nullMessage, nullMessageFailure.getCause());
    assertEquals("target", renderPath(nullMessageFailure));

    assertEquals("Field 'type' must exist", fieldMessageFailure.getOriginalMessage());
    assertEquals("target.type", renderPath(fieldMessageFailure));

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
      assertEquals("action.zoomPercent", renderPath(failure));
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
