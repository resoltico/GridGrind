package dev.erst.gridgrind.json.codec;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.problem.ExplicitNullField;
import dev.erst.gridgrind.core.problem.InvalidJsonException;
import dev.erst.gridgrind.core.problem.InvalidRequestShapeException;
import dev.erst.gridgrind.core.problem.MissingRequiredField;
import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.Optional;
import org.junit.jupiter.api.Test;
import tools.jackson.databind.json.JsonMapper;

/** Checks generic value-fragment codec mechanics; complete documents use their owned readers. */
class GridGrindJsonCodecSupportTest {
  private static final JsonMapper MAPPER = GridGrindJsonMapperSupport.JSON_MAPPER;

  @Test
  void decodeTreeReturnsAValidBoundFragment() throws IOException {
    var tree = MAPPER.readTree("{\"protocolVersion\":1}");
    assertEquals(
        new Fragment(ProtocolVersion.CURRENT),
        GridGrindJsonCodecSupport.decodeTree(
            tree, MAPPER, Fragment.class, JsonReadFailures::syntax));
  }

  @Test
  void decodeTreeRejectsExplicitNullMembersBeforeBinding() throws IOException {
    var tree = MAPPER.readTree("{\"protocolVersion\":null}");
    var failure =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJsonCodecSupport.decodeTree(
                    tree, MAPPER, Fragment.class, JsonReadFailures::syntax));
    var problem = assertInstanceOf(ExplicitNullField.class, failure.requestProblem());
    assertEquals("protocolVersion", problem.jsonPathValue());
    assertEquals(Optional.of("protocolVersion"), failure.jsonPath());
  }

  @Test
  void decodeTreeDetectsMissingRequiredFieldsStructurally() throws IOException {
    var tree = MAPPER.readTree("{}");
    var failure =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJsonCodecSupport.decodeTree(
                    tree, MAPPER, Fragment.class, JsonReadFailures::syntax));
    var problem = assertInstanceOf(MissingRequiredField.class, failure.requestProblem());
    assertEquals("protocolVersion", problem.jsonPathValue());
    assertEquals(Optional.of("protocolVersion"), failure.jsonPath());
  }

  @Test
  void codecSupportMapsStringGrammarFailuresAndNestedArrayNullsBeforeBinding() throws IOException {
    assertEquals(
        "{}",
        GridGrindJsonCodecSupport.readTree("{}", MAPPER, JsonReadFailures::syntax).toString());
    assertThrows(
        InvalidJsonException.class,
        () -> GridGrindJsonCodecSupport.readTree("{", MAPPER, JsonReadFailures::syntax));
    assertThrows(
        InvalidRequestShapeException.class,
        () ->
            GridGrindJsonCodecSupport.decodeTree(
                MAPPER.readTree("[null]"), MAPPER, Object.class, JsonReadFailures::syntax));
    assertThrows(
        InvalidRequestShapeException.class,
        () ->
            GridGrindJsonCodecSupport.decodeTree(
                MAPPER.nullNode(), MAPPER, Object.class, JsonReadFailures::syntax));
  }

  @Test
  void rejectsNonObjectTopLevelPayloadsForRecordFragments() {
    var failure =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJsonCodecSupport.readValue(
                    "[]".getBytes(StandardCharsets.UTF_8),
                    MAPPER,
                    Fragment.class,
                    JsonReadFailures::syntax));
    assertEquals("JSON value has the wrong shape for this field", failure.getMessage());
  }

  @SuppressWarnings("PMD.PublicMemberInNonPublicType")
  // The reflected value fixture needs its public canonical constructor; the JUnit class stays
  // package-private.
  public record Fragment(ProtocolVersion protocolVersion) {
    public Fragment {
      java.util.Objects.requireNonNull(protocolVersion, "protocolVersion must not be null");
    }
  }
}
