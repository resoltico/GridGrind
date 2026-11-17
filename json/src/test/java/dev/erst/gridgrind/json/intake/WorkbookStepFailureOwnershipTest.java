package dev.erst.gridgrind.json.intake;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.problem.RequestProblemSource;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.json.codec.GridGrindJsonMapperSupport;
import java.util.List;
import org.junit.jupiter.api.Test;
import tools.jackson.databind.exc.ValueInstantiationException;

/** Field grammar preserves actual producer failures for the complete document boundary. */
class WorkbookStepFailureOwnershipTest {
  private static final AssertionError FATAL = new AssertionError("synthetic constructor fatal");

  @Test
  void unknownRecordConstructorsRemainUnclassifiedWithTheirActualFieldPath() {
    var mapper = GridGrindJsonMapperSupport.JSON_MAPPER;
    for (String kind : List.of("argument", "null", "state")) {
      var node = mapper.createObjectNode().put("kind", kind);
      try (var parser = mapper.createParser("{}")) {
        var failure =
            assertThrows(
                ValueInstantiationException.class,
                () ->
                    WorkbookStepJsonDeserializer.deserializeField(
                        node, parser, ThrowingRecord.class, "action"));
        assertEquals("action", failure.getPath().getFirst().getPropertyName());
        assertFalse(failure.getCause() instanceof RequestProblemSource);
        assertEquals(
            switch (kind) {
              case "argument" -> IllegalArgumentException.class;
              case "null" -> NullPointerException.class;
              case "state" -> IllegalStateException.class;
              default -> throw new AssertionError("Unexpected fixture");
            },
            failure.getCause().getClass());
      }
    }
  }

  @Test
  void ownedInvariantRemainsTypedWhileFatalErrorsKeepIdentity() {
    var mapper = GridGrindJsonMapperSupport.JSON_MAPPER;
    try (var parser = mapper.createParser("{}")) {
      var known =
          assertThrows(
              ValueInstantiationException.class,
              () ->
                  WorkbookStepJsonDeserializer.deserializeField(
                      mapper.stringNode(""), parser, SheetName.class, "target"));
      assertInstanceOf(RequestProblemSource.class, known.getCause());
      assertEquals("target", known.getPath().getFirst().getPropertyName());
      assertSame(
          FATAL,
          assertThrows(
              AssertionError.class,
              () ->
                  WorkbookStepJsonDeserializer.deserializeField(
                      mapper.createObjectNode().put("kind", "fatal"),
                      parser,
                      ThrowingRecord.class,
                      "action")));
    }
  }

  @Test
  void canonicalConstructorValidationAlsoPreservesFatalIdentity() {
    var parsed =
        TolerantRequestJsonParser.parse(
            "{\"kind\":\"fatal\"}".getBytes(java.nio.charset.StandardCharsets.UTF_8));
    assertTrue(parsed.problems().isEmpty());
    var problems = new java.util.ArrayList<dev.erst.gridgrind.core.problem.RequestBindingFailure>();
    assertSame(
        FATAL,
        assertThrows(
            AssertionError.class,
            () ->
                RequestConstructorValidation.validate(
                    parsed.root(), ThrowingRecord.class, "action", problems)));
    assertTrue(problems.isEmpty());
  }

  /** Unowned record proves constructor failures are not inferred to be malformed input. */
  @SuppressWarnings("PMD.PublicMemberInNonPublicType")
  // Canonical model reflection needs a public record constructor, not a public JUnit owner.
  public record ThrowingRecord(String kind) {
    @SuppressWarnings("PMD.AvoidThrowingNullPointerException")
    // This deliberately unowned constructor supplies the required unexpected-NPE negative control.
    public ThrowingRecord {
      switch (kind) {
        case "argument" -> throw new IllegalArgumentException("synthetic internal argument");
        case "null" -> throw new NullPointerException("synthetic internal null");
        case "state" -> throw new IllegalStateException("synthetic internal state");
        case "fatal" -> throw FATAL;
        default -> throw new AssertionError("Unknown constructor fixture");
      }
    }
  }
}
