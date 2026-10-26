package dev.erst.gridgrind.json.codec;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import dev.erst.gridgrind.core.problem.InvalidJsonException;
import dev.erst.gridgrind.core.problem.InvalidRequestShapeException;
import dev.erst.gridgrind.core.problem.UnknownField;
import java.util.List;
import org.junit.jupiter.api.Test;
import tools.jackson.core.exc.StreamReadException;
import tools.jackson.databind.exc.MismatchedInputException;
import tools.jackson.databind.exc.UnrecognizedPropertyException;

/** Proves public classification is independent of framework exception prose. */
class JsonReadFailuresTest {
  @Test
  void syntaxMessagesNeverEchoFrameworkImplementationDetails() {
    var failure = new StreamReadException(null, "private framework package and internals");
    var publicFailure =
        assertInstanceOf(InvalidJsonException.class, JsonReadFailures.syntax(failure));
    assertEquals("Invalid JSON payload", publicFailure.getMessage());
    assertFalse(publicFailure.getMessage().contains("framework"));
  }

  @Test
  void unknownFieldsComeFromStructuredPropertyData() throws java.io.IOException {
    try (var parser = GridGrindJsonMapperSupport.JSON_MAPPER.createParser("{}")) {
      var failure =
          UnrecognizedPropertyException.from(parser, Object.class, "unexpected", List.of());
      var root = GridGrindJsonMapperSupport.JSON_MAPPER.readTree("{\"unexpected\":0}");
      var publicFailure =
          assertInstanceOf(
              InvalidRequestShapeException.class,
              JsonReadFailures.value(failure, root, Object.class));
      assertEquals(
          "unexpected",
          assertInstanceOf(UnknownField.class, publicFailure.requestProblem()).jsonPathValue());
      assertEquals("Unknown field 'unexpected'", publicFailure.getMessage());
    }
  }

  @Test
  void mismatchedValuesDoNotDependOnFrameworkWording() {
    var first =
        MismatchedInputException.from(null, Object.class, "one arbitrary implementation message");
    var second =
        MismatchedInputException.from(null, Object.class, "different implementation message");
    var root = GridGrindJsonMapperSupport.JSON_MAPPER.readTree("0");
    assertEquals(
        JsonReadFailures.value(first, root, Object.class).getMessage(),
        JsonReadFailures.value(second, root, Object.class).getMessage());
  }
}
