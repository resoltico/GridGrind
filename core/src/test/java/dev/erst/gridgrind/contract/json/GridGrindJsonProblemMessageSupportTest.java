package dev.erst.gridgrind.contract.json;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.InvalidRawFormulaTextException;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.time.DateTimeException;
import java.util.Optional;
import org.junit.jupiter.api.Test;
import tools.jackson.core.JsonParser;
import tools.jackson.core.ObjectReadContext;
import tools.jackson.core.exc.StreamConstraintsException;
import tools.jackson.core.json.JsonFactory;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.exc.InvalidFormatException;
import tools.jackson.databind.exc.MismatchedInputException;
import tools.jackson.databind.json.JsonMapper;

/** Covers residual Jackson message-translation branches owned by the JSON surface. */
class GridGrindJsonProblemMessageSupportTest {
  @Test
  void enumValueMessagesFallBackWhenFieldPathIsMissingOrBlank() throws IOException {
    InvalidFormatException fieldlessEnum =
        InvalidFormatException.from(parser("\"OMEGA\""), "bad enum", "OMEGA", SampleEnum.class);
    InvalidFormatException blankFieldEnum =
        (InvalidFormatException)
            InvalidFormatException.from(parser("\"OMEGA\""), "bad enum", "OMEGA", SampleEnum.class)
                .prependPath(WorkbookPlan.class, "");
    InvalidFormatException namedFieldEnum =
        (InvalidFormatException)
            InvalidFormatException.from(parser("\"OMEGA\""), "bad enum", "OMEGA", SampleEnum.class)
                .prependPath(WorkbookPlan.class, "mode");
    InvalidFormatException nullValueEnum =
        InvalidFormatException.from(parser("null"), "bad enum", null, SampleEnum.class);

    assertEquals(
        "Unsupported value 'OMEGA'; expected one of: ALPHA, BETA",
        GridGrindJsonValueProblemSupport.enumValueMessage(fieldlessEnum));
    assertEquals(
        "Unsupported value 'OMEGA'; expected one of: ALPHA, BETA",
        GridGrindJsonValueProblemSupport.enumValueMessage(blankFieldEnum));
    assertEquals(
        "Unsupported value 'OMEGA' for field 'mode'; expected one of: ALPHA, BETA",
        GridGrindJsonValueProblemSupport.enumValueMessage(namedFieldEnum));
    assertEquals(
        "Unsupported value 'null'; expected one of: ALPHA, BETA",
        GridGrindJsonValueProblemSupport.enumValueMessage(nullValueEnum));
    assertFalse(GridGrindJsonValueProblemSupport.hasNonBlankFieldName(null));
    assertFalse(GridGrindJsonValueProblemSupport.hasNonBlankFieldName(""));
    assertTrue(GridGrindJsonValueProblemSupport.hasNonBlankFieldName("mode"));
  }

  @Test
  void nonEnumInvalidFormatsFlowThroughTheGenericWrongShapeMessage() throws IOException {
    InvalidFormatException nonEnumValue =
        (InvalidFormatException)
            InvalidFormatException.from(
                    parser("\"abc\""),
                    "Cannot deserialize value of type `int` from String \"abc\"",
                    "abc",
                    Integer.class)
                .prependPath(WorkbookPlan.class, "stepCount");

    assertEquals(
        "JSON value has the wrong shape for this field",
        GridGrindJsonProblemMessageSupport.message(nonEnumValue));

    InvalidFormatException nullTargetType =
        new InvalidFormatException(
            parser("\"abc\""),
            "Cannot deserialize value of type `int` from String \"abc\"",
            "abc",
            null);
    assertEquals(
        "JSON value has the wrong shape for this field",
        GridGrindJsonProblemMessageSupport.message(nullTargetType));
  }

  @Test
  void mismatchedInputMessagesCoverBlankAndStartObjectFallbacks() throws IOException {
    try (JsonParser objectParser = parser("{}")) {
      objectParser.nextToken();
      MismatchedInputException objectShape =
          MismatchedInputException.from(
              objectParser, WorkbookPlan.class, "Cannot construct instance");
      assertEquals(
          "JSON object is missing required fields or has the wrong shape",
          GridGrindJsonValueProblemSupport.mismatchedInputMessage(objectShape));
    }

    MismatchedInputException blankMessage =
        MismatchedInputException.from(parser("\"abc\""), String.class, " ");
    assertEquals(
        "Invalid JSON payload",
        GridGrindJsonValueProblemSupport.mismatchedInputMessage(blankMessage));
  }

  @Test
  void cleanJacksonMessageLeavesCreatorNullAndMissingRequiredTextIntact() {
    assertEquals(
        "protocolVersion must not be null",
        GridGrindJsonProblemMessageSupport.cleanJacksonMessage("protocolVersion must not be null"));
    assertEquals(
        "Missing required creator property 'protocolVersion'",
        GridGrindJsonProblemMessageSupport.cleanJacksonMessage(
            "Missing required creator property 'protocolVersion'"));
  }

  @Test
  void messageAndInvalidPayloadOwnTypedProblemSourcesAndJacksonFallbacks() throws IOException {
    JsonNode rootNode = JsonMapper.builder().build().createObjectNode();
    InvalidRequestShapeException typedShape =
        new InvalidRequestShapeException(
            new MissingRequiredField("protocolVersion"),
            Optional.of("protocolVersion"),
            Optional.of(1),
            Optional.of(2),
            null);
    InvalidFormatException enumProblem =
        (InvalidFormatException)
            InvalidFormatException.from(parser("\"OMEGA\""), "bad enum", "OMEGA", SampleEnum.class)
                .prependPath(WorkbookPlan.class, "mode");
    MismatchedInputException databindFailure =
        MismatchedInputException.from(parser("{}"), WorkbookPlan.class, "Cannot deserialize value");

    assertEquals(
        "Missing required field 'protocolVersion'",
        GridGrindJsonProblemMessageSupport.message(typedShape));
    assertEquals(
        "Unsupported value 'OMEGA' for field 'mode'; expected one of: ALPHA, BETA",
        GridGrindJsonProblemMessageSupport.message(enumProblem));
    assertEquals(
        "too deep",
        GridGrindJsonProblemMessageSupport.invalidPayload(
                new StreamConstraintsException("too deep"), rootNode, WorkbookPlan.class)
            .getMessage());
    assertInstanceOf(
        InvalidRequestShapeException.class,
        GridGrindJsonProblemMessageSupport.invalidPayload(databindFailure));
  }

  @Test
  void invalidPayloadRebasesTypedValidationCausesIntoPublicProblemFamilies() throws IOException {
    JsonNode rootNode = JsonMapper.builder().build().createObjectNode();
    InvalidRequestShapeException shapeCause =
        new InvalidRequestShapeException(
            new MissingRequiredField("protocolVersion"),
            Optional.of("protocolVersion"),
            Optional.of(7),
            Optional.of(9),
            null);
    MismatchedInputException shapeCarrier =
        MismatchedInputException.from(parser("{}"), (Class<?>) null, "Cannot deserialize value");
    shapeCarrier.initCause(shapeCause);

    InvalidRequestShapeException rebasedShape =
        assertInstanceOf(
            InvalidRequestShapeException.class,
            GridGrindJsonProblemMessageSupport.invalidPayload(
                shapeCarrier, rootNode, Object.class));
    assertEquals("Missing required field 'protocolVersion'", rebasedShape.getMessage());
    assertEquals(Optional.of("protocolVersion"), rebasedShape.jsonPath());

    InvalidRequestException invariantCause =
        new InvalidRequestException(
            new MessageInvariant("invalid plan identifier", Optional.of("planId")),
            Optional.of("planId"),
            Optional.of(7),
            Optional.of(9),
            null);
    MismatchedInputException invariantProblemCarrier =
        MismatchedInputException.from(parser("{}"), (Class<?>) null, "Cannot deserialize value");
    invariantProblemCarrier.initCause(invariantCause);

    InvalidRequestException rebasedInvariant =
        assertInstanceOf(
            InvalidRequestException.class,
            GridGrindJsonProblemMessageSupport.invalidPayload(
                invariantProblemCarrier, rootNode, Object.class));
    assertEquals("invalid plan identifier", rebasedInvariant.getMessage());
    assertEquals(Optional.of("planId"), rebasedInvariant.jsonPath());

    InvalidRequestShapeException nestedShapeCause =
        new InvalidRequestShapeException(
            new MissingTypeDiscriminator("type"),
            Optional.empty(),
            Optional.of(7),
            Optional.of(9),
            null);
    MismatchedInputException nestedShapeCarrier =
        (MismatchedInputException)
            MismatchedInputException.from(parser("{}"), (Class<?>) null, "Cannot deserialize value")
                .prependPath(WorkbookPlan.class, "target")
                .prependPath(WorkbookPlan.class, 0)
                .prependPath(WorkbookPlan.class, "steps");
    nestedShapeCarrier.initCause(nestedShapeCause);

    InvalidRequestShapeException rebasedNestedShape =
        assertInstanceOf(
            InvalidRequestShapeException.class,
            GridGrindJsonProblemMessageSupport.invalidPayload(
                nestedShapeCarrier, rootNode, Object.class));
    assertEquals("Missing required field 'steps[0].target.type'", rebasedNestedShape.getMessage());
    assertEquals(Optional.of("steps[0].target.type"), rebasedNestedShape.jsonPath());

    MismatchedInputException invariantCarrier =
        MismatchedInputException.from(parser("\"2026-99-99\""), (Class<?>) null, "bad date");
    invariantCarrier.initCause(new DateTimeException("bad date"));

    InvalidRequestException invariant =
        assertInstanceOf(
            InvalidRequestException.class,
            GridGrindJsonProblemMessageSupport.invalidPayload(
                invariantCarrier, rootNode, Object.class));
    assertEquals("bad date", invariant.getMessage());
    assertEquals(Optional.empty(), invariant.jsonPath());

    MismatchedInputException rawFormulaCarrier =
        MismatchedInputException.from(parser("\"bad\""), (Class<?>) null, "bad formula text");
    rawFormulaCarrier.initCause(
        new InvalidRawFormulaTextException("formula text contains a forbidden XML character"));
    FormulaRequestException rawFormula =
        assertInstanceOf(
            FormulaRequestException.class,
            GridGrindJsonProblemMessageSupport.invalidPayload(
                rawFormulaCarrier, rootNode, Object.class));
    assertEquals(GridGrindProblemCode.INVALID_FORMULA_TEXT, rawFormula.problemCode());
  }

  @Test
  void invalidPayloadMergesIndexedProblemPathsWithoutReflection() throws IOException {
    JsonNode rootNode = JsonMapper.builder().build().createObjectNode();
    InvalidRequestShapeException indexedProblem =
        new InvalidRequestShapeException(
            new MessageShape("bad", Optional.of("[0]")),
            Optional.of("[0]"),
            Optional.empty(),
            Optional.empty(),
            null);
    MismatchedInputException matchingIndexedCarrier =
        (MismatchedInputException)
            MismatchedInputException.from(parser("{}"), (Class<?>) null, "bad")
                .prependPath(new Object(), 0)
                .prependPath(new Object(), "steps");
    matchingIndexedCarrier.initCause(indexedProblem);
    MismatchedInputException mismatchedIndexedCarrier =
        (MismatchedInputException)
            MismatchedInputException.from(parser("{}"), (Class<?>) null, "bad")
                .prependPath(new Object(), 1)
                .prependPath(new Object(), "steps");
    mismatchedIndexedCarrier.initCause(indexedProblem);
    InvalidRequestShapeException exactPathProblem =
        new InvalidRequestShapeException(
            new MessageShape("bad", Optional.of("steps[0].target.type")),
            Optional.of("steps[0].target.type"),
            Optional.empty(),
            Optional.empty(),
            null);
    MismatchedInputException exactPathCarrier =
        (MismatchedInputException)
            MismatchedInputException.from(parser("{}"), (Class<?>) null, "bad")
                .prependPath(new Object(), "type")
                .prependPath(new Object(), "target")
                .prependPath(new Object(), 0)
                .prependPath(new Object(), "steps");
    exactPathCarrier.initCause(exactPathProblem);
    InvalidRequestShapeException deeperPathProblem =
        new InvalidRequestShapeException(
            new MessageShape("bad", Optional.of("steps[0].target.type")),
            Optional.of("steps[0].target.type"),
            Optional.empty(),
            Optional.empty(),
            null);
    MismatchedInputException deeperPathCarrier =
        (MismatchedInputException)
            MismatchedInputException.from(parser("{}"), (Class<?>) null, "bad")
                .prependPath(new Object(), "target")
                .prependPath(new Object(), 0)
                .prependPath(new Object(), "steps");
    deeperPathCarrier.initCause(deeperPathProblem);
    InvalidRequestShapeException indexedOwnedProblem =
        new InvalidRequestShapeException(
            new MessageShape("bad", Optional.of("steps[0]")),
            Optional.of("steps[0]"),
            Optional.empty(),
            Optional.empty(),
            null);
    MismatchedInputException indexedOwnedCarrier =
        (MismatchedInputException)
            MismatchedInputException.from(parser("{}"), (Class<?>) null, "bad")
                .prependPath(new Object(), "steps");
    indexedOwnedCarrier.initCause(indexedOwnedProblem);

    assertEquals(
        Optional.of("steps[0]"),
        assertInstanceOf(
                InvalidRequestShapeException.class,
                GridGrindJsonProblemMessageSupport.invalidPayload(
                    matchingIndexedCarrier, rootNode, Object.class))
            .jsonPath());
    assertEquals(
        Optional.of("steps[1][0]"),
        assertInstanceOf(
                InvalidRequestShapeException.class,
                GridGrindJsonProblemMessageSupport.invalidPayload(
                    mismatchedIndexedCarrier, rootNode, Object.class))
            .jsonPath());
    assertEquals(
        Optional.of("steps[0].target.type"),
        assertInstanceOf(
                InvalidRequestShapeException.class,
                GridGrindJsonProblemMessageSupport.invalidPayload(
                    exactPathCarrier, rootNode, Object.class))
            .jsonPath());
    assertEquals(
        Optional.of("steps[0].target.type"),
        assertInstanceOf(
                InvalidRequestShapeException.class,
                GridGrindJsonProblemMessageSupport.invalidPayload(
                    deeperPathCarrier, rootNode, Object.class))
            .jsonPath());
    assertEquals(
        Optional.of("steps[0]"),
        assertInstanceOf(
                InvalidRequestShapeException.class,
                GridGrindJsonProblemMessageSupport.invalidPayload(
                    indexedOwnedCarrier, rootNode, Object.class))
            .jsonPath());
  }

  private static JsonParser parser(String json) throws IOException {
    return new JsonFactory()
        .createParser(
            ObjectReadContext.empty(),
            new ByteArrayInputStream(json.getBytes(StandardCharsets.UTF_8)));
  }

  /** Synthetic enum used to exercise product-owned enum-value error wording. */
  private enum SampleEnum {
    ALPHA,
    BETA
  }
}
