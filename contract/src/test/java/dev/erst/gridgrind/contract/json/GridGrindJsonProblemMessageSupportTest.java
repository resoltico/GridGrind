package dev.erst.gridgrind.contract.json;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import org.junit.jupiter.api.Test;
import tools.jackson.core.JsonParser;
import tools.jackson.core.ObjectReadContext;
import tools.jackson.core.json.JsonFactory;
import tools.jackson.databind.exc.InvalidFormatException;

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
  void nullCreatorMessagesStillMapToMissingRequiredFields() {
    assertEquals(
        "Missing required field 'protocolVersion'",
        GridGrindJsonValueProblemSupport.productOwnedJacksonMessage(
            "protocolVersion must not be null"));
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
