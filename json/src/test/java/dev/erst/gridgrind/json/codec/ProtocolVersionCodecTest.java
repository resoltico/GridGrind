package dev.erst.gridgrind.json.codec;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.catalog.FieldConstraint;
import dev.erst.gridgrind.core.catalog.FieldShape;
import dev.erst.gridgrind.core.catalog.GridGrindProtocolCatalog;
import dev.erst.gridgrind.core.catalog.ScalarType;
import dev.erst.gridgrind.core.problem.InvalidRequestException;
import dev.erst.gridgrind.core.problem.InvalidRequestShapeException;
import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;
import tools.jackson.core.JacksonException;

/** Checks the integer header independently of its Java record representation. */
class ProtocolVersionCodecTest {
  @Test
  void versionIsOneIntegerAndRejectsOtherValuesAndTokenKinds() {
    var mapper = GridGrindJsonMapperSupport.JSON_MAPPER;
    assertArrayEquals(new byte[] {'1'}, mapper.writeValueAsBytes(ProtocolVersion.CURRENT));
    assertEquals(new ProtocolVersion(1), mapper.readValue("1", ProtocolVersion.class));
    for (String invalid :
        List.of("0", "2", "-1", "2147483648", "1.0", "\"1\"", "\"V3\"", "true", "{}")) {
      assertThrows(
          JacksonException.class, () -> mapper.readValue(invalid, ProtocolVersion.class), invalid);
    }
  }

  @Test
  void requestIntakeAcceptsTheIntegerHeaderAndLocatesInvalidHeaders() throws java.io.IOException {
    var plan = GridGrindJson.readRequest(request("1"));
    assertEquals(new ProtocolVersion(1), plan.protocolVersion());
    InvalidRequestException unsupported =
        assertThrows(InvalidRequestException.class, () -> GridGrindJson.readRequest(request("2")));
    assertEquals(Optional.of("protocolVersion"), unsupported.jsonPath());
    for (String invalid : List.of("\"V3\"", "\"1\"", "1.0", "null")) {
      InvalidRequestShapeException failure =
          assertThrows(
              InvalidRequestShapeException.class,
              () -> GridGrindJson.readRequest(request(invalid)),
              invalid);
      assertEquals(Optional.of("protocolVersion"), failure.jsonPath(), invalid);
    }
  }

  @Test
  void catalogPublishesTheNumericConstraintRatherThanAnEnumName() {
    var field =
        GridGrindProtocolCatalog.catalog().requestType().field("protocolVersion").orElseThrow();
    assertEquals(
        new FieldShape.Scalar(ScalarType.NUMBER, List.of(new FieldConstraint.NumberRange(1, 1))),
        field.shape());
    assertEquals(List.of(), field.enumValues());
  }

  private static byte[] request(String version) {
    return ("{\"protocolVersion\":"
            + version
            + ",\"source\":{\"type\":\"NEW\"},\"persistence\":{\"type\":\"NONE\"},\"steps\":[]}")
        .getBytes(StandardCharsets.UTF_8);
  }
}
