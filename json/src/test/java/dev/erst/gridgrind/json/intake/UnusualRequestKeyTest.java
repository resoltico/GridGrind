package dev.erst.gridgrind.json.intake;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.problem.RequestInvalidJson;
import dev.erst.gridgrind.core.problem.RequestUnknownField;
import dev.erst.gridgrind.json.codec.GridGrindJson;
import java.nio.charset.StandardCharsets;
import java.util.Optional;
import java.util.stream.Stream;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.MethodSource;

/** Keeps arbitrary JSON keys distinct from path syntax and the root request location. */
class UnusualRequestKeyTest {
  @ParameterizedTest
  @MethodSource("keys")
  void rejectsUnknownKeysWithTheirExactUnambiguousLocation(String encodedKey, String path) {
    String prefix =
        "{\"protocolVersion\":1,\"source\":{\"type\":\"NEW\"},\"persistence\":{\"type\":\"NONE\"},\"steps\":[],";
    byte[] bytes = (prefix + encodedKey + ":0}").getBytes(StandardCharsets.UTF_8);
    long offset = prefix.getBytes(StandardCharsets.UTF_8).length;
    var analysis = GridGrindJson.analyzeRequest(bytes);

    assertEquals(1, analysis.structuralProblems().size());
    var unknown =
        assertInstanceOf(RequestUnknownField.class, analysis.structuralProblems().getFirst());
    assertEquals(Optional.of(path), unknown.jsonPath());
    assertEquals(Optional.of(offset), unknown.byteOffset());
    assertEquals(Optional.of(offset), analysis.byteOffsetAt(path));
    assertEquals(Optional.of(0L), analysis.byteOffsetAt("request"));
    assertTrue(analysis.completePlan().isEmpty());
    assertTrue(analysis.bindingFailures().isEmpty());
    assertEquals(
        Optional.of(offset),
        RequestJsonTokenLocationSupport.byteOffsetAt(
            TolerantRequestJsonParser.parse(bytes).root(), path));
  }

  static Stream<Arguments> keys() {
    return Stream.of(
        Arguments.of("\"\"", "[\"\"]"),
        Arguments.of("\"    \"", "[\"    \"]"),
        Arguments.of("\"a.b\"", "[\"a.b\"]"),
        Arguments.of("\"x[0]\"", "[\"x[0]\"]"),
        Arguments.of("\"quote\\\"slash\\\\\"", "[\"quote\\\"slash\\\\\"]"),
        Arguments.of("\"request\"", "[\"request\"]"),
        Arguments.of("\"κλειδί\"", "[\"κλειδί\"]"),
        Arguments.of("\"label\"", "label"));
  }

  @Test
  void locatesMixedQuotedKeysAndArrayIndexesUsingUtf8ByteOffsets() {
    String json = "{\"κλειδί\":[{\"x[0]\":[{\"a.b\":0}]}]}";
    String path = "[\"κλειδί\"][0][\"x[0]\"][0][\"a.b\"]";
    long offset =
        json.substring(0, json.indexOf("\"a.b\"")).getBytes(StandardCharsets.UTF_8).length;
    var root = TolerantRequestJsonParser.parse(json.getBytes(StandardCharsets.UTF_8)).root();
    assertEquals(Optional.of(offset), RequestJsonTokenLocationSupport.byteOffsetAt(root, path));
    assertEquals(offset, RequestJsonTokenLocationSupport.index(root).get(path));
    assertEquals(
        Optional.empty(),
        RequestJsonTokenLocationSupport.byteOffsetAt(root, "κλειδί[0].x[0][0].a.b"));
  }

  @Test
  void recoversSyntaxProblemsUnderAWhitespaceKey() {
    byte[] bytes = "{\"    \" 0}".getBytes(StandardCharsets.UTF_8);
    var result = GridGrindJson.analyzeRequest(bytes);
    var syntax =
        result.structuralProblems().stream()
            .filter(RequestInvalidJson.class::isInstance)
            .map(RequestInvalidJson.class::cast)
            .toList();
    assertFalse(syntax.isEmpty());
    assertEquals(Optional.of("[\"    \"]"), syntax.getFirst().jsonPath());
    assertEquals(Optional.of(1L), syntax.getFirst().byteOffset());
    assertTrue(result.completePlan().isEmpty());
  }
}
