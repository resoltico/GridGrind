package dev.erst.gridgrind.contract.json;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.fasterxml.jackson.annotation.JsonIgnore;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Exercises tolerant syntax parsing, recovery, and transport-byte accounting directly. */
class RequestSyntaxSupportTest {
  @Test
  void parsesEveryJsonScalarKindAndEscapedText() {
    RequestSyntaxParseResult parsed =
        TolerantRequestJsonParser.parse(
            """
            {
              "text": "quote: \\" slash: \\/ control: \\b\\f\\n\\r\\t unicode: \\u20ac",
              "trueValue": true,
              "falseValue": false,
              "nothing": null,
              "number": -12.50e+2,
              "array": ["x", 2]
            }
            """
                .getBytes(StandardCharsets.UTF_8));

    RequestJsonObject root = assertInstanceOf(RequestJsonObject.class, parsed.root());
    assertTrue(parsed.problems().isEmpty());
    assertEquals(6, root.members().size());
    assertEquals(
        "quote: \" slash: / control: \b\f\n\r\t unicode: \u20ac",
        assertInstanceOf(RequestJsonString.class, member(root, "text").value()).value());
    assertEquals(
        new RequestJsonBoolean(member(root, "trueValue").value().byteOffset(), true),
        member(root, "trueValue").value());
    assertEquals(
        new RequestJsonBoolean(member(root, "falseValue").value().byteOffset(), false),
        member(root, "falseValue").value());
    assertInstanceOf(RequestJsonNull.class, member(root, "nothing").value());
    assertEquals(
        "-12.50e+2",
        assertInstanceOf(RequestJsonNumber.class, member(root, "number").value()).value());
    assertEquals(
        2,
        assertInstanceOf(RequestJsonArray.class, member(root, "array").value()).elements().size());
  }

  @Test
  void reportsScalarSyntaxFailuresWithoutDiscardingTheRestOfTheObject() {
    assertMessages(
        "{\"control\": \"before\u0001after\"}",
        "Control character is not allowed in a JSON string");
    assertMessages("{\"escape\": \"" + "\\" + "q\"}", "Invalid JSON string escape");
    assertTrue(
        TolerantRequestJsonParser.parse(
                ("{\"shortUnicode\": \"" + "\\" + "u12").getBytes(StandardCharsets.UTF_8))
            .problems()
            .stream()
            .map(RequestStructuralProblem::message)
            .anyMatch("Invalid unicode escape in JSON string"::equals));
    assertMessages(
        "{\"badUnicode\": \"" + "\\" + "uZZZZ\"}", "Invalid unicode escape in JSON string");
    assertTrue(
        TolerantRequestJsonParser.parse(("{\"ending\": \"" + "\\").getBytes(StandardCharsets.UTF_8))
            .problems()
            .stream()
            .map(RequestStructuralProblem::message)
            .anyMatch("Unterminated JSON string"::equals));
    assertInstanceOf(
        RequestJsonBoolean.class,
        TolerantRequestJsonParser.parse("true".getBytes(StandardCharsets.UTF_8)).root());
    assertInstanceOf(
        RequestJsonBoolean.class,
        TolerantRequestJsonParser.parse("false".getBytes(StandardCharsets.UTF_8)).root());
    assertInstanceOf(
        RequestJsonNull.class,
        TolerantRequestJsonParser.parse("null".getBytes(StandardCharsets.UTF_8)).root());
    assertInstanceOf(
        RequestJsonNumber.class,
        TolerantRequestJsonParser.parse("2".getBytes(StandardCharsets.UTF_8)).root());

    List<RequestStructuralProblem> independentlyParsedFailures =
        TolerantRequestJsonParser.parse(
                """
                {"literal": truth, "number": 01, "unknown": bogus, "next": true}
                """
                    .getBytes(StandardCharsets.UTF_8))
            .problems();
    assertEquals(
        List.of("Invalid JSON literal", "Invalid JSON number", "Expected a JSON value"),
        independentlyParsedFailures.stream().map(RequestStructuralProblem::message).toList());
  }

  @Test
  void recoversFromObjectArrayAndDocumentGrammarFaultsAtSafeBoundaries() {
    assertMessages("{unquoted: 1, \"next\": 2}", "Expected an object property name");
    assertMessages("{\"a\" 1, \"next\": 2}", "Expected ':' after object property name");
    assertMessages("{\"a\": 1 garbage, \"next\": 2}", "Expected ',' or '}' after object member");
    assertMessages(
        "{\"a\": 1", "Expected ',' or '}' after object member", "Expected '}' to close object");
    assertMessages(
        "[1 garbage, 2",
        "Expected ',' or ']' after array element",
        "Expected ',' or ']' after array element",
        "Expected ']' to close array");
    assertMessages("[1,", "Expected ']' to close array");
    assertMessages("{\"a\": 1,}", "Trailing comma is not permitted in a JSON object");
    assertMessages("[1,]", "Trailing comma is not permitted in a JSON array");
    assertMessages("true false", "Unexpected token after the root JSON value");
    assertMessages("", "Invalid JSON payload");
    assertMessages("   ", "Invalid JSON payload");
    assertMessages("{\"a\": }", "Expected a JSON value");
    assertInstanceOf(
        RequestJsonArray.class,
        TolerantRequestJsonParser.parse("[]".getBytes(StandardCharsets.UTF_8)).root());
    assertMessages("{broken}", "Expected an object property name");
    assertInstanceOf(
        RequestJsonInvalid.class,
        new RequestTolerantSyntaxParser(RequestUtf8Decoder.decode(new byte[0]))
            .parseDocument()
            .root());
  }

  @Test
  void reportsTheFirstByteForRootLevelNonJsonWhitespace() {
    RequestInvalidJson problem =
        assertInstanceOf(
            RequestInvalidJson.class,
            TolerantRequestJsonParser.parse(
                    Character.toString(0x0B).getBytes(StandardCharsets.UTF_8))
                .problems()
                .getFirst());

    assertEquals("Expected a JSON value", problem.message());
    assertEquals(0L, problem.byteOffset().orElseThrow());
  }

  @Test
  void cursorRecoverySkipsNestedAndQuotedDelimitersWithoutCrossingTheCurrentContainer() {
    RequestSyntaxCursor objectCursor =
        cursorFor("ignored {\"quoted, brace }\": [1, {\"nested\": 2}]}, next");
    objectCursor.recoverObjectMember();
    assertEquals(',', objectCursor.current());
    assertTrue(objectCursor.consume(','));
    assertFalse(objectCursor.consume(','));

    RequestSyntaxCursor arrayCursor =
        cursorFor("ignored [\"quoted, bracket ]\", {\"nested\": 2}], next");
    arrayCursor.recoverArrayElement();
    assertEquals(',', arrayCursor.current());

    RequestSyntaxCursor valueCursor = cursorFor("  :  value");
    valueCursor.recoverToValue();
    assertEquals('v', valueCursor.current());
    valueCursor.consumeInvalidValueToken();
    assertTrue(valueCursor.atEnd());
    assertEquals('\0', valueCursor.currentOrZero());

    RequestSyntaxCursor cursor = cursorFor(" a ");
    assertEquals(' ', cursor.currentOrZero());
    cursor.skipWhitespace();
    assertEquals(1, cursor.position());
    cursor.advanceBy(100);
    assertTrue(cursor.atEnd());
    cursor.advance();
    cursor.moveToEnd();
    assertEquals(cursor.length(), cursor.position());

    for (String delimiterTerminated : List.of("value ", "value,", "value]", "value}")) {
      RequestSyntaxCursor delimiterCursor = cursorFor(delimiterTerminated);
      delimiterCursor.consumeInvalidValueToken();
      assertEquals('v' == delimiterCursor.currentOrZero() ? 0 : 5, delimiterCursor.position());
    }
  }

  @Test
  void recoveryStateTreatsStringsEscapesAndNestedContainersAsOpaqueUntilBalanced() {
    RequestRecoveryState state = new RequestRecoveryState();
    assertFalse(state.shouldStop('x', ',', ']'));
    assertTrue(state.shouldStop(',', ',', ']'));
    assertTrue(state.shouldStop('}', ',', ']'));

    state.consume('{');
    state.consume('[');
    assertFalse(state.shouldStop(',', ',', ']'));
    state.consume(']');
    state.consume('}');
    assertTrue(state.shouldStop(']', ',', ']'));

    state.consume('"');
    assertFalse(state.shouldStop(',', ',', ']'));
    state.consume('\\');
    state.consume('"');
    assertFalse(state.shouldStop(',', ',', ']'));
    state.consume('"');
    assertTrue(state.shouldStop(',', ',', ']'));
    state.consume('x');
    assertTrue(state.shouldStop('x', 'x', 'y'));
    assertTrue(state.shouldStop('y', 'x', 'y'));
    assertFalse(state.shouldStop('z', 'x', 'y'));
  }

  @Test
  void classifiesCreatorTypesWithoutAssumingReflectionSpecificImplementations() {
    ParameterizedType singleArgument = parameterized(List.class, String.class);
    ParameterizedType twoArguments = parameterized(Map.class, String.class, Integer.class);
    Type unknown = new Type() {};

    assertEquals(String.class, RequestTypeSupport.rawType(String.class));
    assertEquals(List.class, RequestTypeSupport.rawType(singleArgument));
    assertEquals(Object.class, RequestTypeSupport.rawType(unknown));
    assertEquals(String.class, RequestTypeSupport.typeArgument(singleArgument));
    assertEquals(Object.class, RequestTypeSupport.typeArgument(twoArguments));
    assertEquals(Object.class, RequestTypeSupport.typeArgument(unknown));
    assertTrue(RequestNodeTypeClassifier.isOptional(parameterized(Optional.class, String.class)));
    assertTrue(RequestNodeTypeClassifier.isCollection(singleArgument));
    assertFalse(RequestNodeTypeClassifier.isCollection(String.class));
    for (Class<?> numberType :
        List.of(
            byte.class,
            Byte.class,
            short.class,
            Short.class,
            int.class,
            Integer.class,
            long.class,
            Long.class,
            float.class,
            Float.class,
            double.class,
            Double.class,
            BigDecimal.class,
            BigInteger.class)) {
      assertTrue(RequestTypeSupport.isNumberType(numberType));
    }
    assertFalse(RequestTypeSupport.isNumberType(String.class));
  }

  @Test
  void recognizesEveryLegalLiteralDelimiterWhenReadingJsonKeywords() {
    assertLiteralDelimiter("true,");
    assertLiteralDelimiter("true]");
    assertLiteralDelimiter("true}");
    assertLiteralDelimiter("true ");
    assertLiteralDelimiter("true\n");
    assertLiteralDelimiter("true\r");
    assertLiteralDelimiter("true\t");
  }

  @Test
  void rejectsAKeywordWhenAFollowingCharacterIsNotAJsonValueDelimiter() {
    List<RequestStructuralProblem> problems = new ArrayList<>();

    assertInstanceOf(
        RequestJsonInvalid.class,
        RequestJsonScalarReader.read(cursorFor("truex"), problems, "field"));

    assertEquals(
        List.of("Invalid JSON literal"),
        problems.stream().map(RequestStructuralProblem::message).toList());
  }

  @Test
  void rejectsNonJsonWhitespaceAsALiteralDelimiter() {
    assertInvalidLiteralDelimiter(Character.toString(0x0B));
    assertInvalidLiteralDelimiter(Character.toString(0x0C));
  }

  @Test
  void rejectsNonNumericScalarsAboveTheNumericAsciiRange() {
    List<RequestStructuralProblem> problems = new ArrayList<>();

    assertInstanceOf(
        RequestJsonInvalid.class,
        RequestJsonScalarReader.read(cursorFor("zebra"), problems, "field"));

    assertEquals(
        List.of("Expected a JSON value"),
        problems.stream().map(RequestStructuralProblem::message).toList());
  }

  private static void assertInvalidLiteralDelimiter(String suffix) {
    List<RequestStructuralProblem> problems = new ArrayList<>();

    assertInstanceOf(
        RequestJsonInvalid.class,
        RequestJsonScalarReader.read(cursorFor("true" + suffix), problems, "field"));
    assertEquals(
        List.of("Invalid JSON literal"),
        problems.stream().map(RequestStructuralProblem::message).toList());
  }

  @Test
  void honorsComponentAndAccessorJsonIgnoreAnnotationsInVisibleRecordContracts() {
    assertEquals(
        List.of("visible"),
        RequestObjectMembers.visibleRecordComponents(ComponentIgnoredRecord.class).stream()
            .map(component -> component.getName())
            .toList());
    assertTrue(RequestObjectMembers.visibleRecordComponents(AccessorIgnoredRecord.class).isEmpty());
    assertEquals(
        List.of(),
        RequestObjectMembers.optionalFieldNames(
            RequestObjectMembers.visibleRecordComponents(ComponentIgnoredRecord.class),
            List.of("visible")));
    assertEquals(
        List.of("visible"),
        RequestObjectMembers.optionalFieldNames(
            RequestObjectMembers.visibleRecordComponents(ComponentIgnoredRecord.class), List.of()));
    assertEquals("hidden", new AccessorIgnoredRecord("hidden").hidden());
  }

  private static void assertLiteralDelimiter(String text) {
    RequestSyntaxCursor cursor = cursorFor(text);
    RequestJsonNode node = RequestJsonScalarReader.read(cursor, new ArrayList<>(), "field");
    assertTrue(assertInstanceOf(RequestJsonBoolean.class, node).value());
    assertEquals(4, cursor.position());
  }

  @Test
  void decodesUtf8WithPreciseByteOffsetsAndRejectsMalformedInput() {
    RequestUtf8DecodeResult decoded =
        RequestUtf8Decoder.decode("A\u00a2\u20ac\uD83D\uDE00".getBytes(StandardCharsets.UTF_8));

    assertEquals("A\u00a2\u20ac\uD83D\uDE00", decoded.text());
    assertEquals(0L, decoded.byteOffsetAt(-1));
    assertEquals(0L, decoded.byteOffsetAt(0));
    assertEquals(1L, decoded.byteOffsetAt(1));
    assertEquals(3L, decoded.byteOffsetAt(2));
    assertEquals(6L, decoded.byteOffsetAt(3));
    assertEquals(6L, decoded.byteOffsetAt(4));
    assertEquals(10L, decoded.byteOffsetAt(5));
    assertEquals(10L, decoded.byteOffsetAt(100));
    assertTrue(decoded.problems().isEmpty());

    RequestUtf8DecodeResult malformed =
        RequestUtf8Decoder.decode(new byte[] {'{', (byte) 0xc3, (byte) 0x28});
    assertEquals("", malformed.text());
    RequestInvalidEncoding problem =
        assertInstanceOf(RequestInvalidEncoding.class, malformed.problems().getFirst());
    assertEquals(1L, problem.byteOffset().orElseThrow());
  }

  @Test
  void enforcesSyntaxValueAndStructuralProblemInvariants() {
    assertEquals(3L, new RequestJsonString(3, "value").byteOffset());
    assertEquals("value", new RequestJsonNumber(0, "value").value());
    assertThrows(IllegalArgumentException.class, () -> new RequestJsonNull(-1));
    assertThrows(NullPointerException.class, () -> new RequestJsonString(0, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new RequestJsonMember("name", -1, new RequestJsonNull(0)));
    assertThrows(NullPointerException.class, () -> new RequestJsonArray(0, null));

    assertEquals("Duplicate key 'field'", new RequestDuplicateKey("", "field", 0, 0).message());
    assertEquals(
        "Missing required field 'field'", new RequestMissingRequiredField("field").message());
    assertEquals(
        "Field 'field' must be omitted when absent; explicit null is not accepted.",
        new RequestExplicitNullField("field", 0).message());
    assertEquals("Unknown field 'field'", new RequestUnknownField("field", 0).message());
    assertEquals(
        "Unknown type value 'UNKNOWN'",
        new RequestUnknownTypeDiscriminator("type", "UNKNOWN", 0).message());
    assertEquals(
        "Unknown type value 'FILE'; use source.type='EXISTING'; similar valid values: EXISTING, NEW",
        new RequestUnknownTypeDiscriminator(
                "source.type",
                "FILE",
                List.of("EXISTING", "NEW"),
                Optional.of("use source.type='EXISTING'"),
                0)
            .message());
    assertEquals(
        "Unsupported value 'NOPE' for field 'version'; expected one of: V2",
        new RequestUnsupportedEnumValue("version", "NOPE", List.of("V2"), 0).message());
    assertEquals(
        "Field 'field' must be a JSON string",
        new RequestMalformedScalar("field", "a JSON string", 0).message());
    assertEquals(
        Optional.empty(),
        new RequestInvalidJson("bad", Optional.empty(), Optional.empty()).jsonPath());
    assertThrows(
        IllegalArgumentException.class,
        () -> new RequestInvalidJson(" ", Optional.empty(), Optional.empty()));
    assertThrows(IllegalArgumentException.class, () -> new RequestInvalidEncoding("bad", -1));
    assertThrows(IllegalArgumentException.class, () -> new RequestDuplicateKey("", "field", -1, 0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new RequestUnsupportedEnumValue("version", "NOPE", List.of(), 0));
    assertEquals(Optional.of("field"), new RequestMissingRequiredField("field").jsonPath());
    assertEquals(Optional.empty(), new RequestMissingRequiredField("field").byteOffset());
    assertEquals(Optional.of(0L), new RequestDuplicateKey("", "field", 0, 0).byteOffset());
    assertEquals(Optional.empty(), new RequestDuplicateKey("", "field", 0, 0).jsonPath());
  }

  private static RequestJsonMember member(RequestJsonObject object, String name) {
    return object.members().stream()
        .filter(member -> member.name().equals(name))
        .findFirst()
        .orElseThrow();
  }

  private static void assertMessages(String json, String... expectedMessages) {
    assertEquals(
        List.of(expectedMessages),
        TolerantRequestJsonParser.parse(json.getBytes(StandardCharsets.UTF_8)).problems().stream()
            .map(RequestStructuralProblem::message)
            .toList());
  }

  private static RequestSyntaxCursor cursorFor(String text) {
    return new RequestSyntaxCursor(
        RequestUtf8Decoder.decode(text.getBytes(StandardCharsets.UTF_8)));
  }

  private static ParameterizedType parameterized(Type rawType, Type... arguments) {
    return new ParameterizedType() {
      @Override
      public Type[] getActualTypeArguments() {
        return arguments.clone();
      }

      @Override
      public Type getRawType() {
        return rawType;
      }

      @Override
      public Type getOwnerType() {
        return null;
      }
    };
  }

  private record ComponentIgnoredRecord(@JsonIgnore String hidden, String visible) {}

  private record AccessorIgnoredRecord(String hidden) {
    @Override
    @JsonIgnore
    public String hidden() {
      return hidden;
    }
  }
}
