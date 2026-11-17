package dev.erst.gridgrind.json.codec;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.problem.AuthorityDocumentIssue;
import dev.erst.gridgrind.core.problem.InvalidAuthorityException;
import dev.erst.gridgrind.core.problem.MessageInvariant;
import dev.erst.gridgrind.core.problem.ValueInvariantException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;
import tools.jackson.databind.exc.InvalidDefinitionException;
import tools.jackson.databind.exc.ValueInstantiationException;

/** Strict authority intake and expected-versus-unexpected boundary classification. */
class AuthorityDocumentTest {
  private static final String VALID =
      """
      {"maxInputSnapshotBytes":0,"readableResources":[],"operationIds":[],
       "targetAuthority":{"type":"WORKBOOK_WIDE"},"publicationAuthority":{"type":"NONE"},
       "allowedSecretReferences":[],"acceptancePolicy":{"type":"MINIMUM_ONLY"}}
      """;

  @Test
  void validExplicitAuthorityBindsWithoutFilesystemAccess() throws java.io.IOException {
    var grant =
        read(
            VALID.replace(
                "\"readableResources\":[]",
                "\"readableResources\":[{\"type\":\"FILE\",\"path\":\"absent-input\"}]"));
    assertEquals(0, grant.maxInputSnapshotBytes());
    assertEquals(1, grant.readableResources().size());
  }

  @Test
  void syntaxShapeAndNullsRejectWholeDocumentsWithSafeOriginalCoordinates() {
    List<String> documents =
        List.of(
            "",
            "null",
            "[]",
            "{",
            VALID + " {}",
            VALID.replace(
                "\"type\":\"WORKBOOK_WIDE\"",
                "\"type\":\"WORKBOOK_WIDE\",\"type\":\"WORKBOOK_WIDE\""),
            VALID.replace("\"maxInputSnapshotBytes\":0,", ""),
            VALID.replace("\"maxInputSnapshotBytes\":0", "\"maxInputSnapshotBytes\":null"),
            VALID.replace("\"readableResources\":[]", "\"readableResources\":[null]"),
            VALID.replace("\"operationIds\":[]", "\"operationIds\":null"),
            VALID.replace("WORKBOOK_WIDE", "PRIVATE_UNKNOWN_TYPE"),
            VALID.replace(
                "\"maxInputSnapshotBytes\":0",
                "\"maxInputSnapshotBytes\":0,\"PRIVATE_UNKNOWN_FIELD\":\"private-content\""),
            VALID.replace(
                "\"readableResources\":[]",
                "\"readableResources\":[{\"type\":\"FILE\",\"path\":\"\\u0000\"}]"));
    for (String document : documents) {
      var failure = assertThrows(InvalidAuthorityException.class, () -> read(document));
      assertFalse(failure.getMessage().contains("private-content"));
      assertFalse(failure.getMessage().contains("PRIVATE_UNKNOWN"));
      assertFalse(failure.documentLocation().toString().contains("PRIVATE_UNKNOWN"));
    }
    var syntax =
        assertThrows(InvalidAuthorityException.class, () -> read("{\n\"maxInputSnapshotBytes\":}"));
    assertEquals(AuthorityDocumentIssue.SYNTAX, syntax.issue());
    assertTrue(syntax.documentLocation().byteOffsetValue().isPresent());
  }

  @Test
  void authorityScalarsRequireTheirDeclaredJsonTokenKinds() throws java.io.IOException {
    assertEquals(5, read(VALID.replace(":0", ":5")).maxInputSnapshotBytes());
    for (String value : List.of("\"5\"", "true", "5.0")) {
      assertThrows(InvalidAuthorityException.class, () -> read(VALID.replace(":0", ":" + value)));
    }
    assertThrows(
        InvalidAuthorityException.class,
        () -> read(VALID.replace("\"operationIds\":[]", "\"operationIds\":[5]")));
    for (String ordinal : List.of("0", "\"0\"")) {
      assertThrows(
          InvalidAuthorityException.class,
          () ->
              read(
                  VALID.replace(
                      "\"publicationAuthority\":{\"type\":\"NONE\"}",
                      "\"publicationAuthority\":{\"type\":\"SAVE_AS\",\"path\":\"result.xlsx\",\"ifExists\":"
                          + ordinal
                          + "}")));
    }
  }

  @Test
  void explicitCoreInvariantsBecomeAuthorityValuesWithoutGuessingConstructorMessages() {
    var budget =
        assertThrows(InvalidAuthorityException.class, () -> read(VALID.replace(":0", ":-1")));
    assertEquals(AuthorityDocumentIssue.VALUE, budget.issue());
    assertEquals(Optional.of("maxInputSnapshotBytes"), budget.documentLocation().jsonPathValue());
    var operation =
        assertThrows(
            InvalidAuthorityException.class,
            () ->
                read(
                    VALID.replace(
                        "\"operationIds\":[]", "\"operationIds\":[\"not-a-machine-token\"]")));
    assertEquals(AuthorityDocumentIssue.VALUE, operation.issue());
  }

  @Test
  void declaredNestedInvariantsRetainTheRealDocumentPath() {
    var secret =
        assertThrows(
            InvalidAuthorityException.class,
            () ->
                read(
                    VALID.replace(
                        "\"allowedSecretReferences\":[]",
                        "\"allowedSecretReferences\":[{\"id\":\"not a secret identifier\"}]")));
    assertEquals(AuthorityDocumentIssue.VALUE, secret.issue());
    assertEquals(
        Optional.of("allowedSecretReferences[0].id"), secret.documentLocation().jsonPathValue());
    String policy =
        "\"acceptancePolicy\":{\"type\":\"REQUIRE_ALL\",\"requirements\":[{\"type\":\"PRESERVE_INSPECTION_FACTS\",\"inspections\":[{\"stepId\":\"inspect\",\"target\":{\"type\":\"SHEET_BY_NAME\",\"name\":\"\"},\"query\":{\"type\":\"GET_SHEET_SUMMARY\"}}]}]}";
    var label =
        assertThrows(
            InvalidAuthorityException.class,
            () -> read(VALID.replace("\"acceptancePolicy\":{\"type\":\"MINIMUM_ONLY\"}", policy)));
    assertEquals(AuthorityDocumentIssue.VALUE, label.issue());
    assertEquals(
        Optional.of("acceptancePolicy.requirements[0].inspections[0].target.name"),
        label.documentLocation().jsonPathValue());
  }

  @Test
  void nestedAssertionConstraintsRemainDeclaredAuthorityValueFailures() {
    for (String assertion :
        List.of(
            "{\"type\":\"EXPECT_ANALYSIS_MAX_SEVERITY\",\"query\":{\"type\":\"GET_SHEET_SUMMARY\"},\"maximumSeverity\":\"ERROR\"}",
            "{\"type\":\"ALL_OF\",\"assertions\":[{\"type\":\"EXPECT_SHEET_PRESENT\"},{\"type\":\"EXPECT_CELL_VALUE\",\"expectedValue\":{\"type\":\"TEXT\",\"value\":\"x\"}}]}")) {
      String policy =
          "\"acceptancePolicy\":{\"type\":\"REQUIRE_ALL\",\"requirements\":[{\"type\":\"TERMINAL_ASSERTIONS\",\"assertions\":[{\"stepId\":\"check\",\"target\":{\"type\":\"SHEET_BY_NAME\",\"name\":\"S\"},\"assertion\":"
              + assertion
              + "}]}]}";
      var failure =
          assertThrows(
              InvalidAuthorityException.class,
              () ->
                  read(VALID.replace("\"acceptancePolicy\":{\"type\":\"MINIMUM_ONLY\"}", policy)));
      assertEquals(AuthorityDocumentIssue.VALUE, failure.issue());
    }
  }

  @Test
  void unknownNestedStepPropertyIsNotEchoedAsAKnownLocation() {
    String policy =
        "\"acceptancePolicy\":{\"type\":\"REQUIRE_ALL\",\"requirements\":[{\"type\":\"PRESERVE_INSPECTION_FACTS\",\"inspections\":[{\"stepId\":\"inspect\",\"target\":{\"type\":\"SHEET_BY_NAME\",\"name\":\"S\"},\"query\":{\"type\":\"GET_SHEET_SUMMARY\"},\"PRIVATE_UNKNOWN_FIELD\":\"private-content\"}]}]}";
    var failure =
        assertThrows(
            InvalidAuthorityException.class,
            () -> read(VALID.replace("\"acceptancePolicy\":{\"type\":\"MINIMUM_ONLY\"}", policy)));
    assertEquals(AuthorityDocumentIssue.SHAPE, failure.issue());
    assertFalse(failure.documentLocation().toString().contains("PRIVATE_UNKNOWN_FIELD"));
    assertFalse(failure.getMessage().contains("private-content"));
  }

  @Test
  void integralBudgetOverflowIsAValueFailureRatherThanInvalidJsonSyntax() {
    var overflow =
        assertThrows(
            InvalidAuthorityException.class,
            () -> read(VALID.replace(":0", ":9223372036854775808")));
    assertEquals(AuthorityDocumentIssue.VALUE, overflow.issue());
  }

  @Test
  void unknownConstructorsDefinitionsAndRawArgumentsRemainUnexpected() {
    var mapper = GridGrindJsonMapperSupport.GRANT_JSON_MAPPER;
    try (var parser = mapper.createParser("{}")) {
      parser.nextToken();
      var type = mapper.constructType(Object.class);
      for (RuntimeException cause :
          List.of(
              new IllegalArgumentException("private-content"),
              new NullPointerException(),
              new IllegalStateException())) {
        var wrapped = ValueInstantiationException.from(parser, "private-content", type, cause);
        var translated = AuthorityReadFailures.translate(wrapped, parser);
        assertInstanceOf(UnexpectedAuthorityReadException.class, translated);
        assertSame(wrapped, translated.getCause());
      }
      var fatal = new AssertionError("synthetic fatal");
      assertSame(
          fatal,
          assertThrows(
              AssertionError.class,
              () ->
                  AuthorityReadFailures.translate(
                      ValueInstantiationException.from(
                          parser, "known constructor wrapper", type, fatal),
                      parser)));
      assertSame(
          fatal,
          assertThrows(
              AssertionError.class,
              () ->
                  AuthorityReadFailures.translate(
                      ValueInstantiationException.from(
                          parser,
                          "reflection constructor wrapper",
                          type,
                          new java.lang.reflect.InvocationTargetException(fatal)),
                      parser)));
      var definition = InvalidDefinitionException.from(parser, "private-content", type);
      assertInstanceOf(
          UnexpectedAuthorityReadException.class,
          AuthorityReadFailures.translate(definition, parser));
      assertInstanceOf(
          UnexpectedAuthorityReadException.class,
          AuthorityReadFailures.translate(new IllegalArgumentException("private-content"), parser));
      var declared =
          new ValueInvariantException(
              new MessageInvariant("Owned invariant", Optional.of("ownedField")));
      var known =
          AuthorityReadFailures.translate(
              ValueInstantiationException.from(parser, "not inspected", type, declared), parser);
      assertInstanceOf(InvalidAuthorityException.class, known);
      assertEquals(
          Optional.of("ownedField"),
          ((InvalidAuthorityException) known).documentLocation().jsonPathValue());
    }
  }

  @Test
  void collectionWrapperAcceptsOnlyADirectDeclaredOwnerFailure() {
    var mapper = GridGrindJsonMapperSupport.GRANT_JSON_MAPPER;
    try (var parser = mapper.createParser("{}")) {
      var declared =
          new dev.erst.gridgrind.core.problem.InvalidRequestException(
              new MessageInvariant("Owned value rule", Optional.of("logical")),
              Optional.of("actual"),
              Optional.empty(),
              Optional.empty(),
              null);
      var wrapper =
          tools.jackson.databind.DatabindException.from(parser, "known framework wrapper", declared)
              .prependPath(Object.class, "outer");
      assertEquals(tools.jackson.databind.DatabindException.class, wrapper.getClass());
      var known =
          assertInstanceOf(
              InvalidAuthorityException.class, AuthorityReadFailures.translate(wrapper, parser));
      assertEquals(Optional.of("outer.actual"), known.documentLocation().jsonPathValue());
      for (RuntimeException failure :
          List.of(new IllegalArgumentException("private"), new IllegalStateException(declared))) {
        var unknown =
            tools.jackson.databind.DatabindException.from(
                    parser, "unknown framework wrapper", failure)
                .prependPath(Object.class, "outer");
        assertInstanceOf(
            UnexpectedAuthorityReadException.class,
            AuthorityReadFailures.translate(unknown, parser));
      }
    }
  }

  private static dev.erst.gridgrind.core.grant.ExecutionGrant read(String json)
      throws java.io.IOException {
    return GridGrindJson.readGrant(
        json.getBytes(StandardCharsets.UTF_8), Path.of("/unopened-authority-root"));
  }
}
