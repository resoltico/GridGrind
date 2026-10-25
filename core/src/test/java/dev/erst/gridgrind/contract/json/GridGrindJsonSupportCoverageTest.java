package dev.erst.gridgrind.contract.json;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolContractSupport;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import java.io.IOException;
import java.math.BigInteger;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;
import tools.jackson.core.JacksonException;
import tools.jackson.core.JsonParser;
import tools.jackson.core.ObjectReadContext;
import tools.jackson.core.exc.StreamConstraintsException;
import tools.jackson.core.json.JsonFactory;
import tools.jackson.databind.JavaType;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.exc.InvalidFormatException;
import tools.jackson.databind.exc.InvalidTypeIdException;
import tools.jackson.databind.exc.MismatchedInputException;
import tools.jackson.databind.exc.UnrecognizedPropertyException;
import tools.jackson.databind.exc.ValueInstantiationException;
import tools.jackson.databind.json.JsonMapper;
import tools.jackson.databind.node.ObjectNode;

/** Covers internal JSON-support branches that are otherwise hard to drive through one request. */
class GridGrindJsonSupportCoverageTest {
  @Test
  void cleanJacksonMessageStripsSourceClausesOnly() {
    assertEquals(
        "Cannot deserialize value of type `int` from String \"abc\" as a subtype of `ignored`"
            + " (for POJO property 'count') (but could if coercion were enabled)",
        GridGrindJsonProblemMessageSupport.cleanJacksonMessage(
            "Cannot deserialize value of type `int` from String \"abc\" as a subtype of `ignored`"
                + " (for POJO property 'count') (but could if coercion were enabled)"
                + " (line noise [Source: REDACTED; line: 1, column: 2])"));
  }

  @Test
  void requestPayloadAndShapeFallbacksCoverTheDedicatedTransportBranches() throws IOException {
    assertInstanceOf(
        InvalidRequestException.class,
        GridGrindJsonProblemMessageSupport.invalidRequestPayload(
            new StreamConstraintsException("too large")));

    JsonMapper mapper = GridGrindJsonMapperSupport.REQUEST_JSON_MAPPER;
    try (JsonParser parser = parser("{}")) {
      parser.nextToken();
      MismatchedInputException mismatch =
          MismatchedInputException.from(parser, String.class, "Expected a string");
      MessageShape shape =
          assertInstanceOf(
              MessageShape.class,
              GridGrindJsonRequestShapeProblemSupport.mismatchedInputProblem(
                      mapper.readTree("{}"), String.class, mismatch)
                  .orElseThrow());
      assertEquals(
          "JSON object is missing required fields or has the wrong shape", shape.message());
    }
  }

  @Test
  void codecSupportMapsStringGrammarFailuresAndNestedArrayNullsBeforeBinding() throws IOException {
    JsonMapper mapper = GridGrindJsonMapperSupport.REQUEST_JSON_MAPPER;

    assertEquals(
        "{}",
        GridGrindJsonCodecSupport.readTree(
                "{}", mapper, GridGrindJsonProblemMessageSupport::invalidRequestPayload)
            .toString());

    assertInstanceOf(
        InvalidJsonException.class,
        assertThrows(
            InvalidJsonException.class,
            () ->
                GridGrindJsonCodecSupport.readTree(
                    "{", mapper, GridGrindJsonProblemMessageSupport::invalidRequestPayload)));
    assertInstanceOf(
        InvalidRequestShapeException.class,
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJsonCodecSupport.decodeTree(
                    mapper.readTree("[null]"),
                    mapper,
                    Object.class,
                    GridGrindJsonProblemMessageSupport::invalidRequestPayload)));
    assertInstanceOf(
        InvalidRequestShapeException.class,
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJsonCodecSupport.decodeTree(
                    mapper.nullNode(),
                    mapper,
                    Object.class,
                    GridGrindJsonProblemMessageSupport::invalidRequestPayload)));
  }

  @Test
  void requestProblemDetectorFindsMissingRequiredFieldsAndTypeDiscriminatorsStructurally()
      throws IOException {
    JsonMapper mapper = GridGrindJsonMapperSupport.REQUEST_JSON_MAPPER;
    JsonNode requestNode = mapper.readTree("{\"steps\":[]}");
    JacksonException missingProtocolVersion =
        structuralFailure(mapper, requestNode, WorkbookPlan.class);

    MissingRequiredField missingField =
        assertInstanceOf(
            MissingRequiredField.class,
            GridGrindJsonRequestProblemDetector.detect(
                    requestNode, WorkbookPlan.class, missingProtocolVersion)
                .orElseThrow());
    assertEquals("protocolVersion", missingField.jsonPathValue());

    JsonNode queryNode = mapper.readTree("{}");
    JacksonException missingQueryType = structuralFailure(mapper, queryNode, InspectionQuery.class);
    MissingTypeDiscriminator missingType =
        assertInstanceOf(
            MissingTypeDiscriminator.class,
            GridGrindJsonRequestProblemDetector.detect(
                    queryNode, InspectionQuery.class, missingQueryType)
                .orElseThrow());
    assertEquals("type", missingType.jsonPathValue());
  }

  @Test
  void requestContractSupportUsesTheEffectiveCreatorAndDiscriminatorContract() {
    assertEquals(
        List.of("protocolVersion", "source", "persistence", "steps"),
        GridGrindProtocolContractSupport.requiredFieldNames(WorkbookPlan.class));
    assertEquals(
        Optional.of("type"),
        GridGrindProtocolContractSupport.discriminatorField(InspectionQuery.class));
    assertEquals(
        Optional.empty(), GridGrindProtocolContractSupport.discriminatorField(WorkbookPlan.class));
  }

  @Test
  void requestProblemHelpersCoverUnknownFieldsEnumsSubtypeFallbacksAndDefaultBranches()
      throws IOException {
    JsonMapper mapper = GridGrindJsonMapperSupport.REQUEST_JSON_MAPPER;
    JsonNode rootNode = mapper.createObjectNode();
    JsonNode queryNode = mapper.readTree("{\"query\":{\"type\":\"NOPE\"}}");
    JsonNode nonStringQueryNode = mapper.readTree("{\"query\":{\"type\":1}}");
    JsonNode queryNodeWithPresentDiscriminator =
        mapper.readTree("{\"query\":{\"type\":\"BROKEN\"}}");
    JsonNode topLevelTypeNode = mapper.readTree("{\"type\":1}");
    JsonNode scalarTypeNode = mapper.readTree("1");
    UnrecognizedPropertyException unknownField =
        UnrecognizedPropertyException.from(
            parser("{}"), WorkbookPlan.class, "extra", java.util.Collections.emptyList());
    InvalidFormatException unsupportedEnum =
        InvalidFormatException.from(parser("null"), "bad enum", null, SampleEnum.class);
    InvalidFormatException namedEnum =
        InvalidFormatException.from(parser("\"OMEGA\""), "bad enum", "OMEGA", SampleEnum.class);
    InvalidFormatException nonEnumFormat =
        InvalidFormatException.from(parser("\"abc\""), "bad number", "abc", Integer.class);
    InvalidFormatException nullTargetFormat =
        new InvalidFormatException(parser("\"abc\""), "bad number", "abc", null);
    InvalidTypeIdException missingTypeWithoutBase =
        InvalidTypeIdException.from(parser("{}"), "bad type", null, null);
    InvalidTypeIdException missingTypeWithoutAnnotation =
        InvalidTypeIdException.from(
            parser("{}"), "bad type", mapper.constructType(WorkbookPlan.class), null);
    InvalidTypeIdException missingTypeWithBlankDiscriminator =
        InvalidTypeIdException.from(
            parser("{}"), "bad type", mapper.constructType(BlankTypeDiscriminator.class), null);
    InvalidTypeIdException nestedMissingTypeWithPresentDiscriminator =
        (InvalidTypeIdException)
            InvalidTypeIdException.from(
                    parser("{}"), "bad type", mapper.constructType(InspectionQuery.class), null)
                .prependPath(WorkbookPlan.class, "query");
    InvalidTypeIdException typo =
        InvalidTypeIdException.from(
            parser("\"COVE_SHEET\""),
            "bad type",
            mapper.constructType(WorkbookMutationAction.class),
            "COVE_SHEET");
    InvalidTypeIdException unknownTypedValue =
        InvalidTypeIdException.from(
            parser("\"NOPE\""),
            "bad type",
            mapper.constructType(WorkbookMutationAction.class),
            "NOPE");
    InvalidTypeIdException nestedUnknownTypedValue =
        (InvalidTypeIdException)
            InvalidTypeIdException.from(
                    parser("\"NOPE\""),
                    "bad type",
                    mapper.constructType(InspectionQuery.class),
                    "NOPE")
                .prependPath(WorkbookPlan.class, "query");
    InvalidTypeIdException nonStringTypedValue =
        (InvalidTypeIdException)
            InvalidTypeIdException.from(
                    parser("1"), "bad type", mapper.constructType(InspectionQuery.class), "1")
                .prependPath(InspectionQuery.class, "type")
                .prependPath(WorkbookPlan.class, "query");
    InvalidTypeIdException topLevelFieldPathNonStringTypedValue =
        (InvalidTypeIdException)
            InvalidTypeIdException.from(
                    parser("1"), "bad type", mapper.constructType(InspectionQuery.class), "1")
                .prependPath(InspectionQuery.class, "type");
    InvalidTypeIdException pathlessUnknownTypedValue =
        InvalidTypeIdException.from(
            parser("1"), "bad type", mapper.constructType(InspectionQuery.class), "1");
    InvalidTypeIdException sourceFile =
        (InvalidTypeIdException)
            InvalidTypeIdException.from(
                    parser("\"FILE\""),
                    "bad type",
                    mapper.constructType(WorkbookPlan.WorkbookSource.class),
                    "FILE")
                .prependPath(WorkbookPlan.class, "source");

    UnknownField unknown =
        assertInstanceOf(
            UnknownField.class,
            GridGrindJsonRequestProblemDetector.detect(rootNode, WorkbookPlan.class, unknownField)
                .orElseThrow());
    assertEquals("extra", unknown.jsonPathValue());

    UnsupportedValue unsupported =
        assertInstanceOf(
            UnsupportedValue.class,
            GridGrindJsonRequestProblemDetector.detect(
                    rootNode, WorkbookPlan.class, unsupportedEnum)
                .orElseThrow());
    assertEquals("null", unsupported.value());
    assertEquals(Optional.empty(), unsupported.jsonPath());
    assertEquals(
        "OMEGA",
        assertInstanceOf(
                UnsupportedValue.class,
                GridGrindJsonRequestTypeProblemSupport.enumValueProblem(namedEnum))
            .value());
    assertTrue(
        GridGrindJsonRequestProblemDetector.detect(
                rootNode, WorkbookPlan.class, new StreamConstraintsException("too deep"))
            .isEmpty());
    assertInstanceOf(
        MessageShape.class,
        GridGrindJsonRequestProblemDetector.detect(rootNode, WorkbookPlan.class, nonEnumFormat)
            .orElseThrow());
    assertInstanceOf(
        MessageShape.class,
        GridGrindJsonRequestProblemDetector.detect(rootNode, Object.class, nullTargetFormat)
            .orElseThrow());

    assertEquals("steps[0]", GridGrindJsonRequestTypeProblemSupport.appendPath("steps", "[0]"));
    assertEquals(
        "extra", GridGrindJsonRequestTypeProblemSupport.fullUnknownFieldPath(unknownField));
    assertTrue(
        GridGrindJsonRequestTypeProblemSupport.pathAlreadyTargetsField(
            "steps[0].action.type", "type"));
    assertTrue(GridGrindJsonRequestTypeProblemSupport.pathAlreadyTargetsField("type", "type"));
    assertFalse(
        GridGrindJsonRequestTypeProblemSupport.pathAlreadyTargetsField("steps[0].action", "type"));
    assertEquals(
        "type",
        assertInstanceOf(
                MissingTypeDiscriminator.class,
                GridGrindJsonRequestTypeProblemSupport.typeProblem(
                    rootNode, missingTypeWithoutBase))
            .jsonPathValue());
    assertEquals(
        "type",
        assertInstanceOf(
                MissingTypeDiscriminator.class,
                GridGrindJsonRequestTypeProblemSupport.typeProblem(
                    rootNode, missingTypeWithoutAnnotation))
            .jsonPathValue());
    assertEquals(
        "type",
        assertInstanceOf(
                MissingTypeDiscriminator.class,
                GridGrindJsonRequestTypeProblemSupport.typeProblem(
                    rootNode, missingTypeWithBlankDiscriminator))
            .jsonPathValue());
    MessageShape pathlessMissingTypeFallback =
        assertInstanceOf(
            MessageShape.class,
            GridGrindJsonRequestTypeProblemSupport.typeProblem(
                scalarTypeNode, missingTypeWithoutBase));
    assertEquals(Optional.empty(), pathlessMissingTypeFallback.jsonPath());
    MessageShape nestedMissingTypeFallback =
        assertInstanceOf(
            MessageShape.class,
            GridGrindJsonRequestTypeProblemSupport.typeProblem(
                queryNodeWithPresentDiscriminator, nestedMissingTypeWithPresentDiscriminator));
    assertEquals(Optional.of("query"), nestedMissingTypeFallback.jsonPath());
    assertEquals(
        "NOPE",
        assertInstanceOf(
                UnknownTypeValue.class,
                GridGrindJsonRequestTypeProblemSupport.typeProblem(rootNode, unknownTypedValue))
            .typeId());
    UnknownTypeValue nestedUnknownType =
        assertInstanceOf(
            UnknownTypeValue.class,
            GridGrindJsonRequestTypeProblemSupport.typeProblem(queryNode, nestedUnknownTypedValue));
    assertEquals(Optional.of("query.type"), nestedUnknownType.jsonPath());
    ActionableShapeMessage nonStringType =
        assertInstanceOf(
            ActionableShapeMessage.class,
            GridGrindJsonRequestTypeProblemSupport.typeProblem(
                nonStringQueryNode, nonStringTypedValue));
    assertEquals("Field 'type' must be a string", nonStringType.message());
    assertEquals(Optional.of("query.type"), nonStringType.jsonPath());
    ActionableShapeMessage topLevelNonStringType =
        assertInstanceOf(
            ActionableShapeMessage.class,
            GridGrindJsonRequestTypeProblemSupport.typeProblem(
                topLevelTypeNode, topLevelFieldPathNonStringTypedValue));
    assertEquals(Optional.of("type"), topLevelNonStringType.jsonPath());
    UnknownTypeValue pathlessUnknownType =
        assertInstanceOf(
            UnknownTypeValue.class,
            GridGrindJsonRequestTypeProblemSupport.typeProblem(
                scalarTypeNode, pathlessUnknownTypedValue));
    assertEquals(Optional.of("type"), pathlessUnknownType.jsonPath());
    assertInstanceOf(
        MessageShape.class,
        GridGrindJsonRequestTypeProblemSupport.typeProblem(
            mapper.createObjectNode().put("type", "BROKEN"), missingTypeWithoutBase));

    assertEquals(
        "JSON object is missing required fields or has the wrong shape",
        GridGrindJsonSubtypeProblemSupport.unknownTypeValueMessage(missingTypeWithoutBase));
    assertTrue(
        GridGrindJsonSubtypeProblemSupport.unknownTypeValueMessage(typo)
            .contains("similar valid values"));
    assertTrue(
        GridGrindJsonSubtypeProblemSupport.unknownTypeValueMessage(sourceFile)
            .contains("source.type='EXISTING'"));
  }

  @Test
  void payloadMetadataRetainsOnlyConcreteTerminalPropertyNames() {
    assertEquals(
        Optional.empty(), GridGrindJsonPayloadMetadataSupport.terminalContainerName(List.of()));
    assertEquals(
        Optional.of("source"),
        GridGrindJsonPayloadMetadataSupport.terminalContainerName(
            List.of(new JacksonException.Reference(WorkbookPlan.class, "source"))));
  }

  @Test
  void subtypeProblemMessagesKeepAlreadyQualifiedDiscriminatorPathsStable() throws IOException {
    JsonMapper mapper = GridGrindJsonMapperSupport.REQUEST_JSON_MAPPER;
    InvalidTypeIdException sourceFileAtType =
        (InvalidTypeIdException)
            InvalidTypeIdException.from(
                    parser("\"FILE\""),
                    "bad type",
                    mapper.constructType(WorkbookPlan.WorkbookSource.class),
                    "FILE")
                .prependPath(WorkbookPlan.WorkbookSource.class, "type")
                .prependPath(WorkbookPlan.class, "source");
    InvalidTypeIdException topLevelType =
        (InvalidTypeIdException)
            InvalidTypeIdException.from(
                    parser("\"FILE\""),
                    "bad type",
                    mapper.constructType(WorkbookPlan.WorkbookSource.class),
                    "FILE")
                .prependPath(WorkbookPlan.WorkbookSource.class, "type");

    assertTrue(
        GridGrindJsonSubtypeProblemSupport.unknownTypeValueMessage(sourceFileAtType)
            .contains("source.type='EXISTING'"));
    assertTrue(
        GridGrindJsonSubtypeProblemSupport.unknownTypeValueMessage(topLevelType)
            .startsWith("Unknown type value 'FILE'"));
  }

  @Test
  void requestShapeHelpersCoverFloatingPointAndFallbackTargetTypes() throws IOException {
    JsonMapper mapper = GridGrindJsonMapperSupport.REQUEST_JSON_MAPPER;
    JsonNode requestNode = mapper.readTree("{\"steps\":[]}");
    MismatchedInputException floatingInteger =
        MismatchedInputException.from(
            parser("2.5"),
            Integer.class,
            "Cannot coerce Floating-point value (2.5) to `int` value");
    MismatchedInputException nullMessageShape =
        MismatchedInputException.from(parser("{}"), Integer.class, null);
    MismatchedInputException missingProtocolVersion =
        MismatchedInputException.from(parser("{}"), (Class<?>) null, "Cannot construct instance");
    MismatchedInputException productOwnedMismatch =
        (MismatchedInputException)
            MismatchedInputException.from(
                    parser("{}"), Object.class, "Field 'target' must be a JSON object")
                .prependPath(WorkbookPlan.class, "target");
    ValueInstantiationException missingFromValueInstantiation =
        ValueInstantiationException.from(
            parser("{}"), "Cannot construct instance", (JavaType) null);

    MessageShape floatingProblem =
        assertInstanceOf(
            MessageShape.class,
            GridGrindJsonRequestShapeProblemSupport.mismatchedInputProblem(
                    requestNode, Integer.class, floatingInteger)
                .orElseThrow());
    assertEquals("JSON value must be an integer value", floatingProblem.message());
    assertEquals(
        "JSON value has the wrong shape for this field",
        assertInstanceOf(
                MessageShape.class,
                GridGrindJsonRequestShapeProblemSupport.mismatchedInputProblem(
                        requestNode, Integer.class, nullMessageShape)
                    .orElseThrow())
            .message());
    MessageShape genericProductMessage =
        assertInstanceOf(
            MessageShape.class,
            GridGrindJsonRequestShapeProblemSupport.mismatchedInputProblem(
                    requestNode, Object.class, productOwnedMismatch)
                .orElseThrow());
    assertEquals("JSON value has the wrong shape for this field", genericProductMessage.message());
    assertEquals(Optional.of("target"), genericProductMessage.jsonPath());
    assertEquals(
        "protocolVersion",
        assertInstanceOf(
                MissingRequiredField.class,
                GridGrindJsonRequestShapeProblemSupport.mismatchedInputProblem(
                        requestNode, WorkbookPlan.class, missingProtocolVersion)
                    .orElseThrow())
            .jsonPathValue());
    assertEquals(
        "protocolVersion",
        assertInstanceOf(
                MissingRequiredField.class,
                GridGrindJsonRequestShapeProblemSupport.valueInstantiationProblem(
                        requestNode, WorkbookPlan.class, missingFromValueInstantiation)
                    .orElseThrow())
            .jsonPathValue());
  }

  @Test
  void requestShapeAndTypeHelpersCoverPreservedMessagesUnknownFieldPathsAndIntegralTargets()
      throws Exception {
    JsonMapper mapper = GridGrindJsonMapperSupport.REQUEST_JSON_MAPPER;
    JsonNode requestNode = mapper.readTree("{\"steps\":[]}");
    ValueInstantiationException blankValueInstantiation =
        ValueInstantiationException.from(parser("{}"), " ", (JavaType) null);
    UnrecognizedPropertyException fieldlessUnknown =
        new SyntheticUnrecognizedPropertyException("extra", List.of());
    UnrecognizedPropertyException alreadyQualifiedUnknown =
        new SyntheticUnrecognizedPropertyException(
            "extra",
            List.of(
                new JacksonException.Reference(WorkbookPlan.class, "steps"),
                new JacksonException.Reference(WorkbookPlan.class, "extra")));
    UnrecognizedPropertyException nestedUnknown =
        new SyntheticUnrecognizedPropertyException(
            "extra",
            List.of(
                new JacksonException.Reference(WorkbookPlan.class, "steps"),
                new JacksonException.Reference(WorkbookPlan.class, 0)));

    assertEquals(
        "JSON object is missing required fields or has the wrong shape",
        assertInstanceOf(
                MessageShape.class,
                GridGrindJsonRequestShapeProblemSupport.valueInstantiationProblem(
                        requestNode, Object.class, blankValueInstantiation)
                    .orElseThrow())
            .message());

    assertEquals(
        "extra", GridGrindJsonRequestTypeProblemSupport.fullUnknownFieldPath(fieldlessUnknown));
    assertEquals(
        "steps.extra",
        GridGrindJsonRequestTypeProblemSupport.fullUnknownFieldPath(alreadyQualifiedUnknown));
    assertEquals(
        "steps[0].extra",
        GridGrindJsonRequestTypeProblemSupport.fullUnknownFieldPath(nestedUnknown));

    try (JsonParser floatingTokenParser = parser("2.5")) {
      floatingTokenParser.nextToken();
      assertTrue(
          GridGrindJsonValueProblemSupport.isFloatingPointIntoInteger(
              MismatchedInputException.from(floatingTokenParser, Integer.class, "bad")));
    }

    assertTrue(floatingPointIntoIntegerDetected(byte.class));
    assertTrue(floatingPointIntoIntegerDetected(short.class));
    assertTrue(floatingPointIntoIntegerDetected(int.class));
    assertTrue(floatingPointIntoIntegerDetected(long.class));
    assertTrue(floatingPointIntoIntegerDetected(Byte.class));
    assertTrue(floatingPointIntoIntegerDetected(Short.class));
    assertTrue(floatingPointIntoIntegerDetected(Integer.class));
    assertTrue(floatingPointIntoIntegerDetected(Long.class));
    assertTrue(floatingPointIntoIntegerDetected(BigInteger.class));
    assertFalse(floatingPointIntoIntegerDetected(String.class));
  }

  @Test
  void rawStepPayloadShapeValidationCoversNoOpAndConflictCases() {
    JsonMapper mapper = JsonMapper.builder().build();
    ObjectNode request = mapper.createObjectNode();

    assertDoesNotThrow(
        () ->
            GridGrindJsonStepPayloadShapeSupport.rejectInvalidStepPayloadShapes(mapper.nullNode()));
    assertDoesNotThrow(
        () -> GridGrindJsonStepPayloadShapeSupport.rejectInvalidStepPayloadShapes(request));

    request.put("steps", 3);
    assertDoesNotThrow(
        () -> GridGrindJsonStepPayloadShapeSupport.rejectInvalidStepPayloadShapes(request));

    request.putArray("steps").add(3);
    assertDoesNotThrow(
        () -> GridGrindJsonStepPayloadShapeSupport.rejectInvalidStepPayloadShapes(request));

    ObjectNode validStep = mapper.createObjectNode();
    validStep.putObject("action").put("type", "ENSURE_SHEET");
    request.set("steps", mapper.createArrayNode().add(validStep));
    assertDoesNotThrow(
        () -> GridGrindJsonStepPayloadShapeSupport.rejectInvalidStepPayloadShapes(request));

    request.set(
        "steps",
        mapper
            .createArrayNode()
            .add(mapper.createObjectNode().put("stepId", "empty").putObject("target")));
    InvalidRequestShapeException missingPayload =
        assertThrows(
            InvalidRequestShapeException.class,
            () -> GridGrindJsonStepPayloadShapeSupport.rejectInvalidStepPayloadShapes(request));
    assertEquals(Optional.of("steps[0]"), missingPayload.jsonPath());

    assertEquals(
        Optional.of("steps[0].assertion"),
        conflictingPayloadPath(
            mapper,
            """
            { "stepId": "both", "target": {}, "action": {}, "assertion": {} }
            """));
    assertEquals(
        Optional.of("steps[0].query"),
        conflictingPayloadPath(
            mapper,
            """
            { "stepId": "both", "target": {}, "action": {}, "query": {} }
            """));
    assertEquals(
        Optional.of("steps[0].query"),
        conflictingPayloadPath(
            mapper,
            """
            { "stepId": "both", "target": {}, "assertion": {}, "query": {} }
            """));
  }

  @Test
  void conflictingPayloadFieldRejectsImpossibleCallers() {
    JsonMapper mapper = JsonMapper.builder().build();
    ObjectNode emptyStep = mapper.createObjectNode();
    ObjectNode actionOnlyStep = mapper.createObjectNode();
    ObjectNode assertionOnlyStep = mapper.createObjectNode();
    actionOnlyStep.putObject("action");
    assertionOnlyStep.putObject("assertion");

    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () -> GridGrindJsonStepPayloadShapeSupport.conflictingPayloadField(emptyStep));
    assertThrows(
        IllegalStateException.class,
        () -> GridGrindJsonStepPayloadShapeSupport.conflictingPayloadField(actionOnlyStep));
    assertThrows(
        IllegalStateException.class,
        () -> GridGrindJsonStepPayloadShapeSupport.conflictingPayloadField(assertionOnlyStep));

    assertEquals("conflictingPayloadField requires at least two payloads", failure.getMessage());
  }

  private static Optional<String> conflictingPayloadPath(JsonMapper mapper, String stepJson) {
    ObjectNode request = mapper.createObjectNode();
    request.set(
        "steps", mapper.createArrayNode().add(assertDoesNotThrow(() -> mapper.readTree(stepJson))));
    InvalidRequestShapeException failure =
        assertThrows(
            InvalidRequestShapeException.class,
            () -> GridGrindJsonStepPayloadShapeSupport.rejectInvalidStepPayloadShapes(request));
    return failure.jsonPath();
  }

  private static JacksonException structuralFailure(
      JsonMapper mapper, JsonNode node, Class<?> targetType) {
    return assertThrows(JacksonException.class, () -> mapper.treeToValue(node, targetType));
  }

  private static boolean floatingPointIntoIntegerDetected(Class<?> targetType) throws IOException {
    try (JsonParser parser = parser("2.5")) {
      return GridGrindJsonValueProblemSupport.isFloatingPointIntoInteger(
          MismatchedInputException.from(
              parser,
              targetType,
              "Cannot coerce Floating-point value (2.5) to `%s` value"
                  .formatted(targetType.getSimpleName())));
    }
  }

  private static tools.jackson.core.JsonParser parser(String json) throws IOException {
    return new JsonFactory()
        .createParser(
            ObjectReadContext.empty(),
            new java.io.ByteArrayInputStream(
                json.getBytes(java.nio.charset.StandardCharsets.UTF_8)));
  }

  @SuppressWarnings("PMD.CommentRequired")
  private enum SampleEnum {
    ALPHA,
    BETA
  }

  /** Synthetic property-binding exception whose path is supplied directly for path-shape tests. */
  private static final class SyntheticUnrecognizedPropertyException
      extends UnrecognizedPropertyException {
    private static final long serialVersionUID = 1L;

    private final List<JacksonException.Reference> syntheticPath;

    private SyntheticUnrecognizedPropertyException(
        String propertyName, List<JacksonException.Reference> syntheticPath) {
      super(null, "bad property", null, WorkbookPlan.class, propertyName, List.of());
      this.syntheticPath = List.copyOf(syntheticPath);
    }

    @Override
    public List<JacksonException.Reference> getPath() {
      return syntheticPath;
    }
  }

  @SuppressWarnings("PMD.CommentRequired")
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "")
  /* Synthetic subtype surface with a blank discriminator property for fallback coverage. */
  private interface BlankTypeDiscriminator {}
}
