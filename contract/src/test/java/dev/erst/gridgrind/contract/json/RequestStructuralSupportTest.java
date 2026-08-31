package dev.erst.gridgrind.contract.json;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.OoxmlEncryptionInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.selector.Selector;
import java.nio.charset.StandardCharsets;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Covers structural validation and independent fragment binding through its package contract. */
class RequestStructuralSupportTest {
  @Test
  void validatesRecordScalarAndUnionShapesWithoutStoppingAtTheFirstProblem() {
    List<RequestStructuralProblem> problems = new ArrayList<>();

    RequestRecordValidator.validate(
        new RequestJsonString(0, "not-an-object"),
        OoxmlEncryptionInput.class,
        "encryption",
        0,
        problems);
    RequestRecordValidator.validate(
        object(member("password", new RequestJsonNumber(1, "7"))),
        OoxmlEncryptionInput.class,
        "encryption",
        0,
        problems);
    RequestNonContainerNodeValidator.validate(
        new RequestJsonString(2, "wrong"), int.class, "integer", 2, problems);
    RequestNonContainerNodeValidator.validate(
        new RequestJsonString(3, "wrong"), boolean.class, "boolean", 3, problems);
    RequestNonContainerNodeValidator.validate(
        new RequestJsonString(3, "text"), StringBuilder.class, "text", 3, problems);
    RequestNodeValidator.validateEnum(
        new RequestJsonString(4, "NOT_A_VERSION"),
        GridGrindProtocolVersion.class,
        "protocolVersion",
        4,
        problems);
    RequestNodeValidator.validateEnum(
        new RequestJsonNumber(5, "2"),
        GridGrindProtocolVersion.class,
        "protocolVersion",
        5,
        problems);
    RequestUnionValidator.validateUnion(
        new RequestJsonNumber(6, "2"), WorkbookPlan.WorkbookSource.class, "source", 6, problems);
    RequestUnionValidator.validateUnion(
        object(), WorkbookPlan.WorkbookSource.class, "source", 7, problems);
    RequestUnionValidator.validateUnion(
        object(member("type", new RequestJsonNumber(8, "2"))),
        WorkbookPlan.WorkbookSource.class,
        "source",
        8,
        problems);
    RequestUnionValidator.validateUnion(
        object(member("type", new RequestJsonString(9, "UNKNOWN"))),
        WorkbookPlan.WorkbookSource.class,
        "source",
        9,
        problems);
    RequestUnionValidator.validateSelector(
        object(member("type", new RequestJsonString(10, "WORKBOOK_CURRENT"))),
        "target",
        10,
        problems);

    assertEquals(
        List.of(
            "Field 'encryption' must be a JSON object",
            "Field 'encryption.password' must be a JSON string",
            "Field 'integer' must be a JSON number",
            "Field 'boolean' must be a JSON boolean",
            "Unsupported value 'NOT_A_VERSION' for field 'protocolVersion'; expected one of: V2",
            "Field 'protocolVersion' must be a JSON string",
            "Field 'source' must be a JSON object",
            "Missing required field 'source.type'",
            "Field 'source.type' must be a JSON string type id",
            "Unknown type value 'UNKNOWN'; valid values: EXISTING, NEW"),
        problems.stream().map(RequestStructuralProblem::message).toList());
  }

  @Test
  void validatesCollectionNullAndNestedRecordBranches() {
    List<RequestStructuralProblem> problems = new ArrayList<>();

    RequestCollectionValidator.validate(
        new RequestJsonString(0, "not-an-array"), String.class, "values", 0, problems);
    RequestCollectionValidator.validate(
        new RequestJsonArray(1, List.of(new RequestJsonNumber(2, "7"))),
        String.class,
        "values",
        1,
        problems);
    RequestNodeValidator.validateNode(new RequestJsonNull(3), String.class, "value", 3, problems);
    RequestNonContainerNodeValidator.validate(
        object(member("password", new RequestJsonString(4, "secret"))),
        OoxmlEncryptionInput.class,
        "encryption",
        4,
        problems);
    RequestNonContainerNodeValidator.validate(
        object(member("type", new RequestJsonString(5, "NEW"))),
        WorkbookPlan.WorkbookSource.class,
        "source",
        5,
        problems);
    RequestNonContainerNodeValidator.validate(
        object(member("type", new RequestJsonString(6, "WORKBOOK_CURRENT"))),
        Selector.class,
        "target",
        6,
        problems);

    assertEquals(
        List.of(
            "Field 'values' must be a JSON array",
            "Field 'values[0]' must be a JSON string",
            "Field 'value' must be omitted when absent; explicit null is not accepted."),
        problems.stream().map(RequestStructuralProblem::message).toList());
  }

  @Test
  void validatesThatNumericAndTemporalScalarsAreRepresentableByTheirCreatorTypes() {
    List<RequestStructuralProblem> problems = new ArrayList<>();

    RequestNonContainerNodeValidator.validate(
        new RequestJsonNumber(10, "999999999999999999999"), int.class, "integer", 10, problems);
    RequestNonContainerNodeValidator.validate(
        new RequestJsonNumber(11, "1e400"), double.class, "floating", 11, problems);
    RequestNonContainerNodeValidator.validate(
        new RequestJsonNumber(12, "1e400"), Float.class, "singlePrecision", 12, problems);
    RequestNonContainerNodeValidator.validate(
        new RequestJsonString(13, "2026-02-30"), LocalDate.class, "date", 13, problems);
    RequestNonContainerNodeValidator.validate(
        new RequestJsonString(14, "2026-02-30T12:00"),
        LocalDateTime.class,
        "dateTime",
        14,
        problems);
    RequestNonContainerNodeValidator.validate(
        new RequestJsonNumber(15, "1.5"), Float.class, "validSinglePrecision", 15, problems);
    RequestNonContainerNodeValidator.validate(
        new RequestJsonNumber(16, "20260705"), LocalDate.class, "dateKind", 16, problems);
    RequestNonContainerNodeValidator.validate(
        new RequestJsonObject(17, List.of()), Object.class, "untyped", 17, problems);

    assertEquals(
        List.of(
            "Field 'integer' must be a JSON integer between -2147483648 and 2147483647",
            "Field 'floating' number '1e400' cannot be represented exactly; store identifiers or precision-sensitive values as TEXT.",
            "Field 'singlePrecision' number '1e400' cannot be represented exactly; store identifiers or precision-sensitive values as TEXT.",
            "Field 'date' must be an ISO-8601 calendar date",
            "Field 'dateTime' must be an ISO-8601 local date-time",
            "Field 'dateKind' must be a JSON string"),
        problems.stream().map(RequestStructuralProblem::message).toList());
    assertEquals(
        2, problems.stream().filter(RequestNumberNotRepresentable.class::isInstance).count());
  }

  @Test
  void rejectsSealedCreatorFamiliesThatFailToDeclareTheirRequiredDiscriminator() {
    assertThrows(
        IllegalStateException.class,
        () ->
            RequestUnionValidator.validateUnion(
                object(), UndiscriminatedUnion.class, "union", 0, new ArrayList<>()));
  }

  @Test
  void independentlyBindsValidFragmentsAndRetainsMalformedSiblingLocations() {
    RequestJsonObject node =
        object(
            member("text", new RequestJsonString(0, "value")),
            member("number", new RequestJsonNumber(1, "7")),
            member("truth", new RequestJsonBoolean(2, true)),
            member("nothing", new RequestJsonNull(3)),
            member(
                "values",
                new RequestJsonArray(
                    4, List.of(new RequestJsonString(5, "one"), new RequestJsonString(6, "two")))));
    List<RequestStructuralProblem> noProblems = new ArrayList<>();

    Optional<Map> bound = bindNode(node, Map.class, "carrier", noProblems);
    assertEquals("value", bound.orElseThrow().get("text"));
    assertEquals("7", bound.orElseThrow().get("number").toString());
    assertEquals(true, bound.orElseThrow().get("truth"));
    assertEquals(null, bound.orElseThrow().get("nothing"));
    assertEquals(List.of("one", "two"), bound.orElseThrow().get("values"));
    assertTrue(bindNode(new RequestJsonInvalid(7), Map.class, "carrier", noProblems).isEmpty());

    List<RequestStructuralProblem> siblingProblems =
        List.of(new RequestUnknownField("carrier.text", 0));
    assertTrue(bindNode(node, Map.class, "carrier", siblingProblems).isEmpty());
    assertTrue(bindNode(node, Map.class, "other", siblingProblems).isPresent());
    assertTrue(bindMember(Optional.empty(), Map.class, "carrier", noProblems).isEmpty());
    assertTrue(
        bindMember(Optional.of(member("carrier", node)), Map.class, "carrier", noProblems)
            .isPresent());
    assertTrue(
        bindNode(
                new RequestJsonString(8, "not-an-encryption-object"),
                OoxmlEncryptionInput.class,
                "encryption",
                noProblems)
            .isEmpty());

    RequestJsonObject duplicateObject =
        object(
            member("first", new RequestJsonString(9, "kept")),
            member("first", new RequestJsonString(10, "discarded")));
    assertEquals(
        "kept", RequestFragmentBinder.toJsonNode(duplicateObject).path("first").stringValue());
    assertEquals(
        "one",
        RequestFragmentBinder.toJsonNode(
                new RequestJsonArray(11, List.of(new RequestJsonString(12, "one"))))
            .get(0)
            .stringValue());
    assertTrue(RequestFragmentBinder.toJsonNode(new RequestJsonNumber(13, "7")).isIntegralNumber());
    assertFalse(
        RequestFragmentBinder.toJsonNode(new RequestJsonNumber(14, "1.5")).isIntegralNumber());
    assertFalse(
        RequestFragmentBinder.toJsonNode(new RequestJsonNumber(15, "1e3")).isIntegralNumber());
    assertFalse(
        RequestFragmentBinder.toJsonNode(new RequestJsonNumber(16, "1E3")).isIntegralNumber());
    assertThrows(
        IllegalStateException.class,
        () -> RequestFragmentBinder.toJsonNode(new RequestJsonInvalid(17)));

    for (RequestStructuralProblem problem :
        List.of(
            new RequestDuplicateKey("", "carrier", 0, 0),
            new RequestDuplicateKey("carrier", "duplicate", 0, 0),
            new RequestDuplicateKey("carrier.child", "duplicate", 0, 0),
            new RequestDuplicateKey("carrier[0]", "duplicate", 0, 0),
            new RequestUnknownField("carrier", 0))) {
      assertTrue(bindNode(node, Map.class, "carrier", List.of(problem)).isEmpty());
    }
    assertTrue(
        bindNode(
                node,
                Map.class,
                "carrier",
                List.of(new RequestDuplicateKey("other", "duplicate", 0, 0)))
            .isPresent());
    assertTrue(
        bindNode(
                node,
                Map.class,
                "carrier",
                List.of(new RequestInvalidJson("grammar", Optional.empty(), Optional.empty())))
            .isPresent());
    assertTrue(
        bindNode(
                node,
                Map.class,
                "carrier",
                List.of(new RequestInvalidJson("grammar", Optional.of("carrier"), Optional.of(0L))))
            .isEmpty());
  }

  @Test
  void collectsAllStepContainerFailuresAndBindsEachPayloadFamilyIndependently() {
    assertTrue(bindSteps(Optional.empty(), List.of()).isEmpty());

    RequestJsonMember explicitNullSteps = member("steps", new RequestJsonNull(0));
    List<RequestStructuralProblem> explicitNullProblems = new ArrayList<>();
    RequestStepFragmentCollector.validateAll(List.of(explicitNullSteps), explicitNullProblems);
    assertTrue(bindSteps(Optional.of(explicitNullSteps), explicitNullProblems).isEmpty());

    RequestJsonMember malformedSteps = member("steps", new RequestJsonString(1, "not-an-array"));
    List<RequestStructuralProblem> malformedProblems = new ArrayList<>();
    RequestStepFragmentCollector.validateAll(List.of(malformedSteps), malformedProblems);
    assertTrue(bindSteps(Optional.of(malformedSteps), malformedProblems).isEmpty());
    assertTrue(bindSteps(Optional.of(malformedSteps), List.of()).isEmpty());

    RequestJsonArray steps =
        parsedSteps(
            """
            {
              "steps": [
                null,
                7,
                {"stepId":"missing-payload","target":{"type":"WORKBOOK_CURRENT"}},
                {"stepId":"two-payloads","target":{"type":"WORKBOOK_CURRENT"},"action":{},"assertion":{}},
                {"stepId":"action","target":{"type":"WORKBOOK_CURRENT"},"action":{"type":"SET_SHEET_ZOOM"}},
                {"stepId":"assertion","target":{"type":"WORKBOOK_CURRENT"},"assertion":{"type":"EXPECT_SHEET_PRESENT"}},
                {"stepId":"query","target":{"type":"WORKBOOK_CURRENT"},"query":{"type":"GET_WORKBOOK_SUMMARY"}}
              ]
            }
            """);
    RequestJsonMember validSteps = member("steps", steps);
    List<RequestStructuralProblem> stepProblems = new ArrayList<>();
    RequestStepFragmentCollector.validateAll(List.of(validSteps), stepProblems);
    List<RequestBoundFragments.Step> collected =
        bindSteps(Optional.of(validSteps), stepProblems).orElseThrow();

    List<String> messages = new ArrayList<>();
    messages.addAll(explicitNullProblems.stream().map(RequestStructuralProblem::message).toList());
    messages.addAll(malformedProblems.stream().map(RequestStructuralProblem::message).toList());
    messages.addAll(stepProblems.stream().map(RequestStructuralProblem::message).toList());
    assertEquals(
        List.of(0, 1, 2, 3, 4, 5, 6),
        collected.stream().map(RequestBoundFragments.Step::index).toList());
    assertTrue(collected.get(6).value().isPresent());
    assertEquals(
        List.of(
            "Field 'steps' must be omitted when absent; explicit null is not accepted.",
            "Field 'steps' must be a JSON array",
            "Field 'steps[0]' must be omitted when absent; explicit null is not accepted.",
            "Field 'steps[1]' must be a JSON object",
            "Field 'steps[2]' must be an object containing exactly one of action, assertion, or query",
            "Missing required field 'steps[3].action.type'",
            "Missing required field 'steps[3].assertion.type'",
            "Field 'steps[3]' must be an object containing exactly one of action, assertion, or query",
            "Missing required field 'steps[4].action.zoomPercent'"),
        messages);
    assertThrows(
        IllegalStateException.class, () -> RequestStepFragmentCollector.payloadType("other"));
  }

  @Test
  void buildsCompletePlansOnlyFromCompleteAndConstructorValidFragments() {
    RequestBoundRoot completeRoot =
        new RequestBoundRoot(
            Optional.of(GridGrindProtocolVersion.V2),
            Optional.of("plan"),
            Optional.of(new WorkbookPlan.WorkbookSource.New()),
            Optional.of(new WorkbookPlan.WorkbookPersistence.None()),
            Optional.of(ExecutionPolicyInput.defaults()),
            Optional.of(FormulaEnvironmentInput.empty()));
    RequestBoundFragments complete =
        new RequestBoundFragments(completeRoot, Optional.of(List.of()));

    assertTrue(complete.completePlan().isPresent());
    assertEquals("plan", complete.planId().orElseThrow());
    assertEquals(GridGrindProtocolVersion.V2, complete.protocolVersion().orElseThrow());
    assertTrue(complete.source().isPresent());
    assertTrue(complete.persistence().isPresent());
    assertTrue(complete.execution().isPresent());
    assertTrue(complete.formulaEnvironment().isPresent());
    assertTrue(complete.steps().isPresent());

    assertFalse(
        new RequestBoundFragments(
                new RequestBoundRoot(
                    Optional.empty(),
                    Optional.empty(),
                    Optional.of(new WorkbookPlan.WorkbookSource.New()),
                    Optional.of(new WorkbookPlan.WorkbookPersistence.None()),
                    Optional.of(ExecutionPolicyInput.defaults()),
                    Optional.of(FormulaEnvironmentInput.empty())),
                Optional.of(List.of()))
            .completePlan()
            .isPresent());
    List<RequestBoundFragments> incompleteFragments =
        List.of(
            new RequestBoundFragments(
                new RequestBoundRoot(
                    Optional.of(GridGrindProtocolVersion.V2),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.of(new WorkbookPlan.WorkbookPersistence.None()),
                    Optional.of(ExecutionPolicyInput.defaults()),
                    Optional.of(FormulaEnvironmentInput.empty())),
                Optional.of(List.of())),
            new RequestBoundFragments(
                new RequestBoundRoot(
                    Optional.of(GridGrindProtocolVersion.V2),
                    Optional.empty(),
                    Optional.of(new WorkbookPlan.WorkbookSource.New()),
                    Optional.empty(),
                    Optional.of(ExecutionPolicyInput.defaults()),
                    Optional.of(FormulaEnvironmentInput.empty())),
                Optional.of(List.of())),
            new RequestBoundFragments(
                new RequestBoundRoot(
                    Optional.of(GridGrindProtocolVersion.V2),
                    Optional.empty(),
                    Optional.of(new WorkbookPlan.WorkbookSource.New()),
                    Optional.of(new WorkbookPlan.WorkbookPersistence.None()),
                    Optional.empty(),
                    Optional.of(FormulaEnvironmentInput.empty())),
                Optional.of(List.of())),
            new RequestBoundFragments(
                new RequestBoundRoot(
                    Optional.of(GridGrindProtocolVersion.V2),
                    Optional.empty(),
                    Optional.of(new WorkbookPlan.WorkbookSource.New()),
                    Optional.of(new WorkbookPlan.WorkbookPersistence.None()),
                    Optional.of(ExecutionPolicyInput.defaults()),
                    Optional.empty()),
                Optional.of(List.of())));
    for (RequestBoundFragments incomplete : incompleteFragments) {
      assertTrue(incomplete.completePlan().isEmpty());
    }
    assertFalse(
        new RequestBoundFragments(
                completeRoot,
                Optional.of(List.of(new RequestBoundFragments.Step(0, Optional.empty()))))
            .completePlan()
            .isPresent());
    assertTrue(new RequestBoundFragments(completeRoot, Optional.empty()).completePlan().isEmpty());
    assertThrows(
        IllegalArgumentException.class, () -> new RequestBoundFragments.Step(-1, Optional.empty()));

    RequestAnalysis invalid =
        new RequestAnalysis(complete, List.of(new RequestUnknownField("extra", 0)));
    assertFalse(invalid.isStructurallyValid());
    assertTrue(invalid.completePlan().isEmpty());
    assertTrue(new RequestAnalysis(complete, List.of()).isStructurallyValid());
    assertEquals(complete.completePlan(), new RequestAnalysis(complete, List.of()).completePlan());

    var duplicateStep =
        GridGrindJson.readRequest(
                """
                {
                  "protocolVersion": "V2",
                  "source": {"type": "NEW"},
                  "persistence": {"type": "NONE"},
                  "steps": [{
                    "stepId": "duplicate",
                    "target": {"type": "WORKBOOK_CURRENT"},
                    "query": {"type": "GET_WORKBOOK_SUMMARY"}
                  }]
                }
                """)
            .steps()
            .getFirst();
    RequestBoundFragments constructorInvalid =
        new RequestBoundFragments(
            completeRoot,
            Optional.of(
                List.of(
                    new RequestBoundFragments.Step(0, Optional.of(duplicateStep)),
                    new RequestBoundFragments.Step(1, Optional.of(duplicateStep)))));
    assertTrue(new RequestAnalysis(constructorInvalid, List.of()).completePlan().isEmpty());
    assertThrows(InvalidRequestException.class, constructorInvalid::completePlan);
  }

  @Test
  void maintainsObjectMemberPathAndPresenceRules() {
    RequestJsonObject object =
        object(
            member("known", new RequestJsonString(0, "first")),
            member("known", new RequestJsonString(1, "second")),
            member("unknown", new RequestJsonString(2, "value")));
    RequestObjectMembers.Index members = RequestObjectMembers.index(object);
    List<RequestStructuralProblem> problems = new ArrayList<>();

    assertEquals(
        "first",
        assertInstanceOf(RequestJsonString.class, members.member("known").orElseThrow().value())
            .value());
    assertEquals(2, members.membersNamed("known").size());
    assertTrue(members.member("known").isPresent());
    assertTrue(members.member("missing").isEmpty());
    RequestObjectMembers.collectFields(
        members, "parent", List.of("required", "known"), List.of(), problems);

    assertEquals("parent.child", RequestObjectMembers.childPath("parent", "child"));
    assertEquals("child", RequestObjectMembers.childPath("", "child"));
    assertEquals(
        List.of("Missing required field 'parent.required'", "Unknown field 'parent.unknown'"),
        problems.stream().map(RequestStructuralProblem::message).toList());
  }

  @Test
  void mapsEveryStructuralProblemToTheSingularExecutionExceptionContract() {
    assertInstanceOf(
        InvalidEncodingException.class,
        GridGrindJson.structuralException(new RequestInvalidEncoding("bad encoding", 0)));
    assertInstanceOf(
        InvalidJsonException.class,
        GridGrindJson.structuralException(
            new RequestInvalidJson("bad JSON", Optional.empty(), Optional.empty())));
    assertInstanceOf(
        InvalidJsonException.class,
        GridGrindJson.structuralException(new RequestDuplicateKey("", "duplicate", 0, 0)));
    for (RequestStructuralProblem problem :
        List.of(
            new RequestUnknownField("unknown", 0),
            new RequestMissingRequiredField("missing"),
            new RequestExplicitNullField("nullValue", 0),
            new RequestMissingTypeDiscriminator("type"),
            new RequestUnknownTypeDiscriminator("type", "UNKNOWN", 0),
            new RequestMalformedScalar("type", "a JSON string type id", 0),
            new RequestMalformedScalar("number", "a JSON number", 0))) {
      assertInstanceOf(
          InvalidRequestShapeException.class, GridGrindJson.structuralException(problem));
    }
    assertEquals(Optional.of("unknown"), new RequestUnknownField("unknown", 0).jsonPath());
    assertEquals(Optional.of(0L), new RequestUnknownField("unknown", 0).byteOffset());
    assertEquals(
        Optional.of("type"), new RequestUnknownTypeDiscriminator("type", "UNKNOWN", 0).jsonPath());
    assertEquals(
        Optional.of(0L), new RequestUnknownTypeDiscriminator("type", "UNKNOWN", 0).byteOffset());
  }

  private static RequestJsonArray parsedSteps(String json) {
    RequestJsonObject root =
        assertInstanceOf(
            RequestJsonObject.class,
            TolerantRequestJsonParser.parse(json.getBytes(StandardCharsets.UTF_8)).root());
    return assertInstanceOf(RequestJsonArray.class, member(root, "steps").value());
  }

  private static <T> Optional<T> bindNode(
      RequestJsonNode node,
      Class<T> type,
      String jsonPath,
      List<RequestStructuralProblem> structuralProblems) {
    return RequestFragmentBinder.bindNode(
        node,
        type,
        jsonPath,
        structuralProblems,
        new ArrayList<>(),
        RequestDiagnosticRedactor.empty());
  }

  private static <T> Optional<T> bindMember(
      Optional<RequestJsonMember> member,
      Class<T> type,
      String jsonPath,
      List<RequestStructuralProblem> structuralProblems) {
    return RequestFragmentBinder.bindMember(
        member,
        type,
        jsonPath,
        structuralProblems,
        new ArrayList<>(),
        RequestDiagnosticRedactor.empty());
  }

  private static Optional<List<RequestBoundFragments.Step>> bindSteps(
      Optional<RequestJsonMember> stepsMember, List<RequestStructuralProblem> structuralProblems) {
    return RequestStepFragmentCollector.bind(
        stepsMember, structuralProblems, new ArrayList<>(), RequestDiagnosticRedactor.empty());
  }

  private static RequestJsonObject object(RequestJsonMember... members) {
    return new RequestJsonObject(0, List.of(members));
  }

  private static RequestJsonMember member(String name, RequestJsonNode value) {
    return new RequestJsonMember(name, value.byteOffset(), value);
  }

  private static RequestJsonMember member(RequestJsonObject object, String name) {
    return object.members().stream()
        .filter(member -> member.name().equals(name))
        .findFirst()
        .orElseThrow();
  }

  /** Synthetic sealed union deliberately missing its required discriminator declaration. */
  private sealed interface UndiscriminatedUnion permits UndiscriminatedVariant {}

  private record UndiscriminatedVariant(String value) implements UndiscriminatedUnion {}
}
