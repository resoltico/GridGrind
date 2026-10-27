package dev.erst.gridgrind.json.intake;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.model.security.OoxmlEncryptionSpec;
import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.InvalidEncodingException;
import dev.erst.gridgrind.core.problem.InvalidJsonException;
import dev.erst.gridgrind.core.problem.InvalidRequestException;
import dev.erst.gridgrind.core.problem.InvalidRequestShapeException;
import dev.erst.gridgrind.core.problem.RequestBoundFragments;
import dev.erst.gridgrind.core.problem.RequestBoundRoot;
import dev.erst.gridgrind.core.problem.RequestDuplicateKey;
import dev.erst.gridgrind.core.problem.RequestExplicitNullField;
import dev.erst.gridgrind.core.problem.RequestIntake;
import dev.erst.gridgrind.core.problem.RequestInvalidEncoding;
import dev.erst.gridgrind.core.problem.RequestInvalidJson;
import dev.erst.gridgrind.core.problem.RequestMalformedScalar;
import dev.erst.gridgrind.core.problem.RequestMissingRequiredField;
import dev.erst.gridgrind.core.problem.RequestMissingTypeDiscriminator;
import dev.erst.gridgrind.core.problem.RequestNumberNotRepresentable;
import dev.erst.gridgrind.core.problem.RequestStructuralProblem;
import dev.erst.gridgrind.core.problem.RequestUnknownField;
import dev.erst.gridgrind.core.problem.RequestUnknownTypeDiscriminator;
import dev.erst.gridgrind.core.vocabulary.PlanId;
import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import dev.erst.gridgrind.json.codec.GridGrindJson;
import java.nio.charset.StandardCharsets;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Covers structural validation and independent fragment binding through its package contract. */
class RequestStructuralSupportTest {
  @Test
  void validatesThatNumericAndTemporalScalarsAreRepresentableByTheirCreatorTypes() {
    List<RequestStructuralProblem> problems = new ArrayList<>();

    RequestScalarTokens.validate(
        new RequestJsonNumber(10, "999999999999999999999"), int.class, "integer", 10, problems);
    RequestScalarTokens.validate(
        new RequestJsonNumber(11, "1e400"), double.class, "floating", 11, problems);
    RequestScalarTokens.validate(
        new RequestJsonNumber(12, "1e400"), Float.class, "singlePrecision", 12, problems);
    RequestScalarTokens.validate(
        new RequestJsonString(13, "2026-02-30"), LocalDate.class, "date", 13, problems);
    RequestScalarTokens.validate(
        new RequestJsonString(14, "2026-02-30T12:00"),
        LocalDateTime.class,
        "dateTime",
        14,
        problems);
    RequestScalarTokens.validate(
        new RequestJsonNumber(15, "1.5"), Float.class, "validSinglePrecision", 15, problems);
    RequestScalarTokens.validate(
        new RequestJsonNumber(16, "20260705"), LocalDate.class, "dateKind", 16, problems);

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
  void independentlyBindsValidFragmentsAndRetainsMalformedSiblingLocations() {
    List<RequestStructuralProblem> noProblems = new ArrayList<>();

    RequestJsonObject encryption =
        object(member("passwordRef", object(member("id", new RequestJsonString(0, "protected")))));
    assertEquals(
        "protected",
        bindNode(encryption, OoxmlEncryptionSpec.class, "carrier", noProblems)
            .orElseThrow()
            .passwordRef()
            .id());
    assertTrue(
        bindNode(new RequestJsonInvalid(7), OoxmlEncryptionSpec.class, "carrier", noProblems)
            .isEmpty());
    List<RequestStructuralProblem> siblingProblems =
        List.of(new RequestUnknownField("carrier.passwordRef", 0));
    assertTrue(
        bindNode(encryption, OoxmlEncryptionSpec.class, "carrier", siblingProblems).isEmpty());
    assertTrue(
        bindNode(encryption, OoxmlEncryptionSpec.class, "other", siblingProblems).isPresent());
    assertTrue(
        bindMember(Optional.empty(), OoxmlEncryptionSpec.class, "carrier", noProblems).isEmpty());
    assertTrue(
        bindMember(
                Optional.of(member("carrier", encryption)),
                OoxmlEncryptionSpec.class,
                "carrier",
                noProblems)
            .isPresent());

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
      assertTrue(
          bindNode(encryption, OoxmlEncryptionSpec.class, "carrier", List.of(problem)).isEmpty());
    }
    assertTrue(
        bindNode(
                encryption,
                OoxmlEncryptionSpec.class,
                "carrier",
                List.of(new RequestDuplicateKey("other", "duplicate", 0, 0)))
            .isPresent());
    assertTrue(
        bindNode(
                encryption,
                OoxmlEncryptionSpec.class,
                "carrier",
                List.of(new RequestInvalidJson("grammar", Optional.empty(), Optional.empty())))
            .isPresent());
    assertTrue(
        bindNode(
                encryption,
                OoxmlEncryptionSpec.class,
                "carrier",
                List.of(new RequestInvalidJson("grammar", Optional.of("carrier"), Optional.of(0L))))
            .isEmpty());
  }

  @Test
  void collectsAllStepContainerFailuresAndBindsEachPayloadFamilyIndependently() {
    assertTrue(bindSteps(Optional.empty(), List.of()).isEmpty());

    RequestJsonMember explicitNullSteps = member("steps", new RequestJsonNull(0));
    List<RequestStructuralProblem> explicitNullProblems = new ArrayList<>();
    validateSteps(explicitNullSteps, explicitNullProblems);
    assertTrue(bindSteps(Optional.of(explicitNullSteps), explicitNullProblems).isEmpty());

    RequestJsonMember malformedSteps = member("steps", new RequestJsonString(1, "not-an-array"));
    List<RequestStructuralProblem> malformedProblems = new ArrayList<>();
    validateSteps(malformedSteps, malformedProblems);
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
    validateSteps(validSteps, stepProblems);
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
                "Missing required field 'steps[4].action.zoomPercent'")
            .stream()
            .sorted()
            .toList(),
        messages.stream().sorted().toList());
  }

  @Test
  void buildsCompletePlansOnlyFromCompleteAndConstructorValidFragments() {
    RequestBoundRoot completeRoot =
        new RequestBoundRoot(
            Optional.of(ProtocolVersion.CURRENT),
            Optional.of(new PlanId("plan")),
            Optional.of(new WorkbookPlan.WorkbookSource.New()),
            Optional.of(new WorkbookPlan.WorkbookPersistence.None()),
            Optional.of(ExecutionPolicySpec.defaults()),
            Optional.of(FormulaEnvironmentSpec.empty()));
    RequestBoundFragments complete =
        new RequestBoundFragments(completeRoot, Optional.of(List.of()));

    assertTrue(complete.completePlan().isPresent());
    assertEquals("plan", complete.planId().orElseThrow().value());
    assertEquals(ProtocolVersion.CURRENT, complete.protocolVersion().orElseThrow());
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
                    Optional.of(ExecutionPolicySpec.defaults()),
                    Optional.of(FormulaEnvironmentSpec.empty())),
                Optional.of(List.of()))
            .completePlan()
            .isPresent());
    List<RequestBoundFragments> incompleteFragments =
        List.of(
            new RequestBoundFragments(
                new RequestBoundRoot(
                    Optional.of(ProtocolVersion.CURRENT),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.of(new WorkbookPlan.WorkbookPersistence.None()),
                    Optional.of(ExecutionPolicySpec.defaults()),
                    Optional.of(FormulaEnvironmentSpec.empty())),
                Optional.of(List.of())),
            new RequestBoundFragments(
                new RequestBoundRoot(
                    Optional.of(ProtocolVersion.CURRENT),
                    Optional.empty(),
                    Optional.of(new WorkbookPlan.WorkbookSource.New()),
                    Optional.empty(),
                    Optional.of(ExecutionPolicySpec.defaults()),
                    Optional.of(FormulaEnvironmentSpec.empty())),
                Optional.of(List.of())),
            new RequestBoundFragments(
                new RequestBoundRoot(
                    Optional.of(ProtocolVersion.CURRENT),
                    Optional.empty(),
                    Optional.of(new WorkbookPlan.WorkbookSource.New()),
                    Optional.of(new WorkbookPlan.WorkbookPersistence.None()),
                    Optional.empty(),
                    Optional.of(FormulaEnvironmentSpec.empty())),
                Optional.of(List.of())),
            new RequestBoundFragments(
                new RequestBoundRoot(
                    Optional.of(ProtocolVersion.CURRENT),
                    Optional.empty(),
                    Optional.of(new WorkbookPlan.WorkbookSource.New()),
                    Optional.of(new WorkbookPlan.WorkbookPersistence.None()),
                    Optional.of(ExecutionPolicySpec.defaults()),
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

    RequestIntake invalid = manualIntake(complete, List.of(new RequestUnknownField("extra", 0)));
    assertFalse(invalid.isStructurallyValid());
    assertTrue(invalid.completePlan().isEmpty());
    assertTrue(manualIntake(complete, List.of()).isStructurallyValid());
    assertEquals(complete.completePlan(), manualIntake(complete, List.of()).completePlan());

    var duplicateStep =
        GridGrindJson.readRequest(
                """
                {
                  "protocolVersion": 1,
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
    assertTrue(manualIntake(constructorInvalid, List.of()).completePlan().isEmpty());
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
        InvalidEncodingException.class, new RequestInvalidEncoding("bad encoding", 0).exception());
    assertInstanceOf(
        InvalidJsonException.class,
        new RequestInvalidJson("bad JSON", Optional.empty(), Optional.empty()).exception());
    assertInstanceOf(
        InvalidJsonException.class, new RequestDuplicateKey("", "duplicate", 0, 0).exception());
    for (RequestStructuralProblem problem :
        List.of(
            new RequestUnknownField("unknown", 0),
            new RequestMissingRequiredField("missing"),
            new RequestExplicitNullField("nullValue", 0),
            new RequestMissingTypeDiscriminator("type"),
            new RequestUnknownTypeDiscriminator("type", "UNKNOWN", 0),
            new RequestMalformedScalar("type", "a JSON string type id", 0),
            new RequestMalformedScalar("number", "a JSON number", 0))) {
      assertInstanceOf(InvalidRequestShapeException.class, problem.exception());
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
        node, type, jsonPath, structuralProblems, new ArrayList<>());
  }

  private static <T> Optional<T> bindMember(
      Optional<RequestJsonMember> member,
      Class<T> type,
      String jsonPath,
      List<RequestStructuralProblem> structuralProblems) {
    return RequestFragmentBinder.bindMember(
        member, type, jsonPath, structuralProblems, new ArrayList<>());
  }

  private static Optional<List<RequestBoundFragments.Step>> bindSteps(
      Optional<RequestJsonMember> stepsMember, List<RequestStructuralProblem> structuralProblems) {
    return RequestStepFragmentCollector.bind(stepsMember, structuralProblems, new ArrayList<>());
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

  private static void validateSteps(
      RequestJsonMember steps, List<RequestStructuralProblem> problems) {
    RequestSchemaStructure.validate(
        object(
            member("protocolVersion", new RequestJsonNumber(0, "1")),
            member("source", object(member("type", new RequestJsonString(0, "NEW")))),
            member("persistence", object(member("type", new RequestJsonString(0, "NONE")))),
            steps),
        problems);
  }

  private static RequestIntake manualIntake(
      dev.erst.gridgrind.core.problem.RequestBoundFragments fragments,
      java.util.List<dev.erst.gridgrind.core.problem.RequestStructuralProblem> problems) {
    return RequestIntakeBinder.bind(
        fragments, problems, java.util.List.of(), java.util.Optional.empty(), java.util.List.of());
  }
}
