package dev.erst.gridgrind.contract.json;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTimeoutPreemptively;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.time.Duration;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Direct contract coverage for tolerant request intake and independently bound fragments. */
class RequestAnalysisTest {
  @Test
  void retainsValidSiblingFragmentsWhileCollectingIndependentShapeProblems() {
    RequestAnalysis analysis =
        GridGrindJson.analyzeRequest(
            """
            {
              "protocolVersion": "V2",
              "planId": 7,
              "source": { "type": "NEW" },
              "persistence": null,
              "extra": true,
              "steps": []
            }
            """
                .getBytes(StandardCharsets.UTF_8));

    assertFalse(analysis.isStructurallyValid());
    assertTrue(analysis.boundFragments().protocolVersion().isPresent());
    assertTrue(analysis.boundFragments().source().isPresent());
    assertTrue(analysis.boundFragments().execution().isPresent());
    assertTrue(analysis.boundFragments().formulaEnvironment().isPresent());
    assertTrue(analysis.boundFragments().persistence().isEmpty());
    assertTrue(analysis.completePlan().isEmpty());
    assertEquals(
        List.of(
            RequestMalformedScalar.class,
            RequestExplicitNullField.class,
            RequestUnknownField.class),
        analysis.structuralProblems().stream().map(Object::getClass).toList());
    assertEquals(
        List.of("planId", "persistence", "extra"),
        analysis.structuralProblems().stream()
            .map(problem -> problem.jsonPath().orElse(""))
            .toList());
  }

  @Test
  void retainsEveryConstructorInvalidFragmentAsAnExplicitBindingFailure() {
    String request =
        """
        {
          "protocolVersion": "V2",
          "source": { "type": "EXISTING", "path": "input.xls" },
          "persistence": { "type": "SAVE_AS", "path": "output.txt", "ifExists": "REJECT" },
          "steps": []
        }
        """;

    RequestAnalysis analysis =
        GridGrindJson.analyzeRequest(request.getBytes(StandardCharsets.UTF_8));

    assertTrue(analysis.isStructurallyValid());
    assertFalse(analysis.isBindable());
    assertEquals(List.of(), analysis.structuralProblems());
    assertEquals(
        List.of("source.path", "persistence.path"),
        analysis.bindingFailures().stream().map(RequestBindingFailure::jsonPath).toList());
    assertEquals(
        List.of(
            requestOffset(request, "\"path\": \"input.xls\""),
            requestOffset(request, "\"path\": \"output.txt\"")),
        analysis.bindingFailures().stream()
            .map(RequestBindingFailure::byteOffset)
            .map(Optional::orElseThrow)
            .toList());
    assertTrue(analysis.boundFragments().source().isEmpty());
    assertTrue(analysis.boundFragments().persistence().isEmpty());
    assertTrue(analysis.completePlan().isEmpty());
    InvalidRequestException failure =
        assertThrows(InvalidRequestException.class, analysis::requireCompletePlan);
    assertEquals(Optional.of("source.path"), failure.jsonPath());
  }

  @Test
  void excludesASyntacticallyMalformedBranchWhileRetainingItsValidSiblings() {
    RequestAnalysis analysis =
        GridGrindJson.analyzeRequest(
            """
            {
              "protocolVersion": "V2",
              "source" { "type": "NEW" },
              "persistence": { "type": "NONE" },
              "steps": []
            }
            """
                .getBytes(StandardCharsets.UTF_8));

    RequestInvalidJson syntaxProblem =
        assertInstanceOf(RequestInvalidJson.class, analysis.structuralProblems().getFirst());
    assertEquals("source", syntaxProblem.jsonPath().orElseThrow());
    assertTrue(analysis.boundFragments().protocolVersion().isPresent());
    assertTrue(analysis.boundFragments().source().isEmpty());
    assertTrue(analysis.boundFragments().persistence().isPresent());
    assertTrue(analysis.boundFragments().execution().isPresent());
    assertTrue(analysis.boundFragments().formulaEnvironment().isPresent());
    assertTrue(analysis.completePlan().isEmpty());
    assertThrows(IllegalStateException.class, analysis::requireCompletePlan);
  }

  @Test
  void distinguishesNonObjectRequestsFromSyntaxFailuresAndRetainsTheManualAnalysisInvariant() {
    RequestAnalysis scalarAnalysis =
        GridGrindJson.analyzeRequest("[]".getBytes(StandardCharsets.UTF_8));

    assertInstanceOf(RequestMalformedScalar.class, scalarAnalysis.structuralProblems().getFirst());

    RequestAnalysis decodedAnalysis =
        GridGrindJson.analyzeRequest(
            """
            {
              "protocolVersion": "V2",
              "source": { "type": "NEW" },
              "persistence": { "type": "NONE" },
              "steps": []
            }
            """
                .getBytes(StandardCharsets.UTF_8));
    assertEquals(
        decodedAnalysis.completePlan().orElseThrow(),
        new RequestAnalysis(decodedAnalysis.boundFragments(), List.of()).requireCompletePlan());
    assertTrue(decodedAnalysis.isBindable());

    RequestAnalysis incompleteManualAnalysis =
        new RequestAnalysis(
            new RequestBoundFragments(
                new RequestBoundRoot(
                    java.util.Optional.empty(),
                    java.util.Optional.empty(),
                    java.util.Optional.empty(),
                    java.util.Optional.empty(),
                    java.util.Optional.empty(),
                    java.util.Optional.empty()),
                java.util.Optional.of(List.of())),
            List.of());
    assertFalse(incompleteManualAnalysis.isBindable());
    assertThrows(IllegalStateException.class, incompleteManualAnalysis::requireCompletePlan);
  }

  @Test
  void preservesPreexistingNonJsonFindingsWhenTheRootIsNotAnObject() {
    RequestAnalysis analysis =
        RequestRootStructuralAnalyzer.analyze(
            new RequestSyntaxParseResult(
                new RequestJsonString(0, "not-an-object"),
                List.of(new RequestUnknownField("unrelated", 0))));

    assertEquals(
        List.of(RequestUnknownField.class, RequestMalformedScalar.class),
        analysis.structuralProblems().stream().map(Object::getClass).toList());
  }

  @Test
  void anchorsDuplicateKeysAtTheirSecondPropertyTokenInUtf8Bytes() {
    String request =
        """
        {"protocolVersion":"V2","source":{"type":"NEW"},"persistence":{"type":"NONE"},"steps":[],"planId":"€","planId":"duplicate"}
        """;
    byte[] bytes = request.getBytes(StandardCharsets.UTF_8);

    RequestAnalysis analysis = GridGrindJson.analyzeRequest(bytes);

    RequestDuplicateKey duplicate =
        assertInstanceOf(
            RequestDuplicateKey.class,
            analysis.structuralProblems().stream()
                .filter(RequestDuplicateKey.class::isInstance)
                .findFirst()
                .orElseThrow());
    assertEquals("", duplicate.containingObjectPath());
    assertEquals("planId", duplicate.key());
    assertEquals(0, duplicate.occurrenceOrdinal());
    assertEquals(secondOccurrence(bytes, "\"planId\""), duplicate.byteOffset().orElseThrow());
  }

  @Test
  void distinguishesEveryDuplicateOccurrenceOfTheSameProperty() {
    String request =
        """
        {"protocolVersion":"V2","source":{"type":"NEW"},"persistence":{"type":"NONE"},"steps":[],"planId":"one","planId":"two","planId":"three"}
        """;
    byte[] bytes = request.getBytes(StandardCharsets.UTF_8);

    List<RequestDuplicateKey> duplicates =
        GridGrindJson.analyzeRequest(bytes).structuralProblems().stream()
            .filter(RequestDuplicateKey.class::isInstance)
            .map(RequestDuplicateKey.class::cast)
            .toList();

    assertEquals(2, duplicates.size());
    assertEquals(
        List.of(0, 1), duplicates.stream().map(RequestDuplicateKey::occurrenceOrdinal).toList());
    assertEquals(
        List.of(secondOccurrence(bytes, "\"planId\""), thirdOccurrence(bytes, "\"planId\"")),
        duplicates.stream().map(duplicate -> duplicate.byteOffset().orElseThrow()).toList());
  }

  @Test
  void validatesEveryKnownDuplicateValueAlongsideItsDuplicateKeyFinding() {
    String request =
        """
        {
          "protocolVersion": "V2",
          "protocolVersion": null,
          "planId": "valid",
          "planId": 7,
          "source": { "type": "EXISTING", "type": null, "path": "source.xlsx", "path": 7 },
          "persistence": { "type": "NONE" },
          "steps": [],
          "steps": null
        }
        """;

    RequestAnalysis analysis =
        GridGrindJson.analyzeRequest(request.getBytes(StandardCharsets.UTF_8));

    assertEquals(
        List.of(
            RequestDuplicateKey.class,
            RequestExplicitNullField.class,
            RequestDuplicateKey.class,
            RequestMalformedScalar.class,
            RequestDuplicateKey.class,
            RequestExplicitNullField.class,
            RequestDuplicateKey.class,
            RequestMalformedScalar.class,
            RequestDuplicateKey.class,
            RequestExplicitNullField.class),
        analysis.structuralProblems().stream().map(Object::getClass).toList());
    List<RequestStructuralProblem> memberProblems =
        analysis.structuralProblems().stream()
            .filter(problem -> !(problem instanceof RequestDuplicateKey))
            .toList();
    assertEquals(
        List.of("protocolVersion", "planId", "source.type", "source.path", "steps"),
        memberProblems.stream().map(problem -> problem.jsonPath().orElseThrow()).toList());
    assertEquals(
        List.of(
            requestOffset(request, "\"protocolVersion\": null"),
            requestOffset(request, "\"planId\": 7"),
            requestOffset(request, "\"type\": null"),
            requestOffset(request, "\"path\": 7"),
            requestOffset(request, "\"steps\": null")),
        memberProblems.stream().map(problem -> problem.byteOffset().orElseThrow()).toList());
    assertTrue(analysis.boundFragments().protocolVersion().isEmpty());
    assertTrue(analysis.boundFragments().planId().isEmpty());
    assertTrue(analysis.boundFragments().source().isEmpty());
    assertTrue(analysis.boundFragments().steps().isEmpty());
  }

  @Test
  void validatesRepeatedRootAndStepMembersBeforeBindingTheirFirstOccurrence() {
    String request =
        """
        {
          "protocolVersion": "V2",
          "source": { "type": "NEW" },
          "source": null,
          "persistence": { "type": "NONE" },
          "steps": [
            {
              "stepId": "valid",
              "stepId": null,
              "target": { "type": "WORKBOOK_CURRENT" },
              "query": { "type": "GET_WORKBOOK_SUMMARY", "type": null }
            }
          ]
        }
        """;

    RequestAnalysis analysis =
        GridGrindJson.analyzeRequest(request.getBytes(StandardCharsets.UTF_8));

    assertEquals(
        List.of(
            RequestDuplicateKey.class,
            RequestExplicitNullField.class,
            RequestDuplicateKey.class,
            RequestExplicitNullField.class,
            RequestDuplicateKey.class,
            RequestExplicitNullField.class),
        analysis.structuralProblems().stream().map(Object::getClass).toList());
    assertEquals(
        List.of("source", "steps[0].stepId", "steps[0].query.type"),
        analysis.structuralProblems().stream()
            .filter(problem -> !(problem instanceof RequestDuplicateKey))
            .map(problem -> problem.jsonPath().orElseThrow())
            .toList());
    assertTrue(analysis.boundFragments().source().isEmpty());
    assertTrue(analysis.boundFragments().steps().orElseThrow().getFirst().value().isEmpty());
    assertTrue(analysis.boundFragments().persistence().isPresent());
  }

  @Test
  void excludesEveryAmbiguousRootFragmentWhileRetainingUnrelatedSiblings() {
    RequestAnalysis analysis =
        GridGrindJson.analyzeRequest(
            """
            {
              "protocolVersion": "V2",
              "protocolVersion": "V2",
              "planId": "first",
              "planId": "second",
              "source": { "type": "NEW" },
              "source": { "type": "EXISTING", "path": "source.xlsx" },
              "persistence": { "type": "NONE" },
              "steps": [],
              "steps": []
            }
            """
                .getBytes(StandardCharsets.UTF_8));

    assertEquals(
        List.of(
            RequestDuplicateKey.class,
            RequestDuplicateKey.class,
            RequestDuplicateKey.class,
            RequestDuplicateKey.class),
        analysis.structuralProblems().stream().map(Object::getClass).toList());
    assertTrue(analysis.boundFragments().protocolVersion().isEmpty());
    assertTrue(analysis.boundFragments().planId().isEmpty());
    assertTrue(analysis.boundFragments().source().isEmpty());
    assertTrue(analysis.boundFragments().steps().isEmpty());
    assertTrue(analysis.boundFragments().persistence().isPresent());
    assertTrue(analysis.boundFragments().execution().isPresent());
    assertTrue(analysis.boundFragments().formulaEnvironment().isPresent());
    assertTrue(analysis.completePlan().isEmpty());
  }

  @Test
  void validatesEveryAuthoredPayloadFamilyWhenStepPayloadCardinalityIsInvalid() {
    RequestAnalysis analysis =
        GridGrindJson.analyzeRequest(
            """
            {
              "protocolVersion": "V2",
              "source": { "type": "NEW" },
              "persistence": { "type": "NONE" },
              "steps": [
                {
                  "stepId": "multiple-payloads",
                  "target": { "type": "WORKBOOK_CURRENT" },
                  "action": null,
                  "query": { "type": null }
                }
              ]
            }
            """
                .getBytes(StandardCharsets.UTF_8));

    assertEquals(
        List.of("steps[0]", "steps[0].action", "steps[0].query.type"),
        analysis.structuralProblems().stream()
            .map(problem -> problem.jsonPath().orElseThrow())
            .toList());
    assertEquals(
        List.of(
            RequestMalformedScalar.class,
            RequestExplicitNullField.class,
            RequestExplicitNullField.class),
        analysis.structuralProblems().stream().map(Object::getClass).toList());
    assertTrue(analysis.boundFragments().steps().orElseThrow().getFirst().value().isEmpty());
  }

  @Test
  void rejectsNonUtf8WithoutTryingToDecodeOrBindIt() {
    byte[] bytes = new byte[] {'{', '"', 'x', '"', ':', (byte) 0xC3, (byte) 0x28};
    RequestAnalysis analysis = GridGrindJson.analyzeRequest(bytes);

    assertEquals(1, analysis.structuralProblems().size());
    RequestInvalidEncoding invalidEncoding =
        assertInstanceOf(RequestInvalidEncoding.class, analysis.structuralProblems().getFirst());
    assertEquals(5L, invalidEncoding.byteOffset().orElseThrow());
    assertTrue(analysis.boundFragments().source().isEmpty());
  }

  @Test
  void terminatesAndReportsMalformedNestedContainersWithoutConsumingTheEnclosingCloser() {
    RequestAnalysis analysis =
        assertTimeoutPreemptively(
            Duration.ofSeconds(1),
            () ->
                GridGrindJson.analyzeRequest(
                    """
                    {
                      "protocolVersion": "V2",
                      "source": { "type": "NEW" },
                      "persistence": { "type": "NONE" },
                      "steps": [}
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));

    assertFalse(analysis.isStructurallyValid());
    assertTrue(
        analysis.structuralProblems().stream().anyMatch(RequestInvalidJson.class::isInstance));
    assertTrue(analysis.boundFragments().source().isPresent());
    assertTrue(analysis.boundFragments().persistence().isPresent());
  }

  @Test
  void rejectsNonJsonWhitespaceRatherThanNormalizingItIntoAValidRequest() {
    String request =
        """
        {
          "protocolVersion": "V2",
          "source": { "type": "NEW" },<non-json-whitespace>
          "persistence": { "type": "NONE" },
          "steps": []
        }
        """
            .replace("<non-json-whitespace>", Character.toString(0x0B));
    RequestAnalysis analysis =
        GridGrindJson.analyzeRequest(request.getBytes(StandardCharsets.UTF_8));

    assertFalse(analysis.isStructurallyValid());
    assertTrue(
        analysis.structuralProblems().stream().anyMatch(RequestInvalidJson.class::isInstance));
  }

  @Test
  void distinguishesMissingFieldsFromExplicitNullWithoutInventingTokenOffsets() {
    RequestAnalysis analysis =
        GridGrindJson.analyzeRequest(
            """
            {
              "protocolVersion": "V2",
              "source": { "type": "NEW" },
              "persistence": { "type": "NONE" },
              "steps": null
            }
            """
                .getBytes(StandardCharsets.UTF_8));

    RequestExplicitNullField nullSteps =
        assertInstanceOf(
            RequestExplicitNullField.class,
            analysis.structuralProblems().stream()
                .filter(RequestExplicitNullField.class::isInstance)
                .findFirst()
                .orElseThrow());
    assertEquals("steps", nullSteps.jsonPath().orElseThrow());
    assertTrue(nullSteps.byteOffset().isPresent());
    assertEquals(
        requestOffset(
            """
            {
              "protocolVersion": "V2",
              "source": { "type": "NEW" },
              "persistence": { "type": "NONE" },
              "steps": null
            }
            """,
            "\"steps\""),
        nullSteps.byteOffset().orElseThrow());
  }

  @Test
  void classifiesNullRootEnumsAndUnionDiscriminatorsAsExplicitNulls() {
    String request =
        """
        {
          "protocolVersion": null,
          "source": { "type": null },
          "persistence": { "type": "NONE" },
          "steps": []
        }
        """;
    RequestAnalysis analysis =
        GridGrindJson.analyzeRequest(request.getBytes(StandardCharsets.UTF_8));

    List<RequestExplicitNullField> nullFields =
        analysis.structuralProblems().stream()
            .filter(RequestExplicitNullField.class::isInstance)
            .map(RequestExplicitNullField.class::cast)
            .toList();

    assertEquals(
        List.of("protocolVersion", "source.type"),
        nullFields.stream().map(field -> field.jsonPath().orElseThrow()).toList());
    assertEquals(
        List.of(
            "Field 'protocolVersion' must be omitted when absent; explicit null is not accepted.",
            "Field 'source.type' must be omitted when absent; explicit null is not accepted."),
        nullFields.stream().map(RequestExplicitNullField::message).toList());
    assertEquals(
        List.of(
            requestOffset(request, "\"protocolVersion\""),
            requestOffset(request, "\"type\": null")),
        nullFields.stream().map(field -> field.byteOffset().orElseThrow()).toList());
  }

  @Test
  void treatsMissingTypeIdsAsRequiredDiscriminatorsRatherThanRecordFields() {
    RequestAnalysis analysis =
        GridGrindJson.analyzeRequest(
            """
            {
              "protocolVersion": "V2",
              "source": {},
              "persistence": { "type": "NONE" },
              "steps": []
            }
            """
                .getBytes(StandardCharsets.UTF_8));

    RequestMissingTypeDiscriminator missingType =
        assertInstanceOf(
            RequestMissingTypeDiscriminator.class, analysis.structuralProblems().getFirst());
    assertEquals("source.type", missingType.jsonPath().orElseThrow());
    assertTrue(missingType.byteOffset().isEmpty());
  }

  @Test
  void anchorsObjectMemberFailuresAtPropertyNamesAndArrayFailuresAtValueTokens() {
    String request =
        """
        {
          "protocolVersion": 7,
          "source": { "type": 1 },
          "persistence": { "type": "NONE" },
          "steps": [7],
          "unexpected": true
        }
        """;
    byte[] bytes = request.getBytes(StandardCharsets.UTF_8);
    List<RequestStructuralProblem> problems =
        GridGrindJson.analyzeRequest(bytes).structuralProblems();

    assertEquals(
        firstOccurrence(bytes, "\"protocolVersion\": 7"),
        problemAt(problems, "protocolVersion").byteOffset().orElseThrow());
    assertEquals(
        firstOccurrence(bytes, "\"type\": 1"),
        problemAt(problems, "source.type").byteOffset().orElseThrow());
    assertEquals(
        firstOccurrence(bytes, "[7]") + 1,
        problemAt(problems, "steps[0]").byteOffset().orElseThrow());
    assertEquals(
        firstOccurrence(bytes, "\"unexpected\""),
        problemAt(problems, "unexpected").byteOffset().orElseThrow());
  }

  @Test
  void collectsIndependentUnionAndScalarFailuresWhileBindingValidSiblingSteps() {
    RequestAnalysis analysis =
        GridGrindJson.analyzeRequest(
            """
            {
              "protocolVersion": "V2",
              "source": { "type": "NEW", "unexpected": true },
              "persistence": { "type": "NOT_A_PERSISTENCE_MODE" },
              "steps": [
                {
                  "stepId": 7,
                  "target": {},
                  "query": { "type": 1 }
                },
                {
                  "stepId": "valid-sibling",
                  "target": { "type": "WORKBOOK_CURRENT" },
                  "query": { "type": "GET_WORKBOOK_SUMMARY" }
                }
              ]
            }
            """
                .getBytes(StandardCharsets.UTF_8));

    assertEquals(
        List.of(
            "source.unexpected",
            "persistence.type",
            "steps[0].stepId",
            "steps[0].query.type",
            "steps[0].target.type"),
        analysis.structuralProblems().stream()
            .map(problem -> problem.jsonPath().orElse(""))
            .toList());
    assertTrue(analysis.boundFragments().steps().orElseThrow().get(0).value().isEmpty());
    assertTrue(analysis.boundFragments().steps().orElseThrow().get(1).value().isPresent());
    assertTrue(analysis.completePlan().isEmpty());
  }

  @Test
  void collectsFieldsFromLaterValidUnionBranchesAfterTheFirstTypeIsMalformed() {
    RequestAnalysis analysis =
        GridGrindJson.analyzeRequest(
            """
            {
              "protocolVersion": "V2",
              "source": {
                "type": null,
                "type": "EXISTING",
                "path": 7,
                "notAllowed": true
              },
              "persistence": { "type": "NONE" },
              "steps": []
            }
            """
                .getBytes(StandardCharsets.UTF_8));

    assertEquals(
        List.of("source.type", "", "source.path", "source.notAllowed"),
        analysis.structuralProblems().stream()
            .map(problem -> problem.jsonPath().orElse(""))
            .toList());
    assertEquals(
        List.of(
            RequestExplicitNullField.class,
            RequestDuplicateKey.class,
            RequestMalformedScalar.class,
            RequestUnknownField.class),
        analysis.structuralProblems().stream().map(Object::getClass).toList());
    assertTrue(analysis.boundFragments().source().isEmpty());
  }

  @Test
  void collectsFieldsUnknownToEveryUnionVariantWhenNoTypeValueCanSelectABranch() {
    RequestAnalysis analysis =
        GridGrindJson.analyzeRequest(
            """
            {
              "protocolVersion": "V2",
              "source": { "type": null, "notAllowed": true },
              "persistence": { "type": "NONE" },
              "steps": []
            }
            """
                .getBytes(StandardCharsets.UTF_8));

    assertEquals(
        List.of("source.type", "source.notAllowed"),
        analysis.structuralProblems().stream()
            .map(problem -> problem.jsonPath().orElseThrow())
            .toList());
    assertEquals(
        List.of(RequestExplicitNullField.class, RequestUnknownField.class),
        analysis.structuralProblems().stream().map(Object::getClass).toList());
    assertTrue(analysis.boundFragments().source().isEmpty());
  }

  @Test
  void deDuplicatesSharedProblemsFromMultipleValidUnionBranches() {
    RequestAnalysis analysis =
        GridGrindJson.analyzeRequest(
            """
            {
              "protocolVersion": "V2",
              "source": { "type": "NEW" },
              "persistence": {
                "type": "OVERWRITE",
                "type": "SAVE_AS",
                "path": "result.xlsx",
                "ifExists": "REPLACE",
                "security": 7
              },
              "steps": []
            }
            """
                .getBytes(StandardCharsets.UTF_8));

    assertEquals(
        List.of(
            RequestDuplicateKey.class,
            RequestUnknownField.class,
            RequestUnknownField.class,
            RequestMalformedScalar.class),
        analysis.structuralProblems().stream().map(Object::getClass).toList());
    assertEquals(
        List.of("persistence.security"),
        analysis.structuralProblems().stream()
            .filter(
                problem -> problem.jsonPath().filter("persistence.security"::equals).isPresent())
            .map(problem -> problem.jsonPath().orElseThrow())
            .toList());
  }

  @Test
  void rejectsMissingAndNullPrimitiveScalarsAsShapeDefectsBeforeRecordBinding() {
    String missingZoom =
        """
        {
          "protocolVersion": "V2",
          "source": { "type": "NEW" },
          "persistence": { "type": "NONE" },
          "steps": [
            {
              "stepId": "zoom",
              "target": { "type": "SHEET_BY_NAME", "name": "Budget" },
              "action": { "type": "SET_SHEET_ZOOM" }
            }
          ]
        }
        """;
    String nullZoom =
        missingZoom.replace("\"SET_SHEET_ZOOM\"", "\"SET_SHEET_ZOOM\", \"zoomPercent\": null");

    InvalidRequestShapeException missing =
        assertThrows(
            InvalidRequestShapeException.class, () -> GridGrindJson.readRequest(missingZoom));
    InvalidRequestShapeException explicitNull =
        assertThrows(InvalidRequestShapeException.class, () -> GridGrindJson.readRequest(nullZoom));

    assertEquals("Missing required field 'steps[0].action.zoomPercent'", missing.getMessage());
    assertEquals(
        "Field 'steps[0].action.zoomPercent' must be omitted when absent; explicit null is not accepted.",
        explicitNull.getMessage());
  }

  @Test
  void reportsAnUnrepresentableNumericScalarAtItsAuthoredFieldBeforeTypedBinding() {
    String request =
        """
        {
          "protocolVersion": "V2",
          "source": { "type": "NEW" },
          "persistence": { "type": "NONE" },
          "steps": [
            {
              "stepId": "zoom",
              "target": { "type": "SHEET_BY_NAME", "name": "Budget" },
              "action": {
                "type": "SET_SHEET_ZOOM",
                "zoomPercent": 999999999999999999999
              }
            }
          ]
        }
        """;
    byte[] bytes = request.getBytes(StandardCharsets.UTF_8);

    RequestAnalysis analysis = GridGrindJson.analyzeRequest(bytes);

    RequestMalformedScalar problem =
        assertInstanceOf(
            RequestMalformedScalar.class,
            problemAt(analysis.structuralProblems(), "steps[0].action.zoomPercent"));
    assertEquals(
        "Field 'steps[0].action.zoomPercent' must be a JSON integer between -2147483648 and 2147483647",
        problem.message());
    assertEquals(
        firstOccurrence(bytes, "\"zoomPercent\": 999999999999999999999"),
        problem.byteOffset().orElseThrow());
    assertTrue(analysis.boundFragments().steps().orElseThrow().getFirst().value().isEmpty());
    assertTrue(analysis.completePlan().isEmpty());
  }

  @Test
  void ordersEveryStructuralProblemKindByPositionThenStableDiagnosticFacts() {
    List<RequestStructuralProblem> ordered =
        RequestStructuralProblemOrder.order(
            List.of(
                new RequestMissingRequiredField("steps[missing"),
                new RequestUnknownField("source.unexpected", 30),
                new RequestInvalidEncoding("invalid UTF-8", 10),
                new RequestInvalidJson(
                    "invalid JSON", java.util.Optional.empty(), java.util.Optional.of(11L)),
                new RequestDuplicateKey("", "planId", 1, 12),
                new RequestDuplicateKey("", "planId", 0, 12),
                new RequestExplicitNullField("steps[bad].stepId", 13),
                new RequestUnknownTypeDiscriminator("steps[3].type", "UNKNOWN", 14),
                new RequestMalformedScalar("steps[1].stepId", "a JSON string", 15),
                new RequestMissingTypeDiscriminator("steps[2].target.type")));

    assertEquals(
        List.of(
            RequestInvalidEncoding.class,
            RequestInvalidJson.class,
            RequestDuplicateKey.class,
            RequestDuplicateKey.class,
            RequestExplicitNullField.class,
            RequestUnknownTypeDiscriminator.class,
            RequestMalformedScalar.class,
            RequestUnknownField.class,
            RequestMissingRequiredField.class,
            RequestMissingTypeDiscriminator.class),
        ordered.stream().map(Object::getClass).toList());
    assertEquals(
        List.of(0, 1),
        ordered.stream()
            .filter(RequestDuplicateKey.class::isInstance)
            .map(RequestDuplicateKey.class::cast)
            .map(RequestDuplicateKey::occurrenceOrdinal)
            .toList());
  }

  @Test
  void ordersTiedFindingsByTheirCanonicalProblemCodeWithoutSkippingAnyVariant() {
    assertEquals(
        List.of(RequestInvalidEncoding.class, RequestInvalidJson.class),
        RequestStructuralProblemOrder.order(
                List.of(
                    new RequestInvalidJson(
                        "invalid JSON", java.util.Optional.empty(), java.util.Optional.of(0L)),
                    new RequestInvalidEncoding("invalid UTF-8", 0)))
            .stream()
            .map(Object::getClass)
            .toList());
    assertEquals(
        List.of(RequestInvalidJson.class, RequestDuplicateKey.class),
        RequestStructuralProblemOrder.order(
                List.of(
                    new RequestInvalidJson(
                        "invalid JSON", java.util.Optional.empty(), java.util.Optional.of(0L)),
                    new RequestDuplicateKey("", "planId", 0, 0)))
            .stream()
            .map(Object::getClass)
            .toList());
    assertEquals(
        List.of(
            RequestUnknownField.class,
            RequestExplicitNullField.class,
            RequestUnknownTypeDiscriminator.class,
            RequestMalformedScalar.class),
        RequestStructuralProblemOrder.order(
                List.of(
                    new RequestUnknownField("field", 0),
                    new RequestExplicitNullField("field", 0),
                    new RequestUnknownTypeDiscriminator("field", "UNKNOWN", 0),
                    new RequestMalformedScalar("field", "a JSON string", 0)))
            .stream()
            .map(Object::getClass)
            .toList());
    assertEquals(
        List.of(RequestMissingRequiredField.class, RequestMissingTypeDiscriminator.class),
        RequestStructuralProblemOrder.order(
                List.of(
                    new RequestMissingTypeDiscriminator("steps[2].type"),
                    new RequestMissingRequiredField("steps[bad].value")))
            .stream()
            .map(Object::getClass)
            .toList());
    assertProblemCodeTie(
        new RequestInvalidEncoding("invalid UTF-8", 0), new RequestUnknownField("field", 0));
    assertProblemCodeTie(
        new RequestInvalidEncoding("invalid UTF-8", 0), new RequestExplicitNullField("field", 0));
    assertProblemCodeTie(
        new RequestInvalidEncoding("invalid UTF-8", 0),
        new RequestUnknownTypeDiscriminator("field", "UNKNOWN", 0));
    assertProblemCodeTie(
        new RequestInvalidEncoding("invalid UTF-8", 0),
        new RequestMalformedScalar("field", "a JSON string", 0));
    assertProblemCodeTie(
        new RequestInvalidJson(
            "invalid JSON", java.util.Optional.empty(), java.util.Optional.empty()),
        new RequestMissingRequiredField("field"));
    assertProblemCodeTie(
        new RequestInvalidJson(
            "invalid JSON", java.util.Optional.empty(), java.util.Optional.empty()),
        new RequestMissingTypeDiscriminator("field"));
  }

  private static void assertProblemCodeTie(
      RequestStructuralProblem first, RequestStructuralProblem second) {
    assertEquals(2, RequestStructuralProblemOrder.order(List.of(first, second)).size());
  }

  private static long secondOccurrence(byte[] bytes, String text) {
    return occurrence(bytes, text, 2);
  }

  private static long thirdOccurrence(byte[] bytes, String text) {
    return occurrence(bytes, text, 3);
  }

  private static long firstOccurrence(byte[] bytes, String text) {
    return occurrence(bytes, text, 1);
  }

  private static long occurrence(byte[] bytes, String text, int requiredOccurrence) {
    byte[] needle = text.getBytes(StandardCharsets.UTF_8);
    int seen = 0;
    for (int index = 0; index <= bytes.length - needle.length; index++) {
      boolean matches = true;
      for (int offset = 0; offset < needle.length; offset++) {
        if (bytes[index + offset] != needle[offset]) {
          matches = false;
          break;
        }
      }
      if (matches) {
        seen++;
      }
      if (seen == requiredOccurrence) {
        return index;
      }
    }
    throw new AssertionError("Occurrence %d was not found".formatted(requiredOccurrence));
  }

  private static long requestOffset(String request, String text) {
    return firstOccurrence(request.getBytes(StandardCharsets.UTF_8), text);
  }

  private static RequestStructuralProblem problemAt(
      List<RequestStructuralProblem> problems, String jsonPath) {
    return problems.stream()
        .filter(problem -> problem.jsonPath().filter(jsonPath::equals).isPresent())
        .findFirst()
        .orElseThrow();
  }
}
