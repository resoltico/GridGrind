package dev.erst.gridgrind.json.codec;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.grant.ExecutionGrant;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.InvalidAuthorityException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.function.Executable;
import tools.jackson.databind.json.JsonMapper;

/** Public generic mappers cannot create complete requests or authority outside owned intake. */
class DocumentReadGuardTest {
  private static final String REQUEST =
      """
      {"protocolVersion":1,"source":{"type":"NEW"},"persistence":{"type":"NONE"},"steps":[]}
      """;

  @Test
  void genericAndRebuiltMappersRejectCompleteRequestDocumentsIncludingDuplicatesAndNulls() {
    String duplicate =
        REQUEST.replace("\"protocolVersion\":1", "\"protocolVersion\":1,\"protocolVersion\":1");
    var mappers = publicMappersAndRebuilds();
    assertFalse(mappers.isEmpty(), "The generic and rebuilt mapper population must be exercised");
    for (JsonMapper mapper : mappers) {
      for (String document : List.of(REQUEST, duplicate, "null")) {
        requireReaderMisuse(
            () -> mapper.readValue(document, WorkbookPlan.class), "GridGrindJson.readRequest(...)");
      }
      requireReaderMisuse(
          () -> mapper.treeToValue(mapper.readTree(REQUEST), WorkbookPlan.class),
          "GridGrindJson.readRequest(...)");
      requireReaderMisuse(
          () -> mapper.readerFor(WorkbookPlan.class).readValue(REQUEST),
          "GridGrindJson.readRequest(...)");
    }
  }

  @Test
  void rawOversizeRequestCannotBypassTheBoundedCanonicalReader() {
    String document = REQUEST.stripTrailing();
    String oversized =
        document.substring(0, document.length() - 1)
            + ",\"planId\":\""
            + "x".repeat((int) GridGrindJson.maxRequestDocumentBytes())
            + "\"}";
    requireReaderMisuse(
        () -> GridGrindJsonMapperSupport.JSON_MAPPER.readValue(oversized, WorkbookPlan.class),
        "GridGrindJson.readRequest(...)");
    assertThrows(IllegalArgumentException.class, () -> GridGrindJson.readRequest(oversized));
  }

  @Test
  void genericAndRebuiltMappersRejectAuthorityDocumentsBeforeDuplicateOrPathBinding()
      throws Exception {
    String valid =
        new String(
            GridGrindJsonOutput.writeGrantBytes(ExecutionGrant.denied(), false),
            StandardCharsets.UTF_8);
    assertTrue(valid.contains("\"maxInputSnapshotBytes\":0"));
    String duplicate =
        valid.replace(
            "\"maxInputSnapshotBytes\":0",
            "\"maxInputSnapshotBytes\":0,\"maxInputSnapshotBytes\":0");
    for (JsonMapper mapper : publicMappersAndRebuilds()) {
      for (String document : List.of(valid, duplicate, "null")) {
        requireReaderMisuse(
            () -> mapper.readValue(document, ExecutionGrant.class), "GridGrindJson.readGrant(...)");
      }
      requireReaderMisuse(
          () -> mapper.treeToValue(mapper.readTree(valid), ExecutionGrant.class),
          "GridGrindJson.readGrant(...)");
    }
    assertEquals(
        ExecutionGrant.denied(),
        GridGrindJson.readGrant(
            valid.getBytes(StandardCharsets.UTF_8), Path.of("/synthetic/root")));
    assertThrows(
        InvalidAuthorityException.class,
        () ->
            GridGrindJson.readGrant(
                duplicate.getBytes(StandardCharsets.UTF_8), Path.of("/synthetic/root")));
  }

  @Test
  void supportedRequestReaderAndGenericTreeAndWriteOperationsRemainAvailable() throws Exception {
    WorkbookPlan plan = GridGrindJson.readRequest(REQUEST);
    assertEquals(plan, GridGrindJson.readRequest(GridGrindJsonOutput.writeRequestBytes(plan)));
    for (JsonMapper mapper : publicMappersAndRebuilds()) {
      assertTrue(mapper.readTree(REQUEST).isObject());
      assertEquals(plan, GridGrindJson.readRequest(mapper.writeValueAsBytes(plan)));
    }
    String duplicate =
        REQUEST.replace("\"protocolVersion\":1", "\"protocolVersion\":1,\"protocolVersion\":1");
    assertThrows(IllegalArgumentException.class, () -> GridGrindJson.readRequest(duplicate));
  }

  private static List<JsonMapper> publicMappersAndRebuilds() {
    return List.of(
        GridGrindJsonMapperSupport.JSON_MAPPER,
        GridGrindJsonMapperSupport.WIRE_JSON_MAPPER,
        GridGrindJsonMapperSupport.PRETTY_WIRE_JSON_MAPPER,
        GridGrindJsonMapperSupport.JSON_MAPPER.rebuild().build(),
        GridGrindJsonMapperSupport.WIRE_JSON_MAPPER.rebuild().build(),
        GridGrindJsonMapperSupport.PRETTY_WIRE_JSON_MAPPER.rebuild().build());
  }

  private static void requireReaderMisuse(Executable action, String reader) {
    var failure = assertThrows(IllegalStateException.class, action);
    assertTrue(failure.getMessage().contains(reader), failure::getMessage);
  }
}
