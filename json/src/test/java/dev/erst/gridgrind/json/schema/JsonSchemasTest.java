package dev.erst.gridgrind.json.schema;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.networknt.schema.InputFormat;
import com.networknt.schema.Schema;
import com.networknt.schema.SchemaRegistry;
import com.networknt.schema.SpecificationVersion;
import dev.erst.gridgrind.core.grant.AcceptancePolicy;
import dev.erst.gridgrind.core.grant.ExecutionGrant;
import dev.erst.gridgrind.core.model.pivot.PivotTableSpec;
import dev.erst.gridgrind.core.model.value.CellReport;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.plan.SheetInspectionResult;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.result.WorkbookResults;
import dev.erst.gridgrind.core.source.SecretReference;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.json.codec.GridGrindJsonMapperSupport;
import dev.erst.gridgrind.json.codec.GridGrindJsonOutput;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;
import tools.jackson.databind.node.ObjectNode;

/** Checks the emitted documents with an independent Draft 2020-12 validator. */
class JsonSchemasTest {
  private static final Path ROOT = Path.of(System.getProperty("gridgrind.root.dir"));
  private static final SchemaRegistry REGISTRY =
      SchemaRegistry.withDefaultDialect(SpecificationVersion.DRAFT_2020_12);

  @Test
  void everyShippedExampleValidatesIndependently() throws IOException {
    Schema request = schema(JsonSchemas.Document.REQUEST);
    try (var paths = Files.list(ROOT.resolve("examples"))) {
      List<Path> examples =
          paths.filter(path -> path.toString().endsWith(".json")).sorted().toList();
      assertEquals(14, examples.size());
      for (Path example : examples) {
        var errors = request.validate(Files.readString(example), InputFormat.JSON);
        assertTrue(errors.isEmpty(), () -> example + ": " + errors);
      }
    }
  }

  @Test
  void theIndependentValidatorRejectsWrongScalarsAndUnknownFields() {
    Schema request = schema(JsonSchemas.Document.REQUEST);
    String valid =
        """
        {"protocolVersion":1,"source":{"type":"NEW"},"persistence":{"type":"NONE"},"steps":[{"stepId":"set","target":{"type":"CELL_BY_ADDRESS","sheetName":"Sheet1","address":"A1"},"action":{"type":"SET_CELL","value":{"type":"NUMBER","value":42}}}]}
        """;
    assertTrue(request.validate(valid, InputFormat.JSON).isEmpty());
    assertFalse(
        request
            .validate(valid.replace("\"value\":42", "\"value\":true"), InputFormat.JSON)
            .isEmpty());
    assertFalse(
        request
            .validate(
                valid.replace("\"protocolVersion\":1", "\"protocolVersion\":\"1\""),
                InputFormat.JSON)
            .isEmpty());
    assertFalse(
        request
            .validate(valid.replace("\"steps\":", "\"extra\":0,\"steps\":"), InputFormat.JSON)
            .isEmpty());
  }

  @Test
  void grantSchemaAcceptsCanonicalAuthoritiesAndOpaqueSecretReferences() throws IOException {
    var grant =
        new ExecutionGrant(
            List.of(new ExecutionGrant.ReadAuthority.File(ROOT.resolve("input.xlsx"))),
            List.of("ENSURE_SHEET"),
            new ExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
            new ExecutionGrant.PublicationAuthority.SaveAs(
                ROOT.resolve("out.xlsx"), WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
            List.of(new SecretReference("report-password")),
            new AcceptancePolicy.RequireAll(
                List.of(new AcceptancePolicy.Requirement.RequireCalculation())));
    byte[] bytes = GridGrindJsonOutput.writeGrantBytes(grant, false);
    Schema schema = schema(JsonSchemas.Document.GRANT);
    var errors = schema.validate(new String(bytes, StandardCharsets.UTF_8), InputFormat.JSON);
    assertTrue(errors.isEmpty(), errors::toString);
    ObjectNode invalid =
        (ObjectNode) GridGrindJsonMapperSupport.JSON_MAPPER.readTree(bytes).deepCopy();
    ((ObjectNode) invalid.path("readableResources").get(0)).put("path", true);
    assertFalse(schema.validate(invalid.toString(), InputFormat.JSON).isEmpty());
    ObjectNode invalidReference =
        (ObjectNode) GridGrindJsonMapperSupport.JSON_MAPPER.readTree(bytes).deepCopy();
    ((ObjectNode) invalidReference.path("allowedSecretReferences").get(0)).put("id", "bad value");
    assertFalse(schema.validate(invalidReference.toString(), InputFormat.JSON).isEmpty());
  }

  @ParameterizedTest
  @ValueSource(doubles = {42.0, Double.NaN, Double.POSITIVE_INFINITY, Double.NEGATIVE_INFINITY})
  void responseSchemaPreservesFactualNumericEncodings(double value) throws IOException {
    var report =
        new CellReport.NumberReport(
            new CellAddress("A1"),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.of(new CellValue.NumberValue(value)),
            Optional.empty());
    var response =
        WorkbookResults.success(
            List.of(),
            List.of(),
            List.of(
                new SheetInspectionResult.CellsResult(
                    new StepId("observed"), new SheetLabel("Sheet1"), List.of(report))));
    byte[] bytes = GridGrindJsonOutput.writeWorkbookResultBytes(response);
    Schema schema = schema(JsonSchemas.Document.RESPONSE);
    var errors = schema.validate(new String(bytes, StandardCharsets.UTF_8), InputFormat.JSON);
    assertTrue(errors.isEmpty(), errors::toString);
    ObjectNode invalid =
        (ObjectNode) GridGrindJsonMapperSupport.JSON_MAPPER.readTree(bytes).deepCopy();
    ((ObjectNode) invalid.path("inspections").get(0).path("cells").get(0).path("value"))
        .put("value", "not-a-number");
    assertFalse(schema.validate(invalid.toString(), InputFormat.JSON).isEmpty());
  }

  @Test
  void defaultsAndSiblingDependenciesComeFromTheModelMetadata() {
    var request =
        GridGrindJsonMapperSupport.JSON_MAPPER.readTree(
            JsonSchemas.document(JsonSchemas.Document.REQUEST));
    var plan = request.path("$defs").path(WorkbookPlan.class.getName());
    assertEquals(
        "FULL_XSSF",
        plan.path("properties")
            .path("execution")
            .path("default")
            .path("mode")
            .path("type")
            .asString());
    assertEquals(
        "FAIL_FAST",
        plan.path("properties").path("execution").path("default").path("assertionMode").asString());
    var displayName =
        request
            .path("$defs")
            .path(PivotTableSpec.DataField.class.getName())
            .path("properties")
            .path("displayName");
    assertEquals("sourceColumnName", displayName.path("x-gridgrind-defaultFrom").asString());
    assertFalse(displayName.has("default"));
  }

  @Test
  void secretReferenceCatalogueShapeMatchesTheActualWireObject() {
    var security =
        new dev.erst.gridgrind.core.model.security.OoxmlOpenSecuritySpec(
            Optional.of(new SecretReference("package-password")));
    var wire = GridGrindJsonMapperSupport.JSON_MAPPER.valueToTree(security);
    assertEquals(
        GridGrindJsonMapperSupport.JSON_MAPPER.readTree(
            "{\"passwordRef\":{\"id\":\"package-password\"}}"),
        wire);
    var component =
        dev.erst.gridgrind.core.model.security.OoxmlOpenSecuritySpec.class.getRecordComponents()[0];
    var field =
        dev.erst.gridgrind.core.catalog.CatalogFieldMetadataSupport.fieldEntry(
            component, java.util.Set.of("passwordRef"));
    assertEquals(
        new dev.erst.gridgrind.core.catalog.FieldShape.PlainTypeGroupRef("secretReferenceType"),
        field.shape());
    assertTrue(
        dev.erst.gridgrind.core.catalog.GridGrindProtocolCatalog.catalog().plainTypes().stream()
            .anyMatch(group -> "secretReferenceType".equals(group.group())));
  }

  @Test
  void committedSchemaArtifactsMatchTheirProducerWithoutRewritingThem() throws IOException {
    Path directory = ROOT.resolve("docs/reference/schema");
    List<String> expectedNames =
        Arrays.stream(JsonSchemas.Document.values())
            .map(JsonSchemas.Document::fileName)
            .sorted()
            .toList();
    try (var files = Files.list(directory)) {
      assertEquals(
          expectedNames,
          files
              .filter(Files::isRegularFile)
              .map(path -> path.getFileName().toString())
              .sorted()
              .toList());
    }
    for (JsonSchemas.Document document : JsonSchemas.Document.values()) {
      assertEquals(
          JsonSchemas.document(document),
          Files.readString(directory.resolve(document.fileName())),
          document.name());
    }
  }

  private static Schema schema(JsonSchemas.Document document) {
    return REGISTRY.getSchema(JsonSchemas.document(document), InputFormat.JSON);
  }
}
