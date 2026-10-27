package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.customxml.CustomXmlDataBindingReport;
import dev.erst.gridgrind.core.model.customxml.CustomXmlExportReport;
import dev.erst.gridgrind.core.model.customxml.CustomXmlMappingLocator;
import dev.erst.gridgrind.core.model.customxml.CustomXmlMappingReport;
import dev.erst.gridgrind.core.model.customxml.ResolvedCustomXmlImport;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.engine.customxml.CustomXmlAdapter;
import java.io.IOException;
import java.util.List;
import java.util.Optional;
import javax.xml.transform.TransformerException;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTMap;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSchema;
import org.w3c.dom.Node;
import org.xml.sax.SAXException;

/** Tests for workbook custom-XML mapping inspection and import-export flows. */
class CustomXmlAdapterTest {
  @Test
  void readsWorkbookCustomXmlMappingsAndExportsXml() throws IOException {
    try (XssfWorkbook workbook = CustomXmlWorkbookSamples.openSimpleCustomXmlWorkbook()) {
      List<CustomXmlMappingReport> mappings = workbook.customXml().mappings();

      assertEquals(1, mappings.size());
      CustomXmlMappingReport mapping = mappings.getFirst();
      assertEquals(1L, mapping.mapId());
      assertEquals("CORSO_mapping", mapping.name());
      assertEquals("CORSO", mapping.rootElement());
      assertEquals("Schema1", mapping.schemaId());
      assertFalse(mapping.settings().showImportExportValidationErrors());
      assertTrue(mapping.settings().autoFit());
      assertFalse(mapping.settings().append());
      assertTrue(mapping.settings().preserveSortAfLayout());
      assertTrue(mapping.settings().preserveFormat());
      assertTrue(mapping.schema().xml().orElseThrow().contains("<xsd:element name=\"CORSO\""));
      assertEquals(8, mapping.linkedCells().size());
      assertEquals("Foglio1", mapping.linkedCells().getFirst().sheetName().value());
      assertEquals("A1", mapping.linkedCells().getFirst().address().value());
      assertEquals("/CORSO/NOME", mapping.linkedCells().getFirst().xpath());
      assertEquals("string", mapping.linkedCells().getFirst().xmlDataType());
      assertEquals(List.of(), mapping.linkedTables());

      CustomXmlExportReport exported =
          workbook
              .customXml()
              .exportMapping(
                  new CustomXmlMappingLocator(
                      java.util.Optional.of(1L), java.util.Optional.of("CORSO_mapping")),
                  true,
                  "UTF-8");

      assertEquals("UTF-8", exported.encoding());
      assertTrue(exported.schemaValidated());
      assertTrue(exported.xml().contains("<CORSO>"));
      assertTrue(exported.xml().contains("<NOME>ro</NOME>"));
      assertTrue(exported.xml().contains("<DOCENTE>ro</DOCENTE>"));
    }
  }

  @Test
  void importsXmlIntoExistingCustomXmlMappingAndSupportsReadCommands() throws IOException {
    try (XssfWorkbook workbook = CustomXmlWorkbookSamples.openSimpleCustomXmlWorkbook()) {
      workbook
          .customXml()
          .importMapping(
              new ResolvedCustomXmlImport(
                  new CustomXmlMappingLocator(
                      java.util.Optional.empty(), java.util.Optional.of("CORSO_mapping")),
                  "<CORSO><NOME>Grid</NOME><DOCENTE>Grind</DOCENTE><TUTOR>Agent</TUTOR>"
                      + "<CDL>Ops</CDL><DURATA>5</DURATA><ARGOMENTO>Audit</ARGOMENTO>"
                      + "<PROGETTO>Parity</PROGETTO><CREDITI>10</CREDITI></CORSO>"));

      dev.erst.gridgrind.core.model.layout.WindowReport.Dense window =
          workbook.sheet("Foglio1").cells().window("A1", 1, 3);
      assertEquals("Grid", window.rows().getFirst().cells().get(0).displayValue().orElseThrow());
      assertEquals("Grind", window.rows().getFirst().cells().get(1).displayValue().orElseThrow());
      assertEquals("Agent", window.rows().getFirst().cells().get(2).displayValue().orElseThrow());

      WorkbookIntrospector introspector = new WorkbookIntrospector();
      dev.erst.gridgrind.core.plan.WorkbookInspectionResult.CustomXmlMappingsResult mappings =
          assertInstanceOf(
              dev.erst.gridgrind.core.plan.WorkbookInspectionResult.CustomXmlMappingsResult.class,
              introspector.execute(
                  workbook,
                  new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                      new StepId("maps"),
                      new dev.erst.gridgrind.core.plan.WorkbookSelector.Current(),
                      new dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery
                          .GetCustomXmlMappings())));
      dev.erst.gridgrind.core.plan.WorkbookInspectionResult.CustomXmlExportResult exported =
          assertInstanceOf(
              dev.erst.gridgrind.core.plan.WorkbookInspectionResult.CustomXmlExportResult.class,
              introspector.execute(
                  workbook,
                  new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                      new StepId("export"),
                      new dev.erst.gridgrind.core.plan.WorkbookSelector.Current(),
                      new dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery
                          .ExportCustomXmlMapping(
                          new CustomXmlMappingLocator(
                              java.util.Optional.of(1L), java.util.Optional.empty()),
                          false,
                          "UTF-8"))));

      assertEquals("maps", mappings.stepId().value());
      assertEquals(1, mappings.mappings().size());
      assertEquals("export", exported.stepId().value());
      assertTrue(exported.export().xml().contains("<NOME>Grid</NOME>"));
    }
  }

  @Test
  void rejectsUnknownCustomXmlMappings() throws IOException {
    try (XssfWorkbook workbook = CustomXmlWorkbookSamples.openSimpleCustomXmlWorkbook()) {
      IllegalArgumentException failure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  workbook
                      .customXml()
                      .exportMapping(
                          new CustomXmlMappingLocator(
                              java.util.Optional.of(99L), java.util.Optional.of("missing")),
                          false,
                          "UTF-8"));
      assertTrue(failure.getMessage().contains("No custom XML mapping matched"));
    }
  }

  @Test
  void exportAndImportWrapExplicitPoiFailures() throws IOException {
    try (XssfWorkbook workbook = CustomXmlWorkbookSamples.openSimpleCustomXmlWorkbook()) {
      CustomXmlAdapter transformerFailingController =
          new CustomXmlAdapter(
              workbook.xssfWorkbook(),
              (mapping, output, encoding, validateSchema) -> {
                throw new TransformerException("boom");
              },
              (mapping, xml) -> {});
      IllegalStateException transformerFailure =
          assertThrows(
              IllegalStateException.class,
              () ->
                  transformerFailingController.exportMapping(
                      new CustomXmlMappingLocator(
                          java.util.Optional.of(1L), java.util.Optional.of("CORSO_mapping")),
                      false,
                      "UTF-8"));
      assertTrue(transformerFailure.getMessage().contains("CORSO_mapping"));

      CustomXmlAdapter validatingController =
          new CustomXmlAdapter(
              workbook.xssfWorkbook(),
              (mapping, output, encoding, validateSchema) -> {
                throw new SAXException("invalid");
              },
              (mapping, xml) -> {});
      IllegalArgumentException validationFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  validatingController.exportMapping(
                      new CustomXmlMappingLocator(
                          java.util.Optional.of(1L), java.util.Optional.empty()),
                      true,
                      "UTF-8"));
      assertTrue(validationFailure.getMessage().contains("Schema validation failed"));

      CustomXmlAdapter ioFailingController =
          new CustomXmlAdapter(
              workbook.xssfWorkbook(),
              (mapping, output, encoding, validateSchema) -> {},
              (mapping, xml) -> {
                throw new IOException("disk");
              });
      IllegalStateException importIoFailure =
          assertThrows(
              IllegalStateException.class,
              () ->
                  ioFailingController.importMapping(
                      new ResolvedCustomXmlImport(
                          new CustomXmlMappingLocator(
                              java.util.Optional.of(1L), java.util.Optional.empty()),
                          "<CORSO/>")));
      assertTrue(importIoFailure.getMessage().contains("CORSO_mapping"));

      CustomXmlAdapter invalidXmlController =
          new CustomXmlAdapter(
              workbook.xssfWorkbook(),
              (mapping, output, encoding, validateSchema) -> {},
              (mapping, xml) -> {
                throw new javax.xml.xpath.XPathExpressionException("xpath");
              });
      IllegalArgumentException invalidXmlFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  invalidXmlController.importMapping(
                      new ResolvedCustomXmlImport(
                          new CustomXmlMappingLocator(
                              java.util.Optional.of(1L), java.util.Optional.empty()),
                          "<CORSO/>")));
      assertTrue(invalidXmlFailure.getMessage().contains("Invalid XML"));
    }
  }

  @Test
  void constructorAndEncodingValidationRejectNullOrBlankInputs() throws IOException {
    try (var workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook()) {
      var exporterFailure =
          assertThrows(
              NullPointerException.class,
              () -> new CustomXmlAdapter(workbook, null, (mapping, xml) -> {}));
      assertEquals("exporter must not be null", exporterFailure.getMessage());
      var importerFailure =
          assertThrows(
              NullPointerException.class,
              () ->
                  new CustomXmlAdapter(
                      workbook, (mapping, output, encoding, validateSchema) -> {}, null));
      assertEquals("importer must not be null", importerFailure.getMessage());
      var encodingFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  new CustomXmlAdapter(workbook)
                      .exportMapping(
                          new CustomXmlMappingLocator(
                              java.util.Optional.of(1L), java.util.Optional.empty()),
                          false,
                          " "));
      assertEquals("encoding must not be blank", encodingFailure.getMessage());
    }
  }

  @Test
  void requireMappingRejectsEmptyAndAmbiguousSelections() throws Exception {
    IllegalArgumentException emptyFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                CustomXmlAdapter.requireMapping(
                    List.<org.apache.poi.xssf.usermodel.XSSFMap>of(),
                    new CustomXmlMappingLocator(
                        java.util.Optional.empty(), java.util.Optional.of("missing"))));
    assertTrue(emptyFailure.getMessage().contains("does not contain any custom XML mappings"));

    try (XssfWorkbook workbook = CustomXmlWorkbookSamples.openSimpleCustomXmlWorkbook()) {
      IllegalArgumentException mixedMismatchFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  CustomXmlAdapter.requireMapping(
                      workbook.xssfWorkbook(),
                      new CustomXmlMappingLocator(
                          java.util.Optional.of(1L), java.util.Optional.of("missing-name"))));
      assertTrue(mixedMismatchFailure.getMessage().contains("No custom XML mapping matched"));

      CTMap duplicateMap = CTMap.Factory.newInstance();
      duplicateMap.setID(2L);
      duplicateMap.setName("CORSO_mapping");
      duplicateMap.setRootElement("CORSO");
      duplicateMap.setSchemaID("Schema1");
      IllegalArgumentException ambiguousFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  CustomXmlAdapter.requireMapping(
                      List.of(
                          workbook.xssfWorkbook().getCustomXMLMappings().iterator().next(),
                          CustomXmlAdapterTestSupport.fakeMap(
                              duplicateMap, null, null, List.of(), List.of())),
                      new CustomXmlMappingLocator(
                          java.util.Optional.empty(), java.util.Optional.of("CORSO_mapping"))));
      assertTrue(ambiguousFailure.getMessage().contains("Multiple custom XML mappings matched"));
    }
  }

  @Test
  void dataBindingReturnsNullWhenCtMapHasNoBinding() {
    assertEquals(Optional.empty(), CustomXmlAdapter.dataBinding(CTMap.Factory.newInstance()));
  }

  @Test
  void dataBindingProjectsUnsetOptionalFields() {
    CTMap ctMap = CTMap.Factory.newInstance();
    ctMap.setID(41L);
    ctMap.setName("OrdersMapping");
    ctMap.setRootElement("Orders");
    ctMap.setSchemaID("Schema42");
    ctMap.setShowImportExportValidationErrors(true);
    ctMap.setAutoFit(false);
    ctMap.setAppend(true);
    ctMap.setPreserveSortAFLayout(false);
    ctMap.setPreserveFormat(true);
    ctMap.addNewDataBinding().setDataBindingLoadMode(7L);

    CustomXmlDataBindingReport dataBinding = CustomXmlAdapter.dataBinding(ctMap).orElseThrow();
    assertEquals(Optional.empty(), dataBinding.dataBindingName());
    assertEquals(Optional.empty(), dataBinding.fileBinding());
    assertEquals(Optional.empty(), dataBinding.connectionId());
    assertEquals(Optional.empty(), dataBinding.fileBindingName());
    assertEquals(7L, dataBinding.loadMode());
  }

  @Test
  void dataBindingProjectsPopulatedFields() {
    CTMap boundCtMap = CTMap.Factory.newInstance();
    boundCtMap.setID(42L);
    boundCtMap.setName("BoundOrders");
    boundCtMap.setRootElement("Orders");
    boundCtMap.setSchemaID("Schema42");
    var fullBinding = boundCtMap.addNewDataBinding();
    fullBinding.setDataBindingName("Binding");
    fullBinding.setFileBinding(true);
    fullBinding.setConnectionID(9L);
    fullBinding.setFileBindingName("orders.xml");
    fullBinding.setDataBindingLoadMode(5L);
    CustomXmlDataBindingReport fullDataBinding =
        CustomXmlAdapter.dataBinding(boundCtMap).orElseThrow();
    assertEquals("Binding", fullDataBinding.dataBindingName().orElseThrow());
    assertEquals(true, fullDataBinding.fileBinding().orElseThrow());
    assertEquals(9L, fullDataBinding.connectionId().orElseThrow());
    assertEquals("orders.xml", fullDataBinding.fileBindingName().orElseThrow());
    assertEquals(5L, fullDataBinding.loadMode());
  }

  @Test
  void snapshotProjectsSchemaAndLinkedTableMetadata() throws Exception {
    CTMap ctMap = CTMap.Factory.newInstance();
    ctMap.setID(41L);
    ctMap.setName("OrdersMapping");
    ctMap.setRootElement("Orders");
    ctMap.setSchemaID("Schema42");
    CTSchema schema = CTSchema.Factory.newInstance();
    schema.setNamespace(" ");
    schema.setSchemaLanguage("XSD");
    schema.setSchemaRef(" ");
    Node schemaNode = XMLHelper.newDocumentBuilder().newDocument().createElement("schema");
    schemaNode.setTextContent("orders");

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Orders");
      sheet.createRow(0).createCell(0).setCellValue("Order");
      CustomXmlAdapterTestSupport.FakeTable table = CustomXmlAdapterTestSupport.fakeTable(sheet);
      CustomXmlAdapterTestSupport.FakeMap map =
          CustomXmlAdapterTestSupport.fakeMap(ctMap, schema, schemaNode, List.of(), List.of(table));

      CustomXmlMappingReport snapshot = CustomXmlAdapter.snapshot(map);
      assertEquals("OrdersMapping", snapshot.name());
      assertEquals(Optional.empty(), snapshot.schema().namespace());
      assertEquals("XSD", snapshot.schema().language().orElseThrow());
      assertEquals(Optional.empty(), snapshot.schema().reference());
      assertTrue(snapshot.schema().xml().isPresent());
      assertEquals(List.of(), snapshot.linkedCells());
      assertEquals(1, snapshot.linkedTables().size());
      assertEquals("OrdersTable", snapshot.linkedTables().getFirst().tableName());
      assertEquals("Orders!A1:B2", snapshot.linkedTables().getFirst().range().value());
      assertEquals("OrdersMapping", CustomXmlAdapter.mappingName(map));
      assertEquals(Optional.empty(), CustomXmlAdapter.blankAsOptional(null));
      assertEquals(Optional.empty(), CustomXmlAdapter.blankAsOptional(" "));
      assertEquals(Optional.of("OrdersMapping"), CustomXmlAdapter.blankAsOptional("OrdersMapping"));
      assertEquals(
          Optional.empty(),
          CustomXmlAdapter.schemaXml(
              CustomXmlAdapterTestSupport.fakeMap(ctMap, schema, null, List.of(), List.of())));

      CTSchema emptySchema = CTSchema.Factory.newInstance();
      CustomXmlMappingReport emptySchemaSnapshot =
          CustomXmlAdapter.snapshot(
              CustomXmlAdapterTestSupport.fakeMap(
                  ctMap, emptySchema, schemaNode, List.of(), List.of(table)));
      assertEquals(Optional.empty(), emptySchemaSnapshot.schema().namespace());
      assertEquals(Optional.empty(), emptySchemaSnapshot.schema().language());
      assertEquals(Optional.empty(), emptySchemaSnapshot.schema().reference());

      CustomXmlMappingReport nullSchemaSnapshot =
          CustomXmlAdapter.snapshot(
              CustomXmlAdapterTestSupport.fakeMap(
                  ctMap, null, schemaNode, List.of(), List.of(table)));
      assertEquals(Optional.empty(), nullSchemaSnapshot.schema().namespace());
      assertEquals(Optional.empty(), nullSchemaSnapshot.schema().language());
      assertEquals(Optional.empty(), nullSchemaSnapshot.schema().reference());

      CTSchema richSchema = CTSchema.Factory.newInstance();
      richSchema.setNamespace("urn:orders");
      richSchema.setSchemaRef("orders.xsd");
      CustomXmlMappingReport richSchemaSnapshot =
          CustomXmlAdapter.snapshot(
              CustomXmlAdapterTestSupport.fakeMap(
                  ctMap, richSchema, schemaNode, List.of(), List.of(table)));
      assertEquals("urn:orders", richSchemaSnapshot.schema().namespace().orElseThrow());
      assertEquals(Optional.empty(), richSchemaSnapshot.schema().language());
      assertEquals("orders.xsd", richSchemaSnapshot.schema().reference().orElseThrow());
    }
  }

  @Test
  void schemaXmlReturnsNullWhenMapHasNoSchemaNode() throws Exception {
    CTMap ctMap = CTMap.Factory.newInstance();
    ctMap.setID(41L);
    ctMap.setName("OrdersMapping");
    CTSchema schema = CTSchema.Factory.newInstance();
    CustomXmlAdapterTestSupport.FakeMap fakeMap =
        CustomXmlAdapterTestSupport.fakeMap(ctMap, schema, null, List.of(), List.of());

    assertEquals(Optional.empty(), CustomXmlAdapter.schemaXml(fakeMap));
    assertEquals(Optional.empty(), CustomXmlAdapter.schemaXml(fakeMap, XMLHelper.newTransformer()));
  }

  @Test
  void schemaXmlWrapsTransformerFailuresWithMappingName() throws Exception {
    CTMap ctMap = CTMap.Factory.newInstance();
    ctMap.setID(41L);
    ctMap.setName("OrdersMapping");
    Node schemaNode = XMLHelper.newDocumentBuilder().newDocument().createElement("schema");
    schemaNode.setTextContent("orders");
    CTSchema schema = CTSchema.Factory.newInstance();
    CustomXmlAdapterTestSupport.FakeMap map =
        CustomXmlAdapterTestSupport.fakeMap(ctMap, schema, schemaNode, List.of(), List.of());

    IllegalStateException schemaFailure =
        assertThrows(
            IllegalStateException.class,
            () ->
                CustomXmlAdapter.schemaXml(map, CustomXmlAdapterTestSupport.failingTransformer()));
    assertTrue(schemaFailure.getMessage().contains("OrdersMapping"));
  }
}
