package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.customxml.CustomXmlDataBindingReport;
import dev.erst.gridgrind.core.model.customxml.CustomXmlExportReport;
import dev.erst.gridgrind.core.model.customxml.CustomXmlMappingLocator;
import dev.erst.gridgrind.core.model.customxml.CustomXmlMappingReport;
import dev.erst.gridgrind.core.model.customxml.ResolvedCustomXmlImport;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.excel.customxml.ExcelCustomXmlController;
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
class ExcelCustomXmlControllerTest {
  @Test
  void readsWorkbookCustomXmlMappingsAndExportsXml() throws IOException {
    try (ExcelWorkbook workbook = CustomXmlWorkbookSamples.openSimpleCustomXmlWorkbook()) {
      List<CustomXmlMappingReport> mappings = workbook.customXml().customXmlMappings();

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
      assertTrue(mapping.schema().xml().contains("<xsd:element name=\"CORSO\""));
      assertEquals(8, mapping.linkedCells().size());
      assertEquals("Foglio1", mapping.linkedCells().getFirst().sheetName().value());
      assertEquals("A1", mapping.linkedCells().getFirst().address().value());
      assertEquals("/CORSO/NOME", mapping.linkedCells().getFirst().xpath());
      assertEquals("string", mapping.linkedCells().getFirst().xmlDataType());
      assertEquals(List.of(), mapping.linkedTables());

      CustomXmlExportReport exported =
          workbook
              .customXml()
              .exportCustomXmlMapping(
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
    try (ExcelWorkbook workbook = CustomXmlWorkbookSamples.openSimpleCustomXmlWorkbook()) {
      workbook
          .customXml()
          .importCustomXmlMapping(
              new ResolvedCustomXmlImport(
                  new CustomXmlMappingLocator(
                      java.util.Optional.empty(), java.util.Optional.of("CORSO_mapping")),
                  "<CORSO><NOME>Grid</NOME><DOCENTE>Grind</DOCENTE><TUTOR>Agent</TUTOR>"
                      + "<CDL>Ops</CDL><DURATA>5</DURATA><ARGOMENTO>Audit</ARGOMENTO>"
                      + "<PROGETTO>Parity</PROGETTO><CREDITI>10</CREDITI></CORSO>"));

      WorkbookSheetResult.Window window = workbook.sheet("Foglio1").cells().window("A1", 1, 3);
      assertEquals("Grid", window.rows().getFirst().cells().get(0).displayValue());
      assertEquals("Grind", window.rows().getFirst().cells().get(1).displayValue());
      assertEquals("Agent", window.rows().getFirst().cells().get(2).displayValue());

      ExcelWorkbookIntrospector introspector = new ExcelWorkbookIntrospector();
      WorkbookCoreResult.CustomXmlMappingsResult mappings =
          assertInstanceOf(
              WorkbookCoreResult.CustomXmlMappingsResult.class,
              introspector.execute(
                  workbook, new WorkbookReadCommand.GetCustomXmlMappings(new StepId("maps"))));
      WorkbookCoreResult.CustomXmlExportResult exported =
          assertInstanceOf(
              WorkbookCoreResult.CustomXmlExportResult.class,
              introspector.execute(
                  workbook,
                  new WorkbookReadCommand.ExportCustomXmlMapping(
                      new StepId("export"),
                      new CustomXmlMappingLocator(
                          java.util.Optional.of(1L), java.util.Optional.empty()),
                      false,
                      "UTF-8")));

      assertEquals("maps", mappings.stepId().value());
      assertEquals(1, mappings.mappings().size());
      assertEquals("export", exported.stepId().value());
      assertTrue(exported.export().xml().contains("<NOME>Grid</NOME>"));
    }
  }

  @Test
  void rejectsUnknownCustomXmlMappings() throws IOException {
    try (ExcelWorkbook workbook = CustomXmlWorkbookSamples.openSimpleCustomXmlWorkbook()) {
      IllegalArgumentException failure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  workbook
                      .customXml()
                      .exportCustomXmlMapping(
                          new CustomXmlMappingLocator(
                              java.util.Optional.of(99L), java.util.Optional.of("missing")),
                          false,
                          "UTF-8"));
      assertTrue(failure.getMessage().contains("No custom XML mapping matched"));
    }
  }

  @Test
  void exportAndImportWrapExplicitPoiFailures() throws IOException {
    try (ExcelWorkbook workbook = CustomXmlWorkbookSamples.openSimpleCustomXmlWorkbook()) {
      ExcelCustomXmlController transformerFailingController =
          new ExcelCustomXmlController(
              (mapping, output, encoding, validateSchema) -> {
                throw new TransformerException("boom");
              },
              (mapping, xml) -> {});
      IllegalStateException transformerFailure =
          assertThrows(
              IllegalStateException.class,
              () ->
                  transformerFailingController.exportMapping(
                      workbook.xssfWorkbook(),
                      new CustomXmlMappingLocator(
                          java.util.Optional.of(1L), java.util.Optional.of("CORSO_mapping")),
                      false,
                      "UTF-8"));
      assertTrue(transformerFailure.getMessage().contains("CORSO_mapping"));

      ExcelCustomXmlController validatingController =
          new ExcelCustomXmlController(
              (mapping, output, encoding, validateSchema) -> {
                throw new SAXException("invalid");
              },
              (mapping, xml) -> {});
      IllegalArgumentException validationFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  validatingController.exportMapping(
                      workbook.xssfWorkbook(),
                      new CustomXmlMappingLocator(
                          java.util.Optional.of(1L), java.util.Optional.empty()),
                      true,
                      "UTF-8"));
      assertTrue(validationFailure.getMessage().contains("Schema validation failed"));

      ExcelCustomXmlController ioFailingController =
          new ExcelCustomXmlController(
              (mapping, output, encoding, validateSchema) -> {},
              (mapping, xml) -> {
                throw new IOException("disk");
              });
      IllegalStateException importIoFailure =
          assertThrows(
              IllegalStateException.class,
              () ->
                  ioFailingController.importMapping(
                      workbook.xssfWorkbook(),
                      new ResolvedCustomXmlImport(
                          new CustomXmlMappingLocator(
                              java.util.Optional.of(1L), java.util.Optional.empty()),
                          "<CORSO/>")));
      assertTrue(importIoFailure.getMessage().contains("CORSO_mapping"));

      ExcelCustomXmlController invalidXmlController =
          new ExcelCustomXmlController(
              (mapping, output, encoding, validateSchema) -> {},
              (mapping, xml) -> {
                throw new javax.xml.xpath.XPathExpressionException("xpath");
              });
      IllegalArgumentException invalidXmlFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  invalidXmlController.importMapping(
                      workbook.xssfWorkbook(),
                      new ResolvedCustomXmlImport(
                          new CustomXmlMappingLocator(
                              java.util.Optional.of(1L), java.util.Optional.empty()),
                          "<CORSO/>")));
      assertTrue(invalidXmlFailure.getMessage().contains("Invalid XML"));
    }
  }

  @Test
  void constructorAndEncodingValidationRejectNullOrBlankInputs() throws IOException {
    NullPointerException exporterFailure =
        assertThrows(
            NullPointerException.class,
            () -> new ExcelCustomXmlController(null, (mapping, xml) -> {}));
    assertEquals("exporter must not be null", exporterFailure.getMessage());

    NullPointerException importerFailure =
        assertThrows(
            NullPointerException.class,
            () ->
                new ExcelCustomXmlController(
                    (mapping, output, encoding, validateSchema) -> {}, null));
    assertEquals("importer must not be null", importerFailure.getMessage());

    try (ExcelWorkbook workbook = CustomXmlWorkbookSamples.openSimpleCustomXmlWorkbook()) {
      IllegalArgumentException noMappingsFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  new ExcelCustomXmlController()
                      .exportMapping(
                          workbook.xssfWorkbook(),
                          new CustomXmlMappingLocator(
                              java.util.Optional.of(1L), java.util.Optional.empty()),
                          false,
                          " "));
      assertEquals("encoding must not be blank", noMappingsFailure.getMessage());
    }
  }

  @Test
  void requireMappingRejectsEmptyAndAmbiguousSelections() throws Exception {
    IllegalArgumentException emptyFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                ExcelCustomXmlController.requireMapping(
                    List.<org.apache.poi.xssf.usermodel.XSSFMap>of(),
                    new CustomXmlMappingLocator(
                        java.util.Optional.empty(), java.util.Optional.of("missing"))));
    assertTrue(emptyFailure.getMessage().contains("does not contain any custom XML mappings"));

    try (ExcelWorkbook workbook = CustomXmlWorkbookSamples.openSimpleCustomXmlWorkbook()) {
      IllegalArgumentException mixedMismatchFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  ExcelCustomXmlController.requireMapping(
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
                  ExcelCustomXmlController.requireMapping(
                      List.of(
                          workbook.xssfWorkbook().getCustomXMLMappings().iterator().next(),
                          ExcelCustomXmlControllerTestSupport.fakeMap(
                              duplicateMap, null, null, List.of(), List.of())),
                      new CustomXmlMappingLocator(
                          java.util.Optional.empty(), java.util.Optional.of("CORSO_mapping"))));
      assertTrue(ambiguousFailure.getMessage().contains("Multiple custom XML mappings matched"));
    }
  }

  @Test
  void dataBindingReturnsNullWhenCtMapHasNoBinding() {
    assertEquals(
        Optional.empty(), ExcelCustomXmlController.dataBinding(CTMap.Factory.newInstance()));
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

    CustomXmlDataBindingReport dataBinding =
        ExcelCustomXmlController.dataBinding(ctMap).orElseThrow();
    org.junit.jupiter.api.Assertions.assertNull(dataBinding.dataBindingName());
    org.junit.jupiter.api.Assertions.assertNull(dataBinding.fileBinding());
    org.junit.jupiter.api.Assertions.assertNull(dataBinding.connectionId());
    org.junit.jupiter.api.Assertions.assertNull(dataBinding.fileBindingName());
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
        ExcelCustomXmlController.dataBinding(boundCtMap).orElseThrow();
    assertEquals("Binding", fullDataBinding.dataBindingName());
    assertEquals(true, fullDataBinding.fileBinding());
    assertEquals(9L, fullDataBinding.connectionId());
    assertEquals("orders.xml", fullDataBinding.fileBindingName());
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
      ExcelCustomXmlControllerTestSupport.FakeTable table =
          ExcelCustomXmlControllerTestSupport.fakeTable(sheet);
      ExcelCustomXmlControllerTestSupport.FakeMap map =
          ExcelCustomXmlControllerTestSupport.fakeMap(
              ctMap, schema, schemaNode, List.of(), List.of(table));

      CustomXmlMappingReport snapshot = ExcelCustomXmlController.snapshot(map);
      assertEquals("OrdersMapping", snapshot.name());
      org.junit.jupiter.api.Assertions.assertNull(snapshot.schema().namespace());
      assertEquals("XSD", snapshot.schema().language());
      org.junit.jupiter.api.Assertions.assertNull(snapshot.schema().reference());
      assertNotNull(snapshot.schema().xml());
      assertEquals(List.of(), snapshot.linkedCells());
      assertEquals(1, snapshot.linkedTables().size());
      assertEquals("OrdersTable", snapshot.linkedTables().getFirst().tableName());
      assertEquals("Orders!A1:B2", snapshot.linkedTables().getFirst().range().value());
      assertEquals("OrdersMapping", ExcelCustomXmlController.mappingName(map));
      assertEquals(Optional.empty(), ExcelCustomXmlController.blankAsOptional(null));
      assertEquals(Optional.empty(), ExcelCustomXmlController.blankAsOptional(" "));
      assertEquals(
          Optional.of("OrdersMapping"), ExcelCustomXmlController.blankAsOptional("OrdersMapping"));
      assertEquals(
          Optional.empty(),
          ExcelCustomXmlController.schemaXml(
              ExcelCustomXmlControllerTestSupport.fakeMap(
                  ctMap, schema, null, List.of(), List.of())));

      CTSchema emptySchema = CTSchema.Factory.newInstance();
      CustomXmlMappingReport emptySchemaSnapshot =
          ExcelCustomXmlController.snapshot(
              ExcelCustomXmlControllerTestSupport.fakeMap(
                  ctMap, emptySchema, schemaNode, List.of(), List.of(table)));
      org.junit.jupiter.api.Assertions.assertNull(emptySchemaSnapshot.schema().namespace());
      org.junit.jupiter.api.Assertions.assertNull(emptySchemaSnapshot.schema().language());
      org.junit.jupiter.api.Assertions.assertNull(emptySchemaSnapshot.schema().reference());

      CustomXmlMappingReport nullSchemaSnapshot =
          ExcelCustomXmlController.snapshot(
              ExcelCustomXmlControllerTestSupport.fakeMap(
                  ctMap, null, schemaNode, List.of(), List.of(table)));
      org.junit.jupiter.api.Assertions.assertNull(nullSchemaSnapshot.schema().namespace());
      org.junit.jupiter.api.Assertions.assertNull(nullSchemaSnapshot.schema().language());
      org.junit.jupiter.api.Assertions.assertNull(nullSchemaSnapshot.schema().reference());

      CTSchema richSchema = CTSchema.Factory.newInstance();
      richSchema.setNamespace("urn:orders");
      richSchema.setSchemaRef("orders.xsd");
      CustomXmlMappingReport richSchemaSnapshot =
          ExcelCustomXmlController.snapshot(
              ExcelCustomXmlControllerTestSupport.fakeMap(
                  ctMap, richSchema, schemaNode, List.of(), List.of(table)));
      assertEquals("urn:orders", richSchemaSnapshot.schema().namespace());
      org.junit.jupiter.api.Assertions.assertNull(richSchemaSnapshot.schema().language());
      assertEquals("orders.xsd", richSchemaSnapshot.schema().reference());
    }
  }

  @Test
  void schemaXmlReturnsNullWhenMapHasNoSchemaNode() throws Exception {
    CTMap ctMap = CTMap.Factory.newInstance();
    ctMap.setID(41L);
    ctMap.setName("OrdersMapping");
    CTSchema schema = CTSchema.Factory.newInstance();
    ExcelCustomXmlControllerTestSupport.FakeMap fakeMap =
        ExcelCustomXmlControllerTestSupport.fakeMap(ctMap, schema, null, List.of(), List.of());

    assertEquals(Optional.empty(), ExcelCustomXmlController.schemaXml(fakeMap));
    assertEquals(
        Optional.empty(), ExcelCustomXmlController.schemaXml(fakeMap, XMLHelper.newTransformer()));
  }

  @Test
  void schemaXmlWrapsTransformerFailuresWithMappingName() throws Exception {
    CTMap ctMap = CTMap.Factory.newInstance();
    ctMap.setID(41L);
    ctMap.setName("OrdersMapping");
    Node schemaNode = XMLHelper.newDocumentBuilder().newDocument().createElement("schema");
    schemaNode.setTextContent("orders");
    CTSchema schema = CTSchema.Factory.newInstance();
    ExcelCustomXmlControllerTestSupport.FakeMap map =
        ExcelCustomXmlControllerTestSupport.fakeMap(
            ctMap, schema, schemaNode, List.of(), List.of());

    IllegalStateException schemaFailure =
        assertThrows(
            IllegalStateException.class,
            () ->
                ExcelCustomXmlController.schemaXml(
                    map, ExcelCustomXmlControllerTestSupport.failingTransformer()));
    assertTrue(schemaFailure.getMessage().contains("OrdersMapping"));
  }

  @Test
  void schemaXmlWrapsTransformerFactoryFailuresWithMappingName() throws Exception {
    CTMap ctMap = CTMap.Factory.newInstance();
    ctMap.setID(52L);
    ctMap.setName("FactoryFailureMap");
    Node schemaNode = XMLHelper.newDocumentBuilder().newDocument().createElement("schema");
    ExcelCustomXmlControllerTestSupport.FakeMap map =
        ExcelCustomXmlControllerTestSupport.fakeMap(
            ctMap, CTSchema.Factory.newInstance(), schemaNode, List.of(), List.of());

    String originalFactory = System.getProperty("javax.xml.transform.TransformerFactory");
    System.setProperty(
        "javax.xml.transform.TransformerFactory", ThrowingTransformerFactory.class.getName());
    try {
      IllegalStateException failure =
          assertThrows(IllegalStateException.class, () -> ExcelCustomXmlController.schemaXml(map));
      assertTrue(failure.getMessage().contains("FactoryFailureMap"));
    } finally {
      if (originalFactory == null) {
        System.clearProperty("javax.xml.transform.TransformerFactory");
      } else {
        System.setProperty("javax.xml.transform.TransformerFactory", originalFactory);
      }
    }
  }
}
