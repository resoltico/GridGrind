package dev.erst.gridgrind.excel;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import javax.xml.xpath.XPathExpressionException;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.extractor.XSSFExportToXml;
import org.apache.poi.xssf.extractor.XSSFImportFromXML;
import org.apache.poi.xssf.usermodel.XSSFMap;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.helpers.XSSFSingleXmlCell;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataBinding;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTMap;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSchema;
import org.w3c.dom.Node;
import org.xml.sax.SAXException;

/** Workbook custom-XML mapping inspect/import/export support built on POI XSSF. */
final class ExcelCustomXmlController {
  private final ExcelCustomXmlExporter exporter;
  private final ExcelCustomXmlImporter importer;

  ExcelCustomXmlController() {
    this(
        (mapping, output, encoding, validateSchema) ->
            new XSSFExportToXml(mapping).exportToXML(output, encoding, validateSchema),
        (mapping, xml) -> new XSSFImportFromXML(mapping).importFromXML(xml));
  }

  ExcelCustomXmlController(ExcelCustomXmlExporter exporter, ExcelCustomXmlImporter importer) {
    this.exporter = Objects.requireNonNull(exporter, "exporter must not be null");
    this.importer = Objects.requireNonNull(importer, "importer must not be null");
  }

  List<ExcelCustomXmlMappingSnapshot> mappings(XSSFWorkbook workbook) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    List<ExcelCustomXmlMappingSnapshot> snapshots = new ArrayList<>();
    for (XSSFMap mapping : workbook.getCustomXMLMappings()) {
      snapshots.add(snapshot(mapping));
    }
    return List.copyOf(snapshots);
  }

  ExcelCustomXmlExportSnapshot exportMapping(
      XSSFWorkbook workbook,
      ExcelCustomXmlMappingLocator locator,
      boolean validateSchema,
      String encoding) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(locator, "locator must not be null");
    Objects.requireNonNull(encoding, "encoding must not be null");
    if (encoding.isBlank()) {
      throw new IllegalArgumentException("encoding must not be blank");
    }

    Charset charset = Charset.forName(encoding);
    XSSFMap mapping = requireMapping(workbook, locator);
    ByteArrayOutputStream output = new ByteArrayOutputStream();
    try {
      exporter.export(mapping, output, charset.name(), validateSchema);
    } catch (TransformerException exception) {
      throw new IllegalStateException(
          "Failed to export custom XML mapping '" + mappingName(mapping) + "'", exception);
    } catch (SAXException exception) {
      throw new IllegalArgumentException(
          "Schema validation failed while exporting custom XML mapping '"
              + mappingName(mapping)
              + "'",
          exception);
    }
    return new ExcelCustomXmlExportSnapshot(
        snapshot(mapping), charset.name(), validateSchema, output.toString(charset));
  }

  void importMapping(XSSFWorkbook workbook, ExcelCustomXmlImportDefinition definition) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(definition, "definition must not be null");

    XSSFMap mapping = requireMapping(workbook, definition.mapping());
    try {
      importer.importXml(mapping, definition.xml());
    } catch (IOException exception) {
      throw new IllegalStateException(
          "Failed to import custom XML mapping '" + mappingName(mapping) + "'", exception);
    } catch (SAXException | XPathExpressionException exception) {
      throw new IllegalArgumentException(
          "Invalid XML for custom XML mapping '" + mappingName(mapping) + "'", exception);
    }
  }

  static XSSFMap requireMapping(XSSFWorkbook workbook, ExcelCustomXmlMappingLocator locator) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    return requireMapping(workbook.getCustomXMLMappings(), locator);
  }

  static XSSFMap requireMapping(
      Collection<? extends XSSFMap> mappings, ExcelCustomXmlMappingLocator locator) {
    Objects.requireNonNull(mappings, "mappings must not be null");
    Objects.requireNonNull(locator, "locator must not be null");
    if (mappings.isEmpty()) {
      throw new IllegalArgumentException("Workbook does not contain any custom XML mappings");
    }

    List<XSSFMap> matches = new ArrayList<>();
    for (XSSFMap mapping : mappings) {
      CTMap ctMap = mapping.getCtMap();
      boolean idMatches = locator.mapId() == null || locator.mapId().equals(ctMap.getID());
      boolean nameMatches =
          locator.name() == null || ctMap.getName().equalsIgnoreCase(locator.name());
      if (idMatches && nameMatches) {
        matches.add(mapping);
      }
    }
    if (matches.isEmpty()) {
      throw new IllegalArgumentException(
          "No custom XML mapping matched mapId=" + locator.mapId() + " name=" + locator.name());
    }
    if (matches.size() > 1) {
      throw new IllegalArgumentException(
          "Multiple custom XML mappings matched mapId="
              + locator.mapId()
              + " name="
              + locator.name());
    }
    return matches.getFirst();
  }

  static ExcelCustomXmlMappingSnapshot snapshot(XSSFMap mapping) {
    CTMap ctMap = mapping.getCtMap();
    CTSchema schema = mapping.getCTSchema();
    return new ExcelCustomXmlMappingSnapshot(
        ctMap.getID(),
        ctMap.getName(),
        ctMap.getRootElement(),
        ctMap.getSchemaID(),
        ctMap.getShowImportExportValidationErrors(),
        ctMap.getAutoFit(),
        ctMap.getAppend(),
        ctMap.getPreserveSortAFLayout(),
        ctMap.getPreserveFormat(),
        schema != null && schema.isSetNamespace() ? nullIfBlank(schema.getNamespace()) : null,
        schema != null && schema.isSetSchemaLanguage()
            ? nullIfBlank(schema.getSchemaLanguage())
            : null,
        schema != null && schema.isSetSchemaRef() ? nullIfBlank(schema.getSchemaRef()) : null,
        schemaXml(mapping).orElse(null),
        dataBinding(ctMap).orElse(null),
        linkedCells(mapping),
        linkedTables(mapping));
  }

  static Optional<ExcelCustomXmlDataBindingSnapshot> dataBinding(CTMap ctMap) {
    if (!ctMap.isSetDataBinding()) {
      return Optional.empty();
    }
    CTDataBinding dataBinding = ctMap.getDataBinding();
    return Optional.of(
        new ExcelCustomXmlDataBindingSnapshot(
            dataBinding.isSetDataBindingName()
                ? nullIfBlank(dataBinding.getDataBindingName())
                : null,
            dataBinding.isSetFileBinding() ? dataBinding.getFileBinding() : null,
            dataBinding.isSetConnectionID() ? dataBinding.getConnectionID() : null,
            dataBinding.isSetFileBindingName()
                ? nullIfBlank(dataBinding.getFileBindingName())
                : null,
            dataBinding.getDataBindingLoadMode()));
  }

  static List<ExcelCustomXmlLinkedCellSnapshot> linkedCells(XSSFMap mapping) {
    List<ExcelCustomXmlLinkedCellSnapshot> snapshots = new ArrayList<>();
    for (XSSFSingleXmlCell linkedCell : mapping.getRelatedSingleXMLCell()) {
      snapshots.add(
          new ExcelCustomXmlLinkedCellSnapshot(
              linkedCell.getReferencedCell().getSheet().getSheetName(),
              linkedCell.getReferencedCell().getAddress().formatAsString(),
              linkedCell.getXpath(),
              linkedCell.getXmlDataType()));
    }
    return List.copyOf(snapshots);
  }

  static List<ExcelCustomXmlLinkedTableSnapshot> linkedTables(XSSFMap mapping) {
    List<ExcelCustomXmlLinkedTableSnapshot> snapshots = new ArrayList<>();
    for (XSSFTable table : mapping.getRelatedTables()) {
      XSSFSheet sheet = table.getXSSFSheet();
      snapshots.add(
          new ExcelCustomXmlLinkedTableSnapshot(
              sheet.getSheetName(),
              table.getName(),
              table.getDisplayName(),
              table.getArea().formatAsString(),
              table.getCommonXpath()));
    }
    return List.copyOf(snapshots);
  }

  static Optional<String> schemaXml(XSSFMap mapping) {
    Node schema = mapping.getSchema();
    if (schema == null) {
      return Optional.empty();
    }
    try {
      return Optional.ofNullable(schemaXml(mapping, schema, XMLHelper.newTransformer()));
    } catch (TransformerException exception) {
      throw new IllegalStateException(
          "Failed to serialize custom XML schema for mapping '" + mappingName(mapping) + "'",
          exception);
    }
  }

  static Optional<String> schemaXml(XSSFMap mapping, Transformer transformer) {
    Node schema = mapping.getSchema();
    if (schema == null) {
      return Optional.empty();
    }
    return Optional.ofNullable(schemaXml(mapping, schema, transformer));
  }

  private static String schemaXml(XSSFMap mapping, Node schema, Transformer transformer) {
    try {
      Objects.requireNonNull(transformer, "transformer must not be null");
      transformer.setOutputProperty(OutputKeys.OMIT_XML_DECLARATION, "yes");
      transformer.setOutputProperty(OutputKeys.INDENT, "no");
      ByteArrayOutputStream output = new ByteArrayOutputStream();
      transformer.transform(new DOMSource(schema), new StreamResult(output));
      return nullIfBlank(output.toString(StandardCharsets.UTF_8));
    } catch (TransformerException exception) {
      throw new IllegalStateException(
          "Failed to serialize custom XML schema for mapping '" + mappingName(mapping) + "'",
          exception);
    }
  }

  static String mappingName(XSSFMap mapping) {
    return mapping.getCtMap().getName();
  }

  static String nullIfBlank(String value) {
    return value == null || value.isBlank() ? null : value;
  }
}
