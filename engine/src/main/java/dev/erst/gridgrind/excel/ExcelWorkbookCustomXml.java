package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.customxml.ExcelCustomXmlController;
import dev.erst.gridgrind.excel.customxml.ExcelCustomXmlExportSnapshot;
import dev.erst.gridgrind.excel.customxml.ExcelCustomXmlImportDefinition;
import dev.erst.gridgrind.excel.customxml.ExcelCustomXmlMappingLocator;
import dev.erst.gridgrind.excel.customxml.ExcelCustomXmlMappingSnapshot;
import java.util.List;
import java.util.Objects;

/** Workbook custom-XML mapping inspection and mutation operations. */
public final class ExcelWorkbookCustomXml {
  private final ExcelWorkbook workbook;

  ExcelWorkbookCustomXml(ExcelWorkbook workbook) {
    this.workbook = Objects.requireNonNull(workbook, "workbook must not be null");
  }

  /** Returns factual workbook custom-XML mapping metadata. */
  public List<ExcelCustomXmlMappingSnapshot> customXmlMappings() {
    return new ExcelCustomXmlController().mappings(workbook.xssfWorkbook());
  }

  /** Exports XML for one existing workbook custom-XML mapping. */
  public ExcelCustomXmlExportSnapshot exportCustomXmlMapping(
      ExcelCustomXmlMappingLocator locator, boolean validateSchema, String encoding) {
    return new ExcelCustomXmlController()
        .exportMapping(workbook.xssfWorkbook(), locator, validateSchema, encoding);
  }

  /** Imports one XML document into one existing workbook custom-XML mapping. */
  public ExcelWorkbook importCustomXmlMapping(ExcelCustomXmlImportDefinition definition) {
    new ExcelCustomXmlController().importMapping(workbook.xssfWorkbook(), definition);
    return workbook;
  }
}
