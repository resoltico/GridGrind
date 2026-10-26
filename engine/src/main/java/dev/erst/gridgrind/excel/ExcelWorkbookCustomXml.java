package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.customxml.CustomXmlExportReport;
import dev.erst.gridgrind.core.model.customxml.CustomXmlMappingLocator;
import dev.erst.gridgrind.core.model.customxml.CustomXmlMappingReport;
import dev.erst.gridgrind.core.model.customxml.ResolvedCustomXmlImport;
import dev.erst.gridgrind.excel.customxml.ExcelCustomXmlController;
import java.util.List;
import java.util.Objects;

/** Workbook custom-XML mapping inspection and mutation operations. */
public final class ExcelWorkbookCustomXml {
  private final ExcelWorkbook workbook;

  ExcelWorkbookCustomXml(ExcelWorkbook workbook) {
    this.workbook = Objects.requireNonNull(workbook, "workbook must not be null");
  }

  /** Returns factual workbook custom-XML mapping metadata. */
  public List<CustomXmlMappingReport> customXmlMappings() {
    return new ExcelCustomXmlController().mappings(workbook.xssfWorkbook());
  }

  /** Exports XML for one existing workbook custom-XML mapping. */
  public CustomXmlExportReport exportCustomXmlMapping(
      CustomXmlMappingLocator locator, boolean validateSchema, String encoding) {
    return new ExcelCustomXmlController()
        .exportMapping(workbook.xssfWorkbook(), locator, validateSchema, encoding);
  }

  /** Imports one XML document into one existing workbook custom-XML mapping. */
  public ExcelWorkbook importCustomXmlMapping(ResolvedCustomXmlImport definition) {
    new ExcelCustomXmlController().importMapping(workbook.xssfWorkbook(), definition);
    return workbook;
  }
}
