package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Workbook-global metadata commands such as custom XML and named ranges. */
public sealed interface WorkbookMetadataCommand extends WorkbookCommand
    permits WorkbookMetadataCommand.ImportCustomXmlMapping,
        WorkbookMetadataCommand.SetNamedRange,
        WorkbookMetadataCommand.DeleteNamedRange {

  /** Imports one XML document into one existing workbook custom-XML mapping. */
  record ImportCustomXmlMapping(ExcelCustomXmlImportDefinition mapping)
      implements WorkbookMetadataCommand {
    public ImportCustomXmlMapping {
      Objects.requireNonNull(mapping, "mapping must not be null");
    }
  }

  /** Creates or replaces one named range in workbook or sheet scope. */
  record SetNamedRange(ExcelNamedRangeDefinition definition) implements WorkbookMetadataCommand {
    public SetNamedRange {
      Objects.requireNonNull(definition, "definition must not be null");
    }
  }

  /** Deletes one existing named range from workbook or sheet scope. */
  record DeleteNamedRange(String name, ExcelNamedRangeScope scope)
      implements WorkbookMetadataCommand {
    public DeleteNamedRange {
      name = ExcelNamedRangeDefinition.validateName(name);
      Objects.requireNonNull(scope, "scope must not be null");
    }
  }
}
