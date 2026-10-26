package dev.erst.gridgrind.excel;

import static dev.erst.gridgrind.excel.WorkbookResultSupport.copyValues;

import dev.erst.gridgrind.core.model.customxml.CustomXmlExportReport;
import dev.erst.gridgrind.core.model.customxml.CustomXmlMappingReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeReport;
import dev.erst.gridgrind.core.model.protection.WorkbookProtectionReport;
import dev.erst.gridgrind.core.model.security.OoxmlPackageSecurityReport;
import dev.erst.gridgrind.core.result.WorkbookSummary;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.util.List;
import java.util.Objects;

/** Workbook-level fact results and summary payloads. */
public sealed interface WorkbookCoreResult extends WorkbookReadIntrospectionResult
    permits WorkbookCoreResult.WorkbookSummaryResult,
        WorkbookCoreResult.PackageSecurityResult,
        WorkbookCoreResult.WorkbookProtectionResult,
        WorkbookCoreResult.CustomXmlMappingsResult,
        WorkbookCoreResult.CustomXmlExportResult,
        WorkbookCoreResult.NamedRangesResult {

  /** Returns workbook-level summary facts. */
  record WorkbookSummaryResult(StepId stepId, WorkbookSummary workbook)
      implements WorkbookCoreResult {
    public WorkbookSummaryResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(workbook, "workbook must not be null");
    }
  }

  /** Returns OOXML package-encryption and package-signature facts. */
  record PackageSecurityResult(StepId stepId, OoxmlPackageSecurityReport security)
      implements WorkbookCoreResult {
    public PackageSecurityResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(security, "security must not be null");
    }
  }

  /** Returns workbook-level protection facts. */
  record WorkbookProtectionResult(StepId stepId, WorkbookProtectionReport protection)
      implements WorkbookCoreResult {
    public WorkbookProtectionResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(protection, "protection must not be null");
    }
  }

  /** Returns factual workbook custom-XML mapping metadata. */
  record CustomXmlMappingsResult(StepId stepId, List<CustomXmlMappingReport> mappings)
      implements WorkbookCoreResult {
    public CustomXmlMappingsResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      mappings = copyValues(mappings, "mappings");
    }
  }

  /** Returns XML exported from one selected workbook custom-XML mapping. */
  record CustomXmlExportResult(StepId stepId, CustomXmlExportReport export)
      implements WorkbookCoreResult {
    public CustomXmlExportResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(export, "export must not be null");
    }
  }

  /** Returns selected named ranges. */
  record NamedRangesResult(StepId stepId, List<NamedRangeReport> namedRanges)
      implements WorkbookCoreResult {
    public NamedRangesResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      namedRanges = copyValues(namedRanges, "namedRanges");
    }
  }
}
