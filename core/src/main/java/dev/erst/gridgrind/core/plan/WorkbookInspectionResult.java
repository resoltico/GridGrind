package dev.erst.gridgrind.core.plan;

import com.fasterxml.jackson.annotation.JsonTypeName;
import dev.erst.gridgrind.core.model.customxml.CustomXmlExportReport;
import dev.erst.gridgrind.core.model.customxml.CustomXmlMappingReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeReport;
import dev.erst.gridgrind.core.model.protection.WorkbookProtectionReport;
import dev.erst.gridgrind.core.model.security.OoxmlPackageSecurityReport;
import dev.erst.gridgrind.core.result.WorkbookSummary;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.util.List;
import java.util.Objects;

/** Workbook-scoped factual inspection results. */
public sealed interface WorkbookInspectionResult extends InspectionIntrospectionResult
    permits WorkbookInspectionResult.WorkbookSummaryResult,
        WorkbookInspectionResult.PackageSecurityResult,
        WorkbookInspectionResult.WorkbookProtectionResult,
        WorkbookInspectionResult.CustomXmlMappingsResult,
        WorkbookInspectionResult.CustomXmlExportResult,
        WorkbookInspectionResult.NamedRangesResult,
        WorkbookInspectionResult.SheetsResult {

  /** Returns workbook-level summary facts. */
  @JsonTypeName("GET_WORKBOOK_SUMMARY")
  record WorkbookSummaryResult(StepId stepId, WorkbookSummary workbook)
      implements WorkbookInspectionResult {
    public WorkbookSummaryResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(workbook, "workbook must not be null");
    }
  }

  /** Returns OOXML package-encryption and package-signature facts. */
  @JsonTypeName("GET_PACKAGE_SECURITY")
  record PackageSecurityResult(StepId stepId, OoxmlPackageSecurityReport security)
      implements WorkbookInspectionResult {
    public PackageSecurityResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(security, "security must not be null");
    }
  }

  /** Returns workbook-level protection facts. */
  @JsonTypeName("GET_WORKBOOK_PROTECTION")
  record WorkbookProtectionResult(StepId stepId, WorkbookProtectionReport protection)
      implements WorkbookInspectionResult {
    public WorkbookProtectionResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(protection, "protection must not be null");
    }
  }

  /** Returns factual workbook custom-XML mapping metadata. */
  @JsonTypeName("GET_CUSTOM_XML_MAPPINGS")
  record CustomXmlMappingsResult(StepId stepId, List<CustomXmlMappingReport> mappings)
      implements WorkbookInspectionResult {
    public CustomXmlMappingsResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      mappings = InspectionResultValidationSupport.copyValues(mappings, "mappings");
    }
  }

  /** Returns XML exported from one selected workbook custom-XML mapping. */
  @JsonTypeName("EXPORT_CUSTOM_XML_MAPPING")
  record CustomXmlExportResult(StepId stepId, CustomXmlExportReport export)
      implements WorkbookInspectionResult {
    public CustomXmlExportResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(export, "export must not be null");
    }
  }

  /** Returns named ranges selected by the originating read operation. */
  @JsonTypeName("GET_NAMED_RANGES")
  record NamedRangesResult(StepId stepId, List<NamedRangeReport> namedRanges)
      implements WorkbookInspectionResult {
    public NamedRangesResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      namedRanges = InspectionResultValidationSupport.copyValues(namedRanges, "namedRanges");
    }
  }

  /** Returns sheet names matched by the originating sheet-presence assertion. */
  @JsonTypeName("GET_SHEETS")
  record SheetsResult(StepId stepId, List<SheetLabel> sheetNames)
      implements WorkbookInspectionResult {
    public SheetsResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      sheetNames = InspectionResultValidationSupport.copyValues(sheetNames, "sheetNames");
    }
  }
}
