package dev.erst.gridgrind.contract.query;

import com.fasterxml.jackson.annotation.JsonTypeName;
import dev.erst.gridgrind.contract.dto.CustomXmlExportReport;
import dev.erst.gridgrind.contract.dto.CustomXmlMappingReport;
import dev.erst.gridgrind.contract.dto.NamedRangeReport;
import dev.erst.gridgrind.contract.dto.OoxmlPackageSecurityReport;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionReport;
import dev.erst.gridgrind.contract.dto.WorkbookSummary;
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
  record WorkbookSummaryResult(String stepId, WorkbookSummary workbook)
      implements WorkbookInspectionResult {
    public WorkbookSummaryResult {
      stepId = InspectionResultValidationSupport.requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(workbook, "workbook must not be null");
    }
  }

  /** Returns OOXML package-encryption and package-signature facts. */
  @JsonTypeName("GET_PACKAGE_SECURITY")
  record PackageSecurityResult(String stepId, OoxmlPackageSecurityReport security)
      implements WorkbookInspectionResult {
    public PackageSecurityResult {
      stepId = InspectionResultValidationSupport.requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(security, "security must not be null");
    }
  }

  /** Returns workbook-level protection facts. */
  @JsonTypeName("GET_WORKBOOK_PROTECTION")
  record WorkbookProtectionResult(String stepId, WorkbookProtectionReport protection)
      implements WorkbookInspectionResult {
    public WorkbookProtectionResult {
      stepId = InspectionResultValidationSupport.requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(protection, "protection must not be null");
    }
  }

  /** Returns factual workbook custom-XML mapping metadata. */
  @JsonTypeName("GET_CUSTOM_XML_MAPPINGS")
  record CustomXmlMappingsResult(String stepId, List<CustomXmlMappingReport> mappings)
      implements WorkbookInspectionResult {
    public CustomXmlMappingsResult {
      stepId = InspectionResultValidationSupport.requireNonBlank(stepId, "stepId");
      mappings = InspectionResultValidationSupport.copyValues(mappings, "mappings");
    }
  }

  /** Returns XML exported from one selected workbook custom-XML mapping. */
  @JsonTypeName("EXPORT_CUSTOM_XML_MAPPING")
  record CustomXmlExportResult(String stepId, CustomXmlExportReport export)
      implements WorkbookInspectionResult {
    public CustomXmlExportResult {
      stepId = InspectionResultValidationSupport.requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(export, "export must not be null");
    }
  }

  /** Returns named ranges selected by the originating read operation. */
  @JsonTypeName("GET_NAMED_RANGES")
  record NamedRangesResult(String stepId, List<NamedRangeReport> namedRanges)
      implements WorkbookInspectionResult {
    public NamedRangesResult {
      stepId = InspectionResultValidationSupport.requireNonBlank(stepId, "stepId");
      namedRanges = InspectionResultValidationSupport.copyValues(namedRanges, "namedRanges");
    }
  }

  /** Returns sheet names matched by the originating sheet-presence assertion. */
  @JsonTypeName("GET_SHEETS")
  record SheetsResult(String stepId, List<String> sheetNames) implements WorkbookInspectionResult {
    public SheetsResult {
      stepId = InspectionResultValidationSupport.requireNonBlank(stepId, "stepId");
      sheetNames = InspectionResultValidationSupport.copyValues(sheetNames, "sheetNames");
    }
  }
}
